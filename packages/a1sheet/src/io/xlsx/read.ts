/**
 * XLSX reader. Ported from ref/xlsxIO.js:346-402, with `DOMParser` replaced by the
 * scanner in ./xml.ts so this works outside a browser.
 *
 * Returns EVERY sheet, in workbook order. The caller rebuilds full `Sheet` objects
 * from these (`makeSheet` plus an overlay), so imported sheets pick up default
 * colWidths, frozen panes, and so on.
 *
 * Trusts well-formed output from Excel, Sheets, and LibreOffice. Charts, pivot
 * tables, data validation, and conditional formatting are IGNORED, not errors.
 */
import { MalformedFileError, UnsupportedFormatError } from "../../errors.js";
import { lettersToCol } from "../../model/address.js";
import type {
  CellKey,
  CellValue,
  Range,
  RawCell,
  StyleObject,
} from "../../model/types.js";
import {
  type AsyncReadOptions,
  countOccurrences,
  createPacer,
} from "../progress.js";
import { listZipEntries, readZipMember } from "../zip/zip.js";
import { excelSerialToDaySerial } from "./dates.js";
import { parseStylesXml } from "./styles.js";
import { colWidthToPx, rowHeightToPx } from "./units.js";
import { findElement, findElements, iterElements, textOf } from "./xml.js";

/** One sheet as it comes out of the file, before becoming a full `Sheet`. */
export interface XlsxSheetData {
  name: string;
  cells: Record<CellKey, RawCell>;
  styles: Record<CellKey, StyleObject>;
  /**
   * The value the writing application last computed for each formula cell.
   * Kept because this engine implements a fraction of Excel's function library:
   * without it, every cell using something unsupported imports as `#NAME?` and a
   * workbook built on dynamic arrays or structured references arrives unreadable.
   */
  cachedValues: Record<CellKey, CellValue>;
  merges: Range[];
  rows: number;
  cols: number;
  /** Column index -> width in px, for columns the file marks as custom. */
  colWidths: Record<number, number>;
  /** Row index -> height in px, for rows the file marks as custom. */
  rowHeights: Record<number, number>;
}

const REF_RE = /^([A-Za-z]+)(\d+)$/;

function parseRef(ref: string): { row: number; col: number } | null {
  const m = ref.match(REF_RE);
  if (!m?.[1] || !m[2]) return null;
  return { col: lettersToCol(m[1].toUpperCase()), row: parseInt(m[2], 10) - 1 };
}

/**
 * Types the `<v>` of a formula cell. Excel writes a formula's string result as
 * `t="str"` and an error as `t="e"`, never as a shared-string index, so there is
 * no string table to consult here.
 */
function cachedValue(
  text: string,
  type: string | undefined,
  style: StyleObject | undefined,
): CellValue {
  if (type === "b") return text === "1";
  if (type === "str" || type === "e" || type === "inlineStr") return text;
  const n = Number.parseFloat(text);
  if (Number.isNaN(n)) return text;
  return style?.numFmt === "date" ? excelSerialToDaySerial(n) : n;
}

/**
 * A .xlsb is also a ZIP, so ZIP magic alone cannot confirm an XLSX. This checks for
 * the XML workbook part and rejects the binary one with a clear message rather
 * than failing deep inside the parser.
 */
function assertXmlWorkbook(names: string[]): void {
  if (names.some((n) => /workbookBin\.bin$/i.test(n))) {
    throw new UnsupportedFormatError(
      ".xlsb (binary BIFF12)",
      "Open it in Excel and save as .xlsx, then retry.",
    );
  }
  if (!names.some((n) => /workbook\.xml$/i.test(n))) {
    throw new MalformedFileError(
      "xl/workbook.xml is missing, so this ZIP is not a spreadsheet",
    );
  }
}

/**
 * Elements handled between checkpoints. Each checkpoint allocates a promise, so
 * checkpointing per cell would cost more than parsing one; a few thousand makes
 * that overhead unmeasurable while still yielding several times a second.
 */
const ELEMENTS_PER_CHECKPOINT = 4096;

/**
 * Reads a .xlsx or .xlsm from a File, Blob, or raw bytes.
 *
 * Pass `signal` to cancel — the read rejects with `AbortedError` at its next
 * checkpoint — and `onProgress` to drive a progress bar. Both are optional and a
 * small file behaves exactly as it did before they existed: the reader only yields
 * once it has actually held the thread for a frame.
 */
export async function readXlsx(
  file: File | Blob | Uint8Array,
  options: AsyncReadOptions = {},
): Promise<XlsxSheetData[]> {
  const bytes =
    file instanceof Uint8Array
      ? file
      : new Uint8Array(await (file as Blob).arrayBuffer());

  const pacer = createPacer(options);
  // Before any inflation, so a signal that is already aborted costs nothing.
  await pacer.checkpoint("decompressing", 0, "archive");

  const members = listZipEntries(bytes);
  assertXmlWorkbook(members.map((m) => m.name));

  // Inflate one member at a time so a large archive does not block in one burst.
  const files: Record<string, Uint8Array> = {};
  for (const [i, member] of members.entries()) {
    files[member.name] = readZipMember(member);
    await pacer.checkpoint("decompressing", (i + 1) / members.length, member.name);
  }

  const names = Object.keys(files);
  const decoder = new TextDecoder();
  const read = (name: string | undefined) =>
    name && files[name] ? decoder.decode(files[name]) : undefined;

  const ssXml = read(names.find((n) => /sharedStrings\.xml$/i.test(n)));
  const sheetFiles = names
    .filter((n) => /^xl\/worksheets\/sheet\d*\.xml$/i.test(n))
    .sort(
      (a, b) =>
        Number(a.match(/(\d+)/)?.[1] ?? 0) - Number(b.match(/(\d+)/)?.[1] ?? 0),
    );
  const sheetXml = sheetFiles.map((name) =>
    decoder.decode(files[name] as Uint8Array),
  );

  // Progress denominator: shared strings plus cells, counted before parsing. It is
  // an estimate — a `<c/>` with no attributes is not counted — and the pacer clamps
  // progress monotonically into 0..1, so an off-by-some never makes the bar jump back.
  const stringCount = ssXml ? countOccurrences(ssXml, "<si") : 0;
  const cellCounts = sheetXml.map((xml) => countOccurrences(xml, "<c "));
  const totalElements = Math.max(
    1,
    stringCount + cellCounts.reduce((sum, n) => sum + n, 0),
  );
  let parsed = 0;

  // Shared strings: `t="s"` cells hold an index into this table.
  const sharedStrings: string[] = [];
  if (ssXml) {
    for (const si of iterElements(ssXml, "si")) {
      sharedStrings.push(textOf(si.inner));
      parsed++;
      if (parsed % ELEMENTS_PER_CHECKPOINT === 0) {
        await pacer.checkpoint("parsing", parsed / totalElements, "shared strings");
      }
    }
  }

  const xfStyles = parseStylesXml(read(names.find((n) => /styles\.xml$/i.test(n))));

  // Sheet names come from workbook.xml; order falls back to the file numbering.
  const wbXml = read(names.find((n) => /xl\/workbook\.xml$/i.test(n)));
  const sheetNames = wbXml
    ? findElements(wbXml, "sheet").map((s, i) => s.attrs.name || `Sheet${i + 1}`)
    : [];

  const sheets: XlsxSheetData[] = [];
  for (const [i, xml] of sheetXml.entries()) {
    const name = sheetNames[i] ?? `Sheet${i + 1}`;
    const cells: Record<CellKey, RawCell> = {};
    const styles: Record<CellKey, StyleObject> = {};
    const cachedValues: Record<CellKey, CellValue> = {};
    let maxR = 0;
    let maxC = 0;

    for (const c of iterElements(xml, "c")) {
      parsed++;
      if (parsed % ELEMENTS_PER_CHECKPOINT === 0) {
        await pacer.checkpoint("parsing", parsed / totalElements, name);
      }

      const ref = c.attrs.r;
      if (!ref) continue;
      const pos = parseRef(ref);
      if (!pos) continue;

      const { row, col } = pos;
      if (row > maxR) maxR = row;
      if (col > maxC) maxC = col;

      const type = c.attrs.t;
      const f = findElement(c.inner, "f");
      const v = findElement(c.inner, "v");

      let value: string;
      if (f) {
        // `<f>` may be empty on the non-anchor cells of a shared or array
        // formula, where the anchor holds the expression. Those cells still
        // carry a `<v>`, so they read as literals rather than as blank formulas.
        const text = textOf(f.inner);
        value = text === "" ? (v ? textOf(v.inner) : "") : `=${text}`;
      } else if (type === "s") {
        const idx = parseInt(v ? textOf(v.inner) : "0", 10);
        value = sharedStrings[idx] ?? "";
      } else if (type === "inlineStr") {
        const is = findElement(c.inner, "is");
        value = is ? textOf(is.inner) : "";
      } else {
        value = v ? textOf(v.inner) : "";
      }

      const key = `${row}_${col}` as CellKey;

      const sIdx = c.attrs.s;
      const style = (sIdx ? xfStyles[parseInt(sIdx, 10)] : null) ?? undefined;
      if (style) styles[key] = style;

      // A date is a plain number in the file; only the format says otherwise.
      // Rebase it here rather than at display time, so the value in the model is
      // a day serial this engine's date functions can do arithmetic on.
      if (style?.numFmt === "date" && !f) {
        const serial = Number.parseFloat(value);
        if (Number.isFinite(serial)) value = String(excelSerialToDaySerial(serial));
      }

      if (value !== "") cells[key] = value;

      // Only formulas get a cached value: for a literal the raw content already
      // is the value, and storing it twice would let the two disagree.
      if (f && v) {
        const text = textOf(v.inner);
        if (text !== "") cachedValues[key] = cachedValue(text, type, style);
      }
    }

    // Excel writes a <col> for runs of columns whether or not the user touched
    // them, and a ht= on rows it merely laid out. `customWidth`/`customHeight`
    // is the flag that means "someone set this deliberately" — without it, a
    // file would import with every column pinned to a width nobody chose.
    const colWidths: Record<number, number> = {};
    for (const col of iterElements(xml, "col")) {
      if (col.attrs.customWidth !== "1" || col.attrs.width === undefined) continue;
      const width = Number.parseFloat(col.attrs.width);
      const min = Number.parseInt(col.attrs.min ?? "", 10);
      const max = Number.parseInt(col.attrs.max ?? "", 10);
      if (!Number.isFinite(width) || !Number.isFinite(min)) continue;
      const last = Number.isFinite(max) ? max : min;
      const px = colWidthToPx(width);
      // min/max are 1-based and inclusive, and one element covers a whole run.
      for (let c = min; c <= last; c++) colWidths[c - 1] = px;
      if (last > maxC + 1) maxC = last - 1;
    }

    const rowHeights: Record<number, number> = {};
    for (const row of iterElements(xml, "row")) {
      if (row.attrs.customHeight !== "1" || row.attrs.ht === undefined) continue;
      const points = Number.parseFloat(row.attrs.ht);
      const index = Number.parseInt(row.attrs.r ?? "", 10);
      if (!Number.isFinite(points) || !Number.isFinite(index)) continue;
      rowHeights[index - 1] = rowHeightToPx(points);
      if (index > maxR + 1) maxR = index - 1;
    }

    const merges: Range[] = [];
    for (const m of iterElements(xml, "mergeCell")) {
      const ref = m.attrs.ref;
      if (!ref) continue;
      const [from, to] = ref.split(":");
      if (!from || !to) continue;
      const a = parseRef(from);
      const b = parseRef(to);
      if (!a || !b) continue;
      merges.push({ r1: a.row, c1: a.col, r2: b.row, c2: b.col });
    }

    sheets.push({
      name,
      cells,
      styles,
      cachedValues,
      merges,
      rows: maxR + 1,
      cols: maxC + 1,
      colWidths,
      rowHeights,
    });
    await pacer.checkpoint("parsing", parsed / totalElements, name);
  }

  pacer.finish(`${sheets.length} sheets`);

  return sheets.length > 0
    ? sheets
    : [
        {
          name: "Sheet1",
          cells: {},
          styles: {},
          cachedValues: {},
          merges: [],
          rows: 1,
          cols: 1,
          colWidths: {},
          rowHeights: {},
        },
      ];
}
