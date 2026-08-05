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
import { lettersToCol } from "../../model/address.js";
import type { CellKey, Range, RawCell, StyleObject } from "../../model/types.js";
import { unzip } from "../zip/zip.js";
import { parseStylesXml } from "./styles.js";
import { findElement, findElements, textOf } from "./xml.js";

/** One sheet as it comes out of the file, before becoming a full `Sheet`. */
export interface XlsxSheetData {
  name: string;
  cells: Record<CellKey, RawCell>;
  styles: Record<CellKey, StyleObject>;
  merges: Range[];
  rows: number;
  cols: number;
}

const REF_RE = /^([A-Za-z]+)(\d+)$/;

function parseRef(ref: string): { row: number; col: number } | null {
  const m = ref.match(REF_RE);
  if (!m?.[1] || !m[2]) return null;
  return { col: lettersToCol(m[1].toUpperCase()), row: parseInt(m[2], 10) - 1 };
}

/**
 * A .xlsb is also a ZIP, so ZIP magic alone cannot confirm an XLSX. This checks for
 * the XML workbook part and rejects the binary one with a clear message rather
 * than failing deep inside the parser.
 */
function assertXmlWorkbook(names: string[]): void {
  if (names.some((n) => /workbookBin\.bin$/i.test(n))) {
    throw new Error(
      "unsupported format: .xlsb (binary BIFF12). Save as .xlsx and retry.",
    );
  }
  if (!names.some((n) => /workbook\.xml$/i.test(n))) {
    throw new Error("not a spreadsheet: xl/workbook.xml is missing");
  }
}

/** Reads a .xlsx or .xlsm from a File, Blob, or raw bytes. */
export async function readXlsx(
  file: File | Blob | Uint8Array,
): Promise<XlsxSheetData[]> {
  const bytes =
    file instanceof Uint8Array
      ? file
      : new Uint8Array(await (file as Blob).arrayBuffer());

  const files = unzip(bytes);
  const names = Object.keys(files);
  assertXmlWorkbook(names);

  const decoder = new TextDecoder();
  const read = (name: string | undefined) =>
    name && files[name] ? decoder.decode(files[name]) : undefined;

  // Shared strings: `t="s"` cells hold an index into this table.
  const ssXml = read(names.find((n) => /sharedStrings\.xml$/i.test(n)));
  const sharedStrings = ssXml
    ? findElements(ssXml, "si").map((si) => textOf(si.inner))
    : [];

  const xfStyles = parseStylesXml(read(names.find((n) => /styles\.xml$/i.test(n))));

  // Sheet names come from workbook.xml; order falls back to the file numbering.
  const wbXml = read(names.find((n) => /xl\/workbook\.xml$/i.test(n)));
  const sheetNames = wbXml
    ? findElements(wbXml, "sheet").map((s, i) => s.attrs.name || `Sheet${i + 1}`)
    : [];

  const sheetFiles = names
    .filter((n) => /^xl\/worksheets\/sheet\d*\.xml$/i.test(n))
    .sort(
      (a, b) =>
        Number(a.match(/(\d+)/)?.[1] ?? 0) - Number(b.match(/(\d+)/)?.[1] ?? 0),
    );

  const sheets = sheetFiles.map((fname, i): XlsxSheetData => {
    const xml = decoder.decode(files[fname] as Uint8Array);
    const cells: Record<CellKey, RawCell> = {};
    const styles: Record<CellKey, StyleObject> = {};
    let maxR = 0;
    let maxC = 0;

    for (const c of findElements(xml, "c")) {
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
        value = `=${textOf(f.inner)}`;
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
      if (value !== "") cells[key] = value;

      const sIdx = c.attrs.s;
      if (sIdx) {
        const style = xfStyles[parseInt(sIdx, 10)];
        if (style) styles[key] = style;
      }
    }

    const merges: Range[] = [];
    for (const m of findElements(xml, "mergeCell")) {
      const ref = m.attrs.ref;
      if (!ref) continue;
      const [from, to] = ref.split(":");
      if (!from || !to) continue;
      const a = parseRef(from);
      const b = parseRef(to);
      if (!a || !b) continue;
      merges.push({ r1: a.row, c1: a.col, r2: b.row, c2: b.col });
    }

    return {
      name: sheetNames[i] ?? `Sheet${i + 1}`,
      cells,
      styles,
      merges,
      rows: maxR + 1,
      cols: maxC + 1,
    };
  });

  return sheets.length > 0
    ? sheets
    : [{ name: "Sheet1", cells: {}, styles: {}, merges: [], rows: 1, cols: 1 }];
}
