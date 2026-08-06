/**
 * XLSX writer. Ported from ref/xlsxIO.js:207-293.
 *
 * Takes one entry PER SHEET (the POC's `buildWorkbookXlsx` is plural — a common
 * misreading). Builds a style table shared across ALL sheets, deduped by
 * `styleKey`, then emits fonts, fills, and cellXfs into styles.xml.
 *
 * Formula cells are written with BOTH the formula and its cached value, which is
 * what lets Excel show a value before it recalculates.
 *
 * Output uses STORE (method 0) for every ZIP entry — valid but larger than
 * necessary. See ../zip/deflate.ts.
 */
import { createEvaluator } from "../../formula/evaluate.js";
import { colToLetters } from "../../model/address.js";
import type {
  CellKey,
  NamedRanges,
  Range,
  RawCell,
  StyleObject,
} from "../../model/types.js";
import { makeZip, type ZipEntry } from "../zip/zip.js";
import { CUSTOM_NUMFMTS, NUMFMT_TO_ID, styleKey } from "./styles.js";
import { pxToColWidth, pxToRowHeight } from "./units.js";
import { colorToRgb, xmlEscape } from "./xml.js";

/** Input shape for one sheet being written. */
export interface XlsxSheetInput {
  name: string;
  cells: Record<CellKey, RawCell>;
  styles: Record<CellKey, StyleObject>;
  merges: Range[];
  namedRanges?: NamedRanges;
  /** Column index -> width in px. Only entries present are written. */
  colWidths?: Record<number, number>;
  /** Row index -> height in px. Only entries present are written. */
  rowHeights?: Record<number, number>;
}

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const NS_REL =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships";

const INTEGER_OR_DECIMAL = /^-?\d+(\.\d+)?$/;

/** Builds a complete .xlsx as bytes. */
export function writeXlsx(sheets: XlsxSheetInput[]): Uint8Array {
  const encoder = new TextEncoder();

  // Style table shared across every sheet. Index 0 is the default, so a cell with
  // no style needs no `s` attribute at all.
  const DEFAULT_STYLE: StyleObject = {};
  const styleList: StyleObject[] = [DEFAULT_STYLE];
  const styleIndex = new Map<string, number>([[styleKey(DEFAULT_STYLE), 0]]);

  function xfIndexFor(style: StyleObject | undefined): number {
    if (!style) return 0;
    const key = styleKey(style);
    const existing = styleIndex.get(key);
    if (existing !== undefined) return existing;
    styleList.push(style);
    const idx = styleList.length - 1;
    styleIndex.set(key, idx);
    return idx;
  }

  const sheetXmls = sheets.map((sheet) => {
    const evaluator = createEvaluator(sheet.cells, sheet.namedRanges ?? {});

    // Extent must cover styled-but-empty cells too, or formatting is lost.
    let maxR = -1;
    let maxC = -1;
    const note = (key: string) => {
      const i = key.indexOf("_");
      const r = Number(key.slice(0, i));
      const c = Number(key.slice(i + 1));
      if (r > maxR) maxR = r;
      if (c > maxC) maxC = c;
    };
    for (const key of Object.keys(sheet.cells)) {
      if (sheet.cells[key as CellKey]) note(key);
    }
    for (const key of Object.keys(sheet.styles ?? {})) note(key);

    // A resized row or column past the last value is still part of the sheet.
    const colWidths = sheet.colWidths ?? {};
    const rowHeights = sheet.rowHeights ?? {};
    for (const key of Object.keys(rowHeights)) {
      const r = Number(key);
      if (r > maxR) maxR = r;
    }
    for (const key of Object.keys(colWidths)) {
      const c = Number(key);
      if (c > maxC) maxC = c;
    }

    const rows = Math.max(maxR + 1, 1);
    const cols = Math.max(maxC + 1, 1);

    let rowsXml = "";
    for (let r = 0; r < rows; r++) {
      let rowCells = "";
      let hasContent = false;

      for (let c = 0; c < cols; c++) {
        const key = `${r}_${c}` as CellKey;
        const raw = sheet.cells[key];
        const style = sheet.styles?.[key];
        if ((raw === undefined || raw === "") && !style) continue;

        hasContent = true;
        const xfIdx = xfIndexFor(style);
        const ref = `${colToLetters(c)}${r + 1}`;
        const sAttr = xfIdx ? ` s="${xfIdx}"` : "";

        if (raw?.startsWith("=")) {
          const val = evaluator.getCellDisplay(r, c);
          const numVal = typeof val === "number" ? val : 0;
          rowCells += `<c r="${ref}"${sAttr}><f>${xmlEscape(raw.slice(1))}</f><v>${numVal}</v></c>`;
        } else if (raw && INTEGER_OR_DECIMAL.test(raw)) {
          rowCells += `<c r="${ref}"${sAttr}><v>${raw}</v></c>`;
        } else if (raw) {
          rowCells += `<c r="${ref}"${sAttr} t="inlineStr"><is><t xml:space="preserve">${xmlEscape(raw)}</t></is></c>`;
        } else {
          rowCells += `<c r="${ref}"${sAttr}/>`;
        }
      }

      // A row with a custom height but no content still has to be written, or
      // the height has nothing to hang off and is lost on the round trip.
      const height = rowHeights[r];
      const heightAttr =
        height === undefined
          ? ""
          : ` ht="${pxToRowHeight(height)}" customHeight="1"`;
      if (hasContent || heightAttr) {
        rowsXml += `<row r="${r + 1}"${heightAttr}>${rowCells}</row>`;
      }
    }

    const merges = sheet.merges ?? [];
    const mergeXml = merges.length
      ? `<mergeCells count="${merges.length}">${merges
          .map(
            (m) =>
              `<mergeCell ref="${colToLetters(m.c1)}${m.r1 + 1}:${colToLetters(m.c2)}${m.r2 + 1}"/>`,
          )
          .join("")}</mergeCells>`
      : "";

    // <cols> must precede <sheetData>; Excel rejects the file otherwise. One
    // element per column rather than per run — runs would need the widths
    // sorted and grouped for a saving that does not matter at this scale.
    const widthEntries = Object.keys(colWidths)
      .map(Number)
      .filter((c) => Number.isFinite(c))
      .sort((a, b) => a - b);
    const colsXml = widthEntries.length
      ? `<cols>${widthEntries
          .map(
            (c) =>
              `<col min="${c + 1}" max="${c + 1}" width="${pxToColWidth(colWidths[c] as number)}" customWidth="1"/>`,
          )
          .join("")}</cols>`
      : "";

    return `${XML_DECL}<worksheet xmlns="${NS_MAIN}">${colsXml}<sheetData>${rowsXml}</sheetData>${mergeXml}</worksheet>`;
  });

  // --- styles.xml, built from the collected style list ---

  const fontList: StyleObject[] = [];
  const fontIndex = new Map<string, number>();
  // Indices 0 and 1 are reserved by the format (none, gray125).
  const fillList: { bg?: string }[] = [{}, {}];
  const fillIndex = new Map<string, number>();

  const xfEntries = styleList.map((s) => {
    const fKey = JSON.stringify([
      !!s.bold,
      !!s.italic,
      !!s.underline,
      s.color ?? "",
    ]);
    let fontId = fontIndex.get(fKey);
    if (fontId === undefined) {
      fontId = fontList.length;
      fontList.push(s);
      fontIndex.set(fKey, fontId);
    }

    let fillId = 0;
    if (s.bg) {
      const existing = fillIndex.get(s.bg);
      if (existing === undefined) {
        fillId = fillList.length;
        fillList.push({ bg: s.bg });
        fillIndex.set(s.bg, fillId);
      } else {
        fillId = existing;
      }
    }

    return {
      fontId,
      fillId,
      numFmtId: NUMFMT_TO_ID[s.numFmt ?? "general"] ?? 0,
      align: s.align ?? "",
      locked: !!s.locked,
    };
  });

  const numFmtsXml = Object.entries(CUSTOM_NUMFMTS)
    .map(
      ([id, code]) => `<numFmt numFmtId="${id}" formatCode="${xmlEscape(code)}"/>`,
    )
    .join("");

  const fontsXml = fontList
    .map(
      (s) =>
        `<font>${s.bold ? "<b/>" : ""}${s.italic ? "<i/>" : ""}${
          s.underline ? "<u/>" : ""
        }<sz val="11"/>${
          s.color ? `<color rgb="${colorToRgb(s.color)}"/>` : `<color theme="1"/>`
        }<name val="Calibri"/></font>`,
    )
    .join("");

  const fillsXml = fillList
    .map((f) =>
      f.bg
        ? `<fill><patternFill patternType="solid"><fgColor rgb="${colorToRgb(f.bg)}"/></patternFill></fill>`
        : `<fill><patternFill patternType="none"/></fill>`,
    )
    .join("");

  const xfsXml = xfEntries
    .map(
      (xf) =>
        `<xf numFmtId="${xf.numFmtId}" fontId="${xf.fontId}" fillId="${xf.fillId}" borderId="0" xfId="0" applyFont="1" applyFill="1" applyNumberFormat="1"${
          xf.align ? ` applyAlignment="1"` : ""
        }${xf.locked ? ` applyProtection="1"` : ""}>${
          xf.align ? `<alignment horizontal="${xf.align}"/>` : ""
        }${xf.locked ? `<protection locked="1"/>` : ""}</xf>`,
    )
    .join("");

  const stylesXml =
    `${XML_DECL}<styleSheet xmlns="${NS_MAIN}">` +
    `<numFmts count="${Object.keys(CUSTOM_NUMFMTS).length}">${numFmtsXml}</numFmts>` +
    `<fonts count="${fontList.length}">${fontsXml}</fonts>` +
    `<fills count="${fillList.length}">${fillsXml}</fills>` +
    `<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>` +
    `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
    `<cellXfs count="${xfEntries.length}">${xfsXml}</cellXfs></styleSheet>`;

  // --- package parts ---

  const contentTypes =
    `${XML_DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
    `<Default Extension="xml" ContentType="application/xml"/>` +
    `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
    `<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>` +
    sheets
      .map(
        (_, i) =>
          `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`,
      )
      .join("") +
    `</Types>`;

  const rootRels =
    `${XML_DECL}<Relationships xmlns="${NS_PKG_REL}">` +
    `<Relationship Id="rId1" Type="${NS_REL}/officeDocument" Target="xl/workbook.xml"/>` +
    `</Relationships>`;

  const definedNames = collectDefinedNames(sheets);
  const workbookXml =
    `${XML_DECL}<workbook xmlns="${NS_MAIN}" xmlns:r="${NS_REL}"><sheets>` +
    sheets
      .map(
        (s, i) =>
          `<sheet name="${xmlEscape(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`,
      )
      .join("") +
    `</sheets>${definedNames}</workbook>`;

  const workbookRels =
    `${XML_DECL}<Relationships xmlns="${NS_PKG_REL}">` +
    sheets
      .map(
        (_, i) =>
          `<Relationship Id="rId${i + 1}" Type="${NS_REL}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`,
      )
      .join("") +
    `<Relationship Id="rId${sheets.length + 1}" Type="${NS_REL}/styles" Target="styles.xml"/>` +
    `</Relationships>`;

  const files: ZipEntry[] = [
    { name: "[Content_Types].xml", data: encoder.encode(contentTypes) },
    { name: "_rels/.rels", data: encoder.encode(rootRels) },
    { name: "xl/workbook.xml", data: encoder.encode(workbookXml) },
    { name: "xl/_rels/workbook.xml.rels", data: encoder.encode(workbookRels) },
    { name: "xl/styles.xml", data: encoder.encode(stylesXml) },
    ...sheetXmls.map((xml, i) => ({
      name: `xl/worksheets/sheet${i + 1}.xml`,
      data: encoder.encode(xml),
    })),
  ];

  return makeZip(files);
}

/**
 * Named ranges are workbook-level in our model, so they are taken from the first
 * sheet that carries them and scoped to that sheet's name on the way out.
 */
function collectDefinedNames(sheets: XlsxSheetInput[]): string {
  const owner = sheets.find(
    (s) => s.namedRanges && Object.keys(s.namedRanges).length > 0,
  );
  if (!owner?.namedRanges) return "";

  const entries = Object.entries(owner.namedRanges);
  if (entries.length === 0) return "";

  const quoted = `'${owner.name.replace(/'/g, "''")}'`;
  return `<definedNames>${entries
    .map(([name, r]) => {
      const ref = `${quoted}!$${colToLetters(r.c1)}$${r.r1 + 1}:$${colToLetters(r.c2)}$${r.r2 + 1}`;
      return `<definedName name="${xmlEscape(name)}">${xmlEscape(ref)}</definedName>`;
    })
    .join("")}</definedNames>`;
}
