/**
 * XLSX writer.
 *
 * Takes one entry PER SHEET. Builds a style table shared across ALL sheets,
 * deduped by `styleKey`, then emits fonts, fills, and cellXfs into styles.xml.
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
  CellImage,
  CellKey,
  CellValue,
  CondFormat,
  DataValidation,
  NamedRanges,
  Range,
  RawCell,
  SheetTable,
  StyleObject,
} from "../../model/types.js";
import { makeZip, type ZipEntry } from "../zip/zip.js";
import { styleKey } from "./styles.js";
import { pxToColWidth, pxToRowHeight } from "./units.js";
import { validationXml } from "./validation.js";
import {
  condFormatXml,
  dxfXml,
  imageKey,
  imageParts,
  RICH_VALUE_STRUCTURE,
  tableXml,
} from "./writeParts.js";
import { buildStylesXml } from "./writeStyles.js";
import { xmlEscape } from "./xml.js";

/** Input shape for one sheet being written. */
export interface XlsxSheetInput {
  name: string;
  cells: Record<CellKey, RawCell>;
  styles: Record<CellKey, StyleObject>;
  merges: Range[];
  namedRanges?: NamedRanges;
  /**
   * Values from a previous import, for formulas this engine cannot evaluate.
   * Passing them keeps an exported formula's `<v>` truthful instead of writing
   * the zero an unevaluable formula would otherwise produce.
   */
  cachedValues?: Record<CellKey, CellValue>;
  /** Column index -> width in px. Only entries present are written. */
  colWidths?: Record<number, number>;
  /** Row index -> height in px. Only entries present are written. */
  rowHeights?: Record<number, number>;
  /** Named tables on this sheet. Written as `xl/tables/*.xml`. */
  tables?: readonly SheetTable[];
  /** Columns and rows to mark hidden. */
  hiddenCols?: ReadonlySet<number>;
  hiddenRows?: ReadonlySet<number>;
  /** Data-validation rules, which is what makes a cell a dropdown. */
  validations?: readonly DataValidation[];
  /** Conditional formats. Their styles are written into `<dxfs>`. */
  condFormats?: readonly CondFormat[];
  /** In-cell images, written through the rich-value chain. */
  images?: Record<CellKey, CellImage>;
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

  // Images are workbook-level: one rich-value table, indexed from every sheet.
  const allImages = sheets.flatMap((sheet, index) =>
    Object.entries(sheet.images ?? {}).map(([key, image]) => ({
      sheet: index,
      key: key as CellKey,
      image,
    })),
  );
  const pictures = imageParts(allImages);

  /** Differential styles, accumulated across every sheet's conditional formats. */
  const dxfs: StyleObject[] = [];
  /** Table parts, numbered across the workbook because their ids must be unique. */
  const tableParts: { path: string; xml: string; sheet: number; relId: string }[] =
    [];

  const sheetXmls = sheets.map((sheet, sheetIndex) => {
    const evaluator = createEvaluator(sheet.cells, sheet.namedRanges ?? {}, {
      cachedValues: sheet.cachedValues ?? {},
    });

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
    // A cell may hold nothing but a picture, and the extent is what decides
    // which cells are written at all.
    for (const key of Object.keys(sheet.images ?? {})) note(key);

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
        const hasImage = sheet.images?.[key] !== undefined;
        if ((raw === undefined || raw === "") && !style && !hasImage) continue;

        hasContent = true;
        const xfIdx = xfIndexFor(style);
        const ref = `${colToLetters(c)}${r + 1}`;
        const sAttr = xfIdx ? ` s="${xfIdx}"` : "";
        // An image cell carries the index into the rich-value chain that holds
        // the picture; without it the formula is just a URL to Excel.
        const vm = pictures?.valueMetadata.get(imageKey(sheetIndex, key));
        const vmAttr = vm === undefined ? "" : ` vm="${vm}"`;

        if (raw?.startsWith("=")) {
          const val = evaluator.getCellDisplay(r, c);
          const f = `<f>${xmlEscape(raw.slice(1))}</f>`;
          // A text result needs t="str", or Excel reads the string as a number
          // and shows 0. Only numbers may be written bare.
          rowCells +=
            typeof val === "number"
              ? `<c r="${ref}"${sAttr}${vmAttr}>${f}<v>${val}</v></c>`
              : `<c r="${ref}"${sAttr}${vmAttr} t="str">${f}<v>${xmlEscape(String(val))}</v></c>`;
        } else if (raw && INTEGER_OR_DECIMAL.test(raw)) {
          rowCells += `<c r="${ref}"${sAttr}${vmAttr}><v>${Number(raw)}</v></c>`;
        } else if (raw) {
          rowCells += `<c r="${ref}"${sAttr}${vmAttr} t="inlineStr"><is><t xml:space="preserve">${xmlEscape(raw)}</t></is></c>`;
        } else {
          rowCells += `<c r="${ref}"${sAttr}${vmAttr}/>`;
        }
      }

      // A row with a custom height but no content still has to be written, or
      // the height has nothing to hang off and is lost on the round trip.
      const height = rowHeights[r];
      const heightAttr =
        height === undefined
          ? ""
          : ` ht="${pxToRowHeight(height)}" customHeight="1"`;
      const hiddenAttr = sheet.hiddenRows?.has(r) ? ` hidden="1"` : "";
      if (hasContent || heightAttr || hiddenAttr) {
        rowsXml += `<row r="${r + 1}"${heightAttr}${hiddenAttr}>${rowCells}</row>`;
      }
    }

    const conditional = condFormatXml(sheet.condFormats ?? [], dxfs.length);
    dxfs.push(...conditional.dxfs);

    const sheetTables = sheet.tables ?? [];
    const tableRels = sheetTables.map((table, i) => {
      const id = tableParts.length + 1;
      const relId = `rIdTable${i + 1}`;
      tableParts.push({
        path: `xl/tables/table${id}.xml`,
        xml: tableXml(table, id),
        sheet: sheetIndex,
        relId,
      });
      return relId;
    });
    const tablePartsXml = tableRels.length
      ? `<tableParts count="${tableRels.length}">${tableRels
          .map((relId) => `<tablePart r:id="${relId}"/>`)
          .join("")}</tableParts>`
      : "";

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
    // A column needs an entry if it has a width OR is hidden; hiding is
    // independent of sizing, and a hidden column at the default width has no
    // width entry to hang off.
    const hiddenColSet = sheet.hiddenCols ?? new Set<number>();
    const colEntries = [
      ...new Set([
        ...Object.keys(colWidths).map(Number).filter(Number.isFinite),
        ...hiddenColSet,
      ]),
    ].sort((a, b) => a - b);
    const colsXml = colEntries.length
      ? `<cols>${colEntries
          .map((c) => {
            const width = colWidths[c];
            const widthAttr =
              width === undefined
                ? ""
                : ` width="${pxToColWidth(width)}" customWidth="1"`;
            const hidden = hiddenColSet.has(c) ? ` hidden="1"` : "";
            return `<col min="${c + 1}" max="${c + 1}"${widthAttr}${hidden}/>`;
          })
          .join("")}</cols>`
      : "";

    // Order is fixed by the schema: cols, sheetData, mergeCells,
    // conditionalFormatting, then tableParts. Excel rejects the file otherwise.
    return (
      `${XML_DECL}<worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_REL}">` +
      `${colsXml}<sheetData>${rowsXml}</sheetData>${mergeXml}` +
      `${conditional.sheet}${validationXml(sheet.validations ?? [])}` +
      `${tablePartsXml}</worksheet>`
    );
  });

  // --- styles.xml, built from the collected style list ---

  const stylesXml = buildStylesXml(styleList, dxfs.map(dxfXml).join(""));

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
    tableParts
      .map(
        (t) =>
          `<Override PartName="/${t.path}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>`,
      )
      .join("") +
    (pictures
      ? // Media extensions are declared as Defaults; everything else is an
        // Override on a specific part.
        [...pictures.extensions]
          .map(
            (ext) =>
              `<Default Extension="${ext}" ContentType="image/${ext === "jpeg" ? "jpeg" : ext}"/>`,
          )
          .join("") +
        `<Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>` +
        `<Override PartName="/xl/richData/rdrichvalue.xml" ContentType="application/vnd.ms-excel.rdrichvalue+xml"/>` +
        `<Override PartName="/xl/richData/rdrichvaluestructure.xml" ContentType="application/vnd.ms-excel.rdrichvaluestructure+xml"/>` +
        `<Override PartName="/xl/richData/rdRichValueWebImage.xml" ContentType="application/vnd.ms-excel.rdrichvaluewebimage+xml"/>`
      : "") +
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
    (pictures
      ? `<Relationship Id="rIdMeta" Type="${NS_REL}/sheetMetadata" Target="metadata.xml"/>` +
        `<Relationship Id="rIdRv" Type="http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue" Target="richData/rdrichvalue.xml"/>` +
        `<Relationship Id="rIdRvs" Type="http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure" Target="richData/rdrichvaluestructure.xml"/>` +
        `<Relationship Id="rIdRvw" Type="http://schemas.microsoft.com/office/2020/07/relationships/rdRichValueWebImage" Target="richData/rdRichValueWebImage.xml"/>`
      : "") +
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
    ...tableParts.map((t) => ({ name: t.path, data: encoder.encode(t.xml) })),
    // A worksheet reaches its tables through its own relationship part.
    ...sheets
      .map((_, i) => ({ index: i, own: tableParts.filter((t) => t.sheet === i) }))
      .filter(({ own }) => own.length > 0)
      .map(({ index, own }) => ({
        name: `xl/worksheets/_rels/sheet${index + 1}.xml.rels`,
        data: encoder.encode(
          `${XML_DECL}<Relationships xmlns="${NS_PKG_REL}">${own
            .map(
              (t) =>
                `<Relationship Id="${t.relId}" Type="${NS_REL}/table" Target="../tables/${t.path.split("/").pop()}"/>`,
            )
            .join("")}</Relationships>`,
        ),
      })),
    ...(pictures
      ? [
          { name: "xl/metadata.xml", data: encoder.encode(pictures.metadataXml) },
          {
            name: "xl/richData/rdrichvalue.xml",
            data: encoder.encode(pictures.richValueXml),
          },
          {
            name: "xl/richData/rdrichvaluestructure.xml",
            data: encoder.encode(RICH_VALUE_STRUCTURE),
          },
          {
            name: "xl/richData/rdRichValueWebImage.xml",
            data: encoder.encode(pictures.webImageXml),
          },
          {
            name: "xl/richData/_rels/rdRichValueWebImage.xml.rels",
            data: encoder.encode(pictures.webImageRels),
          },
          ...pictures.media.map((m) => ({ name: m.path, data: m.bytes })),
        ]
      : []),
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
