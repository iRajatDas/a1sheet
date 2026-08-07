/**
 * The package parts beyond cells and styles: tables, conditional formats, and
 * in-cell images.
 *
 * Each of these lives outside the worksheet's `<sheetData>` and needs its own
 * plumbing — a part, a relationship, and a content-type override — which is why
 * they are here rather than threaded through `write.ts` a field at a time.
 *
 * Reading them without writing them is a quiet data loss: a workbook imported,
 * edited, and exported would come back with its tables flattened to plain cells
 * and its rules and pictures gone.
 */
import { colToLetters } from "../../model/address.js";
import type {
  CellImage,
  CellKey,
  CondFormat,
  CondRule,
  CondScalePoint,
  Range,
  SheetTable,
  StyleObject,
} from "../../model/types.js";
import { colorToRgb, xmlEscape } from "./xml.js";

const NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const NS_REL =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships";
const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

/** `A1:C9`, the form every part here uses to name a region. */
export function rangeRef(r: Range): string {
  return `${colToLetters(r.c1)}${r.r1 + 1}:${colToLetters(r.c2)}${r.r2 + 1}`;
}

// --- tables ---

export interface TablePart {
  /** `xl/tables/tableN.xml`. */
  path: string;
  xml: string;
  /** Relationship id within the owning worksheet's rels. */
  relId: string;
}

/**
 * One `<table>` part.
 *
 * `id` must be unique across the workbook, not the sheet — Excel rejects the file
 * outright on a duplicate, with an error that names nothing useful.
 */
export function tableXml(table: SheetTable, id: number): string {
  const columns = table.columns
    .map(
      (name, i) =>
        `<tableColumn id="${i + 1}" name="${xmlEscape(name || `Column${i + 1}`)}"/>`,
    )
    .join("");

  return (
    `${XML_DECL}<table xmlns="${NS_MAIN}" id="${id}"` +
    ` name="${xmlEscape(table.name)}" displayName="${xmlEscape(table.name)}"` +
    ` ref="${rangeRef(table.range)}"` +
    ` headerRowCount="${table.headerRow ? 1 : 0}" totalsRowShown="0">` +
    `<autoFilter ref="${rangeRef(table.range)}"/>` +
    `<tableColumns count="${table.columns.length}">${columns}</tableColumns>` +
    // A style has to be named or Excel draws the table unstyled. The cells carry
    // their appearance already — flattened there on import — so the least
    // surprising choice is the neutral one rather than a colour nobody picked.
    `<tableStyleInfo name="TableStyleLight1" showFirstColumn="0"` +
    ` showLastColumn="0" showRowStripes="0" showColumnStripes="0"/>` +
    `</table>`
  );
}

// --- conditional formatting ---

const OPERATORS: Record<string, string> = {
  lessThan: "lessThan",
  lessThanOrEqual: "lessThanOrEqual",
  equal: "equal",
  notEqual: "notEqual",
  greaterThanOrEqual: "greaterThanOrEqual",
  greaterThan: "greaterThan",
  between: "between",
  notBetween: "notBetween",
};

/** Points and stops, as `<cfvo>` and `<color>`. */
function scalePoints(points: readonly CondScalePoint[]): string {
  return points
    .map(
      (p) =>
        `<cfvo type="${p.kind}"${p.value === undefined ? "" : ` val="${xmlEscape(p.value)}"`}/>`,
    )
    .join("");
}

/** The body a graphical rule carries instead of a formula. */
function graphicalBody(rule: CondRule): string {
  if (rule.type === "colorScale") {
    return (
      `<colorScale>${scalePoints(rule.stops)}` +
      `${rule.stops.map((s) => `<color rgb="${colorToRgb(s.color)}"/>`).join("")}` +
      `</colorScale>`
    );
  }
  if (rule.type === "dataBar") {
    return (
      `<dataBar>${scalePoints([rule.min, rule.max])}` +
      `<color rgb="${colorToRgb(rule.color)}"/></dataBar>`
    );
  }
  if (rule.type === "iconSet") {
    return `<iconSet iconSet="${xmlEscape(rule.set)}">${scalePoints(rule.thresholds)}</iconSet>`;
  }
  return "";
}

function isGraphical(rule: CondRule): boolean {
  return (
    rule.type === "colorScale" || rule.type === "dataBar" || rule.type === "iconSet"
  );
}

function ruleAttrs(rule: CondRule): string {
  switch (rule.type) {
    case "expression":
      return ` type="expression"`;
    case "cellIs":
      return ` type="cellIs" operator="${OPERATORS[rule.operator] ?? "equal"}"`;
    case "containsText":
      return (
        ` type="${rule.negate ? "notContainsText" : "containsText"}"` +
        ` operator="${rule.negate ? "notContains" : "containsText"}"` +
        ` text="${xmlEscape(rule.text)}"`
      );
    case "containsBlanks":
      return ` type="${rule.negate ? "notContainsBlanks" : "containsBlanks"}"`;
    case "top10":
      return (
        ` type="top10" rank="${rule.rank}"` +
        `${rule.bottom ? ` bottom="1"` : ""}${rule.percent ? ` percent="1"` : ""}`
      );
    case "aboveAverage":
      return (
        ` type="aboveAverage"` +
        `${rule.below ? ` aboveAverage="0"` : ""}` +
        `${rule.orEqual ? ` equalAverage="1"` : ""}`
      );
    default:
      return ` type="${rule.type}"`;
  }
}

function ruleFormulas(rule: CondRule, range: Range): string {
  if (isGraphical(rule)) return graphicalBody(rule);
  if (rule.type === "expression") {
    return `<formula>${xmlEscape(rule.formula)}</formula>`;
  }
  if (rule.type === "cellIs") {
    return rule.operands.map((f) => `<formula>${xmlEscape(f)}</formula>`).join("");
  }
  // The text and blank rules carry a formula Excel uses when it recalculates the
  // rule itself; it must reference the range's top-left cell.
  const anchor = `${colToLetters(range.c1)}${range.r1 + 1}`;
  if (rule.type === "containsText") {
    const test = `NOT(ISERROR(SEARCH("${rule.text.replace(/"/g, '""')}",${anchor})))`;
    return `<formula>${xmlEscape(rule.negate ? `ISERROR(SEARCH("${rule.text}",${anchor}))` : test)}</formula>`;
  }
  if (rule.type !== "containsBlanks") return "";
  const blank = rule.negate ? `LEN(TRIM(${anchor}))>0` : `LEN(TRIM(${anchor}))=0`;
  return `<formula>${xmlEscape(blank)}</formula>`;
}

export interface CondFormatXml {
  /** `<conditionalFormatting>` blocks for the worksheet. */
  sheet: string;
  /** Styles to append to `<dxfs>`, in the order the rules index them. */
  dxfs: StyleObject[];
}

/**
 * Conditional formats as worksheet XML plus the differential styles they point
 * at. `firstDxfId` is where this sheet's styles start in the workbook-wide list.
 */
export function condFormatXml(
  formats: readonly CondFormat[],
  firstDxfId: number,
): CondFormatXml {
  const dxfs: StyleObject[] = [];
  const blocks = formats
    .map((format) => {
      // A graphical rule carries its own colours and needs no dxf; giving it one
      // would leave an empty entry that shifts every later rule's index.
      const graphical = isGraphical(format.rule);
      const dxfId = graphical ? -1 : firstDxfId + dxfs.length;
      if (!graphical) dxfs.push(format.style);
      return (
        `<conditionalFormatting sqref="${rangeRef(format.range)}">` +
        `<cfRule${ruleAttrs(format.rule)}${graphical ? "" : ` dxfId="${dxfId}"`}` +
        ` priority="${format.priority}"` +
        `${format.stopIfTrue ? ` stopIfTrue="1"` : ""}>` +
        `${ruleFormulas(format.rule, format.range)}</cfRule>` +
        `</conditionalFormatting>`
      );
    })
    .join("");

  return { sheet: blocks, dxfs };
}

const EDGES = ["left", "right", "top", "bottom"] as const;

/**
 * A `<dxf>` for a partial style, which states its parts inline rather than by
 * index.
 *
 * Must cover everything a rule can set, including gradients: a differential
 * format that writes only its font and solid fill comes back as an empty `<dxf>`,
 * and an empty one is indistinguishable from no style at all — so the rule that
 * pointed at it is dropped on re-read.
 */
export function dxfXml(style: StyleObject): string {
  const font =
    style.bold || style.italic || style.underline || style.color
      ? `<font>${style.bold ? "<b/>" : ""}${style.italic ? "<i/>" : ""}` +
        `${style.underline ? "<u/>" : ""}` +
        `${style.color ? `<color rgb="${colorToRgb(style.color)}"/>` : ""}</font>`
      : "";

  let fill = "";
  if (style.gradient) {
    const stops = [...style.gradient.stops]
      .sort((a, b) => a.position - b.position)
      .map(
        (stop) =>
          `<stop position="${stop.position}">` +
          `<color rgb="${colorToRgb(stop.color)}"/></stop>`,
      )
      .join("");
    fill = `<fill><gradientFill degree="${style.gradient.degree}">${stops}</gradientFill></fill>`;
  } else if (style.bg) {
    // A dxf's solid fill states its colour as bgColor, not fgColor — the reverse
    // of a cellXf's, and a long-standing trap in the format.
    fill = `<fill><patternFill patternType="solid"><bgColor rgb="${colorToRgb(style.bg)}"/></patternFill></fill>`;
  }

  const borders = style.borders;
  const border = borders
    ? `<border>${EDGES.map((edge) => {
        const set = borders[edge];
        if (!set) return `<${edge}/>`;
        const color = set.color
          ? `<color rgb="${colorToRgb(set.color)}"/>`
          : `<color auto="1"/>`;
        return `<${edge} style="${set.line}">${color}</${edge}>`;
      }).join("")}</border>`
    : "";

  const alignment =
    style.align || style.valign || style.wrap
      ? `<alignment${style.align ? ` horizontal="${style.align}"` : ""}` +
        `${style.valign ? ` vertical="${style.valign === "middle" ? "center" : style.valign}"` : ""}` +
        `${style.wrap ? ` wrapText="1"` : ""}/>`
      : "";

  return `<dxf>${font}${fill}${border}${alignment}</dxf>`;
}

// --- in-cell images ---

const DATA_URI = /^data:([^;]+);base64,(.*)$/;

/** One cell's image, tagged with the sheet it is on. */
export interface SheetImage {
  sheet: number;
  key: CellKey;
  image: CellImage;
}

export interface ImageParts {
  /**
   * `vm` attribute value per cell, 1-based into the value metadata.
   *
   * Keyed by sheet AND cell: a cell key is only unique within a sheet, so
   * indexing by it alone gave every sheet the first sheet's pictures.
   */
  valueMetadata: Map<string, number>;
  /** Media parts to add to the package. */
  media: { path: string; bytes: Uint8Array }[];
  /** Extensions used, so `[Content_Types]` can declare them. */
  extensions: Set<string>;
  metadataXml: string;
  richValueXml: string;
  webImageXml: string;
  webImageRels: string;
}

const MIME_EXTENSIONS: Record<string, string> = {
  "image/png": "png",
  "image/jpeg": "jpeg",
  "image/gif": "gif",
  "image/webp": "webp",
  "image/bmp": "bmp",
};

function decodeBase64(text: string): Uint8Array {
  const binary = atob(text);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes;
}

/**
 * Builds the five parts an in-cell image needs.
 *
 * Excel's chain is vm -> valueMetadata -> futureMetadata -> rdrichvalue ->
 * webImageSrd -> rels -> media, and every link is by index, so the parts have to
 * be built together. An image given only as a URL gets an `address` relationship
 * and no embedded copy; one given as a data URI gets both, since the URL is what
 * a reader without the media part falls back to.
 */
export function imageParts(images: readonly SheetImage[]): ImageParts | null {
  if (images.length === 0) return null;

  const valueMetadata = new Map<string, number>();
  const media: { path: string; bytes: Uint8Array }[] = [];
  const extensions = new Set<string>();
  const rels: string[] = [];
  const srds: string[] = [];
  let relId = 0;
  // One rich value per distinct picture, not per cell. A workbook typically
  // points many cells at the same image — a hundred and forty at twenty in the
  // sample file — and a part each would multiply the file size sevenfold.
  const byImage = new Map<string, number>();

  for (const { sheet, key, image } of images) {
    const cell = imageKey(sheet, key);
    const identity = `${image.src}\u0000${image.alt ?? ""}`;
    const seen = byImage.get(identity);
    if (seen !== undefined) {
      valueMetadata.set(cell, seen);
      continue;
    }
    const index = srds.length;
    byImage.set(identity, index + 1);
    valueMetadata.set(cell, index + 1);

    const parts: string[] = [];
    const address = image.alt ?? (image.src.startsWith("data:") ? null : image.src);
    if (address) {
      relId++;
      rels.push(
        `<Relationship Id="rId${relId}" Type="${NS_REL}/hyperlink"` +
          ` Target="${xmlEscape(address)}" TargetMode="External"/>`,
      );
      parts.push(`<address r:id="rId${relId}"/>`);
    }

    const embedded = image.src.match(DATA_URI);
    const extension = embedded ? MIME_EXTENSIONS[embedded[1] as string] : undefined;
    if (embedded && extension) {
      const path = `xl/media/image${media.length + 1}.${extension}`;
      media.push({ path, bytes: decodeBase64(embedded[2] as string) });
      extensions.add(extension);
      relId++;
      rels.push(
        `<Relationship Id="rId${relId}" Type="${NS_REL}/image"` +
          ` Target="../media/${path.split("/").pop()}"/>`,
      );
      parts.push(`<blip r:id="rId${relId}"/>`);
    }

    srds.push(`<webImageSrd>${parts.join("")}</webImageSrd>`);
  }

  const count = srds.length;
  const metadataXml =
    `${XML_DECL}<metadata xmlns="${NS_MAIN}"` +
    ` xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">` +
    `<metadataTypes count="1"><metadataType name="XLRICHVALUE"` +
    ` minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1"` +
    ` merge="1" splitFirst="1" rowColShift="1" clearFormats="1"` +
    ` clearComments="1" assign="1" coerce="1"/></metadataTypes>` +
    `<futureMetadata name="XLRICHVALUE" count="${count}">` +
    Array.from(
      { length: count },
      (_, i) =>
        `<bk><extLst><ext uri="{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}">` +
        `<xlrd:rvb i="${i}"/></ext></extLst></bk>`,
    ).join("") +
    `</futureMetadata>` +
    `<valueMetadata count="${count}">` +
    Array.from({ length: count }, (_, i) => `<bk><rc t="1" v="${i}"/></bk>`).join(
      "",
    ) +
    `</valueMetadata></metadata>`;

  const richValueXml =
    `${XML_DECL}<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"` +
    ` count="${count}">` +
    Array.from(
      { length: count },
      (_, i) => `<rv s="0"><v>${i}</v><v>1</v><v>0</v><v>0</v></rv>`,
    ).join("") +
    `</rvData>`;

  const webImageXml =
    `${XML_DECL}<webImagesSrd` +
    ` xmlns="http://schemas.microsoft.com/office/spreadsheetml/2020/richdatawebimage"` +
    ` xmlns:r="${NS_REL}">${srds.join("")}</webImagesSrd>`;

  const webImageRels = `${XML_DECL}<Relationships xmlns="${NS_PKG_REL}">${rels.join("")}</Relationships>`;

  return {
    valueMetadata,
    media,
    extensions,
    metadataXml,
    richValueXml,
    webImageXml,
    webImageRels,
  };
}

/** The key `valueMetadata` uses, since a cell key repeats across sheets. */
export function imageKey(sheet: number, key: CellKey): string {
  return `${sheet}:${key}`;
}

/** The structure record the rich values point at — one shape, `_webimage`. */
export const RICH_VALUE_STRUCTURE =
  `${XML_DECL}<rvStructures` +
  ` xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"` +
  ` count="1"><s t="_webimage">` +
  `<k n="WebImageIdentifier" t="i"/><k n="CalcOrigin" t="i"/>` +
  `<k n="ComputedImage" t="b"/><k n="ImageSizing" t="i"/></s></rvStructures>`;
