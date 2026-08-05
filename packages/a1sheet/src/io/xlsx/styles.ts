/**
 * styles.xml read support and the number-format maps.
 * Ported from ref/xlsxIO.js:198-204 and 305-344.
 *
 * A single style table is shared across ALL sheets in the workbook, deduped by
 * `styleKey`. Borders are written as one empty `<border/>` placeholder — see
 * docs/LIMITATIONS.md.
 */
import type { Align, NumFmt, StyleObject } from "../../model/types.js";
import { findElement, findElements, hasElement, rgbToColor } from "./xml.js";

export const NUMFMT_TO_ID: Record<NumFmt, number> = {
  general: 0,
  integer: 1,
  number: 2,
  percent: 10,
  currency: 164,
  date: 14,
};

/** Custom formats we define in styles.xml, keyed by the id we assign. */
export const CUSTOM_NUMFMTS: Record<number, string> = {
  164: '"$"#,##0.00',
};

/** Dedup key for the shared style table. Ported from ref/xlsxIO.js:202-204. */
export function styleKey(style: StyleObject | undefined): string {
  if (!style) return "default";
  return JSON.stringify([
    !!style.bold,
    !!style.italic,
    !!style.underline,
    style.color ?? "",
    style.bg ?? "",
    style.numFmt ?? "general",
    style.align ?? "",
    !!style.locked,
  ]);
}

/**
 * Buckets an OOXML numFmt id (and its format code, for custom ids) into one of our
 * six enum values.
 *
 * Heuristic by design: arbitrary format strings collapse to the nearest supported
 * bucket rather than round-tripping exactly.
 */
export function numFmtToKey(
  id: string,
  customCodes: Record<string, string> = {},
): NumFmt {
  const n = parseInt(id, 10);
  if (n === 0) return "general";
  if (n === 1) return "integer";
  if (n === 2) return "number";
  if (n === 9 || n === 10) return "percent";
  if (n >= 14 && n <= 22) return "date";

  const code = customCodes[id] ?? "";
  if (code.includes("$") || code.includes("€") || code.includes("£")) {
    return "currency";
  }
  if (code.includes("%")) return "percent";
  if (/[ymd]/i.test(code) && (code.includes("/") || code.includes("-"))) {
    return "date";
  }
  return "general";
}

/**
 * Parses styles.xml, mapping cellXfs indices to StyleObjects.
 *
 * Returns `null` at an index whose style is the default/no-op, so cells without
 * real formatting do not get a spurious `styles` entry.
 */
export function parseStylesXml(xml: string | undefined): (StyleObject | null)[] {
  if (!xml) return [];

  const customCodes: Record<string, string> = {};
  for (const n of findElements(xml, "numFmt")) {
    const id = n.attrs.numFmtId;
    const code = n.attrs.formatCode;
    if (id && code) customCodes[id] = code;
  }

  const fontsBlock = findElement(xml, "fonts");
  const fonts = (fontsBlock ? findElements(fontsBlock.inner, "font") : []).map(
    (f) => ({
      bold: hasElement(f.inner, "b"),
      italic: hasElement(f.inner, "i"),
      underline: hasElement(f.inner, "u"),
      color: rgbToColor(findElement(f.inner, "color")?.attrs.rgb),
    }),
  );

  const fillsBlock = findElement(xml, "fills");
  const fills = (fillsBlock ? findElements(fillsBlock.inner, "fill") : []).map(
    (f) => ({
      bg: rgbToColor(findElement(f.inner, "fgColor")?.attrs.rgb),
    }),
  );

  const xfsBlock = findElement(xml, "cellXfs");
  if (!xfsBlock) return [];

  type ParsedFont = (typeof fonts)[number];
  type ParsedFill = (typeof fills)[number];
  const NO_FONT: ParsedFont = {
    bold: false,
    italic: false,
    underline: false,
    color: null,
  };
  const NO_FILL: ParsedFill = { bg: null };

  return findElements(xfsBlock.inner, "xf").map((xf) => {
    const font = fonts[parseInt(xf.attrs.fontId ?? "0", 10)] ?? NO_FONT;
    const fill = fills[parseInt(xf.attrs.fillId ?? "0", 10)] ?? NO_FILL;
    const align = findElement(xf.inner, "alignment")?.attrs.horizontal ?? "";
    const numFmt = numFmtToKey(xf.attrs.numFmtId ?? "0", customCodes);

    const style: StyleObject = {};
    if (font.bold) style.bold = true;
    if (font.italic) style.italic = true;
    if (font.underline) style.underline = true;
    if (font.color) style.color = font.color;
    if (fill.bg) style.bg = fill.bg;
    if (numFmt !== "general") style.numFmt = numFmt;
    if (align === "left" || align === "center" || align === "right") {
      style.align = align as Align;
    }

    return Object.keys(style).length === 0 ? null : style;
  });
}
