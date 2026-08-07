/**
 * Emitting styles.xml.
 *
 * The four tables it contains — fonts, fills, borders, number formats — are each
 * deduped independently and referenced by index from `cellXfs`, which is what the
 * format requires and also what keeps the part small when a thousand cells share
 * one appearance.
 *
 * Split out of write.ts because that file's job is the package and this one's is
 * the style table; together they were one function doing two things.
 */
import type { GradientFill, StyleObject } from "../../model/types.js";
import { CUSTOM_NUMFMTS, NUMFMT_TO_ID } from "./styles.js";
import { colorToRgb, xmlEscape } from "./xml.js";

/** Points, since `<sz>` is in points. The default Excel writes for body text. */
const DEFAULT_FONT_POINTS = 11;
const POINTS_PER_INCH = 72;
const CSS_PIXELS_PER_INCH = 96;
const DEFAULT_FONT_FAMILY = "Calibri";

/**
 * Where our own format codes start. Below `FIRST_CUSTOM_ID` the ids are the
 * format's built-ins, which carry no `formatCode`; a style whose `numFmtCode`
 * came from a file gets a fresh id here so the code survives verbatim rather
 * than being flattened back into one of six buckets.
 */
const FIRST_CUSTOM_ID = 164;

function pxToPoints(px: number): number {
  return Math.round((px * POINTS_PER_INCH) / CSS_PIXELS_PER_INCH);
}

interface Table<T> {
  entries: T[];
  index: Map<string, number>;
}

function table<T>(initial: T[] = []): Table<T> {
  return { entries: initial, index: new Map() };
}

/** Interns a value by key, returning its index. */
function intern<T>(t: Table<T>, key: string, make: () => T): number {
  const existing = t.index.get(key);
  if (existing !== undefined) return existing;
  const id = t.entries.length;
  t.entries.push(make());
  t.index.set(key, id);
  return id;
}

function fontXml(s: StyleObject): string {
  const points =
    s.fontSize === undefined ? DEFAULT_FONT_POINTS : pxToPoints(s.fontSize);
  return (
    `<font>${s.bold ? "<b/>" : ""}${s.italic ? "<i/>" : ""}` +
    `${s.underline ? "<u/>" : ""}<sz val="${points}"/>` +
    `${s.color ? `<color rgb="${colorToRgb(s.color)}"/>` : `<color theme="1"/>`}` +
    `<name val="${xmlEscape(s.fontFamily ?? DEFAULT_FONT_FAMILY)}"/></font>`
  );
}

const GRADIENT_STOP_MAX = 1;

function gradientXml(gradient: GradientFill): string {
  const stops = [...gradient.stops]
    .sort((a, b) => a.position - b.position)
    .map(
      (stop) =>
        `<stop position="${Math.max(0, Math.min(GRADIENT_STOP_MAX, stop.position))}">` +
        `<color rgb="${colorToRgb(stop.color)}"/></stop>`,
    )
    .join("");
  return `<fill><gradientFill degree="${gradient.degree}">${stops}</gradientFill></fill>`;
}

function fillXml(s: StyleObject): string {
  if (s.gradient) return gradientXml(s.gradient);
  if (s.bg) {
    return `<fill><patternFill patternType="solid"><fgColor rgb="${colorToRgb(s.bg)}"/></patternFill></fill>`;
  }
  return `<fill><patternFill patternType="none"/></fill>`;
}

const EDGE_TAGS = [
  ["left", "left"],
  ["right", "right"],
  ["top", "top"],
  ["bottom", "bottom"],
] as const;

function borderXml(s: StyleObject): string {
  const borders = s.borders;
  const edges = EDGE_TAGS.map(([edge, tag]) => {
    const set = borders?.[edge];
    if (!set) return `<${tag}/>`;
    const color = set.color
      ? `<color rgb="${colorToRgb(set.color)}"/>`
      : `<color auto="1"/>`;
    return `<${tag} style="${set.line}">${color}</${tag}>`;
  }).join("");
  return `<border>${edges}<diagonal/></border>`;
}

function alignmentXml(s: StyleObject): string {
  const parts: string[] = [];
  if (s.align) parts.push(`horizontal="${s.align}"`);
  if (s.valign) {
    parts.push(`vertical="${s.valign === "middle" ? "center" : s.valign}"`);
  }
  if (s.wrap) parts.push(`wrapText="1"`);
  if (s.indent) parts.push(`indent="${s.indent}"`);
  if (s.rotation) {
    // Clockwise rotation is written as 90 + the angle; anticlockwise as itself.
    parts.push(`textRotation="${s.rotation < 0 ? 90 - s.rotation : s.rotation}"`);
  }
  return parts.length ? `<alignment ${parts.join(" ")}/>` : "";
}

/** Everything that makes two styles share a font, a fill, or a border. */
const keyOf = {
  font: (s: StyleObject) =>
    JSON.stringify([
      !!s.bold,
      !!s.italic,
      !!s.underline,
      s.color ?? "",
      s.fontFamily ?? "",
      s.fontSize ?? 0,
    ]),
  fill: (s: StyleObject) => JSON.stringify([s.bg ?? "", s.gradient ?? ""]),
  border: (s: StyleObject) => JSON.stringify(s.borders ?? ""),
};

/**
 * @param dxfsXml Differential formats, already serialized. They belong to
 * conditional formatting and to table styling, both of which index into this
 * block by position — so it is passed in whole rather than rebuilt here.
 */
export function buildStylesXml(
  styleList: readonly StyleObject[],
  dxfsXml = "",
): string {
  const fonts = table<string>();
  // The two reserved fills, written verbatim so their indices stay 0 and 1.
  const fills = table<string>([
    `<fill><patternFill patternType="none"/></fill>`,
    `<fill><patternFill patternType="gray125"/></fill>`,
  ]);
  // Index 0 must be the empty border: a cellXf with borderId="0" means no border,
  // and every style without one points at it.
  const borders = table<string>([
    `<border><left/><right/><top/><bottom/><diagonal/></border>`,
  ]);
  const numFmts = table<string>();

  const xfs = styleList.map((s) => {
    const fontId = intern(fonts, keyOf.font(s), () => fontXml(s));
    const hasFill = Boolean(s.bg || s.gradient);
    const fillId = hasFill ? intern(fills, keyOf.fill(s), () => fillXml(s)) : 0;
    const borderId = s.borders
      ? intern(borders, keyOf.border(s), () => borderXml(s))
      : 0;

    // A code from a file wins over the bucket, so a format like "+0;-0;0" comes
    // back out as itself rather than as the nearest of our six.
    const numFmtId = s.numFmtCode
      ? FIRST_CUSTOM_ID +
        Object.keys(CUSTOM_NUMFMTS).length +
        intern(numFmts, s.numFmtCode, () => s.numFmtCode as string)
      : (NUMFMT_TO_ID[s.numFmt ?? "general"] ?? 0);

    return { fontId, fillId, borderId, numFmtId, style: s };
  });

  const customNumFmtsXml = Object.entries(CUSTOM_NUMFMTS)
    .map(
      ([id, code]) => `<numFmt numFmtId="${id}" formatCode="${xmlEscape(code)}"/>`,
    )
    .join("");
  const readNumFmtsXml = numFmts.entries
    .map(
      (code, i) =>
        `<numFmt numFmtId="${FIRST_CUSTOM_ID + Object.keys(CUSTOM_NUMFMTS).length + i}" formatCode="${xmlEscape(code)}"/>`,
    )
    .join("");
  const numFmtCount = Object.keys(CUSTOM_NUMFMTS).length + numFmts.entries.length;

  const xfsXml = xfs
    .map(({ fontId, fillId, borderId, numFmtId, style }) => {
      const alignment = alignmentXml(style);
      const flags =
        `${alignment ? ` applyAlignment="1"` : ""}` +
        `${borderId ? ` applyBorder="1"` : ""}` +
        `${style.locked ? ` applyProtection="1"` : ""}`;
      const body = `${alignment}${style.locked ? `<protection locked="1"/>` : ""}`;
      return (
        `<xf numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}"` +
        ` borderId="${borderId}" xfId="0" applyFont="1" applyFill="1"` +
        ` applyNumberFormat="1"${flags}>${body}</xf>`
      );
    })
    .join("");

  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
    `<numFmts count="${numFmtCount}">${customNumFmtsXml}${readNumFmtsXml}</numFmts>` +
    `<fonts count="${fonts.entries.length}">${fonts.entries.join("")}</fonts>` +
    `<fills count="${fills.entries.length}">${fills.entries.join("")}</fills>` +
    `<borders count="${borders.entries.length}">${borders.entries.join("")}</borders>` +
    `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
    `<cellXfs count="${xfs.length}">${xfsXml}</cellXfs>` +
    // <dxfs> follows <cellXfs> in the schema; putting it earlier is rejected.
    (dxfsXml ? `<dxfs count="${countDxfs(dxfsXml)}">${dxfsXml}</dxfs>` : "") +
    `</styleSheet>`
  );
}

/** How many `<dxf>` elements a serialized block holds. */
function countDxfs(xml: string): number {
  return xml.split("<dxf>").length - 1;
}
