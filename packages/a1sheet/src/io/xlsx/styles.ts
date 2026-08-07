/**
 * styles.xml read support and the number-format maps.
 *
 * A single style table is shared across ALL sheets in the workbook, deduped by
 * `styleKey`.
 *
 * The read side resolves five things per `<xf>`: font, fill, border, number
 * format, and alignment. Each is an index into a separate table, and each index
 * may be inherited from a named cell style rather than stated — see
 * `resolveIds`.
 */
import type {
  Align,
  BorderEdge,
  BorderLine,
  CellBorders,
  GradientFill,
  HexColor,
  NumFmt,
  StyleObject,
  VerticalAlign,
} from "../../model/types.js";
import { EMPTY_PALETTE, resolveColorAttrs, type ThemePalette } from "./palette.js";
import { findElement, findElements, hasElement } from "./xml.js";

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

/**
 * The first id an application may use for its own format codes. Everything below
 * is reserved for the built-ins, which carry no `formatCode` in the file.
 */
const FIRST_CUSTOM_NUMFMT_ID = 164;

/**
 * The built-in format codes, by id. The file does not contain these — every
 * reader is expected to know them — so without this table a column formatted as
 * a plain date or a percentage has no code to render from.
 *
 * Only the ids that mean something distinct are listed; the rest fall back to
 * their bucket. Locale-dependent ids (14-17, 22) are given their en-US forms,
 * which is what Excel shows on a US machine and the closest single answer
 * available without knowing the reader's locale.
 */
const BUILTIN_NUMFMTS: Record<number, string> = {
  1: "0",
  2: "0.00",
  3: "#,##0",
  4: "#,##0.00",
  9: "0%",
  10: "0.00%",
  11: "0.00E+00",
  12: "# ?/?",
  13: "# ??/??",
  14: "mm-dd-yy",
  15: "d-mmm-yy",
  16: "d-mmm",
  17: "mmm-yy",
  18: "h:mm AM/PM",
  19: "h:mm:ss AM/PM",
  20: "h:mm",
  21: "h:mm:ss",
  22: "m/d/yy h:mm",
  37: "#,##0 ;(#,##0)",
  38: "#,##0 ;[Red](#,##0)",
  39: "#,##0.00;(#,##0.00)",
  40: "#,##0.00;[Red](#,##0.00)",
  45: "mm:ss",
  46: "[h]:mm:ss",
  47: "mmss.0",
  48: "##0.0E+0",
  49: "@",
};

/** Dedup key for the shared style table. */
export function styleKey(style: StyleObject | undefined): string {
  if (!style) return "default";
  return JSON.stringify([
    !!style.bold,
    !!style.italic,
    !!style.underline,
    style.color ?? "",
    style.bg ?? "",
    style.gradient ?? "",
    style.numFmt ?? "general",
    style.numFmtCode ?? "",
    style.indent ?? 0,
    style.rotation ?? 0,
    style.align ?? "",
    style.valign ?? "",
    !!style.wrap,
    style.borders ?? "",
    style.fontFamily ?? "",
    style.fontSize ?? 0,
    !!style.locked,
  ]);
}

/**
 * Buckets an OOXML numFmt id (and its format code, for custom ids) into one of our
 * six enum values.
 *
 * Heuristic by design: this is what the format dropdown displays. The literal
 * code travels alongside in `numFmtCode` and is what actually renders.
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
  if (n >= 45 && n <= 47) return "date";

  const code = customCodes[id] ?? "";
  if (code.includes("$") || code.includes("€") || code.includes("£")) {
    return "currency";
  }
  if (code.includes("%")) return "percent";
  if (/[ymd]/i.test(code) && (code.includes("/") || code.includes("-"))) {
    return "date";
  }
  if (/^\[?h/i.test(code) || /h+:mm/i.test(code)) return "date";
  return "general";
}

/** `textRotation="255"` is Excel's stacked-vertical mode rather than an angle. */
const VERTICAL_TEXT = 255;
const QUARTER_TURN = 90;

const BORDER_LINES: readonly BorderLine[] = [
  "hair",
  "thin",
  "medium",
  "thick",
  "double",
  "dotted",
  "dashed",
  "dashDot",
  "dashDotDot",
  "mediumDashed",
  "mediumDashDot",
  "mediumDashDotDot",
  "slantDashDot",
];

function toBorderLine(style: string | undefined): BorderLine | null {
  if (!style || style === "none") return null;
  return BORDER_LINES.find((l) => l === style) ?? "thin";
}

const EDGES = ["top", "right", "bottom", "left"] as const;

/** OOXML names the horizontal edges the same but calls left/right start/end too. */
const EDGE_TAGS: Record<(typeof EDGES)[number], readonly string[]> = {
  top: ["top"],
  right: ["right", "end"],
  bottom: ["bottom"],
  left: ["left", "start"],
};

function parseBorder(inner: string, palette: ThemePalette): CellBorders | null {
  const borders: CellBorders = {};
  for (const edge of EDGES) {
    for (const tag of EDGE_TAGS[edge]) {
      const el = findElement(inner, tag);
      if (!el) continue;
      const line = toBorderLine(el.attrs.style);
      if (!line) continue;
      const colorEl = findElement(el.inner, "color");
      const color = colorEl ? resolveColorAttrs(colorEl.attrs, palette) : null;
      const parsed: BorderEdge = color ? { line, color } : { line };
      borders[edge] = parsed;
      break;
    }
  }
  return Object.keys(borders).length > 0 ? borders : null;
}

interface ParsedFill {
  bg: HexColor | null;
  gradient: GradientFill | null;
}

const NO_FILL: ParsedFill = { bg: null, gradient: null };

/** OOXML's default gradient sweep, used when `<gradientFill>` omits `degree`. */
const DEFAULT_GRADIENT_DEGREE = 0;

function parseFill(inner: string, palette: ThemePalette): ParsedFill {
  const gradientEl = findElement(inner, "gradientFill");
  if (gradientEl) {
    const stops = findElements(gradientEl.inner, "stop")
      .map((stop) => {
        const colorEl = findElement(stop.inner, "color");
        const color = colorEl ? resolveColorAttrs(colorEl.attrs, palette) : null;
        const position = Number.parseFloat(stop.attrs.position ?? "0");
        return color ? { position: position || 0, color } : null;
      })
      .filter((s): s is { position: number; color: HexColor } => s !== null);

    if (stops.length > 0) {
      const degree = Number.parseFloat(gradientEl.attrs.degree ?? "");
      return {
        // The first stop doubles as a flat fallback wherever a gradient cannot
        // be drawn — CSV export, or a consumer reading `bg` directly.
        bg: stops[0]?.color ?? null,
        gradient: {
          degree: Number.isFinite(degree) ? degree : DEFAULT_GRADIENT_DEGREE,
          stops,
        },
      };
    }
  }

  const pattern = findElement(inner, "patternFill");
  if (!pattern) return NO_FILL;
  // "none" is the empty fill every workbook has at index 0 and "gray125" the
  // second one Excel always writes. Neither is a colour — but a DIFFERENTIAL
  // format routinely omits patternType entirely and still means a fill, so the
  // absence of the attribute cannot be treated as "none".
  if (pattern.attrs.patternType === "none") return NO_FILL;

  const fgEl = findElement(pattern.inner, "fgColor");
  const bgEl = findElement(pattern.inner, "bgColor");
  const fg = fgEl ? resolveColorAttrs(fgEl.attrs, palette) : null;
  const behind = bgEl ? resolveColorAttrs(bgEl.attrs, palette) : null;

  // A cell style's solid fill names its colour fgColor; a differential format's
  // names it bgColor. The same fill, the opposite attribute — a long-standing
  // trap in the format, and the reason a conditional format's colour vanished
  // while the identical cell fill survived. Take whichever is there.
  if (pattern.attrs.patternType === "solid" || !pattern.attrs.patternType) {
    return { bg: fg ?? behind, gradient: null };
  }
  // A real pattern puts the pattern's colour in fgColor and the surface behind
  // it in bgColor; since the pattern itself is not drawn, the surface is the
  // better single answer.
  return { bg: behind ?? fg, gradient: null };
}

interface ParsedFont {
  bold: boolean;
  italic: boolean;
  underline: boolean;
  color: HexColor | null;
  family: string | null;
  size: number | null;
}

const NO_FONT: ParsedFont = {
  bold: false,
  italic: false,
  underline: false,
  color: null,
  family: null,
  size: null,
};

const POINTS_PER_INCH = 72;
const CSS_PIXELS_PER_INCH = 96;

/** Font sizes are in points in the file and in pixels everywhere in the DOM. */
function pointsToPx(points: number): number {
  return Math.round((points * CSS_PIXELS_PER_INCH) / POINTS_PER_INCH);
}

function parseFont(inner: string, palette: ThemePalette): ParsedFont {
  const colorEl = findElement(inner, "color");
  const size = Number.parseFloat(findElement(inner, "sz")?.attrs.val ?? "");
  const name = findElement(inner, "name")?.attrs.val;
  return {
    // `<b val="0"/>` is how a theme's font turns bold OFF again, so the presence
    // of the tag is not enough — an explicit "0" means not bold.
    bold: isOn(inner, "b"),
    italic: isOn(inner, "i"),
    underline: isOn(inner, "u"),
    color: colorEl ? resolveColorAttrs(colorEl.attrs, palette) : null,
    family: name ?? null,
    size: Number.isFinite(size) ? pointsToPx(size) : null,
  };
}

/**
 * Whether a boolean font property is switched on.
 *
 * The tag being present is not enough. A differential format turns a property
 * OFF by stating it: `<b val="0"/>`, and for underline `<u val="none"/>`, since
 * `u` carries a style name rather than a flag. Treating presence as truth made
 * every table header underlined.
 */
const OFF_VALUES = new Set(["0", "false", "none"]);

function isOn(inner: string, tag: string): boolean {
  if (!hasElement(inner, tag)) return false;
  const val = findElement(inner, tag)?.attrs.val;
  return val === undefined || !OFF_VALUES.has(val);
}

interface XfIds {
  fontId: number;
  fillId: number;
  borderId: number;
  numFmtId: number;
  /** Raw inner XML of whichever `<xf>` supplied the alignment. */
  alignmentFrom: string;
}

function idOf(attrs: Record<string, string>, name: string): number {
  const n = Number.parseInt(attrs[name] ?? "0", 10);
  return Number.isFinite(n) ? n : 0;
}

/**
 * Works out which font/fill/border/format a `<cellXfs>` entry really uses.
 *
 * A cellXf may inherit from a named cell style through its `xfId`, and Excel
 * usually omits the `applyFont`/`applyBorder` flags that are supposed to say so.
 * The rule that matches observed files: a nonzero id on the cellXf wins, and a
 * zero id falls through to the named style. Index 0 of every table is the
 * "nothing set" entry, so falling through on zero loses nothing and recovers the
 * borders and alignment that titles and headings carry via named styles.
 */
function resolveIds(
  xf: { attrs: Record<string, string>; inner: string },
  styleXfs: readonly { attrs: Record<string, string>; inner: string }[],
): XfIds {
  const parent = styleXfs[idOf(xf.attrs, "xfId")];
  const pick = (name: string) =>
    idOf(xf.attrs, name) || (parent ? idOf(parent.attrs, name) : 0);
  return {
    fontId: pick("fontId"),
    fillId: pick("fillId"),
    borderId: pick("borderId"),
    numFmtId: pick("numFmtId"),
    alignmentFrom: findElement(xf.inner, "alignment")
      ? xf.inner
      : (parent?.inner ?? ""),
  };
}

function toAlign(value: string | undefined): Align | undefined {
  if (value === "left" || value === "center" || value === "right") return value;
  // Excel's "general" is left for text and right for numbers, which the renderer
  // cannot decide from the style alone; leaving it unset gets that behaviour.
  return undefined;
}

function toVerticalAlign(value: string | undefined): VerticalAlign | undefined {
  if (value === "top") return "top";
  if (value === "center") return "middle";
  if (value === "bottom") return "bottom";
  return undefined;
}

/**
 * Parses styles.xml, mapping cellXfs indices to StyleObjects.
 *
 * Returns `null` at an index whose style is the default/no-op, so cells without
 * real formatting do not get a spurious `styles` entry.
 *
 * `themeXml` is `xl/theme/theme1.xml`. Without it, every themed colour — which is
 * most of them in a file Excel wrote — resolves to nothing.
 */
export function parseStylesXml(
  xml: string | undefined,
  themeXml?: ThemePalette,
): (StyleObject | null)[] {
  if (!xml) return [];
  const palette = themeXml ?? EMPTY_PALETTE;

  // Scoped to <numFmts>, not the whole document: <dxfs> holds <numFmt> elements
  // of its own, and a differential format declaring numFmtId="0" as "General"
  // would otherwise give every unformatted cell in the workbook a format code.
  const customCodes = parseNumFmts(findElement(xml, "numFmts")?.inner);

  const fontsBlock = findElement(xml, "fonts");
  const fonts = (fontsBlock ? findElements(fontsBlock.inner, "font") : []).map(
    (f) => parseFont(f.inner, palette),
  );

  const fillsBlock = findElement(xml, "fills");
  const fills = (fillsBlock ? findElements(fillsBlock.inner, "fill") : []).map(
    (f) => parseFill(f.inner, palette),
  );

  const bordersBlock = findElement(xml, "borders");
  const borders = (
    bordersBlock ? findElements(bordersBlock.inner, "border") : []
  ).map((b) => parseBorder(b.inner, palette));

  const styleXfsBlock = findElement(xml, "cellStyleXfs");
  const styleXfs = styleXfsBlock ? findElements(styleXfsBlock.inner, "xf") : [];

  const xfsBlock = findElement(xml, "cellXfs");
  if (!xfsBlock) return [];

  return findElements(xfsBlock.inner, "xf").map((xf) => {
    const ids = resolveIds(xf, styleXfs);
    const font = fonts[ids.fontId] ?? NO_FONT;
    const fill = fills[ids.fillId] ?? NO_FILL;
    const border = borders[ids.borderId] ?? null;
    const alignment = findElement(ids.alignmentFrom, "alignment")?.attrs ?? {};

    return buildStyle({
      font,
      fill,
      border,
      alignment,
      numFmtId: String(ids.numFmtId),
      customCodes,
    });
  });
}

interface StyleParts {
  font: ParsedFont;
  fill: ParsedFill;
  border: CellBorders | null;
  alignment: Record<string, string>;
  numFmtId: string;
  customCodes: Record<string, string>;
}

function buildStyle(parts: StyleParts): StyleObject | null {
  const { font, fill, border, alignment, numFmtId, customCodes } = parts;
  const style: StyleObject = {};

  if (font.bold) style.bold = true;
  if (font.italic) style.italic = true;
  if (font.underline) style.underline = true;
  if (font.color) style.color = font.color;
  if (font.family) style.fontFamily = font.family;
  if (font.size !== null) style.fontSize = font.size;
  if (fill.bg) style.bg = fill.bg;
  if (fill.gradient) style.gradient = fill.gradient;
  if (border) style.borders = border;

  const numFmt = numFmtToKey(numFmtId, customCodes);
  if (numFmt !== "general") style.numFmt = numFmt;
  const code = formatCodeFor(numFmtId, customCodes);
  if (code) style.numFmtCode = code;

  const align = toAlign(alignment.horizontal);
  if (align) style.align = align;
  const valign = toVerticalAlign(alignment.vertical);
  if (valign) style.valign = valign;
  if (alignment.wrapText === "1") style.wrap = true;
  const indent = Number.parseInt(alignment.indent ?? "", 10);
  if (Number.isFinite(indent) && indent > 0) style.indent = indent;
  const rotation = Number.parseInt(alignment.textRotation ?? "", 10);
  if (Number.isFinite(rotation) && rotation !== 0) {
    // The file counts anticlockwise from 0 to 90 and then encodes CLOCKWISE
    // rotation as 91..180, meaning 1..90 degrees the other way. 255 is the
    // stacked-vertical mode, which has no angle and is not represented here.
    style.rotation =
      rotation === VERTICAL_TEXT
        ? 0
        : rotation > QUARTER_TURN
          ? QUARTER_TURN - rotation
          : rotation;
  }

  return Object.keys(style).length === 0 ? null : style;
}

/**
 * Parses a `<dxf>` — a differential format — into a partial style.
 *
 * Unlike a `<cellXfs>` entry, a dxf holds its font, fill, border, and number
 * format inline rather than as indices, so there is no table to resolve against.
 * Everything it omits is inherited from whatever it is layered over, which is why
 * only the keys it states come back.
 */
export function parseDifferentialStyle(
  inner: string,
  palette: ThemePalette,
): StyleObject | null {
  const fontEl = findElement(inner, "font");
  const fillEl = findElement(inner, "fill");
  const borderEl = findElement(inner, "border");
  const numFmtEl = findElement(inner, "numFmt");
  const alignEl = findElement(inner, "alignment");

  const code = numFmtEl?.attrs.formatCode;
  return buildStyle({
    font: fontEl ? parseFont(fontEl.inner, palette) : NO_FONT,
    fill: fillEl ? parseFill(fillEl.inner, palette) : NO_FILL,
    border: borderEl ? parseBorder(borderEl.inner, palette) : null,
    alignment: alignEl?.attrs ?? {},
    // A dxf's format is stated inline, so it needs no id lookup — feeding the id
    // through the custom-code map is how the same builder handles both.
    numFmtId: numFmtEl?.attrs.numFmtId ?? "0",
    customCodes: code ? { [numFmtEl?.attrs.numFmtId ?? "0"]: code } : {},
  });
}

/** `numFmtId` -> `formatCode`, from a `<numFmts>` block. */
export function parseNumFmts(inner: string | undefined): Record<string, string> {
  const codes: Record<string, string> = {};
  if (!inner) return codes;
  for (const n of findElements(inner, "numFmt")) {
    const id = n.attrs.numFmtId;
    const code = n.attrs.formatCode;
    // "General" is the absence of a format, and Excel writes it explicitly in
    // places. Recording it would make every plain cell look formatted.
    if (id && code && code !== "General") codes[id] = code;
  }
  return codes;
}

/** The literal format code for an id, from the file or from the built-in table. */
export function formatCodeFor(
  id: string,
  customCodes: Record<string, string>,
): string | undefined {
  const custom = customCodes[id];
  if (custom) return custom;
  const n = Number.parseInt(id, 10);
  if (!Number.isFinite(n) || n >= FIRST_CUSTOM_NUMFMT_ID) return undefined;
  return BUILTIN_NUMFMTS[n];
}
