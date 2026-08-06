/**
 * Resolving an OOXML `<color>` to a hex value.
 *
 * A colour in styles.xml is written one of four ways, and only one of them is a
 * plain RGB triple:
 *
 *   <color rgb="FF612890"/>                  literal ARGB
 *   <color theme="4" tint="-0.25"/>          index into the workbook theme
 *   <color indexed="64"/>                    index into the legacy 1990s palette
 *   <color auto="1"/>                        "whatever the context wants"
 *
 * Excel writes the themed form for almost everything, which is why reading only
 * `rgb` made a themed workbook import nearly monochrome.
 */
import type { HexColor } from "../../model/types.js";
import { findElement, rgbToColor } from "./xml.js";

/**
 * The theme palette, in the order `theme="N"` indexes it.
 *
 * NOT the order the file lists them in: theme1.xml writes dk1, lt1, dk2, lt2,
 * and Excel indexes them lt1, dk1, lt2, dk2. The first two pairs are swapped, so
 * `theme="0"` is the light background and `theme="1"` is the dark text. Getting
 * this backwards inverts every document — black text on a black fill.
 */
const THEME_SLOTS = [
  "lt1",
  "dk1",
  "lt2",
  "dk2",
  "accent1",
  "accent2",
  "accent3",
  "accent4",
  "accent5",
  "accent6",
  "hlink",
  "folHlink",
] as const;

/** The theme's accent colours, in accent order. Table styles are built on these. */
export interface ThemePalette {
  /** Indexed by `theme="N"`. A slot the theme omits is absent. */
  readonly slots: readonly (HexColor | undefined)[];
}

export const EMPTY_PALETTE: ThemePalette = { slots: [] };

/**
 * The legacy indexed palette, which predates themes and is still emitted by
 * LibreOffice and by anything writing pre-2007-style styles. Entries 0-7 repeat
 * as 8-15 for historical reasons; 64 and 65 are the "system foreground" and
 * "system background" sentinels, which resolve like `auto`.
 */
const INDEXED: readonly (string | null)[] = [
  "000000",
  "FFFFFF",
  "FF0000",
  "00FF00",
  "0000FF",
  "FFFF00",
  "FF00FF",
  "00FFFF",
  "000000",
  "FFFFFF",
  "FF0000",
  "00FF00",
  "0000FF",
  "FFFF00",
  "FF00FF",
  "00FFFF",
  "800000",
  "008000",
  "000080",
  "808000",
  "800080",
  "008080",
  "C0C0C0",
  "808080",
  "9999FF",
  "993366",
  "FFFFCC",
  "CCFFFF",
  "660066",
  "FF8080",
  "0066CC",
  "CCCCFF",
  "000080",
  "FF00FF",
  "FFFF00",
  "00FFFF",
  "800080",
  "800000",
  "008080",
  "0000FF",
  "00CCFF",
  "CCFFFF",
  "CCFFCC",
  "FFFF99",
  "99CCFF",
  "FF99CC",
  "CC99FF",
  "FFCC99",
  "3366FF",
  "33CCCC",
  "99CC00",
  "FFCC00",
  "FF9900",
  "FF6600",
  "666699",
  "969696",
  "003366",
  "339966",
  "003300",
  "333300",
  "993300",
  "993366",
  "333399",
  "333333",
  null,
  null,
];

/**
 * Reads the colour scheme out of `xl/theme/theme1.xml`.
 *
 * dk1 and lt1 are usually `<a:sysClr val="windowText" lastClr="000000"/>` rather
 * than an `srgbClr` — a reference to an OS colour, with the value it last
 * resolved to cached alongside. `lastClr` is that cache, and it is the only
 * sensible thing to use outside Windows.
 */
export function parseThemePalette(xml: string | undefined): ThemePalette {
  if (!xml) return EMPTY_PALETTE;

  const scheme = findElement(xml, "a:clrScheme");
  if (!scheme) return EMPTY_PALETTE;

  const slots = THEME_SLOTS.map((slot) => {
    const el = findElement(scheme.inner, `a:${slot}`);
    if (!el) return undefined;
    const srgb = findElement(el.inner, "a:srgbClr")?.attrs.val;
    if (srgb) return rgbToColor(srgb) ?? undefined;
    const sys = findElement(el.inner, "a:sysClr")?.attrs.lastClr;
    return rgbToColor(sys) ?? undefined;
  });

  return { slots };
}

const RGB_MAX = 255;
const HEX_RADIX = 16;

function channels(color: HexColor): [number, number, number] {
  const h = color.slice(1);
  return [
    parseInt(h.slice(0, 2), HEX_RADIX),
    parseInt(h.slice(2, 4), HEX_RADIX),
    parseInt(h.slice(4, 6), HEX_RADIX),
  ];
}

function toHex(r: number, g: number, b: number): HexColor {
  const part = (n: number) =>
    Math.max(0, Math.min(RGB_MAX, Math.round(n)))
      .toString(HEX_RADIX)
      .padStart(2, "0");
  return `#${part(r)}${part(g)}${part(b)}` as HexColor;
}

/**
 * Lightens (`tint > 0`) or darkens (`tint < 0`) a colour.
 *
 * OOXML specifies this against HSL luminance. This is the RGB approximation every
 * other implementation uses; it differs from Excel by a shade or two on saturated
 * colours and not at all on greys, which is a better trade than carrying an
 * HSL round trip for the sake of the last few units.
 */
export function applyTint(color: HexColor, tint: number): HexColor {
  if (!tint) return color;
  const [r, g, b] = channels(color);
  if (tint > 0) {
    const toward = (c: number) => c + (RGB_MAX - c) * tint;
    return toHex(toward(r), toward(g), toward(b));
  }
  const scale = 1 + tint;
  return toHex(r * scale, g * scale, b * scale);
}

/** Blends two colours. `weight` is how much of `over` to use, 0..1. */
export function mixColors(
  under: HexColor,
  over: HexColor,
  weight: number,
): HexColor {
  const [r1, g1, b1] = channels(under);
  const [r2, g2, b2] = channels(over);
  const at = (a: number, b: number) => a + (b - a) * weight;
  return toHex(at(r1, r2), at(g1, g2), at(b1, b2));
}

/**
 * Resolves a `<color>` element's attributes to a hex value, or null when the file
 * declines to name one — `auto="1"`, an empty element, or an index the palette
 * does not cover. Null means "inherit", never black: defaulting to black is how a
 * missing border colour turns into a black grid.
 */
export function resolveColorAttrs(
  attrs: Record<string, string>,
  palette: ThemePalette,
): HexColor | null {
  if (attrs.rgb) {
    const direct = rgbToColor(attrs.rgb);
    if (direct) return tinted(direct, attrs.tint);
  }

  if (attrs.theme !== undefined) {
    const index = Number.parseInt(attrs.theme, 10);
    const slot = Number.isFinite(index) ? palette.slots[index] : undefined;
    if (slot) return tinted(slot, attrs.tint);
  }

  if (attrs.indexed !== undefined) {
    const index = Number.parseInt(attrs.indexed, 10);
    const entry = Number.isFinite(index) ? INDEXED[index] : undefined;
    const hex = entry ? rgbToColor(entry) : null;
    if (hex) return tinted(hex, attrs.tint);
  }

  return null;
}

function tinted(color: HexColor, tint: string | undefined): HexColor {
  if (tint === undefined) return color;
  const amount = Number.parseFloat(tint);
  return Number.isFinite(amount) ? applyTint(color, amount) : color;
}
