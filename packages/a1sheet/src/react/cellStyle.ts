/**
 * Turning a `StyleObject` into CSS.
 *
 * Separate from `Cell` because it is pure and worth testing directly: a border
 * mapping or a gradient angle is easier to get wrong than to notice, and a render
 * test would only prove that *something* was set.
 */
import type { CSSProperties } from "react";
import type {
  BorderEdge,
  BorderLine,
  GradientFill,
  StyleObject,
} from "../model/types.js";
import type { Theme } from "./theme.js";

/**
 * Excel's thirteen line kinds against the four CSS border styles that can
 * represent them. Width is in pixels.
 *
 * `hair` is Excel's finest line and `thin` is its default; both land on 1px,
 * because a fraction of a pixel renders as either 0 or 1 depending on the
 * device ratio and a border that vanishes at some zoom levels is worse than one
 * a shade too heavy. The `medium*` dashed family differs from the plain dashed
 * family only in weight.
 */
const BORDER_CSS: Record<BorderLine, { width: number; style: string }> = {
  hair: { width: 1, style: "solid" },
  thin: { width: 1, style: "solid" },
  medium: { width: 2, style: "solid" },
  thick: { width: 3, style: "solid" },
  double: { width: 3, style: "double" },
  dotted: { width: 1, style: "dotted" },
  dashed: { width: 1, style: "dashed" },
  dashDot: { width: 1, style: "dashed" },
  dashDotDot: { width: 1, style: "dotted" },
  mediumDashed: { width: 2, style: "dashed" },
  mediumDashDot: { width: 2, style: "dashed" },
  mediumDashDotDot: { width: 2, style: "dotted" },
  slantDashDot: { width: 2, style: "dashed" },
};

function edgeCss(edge: BorderEdge, fallbackColor: string): string {
  const css = BORDER_CSS[edge.line];
  return `${css.width}px ${css.style} ${edge.color ?? fallbackColor}`;
}

/**
 * CSS gradients measure from the direction the sweep points; OOXML's `degree`
 * measures the sweep itself, clockwise from left-to-right. A 90° OOXML gradient
 * runs top-to-bottom, which CSS calls `180deg`.
 */
const OOXML_TO_CSS_DEGREES = 90;
const PERCENT = 100;

function gradientCss(gradient: GradientFill): string {
  const stops = [...gradient.stops]
    .sort((a, b) => a.position - b.position)
    .map((s) => `${s.color} ${Math.round(s.position * PERCENT)}%`)
    .join(", ");
  return `linear-gradient(${gradient.degree + OOXML_TO_CSS_DEGREES}deg, ${stops})`;
}

const JUSTIFY: Record<string, string> = {
  left: "flex-start",
  center: "center",
  right: "flex-end",
};

const ALIGN_ITEMS: Record<string, string> = {
  top: "flex-start",
  middle: "center",
  bottom: "flex-end",
};

/**
 * Everything a cell's appearance comes from, as CSS.
 *
 * The base rule in the stylesheet draws a 1px grid line on the right and bottom
 * of every cell. An explicit border must replace that line rather than sit
 * alongside it, so any edge the style sets is written out in full — and the two
 * edges the grid does NOT draw (top and left) are only written when set, or every
 * cell would gain a double line against its neighbour.
 */
export function cellCss(style: StyleObject, theme: Theme): CSSProperties {
  const color = style.color ?? theme.cellText;
  const css: CSSProperties = {
    fontWeight: style.bold ? 700 : 400,
    fontStyle: style.italic ? "italic" : "normal",
    textDecoration: style.underline ? "underline" : "none",
    color,
    // Always opaque. The selection tint is an ::after overlay, and a frozen cell
    // is sticky, where a transparent background shows the scrolled content through.
    background: style.gradient
      ? gradientCss(style.gradient)
      : (style.bg ?? theme.cellBg),
  };

  if (style.align) {
    css.textAlign = style.align;
    css.justifyContent = JUSTIFY[style.align];
  }
  if (style.valign) css.alignItems = ALIGN_ITEMS[style.valign];
  if (style.fontFamily) {
    // A file names one family and nothing else. Keeping the theme stack behind it
    // means a font the reader does not have falls back to the sheet's own face
    // rather than to whatever the browser defaults to.
    css.fontFamily = `${quoteFamily(style.fontFamily)}, ${theme.fontFamily}`;
  }
  if (style.fontSize !== undefined) css.fontSize = style.fontSize;

  if (style.wrap) {
    css.whiteSpace = "normal";
    css.overflowWrap = "anywhere";
  }

  const borders = style.borders;
  if (borders) {
    if (borders.top) css.borderTop = edgeCss(borders.top, color);
    if (borders.left) css.borderLeft = edgeCss(borders.left, color);
    // Right and bottom always win over the grid line when set, and keep it when not.
    if (borders.right) css.borderRight = edgeCss(borders.right, color);
    if (borders.bottom) css.borderBottom = edgeCss(borders.bottom, color);
  }

  return css;
}

/** A family name with spaces needs quoting before it joins a font stack. */
function quoteFamily(family: string): string {
  return /^[\w-]+$/.test(family) ? family : `"${family.replace(/"/g, "")}"`;
}
