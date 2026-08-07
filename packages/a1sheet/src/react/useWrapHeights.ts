"use client";

/**
 * How tall each row needs to be for its wrapped cells to fit.
 *
 * `wrap` makes text run onto more lines; without this the row keeps the height
 * it had and the extra lines are clipped. Excel and Sheets grow the row instead,
 * and stop growing it the moment you drag it — an explicit height always wins,
 * which is why this returns what a row NEEDS rather than setting anything.
 *
 * The cost is why it is a cache rather than a computation. Row offsets are a
 * whole-sheet quantity, so every wrapped cell has to be accounted for on every
 * edit, not just the visible ones — and measuring one is a greedy line-break
 * over its words, tens of microseconds. So the measurement is keyed on the three
 * things it depends on (the text, the width it wraps into, the font) and reused
 * whenever none of them moved. What an edit costs is then three comparisons per
 * wrapped cell, and measurement only of what actually changed.
 *
 * A merged cell is measured across the columns it spans, because that is the box
 * the browser wraps it in, and it asks only for the height its rows do not
 * already supply between them.
 *
 * A sheet with no wrapped cell allocates nothing and returns a shared empty map,
 * so nothing below pays for a feature it is not using. That the map is the SAME
 * object when nothing changed matters as much: `useRowWindow` builds its offset
 * table from it, and a fresh map every render would rebuild that table every
 * render.
 */
import { useRef } from "react";
import { wrappedLineCount } from "../format/wrapText.js";
import { getMergeAt } from "../model/sheet.js";
import type { CellKey, Sheet, StyleObject } from "../model/types.js";
import {
  CELL_LINE_RATIO,
  CELL_PADDING_X,
  CELL_PADDING_Y,
  type CellFont,
  DEFAULT_COL_WIDTH,
  ROW_HEIGHT,
} from "./constants.js";
import type { MeasureText } from "./useTextMeasurer.js";

/** Row index -> the height its wrapped content needs. Rows that fit are absent. */
export type WrapHeights = ReadonlyMap<number, number>;

const NONE: WrapHeights = new Map();

/** One cell's measurement, and the inputs it was valid for. */
interface Measured {
  text: string;
  width: number;
  font: string;
  /** Height this cell alone demands of the rows it occupies. */
  height: number;
}

interface WrapCache {
  cells: Sheet["cells"];
  styles: Sheet["styles"];
  colWidths: Sheet["colWidths"];
  merges: Sheet["merges"];
  base: CellFont;
  heights: WrapHeights;
  perCell: Map<CellKey, Measured>;
}

export interface WrapHeightsInput {
  sheet: Sheet;
  getDisplay(row: number, col: number): string;
  measure: MeasureText;
  /** The face the theme draws cells in. A cell's own style layers over it. */
  base: CellFont;
}

/**
 * A CSS `font` shorthand for a cell, since a bold or larger cell wraps at a
 * different width and needs taller lines than its neighbours.
 */
function fontOf(style: StyleObject, base: CellFont): string {
  const size = style.fontSize ?? base.size;
  const family = style.fontFamily
    ? `${JSON.stringify(style.fontFamily)}, ${base.family}`
    : base.family;
  const weight = style.bold ? "bold " : "";
  const slant = style.italic ? "italic " : "";
  return `${slant}${weight}${size}px ${family}`;
}

/** The height `lines` of text at this size occupy, padding included. */
function heightFor(lines: number, style: StyleObject, base: CellFont): number {
  const size = style.fontSize ?? base.size;
  const box = Math.round(size * CELL_LINE_RATIO);
  return Math.max(ROW_HEIGHT, lines * box + CELL_PADDING_Y);
}

export function useWrapHeights({
  sheet,
  getDisplay,
  measure,
  base,
}: WrapHeightsInput): WrapHeights {
  const cacheRef = useRef<WrapCache | null>(null);
  const cache = cacheRef.current;

  const { cells, styles, colWidths, merges } = sheet;

  if (
    cache !== null &&
    cache.cells === cells &&
    cache.styles === styles &&
    cache.colWidths === colWidths &&
    cache.merges === merges &&
    cache.base === base
  ) {
    return cache.heights;
  }

  // Every measurement is in terms of the base face, so a theme change is the one
  // input that invalidates the whole cache rather than one entry of it.
  const previous = cache?.base === base ? cache.perCell : undefined;
  const perCell = new Map<CellKey, Measured>();
  const heights = new Map<number, number>();

  // Wrapped cells are found through `styles`, not by scanning the sheet: wrap is
  // a style, so a sheet with none never looks at a cell at all.
  for (const key of Object.keys(styles) as CellKey[]) {
    const style = styles[key];
    if (!style?.wrap) continue;

    const separator = key.indexOf("_");
    const row = Number(key.slice(0, separator));
    const col = Number(key.slice(separator + 1));

    // A merge renders as one box from its top-left cell; the cells it covers
    // render nothing at all, so they have no text to grow anything by.
    const merge =
      sheet.merges.length === 0 ? undefined : getMergeAt(sheet, row, col);
    if (merge && (merge.r1 !== row || merge.c1 !== col)) continue;

    const text = getDisplay(row, col);
    if (text === "") continue;

    const width = wrapWidth(colWidths, merge?.c1 ?? col, merge?.c2 ?? col);
    const font = fontOf(style, base);

    let measured = previous?.get(key);
    if (
      measured === undefined ||
      measured.text !== text ||
      measured.width !== width ||
      measured.font !== font
    ) {
      const lines = wrappedLineCount(text, width, (run) => measure(run, font));
      measured = { text, width, font, height: heightFor(lines, style, base) };
    }
    perCell.set(key, measured);

    // A merge is already as tall as the rows it spans, so only the shortfall is
    // asked for — and it is asked of the LAST of those rows, so the box grows
    // downwards from wherever it starts rather than shoving its own top around.
    const spanned = merge ? merge.r2 - merge.r1 + 1 : 1;
    const demand = measured.height - (spanned - 1) * ROW_HEIGHT;
    if (demand <= ROW_HEIGHT) continue;

    const target = merge ? merge.r2 : row;
    const tallest = heights.get(target) ?? 0;
    if (demand > tallest) heights.set(target, demand);
  }

  const result: WrapHeights = heights.size === 0 ? NONE : heights;
  cacheRef.current = {
    cells,
    styles,
    colWidths,
    merges,
    base,
    heights: result,
    perCell,
  };
  return result;
}

/**
 * The text width of a cell spanning columns `c1..c2` inclusive — its own column
 * when it is not merged.
 *
 * A hidden column has width 0 here, which would break every word onto its own
 * line for no visible benefit. Its natural width is the honest answer.
 */
function wrapWidth(colWidths: Sheet["colWidths"], c1: number, c2: number): number {
  let total = 0;
  for (let c = c1; c <= c2; c++) total += colWidths[c] ?? DEFAULT_COL_WIDTH;
  return total - CELL_PADDING_X;
}
