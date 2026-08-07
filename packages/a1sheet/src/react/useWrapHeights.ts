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
 * A sheet with no wrapped cell allocates nothing and returns a shared empty map,
 * so nothing below pays for a feature it is not using. That the map is the SAME
 * object when nothing changed matters as much: `useRowWindow` builds its offset
 * table from it, and a fresh map every render would rebuild that table every
 * render.
 */
import { useRef } from "react";
import { wrappedLineCount } from "../format/wrapText.js";
import type { CellKey, Sheet, StyleObject } from "../model/types.js";
import {
  CELL_FONT_SIZE,
  CELL_FONT_STACK,
  CELL_LINE_RATIO,
  CELL_PADDING_X,
  CELL_PADDING_Y,
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
  lines: number;
  /** Height this cell alone demands of its row. */
  height: number;
}

interface WrapCache {
  cells: Sheet["cells"];
  styles: Sheet["styles"];
  colWidths: Sheet["colWidths"];
  heights: WrapHeights;
  perCell: Map<CellKey, Measured>;
}

/**
 * A CSS `font` shorthand for a cell, since a bold or larger cell wraps at a
 * different width and needs taller lines than its neighbours.
 */
function fontOf(style: StyleObject): string {
  const size = style.fontSize ?? CELL_FONT_SIZE;
  const family = style.fontFamily
    ? `${JSON.stringify(style.fontFamily)}, ${CELL_FONT_STACK}`
    : CELL_FONT_STACK;
  const weight = style.bold ? "bold " : "";
  const slant = style.italic ? "italic " : "";
  return `${slant}${weight}${size}px ${family}`;
}

/** The height `lines` of text at this size occupy, padding included. */
function heightFor(lines: number, style: StyleObject): number {
  const size = style.fontSize ?? CELL_FONT_SIZE;
  const box = Math.round(size * CELL_LINE_RATIO);
  return Math.max(ROW_HEIGHT, lines * box + CELL_PADDING_Y);
}

export function useWrapHeights(
  sheet: Sheet,
  getDisplay: (row: number, col: number) => string,
  measure: MeasureText,
): WrapHeights {
  const cacheRef = useRef<WrapCache | null>(null);
  const cache = cacheRef.current;

  const { cells, styles, colWidths } = sheet;

  if (
    cache !== null &&
    cache.cells === cells &&
    cache.styles === styles &&
    cache.colWidths === colWidths
  ) {
    return cache.heights;
  }

  // Wrapped cells are found through `styles`, not by scanning the sheet: wrap is
  // a style, so a sheet with none never looks at a cell at all.
  const previous = cache?.perCell;
  const perCell = new Map<CellKey, Measured>();
  const heights = new Map<number, number>();

  for (const key of Object.keys(styles) as CellKey[]) {
    const style = styles[key];
    if (!style?.wrap) continue;

    const separator = key.indexOf("_");
    const row = Number(key.slice(0, separator));
    const col = Number(key.slice(separator + 1));

    const text = getDisplay(row, col);
    if (text === "") continue;

    // A hidden column has width 0 here, which would break every word onto its
    // own line for no visible benefit. Its natural width is the honest answer.
    const width = (colWidths[col] ?? DEFAULT_COL_WIDTH) - CELL_PADDING_X;
    const font = fontOf(style);

    const hit = previous?.get(key);
    const measured =
      hit !== undefined &&
      hit.text === text &&
      hit.width === width &&
      hit.font === font
        ? hit
        : remeasure(text, width, font, style, measure);

    perCell.set(key, measured);
    if (measured.lines < 2) continue;
    const tallest = heights.get(row) ?? 0;
    if (measured.height > tallest) heights.set(row, measured.height);
  }

  const result: WrapHeights = heights.size === 0 ? NONE : heights;
  cacheRef.current = { cells, styles, colWidths, heights: result, perCell };
  return result;
}

function remeasure(
  text: string,
  width: number,
  font: string,
  style: StyleObject,
  measure: MeasureText,
): Measured {
  const lines = wrappedLineCount(text, width, (run) => measure(run, font));
  return { text, width, font, lines, height: heightFor(lines, style) };
}
