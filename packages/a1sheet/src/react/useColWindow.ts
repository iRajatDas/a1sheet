"use client";

/**
 * Column virtualization — the horizontal counterpart to `useRowWindow`.
 *
 * Columns differ from rows in two ways that shape this hook:
 *
 * 1. **Widths vary.** There is no `col * WIDTH` shortcut, so the hook builds a
 *    cumulative offset table once per width change. Every lookup afterwards is
 *    O(1) (`colOffset`) or O(log n) (`colAt`). The grid used to sum widths from
 *    column 0 on every call, which is O(n) per frozen cell and per hit-test.
 * 2. **Tracks stay explicit.** `gridTemplateColumns` still names every column,
 *    so the scroll container keeps its true width and sticky offsets keep
 *    working. Virtualization decides which *cells* exist, not which tracks do —
 *    a track definition costs a string entry, a cell costs a DOM node.
 *
 * Columns hidden by the window simply are not rendered, exactly as with rows.
 */
import { useCallback, useMemo } from "react";
import type { Sheet } from "../model/types.js";
import { BUFFER_COLS, DEFAULT_COL_WIDTH, ROW_HEADER_WIDTH } from "./constants.js";

export interface UseColWindowResult {
  /** Width of one column, falling back to the default. */
  colWidth(col: number): number;
  /** Left edge of a column, measured from the first column (row header excluded). */
  colOffset(col: number): number;
  /** Column index at an x offset measured the same way as `colOffset`. */
  colAt(x: number): number;
  /** Summed width of every column — the grid's true horizontal extent. */
  totalWidth: number;
  /**
   * Ascending column indices to render: the frozen band, the scroll window, and
   * the origin of any merge that spans into it.
   */
  windowCols: number[];
}

const clamp = (v: number, lo: number, hi: number) => Math.max(lo, Math.min(hi, v));

export function useColWindow(
  sheet: Sheet,
  scrollLeft: number,
  viewportWidth: number,
): UseColWindowResult {
  const { numCols, colWidths, merges } = sheet;
  const frozenCols = sheet.frozenCols || 0;

  /** Prefix sums; `offsets[c]` is the left edge of column `c`, `offsets[numCols]` the total. */
  const offsets = useMemo(() => {
    const out: number[] = [0];
    let x = 0;
    for (let c = 0; c < numCols; c++) {
      x += colWidths[c] ?? DEFAULT_COL_WIDTH;
      out.push(x);
    }
    return out;
  }, [numCols, colWidths]);

  const colWidth = useCallback(
    (col: number) => colWidths[col] ?? DEFAULT_COL_WIDTH,
    [colWidths],
  );

  const colOffset = useCallback(
    (col: number) => offsets[clamp(col, 0, numCols)] ?? 0,
    [offsets, numCols],
  );

  const colAt = useCallback(
    (x: number) => {
      let lo = 0;
      let hi = numCols - 1;
      while (lo < hi) {
        const mid = (lo + hi + 1) >> 1;
        if ((offsets[mid] ?? 0) <= x) lo = mid;
        else hi = mid - 1;
      }
      return Math.max(0, lo);
    },
    [offsets, numCols],
  );

  const windowCols = useMemo(() => {
    if (numCols === 0) return [];

    // The frozen band sits over the left of the viewport, so a few columns
    // behind it are drawn needlessly. Accounting for that exactly would couple
    // this to the sticky offsets for the sake of two or three cells.
    const viewLeft = Math.max(0, scrollLeft - ROW_HEADER_WIDTH);
    const last = numCols - 1;
    const first = clamp(colAt(viewLeft) - BUFFER_COLS, 0, last);
    const end = clamp(colAt(viewLeft + viewportWidth) + BUFFER_COLS, 0, last);

    const picked = new Set<number>();
    for (let c = 0; c < frozenCols && c < numCols; c++) picked.add(c);
    for (let c = first; c <= end; c++) picked.add(c);

    // A merge is rendered by its top-left cell. Once that cell scrolls out of
    // the window the whole merged block would vanish, so its origin column is
    // pulled back in.
    for (const m of merges) {
      const c1 = Math.min(m.c1, m.c2);
      const c2 = Math.max(m.c1, m.c2);
      if (c1 < first && c2 >= first) picked.add(c1);
    }

    return [...picked].sort((a, b) => a - b);
  }, [colAt, scrollLeft, viewportWidth, numCols, frozenCols, merges]);

  return {
    colWidth,
    colOffset,
    colAt,
    totalWidth: offsets[numCols] ?? 0,
    windowCols,
  };
}
