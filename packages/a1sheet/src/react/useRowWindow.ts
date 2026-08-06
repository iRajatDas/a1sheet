/**
 * Row virtualization and hidden-row compaction — a two-stage mapping.
 * Ported from ref/Spreadsheet.jsx:386-397.
 *
 * Stage 1: `effectiveHiddenRows` unions manually hidden rows with rows excluded
 *          by an active column filter. `visibleRows` is the compacted array of
 *          absolute row indices in the non-frozen band, hidden rows removed.
 * Stage 2: `windowRows` slices that compacted array from scrollTop and hands
 *          each row a CONSECUTIVE grid line. The lines are consecutive rather
 *          than absolute because rows can differ in height: an unrendered row
 *          has no track, so there is nothing to size it, and placing row 5000
 *          on line 5002 would leave the browser guessing at 5000 heights. A
 *          single spacer item ahead of the window supplies the missing distance
 *          in one go — `leadingSpace` — and the rows below follow in order.
 *
 * Heights come from `sheet.rowHeights`, defaulting to ROW_HEIGHT. While that
 * map is empty every offset is `index * ROW_HEIGHT`, and the hook skips
 * building an offset table at all; a hundred-thousand-row sheet should not pay
 * for a feature it is not using. Once any row is resized the table appears and
 * lookups become a binary search.
 *
 * Frozen rows are a SEPARATE always-rendered band and are assumed never hidden.
 * Adding column hiding, or allowing frozen rows to be filtered out, means
 * revisiting this mapping.
 */
import { useCallback, useMemo } from "react";
import { cellKey } from "../model/address.js";
import type { Sheet } from "../model/types.js";
import { BUFFER_ROWS, HEADER_HEIGHT, ROW_HEIGHT } from "./constants.js";

/** An absolute row index paired with the CSS grid line it renders on. */
export interface WindowRow {
  absRow: number;
  gridRow: number;
}

export interface UseRowWindowResult {
  /** Manual hides unioned with filter exclusions. */
  effectiveHiddenRows: Set<number>;
  /** Absolute row indices of the non-frozen band, hidden rows removed. */
  visibleRows: number[];
  /** The slice to actually render, with grid lines. */
  windowRows: WindowRow[];
  /** The always-rendered frozen band, with grid lines. */
  frozenRowsList: WindowRow[];
  /** Index into `visibleRows` where `windowRows` begins. */
  startVisual: number;
  /** Height of one row, falling back to the default. */
  rowHeight(row: number): number;
  /**
   * Distance from below the column header to the top of a row, hidden rows
   * collapsed out. Null when the row itself is hidden.
   */
  rowTop(absRow: number): number | null;
  /** The row whose band contains `y`, measured the same way as `rowTop`. */
  rowAt(y: number): number;
  /** Height of the frozen band, header excluded. */
  frozenHeight: number;
  /** Distance the spacer item must cover before the first windowed row. */
  leadingSpace: number;
  /** Pixel height of the whole sheet, rendered rows or not. */
  contentHeight: number;
}

const clamp = (v: number, lo: number, hi: number) => Math.max(lo, Math.min(hi, v));

/**
 * A row is filter-hidden when any filtered column's allowed-value set does not
 * contain that row's displayed value for the column. `getDisplay` is passed in
 * rather than the evaluator so this hook stays independent of formula concerns.
 */
export function useRowWindow(
  sheet: Sheet,
  scrollTop: number,
  viewportHeight: number,
  getDisplay: (row: number, col: number) => string,
): UseRowWindowResult {
  const frozenRows = sheet.frozenRows || 0;

  // `sheet.cells` and `sheet.styles` are dependencies twice over: read directly
  // below, and read through `getDisplay`. An edit can change whether a row
  // passes the filter, so dropping either leaves the result stale.
  const effectiveHiddenRows = useMemo(() => {
    const hidden = new Set(sheet.hiddenRows);
    const filterCols = Object.keys(sheet.filters).map(Number);
    if (filterCols.length > 0) {
      // A cell with no content and no style displays the same string in every
      // row of its column, because nothing in that display depends on the row.
      // Resolving it once per column instead of once per row is what keeps this
      // scan affordable on a sparse sheet — it reruns on every edit, since an
      // edit can change whether a row passes the filter.
      const emptyDisplay = new Map<number, string>();
      for (let r = frozenRows; r < sheet.numRows; r++) {
        if (hidden.has(r)) continue;
        for (const c of filterCols) {
          const allowed = sheet.filters[c];
          if (!allowed) continue;
          const key = cellKey(r, c);
          let display: string;
          if (sheet.cells[key] === undefined && sheet.styles[key] === undefined) {
            let cached = emptyDisplay.get(c);
            if (cached === undefined) {
              cached = getDisplay(r, c);
              emptyDisplay.set(c, cached);
            }
            display = cached;
          } else {
            display = getDisplay(r, c);
          }
          if (!allowed.has(display)) {
            hidden.add(r);
            break;
          }
        }
      }
    }
    return hidden;
  }, [
    sheet.hiddenRows,
    sheet.filters,
    sheet.numRows,
    sheet.cells,
    sheet.styles,
    frozenRows,
    getDisplay,
  ]);

  const visibleRows = useMemo(() => {
    const out: number[] = [];
    for (let r = frozenRows; r < sheet.numRows; r++) {
      if (!effectiveHiddenRows.has(r)) out.push(r);
    }
    return out;
  }, [frozenRows, sheet.numRows, effectiveHiddenRows]);

  const rowHeight = useCallback(
    (row: number) => sheet.rowHeights[row] ?? ROW_HEIGHT,
    [sheet.rowHeights],
  );

  /** True while no row has been resized, which is the overwhelmingly common case. */
  const uniform = useMemo(
    () => Object.keys(sheet.rowHeights).length === 0,
    [sheet.rowHeights],
  );

  /** Prefix sums down the frozen band. Small by definition, so always built. */
  const frozenOffsets = useMemo(() => {
    const out: number[] = [0];
    let y = 0;
    for (let r = 0; r < frozenRows; r++) {
      y += sheet.rowHeights[r] ?? ROW_HEIGHT;
      out.push(y);
    }
    return out;
  }, [frozenRows, sheet.rowHeights]);
  const frozenHeight = frozenOffsets[frozenRows] ?? 0;

  /**
   * Prefix sums down the scrollable band, or null while every row is the
   * default height — there `index * ROW_HEIGHT` is exact, and allocating an
   * array as long as the sheet on every filter change would be pure waste.
   */
  const bandOffsets = useMemo(() => {
    if (uniform) return null;
    const out: number[] = [0];
    let y = 0;
    for (const r of visibleRows) {
      y += sheet.rowHeights[r] ?? ROW_HEIGHT;
      out.push(y);
    }
    return out;
  }, [uniform, visibleRows, sheet.rowHeights]);

  /** Top of the nth visible row, relative to the start of the scrollable band. */
  const bandTopAt = useCallback(
    (index: number) => {
      const i = clamp(index, 0, visibleRows.length);
      return bandOffsets ? (bandOffsets[i] ?? 0) : i * ROW_HEIGHT;
    },
    [bandOffsets, visibleRows.length],
  );

  /** Index of the visible row whose track contains `y` within the band. */
  const bandIndexAt = useCallback(
    (y: number) => {
      const last = visibleRows.length - 1;
      if (last < 0) return 0;
      if (!bandOffsets) return clamp(Math.floor(y / ROW_HEIGHT), 0, last);
      let lo = 0;
      let hi = last;
      while (lo < hi) {
        const mid = (lo + hi + 1) >> 1;
        if ((bandOffsets[mid] ?? 0) <= y) lo = mid;
        else hi = mid - 1;
      }
      return lo;
    },
    [bandOffsets, visibleRows.length],
  );

  const bandHeight = bandTopAt(visibleRows.length);

  // The scrollable band starts below the header and the frozen rows, both of
  // which are sticky and therefore do not scroll away.
  const bandScrollTop = Math.max(0, scrollTop - (HEADER_HEIGHT + frozenHeight));
  const startVisual = clamp(
    bandIndexAt(bandScrollTop) - BUFFER_ROWS,
    0,
    visibleRows.length,
  );
  const endVisual = clamp(
    bandIndexAt(bandScrollTop + viewportHeight) + 1 + BUFFER_ROWS,
    0,
    visibleRows.length,
  );

  const windowRows = useMemo(
    () =>
      visibleRows.slice(startVisual, endVisual).map((absRow, i) => ({
        absRow,
        // Consecutive, not absolute: the spacer ahead of the window covers the
        // distance instead. See the note at the top of the file.
        gridRow: frozenRows + i + 3,
      })),
    [visibleRows, startVisual, endVisual, frozenRows],
  );

  const frozenRowsList = useMemo(
    () =>
      Array.from({ length: frozenRows }, (_, i) => ({
        absRow: i,
        gridRow: i + 2,
      })),
    [frozenRows],
  );

  /**
   * `visibleRows` is ascending, so this is a binary search rather than the
   * `indexOf` it replaced. It runs once per highlighted reference on every
   * keystroke of a formula, against an array as long as the sheet.
   */
  const bandIndexOf = useCallback(
    (absRow: number): number | null => {
      let lo = 0;
      let hi = visibleRows.length - 1;
      while (lo <= hi) {
        const mid = (lo + hi) >> 1;
        const value = visibleRows[mid] as number;
        if (value === absRow) return mid;
        if (value < absRow) lo = mid + 1;
        else hi = mid - 1;
      }
      return null;
    },
    [visibleRows],
  );

  const rowTop = useCallback(
    (absRow: number): number | null => {
      if (absRow < frozenRows) return frozenOffsets[absRow] ?? 0;
      const index = bandIndexOf(absRow);
      return index === null ? null : frozenHeight + bandTopAt(index);
    },
    [frozenRows, frozenOffsets, frozenHeight, bandIndexOf, bandTopAt],
  );

  const rowAt = useCallback(
    (y: number): number => {
      if (y < frozenHeight) {
        let row = 0;
        while (row + 1 < frozenRows && (frozenOffsets[row + 1] ?? 0) <= y) row++;
        return row;
      }
      const index = bandIndexAt(y - frozenHeight);
      return visibleRows[index] ?? Math.max(0, sheet.numRows - 1);
    },
    [
      frozenHeight,
      frozenRows,
      frozenOffsets,
      bandIndexAt,
      visibleRows,
      sheet.numRows,
    ],
  );

  return {
    effectiveHiddenRows,
    visibleRows,
    windowRows,
    frozenRowsList,
    startVisual,
    rowHeight,
    rowTop,
    rowAt,
    frozenHeight,
    leadingSpace: bandTopAt(startVisual),
    contentHeight: HEADER_HEIGHT + frozenHeight + bandHeight,
  };
}
