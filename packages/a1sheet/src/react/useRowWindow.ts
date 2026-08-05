/**
 * Row virtualization and hidden-row compaction — a two-stage mapping.
 * Ported from ref/Spreadsheet.jsx:386-397.
 *
 * Stage 1: `effectiveHiddenRows` unions manually hidden rows with rows excluded
 *          by an active column filter. `visibleRows` is the compacted array of
 *          absolute row indices in the non-frozen band, hidden rows removed.
 * Stage 2: `windowRows` slices that compacted array from scrollTop, and carries
 *          the CSS grid line for each row: `frozenRows + visualIndex + 2`
 *          (1 for the header row, 1 because grid lines are 1-based).
 *
 * Unrendered rows simply do not exist as grid items — no spacer divs. That works
 * because `gridAutoRows` sizes implicit tracks and the browser allocates track
 * space up to the highest referenced `gridRow` line.
 *
 * Frozen rows are a SEPARATE always-rendered band and are assumed never hidden.
 * Adding column hiding, or allowing frozen rows to be filtered out, means
 * revisiting this mapping.
 */
import { useMemo } from "react";
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

  // `sheet.cells` is a required dependency even though it is never referenced
  // directly: getDisplay reads through it, so filter results go stale without it.
  // biome-ignore lint/correctness/useExhaustiveDependencies: sheet.cells is required
  const effectiveHiddenRows = useMemo(() => {
    const hidden = new Set(sheet.hiddenRows);
    const filterCols = Object.keys(sheet.filters).map(Number);
    if (filterCols.length > 0) {
      for (let r = frozenRows; r < sheet.numRows; r++) {
        if (hidden.has(r)) continue;
        for (const c of filterCols) {
          const allowed = sheet.filters[c];
          if (!allowed) continue;
          if (!allowed.has(getDisplay(r, c))) {
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

  // The scrollable band starts below the header and the frozen rows, both of
  // which are sticky and therefore do not scroll away.
  const bandScrollTop = Math.max(
    0,
    scrollTop - (HEADER_HEIGHT + frozenRows * ROW_HEIGHT),
  );
  const startVisual = clamp(
    Math.floor(bandScrollTop / ROW_HEIGHT) - BUFFER_ROWS,
    0,
    visibleRows.length,
  );
  const visibleCount = Math.ceil(viewportHeight / ROW_HEIGHT) + BUFFER_ROWS * 2;
  const endVisual = clamp(startVisual + visibleCount, 0, visibleRows.length);

  const windowRows = useMemo(
    () =>
      visibleRows.slice(startVisual, endVisual).map((absRow, i) => ({
        absRow,
        gridRow: frozenRows + startVisual + i + 2,
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

  return {
    effectiveHiddenRows,
    visibleRows,
    windowRows,
    frozenRowsList,
    startVisual,
  };
}
