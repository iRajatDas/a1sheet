/**
 * Sheet operations exposed to the UI: formatting, structure, merges, freeze,
 * column width, sort, filter.
 *
 * Every one goes through the `updateSheet` it is handed — this hook never
 * touches workbook state directly.
 */
import { useCallback, useMemo } from "react";
import { cellKey, normalizeRange } from "../model/address.js";
import {
  deleteCol,
  deleteRow,
  insertCol,
  insertRow,
  sortByColumn,
} from "../model/sheet.js";
import type { Range, Sheet, StyleObject } from "../model/types.js";
import { MIN_COL_WIDTH, MIN_ROW_HEIGHT } from "./constants.js";
import type { SheetPatcher, SheetUpdater } from "./useWorkbook.js";

/** A resize drag can go anywhere, including negative. Rounding keeps offsets integral. */
const clampSize = (px: number, min: number) => Math.max(min, Math.round(px));

export interface UseSheetOpsResult {
  /** Style of the active cell, or an empty object. Drives toolbar toggle states. */
  activeStyle: StyleObject;
  isLocked(row: number, col: number): boolean;
  anyLockedInSelection(): boolean;
  /**
   * Merges a style patch into every selected cell — every range of a
   * multi-selection, not just the primary one. Formatting three separate blocks
   * bold at once is most of the point of Ctrl+click.
   */
  applyStyle(patch: Partial<StyleObject>): void;
  /** Deletes cell contents in every selected range. Refuses if any cell is locked. */
  clearCells(): boolean;
  clearFormatting(): void;
  mergeSelection(): void;
  unmergeSelection(): void;
  freezeToSelection(): void;
  unfreeze(): void;
  insertRowAt(row: number): void;
  deleteRowAt(row: number): void;
  insertColAt(col: number): void;
  deleteColAt(col: number): void;
  /** Grows the sheet by `count` empty rows at the bottom. */
  appendRows(count: number): void;
  setColWidth(col: number, px: number): void;
  setRowHeight(row: number, px: number): void;
  /** Drops the override, returning the row to the default height. */
  resetRowHeight(row: number): void;
  /** Drops the override, returning the column to the default width. */
  resetColWidth(col: number): void;
  setRowLabel(row: number, label: string): void;
  setColLabel(col: number, label: string): void;
  toggleRowHidden(row: number): void;
  toggleColHidden(col: number): void;
  sort(col: number, dir: "asc" | "desc"): void;
  /** Passing null clears the filter for that column. */
  setFilter(col: number, allowed: Set<string> | null): void;
}

/**
 * Visits every cell of every selected range, once.
 *
 * Ranges of a Ctrl+click selection may overlap, and `styles[key] = {...}` twice
 * is only wasted work — but `forEachCell` is also what a future op that is not
 * idempotent would use, so it dedupes here rather than leaving each caller to
 * remember.
 */
function forEachCell(
  ranges: readonly Range[],
  visit: (row: number, col: number) => void,
): void {
  const seen = new Set<string>();
  for (const range of ranges) {
    const b = normalizeRange(range);
    for (let r = b.r1; r <= b.r2; r++) {
      for (let c = b.c1; c <= b.c2; c++) {
        const key = cellKey(r, c);
        if (seen.has(key)) continue;
        seen.add(key);
        visit(r, c);
      }
    }
  }
}

/**
 * @param ranges Every selected range, the primary one first. Cell operations
 * cover all of them; structural ones — merge, freeze, insert — use the primary,
 * because "insert a row at three disjoint places" has no obvious meaning.
 */
/** A copy of `record` without one key. Deleting an entry is what resets a size. */
function without<T>(record: Record<number, T>, key: number): Record<number, T> {
  const out = { ...record };
  delete out[key];
  return out;
}

/** A copy of `set` with `value` added or removed. */
function toggled(set: ReadonlySet<number>, value: number): Set<number> {
  const out = new Set(set);
  if (!out.delete(value)) out.add(value);
  return out;
}

export function useSheetOps(
  sheet: Sheet,
  ranges: readonly Range[],
  updateSheet: (fn: SheetUpdater, addHistory?: boolean) => void,
  patchSheet: (fn: SheetPatcher, addHistory?: boolean) => void,
): UseSheetOpsResult {
  const selection = ranges[0] ?? { r1: 0, c1: 0, r2: 0, c2: 0 };
  const bounds = useMemo(() => normalizeRange(selection), [selection]);

  const isLocked = useCallback(
    (row: number, col: number) => !!sheet.styles[cellKey(row, col)]?.locked,
    [sheet.styles],
  );

  const anyLockedInSelection = useCallback(() => {
    let locked = false;
    forEachCell(ranges, (r, c) => {
      if (isLocked(r, c)) locked = true;
    });
    return locked;
  }, [ranges, isLocked]);

  const applyStyle = useCallback(
    (patch: Partial<StyleObject>) => {
      updateSheet((s) => {
        forEachCell(ranges, (r, c) => {
          const key = cellKey(r, c);
          s.styles[key] = { ...(s.styles[key] ?? {}), ...patch };
        });
        return s;
      });
    },
    [ranges, updateSheet],
  );

  const clearCells = useCallback(() => {
    if (anyLockedInSelection()) return false;
    updateSheet((s) => {
      forEachCell(ranges, (r, c) => {
        delete s.cells[cellKey(r, c)];
      });
      return s;
    });
    return true;
  }, [ranges, anyLockedInSelection, updateSheet]);

  const clearFormatting = useCallback(() => {
    updateSheet((s) => {
      forEachCell(ranges, (r, c) => {
        delete s.styles[cellKey(r, c)];
      });
      return s;
    });
  }, [ranges, updateSheet]);

  return {
    activeStyle: sheet.styles[cellKey(selection.r2, selection.c2)] ?? {},
    isLocked,
    anyLockedInSelection,
    applyStyle,
    clearCells,
    clearFormatting,

    mergeSelection: useCallback(() => {
      if (bounds.r1 === bounds.r2 && bounds.c1 === bounds.c2) return;
      patchSheet((s) => ({ merges: [...s.merges, { ...bounds }] }));
    }, [bounds, patchSheet]),

    unmergeSelection: useCallback(() => {
      patchSheet((s) => ({
        merges: s.merges.filter(
          (m) =>
            m.r2 < bounds.r1 ||
            m.r1 > bounds.r2 ||
            m.c2 < bounds.c1 ||
            m.c1 > bounds.c2,
        ),
      }));
    }, [bounds, patchSheet]),

    // Freezes up through the active cell, which is what the toolbar means by it.
    freezeToSelection: useCallback(() => {
      patchSheet(
        () => ({ frozenRows: selection.r2 + 1, frozenCols: selection.c2 + 1 }),
        false,
      );
    }, [selection.r2, selection.c2, patchSheet]),

    unfreeze: useCallback(() => {
      patchSheet(() => ({ frozenRows: 0, frozenCols: 0 }), false);
    }, [patchSheet]),

    insertRowAt: useCallback(
      (row: number) => updateSheet((s) => insertRow(s, row)),
      [updateSheet],
    ),
    deleteRowAt: useCallback(
      (row: number) => updateSheet((s) => deleteRow(s, row)),
      [updateSheet],
    ),
    insertColAt: useCallback(
      (col: number) => updateSheet((s) => insertCol(s, col)),
      [updateSheet],
    ),
    deleteColAt: useCallback(
      (col: number) => updateSheet((s) => deleteCol(s, col)),
      [updateSheet],
    ),

    // Nothing to shift: the new rows are past everything that exists, so no
    // cell, style, height, or merge moves.
    appendRows: useCallback(
      (count: number) =>
        patchSheet((s) => ({
          numRows: s.numRows + Math.max(0, Math.floor(count)),
        })),
      [patchSheet],
    ),

    // No history: a drag would otherwise push one entry per mousemove.
    setColWidth: useCallback(
      (col: number, px: number) =>
        patchSheet(
          (s) => ({
            colWidths: { ...s.colWidths, [col]: clampSize(px, MIN_COL_WIDTH) },
          }),
          false,
        ),
      [patchSheet],
    ),
    setRowHeight: useCallback(
      (row: number, px: number) =>
        patchSheet(
          (s) => ({
            rowHeights: { ...s.rowHeights, [row]: clampSize(px, MIN_ROW_HEIGHT) },
          }),
          false,
        ),
      [patchSheet],
    ),

    // Deleting the entry, rather than writing the default into it, is what
    // makes a reset free: `useRowWindow` skips its offset table entirely while
    // `rowHeights` is empty, and it can only become empty again this way.
    resetRowHeight: useCallback(
      (row: number) =>
        patchSheet((s) =>
          row in s.rowHeights ? { rowHeights: without(s.rowHeights, row) } : null,
        ),
      [patchSheet],
    ),
    resetColWidth: useCallback(
      (col: number) =>
        patchSheet((s) =>
          col in s.colWidths ? { colWidths: without(s.colWidths, col) } : null,
        ),
      [patchSheet],
    ),

    setRowLabel: useCallback(
      (row: number, label: string) =>
        patchSheet((s) => ({ rowLabels: { ...s.rowLabels, [row]: label } }), false),
      [patchSheet],
    ),
    setColLabel: useCallback(
      (col: number, label: string) =>
        patchSheet((s) => ({ colLabels: { ...s.colLabels, [col]: label } }), false),
      [patchSheet],
    ),

    toggleColHidden: useCallback(
      (col: number) =>
        patchSheet((s) => ({ hiddenCols: toggled(s.hiddenCols, col) })),
      [patchSheet],
    ),

    toggleRowHidden: useCallback(
      (row: number) =>
        patchSheet((s) => ({ hiddenRows: toggled(s.hiddenRows, row) })),
      [patchSheet],
    ),

    sort: useCallback(
      (col: number, dir: "asc" | "desc") =>
        updateSheet((s) => sortByColumn(s, col, dir)),
      [updateSheet],
    ),

    setFilter: useCallback(
      (col: number, allowed: Set<string> | null) =>
        patchSheet((s) => {
          if (allowed !== null)
            return { filters: { ...s.filters, [col]: allowed } };
          return col in s.filters ? { filters: without(s.filters, col) } : null;
        }, false),
      [patchSheet],
    ),
  };
}
