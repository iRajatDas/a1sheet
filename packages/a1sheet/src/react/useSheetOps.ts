/**
 * Sheet operations exposed to the UI: formatting, structure, merges, freeze,
 * column width, sort, filter.
 *
 * Every one goes through the mutators it is handed — this hook never touches
 * workbook state directly.
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
   *
   * Pass `recordHistory: false` for live drags (color picker) after the first
   * event in a session has already recorded undo state.
   */
  applyStyle(patch: Partial<StyleObject>, recordHistory?: boolean): void;
  /** Removes one style field from every cell in the selection. */
  unsetStyle(key: keyof StyleObject): void;
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
  mutateSheet: (fn: SheetUpdater, addHistory?: boolean) => void,
  patchSurface: (fn: SheetPatcher, addHistory?: boolean) => void,
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
    (patch: Partial<StyleObject>, recordHistory = true) => {
      patchSurface((s) => {
        const styles = { ...s.styles };
        forEachCell(ranges, (r, c) => {
          const key = cellKey(r, c);
          styles[key] = { ...(styles[key] ?? {}), ...patch };
        });
        return { styles };
      }, recordHistory);
    },
    [ranges, patchSurface],
  );

  const clearCells = useCallback(() => {
    if (anyLockedInSelection()) return false;
    mutateSheet((s) => {
      forEachCell(ranges, (r, c) => {
        delete s.cells[cellKey(r, c)];
      });
      return s;
    });
    return true;
  }, [ranges, anyLockedInSelection, mutateSheet]);

  const clearFormatting = useCallback(() => {
    patchSurface((s) => {
      const styles = { ...s.styles };
      forEachCell(ranges, (r, c) => {
        delete styles[cellKey(r, c)];
      });
      return { styles };
    });
  }, [ranges, patchSurface]);

  const unsetStyle = useCallback(
    (key: keyof StyleObject) => {
      patchSurface((s) => {
        const styles = { ...s.styles };
        forEachCell(ranges, (r, c) => {
          const k = cellKey(r, c);
          const prev = styles[k];
          if (!prev || !(key in prev)) return;
          const next = { ...prev };
          delete next[key];
          if (Object.keys(next).length === 0) delete styles[k];
          else styles[k] = next;
        });
        return { styles };
      });
    },
    [ranges, patchSurface],
  );

  return {
    activeStyle: sheet.styles[cellKey(selection.r2, selection.c2)] ?? {},
    isLocked,
    anyLockedInSelection,
    applyStyle,
    unsetStyle,
    clearCells,
    clearFormatting,

    mergeSelection: useCallback(() => {
      if (bounds.r1 === bounds.r2 && bounds.c1 === bounds.c2) return;
      patchSurface((s) => ({ merges: [...s.merges, { ...bounds }] }));
    }, [bounds, patchSurface]),

    unmergeSelection: useCallback(() => {
      patchSurface((s) => ({
        merges: s.merges.filter(
          (m) =>
            m.r2 < bounds.r1 ||
            m.r1 > bounds.r2 ||
            m.c2 < bounds.c1 ||
            m.c1 > bounds.c2,
        ),
      }));
    }, [bounds, patchSurface]),

    // Freezes up through the active cell, which is what the toolbar means by it.
    freezeToSelection: useCallback(() => {
      patchSurface(
        () => ({ frozenRows: selection.r2 + 1, frozenCols: selection.c2 + 1 }),
        false,
      );
    }, [selection.r2, selection.c2, patchSurface]),

    unfreeze: useCallback(() => {
      patchSurface(() => ({ frozenRows: 0, frozenCols: 0 }), false);
    }, [patchSurface]),

    insertRowAt: useCallback(
      (row: number) => mutateSheet((s) => insertRow(s, row)),
      [mutateSheet],
    ),
    deleteRowAt: useCallback(
      (row: number) => mutateSheet((s) => deleteRow(s, row)),
      [mutateSheet],
    ),
    insertColAt: useCallback(
      (col: number) => mutateSheet((s) => insertCol(s, col)),
      [mutateSheet],
    ),
    deleteColAt: useCallback(
      (col: number) => mutateSheet((s) => deleteCol(s, col)),
      [mutateSheet],
    ),

    // Nothing to shift: the new rows are past everything that exists, so no
    // cell, style, height, or merge moves.
    appendRows: useCallback(
      (count: number) =>
        patchSurface((s) => ({
          numRows: s.numRows + Math.max(0, Math.floor(count)),
        })),
      [patchSurface],
    ),

    // No history: a drag would otherwise push one entry per mousemove.
    setColWidth: useCallback(
      (col: number, px: number) =>
        patchSurface(
          (s) => ({
            colWidths: { ...s.colWidths, [col]: clampSize(px, MIN_COL_WIDTH) },
          }),
          false,
        ),
      [patchSurface],
    ),
    setRowHeight: useCallback(
      (row: number, px: number) =>
        patchSurface(
          (s) => ({
            rowHeights: { ...s.rowHeights, [row]: clampSize(px, MIN_ROW_HEIGHT) },
          }),
          false,
        ),
      [patchSurface],
    ),

    // Deleting the entry, rather than writing the default into it, is what
    // makes a reset free: `useRowWindow` skips its offset table entirely while
    // `rowHeights` is empty, and it can only become empty again this way.
    resetRowHeight: useCallback(
      (row: number) =>
        patchSurface((s) =>
          row in s.rowHeights ? { rowHeights: without(s.rowHeights, row) } : null,
        ),
      [patchSurface],
    ),
    resetColWidth: useCallback(
      (col: number) =>
        patchSurface((s) =>
          col in s.colWidths ? { colWidths: without(s.colWidths, col) } : null,
        ),
      [patchSurface],
    ),

    setRowLabel: useCallback(
      (row: number, label: string) =>
        patchSurface((s) => ({ rowLabels: { ...s.rowLabels, [row]: label } }), false),
      [patchSurface],
    ),
    setColLabel: useCallback(
      (col: number, label: string) =>
        patchSurface((s) => ({ colLabels: { ...s.colLabels, [col]: label } }), false),
      [patchSurface],
    ),

    toggleColHidden: useCallback(
      (col: number) =>
        patchSurface((s) => ({ hiddenCols: toggled(s.hiddenCols, col) })),
      [patchSurface],
    ),

    toggleRowHidden: useCallback(
      (row: number) =>
        patchSurface((s) => ({ hiddenRows: toggled(s.hiddenRows, row) })),
      [patchSurface],
    ),

    sort: useCallback(
      (col: number, dir: "asc" | "desc") =>
        mutateSheet((s) => sortByColumn(s, col, dir)),
      [mutateSheet],
    ),

    setFilter: useCallback(
      (col: number, allowed: Set<string> | null) =>
        patchSurface((s) => {
          if (allowed !== null)
            return { filters: { ...s.filters, [col]: allowed } };
          return col in s.filters ? { filters: without(s.filters, col) } : null;
        }, false),
      [patchSurface],
    ),
  };
}
