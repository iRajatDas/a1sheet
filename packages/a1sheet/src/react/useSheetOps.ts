/**
 * Sheet operations exposed to the UI: formatting, structure, merges, freeze,
 * column width, sort, filter.
 *
 * Every one goes through the `updateSheet` it is handed — this hook never touches
 * workbook state directly. Ported from the corresponding handlers in
 * ref/Spreadsheet.jsx.
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
import type { SheetUpdater } from "./useWorkbook.js";

/** A resize drag can go anywhere, including negative. Rounding keeps offsets integral. */
const clampSize = (px: number, min: number) => Math.max(min, Math.round(px));

export interface UseSheetOpsResult {
  /** Style of the active cell, or an empty object. Drives toolbar toggle states. */
  activeStyle: StyleObject;
  isLocked(row: number, col: number): boolean;
  anyLockedInSelection(): boolean;
  /** Merges a style patch into every cell of the selection. */
  applyStyle(patch: Partial<StyleObject>): void;
  /** Deletes cell contents in the selection. Refuses if any cell is locked. */
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

export function useSheetOps(
  sheet: Sheet,
  selection: Range,
  updateSheet: (fn: SheetUpdater, addHistory?: boolean) => void,
): UseSheetOpsResult {
  const bounds = useMemo(() => normalizeRange(selection), [selection]);

  const isLocked = useCallback(
    (row: number, col: number) => !!sheet.styles[cellKey(row, col)]?.locked,
    [sheet.styles],
  );

  const anyLockedInSelection = useCallback(() => {
    for (let r = bounds.r1; r <= bounds.r2; r++) {
      for (let c = bounds.c1; c <= bounds.c2; c++) {
        if (isLocked(r, c)) return true;
      }
    }
    return false;
  }, [bounds, isLocked]);

  const applyStyle = useCallback(
    (patch: Partial<StyleObject>) => {
      updateSheet((s) => {
        for (let r = bounds.r1; r <= bounds.r2; r++) {
          for (let c = bounds.c1; c <= bounds.c2; c++) {
            const key = cellKey(r, c);
            s.styles[key] = { ...(s.styles[key] ?? {}), ...patch };
          }
        }
        return s;
      });
    },
    [bounds, updateSheet],
  );

  const clearCells = useCallback(() => {
    if (anyLockedInSelection()) return false;
    updateSheet((s) => {
      for (let r = bounds.r1; r <= bounds.r2; r++) {
        for (let c = bounds.c1; c <= bounds.c2; c++) delete s.cells[cellKey(r, c)];
      }
      return s;
    });
    return true;
  }, [bounds, anyLockedInSelection, updateSheet]);

  const clearFormatting = useCallback(() => {
    updateSheet((s) => {
      for (let r = bounds.r1; r <= bounds.r2; r++) {
        for (let c = bounds.c1; c <= bounds.c2; c++) delete s.styles[cellKey(r, c)];
      }
      return s;
    });
  }, [bounds, updateSheet]);

  return {
    activeStyle: sheet.styles[cellKey(selection.r2, selection.c2)] ?? {},
    isLocked,
    anyLockedInSelection,
    applyStyle,
    clearCells,
    clearFormatting,

    mergeSelection: useCallback(() => {
      if (bounds.r1 === bounds.r2 && bounds.c1 === bounds.c2) return;
      updateSheet((s) => {
        s.merges = [...s.merges, { ...bounds }];
        return s;
      });
    }, [bounds, updateSheet]),

    unmergeSelection: useCallback(() => {
      updateSheet((s) => {
        s.merges = s.merges.filter(
          (m) =>
            m.r2 < bounds.r1 ||
            m.r1 > bounds.r2 ||
            m.c2 < bounds.c1 ||
            m.c1 > bounds.c2,
        );
        return s;
      });
    }, [bounds, updateSheet]),

    // Freezes up through the active cell, matching the POC's toolbar behavior.
    freezeToSelection: useCallback(() => {
      updateSheet((s) => {
        s.frozenRows = selection.r2 + 1;
        s.frozenCols = selection.c2 + 1;
        return s;
      }, false);
    }, [selection.r2, selection.c2, updateSheet]),

    unfreeze: useCallback(() => {
      updateSheet((s) => {
        s.frozenRows = 0;
        s.frozenCols = 0;
        return s;
      }, false);
    }, [updateSheet]),

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
        updateSheet((s) => {
          s.numRows += Math.max(0, Math.floor(count));
          return s;
        }),
      [updateSheet],
    ),

    // No history: a drag would otherwise push one entry per mousemove.
    setColWidth: useCallback(
      (col: number, px: number) =>
        updateSheet((s) => {
          s.colWidths[col] = clampSize(px, MIN_COL_WIDTH);
          return s;
        }, false),
      [updateSheet],
    ),
    setRowHeight: useCallback(
      (row: number, px: number) =>
        updateSheet((s) => {
          s.rowHeights[row] = clampSize(px, MIN_ROW_HEIGHT);
          return s;
        }, false),
      [updateSheet],
    ),

    // Deleting the entry, rather than writing the default into it, is what
    // makes a reset free: `useRowWindow` skips its offset table entirely while
    // `rowHeights` is empty, and it can only become empty again this way.
    resetRowHeight: useCallback(
      (row: number) =>
        updateSheet((s) => {
          delete s.rowHeights[row];
          return s;
        }),
      [updateSheet],
    ),
    resetColWidth: useCallback(
      (col: number) =>
        updateSheet((s) => {
          delete s.colWidths[col];
          return s;
        }),
      [updateSheet],
    ),

    setRowLabel: useCallback(
      (row: number, label: string) =>
        updateSheet((s) => {
          s.rowLabels[row] = label;
          return s;
        }, false),
      [updateSheet],
    ),
    setColLabel: useCallback(
      (col: number, label: string) =>
        updateSheet((s) => {
          s.colLabels[col] = label;
          return s;
        }, false),
      [updateSheet],
    ),

    toggleColHidden: useCallback(
      (col: number) =>
        updateSheet((s) => {
          if (s.hiddenCols.has(col)) s.hiddenCols.delete(col);
          else s.hiddenCols.add(col);
          return s;
        }),
      [updateSheet],
    ),

    toggleRowHidden: useCallback(
      (row: number) =>
        updateSheet((s) => {
          if (s.hiddenRows.has(row)) s.hiddenRows.delete(row);
          else s.hiddenRows.add(row);
          return s;
        }),
      [updateSheet],
    ),

    sort: useCallback(
      (col: number, dir: "asc" | "desc") =>
        updateSheet((s) => sortByColumn(s, col, dir)),
      [updateSheet],
    ),

    setFilter: useCallback(
      (col: number, allowed: Set<string> | null) =>
        updateSheet((s) => {
          if (allowed === null) delete s.filters[col];
          else s.filters[col] = allowed;
          return s;
        }, false),
      [updateSheet],
    ),
  };
}
