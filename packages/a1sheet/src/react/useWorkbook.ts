/**
 * Workbook state and the ONLY sanctioned mutation path.
 *
 * `updateSheet` and `updateWorkbook` clone-on-write and push history. Every other
 * hook receives these as arguments and has no direct access to the workbook
 * setter, which makes "never mutate a sheet in place" a structural property
 * rather than a rule someone has to remember.
 */
import { useCallback, useMemo, useRef, useState } from "react";
import {
  emptyHistory,
  type History,
  canRedo as histCanRedo,
  canUndo as histCanUndo,
  push as histPush,
  redo as histRedo,
  undo as histUndo,
} from "../model/history.js";
import { cloneSheet } from "../model/sheet.js";
import type { Range, Sheet, Workbook } from "../model/types.js";
import {
  addSheet,
  createWorkbook,
  defineName as defineNameIn,
  deleteName,
  deleteSheet,
  renameSheet,
} from "../model/workbook.js";

export type SheetUpdater = (sheet: Sheet) => Sheet;
export type WorkbookUpdater = (wb: Workbook) => Workbook;

export interface UseWorkbookResult {
  workbook: Workbook;
  sheet: Sheet;
  /** Applies `fn` to the active sheet. Pass `addHistory: false` for transient edits. */
  updateSheet(fn: SheetUpdater, addHistory?: boolean): void;
  updateWorkbook(fn: WorkbookUpdater, addHistory?: boolean): void;
  /** Replaces the whole workbook, e.g. after a file import. Always recorded. */
  replaceWorkbook(next: Workbook): void;
  undo(): void;
  redo(): void;
  canUndo: boolean;
  canRedo: boolean;
  /** Appends a sheet and makes it active. */
  addSheetAt(name?: string): void;
  /** No-op when only one sheet remains. */
  deleteSheetAt(index: number): void;
  renameSheetAt(index: number, name: string): void;
  defineName(name: string, range: Range): void;
  deleteNamedRange(name: string): void;
}

export interface UseWorkbookOptions {
  initialWorkbook?: Workbook;
  onChange?: (wb: Workbook) => void;
}

export function useWorkbook(opts: UseWorkbookOptions = {}): UseWorkbookResult {
  const [workbook, setWorkbook] = useState<Workbook>(
    () => opts.initialWorkbook ?? createWorkbook(),
  );
  const [history, setHistory] = useState<History>(emptyHistory);

  // Kept in a ref so the callbacks below stay referentially stable across
  // renders — they are passed down to every child hook and component.
  const onChange = useRef(opts.onChange);
  onChange.current = opts.onChange;

  const commit = useCallback(
    (next: Workbook, prev: Workbook, addHistory: boolean) => {
      if (addHistory) setHistory((h) => histPush(h, prev));
      setWorkbook(next);
      onChange.current?.(next);
    },
    [],
  );

  const updateWorkbook = useCallback((fn: WorkbookUpdater, addHistory = true) => {
    setWorkbook((prev) => {
      const next = fn(prev);
      if (next === prev) return prev;
      if (addHistory) setHistory((h) => histPush(h, prev));
      onChange.current?.(next);
      return next;
    });
  }, []);

  const updateSheet = useCallback(
    (fn: SheetUpdater, addHistory = true) => {
      updateWorkbook((prev) => {
        const active = prev.sheets[prev.activeSheetIndex];
        if (!active) return prev;
        const next = fn(cloneSheet(active));
        if (next === active) return prev;
        return {
          ...prev,
          sheets: prev.sheets.map((s, i) =>
            i === prev.activeSheetIndex ? next : s,
          ),
        };
      }, addHistory);
    },
    [updateWorkbook],
  );

  const replaceWorkbook = useCallback(
    (next: Workbook) => {
      setWorkbook((prev) => {
        commit(next, prev, true);
        return next;
      });
    },
    [commit],
  );

  const undo = useCallback(() => {
    setWorkbook((current) => {
      const r = histUndo(history, current);
      if (!r) return current;
      setHistory(r.history);
      onChange.current?.(r.workbook);
      return r.workbook;
    });
  }, [history]);

  const redo = useCallback(() => {
    setWorkbook((current) => {
      const r = histRedo(history, current);
      if (!r) return current;
      setHistory(r.history);
      onChange.current?.(r.workbook);
      return r.workbook;
    });
  }, [history]);

  const sheet = useMemo(() => {
    const s = workbook.sheets[workbook.activeSheetIndex] ?? workbook.sheets[0];
    if (!s) throw new Error("workbook has no sheets");
    return s;
  }, [workbook]);

  return {
    workbook,
    sheet,
    updateSheet,
    updateWorkbook,
    replaceWorkbook,
    undo,
    redo,
    canUndo: histCanUndo(history),
    canRedo: histCanRedo(history),

    addSheetAt: useCallback(
      (name?: string) => updateWorkbook((wb) => addSheet(wb, name)),
      [updateWorkbook],
    ),
    deleteSheetAt: useCallback(
      (index: number) => updateWorkbook((wb) => deleteSheet(wb, index)),
      [updateWorkbook],
    ),
    renameSheetAt: useCallback(
      (index: number, name: string) =>
        updateWorkbook((wb) => renameSheet(wb, index, name), false),
      [updateWorkbook],
    ),
    defineName: useCallback(
      (name: string, range: Range) =>
        updateWorkbook((wb) => defineNameIn(wb, name, range)),
      [updateWorkbook],
    ),
    deleteNamedRange: useCallback(
      (name: string) => updateWorkbook((wb) => deleteName(wb, name)),
      [updateWorkbook],
    ),
  };
}
