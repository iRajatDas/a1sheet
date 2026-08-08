"use client";

/**
 * Workbook state and the ONLY sanctioned mutation path.
 *
 * `updateSheet` and `updateWorkbook` clone-on-write and push history. Every other
 * hook receives these as arguments and has no direct access to the workbook
 * setter, which makes "never mutate a sheet in place" a structural property
 * rather than a rule someone has to remember.
 */
import { useCallback, useMemo, useRef, useState } from "react";
import { EmptyWorkbookError } from "../errors.js";
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
  duplicateSheet,
  moveSheet,
  renameSheet,
} from "../model/workbook.js";

export type SheetUpdater = (sheet: Sheet) => Sheet;
/** Returns the fields to replace on the active sheet, or null to change nothing. */
export type SheetPatcher = (sheet: Sheet) => Partial<Sheet> | null;
export type WorkbookUpdater = (wb: Workbook) => Workbook;

export interface UseWorkbookResult {
  workbook: Workbook;
  sheet: Sheet;
  /** Applies `fn` to the active sheet. Pass `addHistory: false` for transient edits. */
  updateSheet(fn: SheetUpdater, addHistory?: boolean): void;
  /**
   * Replaces named fields of the active sheet without cloning the rest of it.
   *
   * `updateSheet` hands the updater a full clone, which copies `cells` — 66 ms of
   * the 70 a clone costs at a million filled cells. An operation that touches
   * only a row/column-keyed container (a width, a label, a filter, the frozen
   * counts) does not need that copy, and some of them fire on every mousemove of
   * a resize drag. `fn` returns the fields it is changing, already copied, or
   * null for no change.
   */
  patchSheet(fn: SheetPatcher, addHistory?: boolean): void;
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
  /** Deep-copies a sheet tab (filters/styles included) and activates the copy. */
  duplicateSheetAt(index: number): void;
  /** Reorders sheet tabs. */
  moveSheetAt(from: number, to: number): void;
}

export interface UseWorkbookOptions {
  /** Uncontrolled starting value. Ignored when `workbook` is supplied. */
  initialWorkbook?: Workbook;
  /**
   * Controlled value. When present the hook never holds its own workbook state and
   * every change is reported through `onChange` for the caller to apply.
   */
  workbook?: Workbook;
  onChange?: (wb: Workbook) => void;
}

export function useWorkbook(opts: UseWorkbookOptions = {}): UseWorkbookResult {
  const isControlled = opts.workbook !== undefined;

  const [uncontrolled, setUncontrolled] = useState<Workbook>(
    () => opts.initialWorkbook ?? createWorkbook(),
  );
  const [history, setHistory] = useState<History>(emptyHistory);

  const workbook = isControlled ? (opts.workbook as Workbook) : uncontrolled;

  // Kept in a ref so the callbacks below stay referentially stable across
  // renders — they are passed down to every child hook and component.
  const onChange = useRef(opts.onChange);
  onChange.current = opts.onChange;

  // Mirrors the live workbook so the updater callbacks can read it without taking
  // it as a dependency, which would break their referential stability. Needed in
  // controlled mode, where there is no setState updater to receive `prev`.
  const latest = useRef(workbook);
  latest.current = workbook;

  /**
   * The single write path. In uncontrolled mode this drives internal state; in
   * controlled mode it only reports upward and the caller re-renders us.
   */
  const setWorkbook = useCallback(
    (compute: (prev: Workbook) => Workbook) => {
      if (isControlled) {
        const prev = latest.current;
        const next = compute(prev);
        if (next !== prev) onChange.current?.(next);
        return;
      }
      setUncontrolled(compute);
    },
    [isControlled],
  );

  const updateWorkbook = useCallback(
    (fn: WorkbookUpdater, addHistory = true) => {
      setWorkbook((prev) => {
        const next = fn(prev);
        if (next === prev) return prev;
        if (addHistory) setHistory((h) => histPush(h, prev));
        if (!isControlled) onChange.current?.(next);
        return next;
      });
    },
    [setWorkbook, isControlled],
  );

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

  const patchSheet = useCallback(
    (fn: SheetPatcher, addHistory = true) => {
      updateWorkbook((prev) => {
        const active = prev.sheets[prev.activeSheetIndex];
        if (!active) return prev;
        const patch = fn(active);
        if (patch === null) return prev;
        return {
          ...prev,
          sheets: prev.sheets.map((s, i) =>
            i === prev.activeSheetIndex ? { ...s, ...patch } : s,
          ),
        };
      }, addHistory);
    },
    [updateWorkbook],
  );

  const replaceWorkbook = useCallback(
    (next: Workbook) => updateWorkbook(() => next, true),
    [updateWorkbook],
  );

  const undo = useCallback(() => {
    setWorkbook((current) => {
      const r = histUndo(history, current);
      if (!r) return current;
      setHistory(r.history);
      if (!isControlled) onChange.current?.(r.workbook);
      return r.workbook;
    });
  }, [history, setWorkbook, isControlled]);

  const redo = useCallback(() => {
    setWorkbook((current) => {
      const r = histRedo(history, current);
      if (!r) return current;
      setHistory(r.history);
      if (!isControlled) onChange.current?.(r.workbook);
      return r.workbook;
    });
  }, [history, setWorkbook, isControlled]);

  const sheet = useMemo(() => {
    const s = workbook.sheets[workbook.activeSheetIndex] ?? workbook.sheets[0];
    if (!s) throw new EmptyWorkbookError();
    return s;
  }, [workbook]);

  return {
    workbook,
    sheet,
    updateSheet,
    patchSheet,
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
    duplicateSheetAt: useCallback(
      (index: number) => updateWorkbook((wb) => duplicateSheet(wb, index)),
      [updateWorkbook],
    ),
    moveSheetAt: useCallback(
      (from: number, to: number) => updateWorkbook((wb) => moveSheet(wb, from, to)),
      [updateWorkbook],
    ),
  };
}
