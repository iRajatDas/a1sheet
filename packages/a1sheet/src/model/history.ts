/**
 * Undo/redo as whole-workbook snapshots, not diffs.
 *
 * Snapshots are cheap because `cloneSheet` is shallow per container and only the
 * edited sheet is ever cloned. Capped at HISTORY_LIMIT entries.
 */
import type { Workbook } from "./types.js";

export const HISTORY_LIMIT = 50;

export interface History {
  past: Workbook[];
  future: Workbook[];
}

export function emptyHistory(): History {
  return { past: [], future: [] };
}

/** Records a snapshot and clears the redo stack. */
export function push(history: History, snapshot: Workbook): History {
  const past = [...history.past, snapshot];
  if (past.length > HISTORY_LIMIT) past.shift();
  return { past, future: [] };
}

export function canUndo(history: History): boolean {
  return history.past.length > 0;
}

export function canRedo(history: History): boolean {
  return history.future.length > 0;
}

/** Returns null when there is nothing to undo. */
export function undo(
  history: History,
  current: Workbook,
): { history: History; workbook: Workbook } | null {
  const prev = history.past.at(-1);
  if (!prev) return null;
  return {
    history: {
      past: history.past.slice(0, -1),
      future: [current, ...history.future].slice(0, HISTORY_LIMIT),
    },
    workbook: prev,
  };
}

/** Returns null when there is nothing to redo. */
export function redo(
  history: History,
  current: Workbook,
): { history: History; workbook: Workbook } | null {
  const next = history.future[0];
  if (!next) return null;
  return {
    history: {
      past: [...history.past, current].slice(-HISTORY_LIMIT),
      future: history.future.slice(1),
    },
    workbook: next,
  };
}
