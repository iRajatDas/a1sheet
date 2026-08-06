/**
 * In-progress cell edit.
 *
 * One piece of state shared by the in-cell input AND the formula bar, so the two
 * cannot drift out of sync. Typing in either calls the same `setValue`.
 *
 * `locked` cells cannot be edited (or cleared) but remain selectable and
 * copyable — `startEdit` enforces that.
 */
import { useCallback, useRef, useState } from "react";
import { cellKey } from "../model/address.js";
import type { Sheet } from "../model/types.js";

export interface EditingState {
  row: number;
  col: number;
  value: string;
  /**
   * Caret offset within `value`.
   *
   * Part of the edit rather than of the input element because clicking a cell
   * mid-formula has to decide what to do from where the caret is, and by then
   * the input has lost focus to the click. Both editors report it back on every
   * keystroke and selection change.
   */
  caret: number;
}

export interface UseEditingResult {
  editing: EditingState | null;
  isEditing: boolean;
  /** No-ops on a locked cell. `seed` replaces the content; omit to edit in place. */
  startEdit(sheet: Sheet, row: number, col: number, seed?: string): void;
  /** Sets the text. The caret goes to the end unless `caret` says otherwise. */
  setValue(value: string, caret?: number): void;
  /** Reports where the caret moved to, without changing the text. */
  setCaret(caret: number): void;
  /** Returns the committed value and clears editing state; null if nothing was open. */
  commit(): EditingState | null;
  cancel(): void;
}

export function useEditing(): UseEditingResult {
  const [editing, setEditing] = useState<EditingState | null>(null);

  // `commit` must return the pending value synchronously to its caller (the
  // keydown handler needs it to write the cell). A setState updater cannot be
  // relied on to run synchronously, so the current value is mirrored here.
  const ref = useRef<EditingState | null>(null);

  const set = useCallback((next: EditingState | null) => {
    ref.current = next;
    setEditing(next);
  }, []);

  const startEdit = useCallback(
    (sheet: Sheet, row: number, col: number, seed?: string) => {
      if (sheet.styles[cellKey(row, col)]?.locked) return;
      const value = seed ?? sheet.cells[cellKey(row, col)] ?? "";
      set({ row, col, value, caret: value.length });
    },
    [set],
  );

  const setValue = useCallback(
    (value: string, caret?: number) => {
      if (!ref.current) return;
      set({ ...ref.current, value, caret: caret ?? value.length });
    },
    [set],
  );

  const setCaret = useCallback(
    (caret: number) => {
      if (!ref.current || ref.current.caret === caret) return;
      set({ ...ref.current, caret });
    },
    [set],
  );

  const commit = useCallback(() => {
    const committed = ref.current;
    set(null);
    return committed;
  }, [set]);

  const cancel = useCallback(() => set(null), [set]);

  return {
    editing,
    isEditing: editing !== null,
    startEdit,
    setValue,
    setCaret,
    commit,
    cancel,
  };
}
