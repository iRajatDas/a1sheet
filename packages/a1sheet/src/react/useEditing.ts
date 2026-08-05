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
}

export interface UseEditingResult {
  editing: EditingState | null;
  isEditing: boolean;
  /** No-ops on a locked cell. `seed` replaces the content; omit to edit in place. */
  startEdit(sheet: Sheet, row: number, col: number, seed?: string): void;
  setValue(value: string): void;
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
      set({ row, col, value: seed ?? sheet.cells[cellKey(row, col)] ?? "" });
    },
    [set],
  );

  const setValue = useCallback(
    (value: string) => {
      if (!ref.current) return;
      set({ ...ref.current, value });
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
    commit,
    cancel,
  };
}
