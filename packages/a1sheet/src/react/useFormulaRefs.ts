"use client";

/**
 * Reference picking while a formula is being typed.
 *
 * Two behaviors, one piece of state:
 *
 * - `spans` — every reference in the formula being edited, with the characters
 *   it occupies, so the grid can outline the cells it points at.
 * - `pickAt` / `extendPickTo` — clicking or dragging in the grid writes a
 *   reference into the formula instead of moving the selection.
 *
 * All the string work lives in ../formula/refEditing.ts, which is
 * framework-agnostic and separately tested. This hook is the React-side state
 * for one thing the pure functions cannot know: whether the current click is
 * starting a new reference or dragging the one it just wrote.
 */
import { useCallback, useMemo, useRef } from "react";
import {
  findRefSpans,
  insertRefAtCaret,
  isFormulaSource,
  type RefSpan,
} from "../formula/refEditing.js";
import { toA1 } from "../model/address.js";
import type { EditingState } from "./useEditing.js";

export interface UseFormulaRefsResult {
  /** References in the formula being edited. Empty when not editing one. */
  spans: readonly RefSpan[];
  /** True when the editor should treat grid clicks as reference picking. */
  active: boolean;
  /**
   * Handles a click on a cell. Returns true when it was consumed as a
   * reference; false means the caller should select the cell as normal.
   */
  pickAt(row: number, col: number): boolean;
  /** Grows the reference started by `pickAt` into a range. No-op otherwise. */
  extendPickTo(row: number, col: number): void;
  /** Forgets the in-progress pick, so the next click starts a new reference. */
  endPick(): void;
}

interface Pick {
  anchor: { row: number; col: number };
  /** Where the reference text sits, so a drag rewrites rather than appends. */
  span: { start: number; end: number };
}

/** "A1" for one cell, "A1:B2" for a range, anchored where the drag began. */
function refText(
  anchor: { row: number; col: number },
  row: number,
  col: number,
): string {
  const from = toA1(anchor.row, anchor.col);
  if (anchor.row === row && anchor.col === col) return from;
  return `${from}:${toA1(row, col)}`;
}

export function useFormulaRefs(
  editing: EditingState | null,
  setValue: (value: string, caret?: number) => void,
): UseFormulaRefsResult {
  const pick = useRef<Pick | null>(null);

  const source = editing?.value ?? "";
  const active = isFormulaSource(source);

  const spans = useMemo(
    () => (active ? findRefSpans(source) : []),
    [active, source],
  );

  const pickAt = useCallback(
    (row: number, col: number) => {
      if (!editing || !isFormulaSource(editing.value)) return false;

      const result = insertRefAtCaret(editing.value, editing.caret, toA1(row, col));
      // The caret was after a finished operand, so the click means "select that
      // cell", not "reference it". Say so and let the caller do its normal job.
      if (!result) return false;

      pick.current = { anchor: { row, col }, span: result.span };
      setValue(result.value, result.caret);
      return true;
    },
    [editing, setValue],
  );

  const extendPickTo = useCallback(
    (row: number, col: number) => {
      const current = pick.current;
      if (!current || !editing) return;

      const text = refText(current.anchor, row, col);
      const { start, end } = current.span;
      const next = editing.value.slice(0, start) + text + editing.value.slice(end);

      pick.current = { ...current, span: { start, end: start + text.length } };
      setValue(next, start + text.length);
    },
    [editing, setValue],
  );

  const endPick = useCallback(() => {
    pick.current = null;
  }, []);

  return { spans, active, pickAt, extendPickTo, endPick };
}
