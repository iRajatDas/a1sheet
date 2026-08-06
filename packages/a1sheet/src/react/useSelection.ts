/**
 * Selection state. Component-local — selection is never part of the workbook, so
 * undo does not restore it.
 *
 * `extraRanges` holds Ctrl+click history. It feeds the status bar ONLY: copy,
 * fill, and paste all act on `selection`. That is a deliberate scope cut.
 */
import { useCallback, useState } from "react";
import { normalizeRange } from "../model/address.js";
import type { Range } from "../model/types.js";

export interface UseSelectionResult {
  /**
   * `r1`/`c1` is the anchor, `r2`/`c2` the moving end of the drag. Not normalized,
   * so `r2 < r1` when the drag went upward. Use `bounds` for a normalized rect and
   * `active` for the one cell that typing goes to.
   */
  selection: Range;
  /**
   * The active cell — where typing goes and what the formula bar shows.
   *
   * This is the ANCHOR, not the moving end, matching Excel and Google Sheets:
   * drag D5→F13 and D5 stays the active cell, rendered unfilled inside the tinted
   * range. Reading `selection.r2`/`c2` instead is a bug — that follows the mouse.
   */
  active: { row: number; col: number };
  /** Selection with `r1 <= r2` and `c1 <= c2`. */
  bounds: Range;
  extraRanges: Range[];
  select(range: Range): void;
  selectCell(row: number, col: number): void;
  /** Extends from the existing anchor, as Shift+click and Shift+arrow do. */
  extendTo(row: number, col: number): void;
  addRange(range: Range): void;
  clearExtraRanges(): void;
  /** Moves the active cell, clamped to the sheet extent. */
  move(dRow: number, dCol: number, numRows: number, numCols: number): void;
}

const clamp = (v: number, lo: number, hi: number) => Math.max(lo, Math.min(hi, v));

export function useSelection(initial?: Range): UseSelectionResult {
  const [selection, setSelection] = useState<Range>(
    initial ?? { r1: 0, c1: 0, r2: 0, c2: 0 },
  );
  const [extraRanges, setExtraRanges] = useState<Range[]>([]);

  const selectCell = useCallback((row: number, col: number) => {
    setSelection({ r1: row, c1: col, r2: row, c2: col });
    setExtraRanges([]);
  }, []);

  const extendTo = useCallback((row: number, col: number) => {
    setSelection((s) => ({ ...s, r2: row, c2: col }));
  }, []);

  const move = useCallback(
    (dRow: number, dCol: number, numRows: number, numCols: number) => {
      setSelection((s) => {
        // From the anchor, not the drag end: after selecting D5:F13, an arrow key
        // collapses to a cell next to D5 — the active cell you can see.
        const row = clamp(s.r1 + dRow, 0, numRows - 1);
        const col = clamp(s.c1 + dCol, 0, numCols - 1);
        return { r1: row, c1: col, r2: row, c2: col };
      });
      setExtraRanges([]);
    },
    [],
  );

  return {
    selection,
    active: { row: selection.r1, col: selection.c1 },
    bounds: normalizeRange(selection),
    extraRanges,
    select: setSelection,
    selectCell,
    extendTo,
    addRange: useCallback((r: Range) => setExtraRanges((prev) => [...prev, r]), []),
    clearExtraRanges: useCallback(() => setExtraRanges([]), []),
    move,
  };
}
