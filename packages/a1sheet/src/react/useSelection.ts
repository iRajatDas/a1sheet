/**
 * Selection state. Component-local — selection is never part of the workbook, so
 * undo does not restore it.
 *
 * `extraRanges` holds Ctrl+click history. It feeds the status bar ONLY: copy,
 * fill, and paste all act on `selection`. That is a deliberate scope cut.
 */
import { useCallback, useRef, useState } from "react";
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
  /**
   * Freezes the current selection into `extraRanges` and makes `range` the new
   * active one — what Ctrl+click does, on a cell or on a whole row or column,
   * and what add-mode does on the first arrow key.
   *
   * `addRange` followed by `selectCell` looks equivalent and is not: `selectCell`
   * clears the extras, so the pair discards every range it was meant to keep.
   */
  startNewRange(range: Range): void;
  /**
   * Drops the range containing a cell — Ctrl+clicking a second time on
   * something already selected, which is how a misclick gets taken back.
   *
   * Returns false when the cell is in no range, so the caller can fall through
   * to adding one. Excel removes the CELL and splits the range around it; this
   * removes the whole range, which is predictable in a way that a range silently
   * becoming three is not.
   */
  removeRangeAt(row: number, col: number): boolean;
  /**
   * Excel's Shift+F8 "Add to Selection": while on, moving the cursor keeps the
   * range behind it instead of replacing it, which is the only way to build a
   * discontiguous selection without a mouse.
   */
  addMode: boolean;
  toggleAddMode(): void;
  /** Moves the active cell, clamped to the sheet extent. */
  move(dRow: number, dCol: number, numRows: number, numCols: number): void;
  /** Moves the active cell to an absolute address, clamped the same way. */
  moveTo(row: number, col: number, numRows: number, numCols: number): void;
}

const clamp = (v: number, lo: number, hi: number) => Math.max(lo, Math.min(hi, v));

export function useSelection(initial?: Range): UseSelectionResult {
  const [selection, setSelection] = useState<Range>(
    initial ?? { r1: 0, c1: 0, r2: 0, c2: 0 },
  );
  const [extraRanges, setExtraRanges] = useState<Range[]>([]);
  const [addMode, setAddMode] = useState(false);
  /**
   * True while add-mode is on and the range under the cursor has not been
   * frozen yet. Without it every arrow key in add-mode would deposit another
   * one-cell range, and walking to the next block would leave a trail.
   */
  const unfrozen = useRef(false);

  const selectCell = useCallback((row: number, col: number) => {
    setSelection({ r1: row, c1: col, r2: row, c2: col });
    setExtraRanges([]);
    setAddMode(false);
  }, []);

  const extendTo = useCallback((row: number, col: number) => {
    setSelection((s) => ({ ...s, r2: row, c2: col }));
  }, []);

  const startNewRange = useCallback(
    (range: Range) => {
      setExtraRanges((prev) => [...prev, selection]);
      setSelection(range);
    },
    [selection],
  );

  const removeRangeAt = useCallback(
    (row: number, col: number) => {
      const covers = (r: Range) => {
        const n = normalizeRange(r);
        return row >= n.r1 && row <= n.r2 && col >= n.c1 && col <= n.c2;
      };

      const index = extraRanges.findIndex(covers);
      if (index >= 0) {
        setExtraRanges((prev) => prev.filter((_, i) => i !== index));
        return true;
      }

      // The primary range can only go if another one can take its place: there
      // is always an active cell, and "no selection at all" is not a state the
      // rest of the code is written for.
      const last = extraRanges[extraRanges.length - 1];
      if (last && covers(selection)) {
        setExtraRanges((prev) => prev.slice(0, -1));
        setSelection(last);
        return true;
      }
      return false;
    },
    [extraRanges, selection],
  );

  /**
   * Collapses the selection onto one cell.
   *
   * Reads `selection` from the render rather than through an updater because it
   * has to touch `extraRanges` as well, and a set-inside-an-updater fires twice
   * under StrictMode — which would deposit every add-mode range twice.
   */
  const moveTo = useCallback(
    (row: number, col: number, numRows: number, numCols: number) => {
      const r = clamp(row, 0, numRows - 1);
      const c = clamp(col, 0, numCols - 1);
      if (addMode) {
        if (unfrozen.current) setExtraRanges((prev) => [...prev, selection]);
        unfrozen.current = false;
      } else {
        setExtraRanges([]);
      }
      setSelection({ r1: r, c1: c, r2: r, c2: c });
    },
    [addMode, selection],
  );

  const move = useCallback(
    (dRow: number, dCol: number, numRows: number, numCols: number) => {
      // From the anchor, not the drag end: after selecting D5:F13, an arrow key
      // collapses to a cell next to D5 — the active cell you can see.
      moveTo(selection.r1 + dRow, selection.c1 + dCol, numRows, numCols);
    },
    [moveTo, selection.r1, selection.c1],
  );

  const toggleAddMode = useCallback(() => {
    setAddMode((on) => {
      unfrozen.current = !on;
      return !on;
    });
  }, []);

  return {
    selection,
    active: { row: selection.r1, col: selection.c1 },
    bounds: normalizeRange(selection),
    extraRanges,
    select: setSelection,
    selectCell,
    extendTo,
    addRange: useCallback((r: Range) => setExtraRanges((prev) => [...prev, r]), []),
    clearExtraRanges: useCallback(() => {
      setExtraRanges([]);
      setAddMode(false);
    }, []),
    startNewRange,
    removeRangeAt,
    addMode,
    toggleAddMode,
    move,
    moveTo,
  };
}
