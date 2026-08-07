/**
 * Fill handle drag. Ported from startFillDrag/commitFillDrag in
 * ref/Spreadsheet.jsx.
 *
 * Mechanics preserved from the POC:
 * - On drag start, snapshot the selection bounds.
 * - The Grid tracks the pointer on `window` and computes the hovered row/col
 *   from raw client coordinates. Per-cell onMouseEnter was tried in the POC and
 *   is unreliable during fast drags — and blind to a pointer held outside the
 *   grid, which is where it sits while the sheet auto-scrolls.
 * - Commit extends in whichever direction the drag went. A series extrapolates
 *   the way it is read, so filling upward or leftward reads the source in
 *   reverse — otherwise 1, 2, 3 dragged upward continues 4, 5, 6.
 * - A non-formula source goes through `extrapolateSeries`; a formula source goes
 *   through `shiftFormulaRefs` once per destination cell.
 */
import { useCallback, useState } from "react";
import { shiftFormulaRefs } from "../formula/refs.js";
import { extrapolateSeries } from "../formula/series.js";
import { cellKey, normalizeRange } from "../model/address.js";
import type { Range, Sheet } from "../model/types.js";
import type { SheetUpdater } from "./useWorkbook.js";

/** Which way a keyboard fill runs. The handle takes a target cell instead. */
export type FillDirection = "down" | "right";

export interface UseFillHandleResult {
  /** True while a drag is in progress — the Grid should mount its overlay. */
  dragging: boolean;
  /** Range that would be filled if the drag ended now, for the preview outline. */
  preview: Range | null;
  start(selection: Range): void;
  moveTo(row: number, col: number): void;
  /** Applies the fill and ends the drag. No-op when not dragging. */
  commit(
    updateSheet: (fn: SheetUpdater, addHistory?: boolean) => void,
  ): Range | null;
  cancel(): void;
  /**
   * The keyboard's fill: takes the leading filled lines of `range` as the source
   * and writes them over the rest of it. Ctrl+D and Ctrl+R.
   *
   * Returns false when the selection gives it nothing to do — one line only, or
   * no filled line to start from — so the caller can say why nothing happened.
   */
  within(
    range: Range,
    direction: FillDirection,
    updateSheet: (fn: SheetUpdater, addHistory?: boolean) => void,
  ): boolean;
}

/**
 * Writes the fill of `b` out to the cell the drag ended on, in place on a draft
 * sheet.
 *
 * Module-level and taking the sheet as an argument so the keyboard commands can
 * run exactly the same fill as the handle. A second implementation of series
 * extrapolation and reference shifting is how the two would drift.
 */
function fillFrom(s: Sheet, b: Range, target: { row: number; col: number }): void {
  const up = Math.min(b.r1, target.row);
  const down = Math.max(b.r2, target.row);
  const left = Math.min(b.c1, target.col);
  const right = Math.max(b.c2, target.col);

  /**
   * Fills `count` cells outward from one line of the source.
   *
   * The four directions differ only in the order the source is read and where
   * the results go, which is why they are one function rather than four blocks.
   * A series extrapolates the way it is read, so filling backwards reads the
   * source backwards — otherwise 1, 2, 3 dragged upward continues 4, 5, 6
   * instead of 0, -1, -2.
   */
  const column = (c: number, top: number, bottom: number, up: boolean) =>
    Array.from({ length: bottom - top + 1 }, (_, i) =>
      up
        ? (s.cells[cellKey(bottom - i, c)] ?? "")
        : (s.cells[cellKey(top + i, c)] ?? ""),
    );
  const rowOf = (r: number, from: number, to: number, back: boolean) =>
    Array.from({ length: to - from + 1 }, (_, i) =>
      back
        ? (s.cells[cellKey(r, to - i)] ?? "")
        : (s.cells[cellKey(r, from + i)] ?? ""),
    );
  const fillLine = (
    count: number,
    /** The source cells, in the order the fill will consume them. */
    values: readonly string[],
    write: (offset: number, value: string) => void,
    /** Row and column deltas for shifting a formula's references. */
    shift: (offset: number, sourceIndex: number) => [number, number],
  ) => {
    if (count <= 0 || values.length === 0) return;

    const last = values[values.length - 1];
    if (last?.startsWith("=")) {
      for (let i = 0; i < count; i++) {
        const sourceIndex = i % values.length;
        const raw = values[sourceIndex];
        if (!raw) continue;
        const [dRow, dCol] = shift(i, sourceIndex);
        write(
          i,
          raw.startsWith("=")
            ? `=${shiftFormulaRefs(raw.slice(1), dRow, dCol)}`
            : raw,
        );
      }
      return;
    }

    for (const [i, value] of extrapolateSeries([...values], count).entries()) {
      write(i, value);
    }
  };

  for (let c = b.c1; c <= b.c2; c++) {
    if (down > b.r2) {
      fillLine(
        down - b.r2,
        column(c, b.r1, b.r2, false),
        (offset, value) => {
          s.cells[cellKey(b.r2 + 1 + offset, c)] = value;
        },
        (offset, sourceIndex) => [b.r2 + 1 + offset - (b.r1 + sourceIndex), 0],
      );
    }
    if (up < b.r1) {
      fillLine(
        b.r1 - up,
        column(c, b.r1, b.r2, true),
        (offset, value) => {
          s.cells[cellKey(b.r1 - 1 - offset, c)] = value;
        },
        (offset, sourceIndex) => [b.r1 - 1 - offset - (b.r2 - sourceIndex), 0],
      );
    }
  }

  for (let r = b.r1; r <= b.r2; r++) {
    if (right > b.c2) {
      fillLine(
        right - b.c2,
        rowOf(r, b.c1, b.c2, false),
        (offset, value) => {
          s.cells[cellKey(r, b.c2 + 1 + offset)] = value;
        },
        (offset, sourceIndex) => [0, b.c2 + 1 + offset - (b.c1 + sourceIndex)],
      );
    }
    if (left < b.c1) {
      fillLine(
        b.c1 - left,
        rowOf(r, b.c1, b.c2, true),
        (offset, value) => {
          s.cells[cellKey(r, b.c1 - 1 - offset)] = value;
        },
        (offset, sourceIndex) => [0, b.c1 - 1 - offset - (b.c2 - sourceIndex)],
      );
    }
  }
}

/**
 * How much of a selection is the source of a keyboard fill.
 *
 * The leading run of lines that hold anything, which is the rule that makes one
 * shortcut cover both jobs: A1 filled and A1:A20 selected copies A1 down, and
 * A1:A2 holding 1 and 2 with A1:A20 selected counts on to 20. Excel's Ctrl+D
 * only ever copies its top row, so a numbered column needs the mouse there.
 */
function sourceLines(s: Sheet, b: Range, direction: FillDirection): number {
  const has = (line: number) => {
    if (direction === "down") {
      for (let c = b.c1; c <= b.c2; c++) {
        if ((s.cells[cellKey(line, c)] ?? "") !== "") return true;
      }
      return false;
    }
    for (let r = b.r1; r <= b.r2; r++) {
      if ((s.cells[cellKey(r, line)] ?? "") !== "") return true;
    }
    return false;
  };

  const first = direction === "down" ? b.r1 : b.c1;
  const last = direction === "down" ? b.r2 : b.c2;
  let line = first;
  while (line <= last && has(line)) line++;
  // An empty selection has no source at all; the caller reports that rather
  // than silently writing blanks over the range.
  return line - first;
}

export function useFillHandle(): UseFillHandleResult {
  const [source, setSource] = useState<Range | null>(null);
  const [target, setTarget] = useState<{ row: number; col: number } | null>(null);

  const preview: Range | null = (() => {
    if (!source) return null;
    const b = normalizeRange(source);
    if (!target) return b;
    // The preview grows in whichever direction the pointer went, so dragging up
    // or left previews the cells that would be filled rather than nothing.
    return {
      r1: Math.min(b.r1, target.row),
      c1: Math.min(b.c1, target.col),
      r2: Math.max(b.r2, target.row),
      c2: Math.max(b.c2, target.col),
    };
  })();

  const cancel = useCallback(() => {
    setSource(null);
    setTarget(null);
  }, []);

  const commit = useCallback(
    (
      updateSheet: (fn: SheetUpdater, addHistory?: boolean) => void,
    ): Range | null => {
      if (!source || !target) {
        cancel();
        return null;
      }
      const b = normalizeRange(source);
      const up = Math.min(b.r1, target.row);
      const down = Math.max(b.r2, target.row);
      const left = Math.min(b.c1, target.col);
      const right = Math.max(b.c2, target.col);
      if (up === b.r1 && down === b.r2 && left === b.c1 && right === b.c2) {
        cancel();
        return null;
      }

      updateSheet((s) => {
        fillFrom(s, b, target);
        return s;
      });

      const result = { r1: up, c1: left, r2: down, c2: right };
      cancel();
      return result;
    },
    [source, target, cancel],
  );

  return {
    dragging: source !== null,
    preview,
    start: useCallback((selection: Range) => {
      setSource(normalizeRange(selection));
      setTarget(null);
    }, []),
    moveTo: useCallback((row: number, col: number) => setTarget({ row, col }), []),
    commit,
    cancel,
    within: useCallback((range, direction, updateSheet) => {
      const b = normalizeRange(range);
      const span = direction === "down" ? b.r2 - b.r1 : b.c2 - b.c1;
      if (span < 1) return false;

      // The source run can only be measured against the cells, which live
      // inside the update. `applied` carries the verdict back out.
      let applied = false;
      updateSheet((s) => {
        const lines = sourceLines(s, b, direction);
        if (lines === 0 || lines > span) return s;
        applied = true;
        const source =
          direction === "down"
            ? { ...b, r2: b.r1 + lines - 1 }
            : { ...b, c2: b.c1 + lines - 1 };
        fillFrom(s, source, { row: b.r2, col: b.c2 });
        return s;
      });
      return applied;
    }, []),
  };
}
