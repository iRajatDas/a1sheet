/**
 * Fill handle drag. Ported from startFillDrag/commitFillDrag in
 * ref/Spreadsheet.jsx.
 *
 * Mechanics preserved from the POC:
 * - On drag start, snapshot the selection bounds.
 * - The Grid mounts a transparent overlay across the whole scroll container and
 *   computes the hovered row/col from raw mouse coordinates on mousemove.
 *   Per-cell onMouseEnter was tried in the POC and is unreliable during fast drags.
 * - Commit extends DOWN or RIGHT only. Dragging up or left is a no-op by design.
 * - A non-formula source goes through `extrapolateSeries`; a formula source goes
 *   through `shiftFormulaRefs` once per destination cell.
 */
import { useCallback, useState } from "react";
import { shiftFormulaRefs } from "../formula/refs.js";
import { extrapolateSeries } from "../formula/series.js";
import { cellKey, normalizeRange } from "../model/address.js";
import type { Range } from "../model/types.js";
import type { SheetUpdater } from "./useWorkbook.js";

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
}

export function useFillHandle(): UseFillHandleResult {
  const [source, setSource] = useState<Range | null>(null);
  const [target, setTarget] = useState<{ row: number; col: number } | null>(null);

  const preview: Range | null = (() => {
    if (!source) return null;
    const b = normalizeRange(source);
    if (!target) return b;
    return {
      r1: b.r1,
      c1: b.c1,
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
      const downTo = target.row > b.r2 ? target.row : b.r2;
      const rightTo = target.col > b.c2 ? target.col : b.c2;
      if (downTo === b.r2 && rightTo === b.c2) {
        cancel();
        return null;
      }

      updateSheet((s) => {
        // Vertical fill: one series per column of the source.
        if (downTo > b.r2) {
          const count = downTo - b.r2;
          for (let c = b.c1; c <= b.c2; c++) {
            const srcRows: string[] = [];
            for (let r = b.r1; r <= b.r2; r++) {
              srcRows.push(s.cells[cellKey(r, c)] ?? "");
            }
            const lastIsFormula = srcRows[srcRows.length - 1]?.startsWith("=");
            if (lastIsFormula) {
              const height = b.r2 - b.r1 + 1;
              for (let i = 0; i < count; i++) {
                const srcRow = b.r1 + (i % height);
                const raw = s.cells[cellKey(srcRow, c)];
                if (!raw) continue;
                const destRow = b.r2 + 1 + i;
                s.cells[cellKey(destRow, c)] = raw.startsWith("=")
                  ? `=${shiftFormulaRefs(raw.slice(1), destRow - srcRow, 0)}`
                  : raw;
              }
            } else {
              const filled = extrapolateSeries(srcRows, count);
              filled.forEach((v, i) => {
                s.cells[cellKey(b.r2 + 1 + i, c)] = v;
              });
            }
          }
        }

        // Horizontal fill: one series per row of the source.
        if (rightTo > b.c2) {
          const count = rightTo - b.c2;
          for (let r = b.r1; r <= b.r2; r++) {
            const srcCols: string[] = [];
            for (let c = b.c1; c <= b.c2; c++) {
              srcCols.push(s.cells[cellKey(r, c)] ?? "");
            }
            const lastIsFormula = srcCols[srcCols.length - 1]?.startsWith("=");
            if (lastIsFormula) {
              const width = b.c2 - b.c1 + 1;
              for (let i = 0; i < count; i++) {
                const srcCol = b.c1 + (i % width);
                const raw = s.cells[cellKey(r, srcCol)];
                if (!raw) continue;
                const destCol = b.c2 + 1 + i;
                s.cells[cellKey(r, destCol)] = raw.startsWith("=")
                  ? `=${shiftFormulaRefs(raw.slice(1), 0, destCol - srcCol)}`
                  : raw;
              }
            } else {
              const filled = extrapolateSeries(srcCols, count);
              filled.forEach((v, i) => {
                s.cells[cellKey(r, b.c2 + 1 + i)] = v;
              });
            }
          }
        }

        return s;
      });

      const result = { r1: b.r1, c1: b.c1, r2: downTo, c2: rightTo };
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
  };
}
