/**
 * Copy and paste — consolidated.
 *
 * Two transports reach it: the hidden textarea's onCopy/onPaste events, and
 * navigator.clipboard for the context menu. Both need the same
 * internal-vs-external paste decision, so it lives here once and they are thin
 * callers of `copy` and `paste`.
 *
 * Internal-paste detection:
 * `lastCopied` stores the copied grid, its origin, and the exact serialized text.
 * On paste the incoming clipboard text is compared against that stored text — a
 * match means internal, so relative refs get shifted; anything else is treated as
 * external and pastes values only.
 *
 * Known edge case: copying identical text from another application between an
 * internal copy and paste is misread as internal. Accepted.
 *
 * Paste aligns from the target's top-left corner with NO shape validation
 * against the source.
 */
import { useCallback, useRef, useState } from "react";
import { shiftFormulaRefs } from "../formula/refs.js";
import { cellKey, normalizeRange } from "../model/address.js";
import { rejectCellValue } from "../model/cellValidation.js";
import type { Evaluator } from "../formula/evaluate.js";
import type { Range, Sheet } from "../model/types.js";
import type { SheetUpdater } from "./useWorkbook.js";

export interface CopiedGrid {
  /** Raw cell contents, row-major, from the copied range. */
  grid: string[][];
  origin: { row: number; col: number };
  /** Exact serialized text placed on the clipboard. */
  text: string;
}

export interface UseClipboardResult {
  /**
   * Serializes the selection as TSV and remembers it for internal-paste
   * detection.
   *
   * Takes every selected range. Excel's rule for a multi-range copy is that the
   * ranges have to line up into one block — same columns stacked vertically, or
   * same rows side by side — and it refuses anything else rather than guessing
   * at the shape. Null is that refusal; the caller reports it.
   */
  copy(sheet: Sheet, ranges: readonly Range[]): string | null;
  /** Applies clipboard text at `target`. Returns the range it wrote. */
  paste(
    text: string,
    target: { row: number; col: number },
    updateSheet: (fn: SheetUpdater, addHistory?: boolean) => void,
    options?: {
      evaluator?: Evaluator;
      onReject?: (message: string) => void;
      selection?: Range;
    },
  ): Range;
  lastCopied(): CopiedGrid | null;
  /**
   * The ranges the last copy came from, for the dashed outline the grid draws
   * around them. Empty once the copy has been used or dismissed.
   *
   * State rather than a ref, unlike `lastCopied`: this one is rendered, so a
   * copy has to re-render the grid. What it marks is the SOURCE, which is why it
   * survives the selection moving away to wherever the paste is going.
   */
  copiedRanges: readonly Range[];
  /** Clears the outline — on paste, on Escape, or when an edit invalidates it. */
  clearCopied(): void;
}

/** TSV: tab between columns, newline between rows — what Excel and Sheets use. */
function serialize(grid: string[][]): string {
  return grid.map((row) => row.join("\t")).join("\n");
}

/**
 * Lays several selected blocks out as the one grid a clipboard can hold.
 *
 * Two shapes work and no others: blocks in the same columns, stacked in row
 * order, and blocks in the same rows, placed side by side in column order.
 * Anything else — an L, two overlapping rectangles, blocks of differing widths —
 * has no single sensible flattening, and inventing one would silently paste a
 * shape the user never selected. Excel refuses these too.
 *
 * Returns null on refusal, and the origin of the topmost-leftmost block on
 * success, which is what relative references shift against.
 */
function joinBlocks(
  blocks: readonly Range[],
  read: (block: Range) => string[][],
): { grid: string[][]; origin: { row: number; col: number } } | null {
  const first = blocks[0];
  if (!first) return null;
  if (blocks.length === 1) {
    return { grid: read(first), origin: { row: first.r1, col: first.c1 } };
  }

  const sameCols = blocks.every((b) => b.c1 === first.c1 && b.c2 === first.c2);
  if (sameCols) {
    const ordered = [...blocks].sort((a, b) => a.r1 - b.r1);
    // Overlapping blocks would repeat their shared rows.
    for (let i = 1; i < ordered.length; i++) {
      if ((ordered[i] as Range).r1 <= (ordered[i - 1] as Range).r2) return null;
    }
    const top = ordered[0] as Range;
    return {
      grid: ordered.flatMap(read),
      origin: { row: top.r1, col: top.c1 },
    };
  }

  const sameRows = blocks.every((b) => b.r1 === first.r1 && b.r2 === first.r2);
  if (sameRows) {
    const ordered = [...blocks].sort((a, b) => a.c1 - b.c1);
    for (let i = 1; i < ordered.length; i++) {
      if ((ordered[i] as Range).c1 <= (ordered[i - 1] as Range).c2) return null;
    }
    const left = ordered[0] as Range;
    const grids = ordered.map(read);
    const height = first.r2 - first.r1 + 1;
    const grid = Array.from({ length: height }, (_, r) =>
      grids.flatMap((g) => g[r] ?? []),
    );
    return { grid, origin: { row: left.r1, col: left.c1 } };
  }

  return null;
}

function deserialize(text: string): string[][] {
  return text
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .split("\n")
    .map((line) => line.split("\t"));
}

export function useClipboard(): UseClipboardResult {
  const last = useRef<CopiedGrid | null>(null);
  const [copiedRanges, setCopiedRanges] = useState<readonly Range[]>([]);

  const copy = useCallback((sheet: Sheet, ranges: readonly Range[]) => {
    const blocks = ranges.map(normalizeRange);
    const first = blocks[0];
    if (!first) return null;

    const read = (b: Range) => {
      const grid: string[][] = [];
      for (let r = b.r1; r <= b.r2; r++) {
        const row: string[] = [];
        for (let c = b.c1; c <= b.c2; c++) {
          row.push(sheet.cells[cellKey(r, c)] ?? "");
        }
        grid.push(row);
      }
      return grid;
    };

    const joined = joinBlocks(blocks, read);
    if (!joined) return null;

    const text = serialize(joined.grid);
    last.current = { grid: joined.grid, origin: joined.origin, text };
    setCopiedRanges(blocks);
    return text;
  }, []);

  const paste = useCallback(
    (
      text: string,
      target: { row: number; col: number },
      updateSheet: (fn: SheetUpdater, addHistory?: boolean) => void,
      options?: {
        evaluator?: Evaluator;
        onReject?: (message: string) => void;
        selection?: Range;
      },
    ) => {
      setCopiedRanges([]);
      const internal = last.current && last.current.text === text;
      const grid = internal ? (last.current as CopiedGrid).grid : deserialize(text);
      const origin = internal ? (last.current as CopiedGrid).origin : target;
      const dRow = target.row - origin.row;
      const dCol = target.col - origin.col;

      if (options?.selection) {
        const sel = normalizeRange(options.selection);
        const selRows = sel.r2 - sel.r1 + 1;
        const selCols = sel.c2 - sel.c1 + 1;
        const pasteRows = grid.length;
        const pasteCols = Math.max(...grid.map((r) => r.length), 1);
        if (
          (selRows > 1 || selCols > 1) &&
          (selRows !== pasteRows || selCols !== pasteCols)
        ) {
          options.onReject?.(
            "Paste shape does not match the selection — only the top-left cell is used.",
          );
        }
      }

      updateSheet((s) => {
        grid.forEach((row, ri) => {
          row.forEach((raw, ci) => {
            const r = target.row + ri;
            const c = target.col + ci;
            if (r >= s.numRows || c >= s.numCols) return;
            const key = cellKey(r, c);
            if (s.styles[key]?.locked) return;
            if (options?.evaluator) {
              const rejection = rejectCellValue(s, r, c, raw, options.evaluator);
              if (rejection) {
                options.onReject?.(rejection.message);
                return;
              }
            }
            if (raw === "") {
              delete s.cells[key];
              return;
            }
            // Only an internal paste shifts formula references.
            s.cells[key] =
              internal && raw.startsWith("=")
                ? `=${shiftFormulaRefs(raw.slice(1), dRow, dCol)}`
                : raw;
          });
        });
        return s;
      });

      const rows = grid.length;
      const cols = Math.max(...grid.map((r) => r.length), 1);
      return {
        r1: target.row,
        c1: target.col,
        r2: target.row + rows - 1,
        c2: target.col + cols - 1,
      };
    },
    [],
  );

  return {
    copy,
    paste,
    lastCopied: useCallback(() => last.current, []),
    copiedRanges,
    clearCopied: useCallback(() => setCopiedRanges([]), []),
  };
}
