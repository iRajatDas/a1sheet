/**
 * Copy and paste — consolidated.
 *
 * The POC implemented this TWICE: once on the hidden textarea's onCopy/onPaste
 * events, and again via navigator.clipboard for the context menu. Both encoded
 * the internal-vs-external paste decision and could drift. Here it lives once,
 * with the two transports as thin callers of `copy` and `paste`.
 *
 * Internal-paste detection (preserved from the POC, including its failure mode):
 * `lastCopied` stores the copied grid, its origin, and the exact serialized text.
 * On paste the incoming clipboard text is compared against that stored text — a
 * match means internal, so relative refs get shifted; anything else is treated as
 * external and pastes values only.
 *
 * Known edge case: copying identical text from another application between an
 * internal copy and paste is misread as internal. Accepted, as in the POC.
 *
 * Also preserved: paste aligns from the target's top-left corner with NO shape
 * validation against the source.
 */
import { useCallback, useRef, useState } from "react";
import { shiftFormulaRefs } from "../formula/refs.js";
import { cellKey, normalizeRange } from "../model/address.js";
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
  /** Serializes the range as TSV and remembers it for internal-paste detection. */
  copy(sheet: Sheet, selection: Range): string;
  /** Applies clipboard text at `target`. Returns the range it wrote. */
  paste(
    text: string,
    target: { row: number; col: number },
    updateSheet: (fn: SheetUpdater, addHistory?: boolean) => void,
  ): Range;
  lastCopied(): CopiedGrid | null;
  /**
   * The range the last copy came from, for the dashed outline the grid draws
   * around it. Null once the copy has been used or dismissed.
   *
   * State rather than a ref, unlike `lastCopied`: this one is rendered, so a
   * copy has to re-render the grid. What it marks is the SOURCE, which is why it
   * survives the selection moving away to wherever the paste is going.
   */
  copiedRange: Range | null;
  /** Clears the outline — on paste, on Escape, or when an edit invalidates it. */
  clearCopied(): void;
}

/** TSV: tab between columns, newline between rows — what Excel and Sheets use. */
function serialize(grid: string[][]): string {
  return grid.map((row) => row.join("\t")).join("\n");
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
  const [copiedRange, setCopiedRange] = useState<Range | null>(null);

  const copy = useCallback((sheet: Sheet, selection: Range) => {
    const b = normalizeRange(selection);
    const grid: string[][] = [];
    for (let r = b.r1; r <= b.r2; r++) {
      const row: string[] = [];
      for (let c = b.c1; c <= b.c2; c++) row.push(sheet.cells[cellKey(r, c)] ?? "");
      grid.push(row);
    }
    const text = serialize(grid);
    last.current = { grid, origin: { row: b.r1, col: b.c1 }, text };
    setCopiedRange(b);
    return text;
  }, []);

  const paste = useCallback(
    (
      text: string,
      target: { row: number; col: number },
      updateSheet: (fn: SheetUpdater, addHistory?: boolean) => void,
    ) => {
      setCopiedRange(null);
      const internal = last.current && last.current.text === text;
      const grid = internal ? (last.current as CopiedGrid).grid : deserialize(text);
      const origin = internal ? (last.current as CopiedGrid).origin : target;
      const dRow = target.row - origin.row;
      const dCol = target.col - origin.col;

      updateSheet((s) => {
        grid.forEach((row, ri) => {
          row.forEach((raw, ci) => {
            const r = target.row + ri;
            const c = target.col + ci;
            if (r >= s.numRows || c >= s.numCols) return;
            const key = cellKey(r, c);
            if (s.styles[key]?.locked) return;
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
    copiedRange,
    clearCopied: useCallback(() => setCopiedRange(null), []),
  };
}
