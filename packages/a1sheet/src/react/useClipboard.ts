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
 * On paste the incoming clipboard text is compared against that stored text
 * after normalizing line endings and trailing newlines (browsers often append a
 * trailing `\n` that Excel-style paste still treats as the same copy). A match
 * means internal, so relative refs get shifted; anything else is external —
 * formulas paste as written with no shift (same as pasting `=A1` from a text
 * editor into Excel).
 *
 * Known edge case: copying identical text from another application between an
 * internal copy and paste is misread as internal. Accepted.
 */
import { useCallback, useRef, useState } from "react";
import type { Evaluator } from "../formula/evaluate.js";
import { shiftFormulaRefs } from "../formula/refs.js";
import { checkPasteMerge } from "../model/mergeGuards.js";
import { cellKey, normalizeRange, toA1 } from "../model/address.js";
import { rejectCellValue } from "../model/cellValidation.js";
import type { Range, Sheet, StyleObject } from "../model/types.js";
import type { SheetPatcher, SheetUpdater } from "./useWorkbook.js";

/**
 * What a paste writes. `"all"` is the default keyboard path — values with
 * internal formula shifting. Modes map to common paste-special sets.
 */
export type PasteMode =
  | "all"
  | "values"
  | "formats"
  | "formulas"
  | "transpose"
  | "text";

export interface CopiedGrid {
  /** Raw cell contents, row-major, from the copied range. */
  grid: string[][];
  /** Parallel style grid; undefined cells mean no style was copied. */
  styles: (StyleObject | undefined)[][];
  origin: { row: number; col: number };
  /** Exact serialized text placed on the clipboard. */
  text: string;
}

export interface PasteOptions {
  mode?: PasteMode;
  evaluator?: Evaluator;
  onReject?: (message: string) => void;
  selection?: Range;
  /**
   * Prefer this over `updateSheet` for large pastes — shallow-copies only the
   * maps that change. Falls back to cloning via the third paste argument when
   * omitted.
   */
  patchSheet?: (fn: SheetPatcher, addHistory?: boolean) => void;
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
    options?: PasteOptions,
  ): Range;
  lastCopied(): CopiedGrid | null;
  copiedRanges: readonly Range[];
  clearCopied(): void;
}

/** TSV: tab between columns, newline between rows — what Excel uses. */
function serialize(grid: string[][]): string {
  return grid.map((row) => row.join("\t")).join("\n");
}

/**
 * Browsers and OS clipboards disagree on trailing newlines and `\r\n`.
 * Normalize before deciding whether a paste is our own copy.
 */
export function normalizeClipboardText(text: string): string {
  return text.replace(/\r\n/g, "\n").replace(/\r/g, "\n").replace(/\n+$/u, "");
}

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

function transposeGrid<T>(grid: T[][]): T[][] {
  const rows = grid.length;
  const cols = Math.max(...grid.map((r) => r.length), 0);
  const out: T[][] = [];
  for (let c = 0; c < cols; c++) {
    const row: T[] = [];
    for (let r = 0; r < rows; r++) {
      row.push(grid[r]![c] as T);
    }
    out.push(row);
  }
  return out;
}

function emptyStyles(grid: string[][]): (StyleObject | undefined)[][] {
  return grid.map((row) => row.map(() => undefined));
}

function statusFor(mode: PasteMode, range: Range): string {
  const label =
    range.r1 === range.r2 && range.c1 === range.c2
      ? toA1(range.r1, range.c1)
      : `${toA1(range.r1, range.c1)}:${toA1(range.r2, range.c2)}`;
  switch (mode) {
    case "values":
      return `Pasted values into ${label}`;
    case "formats":
      return `Pasted formats into ${label}`;
    case "formulas":
      return `Pasted formulas into ${label}`;
    case "transpose":
      return `Pasted transposed data into ${label}`;
    case "text":
      return `Pasted as text into ${label}`;
    default:
      return `Pasted into ${label}`;
  }
}

function resolveValue(options: {
  raw: string;
  mode: PasteMode;
  internal: boolean;
  dRow: number;
  dCol: number;
  evaluator?: Evaluator;
  originRow: number;
  originCol: number;
  ri: number;
  ci: number;
}): string {
  const { raw, mode, internal, dRow, dCol, evaluator, originRow, originCol, ri, ci } =
    options;

  if (mode === "text") {
    return raw.startsWith("=") ? `'${raw}` : raw.startsWith("'") ? raw : raw;
  }

  if (mode === "values") {
    if (internal && evaluator) {
      const shown = evaluator.getCellDisplay(originRow + ri, originCol + ci);
      if (typeof shown === "boolean") return shown ? "TRUE" : "FALSE";
      if (shown === undefined || shown === "") return "";
      return String(shown);
    }
    return raw;
  }

  if (mode === "formulas" || mode === "all" || mode === "transpose") {
    if (internal && raw.startsWith("=")) {
      return `=${shiftFormulaRefs(raw.slice(1), dRow, dCol)}`;
    }
    return raw;
  }

  return raw;
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

    const readStyles = (b: Range) => {
      const grid: (StyleObject | undefined)[][] = [];
      for (let r = b.r1; r <= b.r2; r++) {
        const row: (StyleObject | undefined)[] = [];
        for (let c = b.c1; c <= b.c2; c++) {
          row.push(sheet.styles[cellKey(r, c)]);
        }
        grid.push(row);
      }
      return grid;
    };

    const joined = joinBlocks(blocks, read);
    if (!joined) return null;

    const styleJoined = joinBlocks(blocks, (b) =>
      readStyles(b).map((row) =>
        row.map((s) => (s ? JSON.stringify(s) : "")),
      ),
    );
    const styles: (StyleObject | undefined)[][] = styleJoined
      ? styleJoined.grid.map((row) =>
          row.map((raw) => {
            if (!raw) return undefined;
            try {
              return JSON.parse(raw) as StyleObject;
            } catch {
              return undefined;
            }
          }),
        )
      : emptyStyles(joined.grid);

    const text = serialize(joined.grid);
    last.current = {
      grid: joined.grid,
      styles,
      origin: joined.origin,
      text,
    };
    setCopiedRanges(blocks);
    return text;
  }, []);

  const paste = useCallback(
    (
      text: string,
      target: { row: number; col: number },
      updateSheet: (fn: SheetUpdater, addHistory?: boolean) => void,
      options?: PasteOptions,
    ) => {
      setCopiedRanges([]);
      const mode: PasteMode = options?.mode ?? "all";
      const internal = !!(
        last.current &&
        normalizeClipboardText(last.current.text) ===
          normalizeClipboardText(text)
      );
      let grid = internal ? (last.current as CopiedGrid).grid : deserialize(text);
      let styles = internal
        ? (last.current as CopiedGrid).styles
        : emptyStyles(grid);
      const origin = internal ? (last.current as CopiedGrid).origin : target;

      if (mode === "transpose") {
        grid = transposeGrid(grid);
        styles = transposeGrid(styles);
      }

      const dRow = target.row - origin.row;
      const dCol = target.col - origin.col;

      if (options?.selection && mode !== "formats") {
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

      const write = (s: Sheet) => {
        const cells = { ...s.cells };
        const nextStyles = { ...s.styles };
        const cachedValues = { ...s.cachedValues };
        let touched = false;

        grid.forEach((row, ri) => {
          row.forEach((raw, ci) => {
            const r = target.row + ri;
            const c = target.col + ci;
            if (r >= s.numRows || c >= s.numCols) return;
            const key = cellKey(r, c);
            const locked = !!s.styles[key]?.locked;

            if (mode === "formats") {
              const style = styles[ri]?.[ci];
              if (!style) return;
              const { locked: _ignored, ...rest } = style;
              nextStyles[key] = { ...(nextStyles[key] ?? {}), ...rest };
              touched = true;
              return;
            }

            if (locked) return;

            const value = resolveValue({
              raw,
              mode,
              internal,
              dRow,
              dCol,
              evaluator: options?.evaluator,
              originRow: origin.row,
              originCol: origin.col,
              ri,
              ci,
            });

            if (options?.evaluator && mode !== "text") {
              const rejection = rejectCellValue(
                s,
                r,
                c,
                value,
                options.evaluator,
              );
              if (rejection) {
                options.onReject?.(rejection.message);
                return;
              }
            }

            if (value === "") delete cells[key];
            else cells[key] = value;
            delete cachedValues[key];

            if (mode === "all" || mode === "transpose") {
              const style = styles[ri]?.[ci];
              if (style) {
                const { locked: _ignored, ...rest } = style;
                nextStyles[key] = { ...(nextStyles[key] ?? {}), ...rest };
              }
            }
            touched = true;
          });
        });

        if (!touched) return null;
        return { cells, styles: nextStyles, cachedValues };
      };

      const rows = Math.max(grid.length, 1);
      const cols = Math.max(...grid.map((r) => r.length), 1);
      const result = {
        r1: target.row,
        c1: target.col,
        r2: target.row + rows - 1,
        c2: target.col + cols - 1,
      };

      // Merge check needs the live sheet — read via a no-op patch peek.
      let blocked = false;
      const guardWrite = (s: Sheet) => {
        const guard = checkPasteMerge(s, result);
        if (!guard.ok) {
          options?.onReject?.(guard.message);
          blocked = true;
          return null;
        }
        return write(s);
      };

      if (options?.patchSheet) {
        options.patchSheet(guardWrite);
      } else {
        updateSheet((s) => {
          const patch = guardWrite(s);
          if (!patch) return s;
          return { ...s, ...patch };
        });
      }

      if (blocked) return result;

      options?.onReject?.(statusFor(mode, result));
      return result;
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
