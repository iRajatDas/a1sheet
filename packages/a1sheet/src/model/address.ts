/**
 * A1 <-> `{row, col}` conversion and internal cell keys.
 *
 * Ported verbatim from ref/formulaEngine.js:16-35. These are the lowest-level
 * primitives in the library; the formula tokenizer, the XLSX reader/writer, and
 * the grid all depend on them agreeing exactly.
 */
import type { CellKey, CellPos, Range } from "./types.js";

/** 0 -> "A", 25 -> "Z", 26 -> "AA". */
export function colToLetters(n: number): string {
  let s = "";
  let v = n + 1;
  while (v > 0) {
    const m = (v - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    v = Math.floor((v - 1) / 26);
  }
  return s;
}

/** "A" -> 0, "Z" -> 25, "AA" -> 26. Expects uppercase. */
export function lettersToCol(s: string): number {
  let n = 0;
  for (let i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
  return n - 1;
}

/**
 * "A1", "$A$1", "a1" -> `{row, col}`. Absolute markers are accepted and
 * discarded; use the tokenizer when you need the `$` flags preserved.
 *
 * Returns `{row: 0, col: 0}` for unparseable input rather than throwing —
 * matches the POC, and keeps formula evaluation from blowing up on bad refs.
 */
export function parseCellRef(ref: string): CellPos {
  const m = ref.match(/^\$?([A-Za-z]+)\$?(\d+)$/);
  if (!m?.[1] || !m[2]) return { row: 0, col: 0 };
  return { col: lettersToCol(m[1].toUpperCase()), row: parseInt(m[2], 10) - 1 };
}

/** `{row, col}` -> "A1". */
export function toA1(row: number, col: number): string {
  return `${colToLetters(col)}${row + 1}`;
}

/** The internal map key for a cell. */
export function cellKey(row: number, col: number): CellKey {
  return `${row}_${col}`;
}

/** Inverse of `cellKey`. */
export function parseCellKey(key: string): CellPos {
  const i = key.indexOf("_");
  return {
    row: Number(key.slice(0, i)),
    col: Number(key.slice(i + 1)),
  };
}

/** Reorders a range so `r1 <= r2` and `c1 <= c2`. */
export function normalizeRange(r: Range): Range {
  return {
    r1: Math.min(r.r1, r.r2),
    c1: Math.min(r.c1, r.c2),
    r2: Math.max(r.r1, r.r2),
    c2: Math.max(r.c1, r.c2),
  };
}

/** Parses "A1" or "A1:B10" into a range. Returns null if it is not a reference. */
export function parseRangeRef(text: string): Range | null {
  const parts = text.trim().split(":");
  if (parts.length > 2) return null;
  const isRef = (s: string | undefined) =>
    !!s && /^\$?[A-Za-z]{1,3}\$?\d{1,7}$/.test(s);
  if (!isRef(parts[0])) return null;
  const a = parseCellRef(parts[0] as string);
  if (parts.length === 1) return { r1: a.row, c1: a.col, r2: a.row, c2: a.col };
  if (!isRef(parts[1])) return null;
  const b = parseCellRef(parts[1] as string);
  return normalizeRange({ r1: a.row, c1: a.col, r2: b.row, c2: b.col });
}

export function rangeContains(r: Range, row: number, col: number): boolean {
  const n = normalizeRange(r);
  return row >= n.r1 && row <= n.r2 && col >= n.c1 && col <= n.c2;
}
