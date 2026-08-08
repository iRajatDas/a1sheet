/**
 * Headless find & replace over sheet cell text and formulas.
 *
 * Consumers compose UI on top; this module only searches and mutates cells.
 */
import { cellKey, normalizeRange } from "./address.js";
import type { CellKey, Range, RawCell, Sheet } from "./types.js";

export interface FindReplaceOptions {
  /** Substring or whole-cell needle. */
  find: string;
  /** Restrict to this rectangle. Absent means the used area of the sheet. */
  range?: Range;
  matchCase?: boolean;
  matchEntireCell?: boolean;
  /** Search display-ready raw cell text (includes formulas with `=`). */
  searchFormulas?: boolean;
}

export interface FindHit {
  row: number;
  col: number;
  /** Raw cell text that matched. */
  raw: string;
}

export interface ReplaceAllResult {
  sheet: Sheet;
  count: number;
}

function usedBounds(cells: Sheet["cells"]): Range {
  let r2 = 0;
  let c2 = 0;
  for (const key of Object.keys(cells) as CellKey[]) {
    const sep = key.indexOf("_");
    const r = Number(key.slice(0, sep));
    const c = Number(key.slice(sep + 1));
    if (r > r2) r2 = r;
    if (c > c2) c2 = c;
  }
  return { r1: 0, c1: 0, r2, c2 };
}

function matches(raw: string, options: FindReplaceOptions): boolean {
  const { find, matchCase = false, matchEntireCell = false } = options;
  if (find === "") return false;
  const hay = matchCase ? raw : raw.toLowerCase();
  const needle = matchCase ? find : find.toLowerCase();
  if (matchEntireCell) return hay === needle;
  return hay.includes(needle);
}

function replaceIn(
  raw: string,
  find: string,
  replace: string,
  options: Pick<FindReplaceOptions, "matchCase" | "matchEntireCell">,
): string {
  const { matchCase = false, matchEntireCell = false } = options;
  if (matchEntireCell) {
    const hay = matchCase ? raw : raw.toLowerCase();
    const needle = matchCase ? find : find.toLowerCase();
    return hay === needle ? replace : raw;
  }
  if (matchCase) return raw.split(find).join(replace);
  const lower = find.toLowerCase();
  let out = "";
  let i = 0;
  const src = raw;
  while (i < src.length) {
    const slice = src.slice(i);
    const idx = slice.toLowerCase().indexOf(lower);
    if (idx < 0) {
      out += slice;
      break;
    }
    out += slice.slice(0, idx) + replace;
    i += idx + find.length;
  }
  return out;
}

/** All matching cells in row-major order. */
export function findAll(sheet: Sheet, options: FindReplaceOptions): FindHit[] {
  const bounds = normalizeRange(options.range ?? usedBounds(sheet.cells));
  const hits: FindHit[] = [];
  for (let r = bounds.r1; r <= bounds.r2; r++) {
    for (let c = bounds.c1; c <= bounds.c2; c++) {
      const raw = sheet.cells[cellKey(r, c)];
      if (raw === undefined || raw === "") continue;
      if (options.searchFormulas === false && raw.startsWith("=")) continue;
      if (matches(raw, options)) hits.push({ row: r, col: c, raw });
    }
  }
  return hits;
}

/**
 * Next hit after `after`, wrapping to the first hit. Returns null when nothing
 * matches.
 */
export function findNext(
  sheet: Sheet,
  options: FindReplaceOptions & { after?: { row: number; col: number } },
): FindHit | null {
  const hits = findAll(sheet, options);
  if (hits.length === 0) return null;
  const after = options.after;
  if (!after) return hits[0] ?? null;
  for (const hit of hits) {
    if (hit.row > after.row || (hit.row === after.row && hit.col > after.col)) {
      return hit;
    }
  }
  return hits[0] ?? null;
}

export function replaceAll(
  sheet: Sheet,
  options: FindReplaceOptions & { replace: string },
): ReplaceAllResult {
  const hits = findAll(sheet, options);
  if (hits.length === 0) return { sheet, count: 0 };
  const cells: Record<CellKey, RawCell> = { ...sheet.cells };
  for (const hit of hits) {
    const key = cellKey(hit.row, hit.col);
    const raw = cells[key];
    if (raw === undefined) continue;
    cells[key] = replaceIn(raw, options.find, options.replace, options);
  }
  return { sheet: { ...sheet, cells }, count: hits.length };
}

export function replacedStatus(count: number): string {
  const noun = count === 1 ? "occurrence" : "occurrences";
  return `Replaced ${count} ${noun}.`;
}
