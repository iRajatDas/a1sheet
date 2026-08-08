/**
 * Merge intersection guards shared by paste, sort, and filter.
 */
import { cellKey, normalizeRange } from "./address.js";
import { getMergeAt } from "./sheet.js";
import type { Range, Sheet } from "./types.js";

export const MERGE_GUARD_CODES = [
  "PASTE_PARTIAL_MERGE",
  "SORT_PARTIAL_MERGE",
  "FILTER_PARTIAL_MERGE",
] as const;

export type MergeGuardCode = (typeof MERGE_GUARD_CODES)[number];

export type MergeGuardResult =
  | { readonly ok: true }
  | {
      readonly ok: false;
      readonly code: MergeGuardCode;
      readonly message: string;
    };

const MSG = {
  PASTE_PARTIAL_MERGE:
    "You can't perform a paste that partially intersects a merge.",
  SORT_PARTIAL_MERGE: "Merges must be entirely within the sort range.",
  FILTER_PARTIAL_MERGE: "Merges must be entirely within the filter range.",
} as const satisfies Record<MergeGuardCode, string>;

/** True when some — but not all — cells of a merge lie inside `range`. */
export function rangePartiallyIntersectsMerge(
  sheet: Sheet,
  range: Range,
): boolean {
  const dest = normalizeRange(range);
  if (sheet.merges.length === 0) return false;

  const seen = new Set<string>();
  for (let r = dest.r1; r <= dest.r2; r++) {
    for (let c = dest.c1; c <= dest.c2; c++) {
      const merge = getMergeAt(sheet, r, c);
      if (!merge) continue;
      const key = `${merge.r1}_${merge.c1}_${merge.r2}_${merge.c2}`;
      if (seen.has(key)) continue;
      seen.add(key);
      const fullyInside =
        merge.r1 >= dest.r1 &&
        merge.r2 <= dest.r2 &&
        merge.c1 >= dest.c1 &&
        merge.c2 <= dest.c2;
      if (!fullyInside) return true;
    }
  }
  return false;
}

/** True when any merge overlaps `range` without being fully contained. */
export function rangeContainsPartialMerge(sheet: Sheet, range: Range): boolean {
  return rangePartiallyIntersectsMerge(sheet, range);
}

export function checkPasteMerge(
  sheet: Sheet,
  dest: Range,
): MergeGuardResult {
  if (rangePartiallyIntersectsMerge(sheet, dest)) {
    return {
      ok: false,
      code: "PASTE_PARTIAL_MERGE",
      message: MSG.PASTE_PARTIAL_MERGE,
    };
  }
  return { ok: true };
}

export function checkSortMerge(sheet: Sheet, range: Range): MergeGuardResult {
  if (rangeContainsPartialMerge(sheet, range)) {
    return {
      ok: false,
      code: "SORT_PARTIAL_MERGE",
      message: MSG.SORT_PARTIAL_MERGE,
    };
  }
  // Also block when a merge straddles the sort columns but sits inside the
  // row span — Sheets requires merges wholly inside.
  const b = normalizeRange(range);
  for (const merge of sheet.merges) {
    const rowOverlap = merge.r2 >= b.r1 && merge.r1 <= b.r2;
    const colOverlap = merge.c2 >= b.c1 && merge.c1 <= b.c2;
    if (!rowOverlap || !colOverlap) continue;
    const fullyInside =
      merge.r1 >= b.r1 &&
      merge.r2 <= b.r2 &&
      merge.c1 >= b.c1 &&
      merge.c2 <= b.c2;
    if (!fullyInside) {
      return {
        ok: false,
        code: "SORT_PARTIAL_MERGE",
        message: MSG.SORT_PARTIAL_MERGE,
      };
    }
  }
  return { ok: true };
}

export function checkFilterMerge(
  sheet: Sheet,
  range: Range,
): MergeGuardResult {
  if (rangeContainsPartialMerge(sheet, range)) {
    return {
      ok: false,
      code: "FILTER_PARTIAL_MERGE",
      message: MSG.FILTER_PARTIAL_MERGE,
    };
  }
  return { ok: true };
}

/** Used by tests — whether any cell key in range is covered by a merge. */
export function mergeKeysInRange(sheet: Sheet, range: Range): string[] {
  const b = normalizeRange(range);
  const keys: string[] = [];
  for (let r = b.r1; r <= b.r2; r++) {
    for (let c = b.c1; c <= b.c2; c++) {
      if (getMergeAt(sheet, r, c)) keys.push(cellKey(r, c));
    }
  }
  return keys;
}
