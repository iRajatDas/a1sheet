/**
 * Autofill destination validation — typed codes, Sheets-compatible messages.
 *
 * Headless: the fill handle and any custom UI call `previewFillCheck` before
 * writing. Branch on `code`, never on message text.
 */
import { normalizeRange } from "./address.js";
import { getMergeAt } from "./sheet.js";
import type { Range, Sheet } from "./types.js";

export const AUTOFILL_CODES = [
  "DEST_EQUALS_SOURCE",
  "DEST_SOURCE_DIMENSIONS_BOTH_DIFFERENT",
  "PARTIAL_MERGE",
  "CONSTRAIN_DEST_FAILED",
] as const;

export type AutofillCode = (typeof AUTOFILL_CODES)[number];

export type FillCheckResult =
  | { readonly ok: true; readonly dest: Range }
  | {
      readonly ok: false;
      readonly code: AutofillCode;
      readonly message: string;
    };

export interface PreviewFillCheckOptions {
  readonly source: Range;
  readonly target: { readonly row: number; readonly col: number };
}

const MSG = {
  DEST_SOURCE_DIMENSIONS_BOTH_DIFFERENT:
    "You can't auto-fill when both axes of the source and destination are different.",
  PARTIAL_MERGE: "You can't auto-fill a range that partially intersects a merge.",
  CONSTRAIN_DEST_FAILED: "You can't auto-fill past the edge of the sheet.",
  DEST_EQUALS_SOURCE: "Nothing to fill — the destination matches the source.",
} as const satisfies Record<AutofillCode, string>;

/** Destination rectangle covering source and the drag target. */
export function fillDestination(
  source: Range,
  target: { row: number; col: number },
): Range {
  const b = normalizeRange(source);
  return {
    r1: Math.min(b.r1, target.row),
    c1: Math.min(b.c1, target.col),
    r2: Math.max(b.r2, target.row),
    c2: Math.max(b.c2, target.col),
  };
}

/**
 * Whether the destination expands beyond the source on both axes at once.
 *
 * Sheets only allows a fill that shares the source's row span (left/right) or
 * its column span (up/down). Expanding both is rejected.
 */
function bothAxesDifferent(source: Range, dest: Range): boolean {
  const sameCols = dest.c1 === source.c1 && dest.c2 === source.c2;
  const sameRows = dest.r1 === source.r1 && dest.r2 === source.r2;
  return !sameCols && !sameRows;
}

/**
 * A merge is partially intersected when some — but not all — of its cells lie
 * inside `dest`. Entirely inside or entirely outside is fine.
 */
function partialMerge(sheet: Sheet, dest: Range): boolean {
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

export function previewFillCheck(
  sheet: Sheet,
  options: PreviewFillCheckOptions,
): FillCheckResult {
  const source = normalizeRange(options.source);
  const dest = fillDestination(source, options.target);

  if (
    dest.r1 === source.r1 &&
    dest.c1 === source.c1 &&
    dest.r2 === source.r2 &&
    dest.c2 === source.c2
  ) {
    return {
      ok: false,
      code: "DEST_EQUALS_SOURCE",
      message: MSG.DEST_EQUALS_SOURCE,
    };
  }

  if (
    dest.r1 < 0 ||
    dest.c1 < 0 ||
    dest.r2 >= sheet.numRows ||
    dest.c2 >= sheet.numCols
  ) {
    return {
      ok: false,
      code: "CONSTRAIN_DEST_FAILED",
      message: MSG.CONSTRAIN_DEST_FAILED,
    };
  }

  if (bothAxesDifferent(source, dest)) {
    return {
      ok: false,
      code: "DEST_SOURCE_DIMENSIONS_BOTH_DIFFERENT",
      message: MSG.DEST_SOURCE_DIMENSIONS_BOTH_DIFFERENT,
    };
  }

  if (partialMerge(sheet, dest)) {
    return {
      ok: false,
      code: "PARTIAL_MERGE",
      message: MSG.PARTIAL_MERGE,
    };
  }

  return { ok: true, dest };
}
