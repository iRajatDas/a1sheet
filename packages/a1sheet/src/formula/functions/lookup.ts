/**
 * Lookups that take their arrays as values rather than as ranges.
 *
 * VLOOKUP, INDEX, and MATCH predate this and are special-cased in `evaluate.ts`
 * because they were written to need the range AST node. XLOOKUP does not: its
 * arguments can be any array, including one another function computed, which is
 * most of why it replaced VLOOKUP.
 */
import {
  at,
  colCount,
  type FormulaArg,
  type FormulaValue,
  flatten,
  type Matrix,
  rowCount,
  toMatrix,
  toNumber,
  toScalar,
  toText,
} from "../values.js";
import type { FormulaFunction } from "./registry.js";

/**
 * Lookup-key equality: numeric when both sides are numeric, textual otherwise,
 * and case-insensitive, which is how a spreadsheet compares text.
 */
function keyEquals(a: FormulaValue, b: FormulaValue): boolean {
  const an = typeof a === "number" ? a : Number.parseFloat(a as string);
  const bn = typeof b === "number" ? b : Number.parseFloat(b as string);
  if (!Number.isNaN(an) && !Number.isNaN(bn)) return an === bn;
  return toText(a).toLowerCase() === toText(b).toLowerCase();
}

function compareKeys(a: FormulaValue, b: FormulaValue): number {
  const an = typeof a === "number" ? a : Number.parseFloat(a as string);
  const bn = typeof b === "number" ? b : Number.parseFloat(b as string);
  if (!Number.isNaN(an) && !Number.isNaN(bn)) return an - bn;
  return toText(a).toLowerCase().localeCompare(toText(b).toLowerCase());
}

/** A lookup vector as a flat list, whichever axis it runs along. */
function asVector(m: Matrix): FormulaValue[] {
  return flatten([m]);
}

const MATCH_EXACT = 0;
const MATCH_NEXT_SMALLER = -1;
const MATCH_NEXT_LARGER = 1;
const MATCH_WILDCARD = 2;

/**
 * The index of a match, or -1.
 *
 * `mode` follows XMATCH: 0 exact, -1 exact or next smaller, 1 exact or next
 * larger, 2 wildcard. `direction` of -1 searches from the end.
 */
export function findMatch(
  needle: FormulaValue,
  haystack: readonly FormulaValue[],
  mode: number,
  direction: number,
): number {
  const indices =
    direction < 0
      ? haystack.map((_, i) => haystack.length - 1 - i)
      : haystack.map((_, i) => i);

  if (mode === MATCH_WILDCARD) {
    const pattern = wildcardToRegExp(toText(needle));
    for (const i of indices) {
      if (pattern.test(toText(haystack[i]))) return i;
    }
    return -1;
  }

  for (const i of indices) {
    if (keyEquals(haystack[i], needle)) return i;
  }
  if (mode === MATCH_EXACT) return -1;

  // Approximate: the closest value on the requested side. Scanning rather than
  // bisecting, because an unsorted vector would make a binary search silently
  // wrong and a spreadsheet's lookup vectors are rarely long enough to care.
  let best = -1;
  let bestKey: FormulaValue;
  for (const i of indices) {
    const value = haystack[i];
    const cmp = compareKeys(value, needle);
    const onSide = mode === MATCH_NEXT_SMALLER ? cmp < 0 : cmp > 0;
    if (!onSide) continue;
    if (
      best === -1 ||
      (mode === MATCH_NEXT_SMALLER
        ? compareKeys(value, bestKey) > 0
        : compareKeys(value, bestKey) < 0)
    ) {
      best = i;
      bestKey = value;
    }
  }
  return best;
}

/** `*` and `?` as Excel means them, everything else literal. */
export function wildcardToRegExp(pattern: string): RegExp {
  const escaped = pattern.replace(/[.+^${}()|[\]\\]/g, "\\$&");
  return new RegExp(`^${escaped.replace(/\*/g, ".*").replace(/\?/g, ".")}$`, "i");
}

/** The row or column of `results` that lines up with a hit in the lookup vector. */
function resultAt(results: Matrix, index: number, alongRows: boolean): FormulaArg {
  if (alongRows) {
    const row = results[index];
    if (!row) return "#REF!";
    return row.length === 1 ? (row[0] as FormulaValue) : [row];
  }
  const height = rowCount(results);
  if (index >= colCount(results)) return "#REF!";
  if (height === 1) return at(results, 0, index);
  return results.map((row) => [row[index]]);
}

export const lookupFunctions: Record<string, FormulaFunction> = {
  /**
   * VLOOKUP(lookup, table, colIndex, [approximate]).
   *
   * Used to be special-cased in the evaluator because it needed the range AST
   * node to recover the table's 2D shape. Now that a range evaluates to a matrix
   * there is nothing to recover, so it is an ordinary function — and it works on
   * a computed table, which the special case could not do.
   *
   * The fourth argument defaults to TRUE in Excel, meaning approximate. That
   * default is a well-known footgun and it is still the default here, because
   * matching Excel matters more than protecting people from Excel.
   */
  VLOOKUP: (a) => {
    if (a[0] === undefined || a[1] === undefined) return "#VALUE!";
    const table = toMatrix(a[1]);
    const colIndex = Math.round(toNumber(a[2] ?? 1)) - 1;
    const approximate = a[3] === undefined ? true : toNumber(a[3]) !== 0;
    const firstColumn = table.map((row) => row[0]);
    const index = findMatch(
      toScalar(a[0]),
      firstColumn,
      approximate ? MATCH_NEXT_SMALLER : MATCH_EXACT,
      1,
    );
    if (index === -1) return "#N/A";
    return at(table, index, colIndex) ?? "#REF!";
  },

  /** MATCH(lookup, vector, [type]). 1 is approximate, 0 exact, -1 descending. */
  MATCH: (a) => {
    if (a[0] === undefined || a[1] === undefined) return "#VALUE!";
    const type = a[2] === undefined ? 1 : Math.round(toNumber(a[2]));
    // MATCH's type and XMATCH's mode disagree on sign: MATCH 1 means "next
    // smaller" and -1 "next larger", which is the opposite of XMATCH.
    const mode =
      type === 0 ? MATCH_EXACT : type > 0 ? MATCH_NEXT_SMALLER : MATCH_NEXT_LARGER;
    const index = findMatch(toScalar(a[0]), asVector(toMatrix(a[1])), mode, 1);
    return index === -1 ? "#N/A" : index + 1;
  },

  /**
   * INDEX(array, row, [col]).
   *
   * A zero row or column means the whole column or row, which is what makes
   * `INDEX(t, 0, 2)` a way to take a column — and why this returns a matrix.
   */
  INDEX: (a) => {
    if (a[0] === undefined) return "#VALUE!";
    const m = toMatrix(a[0]);
    const rowNum = Math.round(toNumber(a[1] ?? 0));
    const colNum = a[2] === undefined ? 0 : Math.round(toNumber(a[2]));

    // A single-row or single-column array takes one index, addressing along it.
    if (a[2] === undefined && (rowCount(m) === 1 || colCount(m) === 1)) {
      const value = asVector(m)[rowNum - 1];
      return value === undefined ? "#REF!" : value;
    }

    if (rowNum === 0 && colNum === 0) return m;
    if (rowNum === 0) {
      if (colNum > colCount(m)) return "#REF!";
      return m.map((row) => [row[colNum - 1]]);
    }
    if (colNum === 0) {
      const row = m[rowNum - 1];
      return row ? [row] : "#REF!";
    }
    return at(m, rowNum - 1, colNum - 1) ?? "#REF!";
  },

  /**
   * XLOOKUP(lookup, lookupArray, returnArray, [ifNotFound], [matchMode],
   * [searchMode]). Returns the whole matching row or column, so it composes with
   * the array functions in a way VLOOKUP cannot.
   */
  XLOOKUP: (a) => {
    if (a[0] === undefined || a[1] === undefined || a[2] === undefined) {
      return "#VALUE!";
    }
    const lookup = toScalar(a[0]);
    const source = toMatrix(a[1]);
    const results = toMatrix(a[2]);
    const mode = a[4] === undefined ? MATCH_EXACT : Math.round(toNumber(a[4]));
    const searchMode = a[5] === undefined ? 1 : Math.round(toNumber(a[5]));

    // A lookup vector that is a single row means the arrays run left to right.
    const alongRows = rowCount(source) > 1 || colCount(source) === 1;
    const index = findMatch(lookup, asVector(source), mode, searchMode);
    if (index === -1) return a[3] ?? "#N/A";
    return resultAt(results, index, alongRows);
  },

  /** XMATCH(lookup, lookupArray, [matchMode], [searchMode]) — 1-based. */
  XMATCH: (a) => {
    if (a[0] === undefined || a[1] === undefined) return "#VALUE!";
    const mode = a[2] === undefined ? MATCH_EXACT : Math.round(toNumber(a[2]));
    const searchMode = a[3] === undefined ? 1 : Math.round(toNumber(a[3]));
    const index = findMatch(
      toScalar(a[0]),
      asVector(toMatrix(a[1])),
      mode,
      searchMode,
    );
    return index === -1 ? "#N/A" : index + 1;
  },

  /** LOOKUP(lookup, vector, [result]) — the old approximate form. */
  LOOKUP: (a) => {
    if (a[0] === undefined || a[1] === undefined) return "#VALUE!";
    const vector = asVector(toMatrix(a[1]));
    const index = findMatch(toScalar(a[0]), vector, MATCH_NEXT_SMALLER, 1);
    if (index === -1) return "#N/A";
    const results = a[2] === undefined ? vector : asVector(toMatrix(a[2]));
    return results[index] ?? "#N/A";
  },

  /** HLOOKUP(lookup, table, rowIndex, [approximate]) — VLOOKUP on its side. */
  HLOOKUP: (a) => {
    if (a[0] === undefined || a[1] === undefined) return "#VALUE!";
    const table = toMatrix(a[1]);
    const header = table[0] ?? [];
    const rowIndex = Math.round(toNumber(a[2] ?? 1)) - 1;
    const approximate = a[3] === undefined ? true : toNumber(a[3]) !== 0;
    const index = findMatch(
      toScalar(a[0]),
      [...header],
      approximate ? MATCH_NEXT_SMALLER : MATCH_EXACT,
      1,
    );
    if (index === -1) return "#N/A";
    return at(table, rowIndex, index) ?? "#REF!";
  },

  ISBLANK: (a) => {
    const v = toScalar(a[0]);
    return v === undefined || v === "";
  },

  COUNTBLANK: (a) =>
    flatten([a[0]]).filter((v) => v === undefined || v === "").length,

  /** COUNTIF(range, criterion) with Excel's comparison-in-a-string criteria. */
  COUNTIF: (a) => {
    if (a[0] === undefined) return 0;
    const test = criterion(toScalar(a[1]));
    return flatten([a[0]]).filter(test).length;
  },

  /** SUMIF(range, criterion, [sumRange]). */
  SUMIF: (a) => {
    if (a[0] === undefined) return 0;
    const test = criterion(toScalar(a[1]));
    const subject = flatten([a[0]]);
    const summed = a[2] === undefined ? subject : flatten([a[2]]);
    let total = 0;
    for (const [i, v] of subject.entries()) {
      if (test(v)) total += toNumber(summed[i]);
    }
    return total;
  },

  AVERAGEIF: (a) => {
    if (a[0] === undefined) return "#DIV/0!";
    const test = criterion(toScalar(a[1]));
    const subject = flatten([a[0]]);
    const averaged = a[2] === undefined ? subject : flatten([a[2]]);
    const picked = subject
      .map((v, i) => (test(v) ? toNumber(averaged[i]) : null))
      .filter((n): n is number => n !== null);
    if (picked.length === 0) return "#DIV/0!";
    return picked.reduce((x, y) => x + y, 0) / picked.length;
  },
};

const CRITERION_RE = /^(<=|>=|<>|<|>|=)?(.*)$/;

/**
 * Turns `">10"`, `"<>x"`, or `"Ars*"` into a predicate.
 *
 * Excel encodes the comparison inside the criterion string, so `COUNTIF(r, ">10")`
 * has to parse its own argument. A bare value is an equality test, with wildcards
 * honoured.
 */
export function criterion(spec: FormulaValue): (v: FormulaValue) => boolean {
  const text = toText(spec);
  const m = text.match(CRITERION_RE);
  const op = m?.[1] ?? "=";
  const operand = m?.[2] ?? "";

  if ((op === "=" || op === "<>") && /[*?]/.test(operand)) {
    const pattern = wildcardToRegExp(operand);
    return (v) => pattern.test(toText(v)) === (op === "=");
  }

  const target: FormulaValue = Number.isNaN(Number.parseFloat(operand))
    ? operand
    : Number.parseFloat(operand);

  return (v) => {
    if (op === "=") return keyEquals(v, target);
    if (op === "<>") return !keyEquals(v, target);
    const cmp = compareKeys(v, target);
    if (op === "<") return cmp < 0;
    if (op === "<=") return cmp <= 0;
    if (op === ">") return cmp > 0;
    return cmp >= 0;
  };
}
