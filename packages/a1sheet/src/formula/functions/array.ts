/**
 * Dynamic-array functions: the ones that return a shape rather than a value.
 *
 * These are what a modern Excel workbook is built out of. `SORT(UNIQUE(...))`
 * replaces a helper column, `MAP` replaces a filled-down formula, and `HSTACK`
 * assembles a table out of computed columns — so a workbook using them has no
 * intermediate cells to fall back on. Without these, such a file is not partly
 * readable; it is a grid of `#NAME?`.
 */
import {
  at,
  colCount,
  type FormulaArg,
  type FormulaValue,
  flatten,
  isLambda,
  isMatrix,
  type LambdaValue,
  type Matrix,
  rowCount,
  toMatrix,
  toNumber,
  toScalar,
  toText,
} from "../values.js";
import type { FormulaFunction, FormulaHost } from "./registry.js";

/** A rectangular matrix of the given shape, filled by a callback. */
function build(
  rows: number,
  cols: number,
  cell: (row: number, col: number) => FormulaValue,
): Matrix {
  const out: FormulaValue[][] = [];
  for (let r = 0; r < rows; r++) {
    const row: FormulaValue[] = [];
    for (let c = 0; c < cols; c++) row.push(cell(r, c));
    out.push(row);
  }
  return out;
}

/** Excel's cap on a generated array, so a bad SEQUENCE cannot hang the sheet. */
const MAX_GENERATED_CELLS = 1_048_576;

function tooBig(rows: number, cols: number): boolean {
  return (
    !Number.isFinite(rows) ||
    !Number.isFinite(cols) ||
    rows < 1 ||
    cols < 1 ||
    rows * cols > MAX_GENERATED_CELLS
  );
}

/** A key that compares equal for values Excel treats as the same. */
function dedupeKey(v: FormulaValue): string {
  return typeof v === "number" ? `n${v}` : `s${toText(v).toLowerCase()}`;
}

/**
 * Joins a row's keys into one. NUL because a cell can hold any other character,
 * and a separator a cell can contain makes ["a", "b|c"] and ["a|b", "c"] the
 * same row. Written escaped rather than as a literal byte so the file stays
 * text — git treats a source file containing a raw NUL as binary and stops
 * diffing it.
 */
const KEY_SEPARATOR = "\u0000";

/**
 * Comparison for SORT: numbers before text, blanks last, case-insensitive.
 * Matches what a spreadsheet user expects rather than JS's default ordering.
 */
function compareValues(a: FormulaValue, b: FormulaValue): number {
  const aBlank = a === undefined || a === "";
  const bBlank = b === undefined || b === "";
  if (aBlank && bBlank) return 0;
  if (aBlank) return 1;
  if (bBlank) return -1;

  const an = typeof a === "number" ? a : Number.parseFloat(a as string);
  const bn = typeof b === "number" ? b : Number.parseFloat(b as string);
  const aNum = !Number.isNaN(an);
  const bNum = !Number.isNaN(bn);
  if (aNum && bNum) return an - bn;
  if (aNum !== bNum) return aNum ? -1 : 1;
  return toText(a).toLowerCase().localeCompare(toText(b).toLowerCase());
}

/** Sort keys and directions, from SORT's optional 2nd and 3rd arguments. */
function sortSpec(
  indexArg: FormulaArg | undefined,
  orderArg: FormulaArg | undefined,
): { keys: number[]; orders: number[] } {
  const keys = indexArg === undefined ? [1] : flatten([indexArg]).map(toNumber);
  const orders =
    orderArg === undefined ? [] : flatten([orderArg]).map((v) => toNumber(v));
  return { keys: keys.length ? keys : [1], orders };
}

const DESCENDING = -1;

/** Rows of `m` reordered by the given keys. `byColumn` sorts columns instead. */
function sortMatrix(
  m: Matrix,
  keys: readonly number[],
  orders: readonly number[],
  byColumn: boolean,
): Matrix {
  const source = byColumn ? transpose(m) : m;
  const indices = source.map((_, i) => i);

  indices.sort((ia, ib) => {
    for (const [k, key] of keys.entries()) {
      const column = Math.round(key) - 1;
      const direction = (orders[k] ?? orders[0] ?? 1) === DESCENDING ? -1 : 1;
      const cmp = compareValues(source[ia]?.[column], source[ib]?.[column]);
      if (cmp !== 0) return cmp * direction;
    }
    // A stable tie-break, so equal rows keep their original order.
    return ia - ib;
  });

  const sorted = indices.map((i) => source[i] as readonly FormulaValue[]);
  return byColumn ? transpose(sorted) : sorted;
}

function transpose(m: Matrix): Matrix {
  const rows = rowCount(m);
  const cols = colCount(m);
  return build(cols, rows, (r, c) => at(m, c, r));
}

/** Picks columns (or rows) by 1-based index, negative counting from the end. */
function choose(m: Matrix, picks: readonly number[], byColumn: boolean): Matrix {
  const source = byColumn ? transpose(m) : m;
  const out: (readonly FormulaValue[])[] = [];
  for (const pick of picks) {
    const n = Math.round(pick);
    const index = n < 0 ? source.length + n : n - 1;
    const row = source[index];
    if (!row) return [["#VALUE!"]];
    out.push(row);
  }
  return byColumn ? transpose(out) : out;
}

/** Stacks matrices along one axis, padding the short side with #N/A as Excel does. */
function stack(args: readonly FormulaArg[], vertical: boolean): Matrix {
  const parts = args.filter((a) => !isLambda(a)).map(toMatrix);
  if (parts.length === 0) return [[undefined]];

  if (vertical) {
    const width = Math.max(...parts.map(colCount));
    const out: FormulaValue[][] = [];
    for (const part of parts) {
      for (const row of part) {
        out.push(Array.from({ length: width }, (_, c) => row[c] ?? "#N/A"));
      }
    }
    return out;
  }

  const height = Math.max(...parts.map(rowCount));
  return build(
    height,
    parts.reduce((sum, p) => sum + colCount(p), 0),
    () => undefined,
  ).map((_, r) => {
    const row: FormulaValue[] = [];
    for (const part of parts) {
      const width = colCount(part);
      for (let c = 0; c < width; c++) {
        row.push(r < rowCount(part) ? at(part, r, c) : "#N/A");
      }
    }
    return row;
  });
}

/** Applies a lambda that may not be one, so a bad argument is an error not a crash. */
function withLambda(
  arg: FormulaArg | undefined,
  host: FormulaHost,
  use: (call: (args: readonly FormulaArg[]) => FormulaArg) => FormulaArg,
): FormulaArg {
  if (arg === undefined || !isLambda(arg)) return "#VALUE!";
  const lambda = arg as LambdaValue;
  return use((args) => host.call(lambda, args));
}

export const arrayFunctions: Record<string, FormulaFunction> = {
  ROWS: (a) => (a[0] === undefined ? 0 : rowCount(toMatrix(a[0]))),
  COLUMNS: (a) => (a[0] === undefined ? 0 : colCount(toMatrix(a[0]))),

  TRANSPOSE: (a) => (a[0] === undefined ? "#VALUE!" : transpose(toMatrix(a[0]))),

  HSTACK: (a) => stack(a, false),
  VSTACK: (a) => stack(a, true),

  /**
   * SEQUENCE(rows, [cols], [start], [step]). A column by default, because that is
   * what one argument means in Excel.
   */
  SEQUENCE: (a) => {
    const rows = Math.round(toNumber(a[0] ?? 1));
    const cols = a[1] === undefined ? 1 : Math.round(toNumber(a[1]));
    if (tooBig(rows, cols)) return "#VALUE!";
    const start = a[2] === undefined ? 1 : toNumber(a[2]);
    const step = a[3] === undefined ? 1 : toNumber(a[3]);
    return build(rows, cols, (r, c) => start + (r * cols + c) * step);
  },

  /**
   * UNIQUE(array, [byColumn], [exactlyOnce]). Preserves first-seen order, which
   * is what makes `SORT(UNIQUE(...))` produce a stable list.
   */
  UNIQUE: (a) => {
    if (a[0] === undefined) return "#VALUE!";
    const byColumn = toNumber(a[1] ?? false) !== 0;
    const exactlyOnce = toNumber(a[2] ?? false) !== 0;
    const source = byColumn ? transpose(toMatrix(a[0])) : toMatrix(a[0]);

    const counts = new Map<string, number>();
    const order: string[] = [];
    const byKey = new Map<string, readonly FormulaValue[]>();
    for (const row of source) {
      const key = row.map(dedupeKey).join(KEY_SEPARATOR);
      const seen = counts.get(key) ?? 0;
      counts.set(key, seen + 1);
      if (seen === 0) {
        order.push(key);
        byKey.set(key, row);
      }
    }

    const kept = order.filter((k) => !exactlyOnce || counts.get(k) === 1);
    const out = kept.map((k) => byKey.get(k) as readonly FormulaValue[]);
    if (out.length === 0) return "#CALC!";
    return byColumn ? transpose(out) : out;
  },

  /** SORT(array, [sortIndex], [sortOrder], [byColumn]). */
  SORT: (a) => {
    if (a[0] === undefined) return "#VALUE!";
    const { keys, orders } = sortSpec(a[1], a[2]);
    const byColumn = toNumber(a[3] ?? false) !== 0;
    return sortMatrix(toMatrix(a[0]), keys, orders, byColumn);
  },

  /** SORTBY(array, by1, [order1], by2, [order2], …). */
  SORTBY: (a) => {
    if (a[0] === undefined || a[1] === undefined) return "#VALUE!";
    const source = toMatrix(a[0]);
    // Each `by` is a column of keys the same height as the array; sorting a
    // column of indices by them and then reordering keeps the two in step.
    const bys: { keys: Matrix; direction: number }[] = [];
    for (let i = 1; i < a.length; i += 2) {
      const by = a[i];
      if (by === undefined) break;
      bys.push({
        keys: toMatrix(by),
        direction: toNumber(a[i + 1] ?? 1) === DESCENDING ? -1 : 1,
      });
    }

    const indices = source.map((_, i) => i);
    indices.sort((ia, ib) => {
      for (const { keys, direction } of bys) {
        const cmp = compareValues(at(keys, ia, 0), at(keys, ib, 0));
        if (cmp !== 0) return cmp * direction;
      }
      return ia - ib;
    });
    return indices.map((i) => source[i] as readonly FormulaValue[]);
  },

  /** FILTER(array, include, [ifEmpty]). */
  FILTER: (a) => {
    if (a[0] === undefined || a[1] === undefined) return "#VALUE!";
    const source = toMatrix(a[0]);
    const include = toMatrix(a[1]);
    // The mask may be a row or a column; whichever axis it matches is the one
    // being filtered.
    const byRow = rowCount(include) === rowCount(source) && rowCount(source) > 1;

    if (byRow) {
      const kept = source.filter((_, r) => toNumber(at(include, r, 0)) !== 0);
      return kept.length ? kept : (a[2] ?? "#CALC!");
    }
    const width = colCount(source);
    const keep: number[] = [];
    for (let c = 0; c < width; c++) {
      if (toNumber(at(include, 0, c)) !== 0) keep.push(c);
    }
    if (keep.length === 0) return a[2] ?? "#CALC!";
    return source.map((row) => keep.map((c) => row[c]));
  },

  CHOOSECOLS: (a) =>
    a[0] === undefined
      ? "#VALUE!"
      : choose(toMatrix(a[0]), flatten(a.slice(1)).map(toNumber), true),
  CHOOSEROWS: (a) =>
    a[0] === undefined
      ? "#VALUE!"
      : choose(toMatrix(a[0]), flatten(a.slice(1)).map(toNumber), false),

  /** TAKE(array, rows, [cols]). Negative counts from the end. */
  TAKE: (a) => {
    if (a[0] === undefined) return "#VALUE!";
    const m = toMatrix(a[0]);
    const rows = a[1] === undefined ? rowCount(m) : Math.round(toNumber(a[1]));
    const cols = a[2] === undefined ? colCount(m) : Math.round(toNumber(a[2]));
    const kept = rows < 0 ? m.slice(rows) : m.slice(0, rows);
    return kept.map((row) => (cols < 0 ? row.slice(cols) : row.slice(0, cols)));
  },

  /** DROP(array, rows, [cols]). Negative drops from the end. */
  DROP: (a) => {
    if (a[0] === undefined) return "#VALUE!";
    const m = toMatrix(a[0]);
    const rows = a[1] === undefined ? 0 : Math.round(toNumber(a[1]));
    const cols = a[2] === undefined ? 0 : Math.round(toNumber(a[2]));
    const kept = rows < 0 ? m.slice(0, rows) : m.slice(rows);
    const out = kept.map((row) =>
      cols < 0 ? row.slice(0, cols) : row.slice(cols),
    );
    return out.length && out[0]?.length ? out : "#CALC!";
  },

  /** TOROW/TOCOL(array). Flattens row-major, dropping nothing. */
  TOROW: (a) => (a[0] === undefined ? "#VALUE!" : [flatten([a[0]])]),
  TOCOL: (a) => (a[0] === undefined ? "#VALUE!" : flatten([a[0]]).map((v) => [v])),

  /**
   * FLATTEN(range, …). Sheets alias: every argument, row-major, into one column.
   * Prefer TOCOL for Excel-shaped workbooks.
   */
  FLATTEN: (a) => {
    if (a.length === 0) return "#VALUE!";
    return flatten(a).map((v) => [v]);
  },

  /** MAKEARRAY(rows, cols, lambda(row, col)). Indices are 1-based. */
  MAKEARRAY: (a, host) => {
    const rows = Math.round(toNumber(a[0]));
    const cols = Math.round(toNumber(a[1]));
    if (tooBig(rows, cols)) return "#VALUE!";
    return withLambda(a[2], host, (call) =>
      build(rows, cols, (r, c) => toScalar(call([r + 1, c + 1]))),
    );
  },

  /** MAP(array, …, lambda). One lambda argument per array, elementwise. */
  MAP: (a, host) => {
    const arrays = a.slice(0, -1).map(toMatrix);
    const first = arrays[0];
    if (!first) return "#VALUE!";
    return withLambda(a[a.length - 1], host, (call) =>
      build(rowCount(first), colCount(first), (r, c) =>
        toScalar(call(arrays.map((m) => at(m, r, c)))),
      ),
    );
  },

  /** BYROW/BYCOL(array, lambda(vector)) — one result per row or column. */
  BYROW: (a, host) => {
    if (a[0] === undefined) return "#VALUE!";
    const m = toMatrix(a[0]);
    return withLambda(a[1], host, (call) =>
      m.map((row) => [toScalar(call([[row]]))]),
    );
  },

  BYCOL: (a, host) => {
    if (a[0] === undefined) return "#VALUE!";
    const m = transpose(toMatrix(a[0]));
    return withLambda(a[1], host, (call) => [
      m.map((col) => toScalar(call([[col]]))),
    ]);
  },

  /** REDUCE(initial, array, lambda(accumulator, value)). */
  REDUCE: (a, host) => {
    if (a[1] === undefined) return "#VALUE!";
    const values = flatten([a[1]]);
    return withLambda(a[2], host, (call) => {
      let acc: FormulaArg = a[0];
      for (const v of values) acc = call([acc, v]);
      return acc;
    });
  },

  /** SCAN(initial, array, lambda) — REDUCE keeping every intermediate result. */
  SCAN: (a, host) => {
    if (a[1] === undefined) return "#VALUE!";
    const m = toMatrix(a[1]);
    return withLambda(a[2], host, (call) => {
      let acc: FormulaArg = a[0];
      return m.map((row) =>
        row.map((v) => {
          acc = call([acc, v]);
          return toScalar(acc);
        }),
      );
    });
  },

  /** Applies a lambda directly: `LAMBDA(x, x+1)(2)` is written LET-style in files. */
  ISOMITTED: (a) => a[0] === undefined,

  /** SUMPRODUCT(a, b, …) — the sum of elementwise products. */
  SUMPRODUCT: (a) => {
    const parts = a.map(toMatrix);
    const first = parts[0];
    if (!first) return 0;
    let total = 0;
    for (let r = 0; r < rowCount(first); r++) {
      for (let c = 0; c < colCount(first); c++) {
        let product = 1;
        for (const part of parts) product *= toNumber(at(part, r, c));
        total += product;
      }
    }
    return total;
  },

  /** ARRAYTOTEXT(array) — the values joined, which is what a debug view wants. */
  ARRAYTOTEXT: (a) =>
    a[0] === undefined
      ? ""
      : flatten([a[0]])
          .map((v) => toText(v))
          .join(", "),
};

export { isMatrix, transpose };
