/**
 * The values that flow through evaluation, and the coercions between them.
 *
 * Three kinds: scalars, matrices, and lambdas.
 *
 * A **matrix** is how every array value is represented, and it is always two
 * dimensional — Excel has no one-dimensional array, so a single row is `[[a, b]]`
 * and a single column is `[[a], [b]]`. Collapsing those to a flat list is what
 * makes `TRANSPOSE`, `HSTACK`, and `INDEX` impossible to express.
 *
 * A **lambda** is a function value, which only exists so `MAP` and friends have
 * something to be given. It is not a displayable value: a cell that evaluates to
 * one shows an error rather than a stringified function.
 *
 * The coercions are intentionally forgiving — non-numeric text becomes 0 rather
 * than an error — which is what keeps one bad cell from poisoning a column of
 * sums.
 */
import type { Node } from "./ast.js";

/** A single value. `undefined` is an empty cell. */
export type FormulaValue = string | number | boolean | undefined;

/** A rectangular array value, row-major. Always 2D; see the module comment. */
export type Matrix = readonly (readonly FormulaValue[])[];

/** Bound names visible to a formula body, innermost first via prototype chain. */
export type Scope = ReadonlyMap<string, FormulaArg>;

/**
 * A function value from `LAMBDA`. Carries the scope it was defined in, so a
 * lambda passed out of a `LET` can still see that `LET`'s bindings.
 */
export interface LambdaValue {
  readonly kind: "lambda";
  readonly params: readonly string[];
  readonly body: Node;
  readonly scope: Scope;
}

export type FormulaArg = FormulaValue | Matrix | LambdaValue;

export function isMatrix(v: FormulaArg): v is Matrix {
  return Array.isArray(v);
}

export function isLambda(v: FormulaArg): v is LambdaValue {
  return typeof v === "object" && v !== null && !Array.isArray(v);
}

/** Wraps a scalar as a 1x1 matrix. */
export function toMatrix(v: FormulaArg): Matrix {
  if (isMatrix(v)) return v;
  return [[isLambda(v) ? "#VALUE!" : v]];
}

/**
 * The single value a matrix stands for in a scalar position: its top-left.
 *
 * Excel's implicit intersection. An empty matrix has no value at all, which is
 * `#VALUE!` rather than an empty cell — the difference between "this range is
 * blank" and "this operation produced nothing".
 */
export function toScalar(v: FormulaArg): FormulaValue {
  if (isLambda(v)) return "#VALUE!";
  if (!isMatrix(v)) return v;
  const first = v[0];
  if (!first || first.length === 0) return "#VALUE!";
  return first[0];
}

export function rowCount(m: Matrix): number {
  return m.length;
}

/** Widest row, so a ragged matrix still reports a rectangular shape. */
export function colCount(m: Matrix): number {
  let width = 0;
  for (const row of m) if (row.length > width) width = row.length;
  return width;
}

/** Reads a cell of a matrix, or undefined outside it. */
export function at(m: Matrix, row: number, col: number): FormulaValue {
  return m[row]?.[col];
}

export function toNumber(v: FormulaArg): number {
  const s = toScalar(v);
  if (typeof s === "number") return s;
  if (s === true) return 1;
  if (s === false) return 0;
  const n = parseFloat(s as string);
  return Number.isNaN(n) ? 0 : n;
}

export function toText(v: FormulaArg): string {
  if (isMatrix(v)) return v.flat().map(scalarText).join("");
  return scalarText(toScalar(v));
}

function scalarText(v: FormulaValue): string {
  return v === undefined || v === null ? "" : String(v);
}

/** Every scalar in an argument list, matrices flattened row-major. */
export function flatten(args: readonly FormulaArg[]): FormulaValue[] {
  const out: FormulaValue[] = [];
  for (const a of args) {
    if (isMatrix(a)) out.push(...a.flat());
    else if (!isLambda(a)) out.push(a);
  }
  return out;
}

/** Flattens args to numbers, dropping anything non-numeric (so COUNT works). */
export function flattenNums(args: readonly FormulaArg[]): number[] {
  const out: number[] = [];
  for (const v of flatten(args)) {
    const n = typeof v === "number" ? v : parseFloat(v as string);
    if (!Number.isNaN(n) && n !== undefined) out.push(n);
  }
  return out;
}

/** Flattens args to booleans. Empty string is false; any other text is true. */
export function flattenBool(args: readonly FormulaArg[]): boolean[] {
  return flatten(args).map((v) =>
    typeof v === "string" ? v !== "" : !!toNumber(v),
  );
}
