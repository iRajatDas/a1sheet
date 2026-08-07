/**
 * Elementwise operations over matrices — Excel's array arithmetic.
 *
 * `(teams = "Arsenal") * played` is not a scalar expression: both sides are
 * columns, and the result is a column of products. Without this, an operator
 * silently used the top-left of each side and returned one number where the
 * formula meant several, which is worse than an error because it looks like an
 * answer.
 *
 * Broadcasting follows Excel: a 1x1 stretches to anything, a single row stretches
 * down, a single column stretches across, and any other mismatch pads with
 * `#N/A` rather than failing the whole expression.
 */
import {
  colCount,
  type FormulaArg,
  type FormulaValue,
  isMatrix,
  type Matrix,
  rowCount,
  toMatrix,
} from "./values.js";

/** Excel's answer for an element outside a mismatched operand. */
const NOT_AVAILABLE = "#N/A";

interface Shape {
  rows: number;
  cols: number;
}

function shapeOf(v: FormulaArg): Shape {
  if (!isMatrix(v)) return { rows: 1, cols: 1 };
  return { rows: rowCount(v), cols: colCount(v) };
}

/**
 * Reads the element of `m` that lines up with (row, col) of the result.
 *
 * A dimension of extent 1 repeats; any other index past the end is out of the
 * operand, which is `#N/A`.
 */
function element(m: Matrix, shape: Shape, row: number, col: number): FormulaValue {
  const r = shape.rows === 1 ? 0 : row;
  const c = shape.cols === 1 ? 0 : col;
  if (r >= shape.rows || c >= shape.cols) return NOT_AVAILABLE;
  return m[r]?.[c];
}

/**
 * Applies a scalar operation across one or two operands.
 *
 * Returns a scalar when every operand is scalar, so the common case costs
 * nothing and a formula that was not written as an array formula does not
 * suddenly produce one.
 */
export function broadcast2(
  left: FormulaArg,
  right: FormulaArg,
  op: (a: FormulaValue, b: FormulaValue) => FormulaValue,
): FormulaArg {
  const leftIsMatrix = isMatrix(left);
  const rightIsMatrix = isMatrix(right);
  if (!leftIsMatrix && !rightIsMatrix) {
    return op(left as FormulaValue, right as FormulaValue);
  }

  const lm = toMatrix(left);
  const rm = toMatrix(right);
  const ls = shapeOf(left);
  const rs = shapeOf(right);
  const rows = Math.max(ls.rows, rs.rows);
  const cols = Math.max(ls.cols, rs.cols);

  const out: FormulaValue[][] = [];
  for (let r = 0; r < rows; r++) {
    const row: FormulaValue[] = [];
    for (let c = 0; c < cols; c++) {
      row.push(op(element(lm, ls, r, c), element(rm, rs, r, c)));
    }
    out.push(row);
  }
  return out;
}

/** The same for a single operand, e.g. unary minus. */
export function broadcast1(
  value: FormulaArg,
  op: (a: FormulaValue) => FormulaValue,
): FormulaArg {
  if (!isMatrix(value)) return op(value as FormulaValue);
  return value.map((row) => row.map((v) => op(v)));
}
