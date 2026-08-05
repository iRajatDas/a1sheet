/**
 * Token and AST shapes for the formula language.
 *
 * `RefToken` carries the `$` absolute markers separately from the ref text
 * because `shiftFormulaRefs` needs them to decide which refs move during a fill
 * or an internal paste. The AST's `ref` and `range` nodes deliberately drop them
 * — evaluation does not care whether a ref was anchored.
 */

export interface RefToken {
  type: "ref";
  /** Ref text with `$` stripped, uppercased, e.g. "A1". */
  value: string;
  colAbs: boolean;
  rowAbs: boolean;
}

export type Token =
  | RefToken
  | { type: "num"; value: number }
  | { type: "str"; value: string }
  | { type: "name"; value: string }
  | { type: "cmp"; value: CompareOp }
  | { type: "+" | "-" | "*" | "/" | "^" | "(" | ")" | "," | ":" };

export type CompareOp = "=" | "<" | ">" | "<=" | ">=" | "<>";
export type BinaryOp = "+" | "-" | "*" | "/" | "^";

export type Node =
  | { type: "num"; value: number }
  | { type: "str"; value: string }
  | { type: "ref"; value: string }
  | { type: "range"; from: string; to: string }
  | { type: "name"; value: string }
  | { type: "neg"; node: Node }
  | { type: "bin"; op: BinaryOp; left: Node; right: Node }
  | { type: "cmp"; op: CompareOp; left: Node; right: Node }
  | { type: "call"; name: string; args: Node[] };
