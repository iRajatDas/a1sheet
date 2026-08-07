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
  /** Sheet name from a qualified ref, `Sheet2!A1`. Absent means this sheet. */
  sheet?: string;
}

/**
 * A structured reference: a table name plus the bracketed selector after it.
 *
 * Lexed whole rather than as `name` `[` … `]` because the selector's own grammar
 * has nothing to do with the expression grammar — `tbl[[#This Row],[a]:[b]]`
 * contains commas and colons that are not argument separators or range operators.
 */
export interface TableRefToken {
  type: "tableRef";
  table: string;
  /** Everything inside the outermost brackets, verbatim. */
  spec: string;
}

export type Token =
  | RefToken
  | TableRefToken
  | { type: "num"; value: number }
  | { type: "str"; value: string }
  | { type: "name"; value: string }
  | { type: "cmp"; value: CompareOp }
  | { type: "arr"; rows: (number | string | boolean)[][] }
  | { type: "+" | "-" | "*" | "/" | "^" | "&" | "%" | "(" | ")" | "," | ":" };

export type CompareOp = "=" | "<" | ">" | "<=" | ">=" | "<>";
export type BinaryOp = "+" | "-" | "*" | "/" | "^" | "&";

export type Node =
  | { type: "num"; value: number }
  | { type: "str"; value: string }
  | { type: "ref"; value: string; sheet?: string }
  | { type: "range"; from: string; to: string; sheet?: string }
  | { type: "name"; value: string }
  | { type: "tableRef"; table: string; spec: string }
  /** An array literal, `{1,2;3,4}`. Rows separated by `;`, cells by `,`. */
  | { type: "arr"; rows: (number | string | boolean)[][] }
  | { type: "neg"; node: Node }
  | { type: "bin"; op: BinaryOp; left: Node; right: Node }
  | { type: "cmp"; op: CompareOp; left: Node; right: Node }
  | { type: "call"; name: string; args: Node[] };
