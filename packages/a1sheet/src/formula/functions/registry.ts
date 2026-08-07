/**
 * The function library, assembled from one module per category.
 *
 * Adding a function is normally a one-liner in the relevant category module.
 *
 * ONE EXCEPTION, in `evaluate.ts` rather than here: the **lazy** forms. LET and
 * LAMBDA bind names, so their arguments cannot be evaluated before dispatch, and
 * IF, IFS, SWITCH, IFERROR, and IFNA must not evaluate the branch they do not
 * take.
 *
 * VLOOKUP, INDEX, and MATCH used to be a second exception — they needed the raw
 * range AST node to recover a table's 2D shape. A range evaluates to a matrix
 * now, so there is nothing to recover and they are ordinary functions that also
 * work on a computed table.
 */
import type { FormulaArg, LambdaValue } from "../values.js";
import { arrayFunctions } from "./array.js";
import { dateFunctions } from "./date.js";
import { logicFunctions } from "./logic.js";
import { lookupFunctions } from "./lookup.js";
import { mathFunctions } from "./math.js";
import { randomFunctions } from "./random.js";
import { textFunctions } from "./text.js";

/**
 * What a function can do beyond looking at its arguments.
 *
 * `call` exists for the functions that take a lambda — MAP, MAKEARRAY, BYROW,
 * REDUCE. Applying a lambda means evaluating its body in the scope it captured,
 * which is the evaluator's business, so it is handed in rather than reimplemented.
 */
export interface FormulaHost {
  call(lambda: LambdaValue, args: readonly FormulaArg[]): FormulaArg;
  /**
   * The instant this calculation cycle began, in ms since the Unix epoch.
   *
   * `TODAY` and `NOW` read it rather than the clock so that every one of them on
   * a sheet reports the same moment — two `=NOW()` cells calling `Date.now()`
   * separately can land either side of a millisecond, and in Excel they cannot.
   */
  now: number;
}

/**
 * Functions receive their arguments as a single array. A range argument arrives
 * as a `Matrix` — always two-dimensional, so `[[1, 2]]` is a row and
 * `[[1], [2]]` a column.
 */
export type FormulaFunction = (args: FormulaArg[], host: FormulaHost) => FormulaArg;

export const FUNCTIONS: Record<string, FormulaFunction> = {
  ...mathFunctions,
  ...randomFunctions,
  ...textFunctions,
  ...logicFunctions,
  ...dateFunctions,
  ...arrayFunctions,
  ...lookupFunctions,
};

/**
 * Registers a custom function. Uppercases the name to match the tokenizer,
 * which uppercases every identifier.
 */
export function registerFunction(name: string, fn: FormulaFunction): void {
  FUNCTIONS[name.toUpperCase()] = fn;
}
