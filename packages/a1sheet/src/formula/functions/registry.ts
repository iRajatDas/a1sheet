/**
 * The function library, assembled from one module per category.
 *
 * Adding a function is normally a one-liner in the relevant category module.
 *
 * EXCEPTION — shape-sensitive functions: VLOOKUP, INDEX, and MATCH need the raw
 * range AST node to recover 2D shape, which a flattened argument array has
 * already lost. They are special-cased inside evalNode's "call" branch and do
 * NOT live here. A new shape-sensitive function follows that pattern rather than
 * being registered in this table.
 */
import type { FormulaArg, FormulaValue } from "../values.js";
import { dateFunctions } from "./date.js";
import { logicFunctions } from "./logic.js";
import { mathFunctions } from "./math.js";
import { textFunctions } from "./text.js";

/**
 * Functions receive their arguments as a single array, matching the POC. Range
 * arguments arrive as nested arrays, which is why most implementations start by
 * flattening.
 */
export type FormulaFunction = (args: FormulaArg[]) => FormulaValue;

export const FUNCTIONS: Record<string, FormulaFunction> = {
  ...mathFunctions,
  ...textFunctions,
  ...logicFunctions,
  ...dateFunctions,
};

/** Names handled directly in evalNode because they need the range AST node. */
export const SHAPE_SENSITIVE = new Set(["VLOOKUP", "INDEX", "MATCH"]);

/**
 * Registers a custom function. Uppercases the name to match the tokenizer,
 * which uppercases every identifier.
 */
export function registerFunction(name: string, fn: FormulaFunction): void {
  FUNCTIONS[name.toUpperCase()] = fn;
}
