/**
 * Plain-language explanations for the error sentinels a formula can produce.
 *
 * The sentinel in the cell says something is wrong; it does not say what to do
 * about it, and "#CYCLE!" in particular tells a user nothing. Kept out of the
 * evaluator because that runs per cell and must stay allocation-light, and out
 * of the React layer because it is not presentation — a Node consumer reporting
 * on a sheet wants the same sentences.
 */
import { CYCLE_ERROR } from "./evaluate.js";

const EXPLANATIONS: Readonly<Record<string, string>> = {
  [CYCLE_ERROR]:
    "Circular reference: this formula depends on its own result. " +
    "Point it at a different range, or move the total outside the range it sums.",
  "#DIV/0!": "Division by zero. Guard the divisor, e.g. =IF(B1=0, 0, A1/B1).",
  "#NAME?":
    "Unknown name: a function or named range that does not exist. " +
    "Check the spelling.",
  "#REF!": "Invalid reference: it points outside the sheet, or at deleted cells.",
  "#N/A": "No match found.",
  "#VALUE!": "Wrong kind of value: text where a number was expected.",
  "#ERROR!": "The formula could not be parsed. Check for unbalanced brackets.",
} as const;

/**
 * The explanation for an error value, or null when `value` is not one.
 *
 * Returning null rather than a fallback string is what lets a caller say
 * "if (explanation) show it" without first testing whether it is an error.
 */
export function explainErrorValue(value: unknown): string | null {
  if (typeof value !== "string") return null;
  return EXPLANATIONS[value] ?? null;
}
