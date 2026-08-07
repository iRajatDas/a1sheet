/**
 * The error sentinels, in their own module.
 *
 * Not in `evaluate.ts` because the function library needs them — `ISERROR` and
 * `IFERROR` are about nothing else — and a function module importing the
 * evaluator would close a cycle through the registry.
 */

/** Sentinel returned when a formula participates in a reference cycle. */
export const CYCLE_ERROR = "#CYCLE!";

/**
 * The complete set of error sentinels. Membership is checked exactly rather than
 * by a leading "#" so a user who literally types "#hashtag" is not mistaken for
 * an error value.
 */
export const ERROR_VALUES: ReadonlySet<string> = new Set([
  CYCLE_ERROR,
  "#ERROR!",
  "#DIV/0!",
  "#NAME?",
  "#REF!",
  "#N/A",
  "#VALUE!",
  "#NUM!",
  /** A dynamic array that produced nothing, or one that cannot fit. */
  "#CALC!",
  "#SPILL!",
]);

export function isErrorValue(v: unknown): v is string {
  return typeof v === "string" && ERROR_VALUES.has(v);
}
