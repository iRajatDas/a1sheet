/**
 * Logical functions.
 *
 * IF, IFS, SWITCH, IFERROR, and IFNA are NOT here: they must not evaluate the
 * branch they do not take, which means intercepting them before their arguments
 * are evaluated. See LAZY_FORMS in `evaluate.ts`.
 */
import { isErrorValue } from "../errorValues.js";
import { flattenBool, toNumber, toScalar } from "../values.js";
import type { FormulaFunction } from "./registry.js";

export const logicFunctions: Record<string, FormulaFunction> = {
  AND: (a) => flattenBool(a).every(Boolean),
  OR: (a) => flattenBool(a).some(Boolean),
  XOR: (a) => flattenBool(a).filter(Boolean).length % 2 === 1,
  NOT: (a) => !toNumber(a[0]),

  TRUE: () => true,
  FALSE: () => false,

  ISERROR: (a) => isErrorValue(toScalar(a[0])),
  ISNA: (a) => toScalar(a[0]) === "#N/A",
  ISNUMBER: (a) => typeof toScalar(a[0]) === "number",
  ISTEXT: (a) => typeof toScalar(a[0]) === "string",
  ISLOGICAL: (a) => typeof toScalar(a[0]) === "boolean",

  /** NA() is how a formula says "no value here" deliberately. */
  NA: () => "#N/A",
};
