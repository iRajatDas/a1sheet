/**
 * Logical functions. Ported from ref/formulaEngine.js:201-204.
 *
 * Note IF is NOT lazy: evalNode evaluates every argument before dispatching, so
 * both branches are computed. That matters only for cost, not correctness, since
 * there are no side effects — but it does mean `IF(A1<>0, B1/A1, 0)` still
 * divides by zero internally and returns Infinity for the unused branch.
 */
import { flattenBool, toNumber } from "../values.js";
import type { FormulaFunction } from "./registry.js";

export const logicFunctions: Record<string, FormulaFunction> = {
  IF: (a) => {
    const branch = toNumber(a[0]) ? a[1] : a[2];
    return Array.isArray(branch) ? branch[0] : branch;
  },

  AND: (a) => (flattenBool(a).every(Boolean) ? 1 : 0),
  OR: (a) => (flattenBool(a).some(Boolean) ? 1 : 0),
  NOT: (a) => (toNumber(a[0]) ? 0 : 1),
};
