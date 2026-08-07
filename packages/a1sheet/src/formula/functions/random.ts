/**
 * The random functions, which are the engine's only genuinely volatile ones.
 *
 * "Volatile" means the result is not a function of the cells the formula reads,
 * so nothing can decide from the sheet whether it needs recomputing. Excel's
 * answer is to recompute every volatile on every calculation cycle, and that is
 * what happens here — an evaluator is built per cycle and memoizes within it, so
 * `=RAND()` in A1 and `=A1*2` in B1 see the same number, and both change
 * together on the next cycle. A cycle is any edit, or an explicit recalculation.
 *
 * `TODAY` and `NOW` are volatile in the same sense and live in `date.ts`, with
 * the rest of the calendar.
 */
import { toNumber } from "../values.js";
import type { FormulaFunction } from "./registry.js";

/** Excel's error for arguments that describe no possible result. */
const NUM_ERROR = "#NUM!";

export const randomFunctions: Record<string, FormulaFunction> = {
  /** Uniform in [0, 1). Excel's RAND takes no arguments; extras are ignored. */
  RAND: () => Math.random(),

  /**
   * A uniform integer in [bottom, top], both inclusive.
   *
   * Non-integer bounds move inward — bottom up, top down — so the result is
   * always inside the interval the caller named. An interval containing no
   * integer at all is `#NUM!`, the same answer as `bottom > top`.
   */
  RANDBETWEEN: (a) => {
    const bottom = Math.ceil(toNumber(a[0]));
    const top = Math.floor(toNumber(a[1]));
    if (!Number.isFinite(bottom) || !Number.isFinite(top)) return NUM_ERROR;
    if (bottom > top) return NUM_ERROR;
    return bottom + Math.floor(Math.random() * (top - bottom + 1));
  },
};
