/**
 * Date functions. Ported from ref/formulaEngine.js:216-221.
 *
 * Dates are day-serial numbers since the UNIX epoch, not Excel's 1899-12-30
 * epoch. Internally consistent and arithmetic-friendly, but a serial copied out
 * of a1sheet will not match Excel's serial for the same date. See
 * docs/LIMITATIONS.md.
 */
import { DAY_MS, toNumber } from "../values.js";
import type { FormulaFunction } from "./registry.js";

export const dateFunctions: Record<string, FormulaFunction> = {
  TODAY: () => Math.floor(Date.now() / DAY_MS),
  NOW: () => Date.now() / DAY_MS,

  /** Month is 1-indexed on the way in, matching Excel. */
  DATE: (a) =>
    Math.floor(
      Date.UTC(toNumber(a[0]), toNumber(a[1]) - 1, toNumber(a[2])) / DAY_MS,
    ),

  YEAR: (a) => new Date(toNumber(a[0]) * DAY_MS).getUTCFullYear(),
  MONTH: (a) => new Date(toNumber(a[0]) * DAY_MS).getUTCMonth() + 1,
  DAY: (a) => new Date(toNumber(a[0]) * DAY_MS).getUTCDate(),
};
