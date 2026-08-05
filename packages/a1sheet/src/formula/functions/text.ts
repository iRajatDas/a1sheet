/** Text functions. Ported from ref/formulaEngine.js:206-214. */
import { toNumber, toText } from "../values.js";
import type { FormulaFunction } from "./registry.js";

const concat: FormulaFunction = (a) => a.map((v) => toText(v)).join("");

export const textFunctions: Record<string, FormulaFunction> = {
  CONCAT: concat,
  CONCATENATE: concat,

  LEFT: (a) => toText(a[0]).slice(0, a[1] !== undefined ? toNumber(a[1]) : 1),

  RIGHT: (a) => {
    const s = toText(a[0]);
    const n = a[1] !== undefined ? toNumber(a[1]) : 1;
    return s.slice(Math.max(0, s.length - n));
  },

  /** 1-indexed start, matching Excel. */
  MID: (a) =>
    toText(a[0]).slice(toNumber(a[1]) - 1, toNumber(a[1]) - 1 + toNumber(a[2])),

  TRIM: (a) => toText(a[0]).trim(),
  UPPER: (a) => toText(a[0]).toUpperCase(),
  LOWER: (a) => toText(a[0]).toLowerCase(),
  LEN: (a) => toText(a[0]).length,
};
