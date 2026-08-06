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

  /**
   * The URL of an image to draw in the cell.
   *
   * Evaluates to the URL, because a formula's value has to be a scalar — the
   * picture is drawn from `Sheet.images`, which the reader and `setCell` keep in
   * step with the formula. So a cell whose image cannot be fetched still shows
   * where it was meant to come from instead of an error.
   *
   * Excel's further arguments (alt text, sizing, dimensions) are accepted and
   * ignored rather than rejected, so a file using them imports.
   */
  IMAGE: (a) => toText(a[0]),

  TRIM: (a) => toText(a[0]).trim(),
  UPPER: (a) => toText(a[0]).toUpperCase(),
  LOWER: (a) => toText(a[0]).toLowerCase(),
  LEN: (a) => toText(a[0]).length,
};
