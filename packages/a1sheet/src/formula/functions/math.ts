/** Math and statistics functions. Ported from ref/formulaEngine.js:193-200. */
import { flattenNums, toNumber } from "../values.js";
import type { FormulaFunction } from "./registry.js";

const sum = (n: number[]) => n.reduce((x, y) => x + y, 0);

export const mathFunctions: Record<string, FormulaFunction> = {
  SUM: (a) => sum(flattenNums(a)),

  AVERAGE: (a) => {
    const n = flattenNums(a);
    return n.length ? sum(n) / n.length : 0;
  },

  MIN: (a) => {
    const n = flattenNums(a);
    return n.length ? Math.min(...n) : 0;
  },

  MAX: (a) => {
    const n = flattenNums(a);
    return n.length ? Math.max(...n) : 0;
  },

  COUNT: (a) => flattenNums(a).length,

  /** Counts non-empty cells, numeric or not. */
  COUNTA: (a) => {
    let c = 0;
    for (const v of a) {
      if (Array.isArray(v)) {
        for (const x of v) if (x !== "" && x !== undefined) c++;
      } else if (v !== "" && v !== undefined) {
        c++;
      }
    }
    return c;
  },

  ABS: (a) => Math.abs(toNumber(a[0])),

  ROUND: (a) => {
    const n = toNumber(a[0]);
    const d = a[1] !== undefined ? toNumber(a[1]) : 0;
    const f = 10 ** d;
    return Math.round(n * f) / f;
  },
};
