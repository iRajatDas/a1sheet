/** Math and statistics functions. Ported from ref/formulaEngine.js:193-200. */
import { flatten, flattenNums, toNumber } from "../values.js";
import type { FormulaFunction } from "./registry.js";

const sum = (n: number[]) => n.reduce((x, y) => x + y, 0);

export const mathFunctions: Record<string, FormulaFunction> = {
  SUM: (a) => sum(flattenNums(a)),

  PRODUCT: (a) => flattenNums(a).reduce((x, y) => x * y, 1),

  MEDIAN: (a) => {
    const n = flattenNums(a).sort((x, y) => x - y);
    if (n.length === 0) return 0;
    const mid = Math.floor(n.length / 2);
    return n.length % 2
      ? (n[mid] as number)
      : ((n[mid - 1] as number) + (n[mid] as number)) / 2;
  },

  /** Sample standard deviation, which is Excel's STDEV. */
  STDEV: (a) => {
    const n = flattenNums(a);
    if (n.length < 2) return "#DIV/0!";
    const mean = sum(n) / n.length;
    const variance = sum(n.map((x) => (x - mean) ** 2)) / (n.length - 1);
    return Math.sqrt(variance);
  },

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
  COUNTA: (a) => flatten(a).filter((v) => v !== "" && v !== undefined).length,

  ABS: (a) => Math.abs(toNumber(a[0])),
  SIGN: (a) => Math.sign(toNumber(a[0])),
  SQRT: (a) => {
    const n = toNumber(a[0]);
    return n < 0 ? "#NUM!" : Math.sqrt(n);
  },
  POWER: (a) => toNumber(a[0]) ** toNumber(a[1]),
  MOD: (a) => {
    const divisor = toNumber(a[1]);
    if (divisor === 0) return "#DIV/0!";
    // Excel's MOD takes the sign of the divisor, unlike JS's %.
    const n = toNumber(a[0]);
    return n - divisor * Math.floor(n / divisor);
  },
  INT: (a) => Math.floor(toNumber(a[0])),
  TRUNC: (a) => Math.trunc(toNumber(a[0])),
  CEILING: (a) => {
    const step = a[1] === undefined ? 1 : toNumber(a[1]);
    return step === 0 ? 0 : Math.ceil(toNumber(a[0]) / step) * step;
  },
  FLOOR: (a) => {
    const step = a[1] === undefined ? 1 : toNumber(a[1]);
    return step === 0 ? 0 : Math.floor(toNumber(a[0]) / step) * step;
  },
  EXP: (a) => Math.exp(toNumber(a[0])),
  LN: (a) => {
    const n = toNumber(a[0]);
    return n <= 0 ? "#NUM!" : Math.log(n);
  },
  LOG10: (a) => {
    const n = toNumber(a[0]);
    return n <= 0 ? "#NUM!" : Math.log10(n);
  },
  PI: () => Math.PI,

  ROUND: (a) => {
    const n = toNumber(a[0]);
    const d = a[1] !== undefined ? toNumber(a[1]) : 0;
    const f = 10 ** d;
    return Math.round(n * f) / f;
  },
};
