/**
 * Coercion helpers shared by the evaluator and the function library.
 * Ported from ref/formulaEngine.js:163-187.
 *
 * These are intentionally forgiving: non-numeric text coerces to 0 rather than
 * producing an error value, matching the POC's behavior.
 */

/** A value flowing through evaluation. Arrays come from range arguments. */
export type FormulaValue = string | number | boolean | undefined;
export type FormulaArg = FormulaValue | FormulaValue[];

/** Day-serial epoch. Unix-based, NOT Excel's 1899-12-30 — see docs/LIMITATIONS.md. */
export const DAY_MS = 86400000;

export function toNumber(v: FormulaArg): number {
  if (typeof v === "number") return v;
  if (v === true) return 1;
  if (v === false) return 0;
  const n = parseFloat(v as string);
  return Number.isNaN(n) ? 0 : n;
}

export function toText(v: FormulaArg): string {
  if (Array.isArray(v)) return v.map((x) => toText(x)).join("");
  return v === undefined || v === null ? "" : String(v);
}

/** Flattens args to numbers, dropping anything non-numeric (so COUNT works). */
export function flattenNums(args: FormulaArg[]): number[] {
  const out: number[] = [];
  for (const a of args) {
    if (Array.isArray(a)) {
      for (const v of a) {
        const n = parseFloat(v as string);
        if (!Number.isNaN(n)) out.push(n);
      }
    } else {
      const n = parseFloat(a as string);
      if (!Number.isNaN(n)) out.push(n);
    }
  }
  return out;
}

/** Flattens args to booleans. Empty string is false; any other text is true. */
export function flattenBool(args: FormulaArg[]): boolean[] {
  const out: boolean[] = [];
  const push = (v: FormulaValue) =>
    out.push(typeof v === "string" ? v !== "" : !!toNumber(v));
  for (const a of args) {
    if (Array.isArray(a)) for (const v of a) push(v);
    else push(a);
  }
  return out;
}
