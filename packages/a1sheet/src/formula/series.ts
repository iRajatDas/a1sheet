/**
 * Fill-handle series extrapolation. Ported from ref/formulaEngine.js:384-396.
 *
 * Rules, in order:
 *   - all-numeric and 2+ values -> continue the last observed step linearly
 *   - all-numeric and 1 value   -> +1 per step
 *   - anything else             -> cyclic repeat of the source values
 *
 * Only called for non-formula fill sources. A formula source goes through
 * `shiftFormulaRefs` once per destination cell instead.
 */
import type { FormulaValue } from "./values.js";

export function extrapolateSeries(values: FormulaValue[], count: number): string[] {
  const nums = values.map((v) => parseFloat(v as string));
  const allNumeric =
    nums.every((n) => !Number.isNaN(n)) && values.every((v) => v !== "");
  const out: string[] = [];

  if (allNumeric && values.length >= 2) {
    const last = nums[nums.length - 1] as number;
    const step = last - (nums[nums.length - 2] as number);
    for (let i = 0; i < count; i++) out.push(String(last + step * (i + 1)));
  } else if (allNumeric && values.length === 1) {
    const first = nums[0] as number;
    for (let i = 0; i < count; i++) out.push(String(first + (i + 1)));
  } else {
    for (let i = 0; i < count; i++) {
      out.push(String(values[i % values.length] ?? ""));
    }
  }

  return out;
}
