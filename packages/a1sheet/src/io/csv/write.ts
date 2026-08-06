/**
 * CSV writer. Ported from ref/xlsxIO.js:422-441.
 *
 * Writes EVALUATED display values, not raw formulas — a CSV export of `=SUM(A1:A3)`
 * contains the sum. That is the desired behavior for a data interchange format.
 */
import type { Evaluator } from "../../formula/evaluate.js";
import type { CellKey, RawCell } from "../../model/types.js";
import { type CsvInjectionMode, neutralizeCsvValue } from "./sanitize.js";

/** Extent of the non-empty region, used to bound export loops. */
export function getUsedBounds(cells: Record<CellKey, RawCell>): {
  rows: number;
  cols: number;
} {
  let maxR = -1;
  let maxC = -1;
  for (const key of Object.keys(cells) as CellKey[]) {
    const v = cells[key];
    if (v === "" || v == null) continue;
    const i = key.indexOf("_");
    const r = Number(key.slice(0, i));
    const c = Number(key.slice(i + 1));
    if (r > maxR) maxR = r;
    if (c > maxC) maxC = c;
  }
  return { rows: maxR + 1, cols: maxC + 1 };
}

export interface CellsToCsvOptions {
  /**
   * Formula-injection handling. Defaults to "prefix-quote": a value a spreadsheet
   * would execute as a formula is prefixed with the "treat as text" apostrophe.
   *
   * Set to "off" ONLY when the output is not opened by another person's
   * spreadsheet application. See ./sanitize.ts for why this matters.
   */
  injectionMode?: CsvInjectionMode;
}

export function cellsToCSV(
  cells: Record<CellKey, RawCell>,
  evaluator: Evaluator,
  options: CellsToCsvOptions = {},
): string {
  const b = getUsedBounds(cells);
  const lines: string[] = [];

  for (let r = 0; r < b.rows; r++) {
    const parts: string[] = [];
    for (let c = 0; c < b.cols; c++) {
      const raw = evaluator.getCellDisplay(r, c);
      let v = raw === undefined ? "" : String(raw);
      // Neutralize BEFORE quoting: the guard inspects the leading character, and
      // quoting would hide it.
      v = neutralizeCsvValue(v, { mode: options.injectionMode ?? "prefix-quote" });
      if (/[",\n\t]/.test(v)) v = `"${v.replace(/"/g, '""')}"`;
      parts.push(v);
    }
    lines.push(parts.join(","));
  }

  return lines.join("\r\n");
}
