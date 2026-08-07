/**
 * Workbook-level construction and sheet CRUD. Pure functions only.
 */
import { makeSheet } from "./sheet.js";
import type { Range, Sheet, Workbook } from "./types.js";

export function createWorkbook(sheetNames: string[] = ["Sheet1"]): Workbook {
  const names = sheetNames.length > 0 ? sheetNames : ["Sheet1"];
  return {
    sheets: names.map((n) => makeSheet(n)),
    activeSheetIndex: 0,
    namedRanges: {},
  };
}

export function activeSheet(wb: Workbook): Sheet {
  const s = wb.sheets[wb.activeSheetIndex] ?? wb.sheets[0];
  if (!s) throw new Error("workbook has no sheets");
  return s;
}

/** Appends a sheet with a name that does not collide with an existing one. */
export function addSheet(wb: Workbook, name?: string): Workbook {
  const taken = new Set(wb.sheets.map((s) => s.name));
  let n = name ?? `Sheet${wb.sheets.length + 1}`;
  let i = wb.sheets.length + 1;
  while (taken.has(n)) n = `Sheet${++i}`;
  return {
    ...wb,
    sheets: [...wb.sheets, makeSheet(n)],
    activeSheetIndex: wb.sheets.length,
  };
}

/** Deleting the last remaining sheet is a no-op. */
export function deleteSheet(wb: Workbook, index: number): Workbook {
  if (wb.sheets.length <= 1) return wb;
  const sheets = wb.sheets.filter((_, i) => i !== index);
  return {
    ...wb,
    sheets,
    activeSheetIndex: Math.min(wb.activeSheetIndex, sheets.length - 1),
  };
}

export function renameSheet(wb: Workbook, index: number, name: string): Workbook {
  return {
    ...wb,
    sheets: wb.sheets.map((s, i) => (i === index ? { ...s, name } : s)),
  };
}

/**
 * Defines a named range. Names are uppercased; there is no format validation
 * beyond that.
 */
export function defineName(wb: Workbook, name: string, range: Range): Workbook {
  return {
    ...wb,
    namedRanges: { ...wb.namedRanges, [name.toUpperCase()]: range },
  };
}

export function deleteName(wb: Workbook, name: string): Workbook {
  const next = { ...wb.namedRanges };
  delete next[name.toUpperCase()];
  return { ...wb, namedRanges: next };
}
