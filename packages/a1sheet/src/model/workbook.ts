/**
 * Workbook-level construction and sheet CRUD. Pure functions only.
 */
import { cloneSheet, makeSheet, uid } from "./sheet.js";
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

/**
 * Deep-copies a sheet (including filters, styles, merges) and inserts the copy
 * immediately after the source, activating it.
 */
export function duplicateSheet(wb: Workbook, index: number): Workbook {
  const source = wb.sheets[index];
  if (!source) return wb;
  const copy = cloneSheet(source);
  copy.id = uid();
  const taken = new Set(wb.sheets.map((s) => s.name));
  let name = `${source.name} (copy)`;
  let n = 2;
  while (taken.has(name)) name = `${source.name} (copy ${n++})`;
  copy.name = name;
  const sheets = [
    ...wb.sheets.slice(0, index + 1),
    copy,
    ...wb.sheets.slice(index + 1),
  ];
  return { ...wb, sheets, activeSheetIndex: index + 1 };
}

/** Reorders a sheet tab. Out-of-range indices are no-ops. */
export function moveSheet(wb: Workbook, from: number, to: number): Workbook {
  if (
    from < 0 ||
    to < 0 ||
    from >= wb.sheets.length ||
    to >= wb.sheets.length ||
    from === to
  ) {
    return wb;
  }
  const sheets = [...wb.sheets];
  const [moved] = sheets.splice(from, 1);
  if (!moved) return wb;
  sheets.splice(to, 0, moved);
  let active = wb.activeSheetIndex;
  if (active === from) active = to;
  else if (from < active && to >= active) active -= 1;
  else if (from > active && to <= active) active += 1;
  return { ...wb, sheets, activeSheetIndex: active };
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

export function namedRangeAddedStatus(name: string): string {
  return `Named range ${name.toUpperCase()} added`;
}
