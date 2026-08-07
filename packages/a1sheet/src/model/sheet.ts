/**
 * Sheet construction and structural edits.
 *
 * Ported from ref/Spreadsheet.jsx:30-62. Every function here is pure: it takes a
 * sheet and returns a new one. Nothing mutates in place — `useWorkbook` is the
 * only place clone-on-write is coordinated with history.
 */
import { cellKey } from "./address.js";
import type {
  CellImage,
  CellKey,
  CellValue,
  CondFormat,
  RawCell,
  Sheet,
  SheetTable,
  StyleObject,
} from "./types.js";

export const DEFAULT_NUM_ROWS = 200;
export const DEFAULT_NUM_COLS = 26;

export function uid(): string {
  return Math.random().toString(36).slice(2, 9);
}

export function makeSheet(name: string): Sheet {
  return {
    id: uid(),
    name,
    cells: {},
    styles: {},
    cachedValues: {},
    condFormats: [],
    images: {},
    tables: [],
    validations: [],
    spillRanges: {},
    colWidths: {},
    rowHeights: {},
    merges: [],
    frozenRows: 0,
    frozenCols: 0,
    hiddenRows: new Set(),
    hiddenCols: new Set(),
    colLabels: {},
    rowLabels: {},
    filters: {},
    numRows: DEFAULT_NUM_ROWS,
    numCols: DEFAULT_NUM_COLS,
  };
}

/**
 * Shallow-clones every mutable container on the sheet. `merges` entries are
 * cloned too since they are edited in place by resize operations.
 */
export function cloneSheet(sheet: Sheet): Sheet {
  return {
    ...sheet,
    cells: { ...sheet.cells },
    styles: { ...sheet.styles },
    cachedValues: { ...sheet.cachedValues },
    // Frozen values, replaced wholesale rather than edited, so the array itself
    // may be shared. Ranges inside it still move on insert and delete.
    condFormats: sheet.condFormats,
    images: { ...sheet.images },
    tables: sheet.tables,
    validations: sheet.validations,
    spillRanges: { ...sheet.spillRanges },
    colWidths: { ...sheet.colWidths },
    rowHeights: { ...sheet.rowHeights },
    merges: sheet.merges.map((m) => ({ ...m })),
    hiddenRows: new Set(sheet.hiddenRows),
    hiddenCols: new Set(sheet.hiddenCols),
    colLabels: { ...sheet.colLabels },
    rowLabels: { ...sheet.rowLabels },
    filters: { ...sheet.filters },
  };
}

/**
 * Rewrites `"r_c"` keys along one axis. Keys before `at` are untouched; on a
 * delete (`delta < 0`) keys exactly at `at` are dropped.
 *
 * The POC built the new key in one inlined template literal
 * (ref/Spreadsheet.jsx:48-58), where the row and column halves are trivial to
 * transpose. Split into two branches here, and the tests use off-diagonal keys
 * — `1_5`, `5_5` moved along one axis — so a transposition fails them instead
 * of landing on a key that happens to be its own mirror.
 */
function shiftKeys<T>(
  obj: Record<CellKey, T>,
  axis: "row" | "col",
  at: number,
  delta: number,
): Record<CellKey, T> {
  const next: Record<CellKey, T> = {};
  for (const key of Object.keys(obj) as CellKey[]) {
    const i = key.indexOf("_");
    const r = Number(key.slice(0, i));
    const c = Number(key.slice(i + 1));
    const v = axis === "row" ? r : c;
    const value = obj[key] as T;
    if (v < at) {
      next[key] = value;
      continue;
    }
    // The deleted line itself: its keys go nowhere.
    if (delta < 0 && v === at) continue;
    if (axis === "row") next[`${v + delta}_${c}`] = value;
    else next[`${r}_${v + delta}`] = value;
  }
  return next;
}

/**
 * The same shift for a map keyed by a bare row or column index — heights,
 * widths, labels, filters. `shiftKeys` handles the two-dimensional cell maps;
 * these are the one-dimensional ones alongside them, and they move for exactly
 * the same reason: inserting a row above a resized row must carry the height
 * with the row, not leave it behind on the index.
 */
function shiftIndexMap<T>(
  map: Record<number, T>,
  at: number,
  delta: number,
): Record<number, T> {
  const next: Record<number, T> = {};
  for (const key of Object.keys(map)) {
    const index = Number(key);
    const value = map[index] as T;
    if (index < at) next[index] = value;
    else if (delta < 0 && index === at) continue;
    else next[index + delta] = value;
  }
  return next;
}

function shiftIndexSet(set: Set<number>, at: number, delta: number): Set<number> {
  const next = new Set<number>();
  for (const index of set) {
    if (index < at) next.add(index);
    else if (delta < 0 && index === at) continue;
    else next.add(index + delta);
  }
  return next;
}

/**
 * Moves a conditional format's range with the rows or columns it covers.
 *
 * A range whose start is past the insertion point slides whole; one that straddles
 * it grows or shrinks. Without this, inserting a row above a rule leaves the rule
 * a row behind, colouring the wrong cells — the same bug the row and column
 * metadata had.
 */
function shiftCondFormats(
  formats: readonly CondFormat[],
  axis: "row" | "col",
  at: number,
  delta: number,
): readonly CondFormat[] {
  if (formats.length === 0) return formats;
  const startKey = axis === "row" ? "r1" : "c1";
  const endKey = axis === "row" ? "r2" : "c2";
  return formats.map((format) => {
    const start = format.range[startKey];
    const end = format.range[endKey];
    if (end < at) return format;
    return {
      ...format,
      range: {
        ...format.range,
        [startKey]: start >= at ? start + delta : start,
        [endKey]: end + delta,
      },
    };
  });
}

/**
 * Moves a table's range with the rows or columns it covers, on the same terms as
 * a conditional format. A table left behind would resolve its column names
 * against the wrong cells, which is worse than a wrong colour.
 */
function shiftTables(
  tables: readonly SheetTable[],
  axis: "row" | "col",
  at: number,
  delta: number,
): readonly SheetTable[] {
  if (tables.length === 0) return tables;
  const startKey = axis === "row" ? "r1" : "c1";
  const endKey = axis === "row" ? "r2" : "c2";
  return tables.map((table) => {
    const start = table.range[startKey];
    const end = table.range[endKey];
    if (end < at) return table;
    return {
      ...table,
      range: {
        ...table.range,
        [startKey]: start >= at ? start + delta : start,
        [endKey]: end + delta,
      },
    };
  });
}

function shiftRows(sheet: Sheet, at: number, delta: number): Sheet {
  return {
    ...sheet,
    cells: shiftKeys(sheet.cells, "row", at, delta),
    styles: shiftKeys(sheet.styles, "row", at, delta),
    cachedValues: shiftKeys(sheet.cachedValues, "row", at, delta),
    condFormats: shiftCondFormats(sheet.condFormats, "row", at, delta),
    images: shiftKeys(sheet.images, "row", at, delta),
    tables: shiftTables(sheet.tables, "row", at, delta),
    spillRanges: shiftKeys(sheet.spillRanges, "row", at, delta),
    rowHeights: shiftIndexMap(sheet.rowHeights, at, delta),
    rowLabels: shiftIndexMap(sheet.rowLabels, at, delta),
    hiddenRows: shiftIndexSet(sheet.hiddenRows, at, delta),
  };
}

function shiftCols(sheet: Sheet, at: number, delta: number): Sheet {
  return {
    ...sheet,
    cells: shiftKeys(sheet.cells, "col", at, delta),
    styles: shiftKeys(sheet.styles, "col", at, delta),
    cachedValues: shiftKeys(sheet.cachedValues, "col", at, delta),
    condFormats: shiftCondFormats(sheet.condFormats, "col", at, delta),
    images: shiftKeys(sheet.images, "col", at, delta),
    tables: shiftTables(sheet.tables, "col", at, delta),
    spillRanges: shiftKeys(sheet.spillRanges, "col", at, delta),
    colWidths: shiftIndexMap(sheet.colWidths, at, delta),
    colLabels: shiftIndexMap(sheet.colLabels, at, delta),
    hiddenCols: shiftIndexSet(sheet.hiddenCols, at, delta),
    filters: shiftIndexMap(sheet.filters, at, delta),
  };
}

export function insertRow(sheet: Sheet, at: number): Sheet {
  return { ...shiftRows(sheet, at, 1), numRows: sheet.numRows + 1 };
}

export function deleteRow(sheet: Sheet, at: number): Sheet {
  return { ...shiftRows(sheet, at, -1), numRows: Math.max(1, sheet.numRows - 1) };
}

export function insertCol(sheet: Sheet, at: number): Sheet {
  return { ...shiftCols(sheet, at, 1), numCols: sheet.numCols + 1 };
}

export function deleteCol(sheet: Sheet, at: number): Sheet {
  return { ...shiftCols(sheet, at, -1), numCols: Math.max(1, sheet.numCols - 1) };
}

/**
 * Physically rewrites `cells` and `styles` keys to reorder rows by one column.
 *
 * This is a DATA operation, not a view-layer sort — unlike `filters`, clearing it
 * does not restore the original order. Undo goes through history.
 *
 * Sorts the used range only (rows 0..maxUsedRow), leaving trailing empty rows
 * alone. Numeric values sort numerically and before text; blanks sort last in
 * both directions, which is what spreadsheet users expect.
 */
export function sortByColumn(
  sheet: Sheet,
  col: number,
  dir: "asc" | "desc",
): Sheet {
  let maxRow = -1;
  for (const key of Object.keys(sheet.cells) as CellKey[]) {
    if (sheet.cells[key] === "") continue;
    const r = Number(key.slice(0, key.indexOf("_")));
    if (r > maxRow) maxRow = r;
  }
  if (maxRow < 1) return sheet;

  const order = Array.from({ length: maxRow + 1 }, (_, r) => r);
  const valueAt = (r: number) => sheet.cells[cellKey(r, col)] ?? "";

  const sign = dir === "asc" ? 1 : -1;
  order.sort((ra, rb) => {
    const a = valueAt(ra);
    const b = valueAt(rb);
    // Blanks always sink, regardless of direction.
    if (a === "" && b === "") return ra - rb;
    if (a === "") return 1;
    if (b === "") return -1;
    const an = parseFloat(a);
    const bn = parseFloat(b);
    const aNum = !Number.isNaN(an);
    const bNum = !Number.isNaN(bn);
    if (aNum && bNum) return sign * (an - bn);
    if (aNum !== bNum) return sign * (aNum ? -1 : 1);
    return sign * a.localeCompare(b);
  });

  // order[newRow] = oldRow. Rebuild every cell-keyed map from it in one pass.
  // cachedValues moves with the cells for the same reason styles do: an imported
  // value left behind on its old row would be attributed to a different formula.
  const cells: Record<CellKey, RawCell> = {};
  const styles: Record<CellKey, StyleObject> = {};
  const cachedValues: Record<CellKey, CellValue> = {};
  const images: Record<CellKey, CellImage> = {};
  for (const [newRow, oldRow] of order.entries()) {
    for (let c = 0; c < sheet.numCols; c++) {
      const from = cellKey(oldRow, c);
      const to = cellKey(newRow, c);
      const cell = sheet.cells[from];
      if (cell !== undefined) cells[to] = cell;
      const style = sheet.styles[from];
      if (style !== undefined) styles[to] = style;
      const cached = sheet.cachedValues[from];
      if (cached !== undefined) cachedValues[to] = cached;
      const image = sheet.images[from];
      if (image !== undefined) images[to] = image;
    }
  }
  // Preserve anything below the sorted range untouched.
  const keepBelow = <T>(from: Record<CellKey, T>, into: Record<CellKey, T>) => {
    for (const key of Object.keys(from) as CellKey[]) {
      if (Number(key.slice(0, key.indexOf("_"))) > maxRow) {
        into[key] = from[key] as T;
      }
    }
  };
  keepBelow(sheet.cells, cells);
  keepBelow(sheet.styles, styles);
  keepBelow(sheet.cachedValues, cachedValues);
  keepBelow(sheet.images, images);

  return { ...sheet, cells, styles, cachedValues, images };
}

/** Reads a cell's style, or undefined when the cell has no explicit formatting. */
export function getStyle(sheet: Sheet, key: CellKey): StyleObject | undefined {
  return sheet.styles[key];
}

/**
 * Finds the merge covering a cell, if any.
 *
 * Linear scan per call. Fine while sheets have few merges, which is the norm; a
 * sheet with hundreds would want an index. Called once per rendered cell.
 */
export function getMergeAt(sheet: Sheet, row: number, col: number) {
  return sheet.merges.find(
    (m) => row >= m.r1 && row <= m.r2 && col >= m.c1 && col <= m.c2,
  );
}
