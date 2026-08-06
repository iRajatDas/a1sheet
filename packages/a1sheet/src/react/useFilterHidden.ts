"use client";

/**
 * Which rows an active column filter excludes.
 *
 * A row is filter-hidden when any filtered column's allowed-value set does not
 * contain that row's displayed value for the column. That is a whole-sheet
 * question, and an edit can change the answer for the row it touched, so the
 * naive implementation re-evaluates every row on every keystroke — 97 ms at
 * 100k rows, which is a visible stall on each committed edit.
 *
 * So this hook caches the verdict and re-tests only the rows whose raw content
 * in a filtered column actually changed. Detecting those is two object lookups
 * per row against the previous `cells` map, where re-testing is an evaluator
 * call plus number formatting: about twenty times cheaper.
 *
 * The cache is only sound while a displayed value cannot change without its raw
 * text changing, which fails for a formula — `=B1*2` reads the same after B1
 * moves. A filtered column containing one is marked volatile and always fully
 * rescanned.
 *
 * A display also depends on the cell's number format, so `styles` is compared
 * per key alongside `cells`. Comparing whole maps by identity does NOT work
 * here: `cloneSheet` shallow-copies `cells`, `styles`, and `filters` on every
 * write, so all three are new objects after any edit. Their *entries* keep
 * their identity, which is what makes the per-key comparison both correct and
 * cheap.
 */
import { useRef } from "react";
import { cellKey } from "../model/address.js";
import type { Sheet } from "../model/types.js";

const NONE: ReadonlySet<number> = new Set();

interface FilterCache {
  cells: Sheet["cells"];
  styles: Sheet["styles"];
  filters: Sheet["filters"];
  numRows: number;
  frozenRows: number;
  hidden: ReadonlySet<number>;
  /** A filtered column holds a formula, so raw text no longer implies display. */
  volatile: boolean;
}

interface RowVerdict {
  hidden: boolean;
  formula: boolean;
}

/**
 * Same columns filtered, each by the same set. The `filters` map is a fresh
 * object after every write, but an untouched column's `Set` is the same one.
 */
function sameFilters(
  previous: Sheet["filters"],
  next: Sheet["filters"],
  cols: readonly number[],
): boolean {
  if (Object.keys(previous).length !== cols.length) return false;
  for (const col of cols) {
    if (previous[col] !== next[col]) return false;
  }
  return true;
}

export function useFilterHidden(
  sheet: Sheet,
  frozenRows: number,
  getDisplay: (row: number, col: number) => string,
): ReadonlySet<number> {
  const cacheRef = useRef<FilterCache | null>(null);
  const cache = cacheRef.current;

  const filterCols: number[] = [];
  for (const key of Object.keys(sheet.filters)) filterCols.push(Number(key));

  if (filterCols.length === 0) {
    cacheRef.current = null;
    return NONE;
  }

  const sameShape =
    cache !== null &&
    cache.numRows === sheet.numRows &&
    cache.frozenRows === frozenRows &&
    sameFilters(cache.filters, sheet.filters, filterCols);

  if (sameShape && cache.cells === sheet.cells && cache.styles === sheet.styles) {
    return cache.hidden;
  }

  /**
   * A cell with no content and no style displays the same string in every row
   * of its column, because nothing in that display depends on the row. On a
   * sparse sheet almost every row takes this path, so resolving it once per
   * column rather than once per row is most of the scan.
   */
  const emptyDisplay = new Map<number, string>();

  const test = (row: number): RowVerdict => {
    for (const col of filterCols) {
      const allowed = sheet.filters[col];
      if (!allowed) continue;
      const key = cellKey(row, col);
      const raw = sheet.cells[key];
      let display: string;
      if (raw === undefined && sheet.styles[key] === undefined) {
        let cached = emptyDisplay.get(col);
        if (cached === undefined) {
          cached = getDisplay(row, col);
          emptyDisplay.set(col, cached);
        }
        display = cached;
      } else {
        display = getDisplay(row, col);
      }
      if (raw?.startsWith("=")) {
        return { hidden: !allowed.has(display), formula: true };
      }
      if (!allowed.has(display)) return { hidden: true, formula: false };
    }
    return { hidden: false, formula: false };
  };

  if (sameShape && !cache.volatile) {
    const previousCells = cache.cells;
    const previousStyles = cache.styles;
    // Allocated only once something actually moves: an edit outside every
    // filtered column must not invalidate the set for its consumers.
    let hidden: Set<number> | null = null;
    let volatile = false;

    for (let row = frozenRows; row < sheet.numRows; row++) {
      let changed = false;
      for (const col of filterCols) {
        const key = cellKey(row, col);
        if (
          previousCells[key] !== sheet.cells[key] ||
          previousStyles[key] !== sheet.styles[key]
        ) {
          changed = true;
          break;
        }
      }
      if (!changed) continue;

      const verdict = test(row);
      if (verdict.formula) {
        volatile = true;
        break;
      }
      if (verdict.hidden === cache.hidden.has(row)) continue;
      if (hidden === null) hidden = new Set(cache.hidden);
      if (verdict.hidden) hidden.add(row);
      else hidden.delete(row);
    }

    if (!volatile) {
      const result = hidden ?? cache.hidden;
      cacheRef.current = {
        ...cache,
        cells: sheet.cells,
        styles: sheet.styles,
        filters: sheet.filters,
        hidden: result,
      };
      return result;
    }
  }

  const hidden = new Set<number>();
  let volatile = false;
  for (let row = frozenRows; row < sheet.numRows; row++) {
    const verdict = test(row);
    if (verdict.formula) volatile = true;
    if (verdict.hidden) hidden.add(row);
  }

  cacheRef.current = {
    cells: sheet.cells,
    styles: sheet.styles,
    filters: sheet.filters,
    numRows: sheet.numRows,
    frozenRows,
    hidden,
    volatile,
  };
  return hidden;
}
