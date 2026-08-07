/**
 * Structured references — `tblMatches[home_goal]`.
 *
 * A table is a named range with named columns and a header row, and a structured
 * reference addresses it by those names instead of by coordinates. Excel's own
 * templates are built on them, so a workbook using `Format as Table` is full of
 * formulas that mean nothing without this.
 *
 * The selector grammar, all of which appears in real files:
 *
 *   tbl[col]                      the column's data cells
 *   tbl[[col]]                    the same, brackets optional around one name
 *   tbl[[a]:[b]]                  every column from a to b
 *   tbl[#All] / [#Data] / [#Headers]
 *   tbl[[#This Row],[a]]          one cell, on the row the formula is on
 *   tbl[[#This Row],[a]:[b]]      a span of that row
 *   tbl[@col] / tbl[@[col]]       shorthand for [#This Row]
 */
import type { Range } from "../model/types.js";

/** One table as the formula layer needs to see it. */
export interface TableDefinition {
  /** The whole table including its header row, if it has one. */
  range: Range;
  /**
   * The sheet the table's cells are on. A workbook-level defined name is
   * evaluated wherever it is used, so a table it references cannot be assumed to
   * be on the current sheet.
   */
  sheet?: string;
  /** Column names in order, left to right. */
  columns: readonly string[];
  headerRow: boolean;
}

/** Table name -> definition. Lookup is case-insensitive, as Excel's is. */
export interface TableIndex {
  get(name: string): TableDefinition | undefined;
}

export const EMPTY_TABLE_INDEX: TableIndex = { get: () => undefined };

/** Builds a case-insensitive index from a list of named tables. */
export function tableIndex(
  tables: readonly (TableDefinition & { name: string })[],
): TableIndex {
  if (tables.length === 0) return EMPTY_TABLE_INDEX;
  const byLower = new Map<string, TableDefinition>();
  for (const t of tables) byLower.set(t.name.toLowerCase(), t);
  return { get: (name) => byLower.get(name.toLowerCase()) };
}

type Section = "all" | "data" | "headers" | "thisRow";

interface Selector {
  section: Section;
  /** Column names the selector narrows to. Empty means every column. */
  columns: string[];
}

/** Splits a selector on commas that are not inside brackets. */
function topLevelParts(spec: string): string[] {
  const parts: string[] = [];
  let depth = 0;
  let current = "";
  for (const ch of spec) {
    if (ch === "[") depth++;
    if (ch === "]") depth--;
    if (ch === "," && depth === 0) {
      parts.push(current);
      current = "";
      continue;
    }
    current += ch;
  }
  parts.push(current);
  return parts.map((p) => p.trim()).filter((p) => p !== "");
}

/** `[name]` -> `name`; a bare name is returned as it is. */
function unbracket(part: string): string {
  const trimmed = part.trim();
  return trimmed.startsWith("[") && trimmed.endsWith("]")
    ? trimmed.slice(1, -1).trim()
    : trimmed;
}

const SECTIONS: Record<string, Section> = {
  "#all": "all",
  "#data": "data",
  "#headers": "headers",
  "#this row": "thisRow",
};

export function parseSelector(spec: string): Selector | null {
  const trimmed = spec.trim();
  // `@col` and `@[col]` are the shorthand for a this-row reference.
  if (trimmed.startsWith("@")) {
    const inner = parseSelector(trimmed.slice(1));
    return inner ? { ...inner, section: "thisRow" } : null;
  }

  let section: Section = "data";
  const columns: string[] = [];

  for (const part of topLevelParts(trimmed)) {
    // A span, `[a]:[b]`, names its endpoints and everything between them.
    if (part.includes("]:[")) {
      const [from, to] = part.split(":");
      if (from === undefined || to === undefined) return null;
      columns.push(unbracket(from), unbracket(to));
      continue;
    }

    const name = unbracket(part);
    const asSection = SECTIONS[name.toLowerCase()];
    if (asSection) {
      section = asSection;
      continue;
    }
    columns.push(name);
  }

  return { section, columns };
}

export interface ResolveOptions {
  table: TableDefinition;
  spec: string;
  /** The row the formula lives on, for `[#This Row]`. */
  currentRow: number;
}

/**
 * The range a structured reference denotes, or null when the selector names a
 * column the table does not have.
 *
 * A span between two columns spans the columns BETWEEN them too, in table order,
 * which is why the endpoints are resolved to indices rather than kept as names.
 */
export function resolveSelector(opts: ResolveOptions): Range | null {
  const { table, spec, currentRow } = opts;
  const selector = parseSelector(spec);
  if (!selector) return null;

  const lower = table.columns.map((c) => c.toLowerCase());
  const indices: number[] = [];
  for (const name of selector.columns) {
    const index = lower.indexOf(name.toLowerCase());
    if (index === -1) return null;
    indices.push(index);
  }

  const c1 = indices.length ? Math.min(...indices) : 0;
  const c2 = indices.length ? Math.max(...indices) : table.columns.length - 1;
  const left = table.range.c1 + c1;
  const right = table.range.c1 + c2;

  const headerOffset = table.headerRow ? 1 : 0;
  const firstData = table.range.r1 + headerOffset;

  switch (selector.section) {
    case "headers":
      return table.headerRow
        ? { r1: table.range.r1, c1: left, r2: table.range.r1, c2: right }
        : null;
    case "all":
      return { r1: table.range.r1, c1: left, r2: table.range.r2, c2: right };
    case "thisRow":
      // Outside the table the reference has no row to mean, which is #VALUE! in
      // Excel; null lets the caller say so.
      return currentRow < firstData || currentRow > table.range.r2
        ? null
        : { r1: currentRow, c1: left, r2: currentRow, c2: right };
    default:
      return { r1: firstData, c1: left, r2: table.range.r2, c2: right };
  }
}
