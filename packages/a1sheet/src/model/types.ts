/**
 * The workbook data model.
 *
 * Cell addressing is ALWAYS the internal `"${row}_${col}"` form, zero-indexed.
 * `cells` and `styles` are separate parallel maps and are never merged into a
 * single cell object — see docs/superpowers/specs for why, and note that
 * changing it means updating every read site across model/, formula/, io/, and
 * react/.
 */

/** Internal cell key: `"${row}_${col}"`, both zero-indexed. */
export type CellKey = `${number}_${number}`;

/** Zero-indexed grid position. */
export interface CellPos {
  row: number;
  col: number;
}

/**
 * A rectangular range. `r1`/`c1` is the anchor, `r2`/`c2` the active corner.
 * Not guaranteed normalized — `r2 < r1` is legal for a selection dragged
 * upward. Use `normalizeRange` before iterating.
 */
export interface Range {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

export type NumFmt =
  | "general"
  | "integer"
  | "number"
  | "percent"
  | "currency"
  | "date";

export type Align = "left" | "center" | "right";

/** Hex color, `#rrggbb`. */
export type HexColor = `#${string}`;

export interface StyleObject {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  align?: Align;
  color?: HexColor;
  bg?: HexColor;
  numFmt?: NumFmt;
  /**
   * Blocks editing and clearing. Does NOT block selection or copy — a locked
   * cell stays selectable and copyable by design.
   */
  locked?: boolean;
}

/**
 * Raw cell content as typed by the user. A leading `"="` marks a formula;
 * everything else is a literal string.
 */
export type RawCell = string;

export interface Sheet {
  id: string;
  name: string;
  cells: Record<CellKey, RawCell>;
  styles: Record<CellKey, StyleObject>;
  /**
   * Values computed by whatever wrote the file this sheet was imported from,
   * for formula cells only. Consulted ONLY when this engine cannot evaluate the
   * formula itself, so an imported workbook built on functions we do not
   * implement displays its numbers instead of a grid of `#NAME?`.
   *
   * A displayed cached value is a snapshot, not a live result: editing a cell it
   * depends on does not update it, because nothing here can recalculate a
   * formula it could not parse. Editing the formula cell itself drops its entry.
   * Empty for a sheet that was not imported.
   */
  cachedValues: Record<CellKey, CellValue>;
  /** Column index -> width in px. Absent means DEFAULT_COL_WIDTH. */
  colWidths: Record<number, number>;
  /** Row index -> height in px. Absent means ROW_HEIGHT. */
  rowHeights: Record<number, number>;
  merges: Range[];
  /** Count of rows frozen from the top. 0 = none. */
  frozenRows: number;
  /** Count of columns frozen from the left. 0 = none. */
  frozenCols: number;
  /** Manually hidden rows — distinct from rows hidden by an active filter. */
  hiddenRows: Set<number>;
  /** Display-only overrides for "A", "B", … Internal addressing is unaffected. */
  colLabels: Record<number, string>;
  /** Display-only overrides for "1", "2", … Internal addressing is unaffected. */
  rowLabels: Record<number, string>;
  /**
   * Column index -> set of allowed display values. A row is filter-hidden when
   * its value for that column is not in the set. Never mutates `cells`, so
   * clearing a filter always restores the original order and content.
   */
  filters: Record<number, Set<string>>;
  numRows: number;
  numCols: number;
}

/**
 * Named ranges are workbook-level but resolve against whichever sheet is active
 * when a formula uses them. There is no per-sheet scoping yet.
 */
export type NamedRanges = Record<string, Range>;

export interface Workbook {
  sheets: Sheet[];
  activeSheetIndex: number;
  namedRanges: NamedRanges;
}

/** A cell's computed value. Formula errors surface as `"#CYCLE!"`-style strings. */
export type CellValue = string | number | boolean;
