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

/**
 * Vertical placement within the cell box. Absent leaves the grid's default, which
 * is centred — Excel's own default is `bottom`, but a grid of centred rows reads
 * better and every existing sheet already looks that way.
 */
export type VerticalAlign = "top" | "middle" | "bottom";

/** Hex color, `#rrggbb`. */
export type HexColor = `#${string}`;

/**
 * Border line kinds, named after the OOXML values they come from. Each maps to a
 * CSS width and style in the renderer — `hair` is the thinnest line Excel draws,
 * `double` is two lines, and the `*Dash*` family differs from its plain
 * counterpart only in weight.
 */
export type BorderLine =
  | "hair"
  | "thin"
  | "medium"
  | "thick"
  | "double"
  | "dotted"
  | "dashed"
  | "dashDot"
  | "dashDotDot"
  | "mediumDashed"
  | "mediumDashDot"
  | "mediumDashDotDot"
  | "slantDashDot";

export interface BorderEdge {
  line: BorderLine;
  /** Absent means the cell's text colour, which is Excel's `auto`. */
  color?: HexColor;
}

/** Only the edges that are set are present. A cell with none has no `borders`. */
export interface CellBorders {
  top?: BorderEdge;
  right?: BorderEdge;
  bottom?: BorderEdge;
  left?: BorderEdge;
}

/**
 * A linear gradient fill. Excel's gradients are richer than this — it has path
 * gradients too — but a stop list at an angle covers what a spreadsheet uses and
 * maps directly onto a CSS `linear-gradient`.
 */
export interface GradientFill {
  /** Degrees clockwise from a left-to-right sweep, as OOXML states it. */
  degree: number;
  stops: readonly { position: number; color: HexColor }[];
}

export interface StyleObject {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  align?: Align;
  valign?: VerticalAlign;
  /** Wraps text onto more lines instead of clipping it at the cell edge. */
  wrap?: boolean;
  color?: HexColor;
  bg?: HexColor;
  /** Takes precedence over `bg` when both are set, as a gradient covers the box. */
  gradient?: GradientFill;
  borders?: CellBorders;
  /** Font stack. A bare family name from a file; any CSS font list is valid. */
  fontFamily?: string;
  /** Pixels. Absent means the theme's `fontSize`. */
  fontSize?: number;
  numFmt?: NumFmt;
  /**
   * The literal format code from the file, e.g. `"+0;-0;0"` or `"dd/mm/yyyy"`.
   * Preferred over `numFmt` when present: `numFmt` is one of six buckets, and a
   * real workbook's formats do not fit in six. Kept alongside rather than
   * replacing it so the format dropdown still has something to show.
   */
  numFmtCode?: string;
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

export type CondOperator =
  | "lessThan"
  | "lessThanOrEqual"
  | "equal"
  | "notEqual"
  | "greaterThanOrEqual"
  | "greaterThan"
  | "between"
  | "notBetween";

/**
 * What decides whether a conditional format applies to a cell.
 *
 * Deliberately not every rule Excel has: `colorScale`, `dataBar`, and `iconSet`
 * are graphical rather than a style, and `top10`/`aboveAverage` need statistics
 * over the whole range. Those are read and ignored, so a file carrying them
 * imports without them rather than failing.
 */
export type CondRule =
  /** A formula, relative to the top-left of the rule's range. Truthy applies. */
  | { readonly type: "expression"; readonly formula: string }
  | {
      readonly type: "cellIs";
      readonly operator: CondOperator;
      /** One operand, or two for `between`/`notBetween`. */
      readonly operands: readonly string[];
    }
  | {
      readonly type: "containsText";
      readonly text: string;
      readonly negate: boolean;
    }
  | { readonly type: "containsBlanks"; readonly negate: boolean };

export interface CondFormat {
  range: Range;
  /**
   * Excel's rule priority: 1 is checked first. Where two matching rules set the
   * same key the higher-priority one wins, so they are applied lowest-priority
   * first and overwritten.
   */
  priority: number;
  rule: CondRule;
  /** The partial style applied when the rule matches. */
  style: StyleObject;
  /** When this rule matches, no lower-priority rule is considered. */
  stopIfTrue?: boolean;
}

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
  /**
   * Conditional formats, layered over `styles` and under nothing. Evaluated at
   * render time rather than flattened, so a rule keeps reacting to edits — which
   * is the whole point of it being conditional.
   */
  condFormats: readonly CondFormat[];
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
