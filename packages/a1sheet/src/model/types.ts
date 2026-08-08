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
  /** Indent steps from the cell's alignment edge. Excel's unit, not pixels. */
  indent?: number;
  /** Text rotation in degrees, -90 to 90. 255 in the file means vertical. */
  rotation?: number;
  /**
   * When set, the grid renders an interactive checkbox. Cell raw text is
   * `TRUE` / `FALSE` (or empty → unchecked).
   */
  checkbox?: boolean;
  /** Absolute or relative URL; rendered as a link when present. */
  hyperlink?: string;
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
  | { readonly type: "containsBlanks"; readonly negate: boolean }
  /**
   * Rules whose test is a statistic over the whole range rather than over the
   * cell — the top N values, or everything above the mean. Kept as a rule rather
   * than resolved on import because the statistic moves when the values do.
   */
  | {
      readonly type: "top10";
      /** How many, or what percentage when `percent`. */
      readonly rank: number;
      readonly bottom: boolean;
      readonly percent: boolean;
    }
  | {
      readonly type: "aboveAverage";
      readonly below: boolean;
      readonly orEqual: boolean;
    }
  /**
   * A rule that paints rather than styles: a gradient across the range's values,
   * a bar in proportion to them, or an icon per band.
   *
   * Carried as data rather than resolved to a style because the result depends on
   * the whole range — the colour of one cell in a scale is a function of the
   * minimum and maximum of every cell the rule covers.
   */
  | {
      readonly type: "colorScale";
      readonly stops: readonly CondScaleStop[];
    }
  | {
      readonly type: "dataBar";
      readonly color: HexColor;
      readonly min: CondScalePoint;
      readonly max: CondScalePoint;
    }
  | {
      readonly type: "iconSet";
      /** The set's name from the file, e.g. `"3TrafficLights1"`. */
      readonly set: string;
      /** Lower bounds, ascending. The first band needs no threshold. */
      readonly thresholds: readonly CondScalePoint[];
    };

/** How one end of a scale is located: by a percentile, a number, or a formula. */
export interface CondScalePoint {
  readonly kind: "min" | "max" | "num" | "percent" | "percentile" | "formula";
  /** The value or formula, when the kind needs one. */
  readonly value?: string;
}

export interface CondScaleStop extends CondScalePoint {
  readonly color: HexColor;
}

/**
 * An image drawn inside a cell, from `=IMAGE("…")`.
 *
 * `src` is either a `data:` URI of a copy embedded in the file or the source URL
 * the formula names. Embedded is preferred: no request, works offline, and cannot
 * be changed by whoever controls the address.
 */
export interface CellImage {
  src: string;
  /** Alternative text. The source URL, which is the only description available. */
  alt?: string;
}

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

/**
 * A named table over a range, as `Format as Table` creates.
 *
 * Present so structured references (`tblMatches[home_goal]`) can resolve. The
 * table's *appearance* is flattened onto cell styles at import; this is the part
 * formulas need, which cannot be flattened because a reference names a column
 * rather than a coordinate.
 */
export interface SheetTable {
  name: string;
  /** The whole table, header row included when it has one. */
  range: Range;
  /** Column names, left to right. */
  columns: readonly string[];
  headerRow: boolean;
}

/**
 * A data-validation rule over a range.
 *
 * The one that matters in practice is `list`: it is what makes a cell a dropdown,
 * and a workbook that uses them is unusable without them — the constraint is the
 * interface, not a warning about it.
 */
export interface DataValidation {
  range: Range;
  kind: "list" | "whole" | "decimal" | "date" | "textLength" | "custom";
  /**
   * The rule's operands, as formula text. For a list this is one entry: either a
   * range reference or a comma-separated set of literals in quotes.
   */
  formulas: readonly string[];
  /** Excel's `allowBlank`. An empty cell passes any rule when set. */
  allowBlank?: boolean;
  /** Comparison for the numeric and date kinds. */
  operator?: CondOperator;
  /** Shown when the value is rejected. */
  message?: string;
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
  /**
   * Images to draw in cells, from `IMAGE()` formulas. Separate from `cells`
   * because the formula's own value is a URL — the picture is what it means, not
   * what it is.
   */
  images: Record<CellKey, CellImage>;
  /** Named tables on this sheet, for structured references. */
  tables: readonly SheetTable[];
  /** Data-validation rules. A `list` rule renders the cell as a dropdown. */
  validations: readonly DataValidation[];
  /**
   * Where an array formula declares that its result goes, keyed by the anchor.
   *
   * Excel writes a dynamic array's output into the sheet as ordinary values so
   * that other readers can see it, and marks the anchor `<f t="array" ref="…">`.
   * Without knowing the region, the anchor's own output looks like content
   * standing in the way of its spill.
   */
  spillRanges: Record<CellKey, Range>;
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
  /**
   * Manually hidden columns. Rendered as zero-width rather than skipped, so
   * offsets, sticky freeze positions, and hit-testing all keep working without
   * a second notion of "which column is where".
   */
  hiddenCols: Set<number>;
  /** Display-only overrides for "A", "B", … Internal addressing is unaffected. */
  colLabels: Record<number, string>;
  /** Display-only overrides for "1", "2", … Internal addressing is unaffected. */
  rowLabels: Record<number, string>;
  /**
   * Column index -> filter criteria (values and/or fill/text colours). A row is
   * filter-hidden when it fails any criterion on any filtered column. Never
   * mutates `cells`, so clearing a filter always restores the original order.
   */
  filters: Record<number, ColumnFilter>;
  /**
   * Named snapshots of `filters`. Activating one replaces the live filters;
   * deleting/creating does not mutate cells.
   */
  filterViews: Record<string, FilterView>;
  /** Id of the view currently applied, or null when filters are ad hoc. */
  activeFilterViewId: string | null;
  /**
   * Defined names scoped to THIS sheet, which shadow the workbook's of the same
   * name. Excel writes them with a `localSheetId`, and two sheets are free to
   * define the same name differently — which is why they cannot be merged into
   * the workbook's map.
   */
  namedRanges: NamedRanges;
  /** The same, for names holding a formula rather than a range. */
  namedFormulas: Readonly<Record<string, string>>;
  numRows: number;
  numCols: number;
}

/**
 * A name to a region. Held both on the workbook and, for names Excel scoped to
 * one sheet, on the sheet — where they shadow the workbook's.
 */
export type NamedRanges = Record<string, Range>;

/**
 * Per-column filter criteria. Absent fields are unrestricted. Within one
 * column every present field must match (AND); across columns every column
 * must match (AND).
 */
export interface ColumnFilter {
  /** Allowed displayed values. */
  values?: ReadonlySet<string>;
  /** Allowed background colours (`StyleObject.bg`). */
  background?: ReadonlySet<string>;
  /** Allowed text colours (`StyleObject.color`). */
  foreground?: ReadonlySet<string>;
}

/** A named snapshot of column filters that can be re-applied later. */
export interface FilterView {
  id: string;
  name: string;
  filters: Record<number, ColumnFilter>;
}

export interface Workbook {
  sheets: Sheet[];
  activeSheetIndex: number;
  namedRanges: NamedRanges;
  /**
   * Defined names whose value is a formula rather than a range — how a modern
   * workbook names a computed table. Values are formula bodies, without the `=`.
   *
   * Separate from `namedRanges` because the two are used differently: a range is
   * a location the UI can jump to, and a formula is only ever evaluated.
   */
  namedFormulas?: Readonly<Record<string, string>>;
}

/** A cell's computed value. Formula errors surface as `"#CYCLE!"`-style strings. */
export type CellValue = string | number | boolean;
