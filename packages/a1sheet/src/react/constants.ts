/**
 * Layout constants. Ported from ref/Spreadsheet.jsx:22-27.
 *
 * ROW_HEIGHT and DEFAULT_COL_WIDTH are defaults, not assumptions: a row or
 * column with an entry in `rowHeights`/`colWidths` overrides them, and both
 * `useRowWindow` and `useColWindow` resolve positions through cumulative offset
 * tables rather than multiplying by these.
 */
export const ROW_HEIGHT = 26;
export const HEADER_HEIGHT = 26;
export const ROW_HEADER_WIDTH = 44;
export const DEFAULT_COL_WIDTH = 92;

/**
 * The cell's own text metrics, which have to be constants rather than read off
 * the DOM because wrapped row heights are computed during render, before there
 * is a cell to measure.
 *
 * `theme.fontSize` and `theme.fontFamily` default to these two, so the default
 * theme cannot drift from what the measurer assumes. A consumer who overrides
 * either gets wrapped heights computed for the default face — off by a line at
 * worst, and documented in docs/LIMITATIONS.md.
 */
export const CELL_FONT_SIZE = 13;
export const CELL_FONT_STACK =
  'ui-sans-serif, system-ui, -apple-system, "Segoe UI", Roboto, sans-serif';

/**
 * Line box as a multiple of the font size. Unitless in CSS on purpose: a cell
 * with a larger `fontSize` then gets a proportionally taller line without the
 * measurement and the rendering having to agree on a pixel value separately.
 */
export const CELL_LINE_RATIO = 1.3;

/** Total horizontal padding inside a cell — the width wrapping cannot use. */
export const CELL_PADDING_X = 12;

/**
 * Space above and below the text. Chosen so one line at the default size fills
 * exactly ROW_HEIGHT: 13 × 1.3 rounds to 17, and 17 + 9 is 26. A wrapped row is
 * therefore the default height plus one line box per extra line, which is what
 * makes growth look like Sheets rather than like a jump.
 */
export const CELL_PADDING_Y = 9;

/**
 * Floors for a resize drag. Small enough to be useful for a spacer row or
 * column, large enough that the grabber for the next resize is still catchable
 * — a track dragged to zero cannot be dragged back.
 */
export const MIN_ROW_HEIGHT = 12;
export const MIN_COL_WIDTH = 24;

/**
 * Ceiling for auto-fit, which is driven by cell content and would otherwise let
 * one long string size a column off the screen.
 */
export const MAX_AUTOFIT_COL_WIDTH = 600;

/**
 * How many non-empty cells auto-fit measures in a column before it stops and
 * uses the widest it has seen. A column of a hundred thousand values does not
 * need all of them measured to find a good width, and the alternative is a
 * visible stall on a double-click.
 */
export const AUTOFIT_SAMPLE_LIMIT = 2_000;

/** Rows the "add more rows" control offers by default, matching Google Sheets. */
export const ADD_ROWS_DEFAULT = 1_000;

/**
 * Ceiling for one press of that button. Rows are cheap — they are virtualized —
 * but a mistyped number should not turn a sheet into one with ten million of
 * them and no obvious way back.
 */
export const ADD_ROWS_MAX = 100_000;

/** Extra rows rendered above and below the viewport to hide scroll tearing. */
export const BUFFER_ROWS = 6;

/**
 * Extra columns rendered left and right of the viewport. Lower than
 * BUFFER_ROWS because a column is several times wider than a row is tall, so
 * the same pixel margin costs far fewer cells.
 */
export const BUFFER_COLS = 2;

/**
 * Width of the vertical scrollbar's channel and height of the horizontal one.
 * The two are equal so the bars read as one control turned on its side.
 */
export const SCROLLBAR_SIZE = 14;

/**
 * Floor for the thumb. At a hundred thousand rows the proportional thumb is
 * under a pixel tall, which is neither visible nor grabbable.
 */
export const SCROLLBAR_MIN_THUMB = 28;
