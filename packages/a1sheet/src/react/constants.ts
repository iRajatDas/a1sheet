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
