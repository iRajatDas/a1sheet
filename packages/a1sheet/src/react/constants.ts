/**
 * Layout constants. Ported from ref/Spreadsheet.jsx:22-27.
 *
 * ROW_HEIGHT is fixed on purpose: the sticky `top` offsets for frozen rows and
 * the virtualization index math both assume a uniform row height. Making rows
 * variable requires a cumulative offset table in useRowWindow — see
 * docs/LIMITATIONS.md.
 */
export const ROW_HEIGHT = 26;
export const HEADER_HEIGHT = 26;
export const ROW_HEADER_WIDTH = 44;
export const DEFAULT_COL_WIDTH = 92;

/** Extra rows rendered above and below the viewport to hide scroll tearing. */
export const BUFFER_ROWS = 6;
