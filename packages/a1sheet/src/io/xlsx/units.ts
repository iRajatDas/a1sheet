/**
 * Converting OOXML's sizing units to and from pixels.
 *
 * SpreadsheetML does not store pixels. Column widths are in multiples of the
 * width of the widest digit in the workbook's normal font, and row heights are
 * in points. Both conversions are lossy in the last pixel, which is why a value
 * round-trips through `Math.round` rather than being compared exactly.
 *
 * MAX_DIGIT_WIDTH is 7 for Calibri 11, the default Excel writes. A workbook
 * using a different normal font will be read a few pixels off — visibly fine,
 * and the alternative is parsing font metrics out of styles.xml to compute a
 * number that the user is free to drag anyway.
 */

const MAX_DIGIT_WIDTH = 7;
/** Excel pads a column by 5px of cell margin on top of the character width. */
const COL_PADDING_PX = 5;
const POINTS_PER_INCH = 72;
const CSS_PIXELS_PER_INCH = 96;

export function colWidthToPx(width: number): number {
  return Math.round(width * MAX_DIGIT_WIDTH) + COL_PADDING_PX;
}

export function pxToColWidth(px: number): number {
  return (px - COL_PADDING_PX) / MAX_DIGIT_WIDTH;
}

export function rowHeightToPx(points: number): number {
  return Math.round((points * CSS_PIXELS_PER_INCH) / POINTS_PER_INCH);
}

export function pxToRowHeight(px: number): number {
  return (px * POINTS_PER_INCH) / CSS_PIXELS_PER_INCH;
}
