/**
 * Recognizing an `IMAGE()` call in a cell's raw text.
 *
 * Here rather than beside the XLSX reader because the grid needs it too — a cell
 * edited to an IMAGE call should draw its picture — and the React layer must not
 * import the reader, which would pull the ZIP and OOXML code into every bundle.
 *
 * A regex rather than the parser: this runs on raw text while the user is still
 * typing, where the formula is routinely unfinished and would not parse.
 */

/** `_xlfn.` is Excel's prefix for functions newer than the file format. */
const IMAGE_CALL = /^=\s*(?:_xlfn\.)?IMAGE\s*\(\s*"([^"]*)"/i;

/** The URL argument of an `IMAGE(...)` call, or null when it is not one. */
export function imageUrlIn(raw: string): string | null {
  return raw.match(IMAGE_CALL)?.[1] ?? null;
}
