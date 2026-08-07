/**
 * Reading a data-validation rule's meaning out of its formula text.
 *
 * In the model rather than beside the XLSX reader because the grid needs it — a
 * cell with a list rule renders as a dropdown — and the React layer must not
 * import the reader, which would pull the ZIP and OOXML code into every bundle.
 */
import type { DataValidation } from "./types.js";

/** Quoted literals: `"a,b,c"` is a three-item list, not one string. */
const QUOTED_LIST = /^"(.*)"$/s;

/**
 * The choices a `list` rule offers, or null when they come from a range this
 * cannot resolve on its own.
 *
 * Excel writes a literal list as one quoted, comma-separated string and a
 * range-backed one as a reference — so the caller has to evaluate the latter.
 */
export function listLiterals(validation: DataValidation): string[] | null {
  const source = validation.formulas[0];
  if (!source) return null;
  const quoted = source.match(QUOTED_LIST);
  return quoted?.[1] === undefined
    ? null
    : quoted[1].split(",").map((s) => s.trim());
}
