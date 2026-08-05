/**
 * CSV formula-injection mitigation (CWE-1236).
 *
 * A CSV field beginning with `=`, `+`, `-`, `@`, tab, or CR is interpreted as a
 * FORMULA by Excel, LibreOffice, and Google Sheets when the file is opened. So a
 * cell containing
 *
 *   =cmd|'/c calc'!A1
 *
 * becomes command execution on the machine of whoever opens the export. If your
 * app lets users type into a sheet and lets anyone else download it as CSV, that is
 * a real remote-code-execution path, not a theoretical one.
 *
 * Note this is an EXPORT concern. Importing such a value is harmless — a1sheet's
 * evaluator only knows its own ~30 functions and cannot spawn anything.
 *
 * The mitigation prefixes a tab, which every major spreadsheet treats as "this is
 * text" while keeping the value legible. The alternative (a leading `'`) is
 * Excel-specific and shows up as a literal quote elsewhere.
 */

/** Characters that make a spreadsheet treat a field as a formula. */
const DANGEROUS_PREFIXES = ["=", "+", "-", "@", "\t", "\r"] as const;

export type CsvInjectionMode =
  /** Prefix a tab so the value is treated as text. Default. */
  | "prefix-tab"
  /** Prefix a single quote — Excel-friendly, visible elsewhere. */
  | "prefix-quote"
  /** Leave values exactly as-is. Only for trusted, non-shared output. */
  | "off";

export interface SanitizeOptions {
  mode?: CsvInjectionMode;
}

/**
 * True when a spreadsheet would treat this field as a formula.
 *
 * Deliberately includes `-`, which means a negative number like `-5` is also
 * flagged. `neutralizeCsvValue` special-cases plain numbers so they stay numeric —
 * checking the prefix alone is not enough to decide.
 */
export function isFormulaLike(value: string): boolean {
  return DANGEROUS_PREFIXES.some((p) => value.startsWith(p));
}

const PLAIN_NUMBER = /^-?\d+(\.\d+)?([eE][+-]?\d+)?$/;

/**
 * Neutralizes a single CSV field. Plain numbers pass through untouched so numeric
 * columns stay numeric — `-42` is not an injection risk.
 */
export function neutralizeCsvValue(
  value: string,
  { mode = "prefix-tab" }: SanitizeOptions = {},
): string {
  if (mode === "off") return value;
  if (value === "" || !isFormulaLike(value)) return value;
  if (PLAIN_NUMBER.test(value)) return value;
  return mode === "prefix-quote" ? `'${value}` : `\t${value}`;
}

/**
 * Strips a mitigation prefix added by `neutralizeCsvValue`, so a round-trip through
 * our own export and import is lossless.
 */
export function denormalizeCsvValue(value: string): string {
  if (value.startsWith("\t") && isFormulaLike(value.slice(1))) {
    return value.slice(1);
  }
  return value;
}

/**
 * Makes a filename safe for a download attribute: no path separators, no traversal,
 * no control characters, non-empty, length-bounded.
 */
export function safeFilename(name: string, fallback = "spreadsheet"): string {
  const cleaned = name
    // biome-ignore lint/suspicious/noControlCharactersInRegex: stripping controls is the point
    .replace(/[\u0000-\u001f\u007f]/g, "")
    .replace(/[/\\]/g, "-")
    .replace(/\.{2,}/g, ".")
    .replace(/^[.\s]+|[.\s]+$/g, "")
    .slice(0, 200);
  return cleaned || fallback;
}
