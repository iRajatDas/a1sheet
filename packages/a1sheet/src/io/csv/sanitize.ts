/**
 * CSV formula-injection mitigation (CWE-1236).
 *
 * A CSV field beginning with `=`, `+`, `-`, `@`, tab, or CR is interpreted as a
 * FORMULA by Excel, LibreOffice, and other spreadsheet apps when the file is opened. So a
 * cell containing
 *
 *   =cmd|'/c calc'!A1
 *
 * becomes command execution on the machine of whoever opens the export. If your app
 * lets users type into a sheet and lets anyone else download it as CSV, that is a
 * real remote-code-execution path, not a theoretical one.
 *
 * Note this is an EXPORT concern. Importing such a value is harmless — a1sheet's
 * evaluator only knows its own ~30 functions and cannot spawn anything.
 *
 * Why a single quote and not a tab: tab and CR are on the dangerous list precisely
 * BECAUSE spreadsheets strip leading whitespace and then evaluate what follows, so
 * prefixing a tab does not neutralize anything. A leading `'` is the apostrophe
 * "treat as text" marker that Excel, LibreOffice, and Sheets all honor.
 */

/**
 * Characters that make a spreadsheet treat a field as a formula. Tab and CR are
 * included because they are stripped before evaluation, so they can smuggle a
 * payload past a naive first-character check.
 */
const DANGEROUS_PREFIXES = ["=", "+", "-", "@", "\t", "\r"] as const;

export type CsvInjectionMode =
  /** Prefix `'`, the "treat as text" marker. Default, and the only safe mode. */
  | "prefix-quote"
  /** Leave values exactly as-is. Only for output no one else will open. */
  | "off";

export interface SanitizeOptions {
  mode?: CsvInjectionMode;
}

/** True when a spreadsheet would treat this field as a formula. */
export function isFormulaLike(value: string): boolean {
  return DANGEROUS_PREFIXES.some((p) => value.startsWith(p));
}

const PLAIN_NUMBER = /^-?\d+(\.\d+)?([eE][+-]?\d+)?$/;

/**
 * Neutralizes a single CSV field.
 *
 * Plain numbers pass through untouched so numeric columns stay numeric — `-42` is
 * not an injection risk even though it starts with a dangerous character.
 */
export function neutralizeCsvValue(
  value: string,
  { mode = "prefix-quote" }: SanitizeOptions = {},
): string {
  if (mode === "off") return value;
  if (value === "" || !isFormulaLike(value)) return value;
  if (PLAIN_NUMBER.test(value)) return value;
  return `'${value}`;
}

/**
 * Strips a mitigation prefix added by `neutralizeCsvValue`, so a round-trip through
 * our own export and import is lossless.
 *
 * Only removes the quote when what follows is actually formula-like, so a genuine
 * value like `'tis` survives.
 */
export function denormalizeCsvValue(value: string): string {
  if (value.startsWith("'") && isFormulaLike(value.slice(1))) {
    return value.slice(1);
  }
  return value;
}

/**
 * Makes a filename safe for a download attribute: no path separators, no traversal,
 * no control characters, non-empty, length-bounded.
 *
 * Traversal segments are dropped as SEGMENTS rather than pattern-replaced, because
 * collapsing `..` textually leaves debris like `.-.-etc-passwd`.
 */
export function safeFilename(name: string, fallback = "spreadsheet"): string {
  const cleaned = name
    // biome-ignore lint/suspicious/noControlCharactersInRegex: stripping controls is the point
    .replace(/[\u0000-\u001f\u007f]/g, "")
    .split(/[/\\]+/)
    .filter((segment) => segment !== "" && segment !== "." && segment !== "..")
    .join("-")
    .replace(/^[.\s]+|[.\s]+$/g, "")
    .slice(0, 200);
  return cleaned || fallback;
}
