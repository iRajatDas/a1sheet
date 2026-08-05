/**
 * Typed errors with stable codes.
 *
 * Consumers branch on `error.code`, never on message text — messages are free to
 * improve without being a breaking change, codes are not.
 */

export const ERROR_CODES = [
  "UNSUPPORTED_FORMAT",
  "MALFORMED_FILE",
  "NOT_A_ZIP",
  "ABORTED",
  "MISSING_PROVIDER",
  "INVALID_ARGUMENT",
  "EMPTY_WORKBOOK",
] as const;

export type ErrorCode = (typeof ERROR_CODES)[number];

/** Base class for everything this library throws. */
export class A1SheetError extends Error {
  readonly code: ErrorCode;

  constructor(code: ErrorCode, message: string, options?: { cause?: unknown }) {
    super(message, options);
    this.code = code;
    this.name = new.target.name;
  }
}

export function isA1SheetError(e: unknown): e is A1SheetError {
  return e instanceof A1SheetError;
}

/**
 * A recognized but unsupported file format — .xlsb, .xls. `format` is the detected
 * format so a consumer can build its own message.
 */
export class UnsupportedFormatError extends A1SheetError {
  readonly format: string;

  constructor(format: string, remedy: string) {
    super("UNSUPPORTED_FORMAT", `Unsupported format: ${format}. ${remedy}`);
    this.format = format;
  }
}

/** The container was readable but its contents were not what the format requires. */
export class MalformedFileError extends A1SheetError {
  constructor(what: string, options?: { cause?: unknown }) {
    super("MALFORMED_FILE", `Malformed spreadsheet: ${what}`, options);
  }
}

export class NotAZipError extends A1SheetError {
  constructor() {
    super(
      "NOT_A_ZIP",
      "Not a ZIP archive: no end-of-central-directory record found. " +
        "An .xlsx is a ZIP; a legacy .xls is not.",
    );
  }
}

/** Thrown when an `AbortSignal` fires mid-parse. */
export class AbortedError extends A1SheetError {
  constructor(operation: string) {
    super("ABORTED", `${operation} was aborted.`);
  }
}

/** A React primitive used outside its provider — programmer error, fail fast. */
export class MissingProviderError extends A1SheetError {
  constructor(component: string, provider: string) {
    super(
      "MISSING_PROVIDER",
      `<${component}> must be rendered inside <${provider}>. ` +
        `Wrap your tree in <${provider}>, or use the useSheet() hook to build your own.`,
    );
  }
}

export class InvalidArgumentError extends A1SheetError {
  constructor(what: string, expected: string, received: unknown) {
    super(
      "INVALID_ARGUMENT",
      `Invalid ${what}: expected ${expected}, received ${String(received)}.`,
    );
  }
}

export class EmptyWorkbookError extends A1SheetError {
  constructor() {
    super("EMPTY_WORKBOOK", "A workbook must contain at least one sheet.");
  }
}
