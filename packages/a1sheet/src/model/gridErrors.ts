/**
 * Grid-boundary errors — fail-fast for programmer misuse.
 *
 * User-data problems (bad formulas, partial merges) stay soft via status or
 * cell error values. These codes are for callers who pass impossible ranges.
 */
import { A1SheetError, type ErrorCode } from "../errors.js";

export type GridErrorCode = Extract<
  ErrorCode,
  | "GRID_RANGE_UNBOUNDED"
  | "GRID_RANGE_ROW_UNBOUNDED"
  | "GRID_RANGES_DIFFERENT_SHEET"
  | "MERGE_RANGE_SINGLETON"
  | "GROUP_INTERVAL_UNBOUNDED"
  | "FILTER_ID_EXISTS"
  | "FILTER_RANGE_UNBOUNDED"
  | "FORMULA_PARSE_START"
  | "FORMULA_PARSE_ERROR"
>;

export const GRID_ERROR_CODES: readonly GridErrorCode[] = [
  "GRID_RANGE_UNBOUNDED",
  "GRID_RANGE_ROW_UNBOUNDED",
  "GRID_RANGES_DIFFERENT_SHEET",
  "MERGE_RANGE_SINGLETON",
  "GROUP_INTERVAL_UNBOUNDED",
  "FILTER_ID_EXISTS",
  "FILTER_RANGE_UNBOUNDED",
  "FORMULA_PARSE_START",
  "FORMULA_PARSE_ERROR",
] as const;

export class GridError extends A1SheetError {
  declare readonly code: GridErrorCode;

  constructor(code: GridErrorCode, message: string) {
    super(code, message);
    this.name = "GridError";
  }
}

export function isGridError(e: unknown): e is GridError {
  return e instanceof GridError;
}

export function unboundedRange(label: string): GridError {
  return new GridError(
    "GRID_RANGE_UNBOUNDED",
    `Grid range must be bounded: ${label}`,
  );
}

export function rowUnboundedRange(): GridError {
  return new GridError(
    "GRID_RANGE_ROW_UNBOUNDED",
    "Grid range must be row-bounded",
  );
}

export function rangesDifferentSheet(): GridError {
  return new GridError(
    "GRID_RANGES_DIFFERENT_SHEET",
    "Grid ranges must be on same sheet",
  );
}

export function mergeSingleton(label: string): GridError {
  return new GridError(
    "MERGE_RANGE_SINGLETON",
    `merge range must span multiple cells: (${label})`,
  );
}

export function groupIntervalUnbounded(): GridError {
  return new GridError(
    "GROUP_INTERVAL_UNBOUNDED",
    "Group interval must be bounded",
  );
}

export function filterIdExists(id: string): GridError {
  return new GridError("FILTER_ID_EXISTS", `Filter with id ${id} already exists`);
}

export function filterRangeUnbounded(): GridError {
  return new GridError(
    "FILTER_RANGE_UNBOUNDED",
    "filter ranges should be bounded",
  );
}

export function formulaParseStart(): GridError {
  return new GridError(
    "FORMULA_PARSE_START",
    "Formula must end with start element",
  );
}

export function formulaParseError(): GridError {
  return new GridError("FORMULA_PARSE_ERROR", "Formula parse error.");
}
