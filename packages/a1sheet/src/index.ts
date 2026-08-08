/**
 * a1sheet — framework-agnostic entry point.
 *
 * Nothing reachable from this module imports React. The React component and
 * hooks live at "a1sheet/react".
 */

// --------------------------------------------------------------- errors
// Branch on `error.code`, never on message text — codes are stable, messages
// are free to improve.
export type { ErrorCode } from "./errors.js";
export {
  A1SheetError,
  AbortedError,
  EmptyWorkbookError,
  ERROR_CODES,
  InvalidArgumentError,
  isA1SheetError,
  MalformedFileError,
  MissingProviderError,
  NotAZipError,
  UnsupportedFormatError,
} from "./errors.js";
// --------------------------------------------------------------- format
export type { FormatOptions } from "./format/numFmt.js";
export { formatValue, NUM_FMTS } from "./format/numFmt.js";
// -------------------------------------------------------------- formula
export type { BinaryOp, CompareOp, Node, RefToken, Token } from "./formula/ast.js";
export type {
  FormulaArgMeta,
  FormulaArgType,
  FormulaCategory,
  FormulaMeta,
} from "./formula/catalog.js";
export {
  FORMULA_ARG_TYPES,
  FORMULA_CATALOG,
  FORMULA_CATEGORIES,
  formulaMeta,
  isImplementedFormula,
  suggestFormulas,
} from "./formula/catalog.js";
export { explainErrorValue } from "./formula/errorText.js";
export type { EvalContext, Evaluator } from "./formula/evaluate.js";
export {
  CYCLE_ERROR,
  createEvaluator,
  ERROR_VALUES,
  evalNode,
  isErrorValue,
} from "./formula/evaluate.js";
export type { FormulaFunction } from "./formula/functions/registry.js";
export { FUNCTIONS, registerFunction } from "./formula/functions/registry.js";
export { parseFormula } from "./formula/parse.js";
export type { RefInsertion, RefSpan } from "./formula/refEditing.js";
// Locating and inserting references in formula source, for custom editors.
export {
  findRefSpans,
  insertRefAtCaret,
  isFormulaSource,
} from "./formula/refEditing.js";
export { shiftFormulaRefs } from "./formula/refs.js";
export type { Sequence } from "./formula/sequences.js";
export {
  extendSequence,
  matchSequence,
  SEQUENCES,
} from "./formula/sequences.js";
export { extrapolateSeries } from "./formula/series.js";
export { tokenize } from "./formula/tokenize.js";
export type { FormulaArg, FormulaValue } from "./formula/values.js";
export { toNumber, toText } from "./formula/values.js";
export { csvToCells, iterCsvRows, parseCSV } from "./io/csv/read.js";
export { cellsToCSV, getUsedBounds } from "./io/csv/write.js";
export type { DownloadXlsxOptions } from "./io/download.js";
export { downloadCsv, downloadXlsx } from "./io/download.js";
// ------------------------------------------------------------------- io
export type { ReadResult, WorkbookFormat } from "./io/index.js";
export { readWorkbookFile } from "./io/index.js";
export type { AsyncReadOptions, ReadPhase, ReadProgress } from "./io/progress.js";
export { READ_PHASES } from "./io/progress.js";
export type { XlsxSheetData } from "./io/xlsx/read.js";
export { readXlsx } from "./io/xlsx/read.js";
export type { WriteXlsxOptions, XlsxSheetInput } from "./io/xlsx/write.js";
export { writeXlsx } from "./io/xlsx/write.js";
export {
  cellKey,
  colToLetters,
  lettersToCol,
  normalizeRange,
  parseCellKey,
  parseCellRef,
  parseRangeRef,
  rangeContains,
  toA1,
} from "./model/address.js";
export type { AutofillCode, FillCheckResult } from "./model/autofill.js";
export {
  AUTOFILL_CODES,
  fillDestination,
  previewFillCheck,
} from "./model/autofill.js";
export type { FilterInput } from "./model/filters.js";
export {
  activateFilterView,
  cloneColumnFilter,
  cloneFilters,
  colorMovedToTopMessage,
  createFilterView,
  deleteFilterView,
  FILTER_VIEW_MISSING,
  filteredColumnSortedMessage,
  isColumnFilterEmpty,
  normalizeColumnFilter,
  rowMatchesColumnFilter,
} from "./model/filters.js";
export type {
  FindHit,
  FindReplaceOptions,
  ReplaceAllResult,
} from "./model/findReplace.js";
export {
  findAll,
  findNext,
  replaceAll,
  replacedStatus,
} from "./model/findReplace.js";
export type { GridErrorCode } from "./model/gridErrors.js";
export {
  filterIdExists,
  filterRangeUnbounded,
  formulaParseError,
  formulaParseStart,
  GRID_ERROR_CODES,
  GridError,
  groupIntervalUnbounded,
  isGridError,
  mergeSingleton,
  rangesDifferentSheet,
  rowUnboundedRange,
  unboundedRange,
} from "./model/gridErrors.js";
export type { History } from "./model/history.js";
export {
  canRedo,
  canUndo,
  emptyHistory,
  HISTORY_LIMIT,
  push,
  redo,
  undo,
} from "./model/history.js";
export type { MergeGuardCode, MergeGuardResult } from "./model/mergeGuards.js";
export {
  checkFilterMerge,
  checkPasteMerge,
  checkSortMerge,
  MERGE_GUARD_CODES,
  rangePartiallyIntersectsMerge,
} from "./model/mergeGuards.js";
export {
  cloneSheet,
  DEFAULT_NUM_COLS,
  DEFAULT_NUM_ROWS,
  deleteCol,
  deleteRow,
  getMergeAt,
  getStyle,
  insertCol,
  insertRow,
  makeSheet,
  sortByColumn,
  uid,
} from "./model/sheet.js";
export { sortByColor } from "./model/sortByColor.js";
// ---------------------------------------------------------------- model
export type {
  Align,
  CellKey,
  CellPos,
  CellValue,
  ColumnFilter,
  FilterView,
  HexColor,
  NamedRanges,
  NumFmt,
  Range,
  RawCell,
  Sheet,
  StyleObject,
  Workbook,
} from "./model/types.js";
export {
  activeSheet,
  addSheet,
  createWorkbook,
  defineName,
  deleteName,
  deleteSheet,
  duplicateSheet,
  moveSheet,
  namedRangeAddedStatus,
  renameSheet,
} from "./model/workbook.js";
// --------------------------------------------------------------- serials
export type { SerialParts } from "./serial.js";
export { DAY_MS, msToSerial, serialToMs, serialToParts } from "./serial.js";
