"use client";

/**
 * The headless API: composes every other hook into one object.
 *
 * `<Spreadsheet />` is a thin consumer of this. Anyone who wants their own UI uses
 * this hook directly and renders whatever they like.
 *
 * A NEW evaluator is created whenever `cells` or `namedRanges` change. That is
 * intentional and cheap because evaluation is lazy and memoized per evaluator —
 * do not try to persist one across edits.
 */
import { useCallback, useMemo, useState } from "react";
import {
  type CondDecoration,
  condDecorationFor as condDecorationFor_,
  condStyleFor as condStyleFor_,
} from "../format/condFormat.js";
import { formatValue } from "../format/numFmt.js";
import { createEvaluator, type Evaluator } from "../formula/evaluate.js";
import { imageUrlIn } from "../formula/imageCall.js";
import { tableIndex } from "../formula/tableRefs.js";
import type { FormulaValue } from "../formula/values.js";
import { cellKey, normalizeRange } from "../model/address.js";
import { rejectCellValue } from "../model/cellValidation.js";
import type { Range, StyleObject, Workbook } from "../model/types.js";
import { listLiterals } from "../model/validation.js";
import { type CellFont, DEFAULT_CELL_FONT } from "./constants.js";
import { type UseClipboardResult, useClipboard } from "./useClipboard.js";
import { type UseColWindowResult, useColWindow } from "./useColWindow.js";
import { type UseEditingResult, useEditing } from "./useEditing.js";
import { type UseFillHandleResult, useFillHandle } from "./useFillHandle.js";
import { type UseFormulaRefsResult, useFormulaRefs } from "./useFormulaRefs.js";
import { type UseRowWindowResult, useRowWindow } from "./useRowWindow.js";
import { type UseSelectionResult, useSelection } from "./useSelection.js";
import { type UseSheetOpsResult, useSheetOps } from "./useSheetOps.js";
import { type UseWorkbookResult, useWorkbook } from "./useWorkbook.js";

export interface UseSpreadsheetOptions {
  /** Uncontrolled starting workbook. */
  initialWorkbook?: Workbook;
  /** Controlled workbook. Pair with `onChange`. */
  workbook?: Workbook;
  onChange?: (wb: Workbook) => void;
  initialSelection?: Range;
  /**
   * The face cells are drawn in. Wrapped text is measured against it to work out
   * how tall a row has to be, and that has to happen during render, before any
   * cell exists to read metrics off — hence a value in rather than a lookup.
   *
   * `Root` passes its resolved theme's font. Defaults to the default theme's.
   */
  cellFont?: CellFont;
}

export interface UseSpreadsheetResult
  extends UseWorkbookResult,
    UseSelectionResult,
    UseEditingResult,
    UseSheetOpsResult {
  evaluator: Evaluator;
  /** Evaluated value with the cell's number format applied. What the grid shows. */
  getDisplay(row: number, col: number): string;
  /** Evaluated value before formatting. Use for stats and comparisons. */
  getValue(row: number, col: number): FormulaValue;
  /** Raw cell content as typed, formula source included. */
  getRaw(row: number, col: number): string;
  /**
   * The style a cell's conditional formats produce, or undefined when none match.
   * Layered OVER the cell's own style by the renderer.
   */
  condStyleFor(row: number, col: number): StyleObject | undefined;
  /**
   * What a colour scale, data bar, or icon set paints on a cell. Not a style: a
   * bar is drawn behind the text at a width, and an icon beside it.
   */
  condDecorationFor(row: number, col: number): CondDecoration | undefined;
  /**
   * The values a data-validation list allows in a cell, or undefined when it has
   * no list rule. A cell with choices renders as a dropdown.
   */
  choicesFor(row: number, col: number): string[] | undefined;
  rowWindow: UseRowWindowResult;
  colWindow: UseColWindowResult;
  clipboard: UseClipboardResult;
  fill: UseFillHandleResult;
  /** Reference picking and highlighting while a formula is being typed. */
  formulaRefs: UseFormulaRefsResult;
  /** Wire to the scroll container's onScroll and its measured size. */
  setScrollTop(px: number): void;
  setScrollLeft(px: number): void;
  setViewportHeight(px: number): void;
  setViewportWidth(px: number): void;
  /**
   * The same four values read back. Virtualization is driven by these, and so
   * is anything that has to describe the window into the sheet — the grid's own
   * scrollbars are the first such thing.
   */
  scrollTop: number;
  scrollLeft: number;
  viewportHeight: number;
  viewportWidth: number;
  /** Writes a raw value into a cell, respecting `locked`. */
  setCell(row: number, col: number, raw: string): void;
  /** Commits the open edit, optionally moving the selection afterwards. */
  commitEdit(move?: [number, number]): void;
  /** True when the cell falls in the primary selection or any extra range. */
  isSelected(row: number, col: number): boolean;
  /** Every selected range, the primary first. What operations act on. */
  ranges: readonly Range[];
  /** Transient user-facing message, e.g. "That cell is locked." */
  status: string;
  setStatus(message: string): void;
  /**
   * Runs a calculation cycle, refreshing the volatile functions — `RAND`,
   * `RANDBETWEEN`, `TODAY`, `NOW`. Excel's F9.
   *
   * Every other formula is a function of cells that have not changed, so this
   * cannot alter their results. It is not an edit: nothing enters the undo
   * history and `onChange` does not fire.
   */
  recalculate(): void;
}

export function useSpreadsheet(
  opts: UseSpreadsheetOptions = {},
): UseSpreadsheetResult {
  const wb = useWorkbook({
    ...(opts.initialWorkbook ? { initialWorkbook: opts.initialWorkbook } : {}),
    ...(opts.workbook ? { workbook: opts.workbook } : {}),
    ...(opts.onChange ? { onChange: opts.onChange } : {}),
  });
  const selection = useSelection(opts.initialSelection);
  const editing = useEditing();
  const clipboard = useClipboard();
  const fill = useFillHandle();
  const formulaRefs = useFormulaRefs(editing.editing, editing.setValue);

  const clearCopied = clipboard.clearCopied;

  /**
   * Which calculation cycle the sheet is on, and the instant it began.
   *
   * An evaluator memoizes, so a volatile function — `RAND`, `TODAY`, `NOW` —
   * holds its value for the life of one. That is what makes `=RAND()` in A1 and
   * `=A1*2` in B1 agree, and it is why every volatile on the sheet has to move
   * at the same time or none of them do.
   *
   * A cycle begins on an edit, as in Excel, or on an explicit `recalculate`.
   * `serial` distinguishes two cycles that start in the same millisecond, which
   * a timestamp alone cannot.
   */
  const [calculation, setCalculation] = useState(() => ({
    serial: 0,
    at: Date.now(),
  }));
  const beginCalculation = useCallback(
    () => setCalculation((c) => ({ serial: c.serial + 1, at: Date.now() })),
    [],
  );

  /**
   * Every sheet mutation, with the copy outline dropped first.
   *
   * The outline says "this is what a paste will bring in". Once the sheet has
   * changed it may be describing cells that no longer hold what they held, so
   * it goes — which is what Excel and Sheets both do, and it is put HERE rather
   * than in each command so that a new command cannot forget.
   */
  const updateSheet = useCallback<UseWorkbookResult["updateSheet"]>(
    (fn, addHistory) => {
      clearCopied();
      beginCalculation();
      wb.updateSheet(fn, addHistory);
    },
    [wb.updateSheet, clearCopied, beginCalculation],
  );

  const patchSheet = useCallback<UseWorkbookResult["patchSheet"]>(
    (fn, addHistory) => {
      clearCopied();
      beginCalculation();
      wb.patchSheet(fn, addHistory);
    },
    [wb.patchSheet, clearCopied, beginCalculation],
  );

  /**
   * Every selected range, the primary one first. This is what an operation acts
   * on: with a Ctrl+click selection, "the selection" is all of it.
   */
  const ranges = useMemo(
    () => [selection.selection, ...selection.extraRanges],
    [selection.selection, selection.extraRanges],
  );
  const ops = useSheetOps(wb.sheet, ranges, updateSheet, wb.patchSheet);

  const [scrollTop, setScrollTop] = useState(0);
  const [scrollLeft, setScrollLeft] = useState(0);
  const [viewportHeight, setViewportHeight] = useState(500);
  const [viewportWidth, setViewportWidth] = useState(800);
  const [status, setStatus] = useState("");

  const sheets = wb.workbook.sheets;

  /**
   * Tables are indexed across the WHOLE workbook, not per sheet. A defined name
   * is workbook-level and may be used on any sheet while the table it reads sits
   * on another, so each definition carries the sheet its cells are on.
   */
  const tables = useMemo(
    () =>
      tableIndex(
        sheets.flatMap((s) => s.tables.map((t) => ({ ...t, sheet: s.name }))),
      ),
    [sheets],
  );

  /** Cells of every sheet, so a qualified reference can reach them. */
  const sheetCells = useMemo(
    () => sheets.map((s) => ({ name: s.name, cells: s.cells })),
    [sheets],
  );

  const evaluator = useMemo(
    () =>
      createEvaluator(wb.sheet.cells, wb.workbook.namedRanges, {
        cachedValues: wb.sheet.cachedValues,
        tables,
        sheets: sheetCells,
        spillRanges: wb.sheet.spillRanges,
        now: calculation.at,
        // Names the active sheet defines for itself, which shadow the workbook's.
        sheetNamedRanges: wb.sheet.namedRanges,
        sheetNamedFormulas: wb.sheet.namedFormulas,
        ...(wb.workbook.namedFormulas
          ? { namedFormulas: wb.workbook.namedFormulas }
          : {}),
      }),
    [
      wb.sheet.cells,
      wb.workbook.namedRanges,
      wb.workbook.namedFormulas,
      wb.sheet.namedRanges,
      wb.sheet.namedFormulas,
      wb.sheet.cachedValues,
      wb.sheet.spillRanges,
      tables,
      sheetCells,
      // The cycle, not just its instant: `serial` is what forces a rebuild when
      // two cycles begin in the same millisecond, which is the only thing that
      // refreshes RAND.
      calculation,
    ],
  );

  const getValue = useCallback(
    (row: number, col: number) => evaluator.getCellDisplay(row, col),
    [evaluator],
  );

  const getDisplay = useCallback(
    (row: number, col: number) => {
      const value = evaluator.getCellDisplay(row, col);
      if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
      return formatValue(value, wb.sheet.styles[cellKey(row, col)]);
    },
    [evaluator, wb.sheet.styles],
  );

  const getRaw = useCallback(
    (row: number, col: number) => wb.sheet.cells[cellKey(row, col)] ?? "",
    [wb.sheet.cells],
  );

  /**
   * Conditional formatting for a cell, or undefined. Resolved per render rather
   * than cached on the sheet, because a rule's whole purpose is to follow the
   * values — the evaluator memoizes each rule's formula, so a rule over a
   * thousand cells with absolute references costs one evaluation.
   */
  const condStyleFor = useCallback(
    (row: number, col: number) =>
      wb.sheet.condFormats.length === 0
        ? undefined
        : condStyleFor_({ condFormats: wb.sheet.condFormats, evaluator }, row, col),
    [wb.sheet.condFormats, evaluator],
  );

  /**
   * The choices a data-validation list offers for a cell, or undefined.
   *
   * A range-backed list is resolved through the evaluator, so `=$H$1:$H$9` and a
   * literal `"a,b,c"` both come back as the same thing: the strings to offer.
   */
  const choicesFor = useCallback(
    (row: number, col: number): string[] | undefined => {
      for (const validation of wb.sheet.validations) {
        if (validation.kind !== "list") continue;
        const r = normalizeRange(validation.range);
        if (row < r.r1 || row > r.r2 || col < r.c1 || col > r.c2) continue;

        const literals = listLiterals(validation);
        if (literals) return literals;

        const source = validation.formulas[0];
        if (!source) return undefined;
        const value = evaluator.evaluateArray(source.replace(/^=/, ""));
        if (!Array.isArray(value)) return undefined;
        return value
          .flat()
          .map((v) => (v === undefined ? "" : String(v)))
          .filter((v) => v !== "");
      }
      return undefined;
    },
    [wb.sheet.validations, evaluator],
  );

  /** Colour scales, data bars, and icon sets, which paint rather than style. */
  const condDecorationFor = useCallback(
    (row: number, col: number) =>
      wb.sheet.condFormats.length === 0
        ? undefined
        : condDecorationFor_(
            { condFormats: wb.sheet.condFormats, evaluator },
            row,
            col,
          ),
    [wb.sheet.condFormats, evaluator],
  );

  // Held by value, not by reference: a caller passing an object literal every
  // render would otherwise throw away every wrapped-cell measurement each time.
  const fontFamily = opts.cellFont?.family ?? DEFAULT_CELL_FONT.family;
  const fontSize = opts.cellFont?.size ?? DEFAULT_CELL_FONT.size;
  const cellFont = useMemo<CellFont>(
    () => ({ family: fontFamily, size: fontSize }),
    [fontFamily, fontSize],
  );

  const rowWindow = useRowWindow({
    sheet: wb.sheet,
    scrollTop,
    viewportHeight,
    getDisplay,
    cellFont,
  });
  const colWindow = useColWindow(wb.sheet, scrollLeft, viewportWidth);

  const setCell = useCallback(
    (row: number, col: number, raw: string) => {
      const rejection = rejectCellValue(wb.sheet, row, col, raw, evaluator);
      if (rejection) {
        setStatus(rejection.message);
        return;
      }
      patchSheet((sheet) => {
        const key = cellKey(row, col);
        if (sheet.styles[key]?.locked) return null;
        const cells = { ...sheet.cells };
        if (raw === "") delete cells[key];
        else cells[key] = raw;

        const cachedValues = { ...sheet.cachedValues };
        delete cachedValues[key];

        const images = { ...sheet.images };
        const url = imageUrlIn(raw);
        if (url) images[key] = { src: url, alt: url };
        else delete images[key];

        return { cells, cachedValues, images };
      });
    },
    [patchSheet, wb.sheet, evaluator, setStatus],
  );

  const commitEdit = useCallback(
    (move?: [number, number]) => {
      const done = editing.commit();
      // Any reference the user was dragging out ends with the edit that owns it.
      formulaRefs.endPick();
      if (done) setCell(done.row, done.col, done.value);
      if (move) {
        selection.move(move[0], move[1], wb.sheet.numRows, wb.sheet.numCols);
      }
    },
    [editing, formulaRefs, setCell, selection, wb.sheet.numRows, wb.sheet.numCols],
  );

  const isSelected = useCallback(
    (row: number, col: number) => {
      const inRect = (r: Range) => {
        const n = normalizeRange(r);
        return row >= n.r1 && row <= n.r2 && col >= n.c1 && col <= n.c2;
      };
      return inRect(selection.selection) || selection.extraRanges.some(inRect);
    },
    [selection.selection, selection.extraRanges],
  );

  return {
    ...wb,
    updateSheet,
    ...selection,
    ...editing,
    ...ops,
    evaluator,
    getDisplay,
    getValue,
    getRaw,
    condStyleFor,
    condDecorationFor,
    choicesFor,
    rowWindow,
    colWindow,
    clipboard,
    fill,
    formulaRefs,
    setScrollTop,
    setScrollLeft,
    setViewportHeight,
    setViewportWidth,
    scrollTop,
    scrollLeft,
    viewportHeight,
    viewportWidth,
    setCell,
    commitEdit,
    isSelected,
    ranges,
    status,
    setStatus,
    recalculate: beginCalculation,
  };
}
