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
import { condStyleFor as condStyleFor_ } from "../format/condFormat.js";
import { formatValue } from "../format/numFmt.js";
import { createEvaluator, type Evaluator } from "../formula/evaluate.js";
import { imageUrlIn } from "../formula/imageCall.js";
import type { FormulaValue } from "../formula/values.js";
import { cellKey, normalizeRange } from "../model/address.js";
import type { Range, StyleObject, Workbook } from "../model/types.js";
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
  /** Transient user-facing message, e.g. "That cell is locked." */
  status: string;
  setStatus(message: string): void;
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
  const ops = useSheetOps(wb.sheet, selection.selection, wb.updateSheet);

  const [scrollTop, setScrollTop] = useState(0);
  const [scrollLeft, setScrollLeft] = useState(0);
  const [viewportHeight, setViewportHeight] = useState(500);
  const [viewportWidth, setViewportWidth] = useState(800);
  const [status, setStatus] = useState("");

  const evaluator = useMemo(
    () =>
      createEvaluator(
        wb.sheet.cells,
        wb.workbook.namedRanges,
        wb.sheet.cachedValues,
      ),
    [wb.sheet.cells, wb.workbook.namedRanges, wb.sheet.cachedValues],
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

  const rowWindow = useRowWindow(wb.sheet, scrollTop, viewportHeight, getDisplay);
  const colWindow = useColWindow(wb.sheet, scrollLeft, viewportWidth);

  const setCell = useCallback(
    (row: number, col: number, raw: string) => {
      wb.updateSheet((sheet) => {
        const key = cellKey(row, col);
        if (sheet.styles[key]?.locked) return sheet;
        if (raw === "") delete sheet.cells[key];
        else sheet.cells[key] = raw;
        // Whatever an import computed for this cell describes the formula that
        // was here, not the one the user just typed.
        delete sheet.cachedValues[key];

        // Typing an IMAGE() call draws the image, and replacing that formula
        // removes it — otherwise the picture would outlive the formula that
        // asked for it. An imported cell keeps its embedded copy until edited,
        // since the URL is all a typed formula can offer.
        const url = imageUrlIn(raw);
        if (url) sheet.images[key] = { src: url, alt: url };
        else delete sheet.images[key];

        return sheet;
      });
    },
    [wb],
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
    ...selection,
    ...editing,
    ...ops,
    evaluator,
    getDisplay,
    getValue,
    getRaw,
    condStyleFor,
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
    status,
    setStatus,
  };
}
