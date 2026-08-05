/**
 * Shared prop types for the presentational components.
 *
 * Every component here is presentational: it receives the headless API from
 * `useSpreadsheet` plus a resolved theme, and owns no workbook state.
 */
import type { Range, Sheet } from "../../model/types.js";
import type { Theme } from "../theme.js";
import type { UseSpreadsheetResult } from "../useSpreadsheet.js";

export interface BaseProps {
  api: UseSpreadsheetResult;
  theme: Theme;
  /** Class-name prefix, already normalized by `<Spreadsheet />`. */
  prefix: string;
}

export interface ColumnMenuState {
  col: number;
  x: number;
  y: number;
}

export interface ContextMenuState {
  row: number;
  col: number;
  x: number;
  y: number;
}

/** Inline rename target — column header, row header, or sheet tab. */
export interface RenamingState {
  type: "col" | "row" | "sheet";
  index: number;
  value: string;
}

export type { Range, Sheet, Theme };
