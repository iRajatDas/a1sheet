"use client";

/**
 * a1sheet/react — primitives, hooks, and one preset.
 *
 * Two ways in, in order of preference:
 *
 * 1. Compose primitives (the supported extension point):
 *      <Sheet.Root defaultWorkbook={wb}>
 *        <Sheet.Toolbar />
 *        <Sheet.Grid />
 *        <MyFooter />        // your component, our data, via useSheet()
 *      </Sheet.Root>
 *
 * 2. Use the preset when the default arrangement is fine:
 *      <Spreadsheet defaultWorkbook={wb} />
 *
 * There is deliberately no third way. No `show*` props, no layout config — an
 * absent part is a child you did not render.
 *
 * React is a peerDependency. Everything framework-agnostic lives at "a1sheet".
 */

import { AddRows } from "./components/AddRows.js";
import { ColumnMenu } from "./components/ColumnMenu.js";
import { ContextMenu } from "./components/ContextMenu.js";
import { FileMenu } from "./components/FileMenu.js";
import { FormulaBar } from "./components/FormulaBar.js";
import { Grid } from "./components/Grid.js";
import { SheetTabs } from "./components/SheetTabs.js";
import { StatusBar } from "./components/StatusBar.js";
import { Toolbar } from "./components/Toolbar.js";
import { Root } from "./Root.js";

/**
 * Namespaced primitives. `Sheet.Root` provides context; every other member must be
 * rendered inside it and throws a named error if it is not.
 */
export const Sheet = {
  Root,
  Toolbar,
  FileMenu,
  FormulaBar,
  Grid,
  AddRows,
  Tabs: SheetTabs,
  StatusBar,
  ContextMenu,
  ColumnMenu,
} as const;

export { AddRows } from "./components/AddRows.js";
export type { CellProps } from "./components/Cell.js";
export { Cell } from "./components/Cell.js";
export { ColumnMenu } from "./components/ColumnMenu.js";
export { ContextMenu } from "./components/ContextMenu.js";
export { FileMenu } from "./components/FileMenu.js";
export { FormulaBar } from "./components/FormulaBar.js";
export type { GridProps } from "./components/Grid.js";
export { Grid } from "./components/Grid.js";
export { SheetTabs } from "./components/SheetTabs.js";
export { StatusBar } from "./components/StatusBar.js";
export type { ToolbarProps } from "./components/Toolbar.js";
export { Toolbar } from "./components/Toolbar.js";
export type { CellFont } from "./constants.js";
export {
  ADD_ROWS_DEFAULT,
  ADD_ROWS_MAX,
  BUFFER_ROWS,
  DEFAULT_CELL_FONT,
  DEFAULT_COL_WIDTH,
  HEADER_HEIGHT,
  ROW_HEADER_WIDTH,
  ROW_HEIGHT,
} from "./constants.js";
export type { SheetContextValue, SheetUiState } from "./context.js";
// Context access, for consumers writing their own primitives.
export { useSheet, useSheetContext } from "./context.js";
export type { SheetRootHandle, SheetRootProps } from "./Root.js";
// Also exported individually, for named imports and tree-shaking clarity.
export { Root } from "./Root.js";
export type { SlotProps } from "./Slot.js";

// asChild plumbing, for consumers building polymorphic wrappers.
export { Slot } from "./Slot.js";
export type { SpreadsheetProps } from "./Spreadsheet.js";
// The preset.
export { Spreadsheet } from "./Spreadsheet.js";
export { buildCss } from "./styles.js";
export type { Theme } from "./theme.js";
// Theming.
export { defaultTheme, resolveTheme, themeCellFont } from "./theme.js";
// View types.
export type {
  AsChildProps,
  ColumnMenuState,
  ContextMenuState,
  MenuAnchor,
  RenamingState,
} from "./types.js";
export type { CaretBinding } from "./useCaretBinding.js";
export { useCaretBinding } from "./useCaretBinding.js";
export type { CopiedGrid, UseClipboardResult } from "./useClipboard.js";
export { useClipboard } from "./useClipboard.js";
export type { UseColWindowResult } from "./useColWindow.js";
export { useColWindow } from "./useColWindow.js";
export type { EditingState, UseEditingResult } from "./useEditing.js";
export { useEditing } from "./useEditing.js";
export type { UseFillHandleResult } from "./useFillHandle.js";
export { useFillHandle } from "./useFillHandle.js";
export type { UseFormulaRefsResult } from "./useFormulaRefs.js";
export { useFormulaRefs } from "./useFormulaRefs.js";
export type { UseRowWindowResult, WindowRow } from "./useRowWindow.js";
export { useRowWindow } from "./useRowWindow.js";
export type { UseSelectionResult } from "./useSelection.js";
export { useSelection } from "./useSelection.js";
export type { UseSheetOpsResult } from "./useSheetOps.js";
export { useSheetOps } from "./useSheetOps.js";
export type {
  UseSpreadsheetOptions,
  UseSpreadsheetResult,
} from "./useSpreadsheet.js";
// Headless hooks — usable with no a1sheet component at all.
export { useSpreadsheet } from "./useSpreadsheet.js";
export type {
  SheetUpdater,
  UseWorkbookOptions,
  UseWorkbookResult,
  WorkbookUpdater,
} from "./useWorkbook.js";
export { useWorkbook } from "./useWorkbook.js";
