/**
 * a1sheet/react — hooks and the drop-in component.
 *
 * React is a peerDependency. Everything framework-agnostic lives at "a1sheet".
 */

// Presentational components, exported so a custom layout can reuse them.
export type { CellProps } from "./components/Cell.js";
export { Cell } from "./components/Cell.js";
export type { ColumnMenuProps } from "./components/ColumnMenu.js";
export { ColumnMenu } from "./components/ColumnMenu.js";
export type { ContextMenuProps } from "./components/ContextMenu.js";
export { ContextMenu } from "./components/ContextMenu.js";
export { FormulaBar } from "./components/FormulaBar.js";
export type { GridProps } from "./components/Grid.js";
export { Grid } from "./components/Grid.js";
export type {
  BaseProps,
  ColumnMenuState,
  ContextMenuState,
  RenamingState,
} from "./components/props.js";
export type { SheetTabsProps } from "./components/SheetTabs.js";
export { SheetTabs } from "./components/SheetTabs.js";
export { StatusBar } from "./components/StatusBar.js";
export type { ToolbarProps } from "./components/Toolbar.js";
export { Toolbar } from "./components/Toolbar.js";
export {
  BUFFER_ROWS,
  DEFAULT_COL_WIDTH,
  HEADER_HEIGHT,
  ROW_HEADER_WIDTH,
  ROW_HEIGHT,
} from "./constants.js";
export type { SpreadsheetProps } from "./Spreadsheet.js";
export { Spreadsheet } from "./Spreadsheet.js";
export { buildCss } from "./styles.js";
export type { Theme } from "./theme.js";
export { defaultTheme, resolveTheme } from "./theme.js";
export type { CopiedGrid, UseClipboardResult } from "./useClipboard.js";
export { useClipboard } from "./useClipboard.js";
export type { EditingState, UseEditingResult } from "./useEditing.js";
export { useEditing } from "./useEditing.js";
export type { UseFillHandleResult } from "./useFillHandle.js";
export { useFillHandle } from "./useFillHandle.js";
export type { UseRowWindowResult, WindowRow } from "./useRowWindow.js";
export { useRowWindow } from "./useRowWindow.js";
export type { UseSelectionResult } from "./useSelection.js";
export { useSelection } from "./useSelection.js";
export type { UseSheetOpsResult } from "./useSheetOps.js";
export { useSheetOps } from "./useSheetOps.js";
// Headless API — use these to build your own UI.
export type {
  UseSpreadsheetOptions,
  UseSpreadsheetResult,
} from "./useSpreadsheet.js";
export { useSpreadsheet } from "./useSpreadsheet.js";
export type {
  SheetUpdater,
  UseWorkbookOptions,
  UseWorkbookResult,
  WorkbookUpdater,
} from "./useWorkbook.js";
export { useWorkbook } from "./useWorkbook.js";
