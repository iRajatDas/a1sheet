/**
 * The drop-in component. A thin shell: it calls `useSpreadsheet`, resolves the
 * theme, and lays out the presentational children. All state lives in the hooks.
 *
 * Keyboard and clipboard both route through a hidden 1px, opacity-0 `<textarea>`
 * that holds focus whenever a cell is selected but not being edited. It is the
 * target for keydown (navigation, delete, Ctrl+B/I/Z/Y, and "type a character to
 * start editing"), copy, and paste. When a cell IS being edited, focus moves to
 * that cell's `<input>` instead.
 *
 * A new global shortcut belongs in `onKeyDown` below, after the `isEditing` guard.
 */
import { type ReactNode, useCallback, useEffect, useRef, useState } from "react";
import { downloadCsv, downloadXlsx } from "../io/download.js";
import { readWorkbookFile } from "../io/index.js";
import { makeSheet } from "../model/sheet.js";
import type { Sheet, Workbook } from "../model/types.js";
import { ColumnMenu } from "./components/ColumnMenu.js";
import { ContextMenu } from "./components/ContextMenu.js";
import { FormulaBar } from "./components/FormulaBar.js";
import { Grid } from "./components/Grid.js";
import type {
  ColumnMenuState,
  ContextMenuState,
  RenamingState,
} from "./components/props.js";
import { SheetTabs } from "./components/SheetTabs.js";
import { StatusBar } from "./components/StatusBar.js";
import { Toolbar } from "./components/Toolbar.js";
import { buildCss } from "./styles.js";
import { resolveTheme, type Theme } from "./theme.js";
import { useSpreadsheet } from "./useSpreadsheet.js";

export interface SpreadsheetProps {
  initialWorkbook?: Workbook;
  onChange?: (wb: Workbook) => void;
  theme?: Partial<Theme>;
  /** Class-name prefix for the injected CSS. Defaults to "a1s-". */
  classNamePrefix?: string;
  /** Height of the component. Anything CSS accepts. */
  height?: string | number;
  showToolbar?: boolean;
  showFormulaBar?: boolean;
  showSheetTabs?: boolean;
  showStatusBar?: boolean;
  /** Set false to drop the import/export buttons from the toolbar. */
  enableFileIO?: boolean;
}

export function Spreadsheet({
  initialWorkbook,
  onChange,
  theme: themeOverride,
  classNamePrefix = "a1s-",
  height = 600,
  showToolbar = true,
  showFormulaBar = true,
  showSheetTabs = true,
  showStatusBar = true,
  enableFileIO = true,
}: SpreadsheetProps): ReactNode {
  const api = useSpreadsheet({
    ...(initialWorkbook ? { initialWorkbook } : {}),
    ...(onChange ? { onChange } : {}),
  });
  const theme = resolveTheme(themeOverride);
  const prefix = classNamePrefix;

  const hiddenRef = useRef<HTMLTextAreaElement>(null);
  const [renaming, setRenaming] = useState<RenamingState | null>(null);
  const [contextMenu, setContextMenu] = useState<ContextMenuState | null>(null);
  const [columnMenu, setColumnMenu] = useState<ColumnMenuState | null>(null);

  const { sheet, selection, isEditing } = api;

  // Keep the hidden textarea focused whenever a cell is selected and no edit is
  // open. activeSheetIndex is a deliberate trigger: switching sheets reclaims focus.
  // biome-ignore lint/correctness/useExhaustiveDependencies: intentional trigger
  useEffect(() => {
    if (!isEditing && !renaming) hiddenRef.current?.focus();
  }, [isEditing, renaming, api.workbook.activeSheetIndex]);

  // Any click outside a menu dismisses it.
  useEffect(() => {
    if (!contextMenu && !columnMenu) return;
    const close = () => {
      setContextMenu(null);
      setColumnMenu(null);
    };
    window.addEventListener("click", close);
    return () => window.removeEventListener("click", close);
  }, [contextMenu, columnMenu]);

  const onKeyDown = useCallback(
    (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
      if (api.isEditing) return;
      const { numRows, numCols } = sheet;
      const mod = e.ctrlKey || e.metaKey;

      if (mod) {
        const k = e.key.toLowerCase();
        if (k === "z") {
          e.preventDefault();
          api.undo();
          return;
        }
        if (k === "y") {
          e.preventDefault();
          api.redo();
          return;
        }
        if (k === "b") {
          e.preventDefault();
          api.applyStyle({ bold: !api.activeStyle.bold });
          return;
        }
        if (k === "i") {
          e.preventDefault();
          api.applyStyle({ italic: !api.activeStyle.italic });
          return;
        }
        if (k === "u") {
          e.preventDefault();
          api.applyStyle({ underline: !api.activeStyle.underline });
          return;
        }
        // c / v / x fall through to the native copy/paste events below.
        return;
      }

      switch (e.key) {
        case "ArrowUp":
          e.preventDefault();
          if (e.shiftKey) api.extendTo(Math.max(0, selection.r2 - 1), selection.c2);
          else api.move(-1, 0, numRows, numCols);
          return;
        case "ArrowDown":
          e.preventDefault();
          if (e.shiftKey)
            api.extendTo(Math.min(numRows - 1, selection.r2 + 1), selection.c2);
          else api.move(1, 0, numRows, numCols);
          return;
        case "ArrowLeft":
          e.preventDefault();
          if (e.shiftKey) api.extendTo(selection.r2, Math.max(0, selection.c2 - 1));
          else api.move(0, -1, numRows, numCols);
          return;
        case "ArrowRight":
          e.preventDefault();
          if (e.shiftKey)
            api.extendTo(selection.r2, Math.min(numCols - 1, selection.c2 + 1));
          else api.move(0, 1, numRows, numCols);
          return;
        case "Tab":
          e.preventDefault();
          api.move(0, e.shiftKey ? -1 : 1, numRows, numCols);
          return;
        case "Enter":
        case "F2":
          e.preventDefault();
          api.startEdit(sheet, selection.r2, selection.c2);
          return;
        case "Delete":
        case "Backspace":
          e.preventDefault();
          if (!api.clearCells()) {
            api.setStatus("Selection contains locked cells — unlock to clear.");
          }
          return;
        case "Escape":
          api.clearExtraRanges();
          return;
        default:
          break;
      }

      // A printable character starts an edit seeded with that character.
      if (e.key.length === 1 && !e.altKey) {
        e.preventDefault();
        api.startEdit(sheet, selection.r2, selection.c2, e.key);
      }
    },
    [api, sheet, selection],
  );

  const handleImport = useCallback(
    async (file: File) => {
      try {
        const { sheets } = await readWorkbookFile(file);
        const rebuilt: Sheet[] = sheets.map((s) => {
          const base = makeSheet(s.name);
          return {
            ...base,
            cells: s.cells,
            styles: s.styles,
            merges: s.merges,
            numRows: Math.max(base.numRows, s.rows),
            numCols: Math.max(base.numCols, s.cols),
          };
        });
        if (rebuilt.length === 0) throw new Error("no sheets found");
        api.replaceWorkbook({
          sheets: rebuilt,
          activeSheetIndex: 0,
          namedRanges: {},
        });
        api.setStatus(`Imported ${file.name}`);
      } catch (err) {
        api.setStatus(
          `Import failed: ${err instanceof Error ? err.message : String(err)}`,
        );
      }
    },
    [api],
  );

  const childProps = { api, theme, prefix };

  return (
    <div
      className={`${prefix}root`}
      style={{
        fontFamily: theme.fontFamily,
        fontSize: theme.fontSize,
        color: theme.cellText,
        background: theme.cellBg,
        border: `1px solid ${theme.border}`,
        borderRadius: 8,
        display: "flex",
        flexDirection: "column",
        height,
        overflow: "hidden",
        position: "relative",
      }}
    >
      <style>{buildCss(prefix, theme)}</style>

      {showToolbar && (
        <Toolbar
          {...childProps}
          {...(enableFileIO
            ? {
                onImport: handleImport,
                onExportCsv: () =>
                  downloadCsv(sheet.cells, api.evaluator, `${sheet.name}.csv`),
                onExportXlsx: () =>
                  downloadXlsx(
                    api.workbook.sheets.map((s) => ({
                      name: s.name,
                      cells: s.cells,
                      styles: s.styles,
                      merges: s.merges,
                      namedRanges: api.workbook.namedRanges,
                    })),
                  ),
              }
            : {})}
        />
      )}

      {showFormulaBar && <FormulaBar {...childProps} />}

      <textarea
        ref={hiddenRef}
        aria-hidden="true"
        tabIndex={-1}
        value=""
        onChange={() => {}}
        onKeyDown={onKeyDown}
        onCopy={(e) => {
          e.preventDefault();
          e.clipboardData.setData(
            "text/plain",
            api.clipboard.copy(sheet, selection),
          );
        }}
        onCut={(e) => {
          e.preventDefault();
          e.clipboardData.setData(
            "text/plain",
            api.clipboard.copy(sheet, selection),
          );
          api.clearCells();
        }}
        onPaste={(e) => {
          e.preventDefault();
          const text = e.clipboardData.getData("text/plain");
          if (text) {
            api.select(
              api.clipboard.paste(
                text,
                { row: selection.r2, col: selection.c2 },
                api.updateSheet,
              ),
            );
          }
        }}
        style={{
          position: "absolute",
          width: 1,
          height: 1,
          opacity: 0,
          border: 0,
          padding: 0,
          resize: "none",
        }}
      />

      <Grid
        {...childProps}
        renaming={renaming}
        setRenaming={setRenaming}
        onContextMenu={setContextMenu}
        onColumnMenu={setColumnMenu}
      />

      {showSheetTabs && (
        <SheetTabs {...childProps} renaming={renaming} setRenaming={setRenaming} />
      )}
      {showStatusBar && <StatusBar {...childProps} />}

      {contextMenu && (
        <ContextMenu
          {...childProps}
          state={contextMenu}
          onClose={() => setContextMenu(null)}
        />
      )}
      {columnMenu && (
        <ColumnMenu
          {...childProps}
          state={columnMenu}
          onClose={() => setColumnMenu(null)}
        />
      )}
    </div>
  );
}
