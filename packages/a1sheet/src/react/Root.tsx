"use client";

/**
 * `Sheet.Root` — provider, layout container, and owner of keyboard focus.
 *
 * Everything else is a child the consumer chooses to render. There are no
 * `showToolbar`-style props: an absent toolbar is an unrendered `<Sheet.Toolbar/>`.
 *
 * Root owns the hidden 1px, opacity-0 textarea that holds focus whenever a cell is
 * selected but not being edited. It is the target for keydown (navigation, delete,
 * Ctrl+B/I/U/Z/Y, and "type a character to start editing") plus copy, cut, and
 * paste. When a cell IS being edited, focus moves to that cell's input.
 *
 * A new global shortcut goes in `onKeyDown` below, after the `isEditing` guard.
 */
import {
  forwardRef,
  type ReactNode,
  useCallback,
  useEffect,
  useImperativeHandle,
  useMemo,
  useRef,
} from "react";
import type { Workbook } from "../model/types.js";
import { SheetContextProvider, useSheetUiState } from "./context.js";
import { buildCss } from "./styles.js";
import { resolveTheme, type Theme } from "./theme.js";
import { type UseSpreadsheetResult, useSpreadsheet } from "./useSpreadsheet.js";

export interface SheetRootProps {
  /** Uncontrolled initial workbook. */
  defaultWorkbook?: Workbook;
  /** Controlled workbook. Pair with `onWorkbookChange`. */
  workbook?: Workbook;
  onWorkbookChange?: (wb: Workbook) => void;
  theme?: Partial<Theme>;
  /** Class-name prefix for the injected CSS. Defaults to "a1s-". */
  classNamePrefix?: string;
  /** Height of the container. Anything CSS accepts. */
  height?: string | number;
  className?: string;
  style?: React.CSSProperties;
  children?: ReactNode;
}

/** Imperative handle for consumers who need to drive the sheet from outside. */
export interface SheetRootHandle {
  /** Returns keyboard focus to the grid. */
  focus(): void;
  /** The live headless API, for imperative reads and writes. */
  readonly api: UseSpreadsheetResult;
}

export const Root = forwardRef<SheetRootHandle, SheetRootProps>(function Root(
  {
    defaultWorkbook,
    workbook,
    onWorkbookChange,
    theme: themeOverride,
    classNamePrefix = "a1s-",
    height = 600,
    className,
    style,
    children,
  },
  forwardedRef,
) {
  const api = useSpreadsheet({
    ...(defaultWorkbook ? { initialWorkbook: defaultWorkbook } : {}),
    ...(workbook ? { workbook } : {}),
    ...(onWorkbookChange ? { onChange: onWorkbookChange } : {}),
  });

  const theme = resolveTheme(themeOverride);
  const prefix = classNamePrefix;
  const ui = useSheetUiState();
  const focusRef = useRef<HTMLTextAreaElement | null>(null);

  const { sheet, selection, isEditing } = api;

  useImperativeHandle(
    forwardedRef,
    () => ({
      focus: () => focusRef.current?.focus(),
      get api() {
        return api;
      },
    }),
    [api],
  );

  // Reclaim focus whenever no edit or rename is open. activeSheetIndex is a
  // deliberate trigger: switching sheets should return focus to the grid.
  // biome-ignore lint/correctness/useExhaustiveDependencies: intentional trigger
  useEffect(() => {
    if (!isEditing && !ui.renaming) focusRef.current?.focus();
  }, [isEditing, ui.renaming, api.workbook.activeSheetIndex]);

  // Any click outside a menu dismisses it.
  useEffect(() => {
    if (!ui.contextMenu && !ui.columnMenu) return;
    const close = () => ui.closeMenus();
    window.addEventListener("click", close);
    return () => window.removeEventListener("click", close);
  }, [ui]);

  const onKeyDown = useCallback(
    (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
      if (api.isEditing) return;
      const { numRows, numCols } = sheet;

      if (e.ctrlKey || e.metaKey) {
        switch (e.key.toLowerCase()) {
          case "z":
            e.preventDefault();
            api.undo();
            return;
          case "y":
            e.preventDefault();
            api.redo();
            return;
          case "b":
            e.preventDefault();
            api.applyStyle({ bold: !api.activeStyle.bold });
            return;
          case "i":
            e.preventDefault();
            api.applyStyle({ italic: !api.activeStyle.italic });
            return;
          case "u":
            e.preventDefault();
            api.applyStyle({ underline: !api.activeStyle.underline });
            return;
          default:
            // c / v / x fall through to the native clipboard events.
            return;
        }
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
          api.startEdit(sheet, api.active.row, api.active.col);
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
        api.startEdit(sheet, api.active.row, api.active.col, e.key);
      }
    },
    [api, sheet, selection],
  );

  const contextValue = useMemo(
    () => ({ api, theme, prefix, ui, focusRef }),
    [api, theme, prefix, ui],
  );

  return (
    <SheetContextProvider {...contextValue}>
      <div
        className={[`${prefix}root`, className].filter(Boolean).join(" ")}
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
          ...style,
        }}
      >
        <style>{buildCss(prefix, theme)}</style>

        <textarea
          ref={focusRef}
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
            if (!text) return;
            api.select(
              api.clipboard.paste(
                text,
                { row: selection.r2, col: selection.c2 },
                api.updateSheet,
              ),
            );
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

        {children}
      </div>
    </SheetContextProvider>
  );
});
