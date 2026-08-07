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
import { dataEdge, lastUsedCell, lastUsedInRow } from "../model/navigate.js";
import type { Workbook } from "../model/types.js";
import { HEADER_HEIGHT } from "./constants.js";
import { SheetContextProvider, useSheetUiState } from "./context.js";
import { buildCss } from "./styles.js";
import { resolveTheme, type Theme } from "./theme.js";
import { type UseSpreadsheetResult, useSpreadsheet } from "./useSpreadsheet.js";

/**
 * What a copy says when the selected ranges do not line up into one block.
 * Excel's own wording, because it is the message people already know.
 */
const MULTI_COPY_REFUSED =
  "That command cannot be used on multiple selections — the ranges must share their rows or their columns.";

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
      const { row: activeRow, col: activeCol } = api.active;

      /** Go to an absolute cell, or extend to it when Shift is down. */
      const goTo = (row: number, col: number) => {
        if (e.shiftKey) api.extendTo(row, col);
        else api.moveTo(row, col, numRows, numCols);
      };

      /**
       * How far Page Up and Page Down travel. Measured from the row the cursor
       * is on rather than a constant, so a sheet of tall rows pages by what fits
       * on screen instead of overshooting by a screenful.
       */
      const pageRows = () =>
        Math.max(
          1,
          Math.floor(
            (api.viewportHeight - HEADER_HEIGHT) /
              Math.max(1, api.rowWindow.rowHeight(activeRow)),
          ),
        );

      // Ctrl/Cmd + arrow runs to the edge of the block of data, which is the
      // only practical way to cross a long sheet from the keyboard. Handled
      // before the plain shortcut switch because that one keys on the letter.
      if ((e.ctrlKey || e.metaKey) && e.key.startsWith("Arrow")) {
        const step = {
          ArrowUp: { dRow: -1, dCol: 0 },
          ArrowDown: { dRow: 1, dCol: 0 },
          ArrowLeft: { dRow: 0, dCol: -1 },
          ArrowRight: { dRow: 0, dCol: 1 },
        }[e.key];
        if (!step) return;
        e.preventDefault();
        // Shift+Ctrl+Arrow extends, so it walks from the moving end of the
        // selection; a plain jump walks from the active cell.
        const from = e.shiftKey
          ? { row: selection.r2, col: selection.c2 }
          : { row: activeRow, col: activeCol };
        const to = dataEdge(sheet.cells, from.row, from.col, step, {
          numRows,
          numCols,
        });
        goTo(to.row, to.col);
        return;
      }

      if (e.ctrlKey || e.metaKey) {
        switch (e.key) {
          case "Home":
            e.preventDefault();
            goTo(0, 0);
            return;
          case "End": {
            e.preventDefault();
            const end = lastUsedCell(sheet.cells);
            goTo(end.row, end.col);
            return;
          }
          case " ":
            // Ctrl+Space selects the column, Excel's shortcut for it.
            e.preventDefault();
            api.select({ r1: 0, c1: activeCol, r2: numRows - 1, c2: activeCol });
            return;
          default:
            break;
        }
        switch (e.key.toLowerCase()) {
          case "a":
            e.preventDefault();
            api.select({ r1: 0, c1: 0, r2: numRows - 1, c2: numCols - 1 });
            return;
          case "d":
          case "r": {
            e.preventDefault();
            const direction = e.key.toLowerCase() === "d" ? "down" : "right";
            if (!api.fill.within(selection, direction, api.updateSheet)) {
              api.setStatus(
                `Fill ${direction} needs a filled cell to copy and room to copy it into.`,
              );
            }
            return;
          }
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
        case "PageDown":
          e.preventDefault();
          goTo(
            Math.min(
              numRows - 1,
              (e.shiftKey ? selection.r2 : activeRow) + pageRows(),
            ),
            e.shiftKey ? selection.c2 : activeCol,
          );
          return;
        case "PageUp":
          e.preventDefault();
          goTo(
            Math.max(0, (e.shiftKey ? selection.r2 : activeRow) - pageRows()),
            e.shiftKey ? selection.c2 : activeCol,
          );
          return;
        case "Home":
          e.preventDefault();
          goTo(e.shiftKey ? selection.r2 : activeRow, 0);
          return;
        case "End":
          e.preventDefault();
          goTo(
            e.shiftKey ? selection.r2 : activeRow,
            lastUsedInRow(
              sheet.cells,
              e.shiftKey ? selection.r2 : activeRow,
              numCols,
            ),
          );
          return;
        case " ":
          // Shift+Space selects the row, next to Ctrl+Space for the column.
          if (!e.shiftKey) break;
          e.preventDefault();
          api.select({ r1: activeRow, c1: 0, r2: activeRow, c2: numCols - 1 });
          return;
        case "F8":
          // Excel's add-to-selection toggle. Without it a discontiguous
          // selection needs Ctrl+click, and the keyboard cannot make one at all.
          if (!e.shiftKey) break;
          e.preventDefault();
          api.toggleAddMode();
          api.setStatus(
            api.addMode
              ? ""
              : "Add to selection: move to the next range. Shift+F8 again, or Escape to finish.",
          );
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
          // One key ends all three transient states: the copy outline, the
          // extra ranges, and add-mode.
          api.clipboard.clearCopied();
          api.clearExtraRanges();
          api.setStatus("");
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
            const text = api.clipboard.copy(sheet, api.ranges);
            if (text === null) {
              api.setStatus(MULTI_COPY_REFUSED);
              return;
            }
            e.clipboardData.setData("text/plain", text);
          }}
          onCut={(e) => {
            e.preventDefault();
            const text = api.clipboard.copy(sheet, api.ranges);
            if (text === null) {
              api.setStatus(MULTI_COPY_REFUSED);
              return;
            }
            e.clipboardData.setData("text/plain", text);
            api.clearCells();
          }}
          onPaste={(e) => {
            e.preventDefault();
            const text = e.clipboardData.getData("text/plain");
            if (!text) return;
            // Pasting into several ranges at once has no defined meaning, so
            // the extra ranges are dropped and the paste lands where the
            // active range is — the same thing Excel does.
            api.clearExtraRanges();
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
