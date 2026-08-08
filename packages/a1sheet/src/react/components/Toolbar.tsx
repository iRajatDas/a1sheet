"use client";

/**
 * Formatting and structure toolbar shell.
 *
 * Compose atoms from `Sheet.Toolbar.*`, or omit children to render the default
 * set (undo through freeze, without file I/O — add `<Sheet.FileMenu />` yourself).
 *
 *   <Sheet.Toolbar>
 *     <Sheet.Toolbar.Undo asChild>
 *       <Button variant="ghost" size="icon"><UndoIcon /></Button>
 *     </Sheet.Toolbar.Undo>
 *   </Sheet.Toolbar>
 */
import type { ReactNode } from "react";
import { useSheetContext } from "../context.js";
import { mergeClass } from "../primitives/mergeClass.js";
import type { PrimitiveProps } from "../primitives/types.js";
import * as toolbar from "./toolbar/index.js";

export interface ToolbarProps extends PrimitiveProps {
  /**
   * Toolbar body. When omitted, renders the default formatting controls
   * (not including `<Sheet.FileMenu />` — that stays an explicit child).
   */
  children?: ReactNode;
}

function ToolbarRoot({ children, className, style }: ToolbarProps = {}): ReactNode {
  const { theme, prefix } = useSheetContext("Sheet.Toolbar");
  const body = children ?? <toolbar.ToolbarDefaultContent />;

  return (
    <div
      role="toolbar"
      className={mergeClass(`${prefix}toolbar`, className)}
      style={{
        display: "flex",
        gap: 4,
        alignItems: "center",
        padding: "6px 10px",
        borderBottom: `1px solid ${theme.border}`,
        flexWrap: "wrap",
        background: theme.toolbarBg,
        ...style,
      }}
    >
      {body}
    </div>
  );
}

export const Toolbar = Object.assign(ToolbarRoot, {
  Undo: toolbar.ToolbarUndo,
  Redo: toolbar.ToolbarRedo,
  Bold: toolbar.ToolbarBold,
  Italic: toolbar.ToolbarItalic,
  Underline: toolbar.ToolbarUnderline,
  AlignLeft: toolbar.ToolbarAlignLeft,
  AlignCenter: toolbar.ToolbarAlignCenter,
  AlignRight: toolbar.ToolbarAlignRight,
  FontFamily: toolbar.ToolbarFontFamily,
  TextColor: toolbar.ToolbarTextColor,
  FillColor: toolbar.ToolbarFillColor,
  NumFmt: toolbar.ToolbarNumFmt,
  Lock: toolbar.ToolbarLock,
  InsertRow: toolbar.ToolbarInsertRow,
  DeleteRow: toolbar.ToolbarDeleteRow,
  InsertCol: toolbar.ToolbarInsertCol,
  DeleteCol: toolbar.ToolbarDeleteCol,
  Merge: toolbar.ToolbarMerge,
  Unmerge: toolbar.ToolbarUnmerge,
  Freeze: toolbar.ToolbarFreeze,
  Unfreeze: toolbar.ToolbarUnfreeze,
  Separator: toolbar.ToolbarSeparator,
  Group: toolbar.ToolbarGroup,
  Status: toolbar.ToolbarStatus,
  Overflow: toolbar.ToolbarOverflow,
  IconButton: toolbar.IconButton,
  Default: toolbar.ToolbarDefaultContent,
});
