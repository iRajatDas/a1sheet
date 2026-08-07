"use client";

import { type ReactNode, useRef } from "react";
import { useSheetContext } from "../context.js";
import { mergeClass } from "../primitives/mergeClass.js";
import type { PrimitiveProps } from "../primitives/types.js";
import * as parts from "./menu/ContextMenuParts.js";
import { MenuItem, MenuSeparator, useMenuKeyboard } from "./menu/primitives.js";

export interface ContextMenuProps extends PrimitiveProps {
  /** Menu body. Defaults to the standard cell actions. */
  children?: ReactNode;
}

function ContextMenuRoot({
  className,
  style,
  children,
}: ContextMenuProps = {}): ReactNode {
  const { prefix, ui } = useSheetContext("Sheet.ContextMenu");
  const state = ui.contextMenu;
  const onClose = ui.closeMenus;
  const ref = useRef<HTMLDivElement>(null);

  useMenuKeyboard(!!state, onClose, ref);

  if (!state) return null;

  return (
    <div
      ref={ref}
      role="menu"
      className={mergeClass(`${prefix}menu`, className)}
      style={{ left: state.x, top: state.y, ...style }}
      onClick={(e) => e.stopPropagation()}
    >
      {children ?? <parts.ContextMenuDefaultContent />}
    </div>
  );
}

export const ContextMenu = Object.assign(ContextMenuRoot, {
  Copy: parts.ContextMenuCopy,
  Paste: parts.ContextMenuPaste,
  InsertRow: parts.ContextMenuInsertRow,
  DeleteRow: parts.ContextMenuDeleteRow,
  InsertCol: parts.ContextMenuInsertCol,
  DeleteCol: parts.ContextMenuDeleteCol,
  ClearContents: parts.ContextMenuClearContents,
  ClearFormatting: parts.ContextMenuClearFormatting,
  ToggleLock: parts.ContextMenuToggleLock,
  ToggleRowHidden: parts.ContextMenuToggleRowHidden,
  ToggleColHidden: parts.ContextMenuToggleColHidden,
  Separator: MenuSeparator,
  Item: MenuItem,
  Default: parts.ContextMenuDefaultContent,
});
