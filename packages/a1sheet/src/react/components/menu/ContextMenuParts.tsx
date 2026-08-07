"use client";

import type { ReactNode } from "react";
import { useSheetContext } from "../../context.js";
import { MenuItem, MenuSeparator } from "./primitives.js";

function useContextTarget() {
  const { api, ui } = useSheetContext("Sheet.ContextMenu");
  const state = ui.contextMenu;
  if (!state) return null;
  const { row, col } = state;
  const onClose = ui.closeMenus;

  const run = (fn: () => void) => {
    fn();
    onClose();
  };

  return { api, row, col, run, onClose };
}

export function ContextMenuCopy(): ReactNode {
  const ctx = useContextTarget();
  if (!ctx) return null;
  const { api, onClose } = ctx;

  return (
    <MenuItem
      onSelect={() => {
        void (async () => {
          const text = api.clipboard.copy(api.sheet, api.ranges);
          if (text === null) {
            api.setStatus(
              "That command cannot be used on multiple selections — the ranges must share their rows or their columns.",
            );
            onClose();
            return;
          }
          try {
            await navigator.clipboard.writeText(text);
          } catch {
            api.setStatus("Clipboard blocked — use Ctrl+C instead.");
          }
          onClose();
        })();
      }}
    >
      Copy
    </MenuItem>
  );
}

export function ContextMenuPaste(): ReactNode {
  const ctx = useContextTarget();
  if (!ctx) return null;
  const { api, row, col, onClose } = ctx;

  return (
    <MenuItem
      onSelect={() => {
        void (async () => {
          try {
            const text = await navigator.clipboard.readText();
            api.clipboard.paste(text, { row, col }, api.updateSheet, {
        evaluator: api.evaluator,
        onReject: api.setStatus,
        selection: api.selection,
      });
          } catch {
            api.setStatus("Clipboard blocked — use Ctrl+V instead.");
          }
          onClose();
        })();
      }}
    >
      Paste
    </MenuItem>
  );
}

export function ContextMenuInsertRow(): ReactNode {
  const ctx = useContextTarget();
  if (!ctx) return null;
  const { api, row, run } = ctx;
  return (
    <MenuItem onSelect={() => run(() => api.insertRowAt(row))}>
      Insert row above
    </MenuItem>
  );
}

export function ContextMenuDeleteRow(): ReactNode {
  const ctx = useContextTarget();
  if (!ctx) return null;
  const { api, row, run } = ctx;
  return (
    <MenuItem onSelect={() => run(() => api.deleteRowAt(row))}>
      Delete row
    </MenuItem>
  );
}

export function ContextMenuInsertCol(): ReactNode {
  const ctx = useContextTarget();
  if (!ctx) return null;
  const { api, col, run } = ctx;
  return (
    <MenuItem onSelect={() => run(() => api.insertColAt(col))}>
      Insert column left
    </MenuItem>
  );
}

export function ContextMenuDeleteCol(): ReactNode {
  const ctx = useContextTarget();
  if (!ctx) return null;
  const { api, col, run } = ctx;
  return (
    <MenuItem onSelect={() => run(() => api.deleteColAt(col))}>
      Delete column
    </MenuItem>
  );
}

export function ContextMenuClearContents(): ReactNode {
  const ctx = useContextTarget();
  if (!ctx) return null;
  const { api, run } = ctx;
  return (
    <MenuItem
      onSelect={() =>
        run(() => {
          if (!api.clearCells()) {
            api.setStatus("Selection contains locked cells — unlock to clear.");
          }
        })
      }
    >
      Clear contents
    </MenuItem>
  );
}

export function ContextMenuClearFormatting(): ReactNode {
  const ctx = useContextTarget();
  if (!ctx) return null;
  const { api, run } = ctx;
  return (
    <MenuItem onSelect={() => run(api.clearFormatting)}>Clear formatting</MenuItem>
  );
}

export function ContextMenuToggleLock(): ReactNode {
  const ctx = useContextTarget();
  if (!ctx) return null;
  const { api, row, col, run } = ctx;
  const locked = api.isLocked(row, col);
  return (
    <MenuItem onSelect={() => run(() => api.applyStyle({ locked: !locked }))}>
      {locked ? "Unlock cells" : "Lock cells"}
    </MenuItem>
  );
}

export function ContextMenuToggleRowHidden(): ReactNode {
  const ctx = useContextTarget();
  if (!ctx) return null;
  const { api, row, run } = ctx;
  return (
    <MenuItem onSelect={() => run(() => api.toggleRowHidden(row))}>
      {api.sheet.hiddenRows.has(row) ? "Unhide row" : "Hide row"}
    </MenuItem>
  );
}

export function ContextMenuToggleColHidden(): ReactNode {
  const ctx = useContextTarget();
  if (!ctx) return null;
  const { api, col, run } = ctx;
  return (
    <MenuItem onSelect={() => run(() => api.toggleColHidden(col))}>
      {api.sheet.hiddenCols.has(col) ? "Unhide column" : "Hide column"}
    </MenuItem>
  );
}

export function ContextMenuDefaultContent(): ReactNode {
  return (
    <>
      <ContextMenuCopy />
      <ContextMenuPaste />
      <MenuSeparator />
      <ContextMenuInsertRow />
      <ContextMenuDeleteRow />
      <ContextMenuInsertCol />
      <ContextMenuDeleteCol />
      <MenuSeparator />
      <ContextMenuClearContents />
      <ContextMenuClearFormatting />
      <MenuSeparator />
      <ContextMenuToggleLock />
      <ContextMenuToggleRowHidden />
      <ContextMenuToggleColHidden />
    </>
  );
}
