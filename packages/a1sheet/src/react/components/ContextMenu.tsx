"use client";

/**
 * Right-click menu: insert/delete row or column, clear contents, clear
 * formatting, lock/unlock, copy, paste. Ported from the context menu JSX in
 * ref/Spreadsheet.jsx.
 *
 * Its copy and paste items route through `api.clipboard` rather than
 * reimplementing the shift-refs decision, which is what the POC did.
 * `navigator.clipboard` needs permission and can reject, so failures fall back to
 * a status message pointing at Ctrl+C/V.
 */
import type { ReactNode } from "react";
import { useSheetContext } from "../context.js";

/**
 * Renders nothing until a right-click opens it, so the consumer can place it
 * anywhere in the tree — or omit it to disable the menu entirely.
 */
export function ContextMenu(): ReactNode {
  const { api, prefix, ui } = useSheetContext("Sheet.ContextMenu");
  const state = ui.contextMenu;
  const onClose = ui.closeMenus;
  if (!state) return null;
  const { row, col } = state;

  function run(fn: () => void) {
    fn();
    onClose();
  }

  async function copy() {
    const text = api.clipboard.copy(api.sheet, api.selection);
    try {
      await navigator.clipboard.writeText(text);
    } catch {
      api.setStatus("Clipboard blocked — use Ctrl+C instead.");
    }
    onClose();
  }

  async function paste() {
    try {
      const text = await navigator.clipboard.readText();
      api.clipboard.paste(text, { row, col }, api.updateSheet);
    } catch {
      api.setStatus("Clipboard blocked — use Ctrl+V instead.");
    }
    onClose();
  }

  const locked = api.isLocked(row, col);

  return (
    <div
      className={`${prefix}menu`}
      style={{ left: state.x, top: state.y }}
      onClick={(e) => e.stopPropagation()}
      onKeyDown={(e) => {
        if (e.key === "Escape") onClose();
      }}
    >
      <button type="button" onClick={copy}>
        Copy
      </button>
      <button type="button" onClick={paste}>
        Paste
      </button>
      <hr />
      <button type="button" onClick={() => run(() => api.insertRowAt(row))}>
        Insert row above
      </button>
      <button type="button" onClick={() => run(() => api.deleteRowAt(row))}>
        Delete row
      </button>
      <button type="button" onClick={() => run(() => api.insertColAt(col))}>
        Insert column left
      </button>
      <button type="button" onClick={() => run(() => api.deleteColAt(col))}>
        Delete column
      </button>
      <hr />
      <button
        type="button"
        onClick={() =>
          run(() => {
            if (!api.clearCells()) {
              api.setStatus("Selection contains locked cells — unlock to clear.");
            }
          })
        }
      >
        Clear contents
      </button>
      <button type="button" onClick={() => run(api.clearFormatting)}>
        Clear formatting
      </button>
      <hr />
      <button
        type="button"
        onClick={() => run(() => api.applyStyle({ locked: !locked }))}
      >
        {locked ? "Unlock cells" : "Lock cells"}
      </button>
      <button type="button" onClick={() => run(() => api.toggleRowHidden(row))}>
        {api.sheet.hiddenRows.has(row) ? "Unhide row" : "Hide row"}
      </button>
    </div>
  );
}
