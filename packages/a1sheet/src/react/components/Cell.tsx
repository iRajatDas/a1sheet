"use client";

/**
 * A single cell. Ported from `renderCell` in ref/Spreadsheet.jsx:418-465.
 *
 * Appearance resolves in order: base style from `sheet.styles`, then (once
 * implemented) conditional formatting, then the inline style fields. When the cell
 * is being edited it renders an `<input>` that takes focus; otherwise the hidden
 * textarea in `<Spreadsheet />` holds focus for the whole grid.
 */
import type { CSSProperties, ReactNode } from "react";
import { cellKey } from "../../model/address.js";
import { getMergeAt } from "../../model/sheet.js";
import { ROW_HEIGHT } from "../constants.js";
import { useSheetContext } from "../context.js";

export interface CellProps {
  row: number;
  col: number;
  /** CSS grid line this cell renders on, computed by the Grid. */
  gridRow: number;
  /** Sticky offsets for freeze panes, computed by the Grid. */
  stickyStyle: CSSProperties;
}

export function Cell({ row, col, gridRow, stickyStyle }: CellProps): ReactNode {
  const { api, theme, prefix, ui } = useSheetContext("Sheet.Cell");
  const { sheet, selection, editing } = api;

  const merge = getMergeAt(sheet, row, col);
  // A cell covered by a merge is not rendered at all; the merge's top-left cell
  // spans over it.
  if (merge && (merge.r1 !== row || merge.c1 !== col)) return null;

  const key = cellKey(row, col);
  const style = sheet.styles[key] ?? {};
  const selected = api.isSelected(row, col);
  const active = row === selection.r2 && col === selection.c2;
  const isEditing = editing?.row === row && editing.col === col;

  const gridColumn = col + 2;
  const bounds = api.bounds;
  const showFillHandle =
    active && !editing && row === bounds.r2 && col === bounds.c2;

  const classNames = [
    `${prefix}cell`,
    selected ? `${prefix}selected` : "",
    active ? `${prefix}active` : "",
    style.locked ? `${prefix}locked` : "",
  ]
    .filter(Boolean)
    .join(" ");

  return (
    <div
      className={classNames}
      style={{
        gridColumn: merge
          ? `${gridColumn} / span ${merge.c2 - merge.c1 + 1}`
          : gridColumn,
        gridRow: merge ? `${gridRow} / span ${merge.r2 - merge.r1 + 1}` : gridRow,
        height: merge ? ROW_HEIGHT * (merge.r2 - merge.r1 + 1) : ROW_HEIGHT,
        fontWeight: style.bold ? 700 : 400,
        fontStyle: style.italic ? "italic" : "normal",
        textDecoration: style.underline ? "underline" : "none",
        textAlign: style.align ?? "left",
        justifyContent:
          style.align === "center"
            ? "center"
            : style.align === "right"
              ? "flex-end"
              : "flex-start",
        color: style.color ?? theme.cellText,
        background: isEditing
          ? theme.cellBg
          : (style.bg ?? (selected ? undefined : theme.cellBg)),
        ...stickyStyle,
      }}
      onMouseDown={(e) => {
        if (editing) api.commitEdit();
        if (e.ctrlKey || e.metaKey) {
          api.addRange(selection);
          api.selectCell(row, col);
        } else if (e.shiftKey) {
          api.extendTo(row, col);
        } else {
          api.clearExtraRanges();
          api.selectCell(row, col);
        }
      }}
      onMouseEnter={(e) => {
        // Drag-select. Skipped during a fill drag, which owns its own overlay.
        if (e.buttons === 1 && !api.fill.dragging) api.extendTo(row, col);
      }}
      onDoubleClick={() => api.startEdit(sheet, row, col)}
      onContextMenu={(e) => {
        e.preventDefault();
        if (!selected) api.selectCell(row, col);
        ui.setContextMenu({ row, col, x: e.clientX, y: e.clientY });
      }}
      title={style.locked ? "Locked cell" : undefined}
    >
      {isEditing ? (
        <input
          // biome-ignore lint/a11y/noAutofocus: the editor must take the caret
          autoFocus
          value={editing.value}
          onChange={(e) => api.setValue(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === "Enter") {
              e.preventDefault();
              api.commitEdit([1, 0]);
            } else if (e.key === "Tab") {
              e.preventDefault();
              api.commitEdit([0, e.shiftKey ? -1 : 1]);
            } else if (e.key === "Escape") {
              api.cancel();
            }
          }}
          onBlur={() => api.commitEdit()}
        />
      ) : (
        api.getDisplay(row, col)
      )}

      {showFillHandle && (
        <div
          className={`${prefix}fillhandle`}
          onMouseDown={(e) => {
            e.stopPropagation();
            e.preventDefault();
            api.fill.start(selection);
          }}
        />
      )}
    </div>
  );
}
