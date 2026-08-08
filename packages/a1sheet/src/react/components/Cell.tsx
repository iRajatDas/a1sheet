"use client";

/**
 * A single cell.
 *
 * Appearance resolves in order: base style from `sheet.styles`, then
 * conditional formatting, then the inline style fields. When the cell
 * is being edited it renders an `<input>` that takes focus; otherwise the hidden
 * textarea in `<Spreadsheet />` holds focus for the whole grid.
 */
import { type CSSProperties, type ReactNode, useRef } from "react";
import { cellKey } from "../../model/address.js";
import { getMergeAt } from "../../model/sheet.js";
import type { StyleObject } from "../../model/types.js";
import { cellCss } from "../cellStyle.js";
import { useGridRender, useSheetContext } from "../context.js";
import { useCaretBinding } from "../useCaretBinding.js";
import { CondIcon } from "./CondIcon.js";
import { ChevronDownIcon } from "./icons.js";

export interface CellProps {
  row: number;
  col: number;
  /** CSS grid line this cell renders on, computed by the Grid. */
  gridRow: number;
  /** Sticky offsets for freeze panes, computed by the Grid. */
  stickyStyle: CSSProperties;
  /** Stored cell style — lets the grid pass the resolved object directly. */
  styleRef?: StyleObject;
  /** Live toolbar preview — merged over the stored style for selected cells only. */
  stylePreview?: Partial<StyleObject>;
}

const PERCENT = 100;

export function Cell({
  row,
  col,
  gridRow,
  stickyStyle,
  styleRef,
  stylePreview,
}: CellProps): ReactNode {
  const { api, theme, prefix, ui, focusRef, components } =
    useSheetContext("Sheet.Cell");
  const { renderCellContent: gridRenderCell } = useGridRender();
  const { sheet, selection, editing } = api;

  // Every hook runs before the merge bail-out below — a covered cell returns
  // null, and a conditional hook call is an error.
  const isEditing = editing?.row === row && editing.col === col;
  const caret = useCaretBinding(
    isEditing ? editing.caret : undefined,
    api.setCaret,
  );
  const longPressRef = useRef<ReturnType<typeof setTimeout> | null>(null);

  const merge = getMergeAt(sheet, row, col);
  // A cell covered by a merge is not rendered at all; the merge's top-left cell
  // spans over it.
  if (merge && (merge.r1 !== row || merge.c1 !== col)) return null;

  // A merge spans rows that may each be a different height, so the block is as
  // tall as the rows it covers rather than a multiple of one of them.
  let mergedHeight = api.rowWindow.rowHeight(row);
  if (merge) {
    for (let r = merge.r1 + 1; r <= merge.r2; r++) {
      mergedHeight += api.rowWindow.rowHeight(r);
    }
  }

  const key = cellKey(row, col);
  // Base formatting, then conditional formatting over it. That order is Excel's:
  // a rule that matches wins over the cell's own fill and font, which is what
  // makes a rule visible at all.
  const base = styleRef ?? sheet.styles[key] ?? {};
  const conditional = api.condStyleFor(row, col);
  const merged = stylePreview ? { ...base, ...stylePreview } : base;
  const style = conditional ? { ...merged, ...conditional } : merged;
  // An IMAGE() cell draws its picture instead of its value, which is a URL.
  const image = sheet.images[key];
  // A colour scale, data bar, or icon set — painted rather than styled.
  const decoration = api.condDecorationFor(row, col);
  // A data-validation list makes the cell a dropdown.
  const choices = api.choicesFor(row, col);
  const listId = `${prefix}list-${row}-${col}`;
  const selected = api.isSelected(row, col);
  const active = row === api.active.row && col === api.active.col;

  const gridColumn = col + 2;
  const bounds = api.bounds;
  // Bottom-right of the selection rectangle, wherever the anchor happens to be.
  const showFillHandle =
    selected && !editing && row === bounds.r2 && col === bounds.c2;

  const renderCellContent = gridRenderCell ?? components.CellContent;

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
      role="gridcell"
      aria-selected={selected}
      aria-readonly={style.locked ?? false}
      // The cell's address, in the DOM. Both axes are virtualized, so position
      // among the rendered nodes says nothing about which cell this is — this
      // is the stable way to find one, from a test or from consumer code.
      data-row={row}
      data-col={col}
      style={{
        // Sticky offsets first: everything below is the cell's own appearance and
        // must win. Spreading stickyStyle last used to repaint a frozen cell's
        // custom fill with the default background.
        ...stickyStyle,
        gridColumn: merge
          ? `${gridColumn} / span ${merge.c2 - merge.c1 + 1}`
          : gridColumn,
        gridRow: merge ? `${gridRow} / span ${merge.r2 - merge.r1 + 1}` : gridRow,
        height: mergedHeight,
        ...cellCss(style, theme),
        ...(decoration?.background ? { background: decoration.background } : {}),
      }}
      onMouseDown={(e) => {
        // Clicks inside the open editor are the caret's business, not ours.
        if ((e.target as HTMLElement).tagName === "INPUT") return;

        // Two jobs, both load-bearing. A cell is a plain <div>, so a mousedown on
        // it moves focus to <body> — which silently kills every keyboard shortcut
        // and the copy/cut/paste handlers, since all of them live on the hidden
        // textarea. It also starts a native text selection, so dragging across
        // cells highlights their text instead of selecting a range.
        // preventDefault stops both; the focus() call is what puts the keyboard
        // back, because the default it just suppressed was the thing that moved it.
        e.preventDefault();

        // Mid-formula, a click writes a reference and the editor keeps the
        // caret. `pickAt` returns false when the caret is past a finished
        // operand, which means the click really did mean "select this cell".
        if (api.formulaRefs.pickAt(row, col)) return;

        focusRef.current?.focus();
        if (editing) api.commitEdit();
        if (e.ctrlKey || e.metaKey) {
          // Clicking something already selected takes it back out, as Excel
          // does. Otherwise this is one call, because `addRange` then
          // `selectCell` does not work: selectCell clears the extras, so every
          // Ctrl+click discarded the ranges it had just banked.
          if (!api.removeRangeAt(row, col)) {
            api.startNewRange({ r1: row, c1: col, r2: row, c2: col });
          }
        } else if (e.shiftKey) {
          api.extendTo(row, col);
        } else {
          api.clearExtraRanges();
          api.selectCell(row, col);
        }
      }}
      // No onMouseEnter: extending the selection during a drag is the Grid's
      // job now. Per-cell events cannot see a pointer held outside the grid, so
      // they can neither scroll the sheet nor keep extending once it does.
      onDoubleClick={() => api.startEdit(sheet, row, col)}
      onContextMenu={(e) => {
        e.preventDefault();
        if (!selected) api.selectCell(row, col);
        ui.setContextMenu({ row, col, x: e.clientX, y: e.clientY });
      }}
      onTouchStart={(e) => {
        if (e.touches.length !== 1) return;
        const touch = e.touches[0];
        if (!touch) return;
        longPressRef.current = setTimeout(() => {
          if (!selected) api.selectCell(row, col);
          ui.setContextMenu({ row, col, x: touch.clientX, y: touch.clientY });
        }, 500);
      }}
      onTouchEnd={() => {
        if (longPressRef.current) clearTimeout(longPressRef.current);
      }}
      onTouchMove={() => {
        if (longPressRef.current) clearTimeout(longPressRef.current);
      }}
      title={style.locked ? "Locked cell" : undefined}
    >
      {decoration?.barRatio !== undefined && (
        // Behind the text rather than as a background, so the cell's own fill and
        // the selection tint both still show. Not interactive.
        <div
          aria-hidden="true"
          style={{
            position: "absolute",
            left: 0,
            top: 2,
            bottom: 2,
            width: `${decoration.barRatio * PERCENT}%`,
            background: decoration.barColor,
            opacity: 0.45,
            pointerEvents: "none",
          }}
        />
      )}

      {decoration?.iconIndex !== undefined && (
        <span
          aria-hidden="true"
          style={{ flex: "none", marginRight: 4, lineHeight: 1 }}
        >
          <CondIcon set={decoration.iconSet} index={decoration.iconIndex} />
        </span>
      )}

      {isEditing ? null : image ? (
        // Contained rather than stretched, so a crest keeps its proportions in a
        // cell whose size the user chose for the text beside it. The URL is the
        // only description the file offers, so it is the alt text.
        <img
          src={image.src}
          alt={image.alt ?? ""}
          style={{
            maxWidth: "100%",
            maxHeight: "100%",
            objectFit: "contain",
            display: "block",
          }}
        />
      ) : null}

      {isEditing ? (
        <input
          // biome-ignore lint/a11y/noAutofocus: the editor must take the caret
          autoFocus
          ref={caret.ref}
          value={editing.value}
          onSelect={caret.onSelect}
          // The caret is passed through so typing in the middle of a formula
          // does not get yanked to the end on the next render.
          onChange={(e) =>
            api.setValue(e.target.value, e.target.selectionStart ?? undefined)
          }
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
          // A native datalist, so a validated cell can still be TYPED into —
          // Excel allows both, and a <select> would take that away.
          {...(choices ? { list: listId } : {})}
        />
      ) : image ? null : renderCellContent ? (
        renderCellContent({
          row,
          col,
          display: api.getDisplay(row, col),
          raw: api.getRaw(row, col),
          value: api.getValue(row, col),
          style,
          isEditing,
          isSelected: selected,
          isActive: active,
          isLocked: !!style.locked,
        })
      ) : style.checkbox ? (
        <input
          type="checkbox"
          aria-label="Checkbox"
          checked={api.getDisplay(row, col).toUpperCase() === "TRUE"}
          onMouseDown={(e) => {
            e.stopPropagation();
            e.preventDefault();
            api.toggleCheckbox(row, col);
          }}
          onChange={() => {
            /* toggled in onMouseDown so selection does not steal the click */
          }}
        />
      ) : style.hyperlink ? (
        <a
          href={style.hyperlink}
          target="_blank"
          rel="noopener noreferrer"
          onMouseDown={(e) => e.stopPropagation()}
          style={{ color: "inherit", overflow: "hidden", textOverflow: "ellipsis" }}
        >
          {api.getDisplay(row, col) || style.hyperlink}
        </a>
      ) : (
        api.getDisplay(row, col)
      )}

      {isEditing && choices && (
        <datalist id={listId}>
          {choices.map((choice) => (
            <option key={choice} value={choice} />
          ))}
        </datalist>
      )}

      {choices && !isEditing && active && (
        // The affordance. Only on the active cell: an arrow in every validated
        // cell would be a column of chevrons competing with the values.
        <button
          type="button"
          aria-label="Choose a value"
          className={`${prefix}dropdown`}
          onMouseDown={(e) => {
            e.stopPropagation();
            e.preventDefault();
            api.startEdit(sheet, row, col);
          }}
        >
          <ChevronDownIcon />
        </button>
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
