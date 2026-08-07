"use client";

/**
 * Column header dropdown: sort ascending/descending and a checkbox value filter.
 *
 * Sort physically rewrites cell keys — a data operation, undone only via history.
 * Filter never touches cells; it only feeds `effectiveHiddenRows`, so clearing a
 * filter always restores the original row order and content.
 */
import { type ReactNode, useMemo, useState } from "react";
import { getUsedBounds } from "../../io/csv/write.js";
import { useSheetContext } from "../context.js";

/**
 * Renders nothing until a column header dropdown is opened.
 *
 * The panel is split out and keyed by column so its checkbox state resets when a
 * different column's menu is opened. Keeping the state in this always-mounted outer
 * component would carry one column's selections over to the next.
 */
export function ColumnMenu(): ReactNode {
  const { ui } = useSheetContext("Sheet.ColumnMenu");
  const state = ui.columnMenu;
  if (!state) return null;
  return (
    <ColumnMenuPanel key={state.col} col={state.col} x={state.x} y={state.y} />
  );
}

interface PanelProps {
  col: number;
  x: number;
  y: number;
}

function ColumnMenuPanel({ col, x, y }: PanelProps): ReactNode {
  const { api, theme, prefix, ui } = useSheetContext("Sheet.ColumnMenu");
  const onClose = ui.closeMenus;
  const { sheet } = api;

  /** Distinct displayed values in the column, over the used range only. */
  const values = useMemo(() => {
    const rows = getUsedBounds(sheet.cells).rows;
    const set = new Set<string>();
    for (let r = 0; r < rows; r++) set.add(api.getDisplay(r, col));
    return [...set].sort();
  }, [sheet.cells, col, api.getDisplay]);

  const current = sheet.filters[col];
  const [checked, setChecked] = useState<Set<string>>(
    () => new Set(current ?? values),
  );

  function toggle(v: string) {
    setChecked((prev) => {
      const next = new Set(prev);
      if (next.has(v)) next.delete(v);
      else next.add(v);
      return next;
    });
  }

  return (
    <div
      className={`${prefix}menu`}
      style={{ left: x, top: y, minWidth: 200 }}
      onClick={(e) => e.stopPropagation()}
      onKeyDown={(e) => {
        if (e.key === "Escape") onClose();
      }}
    >
      <button
        type="button"
        onClick={() => {
          api.sort(col, "asc");
          onClose();
        }}
      >
        Sort ascending
      </button>
      <button
        type="button"
        onClick={() => {
          api.sort(col, "desc");
          onClose();
        }}
      >
        Sort descending
      </button>
      <hr />

      <div style={{ maxHeight: 200, overflow: "auto", padding: "0 4px" }}>
        {values.map((v) => (
          <label
            key={v}
            style={{
              display: "flex",
              gap: 6,
              alignItems: "center",
              fontSize: 12,
              padding: "3px 4px",
              cursor: "pointer",
            }}
          >
            <input
              type="checkbox"
              checked={checked.has(v)}
              onChange={() => toggle(v)}
            />
            <span
              style={{
                overflow: "hidden",
                textOverflow: "ellipsis",
                whiteSpace: "nowrap",
                color: v === "" ? theme.headerText : undefined,
              }}
            >
              {v === "" ? "(blank)" : v}
            </span>
          </label>
        ))}
      </div>

      <hr />
      <button
        type="button"
        onClick={() => {
          // Filtering on every value is the same as no filter — store null so the
          // fast path in useRowWindow skips the column entirely.
          api.setFilter(col, checked.size === values.length ? null : checked);
          onClose();
        }}
      >
        Apply filter
      </button>
      <button
        type="button"
        onClick={() => {
          api.setFilter(col, null);
          onClose();
        }}
      >
        Clear filter
      </button>
    </div>
  );
}
