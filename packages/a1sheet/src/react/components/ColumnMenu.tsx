/**
 * Column header dropdown: sort ascending/descending and a checkbox value filter.
 * Ported from `ColumnMenu` in ref/Spreadsheet.jsx:715+.
 *
 * Sort physically rewrites cell keys — a data operation, undone only via history.
 * Filter never touches cells; it only feeds `effectiveHiddenRows`, so clearing a
 * filter always restores the original row order and content.
 */
import { type ReactNode, useMemo, useState } from "react";
import { getUsedBounds } from "../../io/csv/write.js";
import type { BaseProps, ColumnMenuState } from "./props.js";

export interface ColumnMenuProps extends BaseProps {
  state: ColumnMenuState;
  onClose(): void;
}

export function ColumnMenu({
  api,
  theme,
  prefix,
  state,
  onClose,
}: ColumnMenuProps): ReactNode {
  const { col } = state;
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
      style={{ left: state.x, top: state.y, minWidth: 200 }}
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
