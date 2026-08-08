"use client";

/**
 * "Add N more rows at the bottom" — the usual bottom-of-grid control.
 *
 * Rendered as a child of `Sheet.Grid`, so it lands inside the scroll container
 * at the end of the content rather than in a bar below it — you reach it by
 * scrolling to the end of the sheet, which is where someone who has run out of
 * rows is already looking.
 *
 * Its content sticks to the left edge so it stays readable however far right
 * the sheet is scrolled.
 */
import { type ReactNode, useState } from "react";
import { ADD_ROWS_DEFAULT, ADD_ROWS_MAX } from "../constants.js";
import { useSheetContext } from "../context.js";
import { mergeClass } from "../primitives/mergeClass.js";
import type { PrimitiveProps } from "../primitives/types.js";

export interface AddRowsProps extends PrimitiveProps {}

export function AddRows({ className, style }: AddRowsProps = {}): ReactNode {
  const { api, theme, prefix } = useSheetContext("Sheet.AddRows");
  const [value, setValue] = useState(String(ADD_ROWS_DEFAULT));

  const parsed = Number.parseInt(value, 10);
  const count = Number.isFinite(parsed)
    ? Math.min(ADD_ROWS_MAX, Math.max(1, parsed))
    : 0;

  return (
    <div
      className={mergeClass(`${prefix}addrows`, className)}
      style={{
        display: "flex",
        alignItems: "center",
        gap: 8,
        padding: "10px 12px",
        background: theme.headerBg,
        borderTop: `1px solid ${theme.headerBorder}`,
        fontSize: theme.fontSize,
        color: theme.cellText,
        // The grid is wider than the viewport; without this the control drifts
        // off to the left as soon as the sheet is scrolled sideways.
        position: "sticky",
        left: 0,
        width: "100%",
        ...style,
      }}
    >
      <button
        type="button"
        className={`${prefix}btn`}
        disabled={count === 0}
        onClick={() => api.appendRows(count)}
      >
        Add
      </button>
      <input
        className={`${prefix}input`}
        type="number"
        min={1}
        max={ADD_ROWS_MAX}
        aria-label="Number of rows to add"
        value={value}
        onChange={(e) => setValue(e.target.value)}
        onKeyDown={(e) => {
          if (e.key !== "Enter" || count === 0) return;
          e.preventDefault();
          api.appendRows(count);
        }}
        style={{ width: 72 }}
      />
      <span>more rows at the bottom</span>
    </div>
  );
}
