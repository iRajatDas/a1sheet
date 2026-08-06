"use client";

/**
 * Formatting and structure toolbar. Ported from ref/Spreadsheet.jsx:509-542.
 *
 * Freeze freezes up through the current selection. Anything else you want in
 * the bar is a child — `<Sheet.FileMenu />`, or your own buttons using the
 * `${prefix}btn` class:
 *
 *   <Sheet.Toolbar>
 *     <Sheet.FileMenu />
 *     <button className="a1s-btn" onClick={save}>Save</button>
 *   </Sheet.Toolbar>
 */
import type { ReactNode } from "react";
import { NUM_FMTS } from "../../format/numFmt.js";
import type { NumFmt } from "../../model/types.js";
import { useSheetContext } from "../context.js";

export interface ToolbarProps {
  /** Rendered at the end of the bar, after a separator. */
  children?: ReactNode;
}

const NUM_FMT_LABELS: Record<NumFmt, string> = {
  general: "General",
  integer: "Integer",
  number: "0.00",
  percent: "Percent",
  currency: "Currency",
  date: "Date",
};

export function Toolbar({ children }: ToolbarProps = {}): ReactNode {
  const { api, theme, prefix } = useSheetContext("Sheet.Toolbar");
  const s = api.activeStyle;
  const btn = (on?: boolean) => `${prefix}btn${on ? ` ${prefix}on` : ""}`;

  return (
    <div
      style={{
        display: "flex",
        gap: 6,
        alignItems: "center",
        padding: "6px 10px",
        borderBottom: `1px solid ${theme.border}`,
        flexWrap: "wrap",
        background: theme.toolbarBg,
      }}
    >
      <button
        type="button"
        className={btn()}
        onClick={api.undo}
        disabled={!api.canUndo}
        title="Undo (Ctrl+Z)"
      >
        ↶
      </button>
      <button
        type="button"
        className={btn()}
        onClick={api.redo}
        disabled={!api.canRedo}
        title="Redo (Ctrl+Y)"
      >
        ↷
      </button>
      <span className={`${prefix}sep`} />

      <button
        type="button"
        className={btn(s.bold)}
        onClick={() => api.applyStyle({ bold: !s.bold })}
        title="Bold (Ctrl+B)"
      >
        <b>B</b>
      </button>
      <button
        type="button"
        className={btn(s.italic)}
        onClick={() => api.applyStyle({ italic: !s.italic })}
        title="Italic (Ctrl+I)"
      >
        <i>I</i>
      </button>
      <button
        type="button"
        className={btn(s.underline)}
        onClick={() => api.applyStyle({ underline: !s.underline })}
        title="Underline"
      >
        <u>U</u>
      </button>
      <button
        type="button"
        className={btn(s.align === "left")}
        onClick={() => api.applyStyle({ align: "left" })}
        title="Align left"
      >
        ⯇
      </button>
      <button
        type="button"
        className={btn(s.align === "center")}
        onClick={() => api.applyStyle({ align: "center" })}
        title="Align center"
      >
        ≡
      </button>
      <button
        type="button"
        className={btn(s.align === "right")}
        onClick={() => api.applyStyle({ align: "right" })}
        title="Align right"
      >
        ⯈
      </button>

      <input
        type="color"
        title="Text color"
        aria-label="Text color"
        value={s.color ?? theme.cellText}
        onChange={(e) => api.applyStyle({ color: e.target.value as `#${string}` })}
        style={{
          width: 28,
          height: 28,
          padding: 0,
          border: `1px solid ${theme.buttonBorder}`,
          borderRadius: 6,
        }}
      />
      <input
        type="color"
        title="Fill color"
        aria-label="Fill color"
        value={s.bg ?? "#ffffff"}
        onChange={(e) => api.applyStyle({ bg: e.target.value as `#${string}` })}
        style={{
          width: 28,
          height: 28,
          padding: 0,
          border: `1px solid ${theme.buttonBorder}`,
          borderRadius: 6,
        }}
      />

      <select
        className={`${prefix}btn`}
        aria-label="Number format"
        value={s.numFmt ?? "general"}
        onChange={(e) => api.applyStyle({ numFmt: e.target.value as NumFmt })}
      >
        {NUM_FMTS.map((f) => (
          <option key={f} value={f}>
            {NUM_FMT_LABELS[f]}
          </option>
        ))}
      </select>

      <button
        type="button"
        className={btn(s.locked)}
        onClick={() => api.applyStyle({ locked: !s.locked })}
        title="Lock or unlock the selection"
      >
        🔒
      </button>
      <span className={`${prefix}sep`} />

      <button
        type="button"
        className={btn()}
        onClick={() => api.insertRowAt(api.selection.r2)}
      >
        +Row
      </button>
      <button
        type="button"
        className={btn()}
        onClick={() => api.deleteRowAt(api.selection.r2)}
      >
        −Row
      </button>
      <button
        type="button"
        className={btn()}
        onClick={() => api.insertColAt(api.selection.c2)}
      >
        +Col
      </button>
      <button
        type="button"
        className={btn()}
        onClick={() => api.deleteColAt(api.selection.c2)}
      >
        −Col
      </button>
      <button type="button" className={btn()} onClick={api.mergeSelection}>
        Merge
      </button>
      <button type="button" className={btn()} onClick={api.unmergeSelection}>
        Unmerge
      </button>
      <button
        type="button"
        className={btn()}
        onClick={api.freezeToSelection}
        title="Freeze up through the selection"
      >
        Freeze
      </button>
      <button type="button" className={btn()} onClick={api.unfreeze}>
        Unfreeze
      </button>

      {children && (
        <>
          <span className={`${prefix}sep`} />
          {children}
        </>
      )}

      {api.status && (
        <span style={{ fontSize: 12, color: theme.headerText }}>{api.status}</span>
      )}
    </div>
  );
}
