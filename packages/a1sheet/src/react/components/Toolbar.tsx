"use client";

/**
 * Formatting and structure toolbar.
 *
 * Every button is icon-only and carries a `title` and an `aria-label` with the
 * same words, so the tooltip and the accessible name cannot drift.
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
import {
  AlignCenterIcon,
  AlignLeftIcon,
  AlignRightIcon,
  BoldIcon,
  DeleteColIcon,
  DeleteRowIcon,
  FreezeIcon,
  InsertColIcon,
  InsertRowIcon,
  ItalicIcon,
  LockIcon,
  MergeIcon,
  RedoIcon,
  UnderlineIcon,
  UndoIcon,
  UnfreezeIcon,
  UnlockIcon,
  UnmergeIcon,
} from "./icons.js";

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

/** Both colour wells are the height of a button, so the bar stays one line. */
const COLOR_WELL_SIZE = 28;

export function Toolbar({ children }: ToolbarProps = {}): ReactNode {
  const { api, theme, prefix } = useSheetContext("Sheet.Toolbar");
  const s = api.activeStyle;
  const base = `${prefix}btn ${prefix}iconbtn`;
  const cls = (on?: boolean) => (on ? `${base} ${prefix}on` : base);

  return (
    <div
      style={{
        display: "flex",
        gap: 4,
        alignItems: "center",
        padding: "6px 10px",
        borderBottom: `1px solid ${theme.border}`,
        flexWrap: "wrap",
        background: theme.toolbarBg,
      }}
    >
      <button
        type="button"
        className={base}
        onClick={api.undo}
        disabled={!api.canUndo}
        title="Undo (Ctrl+Z)"
        aria-label="Undo (Ctrl+Z)"
      >
        <UndoIcon />
      </button>
      <button
        type="button"
        className={base}
        onClick={api.redo}
        disabled={!api.canRedo}
        title="Redo (Ctrl+Y)"
        aria-label="Redo (Ctrl+Y)"
      >
        <RedoIcon />
      </button>
      <span className={`${prefix}sep`} />

      <button
        type="button"
        className={cls(s.bold)}
        onClick={() => api.applyStyle({ bold: !s.bold })}
        title="Bold (Ctrl+B)"
        aria-label="Bold (Ctrl+B)"
        aria-pressed={s.bold ?? false}
      >
        <BoldIcon />
      </button>
      <button
        type="button"
        className={cls(s.italic)}
        onClick={() => api.applyStyle({ italic: !s.italic })}
        title="Italic (Ctrl+I)"
        aria-label="Italic (Ctrl+I)"
        aria-pressed={s.italic ?? false}
      >
        <ItalicIcon />
      </button>
      <button
        type="button"
        className={cls(s.underline)}
        onClick={() => api.applyStyle({ underline: !s.underline })}
        title="Underline"
        aria-label="Underline"
        aria-pressed={s.underline ?? false}
      >
        <UnderlineIcon />
      </button>
      <span className={`${prefix}sep`} />

      <button
        type="button"
        className={cls(s.align === "left")}
        onClick={() => api.applyStyle({ align: "left" })}
        title="Align left"
        aria-label="Align left"
        aria-pressed={s.align === "left"}
      >
        <AlignLeftIcon />
      </button>
      <button
        type="button"
        className={cls(s.align === "center")}
        onClick={() => api.applyStyle({ align: "center" })}
        title="Align center"
        aria-label="Align center"
        aria-pressed={s.align === "center"}
      >
        <AlignCenterIcon />
      </button>
      <button
        type="button"
        className={cls(s.align === "right")}
        onClick={() => api.applyStyle({ align: "right" })}
        title="Align right"
        aria-label="Align right"
        aria-pressed={s.align === "right"}
      >
        <AlignRightIcon />
      </button>

      <input
        type="color"
        title="Text color"
        aria-label="Text color"
        value={s.color ?? theme.cellText}
        onChange={(e) => api.applyStyle({ color: e.target.value as `#${string}` })}
        style={{
          width: COLOR_WELL_SIZE,
          height: COLOR_WELL_SIZE,
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
          width: COLOR_WELL_SIZE,
          height: COLOR_WELL_SIZE,
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
        className={cls(s.locked)}
        onClick={() => api.applyStyle({ locked: !s.locked })}
        title={s.locked ? "Unlock the selection" : "Lock the selection"}
        aria-label={s.locked ? "Unlock the selection" : "Lock the selection"}
        aria-pressed={s.locked ?? false}
      >
        {s.locked ? <LockIcon /> : <UnlockIcon />}
      </button>
      <span className={`${prefix}sep`} />

      <button
        type="button"
        className={base}
        onClick={() => api.insertRowAt(api.selection.r2)}
        title="Insert row"
        aria-label="Insert row"
      >
        <InsertRowIcon />
      </button>
      <button
        type="button"
        className={base}
        onClick={() => api.deleteRowAt(api.selection.r2)}
        title="Delete row"
        aria-label="Delete row"
      >
        <DeleteRowIcon />
      </button>
      <button
        type="button"
        className={base}
        onClick={() => api.insertColAt(api.selection.c2)}
        title="Insert column"
        aria-label="Insert column"
      >
        <InsertColIcon />
      </button>
      <button
        type="button"
        className={base}
        onClick={() => api.deleteColAt(api.selection.c2)}
        title="Delete column"
        aria-label="Delete column"
      >
        <DeleteColIcon />
      </button>
      <span className={`${prefix}sep`} />

      <button
        type="button"
        className={base}
        onClick={api.mergeSelection}
        title="Merge cells"
        aria-label="Merge cells"
      >
        <MergeIcon />
      </button>
      <button
        type="button"
        className={base}
        onClick={api.unmergeSelection}
        title="Unmerge cells"
        aria-label="Unmerge cells"
      >
        <UnmergeIcon />
      </button>
      <button
        type="button"
        className={base}
        onClick={api.freezeToSelection}
        title="Freeze up through the selection"
        aria-label="Freeze up through the selection"
      >
        <FreezeIcon />
      </button>
      <button
        type="button"
        className={base}
        onClick={api.unfreeze}
        title="Unfreeze"
        aria-label="Unfreeze"
      >
        <UnfreezeIcon />
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
