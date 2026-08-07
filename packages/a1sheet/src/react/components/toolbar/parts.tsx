"use client";

import type { ReactNode } from "react";
import { NUM_FMTS } from "../../../format/numFmt.js";
import type { NumFmt } from "../../../model/types.js";
import { useSheetContext } from "../../context.js";
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
} from "../icons.js";
import { IconButton, type IconButtonProps } from "./IconButton.js";
import { ColorWell } from "./ColorWell.js";
import { FontFamilyMenu } from "./FontFamilyMenu.js";
import { ToolbarSeparator } from "./chrome.js";

const NUM_FMT_LABELS: Record<NumFmt, string> = {
  general: "General",
  integer: "Integer",
  number: "0.00",
  percent: "Percent",
  currency: "Currency",
  date: "Date",
};

type Btn = Omit<IconButtonProps, "children" | "onClick" | "label">;

export function ToolbarUndo(props: Btn = {}): ReactNode {
  const { api } = useSheetContext("Sheet.Toolbar.Undo");
  return (
    <IconButton
      label="Undo (Ctrl+Z)"
      disabled={!api.canUndo}
      onClick={api.undo}
      {...props}
    >
      <UndoIcon />
    </IconButton>
  );
}

export function ToolbarRedo(props: Btn = {}): ReactNode {
  const { api } = useSheetContext("Sheet.Toolbar.Redo");
  return (
    <IconButton
      label="Redo (Ctrl+Y)"
      disabled={!api.canRedo}
      onClick={api.redo}
      {...props}
    >
      <RedoIcon />
    </IconButton>
  );
}

export function ToolbarBold(props: Btn = {}): ReactNode {
  const { api } = useSheetContext("Sheet.Toolbar.Bold");
  const on = api.activeStyle.bold;
  return (
    <IconButton
      label="Bold (Ctrl+B)"
      pressed={on ?? false}
      onClick={() => api.applyStyle({ bold: !on })}
      {...props}
    >
      <BoldIcon />
    </IconButton>
  );
}

export function ToolbarItalic(props: Btn = {}): ReactNode {
  const { api } = useSheetContext("Sheet.Toolbar.Italic");
  const on = api.activeStyle.italic;
  return (
    <IconButton
      label="Italic (Ctrl+I)"
      pressed={on ?? false}
      onClick={() => api.applyStyle({ italic: !on })}
      {...props}
    >
      <ItalicIcon />
    </IconButton>
  );
}

export function ToolbarUnderline(props: Btn = {}): ReactNode {
  const { api } = useSheetContext("Sheet.Toolbar.Underline");
  const on = api.activeStyle.underline;
  return (
    <IconButton
      label="Underline"
      pressed={on ?? false}
      onClick={() => api.applyStyle({ underline: !on })}
      {...props}
    >
      <UnderlineIcon />
    </IconButton>
  );
}

export function ToolbarAlignLeft(props: Btn = {}): ReactNode {
  const { api } = useSheetContext("Sheet.Toolbar.AlignLeft");
  const on = api.activeStyle.align === "left";
  return (
    <IconButton
      label="Align left"
      pressed={on}
      onClick={() => api.applyStyle({ align: "left" })}
      {...props}
    >
      <AlignLeftIcon />
    </IconButton>
  );
}

export function ToolbarAlignCenter(props: Btn = {}): ReactNode {
  const { api } = useSheetContext("Sheet.Toolbar.AlignCenter");
  const on = api.activeStyle.align === "center";
  return (
    <IconButton
      label="Align center"
      pressed={on}
      onClick={() => api.applyStyle({ align: "center" })}
      {...props}
    >
      <AlignCenterIcon />
    </IconButton>
  );
}

export function ToolbarAlignRight(props: Btn = {}): ReactNode {
  const { api } = useSheetContext("Sheet.Toolbar.AlignRight");
  const on = api.activeStyle.align === "right";
  return (
    <IconButton
      label="Align right"
      pressed={on}
      onClick={() => api.applyStyle({ align: "right" })}
      {...props}
    >
      <AlignRightIcon />
    </IconButton>
  );
}

export function ToolbarFontFamily(): ReactNode {
  return <FontFamilyMenu />;
}

export function ToolbarTextColor(): ReactNode {
  const { theme } = useSheetContext("Sheet.Toolbar.TextColor");
  return (
    <ColorWell field="color" label="Text color" fallback={theme.cellText} />
  );
}

export function ToolbarFillColor(): ReactNode {
  const { theme } = useSheetContext("Sheet.Toolbar.FillColor");
  return (
    <ColorWell field="bg" label="Fill color" fallback={theme.cellBg} />
  );
}

export function ToolbarNumFmt(): ReactNode {
  const { api, prefix } = useSheetContext("Sheet.Toolbar.NumFmt");
  const s = api.activeStyle;
  return (
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
  );
}

export function ToolbarLock(props: Btn = {}): ReactNode {
  const { api } = useSheetContext("Sheet.Toolbar.Lock");
  const locked = api.activeStyle.locked;
  return (
    <IconButton
      label={locked ? "Unlock the selection" : "Lock the selection"}
      pressed={locked ?? false}
      onClick={() => api.applyStyle({ locked: !locked })}
      {...props}
    >
      {locked ? <LockIcon /> : <UnlockIcon />}
    </IconButton>
  );
}

export function ToolbarInsertRow(props: Btn = {}): ReactNode {
  const { api } = useSheetContext("Sheet.Toolbar.InsertRow");
  return (
    <IconButton
      label="Insert row"
      onClick={() => api.insertRowAt(api.selection.r2)}
      {...props}
    >
      <InsertRowIcon />
    </IconButton>
  );
}

export function ToolbarDeleteRow(props: Btn = {}): ReactNode {
  const { api } = useSheetContext("Sheet.Toolbar.DeleteRow");
  return (
    <IconButton
      label="Delete row"
      onClick={() => api.deleteRowAt(api.selection.r2)}
      {...props}
    >
      <DeleteRowIcon />
    </IconButton>
  );
}

export function ToolbarInsertCol(props: Btn = {}): ReactNode {
  const { api } = useSheetContext("Sheet.Toolbar.InsertCol");
  return (
    <IconButton
      label="Insert column"
      onClick={() => api.insertColAt(api.selection.c2)}
      {...props}
    >
      <InsertColIcon />
    </IconButton>
  );
}

export function ToolbarDeleteCol(props: Btn = {}): ReactNode {
  const { api } = useSheetContext("Sheet.Toolbar.DeleteCol");
  return (
    <IconButton
      label="Delete column"
      onClick={() => api.deleteColAt(api.selection.c2)}
      {...props}
    >
      <DeleteColIcon />
    </IconButton>
  );
}

export function ToolbarMerge(props: Btn = {}): ReactNode {
  const { api } = useSheetContext("Sheet.Toolbar.Merge");
  return (
    <IconButton label="Merge cells" onClick={api.mergeSelection} {...props}>
      <MergeIcon />
    </IconButton>
  );
}

export function ToolbarUnmerge(props: Btn = {}): ReactNode {
  const { api } = useSheetContext("Sheet.Toolbar.Unmerge");
  return (
    <IconButton label="Unmerge cells" onClick={api.unmergeSelection} {...props}>
      <UnmergeIcon />
    </IconButton>
  );
}

export function ToolbarFreeze(props: Btn = {}): ReactNode {
  const { api } = useSheetContext("Sheet.Toolbar.Freeze");
  return (
    <IconButton
      label="Freeze up through the selection"
      onClick={api.freezeToSelection}
      {...props}
    >
      <FreezeIcon />
    </IconButton>
  );
}

export function ToolbarUnfreeze(props: Btn = {}): ReactNode {
  const { api } = useSheetContext("Sheet.Toolbar.Unfreeze");
  return (
    <IconButton label="Unfreeze" onClick={api.unfreeze} {...props}>
      <UnfreezeIcon />
    </IconButton>
  );
}

/** The default toolbar body — what `<Spreadsheet />` composes. */
export function ToolbarDefaultContent(): ReactNode {
  return (
    <>
      <ToolbarUndo />
      <ToolbarRedo />
      <ToolbarSeparator />
      <ToolbarSeparator />
      <ToolbarFontFamily />
      <ToolbarBold />
      <ToolbarItalic />
      <ToolbarUnderline />
      <ToolbarSeparator />
      <ToolbarAlignLeft />
      <ToolbarAlignCenter />
      <ToolbarAlignRight />
      <ToolbarTextColor />
      <ToolbarFillColor />
      <ToolbarNumFmt />
      <ToolbarLock />
      <ToolbarSeparator />
      <ToolbarInsertRow />
      <ToolbarDeleteRow />
      <ToolbarInsertCol />
      <ToolbarDeleteCol />
      <ToolbarSeparator />
      <ToolbarMerge />
      <ToolbarUnmerge />
      <ToolbarFreeze />
      <ToolbarUnfreeze />
    </>
  );
}
