"use client";

import { type ReactNode, useEffect, useRef } from "react";
import type { StyleObject } from "../../../model/types.js";
import { useSheetContext } from "../../context.js";
import { useStylePreview } from "../../stylePreview.js";

type ColorField = "color" | "bg";

export interface ColorWellProps {
  field: ColorField;
  label: string;
  fallback: string;
}

/**
 * Native color input — live preview via `StylePreviewProvider`, one workbook
 * commit when the picker closes (`change`). Avoids rebuilding the grid on every
 * hue tick while the OS color dialog is open.
 */
export function ColorWell({ field, label, fallback }: ColorWellProps): ReactNode {
  const { api, prefix } = useSheetContext("Sheet.Toolbar.ColorWell");
  const { patch, setPreview, clearPreview } = useStylePreview();
  const committed = api.activeStyle[field] ?? fallback;
  const preview = patch[field];
  const value = typeof preview === "string" ? preview : committed;
  const anchor = `${api.active.row}_${api.active.col}`;
  const open = useRef(false);

  // biome-ignore lint/correctness/useExhaustiveDependencies: reset when active cell changes
  useEffect(() => {
    clearPreview();
    open.current = false;
  }, [anchor, clearPreview]);

  return (
    <input
      type="color"
      className={`${prefix}colorwell`}
      title={label}
      aria-label={label}
      value={value}
      onFocus={() => {
        open.current = true;
      }}
      onInput={(e) => {
        if (!open.current) return;
        setPreview({ [field]: e.currentTarget.value } as Partial<StyleObject>);
      }}
      onChange={(e) => {
        const hex = e.currentTarget.value as `#${string}`;
        clearPreview();
        open.current = false;
        api.applyStyle({ [field]: hex } as Partial<StyleObject>);
      }}
      onBlur={() => {
        clearPreview();
        open.current = false;
      }}
    />
  );
}
