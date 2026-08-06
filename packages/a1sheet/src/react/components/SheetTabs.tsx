"use client";

/**
 * Sheet tab strip: switch, add, delete, and rename (double-click) sheets.
 * Ported from the sheet tab JSX in ref/Spreadsheet.jsx.
 *
 * Renaming shares the `RenamingState` shape with the column and row headers.
 */
import type { ReactNode } from "react";
import { useSheetContext } from "../context.js";
import { PlusIcon, TrashIcon } from "./icons.js";

export function SheetTabs(): ReactNode {
  const { api, theme, prefix, ui } = useSheetContext("Sheet.Tabs");
  const { renaming, setRenaming } = ui;
  const { workbook } = api;

  return (
    <div
      style={{
        display: "flex",
        alignItems: "flex-end",
        gap: 2,
        padding: "4px 8px 0",
        borderTop: `1px solid ${theme.border}`,
        background: theme.headerBg,
      }}
    >
      {workbook.sheets.map((s, i) => {
        const isRenaming = renaming?.type === "sheet" && renaming.index === i;
        const isActive = i === workbook.activeSheetIndex;
        return (
          <div
            key={s.id}
            className={`${prefix}tab${isActive ? ` ${prefix}on` : ""}`}
            onMouseDown={() =>
              api.updateWorkbook((wb) => ({ ...wb, activeSheetIndex: i }), false)
            }
            onDoubleClick={() =>
              setRenaming({ type: "sheet", index: i, value: s.name })
            }
          >
            {isRenaming ? (
              <input
                // biome-ignore lint/a11y/noAutofocus: inline rename must take the caret
                autoFocus
                value={renaming.value}
                style={{ width: 90 }}
                onChange={(e) =>
                  setRenaming({ ...renaming, value: e.target.value })
                }
                onKeyDown={(e) => {
                  if (e.key === "Enter") {
                    api.renameSheetAt(i, renaming.value);
                    setRenaming(null);
                  }
                  if (e.key === "Escape") setRenaming(null);
                }}
                onBlur={() => setRenaming(null)}
              />
            ) : (
              s.name
            )}
          </div>
        );
      })}

      <button
        type="button"
        className={`${prefix}btn ${prefix}iconbtn`}
        title="Add sheet"
        aria-label="Add sheet"
        style={{ marginLeft: 6, marginBottom: 4 }}
        onClick={() => api.addSheetAt()}
      >
        <PlusIcon />
      </button>
      {workbook.sheets.length > 1 && (
        <button
          type="button"
          className={`${prefix}btn ${prefix}iconbtn`}
          title="Delete the active sheet"
          aria-label="Delete the active sheet"
          style={{ marginBottom: 4 }}
          onClick={() => api.deleteSheetAt(workbook.activeSheetIndex)}
        >
          <TrashIcon />
        </button>
      )}
    </div>
  );
}
