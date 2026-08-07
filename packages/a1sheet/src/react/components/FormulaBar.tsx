"use client";

/**
 * Name box plus formula input.
 *
 * Name box resolution order on Enter:
 *   1. looks like A1 or A1:B2      -> jump the selection there
 *   2. matches an existing name    -> select that range
 *   3. otherwise                   -> define a new named range from the current
 *                                     selection, uppercased. No further validation.
 *
 * The formula input writes through `api.setValue` — the same state the in-cell
 * editor uses — so the two cannot drift.
 */
import { type ReactNode, useState } from "react";
import { normalizeRange, parseRangeRef, toA1 } from "../../model/address.js";
import { useSheetContext } from "../context.js";
import { useCaretBinding } from "../useCaretBinding.js";

export function FormulaBar(): ReactNode {
  const { api, theme, prefix } = useSheetContext("Sheet.FormulaBar");
  const [nameBox, setNameBox] = useState("");
  const { selection, editing, active } = api;
  const caret = useCaretBinding(editing?.caret, api.setCaret);

  // The formula bar always shows the active cell — the anchor — not the moving end
  // of a drag, which would make the bar flicker through every cell you drag over.
  const editingHere = editing?.row === active.row && editing.col === active.col;
  const value = editingHere ? editing.value : api.getRaw(active.row, active.col);

  function commitNameBox() {
    const text = nameBox.trim();
    if (!text) return;

    const asRange = parseRangeRef(text);
    if (asRange) {
      api.select(asRange);
      setNameBox("");
      return;
    }

    const upper = text.toUpperCase();
    const existing = api.workbook.namedRanges[upper];
    if (existing) {
      api.select(existing);
      setNameBox("");
      return;
    }

    api.defineName(upper, normalizeRange(selection));
    api.setStatus(`Defined ${upper}`);
    setNameBox("");
  }

  return (
    <div
      style={{
        display: "flex",
        alignItems: "center",
        gap: 8,
        padding: "5px 10px",
        borderBottom: `1px solid ${theme.border}`,
        background: theme.toolbarBg,
      }}
    >
      <input
        className={`${prefix}input`}
        aria-label="Name box"
        value={nameBox}
        placeholder={toA1(active.row, active.col)}
        onChange={(e) => setNameBox(e.target.value)}
        onKeyDown={(e) => {
          if (e.key === "Enter") commitNameBox();
          if (e.key === "Escape") setNameBox("");
        }}
        style={{ width: 96, fontSize: 12 }}
        title="Type a cell or range to jump to it, or a name to define a named range for the current selection"
      />
      <span style={{ color: theme.headerText, fontStyle: "italic" }}>fx</span>
      <input
        className={`${prefix}input`}
        aria-label="Formula"
        ref={caret.ref}
        value={value}
        onSelect={caret.onSelect}
        onChange={(e) => {
          const at = e.target.selectionStart ?? undefined;
          if (!editingHere) {
            api.startEdit(api.sheet, active.row, active.col, e.target.value);
          } else {
            api.setValue(e.target.value, at);
          }
        }}
        onKeyDown={(e) => {
          if (e.key === "Enter") api.commitEdit([1, 0]);
          if (e.key === "Escape") api.cancel();
        }}
        onBlur={() => {
          if (api.isEditing) api.commitEdit();
        }}
        style={{ flex: 1, fontFamily: theme.monoFontFamily }}
      />
    </div>
  );
}
