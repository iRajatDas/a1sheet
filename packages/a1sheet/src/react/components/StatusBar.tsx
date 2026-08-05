/**
 * Aggregate stats for the current selection. Ported from
 * ref/Spreadsheet.jsx:699-714.
 *
 * This is the ONLY consumer of `extraRanges` (Ctrl+click multi-select). Copy,
 * fill, and paste all ignore it and act on the primary selection.
 */
import type { ReactNode } from "react";
import { normalizeRange, toA1 } from "../../model/address.js";
import type { Range } from "../../model/types.js";
import type { BaseProps } from "./props.js";

export function StatusBar({ api, theme, prefix }: BaseProps): ReactNode {
  const ranges: Range[] = [api.selection, ...api.extraRanges];

  let count = 0;
  let numericCount = 0;
  let sum = 0;
  const seen = new Set<string>();

  for (const range of ranges) {
    const b = normalizeRange(range);
    for (let r = b.r1; r <= b.r2; r++) {
      for (let c = b.c1; c <= b.c2; c++) {
        // Ctrl+click ranges can overlap; count each cell once.
        const key = `${r}_${c}`;
        if (seen.has(key)) continue;
        seen.add(key);
        const v = api.getValue(r, c);
        if (v === "" || v === undefined) continue;
        count++;
        const n = typeof v === "number" ? v : parseFloat(String(v));
        if (!Number.isNaN(n)) {
          numericCount++;
          sum += n;
        }
      }
    }
  }

  const b = normalizeRange(api.selection);
  const rangeLabel =
    b.r1 === b.r2 && b.c1 === b.c2
      ? toA1(b.r1, b.c1)
      : `${toA1(b.r1, b.c1)}:${toA1(b.r2, b.c2)}`;

  const fmt = (n: number) =>
    Number.isInteger(n) ? String(n) : n.toFixed(2).replace(/\.?0+$/, "");

  return (
    <div
      className={`${prefix}status`}
      style={{
        display: "flex",
        gap: 16,
        alignItems: "center",
        padding: "4px 10px",
        borderTop: `1px solid ${theme.border}`,
        background: theme.headerBg,
        fontSize: 12,
        color: theme.headerText,
      }}
    >
      <span>{rangeLabel}</span>
      {api.extraRanges.length > 0 && (
        <span>{api.extraRanges.length + 1} ranges</span>
      )}
      <span>Count: {count}</span>
      {numericCount > 0 && (
        <>
          <span>Sum: {fmt(sum)}</span>
          <span>Avg: {fmt(sum / numericCount)}</span>
        </>
      )}
    </div>
  );
}
