/**
 * Display formatting. Ported from ref/Spreadsheet.jsx:64-78.
 *
 * Applied to the *evaluated* value on its way to the screen or to CSV — never to
 * the raw cell content, which stays exactly as the user typed it.
 */

import { DAY_MS } from "../formula/values.js";
import type { NumFmt, StyleObject } from "../model/types.js";

export const NUM_FMTS: NumFmt[] = [
  "general",
  "integer",
  "number",
  "percent",
  "currency",
  "date",
];

export interface FormatOptions {
  /** Passed to toLocaleString. Defaults to the runtime locale. */
  locale?: string;
  /** ISO 4217 code used by the "currency" format. Defaults to USD. */
  currency?: string;
}

export function formatValue(
  raw: string | number | undefined,
  style?: StyleObject,
  opts: FormatOptions = {},
): string {
  if (raw === undefined || raw === "") return "";

  const numFmt = style?.numFmt;
  if (!numFmt || numFmt === "general") return String(raw);

  const n = typeof raw === "number" ? raw : parseFloat(raw);
  if (Number.isNaN(n)) return String(raw);

  const locale = opts.locale;
  const currency = opts.currency ?? "USD";

  switch (numFmt) {
    case "integer":
      return Math.round(n).toLocaleString(locale);
    case "number":
      return n.toLocaleString(locale, {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2,
      });
    case "percent":
      return `${(n * 100).toFixed(2)}%`;
    case "currency":
      return n.toLocaleString(locale, { style: "currency", currency });
    case "date": {
      const d = new Date(n * DAY_MS);
      return Number.isNaN(d.getTime())
        ? String(raw)
        : (d.toISOString().slice(0, 10) as string);
    }
    default:
      return String(raw);
  }
}
