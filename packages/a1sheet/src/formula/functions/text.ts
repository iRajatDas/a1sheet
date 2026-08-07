/** Text functions. */
import { applyFormatCode } from "../../format/formatCode.js";
import { flatten, toNumber, toScalar, toText } from "../values.js";
import type { FormulaFunction } from "./registry.js";

const concat: FormulaFunction = (a) => a.map((v) => toText(v)).join("");

export const textFunctions: Record<string, FormulaFunction> = {
  CONCAT: concat,
  CONCATENATE: concat,

  LEFT: (a) => toText(a[0]).slice(0, a[1] !== undefined ? toNumber(a[1]) : 1),

  RIGHT: (a) => {
    const s = toText(a[0]);
    const n = a[1] !== undefined ? toNumber(a[1]) : 1;
    return s.slice(Math.max(0, s.length - n));
  },

  /** 1-indexed start, matching Excel. */
  MID: (a) =>
    toText(a[0]).slice(toNumber(a[1]) - 1, toNumber(a[1]) - 1 + toNumber(a[2])),

  /**
   * The URL of an image to draw in the cell.
   *
   * Evaluates to the URL, because a formula's value has to be a scalar — the
   * picture is drawn from `Sheet.images`, which the reader and `setCell` keep in
   * step with the formula. So a cell whose image cannot be fetched still shows
   * where it was meant to come from instead of an error.
   *
   * Excel's further arguments (alt text, sizing, dimensions) are accepted and
   * ignored rather than rejected, so a file using them imports.
   */
  IMAGE: (a) => toText(a[0]),

  TRIM: (a) => toText(a[0]).trim(),
  UPPER: (a) => toText(a[0]).toUpperCase(),
  LOWER: (a) => toText(a[0]).toLowerCase(),
  LEN: (a) => toText(a[0]).length,

  /** TEXTJOIN(delimiter, ignoreEmpty, …). The modern CONCAT. */
  TEXTJOIN: (a) => {
    const delimiter = toText(a[0]);
    const ignoreEmpty = toNumber(a[1]) !== 0;
    const parts = flatten(a.slice(2)).map((v) => toText(v));
    return (ignoreEmpty ? parts.filter((p) => p !== "") : parts).join(delimiter);
  },

  /** TEXTSPLIT(text, colDelimiter, [rowDelimiter]) — a matrix, not a string. */
  TEXTSPLIT: (a) => {
    const text = toText(a[0]);
    const colBy = toText(a[1]);
    const rowBy = a[2] === undefined ? "" : toText(a[2]);
    const rows = rowBy === "" ? [text] : text.split(rowBy);
    return rows.map((row) => (colBy === "" ? [row] : row.split(colBy)));
  },

  SUBSTITUTE: (a) => {
    const text = toText(a[0]);
    const from = toText(a[1]);
    const to = toText(a[2]);
    if (from === "") return text;
    // A fourth argument replaces only that occurrence, 1-based.
    if (a[3] === undefined) return text.split(from).join(to);
    const which = Math.round(toNumber(a[3]));
    let seen = 0;
    return text.split(from).reduce((acc, part, i) => {
      if (i === 0) return part;
      seen++;
      return acc + (seen === which ? to : from) + part;
    }, "");
  },

  REPLACE: (a) => {
    const text = toText(a[0]);
    const start = Math.round(toNumber(a[1])) - 1;
    const count = Math.round(toNumber(a[2]));
    return text.slice(0, start) + toText(a[3]) + text.slice(start + count);
  },

  FIND: (a) => {
    const at = toText(a[1]).indexOf(
      toText(a[0]),
      Math.round(toNumber(a[2] ?? 1)) - 1,
    );
    return at === -1 ? "#VALUE!" : at + 1;
  },

  /** SEARCH is FIND without case sensitivity. */
  SEARCH: (a) => {
    const at = toText(a[1])
      .toLowerCase()
      .indexOf(toText(a[0]).toLowerCase(), Math.round(toNumber(a[2] ?? 1)) - 1);
    return at === -1 ? "#VALUE!" : at + 1;
  },

  REPT: (a) => toText(a[0]).repeat(Math.max(0, Math.round(toNumber(a[1])))),

  /** TEXT(value, formatCode) — the format engine, reached from a formula. */
  TEXT: (a) => {
    const value = toScalar(a[0]);
    const subject = typeof value === "number" ? value : toText(value);
    return applyFormatCode(toText(a[1]), subject) ?? String(subject);
  },

  VALUE: (a) => {
    const n = Number.parseFloat(toText(a[0]));
    return Number.isNaN(n) ? "#VALUE!" : n;
  },

  CHAR: (a) => String.fromCodePoint(Math.round(toNumber(a[0]))),
  CODE: (a) => toText(a[0]).codePointAt(0) ?? 0,
  PROPER: (a) =>
    toText(a[0]).replace(
      /\w\S*/g,
      (w) => w[0]?.toUpperCase() + w.slice(1).toLowerCase(),
    ),
};
