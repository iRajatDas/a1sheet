/**
 * `<dataValidation>` — the rules that constrain what a cell may hold.
 *
 * The one that matters in practice is `type="list"`, which is what makes a cell
 * a dropdown. A workbook built on them is not merely less helpful without them:
 * the list IS the interface, and a user faced with a free-text cell has no way to
 * know what the sheet will accept.
 */
import { lettersToCol } from "../../model/address.js";
import type { CondOperator, DataValidation, Range } from "../../model/types.js";
import { findElements, textOf } from "./xml.js";

const REF_RE = /^\$?([A-Za-z]+)\$?(\d+)$/;

function parseRef(ref: string): { row: number; col: number } | null {
  const m = ref.match(REF_RE);
  if (!m?.[1] || !m[2]) return null;
  return {
    col: lettersToCol(m[1].toUpperCase()),
    row: Number.parseInt(m[2], 10) - 1,
  };
}

function parseRange(ref: string): Range | null {
  const [from, to] = ref.split(":");
  if (!from) return null;
  const a = parseRef(from);
  if (!a) return null;
  const b = to ? parseRef(to) : a;
  if (!b) return null;
  return {
    r1: Math.min(a.row, b.row),
    c1: Math.min(a.col, b.col),
    r2: Math.max(a.row, b.row),
    c2: Math.max(a.col, b.col),
  };
}

const KINDS: Record<string, DataValidation["kind"]> = {
  list: "list",
  whole: "whole",
  decimal: "decimal",
  date: "date",
  time: "date",
  textLength: "textLength",
  custom: "custom",
};

const OPERATORS: readonly CondOperator[] = [
  "lessThan",
  "lessThanOrEqual",
  "equal",
  "notEqual",
  "greaterThanOrEqual",
  "greaterThan",
  "between",
  "notBetween",
];

export function parseValidations(sheetXml: string): DataValidation[] {
  const out: DataValidation[] = [];

  for (const el of findElements(sheetXml, "dataValidation")) {
    const kind = KINDS[el.attrs.type ?? ""];
    const sqref = el.attrs.sqref;
    if (!kind || !sqref) continue;

    const formulas = findElements(el.inner, "formula1")
      .concat(findElements(el.inner, "formula2"))
      .map((f) => textOf(f.inner))
      .filter((f) => f !== "");

    const operator = OPERATORS.find((o) => o === el.attrs.operator);
    const message = el.attrs.error;

    // One rule may cover several ranges, space-separated.
    for (const part of sqref.split(/\s+/)) {
      const range = parseRange(part);
      if (!range) continue;
      out.push({
        range,
        kind,
        formulas,
        ...(el.attrs.allowBlank === "1" ? { allowBlank: true } : {}),
        ...(operator ? { operator } : {}),
        ...(message ? { message } : {}),
      });
    }
  }

  return out;
}

/** `A1:C9`, the reference form the file uses. */
function toRef(range: Range): string {
  const col = (c: number) => {
    let letters = "";
    let n = c;
    while (n >= 0) {
      letters = String.fromCharCode((n % 26) + 65) + letters;
      n = Math.floor(n / 26) - 1;
    }
    return letters;
  };
  return `${col(range.c1)}${range.r1 + 1}:${col(range.c2)}${range.r2 + 1}`;
}

function escapeXml(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/** The `<dataValidations>` block for a worksheet, or an empty string. */
export function validationXml(validations: readonly DataValidation[]): string {
  if (validations.length === 0) return "";
  const body = validations
    .map((v) => {
      const formulas = v.formulas
        .map((f, i) => `<formula${i + 1}>${escapeXml(f)}</formula${i + 1}>`)
        .join("");
      return (
        `<dataValidation type="${v.kind}"` +
        `${v.operator ? ` operator="${v.operator}"` : ""}` +
        `${v.allowBlank ? ` allowBlank="1"` : ""}` +
        `${v.message ? ` error="${escapeXml(v.message)}" showErrorMessage="1"` : ""}` +
        ` sqref="${toRef(v.range)}">${formulas}</dataValidation>`
      );
    })
    .join("");
  return `<dataValidations count="${validations.length}">${body}</dataValidations>`;
}
