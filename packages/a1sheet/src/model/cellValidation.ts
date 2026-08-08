/**
 * Whether a raw cell value satisfies a data-validation rule on that cell.
 */

import type { Evaluator } from "../formula/evaluate.js";
import { normalizeRange } from "./address.js";
import type { DataValidation, Sheet } from "./types.js";
import { listLiterals } from "./validation.js";

export interface ValidationRejection {
  message: string;
}

function validationFor(
  sheet: Sheet,
  row: number,
  col: number,
): DataValidation | undefined {
  for (const rule of sheet.validations) {
    const r = normalizeRange(rule.range);
    if (row >= r.r1 && row <= r.r2 && col >= r.c1 && col <= r.c2) return rule;
  }
  return undefined;
}

function compareNumber(
  value: number,
  operator: DataValidation["operator"],
  bound: number,
): boolean {
  switch (operator) {
    case "greaterThan":
      return value > bound;
    case "greaterThanOrEqual":
      return value >= bound;
    case "lessThan":
      return value < bound;
    case "lessThanOrEqual":
      return value <= bound;
    case "equal":
      return value === bound;
    case "notEqual":
      return value !== bound;
    case "between": {
      // between uses two formula operands — handled by caller
      return true;
    }
    default:
      return true;
  }
}

/**
 * Returns a rejection message when `raw` breaks the cell's rule, or undefined
 * when the value is accepted. List rules are enforced; other kinds are checked
 * when their operands parse as plain numbers or lengths.
 */
export function rejectCellValue(
  sheet: Sheet,
  row: number,
  col: number,
  raw: string,
  evaluator: Evaluator,
): ValidationRejection | undefined {
  const rule = validationFor(sheet, row, col);
  if (!rule) return undefined;
  if (raw === "" && rule.allowBlank) return undefined;

  switch (rule.kind) {
    case "list": {
      const literals = listLiterals(rule);
      const choices =
        literals ??
        (() => {
          const source = rule.formulas[0];
          if (!source) return null;
          const value = evaluator.evaluateArray(source.replace(/^=/, ""));
          if (!Array.isArray(value)) return null;
          return value.flat().map((v) => (v === undefined ? "" : String(v)));
        })();
      if (!choices) return undefined;
      if (!choices.includes(raw)) {
        return {
          message: rule.message ?? "That value is not in the allowed list.",
        };
      }
      return undefined;
    }
    case "whole":
    case "decimal": {
      const n = Number(raw);
      if (Number.isNaN(n)) {
        return { message: rule.message ?? "That value must be a number." };
      }
      const bound = Number(rule.formulas[0]);
      if (Number.isNaN(bound)) return undefined;
      if (!compareNumber(n, rule.operator ?? "equal", bound)) {
        return { message: rule.message ?? "That value is not allowed here." };
      }
      return undefined;
    }
    case "textLength": {
      const len = raw.length;
      const bound = Number(rule.formulas[0]);
      if (Number.isNaN(bound)) return undefined;
      if (!compareNumber(len, rule.operator ?? "equal", bound)) {
        return { message: rule.message ?? "Text length is not allowed here." };
      }
      return undefined;
    }
    default:
      return undefined;
  }
}
