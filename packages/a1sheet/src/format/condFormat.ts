/**
 * Deciding which conditional formats apply to a cell.
 *
 * Framework-agnostic on purpose: this is the behaviour, and `Cell` is the thin
 * renderer over it. It needs an evaluator because a rule's test is a formula.
 */
import type { Evaluator } from "../formula/evaluate.js";
import { isErrorValue } from "../formula/evaluate.js";
import { shiftFormulaRefs } from "../formula/refs.js";
import type { FormulaValue } from "../formula/values.js";
import { normalizeRange } from "../model/address.js";
import type { CondFormat, CondRule, StyleObject } from "../model/types.js";

function contains(range: CondFormat["range"], row: number, col: number): boolean {
  const n = normalizeRange(range);
  return row >= n.r1 && row <= n.r2 && col >= n.c1 && col <= n.c2;
}

/** Excel's truthiness for a rule's formula: TRUE, or any nonzero number. */
function isTruthy(value: FormulaValue): boolean {
  if (typeof value === "boolean") return value;
  if (typeof value === "number") return value !== 0;
  if (isErrorValue(value)) return false;
  const text = String(value).toUpperCase();
  if (text === "TRUE") return true;
  if (text === "FALSE") return false;
  const n = Number.parseFloat(text);
  return !Number.isNaN(n) && n !== 0;
}

function compare(
  operator: string,
  value: number | string,
  operands: readonly (number | string)[],
): boolean {
  const [a, b] = operands;
  if (a === undefined) return false;
  switch (operator) {
    case "lessThan":
      return value < a;
    case "lessThanOrEqual":
      return value <= a;
    case "equal":
      return value === a;
    case "notEqual":
      return value !== a;
    case "greaterThanOrEqual":
      return value >= a;
    case "greaterThan":
      return value > a;
    case "between":
      return b !== undefined && value >= a && value <= b;
    case "notBetween":
      return b === undefined || value < a || value > b;
    default:
      return false;
  }
}

/** Numeric where both sides are numeric, textual otherwise. */
function comparable(value: unknown): number | string {
  if (typeof value === "number") return value;
  const text = String(value);
  const n = Number.parseFloat(text);
  return Number.isNaN(n) ? text : n;
}

interface MatchContext {
  evaluator: Evaluator;
  row: number;
  col: number;
  /** Top-left of the rule's range, which relative references are relative to. */
  anchorRow: number;
  anchorCol: number;
}

function matches(rule: CondRule, ctx: MatchContext): boolean {
  const { evaluator, row, col } = ctx;

  if (rule.type === "expression") {
    // A rule's formula is written for the top-left cell of its range and applies
    // to the rest shifted, exactly as a copied formula would be. Absolute refs
    // stay put, which is why `$C$4` tests one cell for the whole range.
    const shifted = shiftFormulaRefs(
      rule.formula,
      row - ctx.anchorRow,
      col - ctx.anchorCol,
    );
    return isTruthy(evaluator.evaluate(shifted));
  }

  const value = evaluator.getCellDisplay(row, col);

  if (rule.type === "containsBlanks") {
    const blank = value === "" || value === undefined;
    return rule.negate ? !blank : blank;
  }

  if (rule.type === "containsText") {
    const has = String(value).includes(rule.text);
    return rule.negate ? !has : has;
  }

  const operands = rule.operands.map((f) => comparable(evaluator.evaluate(f)));
  return compare(rule.operator, comparable(value), operands);
}

export interface CondStyleOptions {
  condFormats: readonly CondFormat[];
  evaluator: Evaluator;
}

/**
 * The combined style from every rule matching a cell, or undefined when none do.
 *
 * Applied lowest priority first so a higher-priority rule's keys overwrite,
 * matching Excel — where priority 1 is the most important. `stopIfTrue` cuts the
 * list off at the first match in priority order.
 */
export function condStyleFor(
  opts: CondStyleOptions,
  row: number,
  col: number,
): StyleObject | undefined {
  const applicable = opts.condFormats.filter((f) => contains(f.range, row, col));
  if (applicable.length === 0) return undefined;

  const byPriority = [...applicable].sort((a, b) => a.priority - b.priority);
  const hits: StyleObject[] = [];
  for (const format of byPriority) {
    const hit = matches(format.rule, {
      evaluator: opts.evaluator,
      row,
      col,
      anchorRow: normalizeRange(format.range).r1,
      anchorCol: normalizeRange(format.range).c1,
    });
    if (!hit) continue;
    hits.push(format.style);
    if (format.stopIfTrue) break;
  }
  if (hits.length === 0) return undefined;

  // Reversed: the last spread wins, and the most important rule must.
  return Object.assign({}, ...hits.reverse());
}
