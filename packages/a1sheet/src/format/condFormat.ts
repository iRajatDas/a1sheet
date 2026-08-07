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
import type {
  CondFormat,
  CondRule,
  CondScalePoint,
  Range,
  StyleObject,
} from "../model/types.js";

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

/**
 * What a graphical rule paints over a cell.
 *
 * Separate from `StyleObject` because it is not formatting: a data bar is drawn
 * behind the text at a width, and an icon is drawn beside it. Only one applies
 * per cell — the highest-priority graphical rule — because two bars in one cell
 * mean nothing.
 */
export interface CondDecoration {
  /** A fill computed from the cell's position between the scale's stops. */
  background?: string;
  /** 0..1 of the cell's width, for a data bar. */
  barRatio?: number;
  barColor?: string;
  /** The icon's band, 0-based from the lowest. */
  iconIndex?: number;
  iconSet?: string;
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

  // A graphical rule is not a style and never "matches"; it is resolved by
  // `condDecorationFor`, which needs the whole range to scale against.
  if (
    rule.type === "colorScale" ||
    rule.type === "dataBar" ||
    rule.type === "iconSet"
  ) {
    return false;
  }

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
/**
 * Where a scale point sits on the range's numbers.
 *
 * `min` and `max` are the range's own extremes, a percentile is over its sorted
 * values, and `num`/`formula` are absolute. Returning NaN for a point that cannot
 * be located lets the caller skip the rule rather than draw something arbitrary.
 */
function pointValue(
  point: CondScalePoint,
  sorted: readonly number[],
  evaluator: Evaluator,
): number {
  const lowest = sorted[0] ?? 0;
  const highest = sorted[sorted.length - 1] ?? 0;
  switch (point.kind) {
    case "min":
      return lowest;
    case "max":
      return highest;
    case "percent": {
      const pct = Number.parseFloat(point.value ?? "");
      return lowest + ((highest - lowest) * pct) / PERCENT;
    }
    case "percentile": {
      const pct = Number.parseFloat(point.value ?? "");
      if (sorted.length === 0) return Number.NaN;
      const at = ((sorted.length - 1) * pct) / PERCENT;
      const low = sorted[Math.floor(at)] ?? 0;
      const high = sorted[Math.ceil(at)] ?? low;
      return low + (high - low) * (at - Math.floor(at));
    }
    case "formula":
      return toNumberOr(evaluator.evaluate(point.value ?? ""), Number.NaN);
    default:
      return Number.parseFloat(point.value ?? "");
  }
}

const PERCENT = 100;

function toNumberOr(value: unknown, fallback: number): number {
  const n = typeof value === "number" ? value : Number.parseFloat(String(value));
  return Number.isNaN(n) ? fallback : n;
}

/** Blends two `#rrggbb` colours. */
function blend(from: string, to: string, ratio: number): string {
  const parse = (hex: string) => [
    Number.parseInt(hex.slice(1, 3), 16),
    Number.parseInt(hex.slice(3, 5), 16),
    Number.parseInt(hex.slice(5, 7), 16),
  ];
  const [r1, g1, b1] = parse(from) as [number, number, number];
  const [r2, g2, b2] = parse(to) as [number, number, number];
  const at = (a: number, b: number) =>
    Math.round(a + (b - a) * Math.max(0, Math.min(1, ratio)))
      .toString(16)
      .padStart(2, "0");
  return `#${at(r1, r2)}${at(g1, g2)}${at(b1, b2)}`;
}

/**
 * The decoration a graphical rule paints on a cell, or undefined.
 *
 * Every such rule scales against the whole range, so the range's numbers are
 * gathered once per call. That is the cost of the feature: a colour scale over a
 * thousand cells reads a thousand values to colour one. The evaluator memoizes
 * per cell, so it is one pass per render rather than one per cell.
 */
export function condDecorationFor(
  opts: CondStyleOptions,
  row: number,
  col: number,
): CondDecoration | undefined {
  const applicable = opts.condFormats
    .filter((f) => contains(f.range, row, col))
    .filter(
      (f) =>
        f.rule.type === "colorScale" ||
        f.rule.type === "dataBar" ||
        f.rule.type === "iconSet",
    )
    .sort((a, b) => a.priority - b.priority);

  const format = applicable[0];
  if (!format) return undefined;

  const values = rangeNumbers(format.range, opts.evaluator);
  const own = toNumberOr(opts.evaluator.getCellDisplay(row, col), Number.NaN);
  if (Number.isNaN(own) || values.length === 0) return undefined;

  const rule = format.rule;

  if (rule.type === "colorScale") {
    const points = rule.stops.map((stop) =>
      pointValue(stop, values, opts.evaluator),
    );
    // Between which pair of stops does this cell fall?
    for (let i = 1; i < points.length; i++) {
      const low = points[i - 1] as number;
      const high = points[i] as number;
      if (own > high && i < points.length - 1) continue;
      const span = high - low;
      const ratio = span === 0 ? 0 : (own - low) / span;
      return {
        background: blend(
          rule.stops[i - 1]?.color as string,
          rule.stops[i]?.color as string,
          ratio,
        ),
      };
    }
    return undefined;
  }

  if (rule.type === "dataBar") {
    const low = pointValue(rule.min, values, opts.evaluator);
    const high = pointValue(rule.max, values, opts.evaluator);
    const span = high - low;
    return {
      barRatio: span === 0 ? 0 : Math.max(0, Math.min(1, (own - low) / span)),
      barColor: rule.color,
    };
  }

  if (rule.type !== "iconSet") return undefined;

  // The highest threshold the value clears.
  let band = 0;
  for (const [i, threshold] of rule.thresholds.entries()) {
    if (own >= pointValue(threshold, values, opts.evaluator)) band = i;
  }
  return { iconIndex: band, iconSet: rule.set };
}

/** Every numeric value in a range, sorted — what a scale is measured against. */
function rangeNumbers(range: Range, evaluator: Evaluator): number[] {
  const n = normalizeRange(range);
  const out: number[] = [];
  for (let r = n.r1; r <= n.r2; r++) {
    for (let c = n.c1; c <= n.c2; c++) {
      const value = evaluator.getCellDisplay(r, c);
      const num =
        typeof value === "number" ? value : Number.parseFloat(String(value));
      if (!Number.isNaN(num)) out.push(num);
    }
  }
  return out.sort((a, b) => a - b);
}

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
