/**
 * Formula evaluation. Ported from ref/formulaEngine.js:226-358.
 *
 * Design notes:
 *
 * - `createEvaluator` is LAZY. `getCellDisplay` computes on demand, memoizes into
 *   a Map, and detects cycles with a `visiting` Set that yields "#CYCLE!".
 * - A NEW evaluator is created every time `cells` or `namedRanges` changes (a
 *   `useMemo` in the React layer). That is intentional and cheap precisely
 *   because it is lazy. Do not persist or incrementally mutate one.
 * - VLOOKUP, INDEX, and MATCH are special-cased in the "call" branch because they
 *   need the raw range AST node to recover 2D shape. They are not in FUNCTIONS.
 *   A new shape-sensitive function follows that pattern.
 */
import { parseCellRef } from "../model/address.js";
import type { CellKey, CellValue, NamedRanges, RawCell } from "../model/types.js";
import type { CompareOp, Node } from "./ast.js";
import { FUNCTIONS } from "./functions/registry.js";
import { parseFormula } from "./parse.js";
import { tokenize } from "./tokenize.js";
import type { FormulaArg, FormulaValue } from "./values.js";
import { toNumber, toText } from "./values.js";

/** Sentinel returned when a formula participates in a reference cycle. */
export const CYCLE_ERROR = "#CYCLE!";

/**
 * The complete set of error sentinels. Membership is checked exactly rather than
 * by a leading "#" so a user who literally types "#hashtag" is not mistaken for
 * an error value.
 */
export const ERROR_VALUES: ReadonlySet<string> = new Set([
  CYCLE_ERROR,
  "#ERROR!",
  "#DIV/0!",
  "#NAME?",
  "#REF!",
  "#N/A",
  "#VALUE!",
]);

export function isErrorValue(v: unknown): v is string {
  return typeof v === "string" && ERROR_VALUES.has(v);
}

/**
 * Thrown when evaluation re-enters a cell already being evaluated. Unwinds to
 * every frame in the cycle so each participating cell reports "#CYCLE!".
 *
 * Needed because `ctx.getValue` coerces through `toNumber`, which would otherwise
 * turn the sentinel into 0 — the POC's cycle detection therefore only ever caught
 * a DIRECT self-reference, and silently returned 0 for any longer cycle.
 */
class CycleSignal extends Error {
  constructor() {
    super(CYCLE_ERROR);
    this.name = "CycleSignal";
  }
}

export interface EvalContext {
  /** Numeric value of a single cell by position. */
  getValue(row: number, col: number): number;
  /** Flattened row-major values for a range given as two A1 refs. */
  getRange(fromRef: string, toRef: string): FormulaValue[];
  /** Row-major 2D values, for shape-sensitive functions. */
  getRange2D(fromRef: string, toRef: string): FormulaValue[][];
  /** Flattened row-major values for absolute bounds. Used by named ranges. */
  getRangeAbs(r1: number, c1: number, r2: number, c2: number): FormulaValue[];
  namedRanges: NamedRanges;
}

export interface Evaluator {
  /**
   * The computed value for a cell. Returns "" for an empty cell, the raw text for
   * a literal, and an error sentinel ("#CYCLE!", "#ERROR!", "#DIV/0!", "#NAME?",
   * "#REF!", "#N/A") for a formula that cannot be resolved.
   */
  getCellDisplay(row: number, col: number): FormulaValue;
}

/**
 * Comparison with Excel-ish coercion: if both sides parse as numbers, compare
 * numerically; otherwise compare as text. Returns 1 or 0, not a boolean, so the
 * result feeds arithmetic directly.
 */
function evalCompare(op: CompareOp, l: FormulaArg, r: FormulaArg): number {
  const ln = typeof l === "number" ? l : parseFloat(l as string);
  const rn = typeof r === "number" ? r : parseFloat(r as string);
  const bothNum = !Number.isNaN(ln) && !Number.isNaN(rn);
  const a: string | number = bothNum ? ln : toText(l);
  const b: string | number = bothNum ? rn : toText(r);
  switch (op) {
    case "=":
      return a === b ? 1 : 0;
    case "<":
      return a < b ? 1 : 0;
    case ">":
      return a > b ? 1 : 0;
    case "<=":
      return a <= b ? 1 : 0;
    case ">=":
      return a >= b ? 1 : 0;
    case "<>":
      return a !== b ? 1 : 0;
    default:
      return 0;
  }
}

/**
 * Lookup-key equality: numeric comparison when BOTH sides are numeric, text
 * comparison otherwise.
 *
 * The `both` guard is essential. ref/formulaEngine.js:277 used
 * `toNumber(a) === toNumber(b)`, and toNumber coerces non-numeric text to 0 — so
 * every pair of text keys compared equal and VLOOKUP always returned the FIRST
 * row. Text lookups were completely broken.
 */
function looseEquals(a: FormulaValue, b: FormulaValue): boolean {
  const an = parseFloat(a as string);
  const bn = parseFloat(b as string);
  if (!Number.isNaN(an) && !Number.isNaN(bn)) return an === bn;
  return toText(a) === toText(b);
}

export function evalNode(node: Node, ctx: EvalContext): FormulaArg {
  switch (node.type) {
    case "num":
      return node.value;

    case "str":
      return node.value;

    case "neg":
      return -toNumber(evalNode(node.node, ctx));

    case "cmp":
      return evalCompare(
        node.op,
        evalNode(node.left, ctx),
        evalNode(node.right, ctx),
      );

    case "bin": {
      const l = toNumber(evalNode(node.left, ctx));
      const r = toNumber(evalNode(node.right, ctx));
      switch (node.op) {
        case "+":
          return l + r;
        case "-":
          return l - r;
        case "*":
          return l * r;
        case "/":
          return r === 0 ? "#DIV/0!" : l / r;
        case "^":
          return l ** r;
        default:
          return 0;
      }
    }

    case "ref": {
      const { row, col } = parseCellRef(node.value);
      return ctx.getValue(row, col);
    }

    case "range":
      return ctx.getRange(node.from, node.to);

    case "name": {
      const nr = ctx.namedRanges?.[node.value];
      if (nr) return ctx.getRangeAbs(nr.r1, nr.c1, nr.r2, nr.c2);
      return "#NAME?";
    }

    case "call": {
      const name = node.name;

      // --- shape-sensitive functions: need the range AST node, not values ---

      if (name === "VLOOKUP") {
        const first = node.args[0];
        const rangeArg = node.args[1];
        const third = node.args[2];
        if (!first || !rangeArg || rangeArg.type !== "range") return "#REF!";
        const lookupVal = evalNode(first, ctx) as FormulaValue;
        const table = ctx.getRange2D(rangeArg.from, rangeArg.to);
        const colIdx = Math.round(toNumber(third ? evalNode(third, ctx) : 1)) - 1;
        for (const row of table) {
          if (looseEquals(row[0], lookupVal)) {
            return row[colIdx] !== undefined ? row[colIdx] : "#REF!";
          }
        }
        return "#N/A";
      }

      if (name === "MATCH") {
        const first = node.args[0];
        const rangeArg = node.args[1];
        if (!first || !rangeArg) return "#REF!";
        const lookupVal = evalNode(first, ctx) as FormulaValue;
        const arr =
          rangeArg.type === "range"
            ? ctx.getRange(rangeArg.from, rangeArg.to)
            : [evalNode(rangeArg, ctx) as FormulaValue];
        for (let i = 0; i < arr.length; i++) {
          if (looseEquals(arr[i], lookupVal)) return i + 1;
        }
        return "#N/A";
      }

      if (name === "INDEX") {
        const rangeArg = node.args[0];
        const rowArg = node.args[1];
        const colArg = node.args[2];
        if (rangeArg?.type !== "range" || !rowArg) return "#REF!";
        const table = ctx.getRange2D(rangeArg.from, rangeArg.to);
        const rowNum = Math.round(toNumber(evalNode(rowArg, ctx)));
        const colNum = colArg ? Math.round(toNumber(evalNode(colArg, ctx))) : 1;
        const row = table[rowNum - 1];
        if (!row) return "#REF!";
        return row[colNum - 1] !== undefined ? row[colNum - 1] : "#REF!";
      }

      // --- everything else comes from the registry ---

      const fn = FUNCTIONS[name];
      if (!fn) return "#NAME?";
      const argVals: FormulaArg[] = node.args.map((a) =>
        a.type === "range" ? ctx.getRange(a.from, a.to) : evalNode(a, ctx),
      );

      // Propagate errors out of arguments instead of coercing them away. The POC
      // fed them straight into the function, where flattenNums silently dropped
      // them — so `=SUM(UNDEFINED_NAME)` returned 0 rather than #NAME?.
      for (const arg of argVals) {
        if (isErrorValue(arg)) return arg;
      }

      return fn(argVals);
    }

    default:
      return "";
  }
}

/**
 * Values a previous application computed, consulted only where this engine
 * fails. See `Sheet.cachedValues` for the staleness contract.
 */
export type CachedValues = Readonly<Record<CellKey, CellValue>>;

export function createEvaluator(
  cells: Record<CellKey, RawCell>,
  namedRanges: NamedRanges,
  cachedValues: CachedValues = {},
): Evaluator {
  const cache = new Map<string, FormulaValue>();
  const visiting = new Set<string>();

  /**
   * Substitutes the imported value for a formula this engine could not resolve.
   *
   * Only on failure, and never in place of a formula we CAN evaluate: preferring
   * the import would freeze the sheet, since an edit to a cell's inputs would
   * recompute a result nothing ever displays.
   */
  function orImported(key: string, failure: FormulaValue): FormulaValue {
    return cachedValues[key as CellKey] ?? failure;
  }

  /**
   * Internal evaluation. THROWS CycleSignal on a reference cycle so the signal
   * reaches every frame — `getValue` coerces through `toNumber`, which would
   * flatten a returned sentinel to 0 one level up. The public `getCellDisplay`
   * below is the boundary that turns the signal back into a value.
   */
  function compute(row: number, col: number): FormulaValue {
    const key = `${row}_${col}`;

    const cached = cache.get(key);
    if (cached !== undefined) return cached;

    if (visiting.has(key)) {
      cache.set(key, CYCLE_ERROR);
      throw new CycleSignal();
    }

    const raw = cells[key as CellKey];
    if (raw === undefined || raw === "") {
      cache.set(key, "");
      return "";
    }

    if (raw[0] === "=") {
      visiting.add(key);
      let result: FormulaValue;
      try {
        const ast = parseFormula(tokenize(raw.slice(1)));
        const value = evalNode(ast, ctx);
        // A range that leaks into a scalar position collapses to its text form,
        // matching what the POC did implicitly by stringifying the array.
        result = Array.isArray(value) ? toText(value) : value;
        if (typeof result === "number" && !Number.isFinite(result)) {
          result = "#ERROR!";
        }
      } catch (e) {
        visiting.delete(key);
        if (e instanceof CycleSignal) {
          // Mark this cell and keep unwinding, so every cell in the cycle
          // reports #CYCLE! rather than only the one that closed it.
          cache.set(key, CYCLE_ERROR);
          throw e;
        }
        const failed = orImported(key, "#ERROR!");
        cache.set(key, failed);
        return failed;
      }
      visiting.delete(key);
      // An error sentinel means the formula referenced something we do not have —
      // an unimplemented function, a structured reference — which is exactly the
      // case the import's own result covers.
      if (isErrorValue(result)) result = orImported(key, result);
      cache.set(key, result);
      return result;
    }

    cache.set(key, raw);
    return raw;
  }

  function getCellDisplay(row: number, col: number): FormulaValue {
    try {
      return compute(row, col);
    } catch (e) {
      if (e instanceof CycleSignal) return CYCLE_ERROR;
      throw e;
    }
  }

  const ctx: EvalContext = {
    namedRanges,

    getValue: (row, col) => toNumber(compute(row, col)),

    getRangeAbs: (r1, c1, r2, c2) => {
      const out: FormulaValue[] = [];
      for (let r = Math.min(r1, r2); r <= Math.max(r1, r2); r++) {
        for (let c = Math.min(c1, c2); c <= Math.max(c1, c2); c++) {
          out.push(compute(r, c));
        }
      }
      return out;
    },

    getRange: (fromRef, toRef) => {
      const a = parseCellRef(fromRef);
      const b = parseCellRef(toRef);
      return ctx.getRangeAbs(a.row, a.col, b.row, b.col);
    },

    getRange2D: (fromRef, toRef) => {
      const a = parseCellRef(fromRef);
      const b = parseCellRef(toRef);
      const r1 = Math.min(a.row, b.row);
      const r2 = Math.max(a.row, b.row);
      const c1 = Math.min(a.col, b.col);
      const c2 = Math.max(a.col, b.col);
      const rows: FormulaValue[][] = [];
      for (let r = r1; r <= r2; r++) {
        const row: FormulaValue[] = [];
        for (let c = c1; c <= c2; c++) row.push(compute(r, c));
        rows.push(row);
      }
      return rows;
    },
  };

  return { getCellDisplay };
}
