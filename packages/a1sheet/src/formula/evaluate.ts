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
import type {
  CellKey,
  CellValue,
  NamedRanges,
  Range,
  RawCell,
} from "../model/types.js";
import type { BinaryOp, CompareOp, Node } from "./ast.js";
import { broadcast1, broadcast2 } from "./broadcast.js";
import { FUNCTIONS } from "./functions/registry.js";
import { parseFormula } from "./parse.js";
import {
  EMPTY_TABLE_INDEX,
  resolveSelector,
  type TableIndex,
} from "./tableRefs.js";
import { tokenize } from "./tokenize.js";
import type {
  FormulaArg,
  FormulaValue,
  LambdaValue,
  Matrix,
  Scope,
} from "./values.js";
import {
  at,
  colCount,
  isLambda,
  isMatrix,
  rowCount,
  toNumber,
  toScalar,
  toText,
} from "./values.js";

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
  /** Numeric value of a single cell on the current sheet. */
  getValue(row: number, col: number): number;
  /**
   * A single cell. `sheet` names another sheet, from a qualified reference;
   * `undefined` means the one the formula is on. An unknown name is `#REF!`.
   */
  getCell(row: number, col: number, sheet?: string): FormulaValue;
  /** Row-major values for a range given as two A1 refs. */
  getRange(fromRef: string, toRef: string, sheet?: string): Matrix;
  /** Row-major values for absolute bounds. Used by named ranges and tables. */
  getRangeAbs(
    r1: number,
    c1: number,
    r2: number,
    c2: number,
    sheet?: string,
  ): Matrix;
  /** Evaluates a workbook-level defined name that holds a formula, not a range. */
  getNamedFormula(name: string): FormulaArg | undefined;
  namedRanges: NamedRanges;
  /**
   * Names bound by an enclosing `LET` or `LAMBDA`, checked before named ranges.
   * A fresh context is derived for each scope rather than mutated, so a lambda
   * captures the bindings it was defined under.
   */
  scope: Scope;
  /** Named tables, for structured references like `tbl[column]`. */
  tables: TableIndex;
  /**
   * The row the formula being evaluated lives on, for `[#This Row]`. A formula
   * belonging to no cell — a conditional-format rule — reports -1, which no
   * this-row reference can resolve against.
   */
  currentRow: number;
}

export interface Evaluator {
  /**
   * The computed value for a cell. Returns "" for an empty cell, the raw text for
   * a literal, and an error sentinel ("#CYCLE!", "#ERROR!", "#DIV/0!", "#NAME?",
   * "#REF!", "#N/A") for a formula that cannot be resolved.
   */
  getCellDisplay(row: number, col: number): FormulaValue;
  /**
   * The whole value of a cell, which for a formula returning an array is the
   * array. `getCellDisplay` gives the single value shown IN the anchor cell; this
   * gives what the formula actually produced, which is what spilling needs.
   */
  getCellValue(row: number, col: number): FormulaArg;
  /**
   * Evaluates a formula body — WITHOUT the leading "=" — that belongs to no cell.
   * Conditional-format rules are the first such thing: their formulas are
   * attached to a range rather than living in a cell.
   *
   * Returns an error sentinel rather than throwing, on the same terms as
   * `getCellDisplay`.
   */
  evaluate(formulaBody: string): FormulaValue;
  /**
   * The same, keeping an array result whole rather than collapsing it to its
   * top-left. `evaluate` is what a conditional-format rule wants; this is what a
   * caller inspecting a dynamic array wants.
   */
  evaluateArray(formulaBody: string): FormulaArg;
}

/**
 * Comparison with Excel-ish coercion: if both sides parse as numbers, compare
 * numerically; otherwise compare as text. Returns 1 or 0, not a boolean, so the
 * result feeds arithmetic directly.
 */
/**
 * An error on either side of an operator is the operator's result.
 *
 * Excel propagates errors through arithmetic and comparison, and it matters more
 * than it looks: without it, `IF(COUNTBLANK(bad)=0, 1, 0)` compared `#NAME?`
 * against 0, found them unequal, and quietly returned 0 — an error turned into a
 * plausible answer by the guard that was supposed to be checking for one.
 */
function errorIn(a: FormulaValue, b?: FormulaValue): string | null {
  if (isErrorValue(a)) return a;
  if (b !== undefined && isErrorValue(b)) return b;
  return null;
}

function compareScalars(
  op: CompareOp,
  l: FormulaValue,
  r: FormulaValue,
): FormulaValue {
  const failed = errorIn(l, r);
  if (failed) return failed;
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
/** Names that are values rather than references. */
const NAME_CONSTANTS: Record<string, FormulaArg> = {
  TRUE: true,
  FALSE: false,
};

function applyBinary(op: BinaryOp, a: FormulaValue, b: FormulaValue): FormulaValue {
  const failed = errorIn(a, b);
  if (failed) return failed;
  // `&` is text concatenation and must not go through toNumber.
  if (op === "&") return toText(a) + toText(b);
  const l = toNumber(a);
  const r = toNumber(b);
  switch (op) {
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

export function looseEquals(a: FormulaValue, b: FormulaValue): boolean {
  const an = parseFloat(a as string);
  const bn = parseFloat(b as string);
  if (!Number.isNaN(an) && !Number.isNaN(bn)) return an === bn;
  return toText(a) === toText(b);
}

/**
 * Functions whose arguments must NOT all be evaluated before dispatch.
 *
 * `LET` and `LAMBDA` bind names, so their arguments have to be evaluated in a
 * scope that does not exist yet. `IF`, `IFS`, `SWITCH`, and `IFERROR` must not
 * evaluate the branch they do not take — for `IFERROR` that is the entire point,
 * and for `IF` it is the difference between a guard working and dividing by zero
 * anyway.
 */
const LAZY_FORMS = new Set([
  "LET",
  "LAMBDA",
  "IF",
  "IFS",
  "SWITCH",
  "IFERROR",
  "IFNA",
]);

/**
 * Functions that build a reference rather than read one.
 *
 * They cannot be ordinary registry functions because a function receives values,
 * and these need to produce a region for the evaluator to read — `OFFSET(A1,1,0)`
 * means "the cell below A1", not "the value in A1, moved".
 */
const REFERENCE_FORMS = new Set(["OFFSET", "INDIRECT"]);

function evalReferenceForm(
  name: string,
  args: readonly Node[],
  ctx: EvalContext,
): FormulaArg {
  if (name === "INDIRECT") {
    // The address is computed, which is exactly what makes INDIRECT volatile in
    // Excel and why nothing here caches on it.
    const first = args[0];
    if (!first) return "#VALUE!";
    const text = toText(evalNode(first, ctx));
    const [sheet, address] = splitQualified(text);
    const range = parseA1Range(address);
    if (!range) return "#REF!";
    return ctx.getRangeAbs(range.r1, range.c1, range.r2, range.c2, sheet);
  }

  // OFFSET(reference, rows, cols, [height], [width]).
  const anchor = args[0];
  if (!anchor) return "#VALUE!";
  const base = rangeOf(anchor);
  if (!base) return "#REF!";

  const dRow = Math.round(toNumber(evalNode(args[1] as Node, ctx)));
  const dCol = args[2] ? Math.round(toNumber(evalNode(args[2], ctx))) : 0;
  const height = args[3]
    ? Math.round(toNumber(evalNode(args[3], ctx)))
    : base.r2 - base.r1 + 1;
  const width = args[4]
    ? Math.round(toNumber(evalNode(args[4], ctx)))
    : base.c2 - base.c1 + 1;

  const r1 = base.r1 + dRow;
  const c1 = base.c1 + dCol;
  if (r1 < 0 || c1 < 0 || height < 1 || width < 1) return "#REF!";
  return ctx.getRangeAbs(r1, c1, r1 + height - 1, c1 + width - 1, base.sheet);
}

/** The region an AST node names, for the functions that take a reference. */
function rangeOf(node: Node): (Range & { sheet?: string | undefined }) | null {
  if (node.type === "ref") {
    const { row, col } = parseCellRef(node.value);
    return { r1: row, c1: col, r2: row, c2: col, sheet: node.sheet };
  }
  if (node.type === "range") {
    const a = parseCellRef(node.from);
    const b = parseCellRef(node.to);
    return {
      r1: Math.min(a.row, b.row),
      c1: Math.min(a.col, b.col),
      r2: Math.max(a.row, b.row),
      c2: Math.max(a.col, b.col),
      sheet: node.sheet,
    };
  }
  return null;
}

const QUALIFIED = /^(?:'((?:[^']|'')*)'|([A-Za-z_][A-Za-z0-9_. ]*))!(.*)$/;

/** Splits `Sheet2!A1:B9` into its sheet name and its address. */
function splitQualified(text: string): [string | undefined, string] {
  const m = text.match(QUALIFIED);
  if (!m) return [undefined, text];
  const sheet = (m[1] ?? m[2]) as string;
  return [sheet.replace(/''/g, "'"), m[3] as string];
}

const A1_RANGE =
  /^\$?([A-Za-z]{1,3})\$?(\d{1,7})(?::\$?([A-Za-z]{1,3})\$?(\d{1,7}))?$/;

function parseA1Range(text: string): Range | null {
  const m = text.trim().match(A1_RANGE);
  if (!m) return null;
  const from = parseCellRef(`${m[1]}${m[2]}`);
  const to = m[3] ? parseCellRef(`${m[3]}${m[4]}`) : from;
  return {
    r1: Math.min(from.row, to.row),
    c1: Math.min(from.col, to.col),
    r2: Math.max(from.row, to.row),
    c2: Math.max(from.col, to.col),
  };
}

function evalCall(
  node: { type: "call"; name: string; args: Node[] },
  ctx: EvalContext,
): FormulaArg {
  const name = node.name;

  if (LAZY_FORMS.has(name)) return evalLazy(name, node.args, ctx);
  if (REFERENCE_FORMS.has(name)) return evalReferenceForm(name, node.args, ctx);

  const fn = FUNCTIONS[name];
  if (!fn) return "#NAME?";

  const argVals = node.args.map((a) => evalNode(a, ctx));

  // Propagate errors out of arguments instead of coercing them away. The POC fed
  // them straight into the function, where flattenNums silently dropped them —
  // so `=SUM(UNDEFINED_NAME)` returned 0 rather than #NAME?. A matrix argument is
  // NOT scanned: a range holding one error value does not poison SUM in Excel.
  for (const arg of argVals) {
    if (isErrorValue(arg)) return arg;
  }

  return fn(argVals, { call: (lambda, args) => applyLambda(lambda, args, ctx) });
}

/**
 * Applies a lambda to arguments, in the scope it closed over.
 *
 * The lambda's own scope is the parent, not the caller's — that is what makes a
 * lambda returned from a `LET` still see the `LET`'s bindings.
 */
function applyLambda(
  lambda: LambdaValue,
  args: readonly FormulaArg[],
  ctx: EvalContext,
): FormulaArg {
  const scope = new Map(lambda.scope);
  for (const [i, param] of lambda.params.entries()) {
    scope.set(param, args[i]);
  }
  return evalNode(lambda.body, { ...ctx, scope });
}

/** How deep a chain of nested LET/LAMBDA scopes may go before we call it a loop. */
const MAX_SCOPE_DEPTH = 64;

function evalLazy(
  name: string,
  args: readonly Node[],
  ctx: EvalContext,
): FormulaArg {
  if (ctx.scope.size > MAX_SCOPE_DEPTH) return "#ERROR!";

  if (name === "LAMBDA") {
    // Every argument but the last is a parameter name; the last is the body.
    const body = args[args.length - 1];
    if (!body) return "#VALUE!";
    const params: string[] = [];
    for (const param of args.slice(0, -1)) {
      if (param.type !== "name") return "#VALUE!";
      params.push(param.value);
    }
    return { kind: "lambda", params, body, scope: ctx.scope };
  }

  if (name === "LET") {
    // name, value, name, value, …, result. Each value sees the bindings before
    // it, which is what lets a LET build on itself.
    const scope = new Map(ctx.scope);
    let i = 0;
    while (i + 2 < args.length) {
      const nameNode = args[i];
      const valueNode = args[i + 1];
      if (!nameNode || !valueNode || nameNode.type !== "name") return "#VALUE!";
      scope.set(nameNode.value, evalNode(valueNode, { ...ctx, scope }));
      i += 2;
    }
    const result = args[i];
    return result ? evalNode(result, { ...ctx, scope }) : "#VALUE!";
  }

  if (name === "IF") {
    const condition = args[0];
    if (!condition) return "#VALUE!";
    const test = evalNode(condition, ctx);
    if (isErrorValue(test)) return test;
    // An array condition makes the whole IF an array operation.
    if (isMatrix(test)) {
      return broadcast1(test, (v) =>
        toScalar(evalIfBranch(toNumber(v) ? args[1] : args[2], ctx)),
      );
    }
    return evalIfBranch(toNumber(test) ? args[1] : args[2], ctx);
  }

  if (name === "IFS") {
    // Condition/value pairs, first match wins. No match is #N/A, as in Excel.
    for (let i = 0; i + 1 < args.length; i += 2) {
      const condition = args[i] as Node;
      const test = evalNode(condition, ctx);
      if (isErrorValue(test)) return test;
      if (toNumber(test)) return evalNode(args[i + 1] as Node, ctx);
    }
    return "#N/A";
  }

  if (name === "SWITCH") {
    const subjectNode = args[0];
    if (!subjectNode) return "#VALUE!";
    const subject = toScalar(evalNode(subjectNode, ctx));
    let i = 1;
    for (; i + 1 < args.length; i += 2) {
      const candidate = toScalar(evalNode(args[i] as Node, ctx));
      if (looseEquals(candidate, subject)) {
        return evalNode(args[i + 1] as Node, ctx);
      }
    }
    // A trailing odd argument is the default.
    return i < args.length ? evalNode(args[i] as Node, ctx) : "#N/A";
  }

  // IFERROR and IFNA: evaluate the guarded expression, and only on the error it
  // catches evaluate the fallback.
  const guarded = args[0];
  if (!guarded) return "#VALUE!";
  const value = evalNode(guarded, ctx);
  const caught = name === "IFNA" ? value === "#N/A" : isErrorValue(toScalar(value));
  if (!caught) return value;
  const fallback = args[1];
  return fallback ? evalNode(fallback, ctx) : "";
}

function evalIfBranch(branch: Node | undefined, ctx: EvalContext): FormulaArg {
  // Excel's omitted branch is FALSE, not blank.
  return branch ? evalNode(branch, ctx) : false;
}

/** The range a structured reference denotes, as values. */
function resolveTableRef(
  node: { table: string; spec: string },
  ctx: EvalContext,
): FormulaArg {
  const table = ctx.tables.get(node.table);
  if (!table) return "#NAME?";
  const range = resolveSelector({
    table,
    spec: node.spec,
    currentRow: ctx.currentRow,
  });
  if (!range) return "#REF!";
  // A table's cells are on the sheet that defines it, which need not be the sheet
  // the formula is on: a workbook-level defined name is evaluated wherever it is
  // used, and its tables stay where they are.
  return ctx.getRangeAbs(range.r1, range.c1, range.r2, range.c2, table.sheet);
}

export function evalNode(node: Node, ctx: EvalContext): FormulaArg {
  switch (node.type) {
    case "num":
      return node.value;

    case "str":
      return node.value;

    // Every operator broadcasts. `(teams=team) * played` over two columns is a
    // column of products, and taking the top-left of each side instead would
    // return one number where the formula meant many.
    case "neg":
      return broadcast1(evalNode(node.node, ctx), (v) =>
        isErrorValue(v) ? v : -toNumber(v),
      );

    case "cmp": {
      const left = evalNode(node.left, ctx);
      const right = evalNode(node.right, ctx);
      return broadcast2(left, right, (a, b) => compareScalars(node.op, a, b));
    }

    case "bin": {
      const left = evalNode(node.left, ctx);
      const right = evalNode(node.right, ctx);
      return broadcast2(left, right, (a, b) => applyBinary(node.op, a, b));
    }

    case "ref": {
      const { row, col } = parseCellRef(node.value);
      return ctx.getCell(row, col, node.sheet);
    }

    case "range":
      return ctx.getRange(node.from, node.to, node.sheet);

    case "tableRef":
      return resolveTableRef(node, ctx);

    case "name": {
      // A LET or LAMBDA binding shadows a workbook name, as it does in Excel.
      const bound = ctx.scope.get(node.value);
      if (bound !== undefined) return bound;
      const nr = ctx.namedRanges?.[node.value];
      if (nr) return ctx.getRangeAbs(nr.r1, nr.c1, nr.r2, nr.c2);
      // A defined name may hold a formula rather than a range — that is how a
      // modern workbook names a computed table.
      const computed = ctx.getNamedFormula(node.value);
      if (computed !== undefined) return computed;
      const constant = NAME_CONSTANTS[node.value];
      if (constant !== undefined) return constant;
      return "#NAME?";
    }

    case "call":
      return evalCall(node, ctx);

    case "arr":
      return node.rows as Matrix;

    default:
      return "";
  }
}

/**
 * Values a previous application computed, consulted only where this engine
 * fails. See `Sheet.cachedValues` for the staleness contract.
 */
export type CachedValues = Readonly<Record<CellKey, CellValue>>;

/** One other sheet a qualified reference can reach. */
export interface SheetCells {
  name: string;
  cells: Record<CellKey, RawCell>;
}

export interface EvaluatorOptions {
  /** Values a previous application computed, for formulas we cannot evaluate. */
  cachedValues?: CachedValues;
  /** Named tables, for structured references. */
  tables?: TableIndex;
  /**
   * The other sheets in the workbook, so `Sheet2!A1` resolves. Omit for a
   * single-sheet evaluator, where a qualified reference is `#REF!`.
   */
  sheets?: readonly SheetCells[];
  /**
   * Workbook defined names that hold a formula rather than a range — how a modern
   * workbook names a computed table. Values are formula bodies, without the `=`.
   */
  namedFormulas?: Readonly<Record<string, string>>;
  /**
   * Regions an array formula declares as its output, keyed by the anchor.
   *
   * An imported sheet has the array's result written into it as ordinary values,
   * because that is how Excel lets other readers see one. Those cells belong to
   * the anchor: they must not block its spill, and they must give way to a
   * recomputed result rather than shadowing it with a stale one.
   */
  spillRanges?: Readonly<Record<CellKey, Range>>;
}

export function createEvaluator(
  cells: Record<CellKey, RawCell>,
  namedRanges: NamedRanges,
  opts: EvaluatorOptions = {},
): Evaluator {
  const cachedValues = opts.cachedValues ?? {};
  const tables = opts.tables ?? EMPTY_TABLE_INDEX;
  const namedFormulas = opts.namedFormulas ?? {};
  const declaredSpills = Object.entries(opts.spillRanges ?? {}) as [
    CellKey,
    Range,
  ][];

  /** The anchor whose declared output covers this cell, if any but itself. */
  function anchorOver(row: number, col: number): [number, number] | null {
    for (const [key, range] of declaredSpills) {
      const separator = key.indexOf("_");
      const anchorRow = Number(key.slice(0, separator));
      const anchorCol = Number(key.slice(separator + 1));
      if (anchorRow === row && anchorCol === col) continue;
      if (
        row >= range.r1 &&
        row <= range.r2 &&
        col >= range.c1 &&
        col <= range.c2
      ) {
        return [anchorRow, anchorCol];
      }
    }
    return null;
  }

  /**
   * Cells by sheet name, for qualified references. The sheet the evaluator was
   * built for is the unnamed default and is NOT in here — it is `cells`.
   */
  const otherSheets = new Map<string, Record<CellKey, RawCell>>();
  for (const sheet of opts.sheets ?? []) {
    otherSheets.set(sheet.name.toLowerCase(), sheet.cells);
  }

  /** Which cell map a reference addresses. `null` means the name is unknown. */
  function cellsOf(sheet: string | undefined): Record<CellKey, RawCell> | null {
    if (sheet === undefined) return cells;
    const named = otherSheets.get(sheet.toLowerCase());
    return named ?? null;
  }
  const cache = new Map<string, FormulaArg>();
  const visiting = new Set<string>();

  /**
   * Substitutes the imported value for a formula this engine could not resolve.
   *
   * Only on failure, and never in place of a formula we CAN evaluate: preferring
   * the import would freeze the sheet, since an edit to a cell's inputs would
   * recompute a result nothing ever displays.
   */
  function orImported(key: string, failure: FormulaArg): FormulaArg {
    return cachedValues[key as CellKey] ?? failure;
  }

  /**
   * The whole value of a cell — an array formula's array, not just its corner.
   *
   * THROWS CycleSignal on a reference cycle so the signal reaches every frame:
   * `getValue` coerces through `toNumber`, which would flatten a returned
   * sentinel to 0 one level up. `getCellDisplay` is the boundary that turns the
   * signal back into a value.
   */
  function compute(row: number, col: number, sheet?: string): FormulaArg {
    const source = cellsOf(sheet);
    if (source === null) return "#REF!";
    // The cache is keyed by sheet as well as position, so the same address on two
    // sheets does not collide.
    const key = sheet === undefined ? `${row}_${col}` : `${sheet}!${row}_${col}`;
    const cellKeyOf = `${row}_${col}` as CellKey;

    const cached = cache.get(key);
    if (cached !== undefined) return cached;

    if (visiting.has(key)) {
      cache.set(key, CYCLE_ERROR);
      throw new CycleSignal();
    }

    const raw = source[cellKeyOf];

    // A cell inside a declared output region is the anchor's, not its own. Its
    // stored value is what the writing application computed, kept as the fallback
    // for when we cannot recompute the anchor.
    if (sheet === undefined && declaredSpills.length > 0) {
      const anchor = anchorOver(row, col);
      if (anchor && !visiting.has(`${anchor[0]}_${anchor[1]}`)) {
        const whole = compute(anchor[0], anchor[1]);
        if (isMatrix(whole)) {
          const value = at(whole, row - anchor[0], col - anchor[1]);
          cache.set(key, value ?? "");
          return value ?? "";
        }
      }
    }

    if (raw === undefined || raw === "") {
      // Not a value of its own: it may be a cell another formula spills onto.
      // Only the evaluator's own sheet has a spill index; a cross-sheet spill
      // would need one per sheet, which no observed file needs.
      const spilled = sheet === undefined ? spilledInto(row, col) : "";
      cache.set(key, spilled);
      return spilled;
    }

    if (raw[0] !== "=") {
      cache.set(key, raw);
      return raw;
    }

    visiting.add(key);
    let result: FormulaArg;
    try {
      result = evalNode(astFor(raw), {
        ...ctx,
        currentRow: row,
        scope: EMPTY_SCOPE,
      });
      if (typeof result === "number" && !Number.isFinite(result)) {
        result = "#ERROR!";
      }
    } catch (e) {
      visiting.delete(key);
      if (e instanceof CycleSignal) {
        // Mark this cell and keep unwinding, so every cell in the cycle reports
        // #CYCLE! rather than only the one that closed it.
        cache.set(key, CYCLE_ERROR);
        throw e;
      }
      const failed = orImported(cellKeyOf, "#ERROR!");
      cache.set(key, failed);
      return failed;
    }
    visiting.delete(key);

    // An error sentinel means the formula referenced something we do not have —
    // an unimplemented function, a structured reference — which is exactly the
    // case the import's own result covers.
    if (isErrorValue(result)) result = orImported(cellKeyOf, result);
    // An array needs somewhere to go. Checked here rather than only while the
    // spill index is built, because the anchor may be asked for first and must
    // not report a value it cannot actually place.
    if (isMatrix(result) && blocks(source, row, col, result)) result = "#SPILL!";
    cache.set(key, result);
    return result;
  }

  /**
   * True when a cell other than the anchor already holds content of its own.
   *
   * Cells inside the anchor's DECLARED output region do not count: they hold the
   * array's own previously-written result, and treating them as an obstruction
   * makes every imported dynamic array report #SPILL! against itself.
   */
  function blocks(
    source: Record<CellKey, RawCell>,
    row: number,
    col: number,
    value: Matrix,
  ): boolean {
    const declared = opts.spillRanges?.[`${row}_${col}` as CellKey];
    const height = rowCount(value);
    const width = colCount(value);
    for (let r = 0; r < height; r++) {
      for (let c = 0; c < width; c++) {
        if (r === 0 && c === 0) continue;
        if (!source[`${row + r}_${col + c}` as CellKey]) continue;
        const inDeclared =
          declared !== undefined &&
          row + r >= declared.r1 &&
          row + r <= declared.r2 &&
          col + c >= declared.c1 &&
          col + c <= declared.c2;
        if (!inDeclared) return true;
      }
    }
    return false;
  }

  /**
   * A cell's value as a formula sees it, with a numeric literal read as a number.
   *
   * `compute` deliberately keeps a literal's raw text, because that text is what
   * the cell DISPLAYS — "1.50" typed by a user must not come back as "1.5". A
   * formula wants the number, and it must not get the 0 that blanket coercion
   * gave text: `=A1` on a cell reading "hello" used to evaluate to 0.
   */
  function typedCell(row: number, col: number, sheet?: string): FormulaValue {
    return coerceLiteral(toScalar(compute(row, col, sheet)));
  }

  /** Parsed once per distinct formula text, which repeats down a filled column. */
  const asts = new Map<string, Node>();

  function astFor(raw: string): Node {
    const hit = asts.get(raw);
    if (hit) return hit;
    const ast = parseFormula(tokenize(raw.slice(1)));
    asts.set(raw, ast);
    return ast;
  }

  // --- spilling ---

  /**
   * Where each spilling formula's array lands, keyed by the cell it covers.
   *
   * Built once, lazily, on the first query for an empty cell — and only over the
   * formulas that COULD return an array, which `canSpill` decides from the parse
   * tree. Most formulas cannot, so a sheet of `=A1*B1` builds an empty index and
   * pays a walk of its ASTs rather than a full recalculation.
   */
  let spills: Map<string, FormulaValue> | null = null;
  let buildingSpills = false;

  function spilledInto(row: number, col: number): FormulaValue {
    if (spills === null) {
      // Re-entrant while building: a candidate formula reading an empty cell gets
      // "" rather than recursing. A spill that depends on another spill's tail is
      // therefore not seen — documented in docs/LIMITATIONS.md.
      if (buildingSpills) return "";
      spills = buildSpills();
    }
    return spills.get(`${row}_${col}`) ?? "";
  }

  function buildSpills(): Map<string, FormulaValue> {
    const map = new Map<string, FormulaValue>();
    buildingSpills = true;
    try {
      for (const key of Object.keys(cells) as CellKey[]) {
        const raw = cells[key];
        if (raw === undefined || raw[0] !== "=") continue;

        // Text prefilter BEFORE parsing. The scan visits every formula on the
        // sheet, and parsing each one cost more than the whole index was worth:
        // 100k distinct `=A1*2` formulas parse to 100k ASTs to prove that none of
        // them can spill. The regex rejects those without allocating anything.
        if (!mightSpill(raw)) continue;

        let ast: Node;
        try {
          ast = astFor(raw);
        } catch {
          continue;
        }
        if (!canSpill(ast)) continue;

        const separator = key.indexOf("_");
        const row = Number(key.slice(0, separator));
        const col = Number(key.slice(separator + 1));

        let value: FormulaArg;
        try {
          value = compute(row, col);
        } catch {
          continue;
        }
        if (!isMatrix(value)) continue;

        // The anchor shows the top-left itself; only the rest is a spill. A
        // blocked one never reaches here: `compute` has already turned it into
        // #SPILL!, so it is not a matrix.
        const height = rowCount(value);
        const width = colCount(value);
        for (let r = 0; r < height; r++) {
          for (let c = 0; c < width; c++) {
            if (r === 0 && c === 0) continue;
            map.set(`${row + r}_${col + c}`, at(value, r, c) ?? "");
          }
        }
      }
    } finally {
      buildingSpills = false;
    }
    return map;
  }

  function getCellValue(row: number, col: number): FormulaArg {
    try {
      return compute(row, col);
    } catch (e) {
      if (e instanceof CycleSignal) return CYCLE_ERROR;
      throw e;
    }
  }

  /**
   * What the cell SHOWS. An array formula displays its top-left value in the
   * anchor and the rest in the cells it spills onto, so a matrix collapses here
   * rather than being stringified.
   */
  function getCellDisplay(row: number, col: number): FormulaValue {
    const value = getCellValue(row, col);
    return isLambda(value) ? "#VALUE!" : toScalar(value);
  }

  const ctx: EvalContext = {
    namedRanges,
    tables,
    scope: EMPTY_SCOPE,
    currentRow: -1,

    getCell: (row, col, sheet) => typedCell(row, col, sheet),
    getValue: (row, col) => toNumber(compute(row, col)),
    getNamedFormula: (name) => namedFormula(name),

    getRangeAbs: (r1, c1, r2, c2, sheet) => {
      const rows: FormulaValue[][] = [];
      for (let r = Math.min(r1, r2); r <= Math.max(r1, r2); r++) {
        const row: FormulaValue[] = [];
        for (let c = Math.min(c1, c2); c <= Math.max(c1, c2); c++) {
          row.push(typedCell(r, c, sheet));
        }
        rows.push(row);
      }
      return rows;
    },

    getRange: (fromRef, toRef, sheet) => {
      const a = parseCellRef(fromRef);
      const b = parseCellRef(toRef);
      return ctx.getRangeAbs(a.row, a.col, b.row, b.col, sheet);
    },
  };

  /**
   * A defined name holding a formula, evaluated once and memoized.
   *
   * `visitingNames` is a separate cycle guard from the cell one: a name that
   * refers to itself would otherwise recurse without ever touching a cell.
   */
  const namedResults = new Map<string, FormulaArg>();
  const visitingNames = new Set<string>();

  function namedFormula(name: string): FormulaArg | undefined {
    const body = namedFormulas[name] ?? namedFormulas[name.toUpperCase()];
    if (body === undefined) return undefined;

    const hit = namedResults.get(name);
    if (hit !== undefined) return hit;
    if (visitingNames.has(name)) return CYCLE_ERROR;

    visitingNames.add(name);
    let result: FormulaArg;
    try {
      result = evalNode(parseFormula(tokenize(body)), {
        ...ctx,
        scope: EMPTY_SCOPE,
      });
    } catch (e) {
      if (e instanceof CycleSignal) throw e;
      result = "#ERROR!";
    } finally {
      visitingNames.delete(name);
    }
    namedResults.set(name, result);
    return result;
  }

  /**
   * Memoized on the formula text. A conditional format asks the same question of
   * every cell in its range, and with absolute references — which is the common
   * case — that is one distinct question for the whole range.
   */
  const expressions = new Map<string, FormulaArg>();

  function evaluateArray(formulaBody: string): FormulaArg {
    const hit = expressions.get(formulaBody);
    if (hit !== undefined) return hit;

    let result: FormulaArg;
    try {
      result = evalNode(parseFormula(tokenize(formulaBody)), ctx);
    } catch (e) {
      if (e instanceof CycleSignal) result = CYCLE_ERROR;
      else result = "#ERROR!";
    }
    expressions.set(formulaBody, result);
    return result;
  }

  const evaluate = (formulaBody: string): FormulaValue =>
    toScalar(evaluateArray(formulaBody));

  return { getCellDisplay, getCellValue, evaluate, evaluateArray };
}

const EMPTY_SCOPE: Scope = new Map();

/** Text that is entirely a number, so "1e3" counts and "1 apple" does not. */
const NUMERIC_TEXT = /^[+-]?(\d+\.?\d*|\.\d+)(e[+-]?\d+)?$/i;

function coerceLiteral(v: FormulaValue): FormulaValue {
  if (typeof v !== "string" || v === "") return v;
  if (NUMERIC_TEXT.test(v)) return Number.parseFloat(v);
  const upper = v.toUpperCase();
  if (upper === "TRUE") return true;
  if (upper === "FALSE") return false;
  return v;
}

/**
 * Whether a formula could possibly return an array, decided from its parse tree.
 *
 * The spill index has to evaluate every candidate formula, so narrowing the
 * candidates is what keeps it from being a full recalculation of the sheet. A
 * formula spills only if its ROOT produces an array — an array used inside `SUM`
 * does not — so this looks at the root and follows only the paths a value can
 * reach it by.
 */
function canSpill(node: Node): boolean {
  switch (node.type) {
    case "arr":
    case "range":
    case "tableRef":
    case "name":
      return true;
    case "call":
      return ARRAY_RETURNING.has(node.name) || LAZY_PASSTHROUGH.has(node.name);
    case "bin":
    case "cmp":
      return canSpill(node.left) || canSpill(node.right);
    case "neg":
      return canSpill(node.node);
    default:
      return false;
  }
}

/** Functions whose result can be an array. */
const ARRAY_RETURNING: ReadonlySet<string> = new Set([
  "INDIRECT",
  "SEQUENCE",
  "UNIQUE",
  "SORT",
  "SORTBY",
  "FILTER",
  "HSTACK",
  "VSTACK",
  "TRANSPOSE",
  "CHOOSECOLS",
  "CHOOSEROWS",
  "TAKE",
  "DROP",
  "TOROW",
  "TOCOL",
  "MAKEARRAY",
  "MAP",
  "BYROW",
  "BYCOL",
  "SCAN",
  "TEXTSPLIT",
  "XLOOKUP",
  "OFFSET",
  "INDEX",
  "LOOKUP",
]);

/** Forms that return one of their arguments, so an array can pass through them. */
const LAZY_PASSTHROUGH: ReadonlySet<string> = new Set([
  "LET",
  "IF",
  "IFS",
  "SWITCH",
  "IFERROR",
  "IFNA",
  "REDUCE",
]);

/**
 * A cheap text test for whether a formula is worth parsing to ask `canSpill`.
 *
 * Conservative in the safe direction: it may say yes to a formula that turns out
 * not to spill, and `canSpill` then rejects it properly. It must never say no to
 * one that does, so it matches any array-returning or pass-through function name,
 * any array literal, and any formula with no call at all — a bare `=Name`,
 * `=A1:B9`, or `=tbl[col]`.
 */
const SPILL_HINT = new RegExp(
  `[{[]|\\b(?:${[...ARRAY_RETURNING, ...LAZY_PASSTHROUGH].join("|")})\\s*\\(`,
  "i",
);

/** `=OverallTable` — a defined name on its own, which is a spill if it is an array. */
const BARE_NAME = /^=\s*[A-Za-z_][A-Za-z0-9_.]*\s*$/;

function mightSpill(raw: string): boolean {
  // Ordered by cost. The overwhelmingly common formula on a large sheet is
  // something like `=A1*2`, and it has to be rejected without allocating: an
  // array literal, a structured reference, and a range are all single characters
  // to look for, and only a formula with a call is worth the alternation regex.
  if (raw.includes("{") || raw.includes("[") || raw.includes(":")) return true;
  if (!raw.includes("(")) return BARE_NAME.test(raw);
  return SPILL_HINT.test(raw);
}
