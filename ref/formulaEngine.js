/* =============================================================================
   FORMULA ENGINE
   Tokenizer -> recursive-descent parser -> tree-walking evaluator.
   Supports: + - * / ^, parentheses, A1 / $A$1 / A$1 / $A1 refs, A1:B10 ranges,
   named ranges, and a function library (math, text, lookup, date, logic).

   Known simplifications (by design, to keep this readable & hackable):
   - Single-sheet formulas only. No `Sheet2!A1` cross-sheet refs yet.
   - MATCH/VLOOKUP only do exact matches (no approximate/sorted lookup).
   - Dates are day-serial numbers since the Unix epoch (not Excel's 1900
     epoch) — internally consistent for arithmetic, just not byte-identical
     to what Excel would compute for the same serial number.
   ============================================================================= */

/* ------------------------------- A1 helpers ------------------------------- */
export function colToLetters(n) {
  let s = "";
  n += 1;
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}
export function lettersToCol(s) {
  let n = 0;
  for (let i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
  return n - 1;
}
export function parseCellRef(ref) {
  const m = ref.match(/^\$?([A-Za-z]+)\$?(\d+)$/);
  if (!m) return { row: 0, col: 0 };
  return { col: lettersToCol(m[1].toUpperCase()), row: parseInt(m[2], 10) - 1 };
}

/* -------------------------------- Tokenizer -------------------------------- */
// Cell refs are matched with an explicit regex (supports $ absolute markers)
// BEFORE generic identifier scanning, so "SUM" stays a name but "A1"/"$A$1"
// become ref tokens carrying colAbs/rowAbs flags used later for fill/copy.
const REF_RE = /^(\$?)([A-Za-z]{1,3})(\$?)([0-9]{1,7})(?![A-Za-z0-9_])/;

export function tokenize(src) {
  const tokens = [];
  let i = 0;
  while (i < src.length) {
    const ch = src[i];
    if (/\s/.test(ch)) { i++; continue; }

    if (ch === "$" || /[A-Za-z]/.test(ch)) {
      const m = REF_RE.exec(src.slice(i));
      if (m) {
        tokens.push({ type: "ref", value: m[2].toUpperCase() + m[4], colAbs: !!m[1], rowAbs: !!m[3] });
        i += m[0].length;
        continue;
      }
    }
    if (/[0-9.]/.test(ch)) {
      let j = i;
      while (j < src.length && /[0-9.]/.test(src[j])) j++;
      tokens.push({ type: "num", value: parseFloat(src.slice(i, j)) });
      i = j;
      continue;
    }
    if (/[A-Za-z_]/.test(ch)) {
      let j = i;
      while (j < src.length && /[A-Za-z0-9_]/.test(src[j])) j++;
      tokens.push({ type: "name", value: src.slice(i, j).toUpperCase() });
      i = j;
      continue;
    }
    if ("+-*/^(),:".includes(ch)) { tokens.push({ type: ch }); i++; continue; }
    if (ch === '"') {
      let j = i + 1, s = "";
      while (j < src.length && src[j] !== '"') { s += src[j]; j++; }
      tokens.push({ type: "str", value: s });
      i = j + 1;
      continue;
    }
    if (ch === "=" || ch === "<" || ch === ">") {
      let op = ch, j = i + 1;
      if (src[j] === "=") { op += "="; j++; }
      tokens.push({ type: "cmp", value: op });
      i = j;
      continue;
    }
    i++; // skip anything unrecognized rather than hard-fail the whole sheet
  }
  return tokens;
}

/* --------------------------------- Parser ---------------------------------- */
export function parseFormula(tokens) {
  let pos = 0;
  const peek = () => tokens[pos];
  const next = () => tokens[pos++];

  function parseCompare() {
    let node = parseExpr();
    while (peek() && peek().type === "cmp") { const op = next().value; node = { type: "cmp", op, left: node, right: parseExpr() }; }
    return node;
  }
  function parseExpr() {
    let node = parseTerm();
    while (peek() && (peek().type === "+" || peek().type === "-")) { const op = next().type; node = { type: "bin", op, left: node, right: parseTerm() }; }
    return node;
  }
  function parseTerm() {
    let node = parseUnary();
    while (peek() && (peek().type === "*" || peek().type === "/")) { const op = next().type; node = { type: "bin", op, left: node, right: parseUnary() }; }
    return node;
  }
  function parseUnary() {
    if (peek() && peek().type === "-") { next(); return { type: "neg", node: parseUnary() }; }
    return parsePower();
  }
  function parsePower() {
    let node = parsePrimary();
    if (peek() && peek().type === "^") { next(); node = { type: "bin", op: "^", left: node, right: parseUnary() }; }
    return node;
  }
  function parseArg() {
    if (peek() && peek().type === "ref") {
      const save = pos;
      const t = next();
      if (peek() && peek().type === ":") { next(); const t2 = next(); return { type: "range", from: t.value, to: t2.value }; }
      pos = save;
    }
    return parseCompare();
  }
  function parsePrimary() {
    const t = peek();
    if (!t) return { type: "num", value: 0 };
    if (t.type === "num") { next(); return { type: "num", value: t.value }; }
    if (t.type === "str") { next(); return { type: "str", value: t.value }; }
    if (t.type === "ref") {
      next();
      if (peek() && peek().type === ":") { next(); const t2 = next(); return { type: "range", from: t.value, to: t2.value }; }
      return { type: "ref", value: t.value };
    }
    if (t.type === "name") {
      next();
      if (peek() && peek().type === "(") {
        next();
        const args = [];
        if (!(peek() && peek().type === ")")) {
          args.push(parseArg());
          while (peek() && peek().type === ",") { next(); args.push(parseArg()); }
        }
        if (peek() && peek().type === ")") next();
        return { type: "call", name: t.value, args };
      }
      return { type: "name", value: t.value };
    }
    if (t.type === "(") { next(); const node = parseCompare(); if (peek() && peek().type === ")") next(); return node; }
    next();
    return { type: "num", value: 0 };
  }
  return parseCompare();
}

/* -------------------------------- Utilities -------------------------------- */
export function toNumber(v) {
  if (typeof v === "number") return v;
  if (v === true) return 1;
  if (v === false) return 0;
  const n = parseFloat(v);
  return isNaN(n) ? 0 : n;
}
export function toText(v) {
  if (Array.isArray(v)) return v.map(toText).join("");
  return v === undefined || v === null ? "" : String(v);
}
function flattenNums(args) {
  const out = [];
  for (const a of args) {
    if (Array.isArray(a)) for (const v of a) { const n = parseFloat(v); if (!isNaN(n)) out.push(n); }
    else { const n = parseFloat(a); if (!isNaN(n)) out.push(n); }
  }
  return out;
}
function flattenBool(args) {
  const out = [];
  const push = (v) => out.push(typeof v === "string" ? v !== "" : !!toNumber(v));
  for (const a of args) { if (Array.isArray(a)) a.forEach(push); else push(a); }
  return out;
}

/* ------------------------------ Function library ---------------------------- */
const DAY_MS = 86400000;
export const FUNCTIONS = {
  // math / stats
  SUM: (a) => flattenNums(a).reduce((x, y) => x + y, 0),
  AVERAGE: (a) => { const n = flattenNums(a); return n.length ? n.reduce((x, y) => x + y, 0) / n.length : 0; },
  MIN: (a) => { const n = flattenNums(a); return n.length ? Math.min(...n) : 0; },
  MAX: (a) => { const n = flattenNums(a); return n.length ? Math.max(...n) : 0; },
  COUNT: (a) => flattenNums(a).length,
  COUNTA: (a) => { let c = 0; const flat = []; for (const v of a) Array.isArray(v) ? flat.push(...v) : flat.push(v); for (const v of flat) if (v !== "" && v !== undefined) c++; return c; },
  ABS: (a) => Math.abs(toNumber(a[0])),
  ROUND: (a) => { const n = toNumber(a[0]); const d = a[1] !== undefined ? toNumber(a[1]) : 0; const f = Math.pow(10, d); return Math.round(n * f) / f; },
  IF: (a) => (toNumber(a[0]) ? a[1] : a[2]),
  AND: (a) => (flattenBool(a).every(Boolean) ? 1 : 0),
  OR: (a) => (flattenBool(a).some(Boolean) ? 1 : 0),
  NOT: (a) => (toNumber(a[0]) ? 0 : 1),
  // text
  CONCAT: (a) => a.map(toText).join(""),
  CONCATENATE: (a) => a.map(toText).join(""),
  LEFT: (a) => toText(a[0]).slice(0, a[1] !== undefined ? toNumber(a[1]) : 1),
  RIGHT: (a) => { const s = toText(a[0]); const n = a[1] !== undefined ? toNumber(a[1]) : 1; return s.slice(Math.max(0, s.length - n)); },
  MID: (a) => toText(a[0]).slice(toNumber(a[1]) - 1, toNumber(a[1]) - 1 + toNumber(a[2])),
  TRIM: (a) => toText(a[0]).trim(),
  UPPER: (a) => toText(a[0]).toUpperCase(),
  LOWER: (a) => toText(a[0]).toLowerCase(),
  LEN: (a) => toText(a[0]).length,
  // date (day-serial numbers since Unix epoch)
  TODAY: () => Math.floor(Date.now() / DAY_MS),
  NOW: () => Date.now() / DAY_MS,
  DATE: (a) => Math.floor(Date.UTC(toNumber(a[0]), toNumber(a[1]) - 1, toNumber(a[2])) / DAY_MS),
  YEAR: (a) => new Date(toNumber(a[0]) * DAY_MS).getUTCFullYear(),
  MONTH: (a) => new Date(toNumber(a[0]) * DAY_MS).getUTCMonth() + 1,
  DAY: (a) => new Date(toNumber(a[0]) * DAY_MS).getUTCDate(),
};
export { DAY_MS };

/* --------------------------------- Evaluator -------------------------------- */
function evalCompare(op, l, r) {
  const ln = typeof l === "number" ? l : parseFloat(l);
  const rn = typeof r === "number" ? r : parseFloat(r);
  const bothNum = !isNaN(ln) && !isNaN(rn);
  const a = bothNum ? ln : toText(l);
  const b = bothNum ? rn : toText(r);
  switch (op) {
    case "=": return a === b ? 1 : 0;
    case "<": return a < b ? 1 : 0;
    case ">": return a > b ? 1 : 0;
    case "<=": return a <= b ? 1 : 0;
    case ">=": return a >= b ? 1 : 0;
    case "<>": return a !== b ? 1 : 0;
    default: return 0;
  }
}

export function evalNode(node, ctx) {
  switch (node.type) {
    case "num": return node.value;
    case "str": return node.value;
    case "neg": return -toNumber(evalNode(node.node, ctx));
    case "cmp": return evalCompare(node.op, evalNode(node.left, ctx), evalNode(node.right, ctx));
    case "bin": {
      const l = toNumber(evalNode(node.left, ctx));
      const r = toNumber(evalNode(node.right, ctx));
      switch (node.op) {
        case "+": return l + r;
        case "-": return l - r;
        case "*": return l * r;
        case "/": return r === 0 ? "#DIV/0!" : l / r;
        case "^": return Math.pow(l, r);
        default: return 0;
      }
    }
    case "ref": { const { row, col } = parseCellRef(node.value); return ctx.getValue(row, col); }
    case "range": return ctx.getRange(node.from, node.to);
    case "name": {
      const nr = ctx.namedRanges && ctx.namedRanges[node.value];
      if (nr) return ctx.getRangeAbs(nr.r1, nr.c1, nr.r2, nr.c2);
      return "#NAME?";
    }
    case "call": {
      const name = node.name;
      if (name === "VLOOKUP") {
        const lookupVal = evalNode(node.args[0], ctx);
        const rangeArg = node.args[1];
        if (!rangeArg || rangeArg.type !== "range") return "#REF!";
        const table = ctx.getRange2D(rangeArg.from, rangeArg.to);
        const colIdx = Math.round(toNumber(evalNode(node.args[2], ctx))) - 1;
        for (const row of table) {
          if (toText(row[0]) === toText(lookupVal) || toNumber(row[0]) === toNumber(lookupVal)) {
            return row[colIdx] !== undefined ? row[colIdx] : "#REF!";
          }
        }
        return "#N/A";
      }
      if (name === "MATCH") {
        const lookupVal = evalNode(node.args[0], ctx);
        const rangeArg = node.args[1];
        const arr = rangeArg.type === "range" ? ctx.getRange(rangeArg.from, rangeArg.to) : [evalNode(rangeArg, ctx)];
        for (let i = 0; i < arr.length; i++) if (toText(arr[i]) === toText(lookupVal) || toNumber(arr[i]) === toNumber(lookupVal)) return i + 1;
        return "#N/A";
      }
      if (name === "INDEX") {
        const rangeArg = node.args[0];
        if (!rangeArg || rangeArg.type !== "range") return "#REF!";
        const table = ctx.getRange2D(rangeArg.from, rangeArg.to);
        const rowNum = Math.round(toNumber(evalNode(node.args[1], ctx)));
        const colNum = node.args[2] !== undefined ? Math.round(toNumber(evalNode(node.args[2], ctx))) : 1;
        const row = table[rowNum - 1];
        if (!row) return "#REF!";
        return row[colNum - 1] !== undefined ? row[colNum - 1] : "#REF!";
      }
      const fn = FUNCTIONS[name];
      if (!fn) return "#NAME?";
      const argVals = node.args.map((a) => (a.type === "range" ? ctx.getRange(a.from, a.to) : evalNode(a, ctx)));
      return fn(argVals);
    }
    default: return "";
  }
}

/** Lazy, memoized evaluator over a snapshot of raw cell text for one sheet. */
export function createEvaluator(cells, namedRanges) {
  const cache = new Map();
  const visiting = new Set();

  function getCellDisplay(row, col) {
    const key = `${row}_${col}`;
    if (cache.has(key)) return cache.get(key);
    if (visiting.has(key)) { cache.set(key, "#CYCLE!"); return "#CYCLE!"; }
    const raw = cells[key];
    if (raw === undefined || raw === "") { cache.set(key, ""); return ""; }
    if (typeof raw === "string" && raw[0] === "=") {
      visiting.add(key);
      let result;
      try {
        const ast = parseFormula(tokenize(raw.slice(1)));
        result = evalNode(ast, ctx);
        if (typeof result === "number" && !isFinite(result)) result = "#ERROR!";
      } catch (e) { result = "#ERROR!"; }
      visiting.delete(key);
      cache.set(key, result);
      return result;
    }
    cache.set(key, raw);
    return raw;
  }
  const ctx = {
    namedRanges,
    getValue: (row, col) => toNumber(getCellDisplay(row, col)),
    getRangeAbs: (r1, c1, r2, c2) => {
      const out = [];
      for (let r = Math.min(r1, r2); r <= Math.max(r1, r2); r++)
        for (let c = Math.min(c1, c2); c <= Math.max(c1, c2); c++) out.push(getCellDisplay(r, c));
      return out;
    },
    getRange: (fromRef, toRef) => {
      const a = parseCellRef(fromRef), b = parseCellRef(toRef);
      return ctx.getRangeAbs(a.row, a.col, b.row, b.col);
    },
    getRange2D: (fromRef, toRef) => {
      const a = parseCellRef(fromRef), b = parseCellRef(toRef);
      const r1 = Math.min(a.row, b.row), r2 = Math.max(a.row, b.row);
      const c1 = Math.min(a.col, b.col), c2 = Math.max(a.col, b.col);
      const rows = [];
      for (let r = r1; r <= r2; r++) { const row = []; for (let c = c1; c <= c2; c++) row.push(getCellDisplay(r, c)); rows.push(row); }
      return rows;
    },
  };
  return { getCellDisplay };
}

/* ------------------------- Relative-reference shifting ---------------------- */
// Used by the fill handle and by internal copy/paste: shifts every non-$
// reference in a formula by (dRow, dCol), the way Excel does when you drag
// or copy-paste a formula to a new cell.
export function shiftFormulaRefs(formulaText, dRow, dCol) {
  const tokens = tokenize(formulaText);
  const parts = tokens.map((t) => {
    if (t.type === "ref") {
      const { row, col } = parseCellRef(t.value);
      const newCol = t.colAbs ? col : col + dCol;
      const newRow = t.rowAbs ? row : row + dRow;
      if (newCol < 0 || newRow < 0) return "#REF!";
      return (t.colAbs ? "$" : "") + colToLetters(newCol) + (t.rowAbs ? "$" : "") + (newRow + 1);
    }
    if (t.type === "num") return String(t.value);
    if (t.type === "str") return `"${t.value}"`;
    if (t.type === "name") return t.value;
    if (t.type === "cmp") return t.value;
    return t.type;
  });
  return parts.join("");
}

/** Detects a simple linear (or repeating) series for fill-handle drag. */
export function extrapolateSeries(values, count) {
  const nums = values.map((v) => parseFloat(v));
  const allNumeric = nums.every((n) => !isNaN(n)) && values.every((v) => v !== "");
  const out = [];
  if (allNumeric && values.length >= 2) {
    const step = nums[nums.length - 1] - nums[nums.length - 2];
    for (let i = 0; i < count; i++) out.push(String(nums[nums.length - 1] + step * (i + 1)));
  } else if (allNumeric && values.length === 1) {
    for (let i = 0; i < count; i++) out.push(String(nums[0] + (i + 1)));
  } else {
    for (let i = 0; i < count; i++) out.push(values[i % values.length]);
  }
  return out;
}
