/**
 * Formula tokenizer. Ported from ref/formulaEngine.js:41-90.
 *
 * Unrecognized characters are skipped rather than throwing — a malformed
 * formula in one cell must not take down evaluation for the whole sheet.
 */
import type { CompareOp, Token } from "./ast.js";

/**
 * Cell refs are matched with an explicit regex (supporting `$` absolute markers)
 * BEFORE generic identifier scanning, so "SUM" stays a name while "A1" and
 * "$A$1" become ref tokens carrying colAbs/rowAbs.
 *
 * The trailing negative lookahead is what stops "A1B" from lexing as ref "A1"
 * followed by name "B".
 */
const REF_RE = /^(\$?)([A-Za-z]{1,3})(\$?)([0-9]{1,7})(?![A-Za-z0-9_.])/;

/**
 * A sheet qualifier before a ref: `Sheet2!`, or `'My Sheet'!` when the name
 * contains anything that would not lex as an identifier.
 *
 * Matched before the ref so `Sheet1!A1` does not lex as name `Sheet1`, a stray
 * `!`, and ref `A1` — which is what made cross-sheet formulas silently wrong
 * rather than merely unsupported.
 */
const SHEET_QUALIFIER = /^(?:'((?:[^']|'')*)'|([A-Za-z_][A-Za-z0-9_. ]*))!/;

const PUNCT = "+-*/^&%(),:";

/**
 * Identifiers may contain a dot, so `_xlfn.IFS` and a table name like
 * `Sales.2024` lex as one name. The reader strips Excel's `_xlfn.`/`_xlws.`
 * prefixes, but a formula typed by hand or written by another tool may not have
 * been through it.
 */
const IDENT_CHAR = /[A-Za-z0-9_.]/;

/**
 * Reads a bracketed structured-reference selector, honouring nesting.
 *
 * `tbl[[#This Row],[a]:[b]]` has brackets inside brackets, so stopping at the
 * first `]` would cut the selector in half and leave the rest to be lexed as
 * expression tokens.
 */
function readSelector(
  src: string,
  open: number,
): { spec: string; end: number } | null {
  let depth = 0;
  for (let i = open; i < src.length; i++) {
    const ch = src[i];
    if (ch === "[") depth++;
    else if (ch === "]") {
      depth--;
      if (depth === 0) return { spec: src.slice(open + 1, i), end: i + 1 };
    }
  }
  return null;
}

/** Parses `{1,2;3,4}` into rows of literals. */
function readArrayLiteral(
  src: string,
  open: number,
): { rows: (number | string | boolean)[][]; end: number } | null {
  const close = src.indexOf("}", open);
  if (close === -1) return null;
  const body = src.slice(open + 1, close);
  const rows = body.split(";").map((row) =>
    row.split(",").map((cell) => {
      const text = cell.trim();
      if (text.startsWith('"')) return text.slice(1, -1);
      if (text.toUpperCase() === "TRUE") return true;
      if (text.toUpperCase() === "FALSE") return false;
      const n = Number.parseFloat(text);
      return Number.isNaN(n) ? text : n;
    }),
  );
  return { rows, end: close + 1 };
}

export function tokenize(src: string): Token[] {
  const tokens: Token[] = [];
  let i = 0;

  while (i < src.length) {
    const ch = src[i] as string;

    if (/\s/.test(ch)) {
      i++;
      continue;
    }

    // Refs first — see REF_RE above.
    if (ch === "$" || ch === "'" || /[A-Za-z]/.test(ch)) {
      const rest = src.slice(i);
      const qualifier = SHEET_QUALIFIER.exec(rest);
      const afterQualifier = qualifier ? qualifier[0].length : 0;
      const m = REF_RE.exec(rest.slice(afterQualifier));
      if (m) {
        // A quoted name escapes its own quote by doubling it.
        const sheet = qualifier
          ? ((qualifier[1] ?? qualifier[2]) as string).replace(/''/g, "'")
          : undefined;
        tokens.push({
          type: "ref",
          value: (m[2] as string).toUpperCase() + (m[4] as string),
          colAbs: !!m[1],
          rowAbs: !!m[3],
          ...(sheet ? { sheet } : {}),
        });
        i += afterQualifier + m[0].length;
        continue;
      }
      // A qualifier followed by a table name, `Sheet2!tbl[col]`, is still a
      // structured reference; the sheet is implied by the table.
      if (qualifier && /^[A-Za-z_]/.test(rest.slice(afterQualifier))) {
        i += afterQualifier;
        continue;
      }
    }

    if (/[0-9.]/.test(ch)) {
      let j = i;
      while (j < src.length && /[0-9.]/.test(src[j] as string)) j++;
      tokens.push({ type: "num", value: parseFloat(src.slice(i, j)) });
      i = j;
      continue;
    }

    if (/[A-Za-z_]/.test(ch)) {
      let j = i;
      while (j < src.length && IDENT_CHAR.test(src[j] as string)) j++;
      const raw = src.slice(i, j);

      // A `[` immediately after a name makes it a structured reference. The
      // table name keeps its case, since it is matched against the file's.
      if (src[j] === "[") {
        const selector = readSelector(src, j);
        if (selector) {
          tokens.push({ type: "tableRef", table: raw, spec: selector.spec });
          i = selector.end;
          continue;
        }
      }

      tokens.push({ type: "name", value: raw.toUpperCase() });
      i = j;
      continue;
    }

    if (ch === "{") {
      const literal = readArrayLiteral(src, i);
      if (literal) {
        tokens.push({ type: "arr", rows: literal.rows });
        i = literal.end;
        continue;
      }
    }

    if (PUNCT.includes(ch)) {
      tokens.push({ type: ch as "+" });
      i++;
      continue;
    }

    if (ch === '"') {
      let j = i + 1;
      let s = "";
      while (j < src.length && src[j] !== '"') {
        s += src[j];
        j++;
      }
      tokens.push({ type: "str", value: s });
      i = j + 1;
      continue;
    }

    if (ch === "=" || ch === "<" || ch === ">") {
      let op = ch;
      let j = i + 1;
      // BUGFIX vs ref/formulaEngine.js:80-86, which only ever appended "=" and
      // so could not produce "<>". `A1<>B1` lexed as cmp "<" followed by cmp
      // ">", and parsePrimary swallowed the ">" as a literal 0 — silently
      // wrong, never an error. evalCompare already had a "<>" case waiting.
      if (src[j] === "=" || (ch === "<" && src[j] === ">")) {
        op += src[j] as string;
        j++;
      }
      tokens.push({ type: "cmp", value: op as CompareOp });
      i = j;
      continue;
    }

    i++;
  }

  return tokens;
}
