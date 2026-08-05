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
const REF_RE = /^(\$?)([A-Za-z]{1,3})(\$?)([0-9]{1,7})(?![A-Za-z0-9_])/;

const PUNCT = "+-*/^(),:";

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
    if (ch === "$" || /[A-Za-z]/.test(ch)) {
      const m = REF_RE.exec(src.slice(i));
      if (m) {
        tokens.push({
          type: "ref",
          value: (m[2] as string).toUpperCase() + (m[4] as string),
          colAbs: !!m[1],
          rowAbs: !!m[3],
        });
        i += m[0].length;
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
      while (j < src.length && /[A-Za-z0-9_]/.test(src[j] as string)) j++;
      tokens.push({ type: "name", value: src.slice(i, j).toUpperCase() });
      i = j;
      continue;
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
