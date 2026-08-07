/**
 * Recursive-descent parser. Ported from ref/formulaEngine.js:93-160.
 *
 * Precedence, loosest to tightest:
 *   parseCompare (= < > <= >= <>)
 *     parseExpr    (+ -)
 *       parseTerm  (* /)
 *         parseUnary (-x)
 *           parsePower (^, right-associative via parseUnary)
 *             parsePrimary (literals, refs, ranges, calls, parens)
 *
 * Like the tokenizer, this never throws. Unexpected input degrades to `0` so a
 * single bad formula cannot break the sheet.
 */
import type { BinaryOp, CompareOp, Node, Token } from "./ast.js";

const ZERO: Node = { type: "num", value: 0 };
const PERCENT = 100;

export function parseFormula(tokens: Token[]): Node {
  let pos = 0;
  const peek = (): Token | undefined => tokens[pos];
  const next = (): Token | undefined => tokens[pos++];

  function parseCompare(): Node {
    let node = parseConcat();
    for (let t = peek(); t?.type === "cmp"; t = peek()) {
      next();
      node = {
        type: "cmp",
        op: t.value as CompareOp,
        left: node,
        right: parseConcat(),
      };
    }
    return node;
  }

  /**
   * `&` binds looser than arithmetic and tighter than comparison, which is why it
   * sits between them rather than alongside `+`.
   */
  function parseConcat(): Node {
    let node = parseExpr();
    for (let t = peek(); t?.type === "&"; t = peek()) {
      next();
      node = { type: "bin", op: "&", left: node, right: parseExpr() };
    }
    return node;
  }

  function parseExpr(): Node {
    let node = parseTerm();
    for (let t = peek(); t?.type === "+" || t?.type === "-"; t = peek()) {
      next();
      node = { type: "bin", op: t.type, left: node, right: parseTerm() };
    }
    return node;
  }

  function parseTerm(): Node {
    let node = parseUnary();
    for (let t = peek(); t?.type === "*" || t?.type === "/"; t = peek()) {
      next();
      node = { type: "bin", op: t.type, left: node, right: parseUnary() };
    }
    return node;
  }

  function parseUnary(): Node {
    if (peek()?.type === "-") {
      next();
      return { type: "neg", node: parseUnary() };
    }
    return parsePower();
  }

  function parsePower(): Node {
    let node = parsePostfix();
    if (peek()?.type === "^") {
      next();
      node = { type: "bin", op: "^" as BinaryOp, left: node, right: parseUnary() };
    }
    return node;
  }

  /** `50%` is a postfix operator, not a stray token to skip. */
  function parsePostfix(): Node {
    let node = parsePrimary();
    while (peek()?.type === "%") {
      next();
      node = {
        type: "bin",
        op: "/" as BinaryOp,
        left: node,
        right: { type: "num", value: PERCENT },
      };
    }
    return node;
  }

  /**
   * Argument position is the only place a bare `A1:B2` range is recognized
   * without parens. Speculatively consumes a ref, and rewinds if no `:` follows.
   */
  function parseArg(): Node {
    if (peek()?.type === "ref") {
      const save = pos;
      const t = next();
      if (peek()?.type === ":") {
        next();
        const t2 = next();
        if (t?.type === "ref" && t2?.type === "ref") {
          return {
            type: "range",
            from: t.value,
            to: t2.value,
            ...(t.sheet ? { sheet: t.sheet } : {}),
          };
        }
      }
      pos = save;
    }
    return parseCompare();
  }

  function parsePrimary(): Node {
    const t = peek();
    if (!t) return ZERO;

    if (t.type === "num") {
      next();
      return { type: "num", value: t.value };
    }

    if (t.type === "str") {
      next();
      return { type: "str", value: t.value };
    }

    if (t.type === "arr") {
      next();
      return { type: "arr", rows: t.rows };
    }

    if (t.type === "tableRef") {
      next();
      return { type: "tableRef", table: t.table, spec: t.spec };
    }

    if (t.type === "ref") {
      next();
      const sheet = t.sheet ? { sheet: t.sheet } : {};
      if (peek()?.type === ":") {
        next();
        const t2 = next();
        if (t2?.type === "ref") {
          return { type: "range", from: t.value, to: t2.value, ...sheet };
        }
        return { type: "ref", value: t.value, ...sheet };
      }
      return { type: "ref", value: t.value, ...sheet };
    }

    if (t.type === "name") {
      next();
      if (peek()?.type === "(") {
        next();
        const args: Node[] = [];
        if (peek()?.type !== ")") {
          args.push(parseArg());
          while (peek()?.type === ",") {
            next();
            args.push(parseArg());
          }
        }
        if (peek()?.type === ")") next();
        return { type: "call", name: t.value, args };
      }
      return { type: "name", value: t.value };
    }

    if (t.type === "(") {
      next();
      const node = parseCompare();
      if (peek()?.type === ")") next();
      return node;
    }

    next();
    return ZERO;
  }

  return parseCompare();
}
