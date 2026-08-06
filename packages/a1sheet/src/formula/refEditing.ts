/**
 * Locating and inserting cell references inside formula source, while it is
 * being typed.
 *
 * Separate from ./tokenize.ts on purpose. The tokenizer throws source positions
 * away — it produces a `ref` token with a normalized value, which is exactly
 * right for evaluation and useless for an editor, which needs to know that
 * characters 5..10 of the text are the reference so it can underline them and
 * replace them when you click another cell.
 *
 * Separate from ./refs.ts too: that shifts references when a formula is copied,
 * which happens to a committed formula. This is about the one being typed.
 *
 * Nothing here evaluates anything, and nothing here knows about React.
 */
import { normalizeRange, parseCellRef } from "../model/address.js";
import type { Range } from "../model/types.js";

/** One reference found in the source, with the span it occupies. */
export interface RefSpan {
  /** Character offset into the source string, leading "=" included. */
  start: number;
  /** Offset one past the last character. */
  end: number;
  /** Exactly as typed: "a1", "$A$1", "B2:C4". */
  text: string;
  /** The cells the reference covers, normalized. */
  range: Range;
  /**
   * Refs that resolve to the same range share a group. A consumer colors by
   * group so `=A1+A1` underlines both occurrences identically, as Sheets does.
   * Numbered by first appearance, so it is stable while you type.
   */
  group: number;
}

/** One cell: optional `$`, 1-3 letters, optional `$`, 1-7 digits. */
const CELL = String.raw`\$?[A-Za-z]{1,3}\$?[0-9]{1,7}`;

/**
 * A cell or a range. The trailing lookahead is what stops "A1B" matching as
 * "A1", mirroring the tokenizer's rule.
 */
const REF_RE = new RegExp(`${CELL}(?::${CELL})?(?![A-Za-z0-9_])`, "g");

/** Characters after which a reference may legally begin. */
const OPERAND_POSITION = new Set(["=", "+", "-", "*", "/", "^", "(", ",", ":"]);

/**
 * Every reference in `source`, in order.
 *
 * Two things are deliberately skipped:
 *
 * - text inside string literals, so `="A1"` underlines nothing;
 * - a match immediately followed by "(", which is a function call. `LOG10(` is
 *   three letters and two digits, so it matches the reference pattern exactly.
 *   The tokenizer has the same blind spot and resolves it later; here it would
 *   be visible as a stray underline under a function name.
 */
export function findRefSpans(source: string): RefSpan[] {
  const spans: RefSpan[] = [];
  const groups = new Map<string, number>();

  for (const [start, end] of scannableSpans(source)) {
    REF_RE.lastIndex = 0;
    const segment = source.slice(start, end);
    let m = REF_RE.exec(segment);

    while (m) {
      const text = m[0];
      const at = start + m.index;
      const after = source[at + text.length];

      if (after !== "(") {
        const range = toRange(text);
        const key = `${range.r1}:${range.c1}:${range.r2}:${range.c2}`;
        let group = groups.get(key);
        if (group === undefined) {
          group = groups.size;
          groups.set(key, group);
        }
        spans.push({ start: at, end: at + text.length, text, range, group });
      }

      m = REF_RE.exec(segment);
    }
  }

  return spans;
}

/**
 * The `[start, end)` regions of `source` that are outside string literals.
 *
 * An unterminated quote runs to the end of the source, which is the common case
 * mid-typing: `=CONCAT("hello` should not start matching refs again.
 */
function scannableSpans(source: string): Array<[number, number]> {
  const out: Array<[number, number]> = [];
  let cursor = 0;

  for (let i = 0; i < source.length; i++) {
    if (source[i] !== '"') continue;
    out.push([cursor, i]);
    const close = source.indexOf('"', i + 1);
    if (close === -1) return out;
    cursor = close + 1;
    i = close;
  }

  out.push([cursor, source.length]);
  return out;
}

function toRange(text: string): Range {
  const [from, to] = text.split(":");
  const a = parseCellRef(from as string);
  if (to === undefined) return { r1: a.row, c1: a.col, r2: a.row, c2: a.col };
  const b = parseCellRef(to);
  return normalizeRange({ r1: a.row, c1: a.col, r2: b.row, c2: b.col });
}

export interface RefInsertion {
  /** The new source. */
  value: string;
  /** Where the caret belongs afterwards — just past the inserted reference. */
  caret: number;
  /** The span the reference now occupies, so a drag can keep rewriting it. */
  span: { start: number; end: number };
}

/**
 * Puts `ref` into `source` at `caret`, the way clicking a cell mid-formula does.
 *
 * Returns null when the caret is not somewhere a reference can go — after a
 * complete operand like `=A1+2` with the caret at the end. That is not a
 * failure: it is how the caller knows the click meant "I am done here, select
 * that cell instead", which is what Sheets does too.
 *
 * Two cases produce an insertion:
 *
 *   =SUM(A1,|        caret follows an operator or separator -> insert
 *   =SUM(A1|         caret sits on a reference              -> replace it
 */
export function insertRefAtCaret(
  source: string,
  caret: number,
  ref: string,
): RefInsertion | null {
  if (!source.startsWith("=")) return null;
  const at = Math.max(0, Math.min(caret, source.length));

  // Replacing wins over inserting: with the caret at the end of `=A1`, clicking
  // B2 should give `=B2`, not `=A1B2`.
  const existing = findRefSpans(source).find(
    (span) => at >= span.start && at <= span.end,
  );
  if (existing) return spliced(source, existing.start, existing.end, ref);

  const before = source.slice(0, at).trimEnd();
  const previous = before.at(-1);
  if (previous === undefined || !OPERAND_POSITION.has(previous)) return null;

  return spliced(source, at, at, ref);
}

function spliced(
  source: string,
  start: number,
  end: number,
  ref: string,
): RefInsertion {
  return {
    value: source.slice(0, start) + ref + source.slice(end),
    caret: start + ref.length,
    span: { start, end: start + ref.length },
  };
}

/** True when `source` is a formula, i.e. the editor should behave specially. */
export function isFormulaSource(source: string): boolean {
  return source.startsWith("=");
}
