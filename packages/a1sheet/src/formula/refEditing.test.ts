import { describe, expect, test } from "bun:test";
import { findRefSpans, insertRefAtCaret, isFormulaSource } from "./refEditing.js";

/** Compact view of a span list: the text and where it sits. */
const at = (source: string) =>
  findRefSpans(source).map((s) => `${s.text}@${s.start}-${s.end}`);

describe("findRefSpans", () => {
  test("finds a single cell reference with its position", () => {
    expect(findRefSpans("=A1")[0]).toMatchObject({
      start: 1,
      end: 3,
      text: "A1",
      range: { r1: 0, c1: 0, r2: 0, c2: 0 },
    });
  });

  test("finds a range as one span, not two", () => {
    expect(at("=SUM(B2:C4)")).toEqual(["B2:C4@5-10"]);
    expect(findRefSpans("=SUM(B2:C4)")[0]?.range).toEqual({
      r1: 1,
      c1: 1,
      r2: 3,
      c2: 2,
    });
  });

  test("finds several references in order", () => {
    expect(at("=A1+B2*C3")).toEqual(["A1@1-3", "B2@4-6", "C3@7-9"]);
  });

  test("keeps absolute markers in the text and ignores them in the range", () => {
    const [span] = findRefSpans("=$A$1");
    expect(span?.text).toBe("$A$1");
    expect(span?.range).toEqual({ r1: 0, c1: 0, r2: 0, c2: 0 });
  });

  test("lowercase references are found", () => {
    expect(at("=a1+b2")).toEqual(["a1@1-3", "b2@4-6"]);
  });

  test("a function name is not a reference", () => {
    // LOG10 is three letters and two digits — it matches the reference pattern
    // exactly, and only the following "(" tells them apart.
    expect(at("=LOG10(A1)")).toEqual(["A1@7-9"]);
    expect(at("=SUM(A1)")).toEqual(["A1@5-7"]);
  });

  test("text inside a string literal is not a reference", () => {
    expect(at('="A1"')).toEqual([]);
    expect(at('=CONCAT("A1",B2)')).toEqual(["B2@13-15"]);
  });

  test("an unterminated string swallows the rest of the source", () => {
    // Mid-typing state. Resuming the scan after the quote would underline
    // references inside what is about to become a string.
    expect(at('=CONCAT(A1,"B2')).toEqual(["A1@8-10"]);
  });

  test("A1B is not a reference", () => {
    expect(at("=A1B")).toEqual([]);
  });

  test("a bare value has no references", () => {
    expect(at("hello")).toEqual([]);
    expect(at("=1+2")).toEqual([]);
  });
});

describe("findRefSpans grouping", () => {
  test("the same reference twice shares a group", () => {
    const groups = findRefSpans("=A1+A1").map((s) => s.group);
    expect(groups).toEqual([0, 0]);
  });

  test("different references get different groups, numbered by appearance", () => {
    const spans = findRefSpans("=B2+A1+B2+C3");
    expect(spans.map((s) => s.text)).toEqual(["B2", "A1", "B2", "C3"]);
    expect(spans.map((s) => s.group)).toEqual([0, 1, 0, 2]);
  });

  test("references differing only in $ markers share a group", () => {
    // They address the same cell, so they should read as the same reference.
    expect(findRefSpans("=A1+$A$1").map((s) => s.group)).toEqual([0, 0]);
  });
});

describe("insertRefAtCaret", () => {
  test("inserts after an operator", () => {
    expect(insertRefAtCaret("=A1+", 4, "B2")).toMatchObject({
      value: "=A1+B2",
      caret: 6,
    });
  });

  test("inserts after a comma inside a call", () => {
    expect(insertRefAtCaret("=SUM(A1,", 8, "B2")?.value).toBe("=SUM(A1,B2");
  });

  test("inserts directly after the leading equals", () => {
    expect(insertRefAtCaret("=", 1, "B2")?.value).toBe("=B2");
  });

  test("inserts after an open paren", () => {
    expect(insertRefAtCaret("=SUM(", 5, "B2")?.value).toBe("=SUM(B2");
  });

  test("replaces the reference the caret sits at the end of", () => {
    expect(insertRefAtCaret("=A1", 3, "B2")).toMatchObject({
      value: "=B2",
      caret: 3,
    });
  });

  test("replaces the reference the caret sits inside", () => {
    expect(insertRefAtCaret("=SUM(A1,C3)", 6, "B2")?.value).toBe("=SUM(B2,C3)");
  });

  test("replacing keeps the rest of the formula intact", () => {
    expect(insertRefAtCaret("=A1*2+C3", 3, "Z9")?.value).toBe("=Z9*2+C3");
  });

  test("returns the span so a drag can keep rewriting it", () => {
    const first = insertRefAtCaret("=SUM(", 5, "B2");
    expect(first?.span).toEqual({ start: 5, end: 7 });
    // A drag extends the same span rather than appending.
    const dragged = insertRefAtCaret(
      first?.value as string,
      first?.caret as number,
      "B2:C4",
    );
    expect(dragged?.value).toBe("=SUM(B2:C4");
  });

  test("refuses when the caret follows a complete operand", () => {
    // `=A1+2|` — clicking a cell here means "select it", not "extend this".
    expect(insertRefAtCaret("=A1+2", 5, "B2")).toBeNull();
    expect(insertRefAtCaret("=SUM(A1)", 8, "B2")).toBeNull();
  });

  test("refuses on a non-formula", () => {
    expect(insertRefAtCaret("hello", 5, "B2")).toBeNull();
    expect(insertRefAtCaret("", 0, "B2")).toBeNull();
  });

  test("whitespace before the caret does not block an insertion", () => {
    expect(insertRefAtCaret("=A1 + ", 6, "B2")?.value).toBe("=A1 + B2");
  });

  test("a caret past the end is clamped rather than corrupting the source", () => {
    expect(insertRefAtCaret("=A1+", 99, "B2")?.value).toBe("=A1+B2");
  });
});

describe("isFormulaSource", () => {
  test("only a leading equals counts", () => {
    expect(isFormulaSource("=A1")).toBe(true);
    expect(isFormulaSource("A1")).toBe(false);
    expect(isFormulaSource("")).toBe(false);
  });
});
