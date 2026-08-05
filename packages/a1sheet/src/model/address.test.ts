import { describe, expect, test } from "bun:test";
import {
  cellKey,
  colToLetters,
  lettersToCol,
  normalizeRange,
  parseCellKey,
  parseCellRef,
  parseRangeRef,
  rangeContains,
  toA1,
} from "./address.js";

describe("colToLetters / lettersToCol", () => {
  test.each([
    [0, "A"],
    [1, "B"],
    [25, "Z"],
    [26, "AA"],
    [27, "AB"],
    [51, "AZ"],
    [52, "BA"],
    [701, "ZZ"],
    [702, "AAA"],
  ])("column %i is %s", (index, letters) => {
    expect(colToLetters(index)).toBe(letters);
    expect(lettersToCol(letters)).toBe(index);
  });

  test("round-trips across the whole 3-letter range", () => {
    for (let i = 0; i < 16384; i++) {
      expect(lettersToCol(colToLetters(i))).toBe(i);
    }
  });
});

describe("parseCellRef", () => {
  test("parses plain and absolute refs identically", () => {
    for (const ref of ["A1", "$A1", "A$1", "$A$1", "a1"]) {
      expect(parseCellRef(ref)).toEqual({ row: 0, col: 0 });
    }
  });

  test("is 0-indexed on both axes", () => {
    expect(parseCellRef("B3")).toEqual({ row: 2, col: 1 });
    expect(parseCellRef("AA10")).toEqual({ row: 9, col: 26 });
  });

  test("degrades to A1 rather than throwing on garbage", () => {
    expect(parseCellRef("not-a-ref")).toEqual({ row: 0, col: 0 });
    expect(parseCellRef("")).toEqual({ row: 0, col: 0 });
  });
});

describe("toA1", () => {
  test("is the inverse of parseCellRef", () => {
    for (const [row, col] of [
      [0, 0],
      [9, 26],
      [199, 25],
    ] as const) {
      expect(parseCellRef(toA1(row, col))).toEqual({ row, col });
    }
  });
});

describe("cellKey / parseCellKey", () => {
  test("builds the internal r_c form", () => {
    expect(cellKey(0, 0)).toBe("0_0");
    expect(cellKey(12, 3)).toBe("12_3");
  });

  test("round-trips multi-digit indices", () => {
    expect(parseCellKey(cellKey(123, 45))).toEqual({ row: 123, col: 45 });
  });
});

describe("normalizeRange", () => {
  test("reorders a range dragged up and left", () => {
    expect(normalizeRange({ r1: 5, c1: 5, r2: 1, c2: 2 })).toEqual({
      r1: 1,
      c1: 2,
      r2: 5,
      c2: 5,
    });
  });

  test("leaves an already-ordered range alone", () => {
    const r = { r1: 0, c1: 0, r2: 3, c2: 4 };
    expect(normalizeRange(r)).toEqual(r);
  });
});

describe("parseRangeRef", () => {
  test("parses a single cell as a degenerate range", () => {
    expect(parseRangeRef("B2")).toEqual({ r1: 1, c1: 1, r2: 1, c2: 1 });
  });

  test("parses a range and normalizes it", () => {
    expect(parseRangeRef("A1:B10")).toEqual({ r1: 0, c1: 0, r2: 9, c2: 1 });
    expect(parseRangeRef("B10:A1")).toEqual({ r1: 0, c1: 0, r2: 9, c2: 1 });
  });

  test("accepts absolute markers", () => {
    expect(parseRangeRef("$A$1:$B$2")).toEqual({ r1: 0, c1: 0, r2: 1, c2: 1 });
  });

  test("rejects non-references so the name box can fall through", () => {
    expect(parseRangeRef("TOTALS")).toBeNull();
    expect(parseRangeRef("A1:B2:C3")).toBeNull();
    expect(parseRangeRef("")).toBeNull();
  });
});

describe("rangeContains", () => {
  test("is inclusive on both bounds and order-independent", () => {
    const r = { r1: 3, c1: 3, r2: 1, c2: 1 };
    expect(rangeContains(r, 1, 1)).toBe(true);
    expect(rangeContains(r, 3, 3)).toBe(true);
    expect(rangeContains(r, 2, 2)).toBe(true);
    expect(rangeContains(r, 0, 2)).toBe(false);
    expect(rangeContains(r, 4, 2)).toBe(false);
  });
});
