import { describe, expect, test } from "bun:test";
import type { CellKey, NamedRanges, RawCell } from "../model/types.js";
import { CYCLE_ERROR, createEvaluator } from "./evaluate.js";

/** Evaluates one formula against an optional sheet of literals. */
function evalOne(
  formula: string,
  cells: Record<string, string> = {},
  namedRanges: NamedRanges = {},
) {
  const all = { ...cells, "99_99": formula } as Record<CellKey, RawCell>;
  return createEvaluator(all, namedRanges).getCellDisplay(99, 99);
}

const GRID = {
  "0_0": "1",
  "1_0": "2",
  "2_0": "3",
  "0_1": "10",
  "1_1": "20",
  "2_1": "30",
};

describe("arithmetic", () => {
  test.each([
    ["=1+2", 3],
    ["=10-4", 6],
    ["=3*4", 12],
    ["=10/4", 2.5],
    ["=2^10", 1024],
    ["=-5+2", -3],
    ["=2+3*4", 14],
    ["=(2+3)*4", 20],
    ["=2^3^2", 512], // ^ is right-associative
  ])("%s = %p", (formula, expected) => {
    expect(evalOne(formula)).toBe(expected);
  });

  test("division by zero is an error value, not Infinity", () => {
    expect(evalOne("=1/0")).toBe("#DIV/0!");
  });
});

describe("comparisons", () => {
  test.each([
    ["=1=1", 1],
    ["=1=2", 0],
    ["=2>1", 1],
    ["=1<2", 1],
    ["=2>=2", 1],
    ["=2<=1", 0],
    ["=1<>2", 1],
    ["=1<>1", 0],
  ])("%s = %p", (formula, expected) => {
    expect(evalOne(formula)).toBe(expected);
  });

  test("<> works end to end (was silently broken in the POC)", () => {
    expect(evalOne("=A1<>A2", GRID)).toBe(1);
    expect(evalOne("=A1<>A1", GRID)).toBe(0);
  });

  test("falls back to text comparison when not both numeric", () => {
    expect(evalOne('="abc"="abc"')).toBe(1);
    expect(evalOne('="abc"="abd"')).toBe(0);
  });
});

describe("references", () => {
  test("reads a single cell", () => {
    expect(evalOne("=A1", GRID)).toBe(1);
    expect(evalOne("=B3", GRID)).toBe(30);
  });

  test("absolute markers do not change the value read", () => {
    expect(evalOne("=$A$1", GRID)).toBe(1);
    expect(evalOne("=A$1", GRID)).toBe(1);
  });

  test("an empty cell reads as 0 in arithmetic", () => {
    expect(evalOne("=Z50+1", GRID)).toBe(1);
  });

  test("resolves chains of formulas", () => {
    const cells = { "0_0": "5", "1_0": "=A1*2", "2_0": "=A2+1" };
    expect(
      createEvaluator(cells as Record<CellKey, RawCell>, {}).getCellDisplay(2, 0),
    ).toBe(11);
  });
});

describe("ranges and functions", () => {
  test.each([
    ["=SUM(A1:A3)", 6],
    ["=AVERAGE(A1:A3)", 2],
    ["=MIN(A1:B3)", 1],
    ["=MAX(A1:B3)", 30],
    ["=COUNT(A1:A3)", 3],
    ["=SUM(A1:A3,B1:B3)", 66],
  ])("%s = %p", (formula, expected) => {
    expect(evalOne(formula, GRID)).toBe(expected);
  });

  test("COUNTA counts non-numeric text but not blanks", () => {
    const cells = { "0_0": "a", "1_0": "", "2_0": "7" };
    expect(evalOne("=COUNTA(A1:A3)", cells)).toBe(2);
  });

  test("text functions", () => {
    expect(evalOne('=UPPER("abc")')).toBe("ABC");
    expect(evalOne('=LEN("hello")')).toBe(5);
    expect(evalOne('=LEFT("hello",2)')).toBe("he");
    expect(evalOne('=RIGHT("hello",2)')).toBe("lo");
    expect(evalOne('=MID("hello",2,3)')).toBe("ell");
    expect(evalOne('=CONCAT("a","b","c")')).toBe("abc");
  });

  test("logical functions", () => {
    expect(evalOne("=IF(1>0,10,20)")).toBe(10);
    expect(evalOne("=IF(1<0,10,20)")).toBe(20);
    // Booleans, not 1 and 0. Excel's AND returns TRUE, and the grid renders a
    // boolean as TRUE/FALSE — returning a number showed "1" where Excel shows
    // "TRUE", and made ISLOGICAL impossible to write.
    expect(evalOne("=AND(1,1)")).toBe(true);
    expect(evalOne("=AND(1,0)")).toBe(false);
    expect(evalOne("=OR(0,1)")).toBe(true);
    expect(evalOne("=NOT(0)")).toBe(true);
  });

  test("an unknown function is #NAME?", () => {
    expect(evalOne("=NOPE(1)")).toBe("#NAME?");
  });
});

describe("lookup functions", () => {
  const table = {
    "0_0": "apple",
    "0_1": "1.50",
    "1_0": "banana",
    "1_1": "0.75",
    "2_0": "cherry",
    "2_1": "4.00",
  };

  test("VLOOKUP finds a text match", () => {
    // A numeric cell reads back as a number now. `=A1` on a cell holding "0.75"
    // used to evaluate to the STRING "0.75", so arithmetic on it went through a
    // coercion that turned any non-numeric text into 0.
    expect(evalOne('=VLOOKUP("banana",A1:B3,2)', table)).toBe(0.75);
  });

  test("VLOOKUP approximates by default, as Excel does", () => {
    // The fourth argument defaults to TRUE in Excel, meaning "or the next
    // smaller" — so "durian" finds cherry's row rather than reporting #N/A. A
    // well-known footgun, kept because an imported formula has to compute what
    // Excel computes.
    expect(evalOne('=VLOOKUP("durian",A1:B3,2)', table)).toBe(4);
    expect(evalOne('=VLOOKUP("durian",A1:B3,2,FALSE)', table)).toBe("#N/A");
  });

  test("MATCH returns a 1-indexed position", () => {
    expect(evalOne('=MATCH("cherry",A1:A3)', table)).toBe(3);
    // Type defaults to 1 — approximate — for the same reason as VLOOKUP.
    expect(evalOne('=MATCH("durian",A1:A3,0)', table)).toBe("#N/A");
  });

  test("INDEX is 1-indexed on both axes", () => {
    expect(evalOne("=INDEX(A1:B3,2,1)", table)).toBe("banana");
    expect(evalOne("=INDEX(A1:B3,9,1)", table)).toBe("#REF!");
  });

  test("INDEX/MATCH compose", () => {
    // 4, not "4.00": the cell's text is a number and reads back as one.
    expect(evalOne('=INDEX(B1:B3,MATCH("cherry",A1:A3))', table)).toBe(4);
  });
});

describe("named ranges", () => {
  test("resolves a name to its range", () => {
    const names: NamedRanges = { TOTALS: { r1: 0, c1: 0, r2: 2, c2: 0 } };
    expect(evalOne("=SUM(TOTALS)", GRID, names)).toBe(6);
  });

  test("an undefined name is #NAME?", () => {
    expect(evalOne("=SUM(NOPE)", GRID)).toBe("#NAME?");
  });
});

describe("cycles and errors", () => {
  test("a self-reference is #CYCLE!", () => {
    const cells = { "0_0": "=A1" } as Record<CellKey, RawCell>;
    expect(createEvaluator(cells, {}).getCellDisplay(0, 0)).toBe(CYCLE_ERROR);
  });

  test("a mutual reference is #CYCLE!", () => {
    const cells = { "0_0": "=A2", "1_0": "=A1" } as Record<CellKey, RawCell>;
    expect(createEvaluator(cells, {}).getCellDisplay(0, 0)).toBe(CYCLE_ERROR);
  });

  test("a longer cycle is caught", () => {
    const cells = {
      "0_0": "=A2",
      "1_0": "=A3",
      "2_0": "=A1",
    } as Record<CellKey, RawCell>;
    expect(createEvaluator(cells, {}).getCellDisplay(0, 0)).toBe(CYCLE_ERROR);
  });

  test("literals pass through untouched", () => {
    const cells = { "0_0": "hello" } as Record<CellKey, RawCell>;
    expect(createEvaluator(cells, {}).getCellDisplay(0, 0)).toBe("hello");
  });

  test("an empty cell is the empty string", () => {
    expect(createEvaluator({}, {}).getCellDisplay(5, 5)).toBe("");
  });
});

describe("memoization", () => {
  test("a repeated read does not recompute", () => {
    // TODAY() is stable within one evaluator because results are cached.
    const ev = createEvaluator(
      { "0_0": "=NOW()", "1_0": "=A1" } as Record<CellKey, RawCell>,
      {},
    );
    expect(ev.getCellDisplay(0, 0)).toBe(ev.getCellDisplay(0, 0));
    expect(ev.getCellDisplay(1, 0)).toBe(ev.getCellDisplay(0, 0));
  });
});
