/**
 * Dynamic arrays: values that are a shape rather than a number.
 *
 * This is what a modern Excel workbook is made of. `SORT(UNIQUE(...))` replaces a
 * helper column and `MAP` replaces a filled-down formula, so a file using them
 * has no intermediate cells to fall back on — without arrays it is not partly
 * readable, it is a grid of `#NAME?`.
 */
import { describe, expect, test } from "bun:test";
import type { CellKey, RawCell } from "../model/types.js";
import { createEvaluator } from "./evaluate.js";

const GRID = {
  "0_0": "3",
  "1_0": "1",
  "2_0": "2",
  "3_0": "1",
  "0_1": "c",
  "1_1": "a",
  "2_1": "b",
  "3_1": "a",
} as Record<CellKey, RawCell>;

function evaluator(cells: Record<CellKey, RawCell> = GRID) {
  return createEvaluator(cells, {});
}

/** The whole array a formula produces, as plain rows. */
function array(formula: string, cells?: Record<CellKey, RawCell>) {
  const value = evaluator(cells).evaluateArray(formula);
  return Array.isArray(value) ? value.map((row) => [...row]) : value;
}

const one = (formula: string, cells?: Record<CellKey, RawCell>) =>
  evaluator(cells).evaluate(formula);

describe("arrays are two-dimensional", () => {
  test("a column range is a column, not a flat list", () => {
    // Excel has no 1D array. Collapsing a column to a list is what makes
    // TRANSPOSE and HSTACK impossible to express.
    expect(array("A1:A3")).toEqual([[3], [1], [2]]);
    expect(array("A1:B1")).toEqual([[3, "c"]]);
  });

  test("an array literal keeps its rows", () => {
    expect(array("{1,2;3,4}")).toEqual([
      [1, 2],
      [3, 4],
    ]);
  });

  test("a numeric cell reads back as a number", () => {
    // `=A1` used to give the string "3", so arithmetic went through a coercion
    // that turned any non-numeric text into 0.
    expect(one("A1")).toBe(3);
    expect(one("B1")).toBe("c");
  });
});

describe("operators broadcast", () => {
  test("two columns combine elementwise", () => {
    expect(array("A1:A3*2")).toEqual([[6], [2], [4]]);
  });

  test("a comparison over a column yields a column of results", () => {
    // `(teams = team) * played` is the shape every SUMPRODUCT-style formula
    // takes. Using the top-left of each side returns one number where the
    // formula meant several — an answer, and the wrong one.
    expect(array('B1:B3="a"')).toEqual([[0], [1], [0]]);
    expect(one('SUM((B1:B4="a")*A1:A4)')).toBe(2);
  });

  test("a single value stretches to the other operand's shape", () => {
    expect(array("{1;2;3}+10")).toEqual([[11], [12], [13]]);
  });

  test("a row and a column make a grid", () => {
    expect(array("{1,2}*{10;20}")).toEqual([
      [10, 20],
      [20, 40],
    ]);
  });

  test("mismatched extents pad with #N/A rather than failing", () => {
    expect(array("{1;2;3}+{10;20}")).toEqual([[11], [22], ["#N/A"]]);
  });

  test("an error on either side is the result", () => {
    // Without this, `IF(COUNTBLANK(bad)=0, 1, 0)` compared #NAME? against 0,
    // found them unequal, and returned 0 — a guard turning an error into a
    // plausible answer.
    expect(one("NOSUCH(1)=0")).toBe("#NAME?");
    expect(one("IF(NOSUCH(1)=0,1,0)")).toBe("#NAME?");
  });
});

describe("generating and reshaping", () => {
  test("SEQUENCE builds a column by default and a grid on request", () => {
    expect(array("SEQUENCE(3)")).toEqual([[1], [2], [3]]);
    expect(array("SEQUENCE(2,2)")).toEqual([
      [1, 2],
      [3, 4],
    ]);
    expect(array("SEQUENCE(2,2,10,5)")).toEqual([
      [10, 15],
      [20, 25],
    ]);
  });

  test("TRANSPOSE swaps the axes", () => {
    expect(array("TRANSPOSE({1,2;3,4})")).toEqual([
      [1, 3],
      [2, 4],
    ]);
  });

  test("HSTACK and VSTACK assemble, padding the short side", () => {
    expect(array("HSTACK({1;2},{3;4})")).toEqual([
      [1, 3],
      [2, 4],
    ]);
    expect(array("VSTACK({1,2},{3,4})")).toEqual([
      [1, 2],
      [3, 4],
    ]);
    expect(array("VSTACK({1,2},{3})")).toEqual([
      [1, 2],
      [3, "#N/A"],
    ]);
  });

  test("CHOOSECOLS and CHOOSEROWS pick, negatives from the end", () => {
    expect(array("CHOOSECOLS({1,2,3;4,5,6},3,1)")).toEqual([
      [3, 1],
      [6, 4],
    ]);
    expect(array("CHOOSEROWS({1;2;3},-1)")).toEqual([[3]]);
  });

  test("TAKE and DROP trim from either end", () => {
    expect(array("TAKE({1;2;3},2)")).toEqual([[1], [2]]);
    expect(array("TAKE({1;2;3},-1)")).toEqual([[3]]);
    expect(array("DROP({1;2;3},1)")).toEqual([[2], [3]]);
  });

  test("ROWS and COLUMNS measure", () => {
    expect(one("ROWS(A1:A3)")).toBe(3);
    expect(one("COLUMNS({1,2,3})")).toBe(3);
  });
});

describe("filtering and ordering", () => {
  test("UNIQUE keeps first-seen order", () => {
    // Which is what makes SORT(UNIQUE(...)) produce a stable list.
    expect(array('UNIQUE({"b";"a";"b";"c"})')).toEqual([["b"], ["a"], ["c"]]);
  });

  test("SORT puts numbers before text and blanks last", () => {
    expect(array('SORT({"b";1;"a"})')).toEqual([[1], ["a"], ["b"]]);
  });

  test("SORT takes a key column and a direction", () => {
    expect(array('SORT({1,"b";3,"a";2,"c"},2)')).toEqual([
      [3, "a"],
      [1, "b"],
      [2, "c"],
    ]);
    expect(array("SORT({1;3;2},1,-1)")).toEqual([[3], [2], [1]]);
  });

  test("FILTER keeps the rows its mask marks", () => {
    expect(array("FILTER({1;2;3},{1;0;1})")).toEqual([[1], [3]]);
  });

  test("FILTER with nothing left uses its fallback", () => {
    expect(one('FILTER({1;2},{0;0},"none")')).toBe("none");
    expect(one("FILTER({1;2},{0;0})")).toBe("#CALC!");
  });

  test("SORTBY orders one array by another", () => {
    expect(array('SORTBY({"a";"b";"c"},{3;1;2})')).toEqual([["b"], ["c"], ["a"]]);
  });
});

describe("LET and LAMBDA", () => {
  test("LET binds names, and each value sees the ones before it", () => {
    expect(one("LET(x, 2, y, x*3, x+y)")).toBe(8);
  });

  test("a LET binding shadows a function-free name", () => {
    expect(one("LET(total, 5, total)")).toBe(5);
  });

  test("MAP applies a lambda elementwise", () => {
    expect(array("MAP({1;2;3}, LAMBDA(v, v*2))")).toEqual([[2], [4], [6]]);
  });

  test("MAP takes one array per lambda parameter", () => {
    expect(array("MAP({1;2},{10;20}, LAMBDA(a,b, a+b))")).toEqual([[11], [22]]);
  });

  test("MAKEARRAY is given 1-based indices", () => {
    expect(array("MAKEARRAY(2,2, LAMBDA(r,c, r*10+c))")).toEqual([
      [11, 12],
      [21, 22],
    ]);
  });

  test("BYROW and BYCOL reduce along one axis", () => {
    expect(array("BYROW({1,2;3,4}, LAMBDA(r, SUM(r)))")).toEqual([[3], [7]]);
    expect(array("BYCOL({1,2;3,4}, LAMBDA(c, SUM(c)))")).toEqual([[4, 6]]);
  });

  test("REDUCE folds and SCAN keeps the intermediates", () => {
    expect(one("REDUCE(0,{1;2;3}, LAMBDA(acc,v, acc+v))")).toBe(6);
    expect(array("SCAN(0,{1;2;3}, LAMBDA(acc,v, acc+v))")).toEqual([[1], [3], [6]]);
  });

  test("a lambda sees the LET bindings it was defined under", () => {
    expect(one("LET(n, 10, REDUCE(0, {1;2}, LAMBDA(a,v, a+v*n)))")).toBe(30);
  });

  test("a lambda in a cell position is an error, not a stringified function", () => {
    expect(one("LAMBDA(x, x)")).toBe("#VALUE!");
  });
});

describe("the lazy forms", () => {
  test("IF does not evaluate the branch it does not take", () => {
    // The whole point of a guard. `IF(A1<>0, B1/A1, 0)` must not divide by zero
    // just because both branches were computed before dispatch.
    expect(one("IF(FALSE, 1/0, 99)")).toBe(99);
  });

  test("IFERROR catches, and does not fire when there is nothing to catch", () => {
    expect(one('IFERROR(1/0, "caught")')).toBe("caught");
    expect(one('IFERROR(4, "caught")')).toBe(4);
  });

  test("IFS takes the first true branch", () => {
    expect(one("IFS(FALSE, 1, TRUE, 2, TRUE, 3)")).toBe(2);
    expect(one("IFS(FALSE, 1)")).toBe("#N/A");
  });

  test("SWITCH matches, with a trailing default", () => {
    expect(one('SWITCH(2, 1, "one", 2, "two", "other")')).toBe("two");
    expect(one('SWITCH(9, 1, "one", "other")')).toBe("other");
  });

  test("an IF over an array condition is an array operation", () => {
    expect(array("IF({1;0;1}, 10, 20)")).toEqual([[10], [20], [10]]);
  });
});

describe("XLOOKUP", () => {
  test("finds a value and returns the matching entry", () => {
    expect(one('XLOOKUP("b",{"a";"b";"c"},{1;2;3})')).toBe(2);
  });

  test("returns its fallback rather than #N/A when told to", () => {
    expect(one('XLOOKUP("z",{"a";"b"},{1;2},"missing")')).toBe("missing");
    expect(one('XLOOKUP("z",{"a";"b"},{1;2})')).toBe("#N/A");
  });

  test("matches exactly by default, unlike VLOOKUP", () => {
    // The default that made XLOOKUP worth having.
    expect(one('XLOOKUP("bb",{"a";"b";"c"},{1;2;3})')).toBe("#N/A");
    expect(one('XLOOKUP("bb",{"a";"b";"c"},{1;2;3},"none",-1)')).toBe(2);
  });

  test("searches from the end when asked", () => {
    expect(one('XLOOKUP("a",{"a";"b";"a"},{1;2;3},"none",0,-1)')).toBe(3);
  });
});

describe("spilling", () => {
  const spilling = {
    "0_0": "=SEQUENCE(3)",
    "0_2": "=TRANSPOSE({1,2})",
  } as Record<CellKey, RawCell>;

  test("an array formula fills the cells below and beside its anchor", () => {
    const ev = evaluator(spilling);
    expect(ev.getCellDisplay(0, 0)).toBe(1);
    expect(ev.getCellDisplay(1, 0)).toBe(2);
    expect(ev.getCellDisplay(2, 0)).toBe(3);
    // …and stops.
    expect(ev.getCellDisplay(3, 0)).toBe("");
  });

  test("a row-shaped array spills sideways", () => {
    const ev = evaluator(spilling);
    expect(ev.getCellDisplay(0, 2)).toBe(1);
    expect(ev.getCellDisplay(1, 2)).toBe(2);
  });

  test("a cell in the way blocks the spill", () => {
    // Excel reports #SPILL! on the anchor rather than overwriting, and silently
    // truncating would lose data the formula meant to produce.
    const blocked = {
      "0_0": "=SEQUENCE(3)",
      "1_0": "in the way",
    } as Record<CellKey, RawCell>;
    const ev = evaluator(blocked);

    expect(ev.getCellDisplay(0, 0)).toBe("#SPILL!");
    expect(ev.getCellDisplay(1, 0)).toBe("in the way");
  });

  test("one array reads another's spilled tail", () => {
    // The whole point of chaining. A1 spills 1,2,3 down; A2 and A3 hold no text
    // of their own, so SORT sees them only if the index resolves A1 first.
    const chained = {
      "0_0": "=SEQUENCE(3)",
      "0_2": "=SORT(A1:A3, 1, -1)",
    } as Record<CellKey, RawCell>;
    const ev = evaluator(chained);

    expect(ev.getCellDisplay(0, 2)).toBe(3);
    expect(ev.getCellDisplay(1, 2)).toBe(2);
    expect(ev.getCellDisplay(2, 2)).toBe(1);
  });

  test("the chain holds however the cells are read first", () => {
    // Reading the tail before the anchor used to cache the blank it found.
    const chained = {
      "0_0": "=SEQUENCE(3)",
      "0_2": "=SUM(A1:A3)",
    } as Record<CellKey, RawCell>;
    const ev = evaluator(chained);

    expect(ev.getCellDisplay(1, 0)).toBe(2);
    expect(ev.getCellDisplay(0, 2)).toBe(6);
  });

  test("a chain three deep resolves", () => {
    const chained = {
      "0_0": "=SEQUENCE(3)",
      "0_2": "=SORT(A1:A3, 1, -1)",
      "0_4": "=SORT(C1:C3)",
    } as Record<CellKey, RawCell>;
    const ev = evaluator(chained);

    expect(ev.getCellDisplay(0, 4)).toBe(1);
    expect(ev.getCellDisplay(2, 4)).toBe(3);
  });

  test("two arrays that each read the other's tail terminate", () => {
    // A genuine reference cycle. It must not hang or blow the stack; one of the
    // two reads blanks, which is what Excel calls a circular reference.
    const mutual = {
      "0_0": "=SORT(C1:C2)",
      "0_2": "=SORT(A1:A2)",
    } as Record<CellKey, RawCell>;
    const ev = evaluator(mutual);

    expect(() => ev.getCellDisplay(0, 0)).not.toThrow();
    expect(() => ev.getCellDisplay(0, 2)).not.toThrow();
  });

  test("a spill on another sheet is visible across the reference", () => {
    const here = { "0_0": "=SUM(Data!A1:A3)" } as Record<CellKey, RawCell>;
    const ev = createEvaluator(
      here,
      {},
      {
        sheets: [
          {
            name: "Data",
            cells: { "0_0": "=SEQUENCE(3)" } as Record<CellKey, RawCell>,
          },
        ],
      },
    );

    expect(ev.getCellDisplay(0, 0)).toBe(6);
  });

  test("a formula that cannot spill costs no index", () => {
    // The index has to evaluate every candidate, so narrowing the candidates is
    // what keeps it from being a full recalculation. Asserted by behaviour: a
    // sheet of plain arithmetic still answers for its empty cells.
    const plain = { "0_0": "1", "0_1": "=A1*2" } as Record<CellKey, RawCell>;
    expect(evaluator(plain).getCellDisplay(5, 5)).toBe("");
  });
});
