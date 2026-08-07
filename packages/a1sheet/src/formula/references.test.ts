/**
 * The three ways a formula reaches cells it does not sit next to: structured
 * references into a table, qualified references into another sheet, and defined
 * names that hold a formula.
 *
 * All three appear together in a modern workbook — a name defined once, used on
 * a second sheet, reading a table on a third — so each is useless without the
 * others.
 */
import { describe, expect, test } from "bun:test";
import type { CellKey, RawCell } from "../model/types.js";
import { createEvaluator } from "./evaluate.js";
import { tableIndex } from "./tableRefs.js";

/** A three-column table with a header row, at A1:C4 of a sheet called Data. */
const DATA = {
  "0_0": "team",
  "0_1": "won",
  "0_2": "drawn",
  "1_0": "Arsenal",
  "1_1": "20",
  "1_2": "14",
  "2_0": "Chelsea",
  "2_1": "20",
  "2_2": "9",
  "3_0": "Fulham",
  "3_1": "15",
  "3_2": "9",
} as Record<CellKey, RawCell>;

const TABLES = tableIndex([
  {
    name: "tblTeams",
    sheet: "Data",
    range: { r1: 0, c1: 0, r2: 3, c2: 2 },
    columns: ["team", "won", "drawn"],
    headerRow: true,
  },
]);

/** An evaluator on `Data` itself, so a this-row reference has a row to mean. */
function onData(cells: Record<CellKey, RawCell> = DATA) {
  return createEvaluator(
    cells,
    {},
    {
      tables: TABLES,
      sheets: [{ name: "Data", cells: DATA }],
    },
  );
}

describe("structured references", () => {
  const ev = onData();
  const array = (f: string) => {
    const v = ev.evaluateArray(f);
    return Array.isArray(v) ? v.map((r) => [...r]) : v;
  };

  test("a column name selects that column's data, not its header", () => {
    expect(array("tblTeams[won]")).toEqual([[20], [20], [15]]);
  });

  test("brackets around a single name are optional", () => {
    expect(array("tblTeams[[won]]")).toEqual([[20], [20], [15]]);
  });

  test("a span covers the columns between its endpoints", () => {
    expect(array("tblTeams[[won]:[drawn]]")).toEqual([
      [20, 14],
      [20, 9],
      [15, 9],
    ]);
  });

  test("#Headers and #All address the parts a column selector skips", () => {
    expect(array("tblTeams[[#Headers],[won]]")).toEqual([["won"]]);
    expect(array("tblTeams[#All]")).toHaveLength(4);
  });

  test("the whole table with no column named is every data column", () => {
    expect(array("tblTeams[#Data]")).toEqual([
      ["Arsenal", 20, 14],
      ["Chelsea", 20, 9],
      ["Fulham", 15, 9],
    ]);
  });

  test("a name the table does not have is #REF!, not a silent empty range", () => {
    expect(ev.evaluate("tblTeams[nosuch]")).toBe("#REF!");
  });

  test("an unknown table is #NAME?", () => {
    expect(ev.evaluate("nosuchtable[col]")).toBe("#NAME?");
  });

  test("the name is matched case-insensitively, as Excel matches it", () => {
    expect(array("TBLTEAMS[WON]")).toEqual([[20], [20], [15]]);
  });

  test("SUM over a column works, which is the common use", () => {
    expect(ev.evaluate("SUM(tblTeams[won])")).toBe(55);
  });
});

describe("[#This Row]", () => {
  test("resolves against the row the formula is on", () => {
    const cells = {
      ...DATA,
      "1_3": "=SUM(tblTeams[[#This Row],[won]:[drawn]])",
      "2_3": "=SUM(tblTeams[[#This Row],[won]:[drawn]])",
    } as Record<CellKey, RawCell>;
    const ev = onData(cells);

    expect(ev.getCellDisplay(1, 3)).toBe(34);
    expect(ev.getCellDisplay(2, 3)).toBe(29);
  });

  test("the @ shorthand means the same thing", () => {
    const cells = { ...DATA, "1_3": "=tblTeams[@won]" } as Record<CellKey, RawCell>;
    expect(onData(cells).getCellDisplay(1, 3)).toBe(20);
  });

  test("outside the table it has no row to mean", () => {
    const cells = { ...DATA, "9_3": "=tblTeams[@won]" } as Record<CellKey, RawCell>;
    expect(onData(cells).getCellDisplay(9, 3)).toBe("#REF!");
  });
});

describe("cross-sheet references", () => {
  const other = { "0_0": "7", "1_0": "8" } as Record<CellKey, RawCell>;
  const here = {
    "0_0": "=Data!B2",
    "1_0": "=SUM(Other!A1:A2)",
    "2_0": "='My Sheet'!A1",
    "3_0": "=Nope!A1",
  } as Record<CellKey, RawCell>;

  const ev = createEvaluator(
    here,
    {},
    {
      sheets: [
        { name: "Data", cells: DATA },
        { name: "Other", cells: other },
        { name: "My Sheet", cells: other },
      ],
    },
  );

  test("a qualified cell reads the other sheet", () => {
    expect(ev.getCellDisplay(0, 0)).toBe(20);
  });

  test("a qualified range does too", () => {
    expect(ev.getCellDisplay(1, 0)).toBe(15);
  });

  test("a quoted name works, which is how a name with a space is written", () => {
    // Unquoted, `My Sheet!A1` would lex as a name, a name, and a ref — which is
    // how a cross-sheet formula was silently wrong rather than merely refused.
    expect(ev.getCellDisplay(2, 0)).toBe(7);
  });

  test("an unknown sheet is #REF!", () => {
    expect(ev.getCellDisplay(3, 0)).toBe("#REF!");
  });
});

describe("defined names holding a formula", () => {
  const ev = createEvaluator(
    {} as Record<CellKey, RawCell>,
    {},
    {
      tables: TABLES,
      sheets: [{ name: "Data", cells: DATA }],
      namedFormulas: {
        TOTALWON: "SUM(tblTeams[won])",
        // One name built on another, which is how these are usually written.
        SUMMARY: "TOTALWON*2",
        SELFREF: "SELFREF+1",
        TEAMLIST: "SORT(tblTeams[team])",
      },
    },
  );

  test("a name whose value is a formula evaluates it", () => {
    expect(ev.evaluate("TOTALWON")).toBe(55);
  });

  test("a name may be defined in terms of another", () => {
    expect(ev.evaluate("SUMMARY")).toBe(110);
  });

  test("a name that returns an array keeps its shape", () => {
    const value = ev.evaluateArray("TEAMLIST");
    expect(Array.isArray(value) && value.map((r) => [...r])).toEqual([
      ["Arsenal"],
      ["Chelsea"],
      ["Fulham"],
    ]);
  });

  test("a self-referencing name reports a cycle instead of recursing", () => {
    expect(ev.evaluate("SELFREF")).toBe("#CYCLE!");
  });

  test("an undefined name is still #NAME?", () => {
    expect(ev.evaluate("NOSUCHNAME")).toBe("#NAME?");
  });
});

describe("references built from a computation", () => {
  const grid = {
    "0_0": "1",
    "1_0": "2",
    "2_0": "3",
    "0_1": "10",
    "1_1": "20",
    "2_1": "30",
  } as Record<CellKey, RawCell>;
  const ev = createEvaluator(
    grid,
    {},
    {
      sheets: [
        { name: "Other", cells: { "0_0": "99" } as Record<CellKey, RawCell> },
      ],
    },
  );

  test("OFFSET moves a reference and can resize it", () => {
    // Not an ordinary function: `OFFSET(A1,1,0)` means "the cell below A1", so
    // it has to produce a region rather than a value that has been moved.
    expect(ev.evaluate("OFFSET(A1,1,0)")).toBe(2);
    expect(ev.evaluate("SUM(OFFSET(A1,0,0,3,1))")).toBe(6);
    expect(ev.evaluate("SUM(OFFSET(A1,0,0,3,2))")).toBe(66);
  });

  test("OFFSET off the sheet is #REF!", () => {
    expect(ev.evaluate("OFFSET(A1,-1,0)")).toBe("#REF!");
    expect(ev.evaluate("OFFSET(A1,0,0,0,1)")).toBe("#REF!");
  });

  test("INDIRECT turns text into a reference, including a qualified one", () => {
    expect(ev.evaluate('INDIRECT("B2")')).toBe(20);
    expect(ev.evaluate('SUM(INDIRECT("A1:A3"))')).toBe(6);
    expect(ev.evaluate('INDIRECT("Other!A1")')).toBe(99);
  });

  test("INDIRECT on something that is not an address is #REF!", () => {
    expect(ev.evaluate('INDIRECT("nonsense")')).toBe("#REF!");
  });
});
