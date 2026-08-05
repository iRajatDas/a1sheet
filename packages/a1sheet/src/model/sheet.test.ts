import { describe, expect, test } from "bun:test";
import {
  cloneSheet,
  deleteCol,
  deleteRow,
  getMergeAt,
  insertCol,
  insertRow,
  makeSheet,
} from "./sheet.js";

function sheetWith(cells: Record<string, string>) {
  const s = makeSheet("S");
  Object.assign(s.cells, cells);
  return s;
}

describe("cloneSheet", () => {
  test("does not share mutable containers with the original", () => {
    const a = makeSheet("A");
    a.cells["0_0"] = "x";
    a.hiddenRows.add(3);
    a.merges.push({ r1: 0, c1: 0, r2: 1, c2: 1 });

    const b = cloneSheet(a);
    b.cells["0_0"] = "y";
    b.hiddenRows.add(4);
    (b.merges[0] as { r1: number }).r1 = 9;

    expect(a.cells["0_0"]).toBe("x");
    expect(a.hiddenRows.has(4)).toBe(false);
    expect(a.merges[0]?.r1).toBe(0);
  });
});

describe("insertRow", () => {
  test("shifts rows at and below the insertion point", () => {
    const s = insertRow(sheetWith({ "0_0": "a", "1_0": "b", "2_0": "c" }), 1);
    expect(s.cells).toEqual({ "0_0": "a", "2_0": "b", "3_0": "c" });
    expect(s.numRows).toBe(201);
  });

  test("leaves other columns' row indices consistent", () => {
    const s = insertRow(sheetWith({ "1_0": "a", "1_5": "b" }), 0);
    expect(s.cells).toEqual({ "2_0": "a", "2_5": "b" });
  });
});

describe("deleteRow", () => {
  test("drops the row and pulls the rest up", () => {
    const s = deleteRow(sheetWith({ "0_0": "a", "1_0": "b", "2_0": "c" }), 1);
    expect(s.cells).toEqual({ "0_0": "a", "1_0": "c" });
    expect(s.numRows).toBe(199);
  });

  test("never shrinks below one row", () => {
    let s = makeSheet("S");
    s.numRows = 1;
    s = deleteRow(s, 0);
    expect(s.numRows).toBe(1);
  });
});

describe("insertCol / deleteCol", () => {
  test("insertCol shifts columns at and after the insertion point", () => {
    const s = insertCol(sheetWith({ "0_0": "a", "0_1": "b" }), 1);
    expect(s.cells).toEqual({ "0_0": "a", "0_2": "b" });
    expect(s.numCols).toBe(27);
  });

  test("deleteCol drops the column and pulls the rest left", () => {
    const s = deleteCol(sheetWith({ "0_0": "a", "0_1": "b", "0_2": "c" }), 1);
    expect(s.cells).toEqual({ "0_0": "a", "0_1": "c" });
    expect(s.numCols).toBe(25);
  });

  test("row and column shifts are independent axes", () => {
    const s = insertCol(sheetWith({ "5_5": "x" }), 0);
    expect(s.cells).toEqual({ "5_6": "x" });
  });
});

describe("styles shift alongside cells", () => {
  test("insertRow moves a style with its cell", () => {
    const base = makeSheet("S");
    base.cells["1_0"] = "a";
    base.styles["1_0"] = { bold: true };
    const s = insertRow(base, 0);
    expect(s.cells["2_0"]).toBe("a");
    expect(s.styles["2_0"]).toEqual({ bold: true });
    expect(s.styles["1_0"]).toBeUndefined();
  });
});

describe("getMergeAt", () => {
  test("finds the merge covering an interior cell", () => {
    const s = makeSheet("S");
    s.merges.push({ r1: 1, c1: 1, r2: 3, c2: 3 });
    expect(getMergeAt(s, 2, 2)).toEqual({ r1: 1, c1: 1, r2: 3, c2: 3 });
    expect(getMergeAt(s, 0, 0)).toBeUndefined();
  });
});
