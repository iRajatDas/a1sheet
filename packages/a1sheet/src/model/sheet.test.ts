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

describe("per-axis maps shift alongside cells", () => {
  // These are keyed by a bare row or column index rather than by cell key, and
  // they used to be left behind entirely: inserting a row above a resized one
  // moved the content down and left the height on the row that used to be there.

  test("insertRow carries a row's height, label, and hidden flag with it", () => {
    const base = makeSheet("S");
    base.rowHeights[1] = 80;
    base.rowLabels[1] = "total";
    base.hiddenRows.add(1);

    const s = insertRow(base, 0);

    expect(s.rowHeights[2]).toBe(80);
    expect(s.rowHeights[1]).toBeUndefined();
    expect(s.rowLabels[2]).toBe("total");
    expect(s.hiddenRows.has(2)).toBe(true);
    expect(s.hiddenRows.has(1)).toBe(false);
  });

  test("deleteRow drops the deleted row's height and pulls the rest up", () => {
    const base = makeSheet("S");
    base.rowHeights[1] = 80;
    base.rowHeights[2] = 40;

    const s = deleteRow(base, 1);

    expect(s.rowHeights[1]).toBe(40);
    expect(s.rowHeights[2]).toBeUndefined();
  });

  test("insertCol carries a column's width, label, and filter with it", () => {
    const base = makeSheet("S");
    base.colWidths[1] = 300;
    base.colLabels[1] = "Amount";
    base.filters[1] = new Set(["a"]);

    const s = insertCol(base, 0);

    expect(s.colWidths[2]).toBe(300);
    expect(s.colWidths[1]).toBeUndefined();
    expect(s.colLabels[2]).toBe("Amount");
    expect(s.filters[2]).toEqual(new Set(["a"]));
  });

  test("rows untouched by the insert keep their heights", () => {
    const base = makeSheet("S");
    base.rowHeights[0] = 80;

    const s = insertRow(base, 5);

    expect(s.rowHeights[0]).toBe(80);
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
