import { describe, expect, test } from "bun:test";
import {
  createWorkbook,
  defineName,
  deleteSheet,
  duplicateSheet,
  moveSheet,
  namedRangeAddedStatus,
} from "./workbook.js";

describe("sheet tab ops", () => {
  test("duplicateSheet deep-copies filters and activates the copy", () => {
    let wb = createWorkbook(["A", "B"]);
    wb.sheets[0]!.cells["0_0"] = "x";
    wb.sheets[0]!.filters[0] = { values: new Set(["x"]) };
    wb = duplicateSheet(wb, 0);
    expect(wb.sheets).toHaveLength(3);
    expect(wb.activeSheetIndex).toBe(1);
    expect(wb.sheets[1]!.name).toBe("A (copy)");
    expect(wb.sheets[1]!.cells["0_0"]).toBe("x");
    expect(wb.sheets[1]!.filters[0]?.values).toEqual(new Set(["x"]));
    expect(wb.sheets[1]!.id).not.toBe(wb.sheets[0]!.id);
  });

  test("deleteSheet refuses to delete the last sheet", () => {
    const wb = createWorkbook(["Only"]);
    expect(deleteSheet(wb, 0).sheets).toHaveLength(1);
  });

  test("moveSheet reorders and keeps the active sheet", () => {
    let wb = createWorkbook(["A", "B", "C"]);
    wb = { ...wb, activeSheetIndex: 1 };
    wb = moveSheet(wb, 0, 2);
    expect(wb.sheets.map((s) => s.name)).toEqual(["B", "C", "A"]);
    expect(wb.activeSheetIndex).toBe(0);
  });

  test("named range status", () => {
    const wb = defineName(createWorkbook(), "Revenue", {
      r1: 0,
      c1: 0,
      r2: 10,
      c2: 0,
    });
    expect(wb.namedRanges.REVENUE).toBeDefined();
    expect(namedRangeAddedStatus("Revenue")).toBe("Named range REVENUE added");
  });
});
