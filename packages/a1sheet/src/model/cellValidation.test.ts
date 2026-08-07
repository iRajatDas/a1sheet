import { describe, expect, test } from "bun:test";
import { createEvaluator } from "../formula/evaluate.js";
import { rejectCellValue } from "./cellValidation.js";
import { makeSheet } from "./sheet.js";

describe("rejectCellValue", () => {
  test("rejects a list value that is not allowed", () => {
    const sheet = makeSheet("S");
    sheet.validations = [
      {
        range: { r1: 0, c1: 0, r2: 0, c2: 0 },
        kind: "list",
        formulas: ['"a,b,c"'],
      },
    ];
    const ev = createEvaluator(sheet.cells, {});
    expect(rejectCellValue(sheet, 0, 0, "z", ev)?.message).toContain("list");
    expect(rejectCellValue(sheet, 0, 0, "b", ev)).toBeUndefined();
  });

  test("allows blank when allowBlank is set", () => {
    const sheet = makeSheet("S");
    sheet.validations = [
      {
        range: { r1: 0, c1: 0, r2: 0, c2: 0 },
        kind: "list",
        formulas: ['"a,b"'],
        allowBlank: true,
      },
    ];
    const ev = createEvaluator(sheet.cells, {});
    expect(rejectCellValue(sheet, 0, 0, "", ev)).toBeUndefined();
  });
});
