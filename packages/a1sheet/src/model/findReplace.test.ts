import { describe, expect, test } from "bun:test";
import {
  findAll,
  findNext,
  replaceAll,
  replacedStatus,
} from "./findReplace.js";
import { makeSheet } from "./sheet.js";

describe("findReplace", () => {
  test("findAll matches case-insensitively by default", () => {
    const sheet = makeSheet("S");
    sheet.cells["0_0"] = "Hello";
    sheet.cells["1_0"] = "HELLO";
    sheet.cells["2_0"] = "nope";
    expect(findAll(sheet, { find: "hello" })).toHaveLength(2);
  });

  test("matchEntireCell and matchCase", () => {
    const sheet = makeSheet("S");
    sheet.cells["0_0"] = "cat";
    sheet.cells["1_0"] = "caterpillar";
    sheet.cells["2_0"] = "Cat";
    expect(findAll(sheet, { find: "cat", matchEntireCell: true })).toHaveLength(
      2,
    );
    expect(
      findAll(sheet, { find: "Cat", matchCase: true, matchEntireCell: true }),
    ).toHaveLength(1);
  });

  test("findNext walks then wraps", () => {
    const sheet = makeSheet("S");
    sheet.cells["0_0"] = "x";
    sheet.cells["0_1"] = "x";
    sheet.cells["1_0"] = "x";
    const first = findNext(sheet, { find: "x" });
    expect(first).toEqual({ row: 0, col: 0, raw: "x" });
    const second = findNext(sheet, { find: "x", after: first! });
    expect(second).toEqual({ row: 0, col: 1, raw: "x" });
    const wrap = findNext(sheet, { find: "x", after: { row: 1, col: 0 } });
    expect(wrap).toEqual({ row: 0, col: 0, raw: "x" });
  });

  test("replaceAll rewrites formulas and reports count", () => {
    const sheet = makeSheet("S");
    sheet.cells["0_0"] = "=A1+1";
    sheet.cells["1_0"] = "A1";
    const { sheet: next, count } = replaceAll(sheet, {
      find: "A1",
      replace: "B1",
    });
    expect(count).toBe(2);
    expect(next.cells["0_0"]).toBe("=B1+1");
    expect(next.cells["1_0"]).toBe("B1");
    expect(replacedStatus(count)).toBe("Replaced 2 occurrences.");
  });

  test("searchFormulas false skips formula cells", () => {
    const sheet = makeSheet("S");
    sheet.cells["0_0"] = "=SUM(1)";
    sheet.cells["1_0"] = "SUM";
    expect(
      findAll(sheet, { find: "SUM", searchFormulas: false }),
    ).toHaveLength(1);
  });
});
