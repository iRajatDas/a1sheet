/**
 * The stories are thin — a workbook and a tree of primitives — so the part that
 * can quietly rot is the fixture data: a formula that does not parse, a merge
 * outside the sheet, a filter naming a value no row has. A broken fixture makes
 * a story look like a broken library.
 *
 * The trees themselves are covered by the library's own composition tests.
 */
import { describe, expect, test } from "bun:test";
import { createEvaluator, isErrorValue } from "a1sheet";
import * as fixtures from "./fixtures.js";

const WORKBOOKS = [
  ["budget", fixtures.budget()],
  ["formulas", fixtures.formulas()],
  ["layout", fixtures.layout()],
  ["filtered", fixtures.filtered()],
  ["salesReport", fixtures.salesReport()],
] as const;

describe("story fixtures", () => {
  for (const [name, wb] of WORKBOOKS) {
    test(`${name} is a well-formed workbook`, () => {
      expect(wb.sheets.length).toBeGreaterThan(0);
      for (const sheet of wb.sheets) {
        expect(sheet.numRows).toBeGreaterThan(0);
        expect(sheet.numCols).toBeGreaterThan(0);

        for (const m of sheet.merges) {
          expect(m.r2).toBeLessThan(sheet.numRows);
          expect(m.c2).toBeLessThan(sheet.numCols);
        }
        // A cell outside the sheet's extent is invisible, which reads as a
        // fixture that silently lost half its data.
        for (const key of Object.keys(sheet.cells)) {
          const [r, c] = key.split("_").map(Number) as [number, number];
          expect(r).toBeLessThan(sheet.numRows);
          expect(c).toBeLessThan(sheet.numCols);
        }
      }
    });
  }

  test("only the cells meant to be errors evaluate to errors", () => {
    // The formulas story deliberately shows #DIV/0!, #CYCLE! and #NAME?. Every
    // other formula in every other fixture must actually compute.
    const broken: string[] = [];
    for (const [name, wb] of WORKBOOKS) {
      if (name === "formulas") continue;
      for (const sheet of wb.sheets) {
        const evaluator = createEvaluator(sheet.cells, wb.namedRanges);
        for (const [key, raw] of Object.entries(sheet.cells)) {
          if (!raw.startsWith("=")) continue;
          const [r, c] = key.split("_").map(Number) as [number, number];
          const value = evaluator.getCellDisplay(r, c);
          if (isErrorValue(value))
            broken.push(`${name} ${key}: ${raw} -> ${value}`);
        }
      }
    }
    expect(broken).toEqual([]);
  });

  test("the formulas story really does produce the errors it advertises", () => {
    const sheet = fixtures.formulas().sheets[0];
    if (!sheet) throw new Error("no sheet");
    const evaluator = createEvaluator(sheet.cells, {});

    expect(isErrorValue(evaluator.getCellDisplay(8, 0))).toBe(true);
    expect(isErrorValue(evaluator.getCellDisplay(9, 0))).toBe(true);
    expect(isErrorValue(evaluator.getCellDisplay(10, 0))).toBe(true);
  });

  test("the filter names values that rows actually have", () => {
    // A filter matching nothing renders an empty sheet, which looks like a bug
    // in virtualization rather than a deliberate fixture.
    const sheet = fixtures.filtered().sheets[0];
    if (!sheet) throw new Error("no sheet");

    for (const [col, allowed] of Object.entries(sheet.filters)) {
      const present = new Set<string>();
      for (let r = 0; r < sheet.numRows; r++) {
        const value = sheet.cells[`${r}_${Number(col)}`];
        if (value !== undefined) present.add(value);
      }
      for (const value of allowed) expect(present.has(value)).toBe(true);
    }
  });

  test("the large fixture is large in rows, not in DOM-sized promises", () => {
    const sheet = fixtures.large(100_000).sheets[0];
    if (!sheet) throw new Error("no sheet");
    expect(sheet.numRows).toBe(100_001);
    expect(Object.keys(sheet.cells).length).toBe(100_000 * 4 + 4);
  });
});
