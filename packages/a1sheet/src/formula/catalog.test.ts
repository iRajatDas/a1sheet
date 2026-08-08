import { describe, expect, test } from "bun:test";
import type { CellKey, RawCell } from "../model/types.js";
import {
  FORMULA_ARG_TYPES,
  FORMULA_CATALOG,
  FORMULA_CATEGORIES,
  formulaMeta,
  isImplementedFormula,
  suggestFormulas,
} from "./catalog.js";
import { createEvaluator } from "./evaluate.js";

function array(formula: string, cells: Record<string, string> = {}): unknown {
  const value = createEvaluator(
    cells as Record<CellKey, RawCell>,
    {},
  ).evaluateArray(formula);
  return Array.isArray(value) ? value.map((row) => [...row]) : value;
}

describe("formula catalog", () => {
  test("exposes the 16 Sheets categories", () => {
    expect(FORMULA_CATEGORIES).toHaveLength(16);
    expect(FORMULA_CATEGORIES[0]).toBe("DATE");
    expect(FORMULA_CATEGORIES[14]).toBe("ARRAY");
  });

  test("exposes the help-content argument types", () => {
    expect(FORMULA_ARG_TYPES).toContain("NUMBER");
    expect(FORMULA_ARG_TYPES).toContain("RANGE");
    expect(FORMULA_ARG_TYPES).toContain("SPARKLINE_SECOND_ARGUMENT");
  });

  test("catalog entries are implementable", () => {
    for (const entry of FORMULA_CATALOG) {
      expect(isImplementedFormula(entry.name)).toBe(true);
    }
  });

  test("suggestFormulas matches a prefix", () => {
    const hits = suggestFormulas("XL");
    expect(hits.map((h) => h.name)).toEqual(["XLOOKUP"]);
    expect(formulaMeta("xlookup")?.category).toBe("LOOKUP");
  });
});

describe("FLATTEN", () => {
  test("stacks multiple ranges into one column", () => {
    expect(
      array("FLATTEN(A1:B1,A2:B2)", {
        "0_0": "1",
        "0_1": "2",
        "1_0": "3",
        "1_1": "4",
      }),
    ).toEqual([[1], [2], [3], [4]]);
  });
});
