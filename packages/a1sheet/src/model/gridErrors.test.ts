import { describe, expect, test } from "bun:test";
import {
  filterIdExists,
  type GridError,
  isGridError,
  mergeSingleton,
  unboundedRange,
} from "./gridErrors.js";
import { makeSheet } from "./sheet.js";
import type { Range } from "./types.js";

function assertMergeMultiCell(range: Range): void {
  if (range.r1 === range.r2 && range.c1 === range.c2) {
    throw mergeSingleton(`${range.r1},${range.c1}`);
  }
}

function assertBounded(range: Range, label: string): void {
  if (
    !Number.isFinite(range.r1) ||
    !Number.isFinite(range.r2) ||
    !Number.isFinite(range.c1) ||
    !Number.isFinite(range.c2)
  ) {
    throw unboundedRange(label);
  }
}

describe("GridError", () => {
  test("merge of a single cell throws MERGE_RANGE_SINGLETON", () => {
    try {
      assertMergeMultiCell({ r1: 0, c1: 0, r2: 0, c2: 0 });
      throw new Error("should have thrown");
    } catch (e) {
      expect(isGridError(e)).toBe(true);
      expect((e as GridError).code).toBe("MERGE_RANGE_SINGLETON");
    }
  });

  test("a multi-cell merge is accepted", () => {
    expect(() =>
      assertMergeMultiCell({ r1: 0, c1: 0, r2: 1, c2: 0 }),
    ).not.toThrow();
  });

  test("unbounded range throws GRID_RANGE_UNBOUNDED", () => {
    try {
      assertBounded({ r1: 0, c1: 0, r2: Number.POSITIVE_INFINITY, c2: 0 }, "A1:∞");
      throw new Error("should have thrown");
    } catch (e) {
      expect((e as GridError).code).toBe("GRID_RANGE_UNBOUNDED");
    }
  });

  test("duplicate filter id throws FILTER_ID_EXISTS", () => {
    const err = filterIdExists("f1");
    expect(err.code).toBe("FILTER_ID_EXISTS");
  });

  test("consumers branch on code, not message", () => {
    const sheet = makeSheet("S");
    try {
      assertMergeMultiCell({ r1: 2, c1: 2, r2: 2, c2: 2 });
    } catch (e) {
      if (isGridError(e) && e.code === "MERGE_RANGE_SINGLETON") {
        expect(sheet.merges).toEqual([]);
        return;
      }
      throw e;
    }
  });
});
