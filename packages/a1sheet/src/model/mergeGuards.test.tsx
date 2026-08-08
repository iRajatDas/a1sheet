import { describe, expect, test } from "bun:test";
import {
  checkFilterMerge,
  checkPasteMerge,
  checkSortMerge,
} from "../model/mergeGuards.js";
import { makeSheet } from "../model/sheet.js";
import { act, render } from "@testing-library/react";
import type { CellKey, Sheet, Workbook } from "../model/types.js";
import { useSpreadsheet } from "../react/useSpreadsheet.js";

describe("merge guards", () => {
  test("paste blocks a destination that cuts a merge", () => {
    const sheet = makeSheet("S");
    sheet.merges = [{ r1: 1, c1: 1, r2: 3, c2: 2 }];
    const result = checkPasteMerge(sheet, {
      r1: 0,
      c1: 1,
      r2: 2,
      c2: 1,
    });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.code).toBe("PASTE_PARTIAL_MERGE");
  });

  test("sort blocks a range that partially covers a merge", () => {
    const sheet = makeSheet("S");
    sheet.merges = [{ r1: 0, c1: 0, r2: 2, c2: 1 }];
    const result = checkSortMerge(sheet, { r1: 0, c1: 0, r2: 5, c2: 0 });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.code).toBe("SORT_PARTIAL_MERGE");
  });

  test("filter blocks a range that cuts a merge", () => {
    const sheet = makeSheet("S");
    sheet.merges = [{ r1: 0, c1: 0, r2: 2, c2: 0 }];
    const result = checkFilterMerge(sheet, { r1: 0, c1: 0, r2: 1, c2: 0 });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.code).toBe("FILTER_PARTIAL_MERGE");
  });
});

function workbook(sheet: Partial<Sheet>): Workbook {
  return {
    sheets: [{ ...makeSheet("S"), ...sheet }],
    activeSheetIndex: 0,
    namedRanges: {},
  };
}

describe("merge guards via public API", () => {
  test("paste that cuts a merge sets status and writes nothing", () => {
    let api: ReturnType<typeof useSpreadsheet> | undefined;
    function Probe() {
      api = useSpreadsheet({
        initialWorkbook: workbook({
          cells: { "0_0": "x" } as Record<CellKey, string>,
          merges: [{ r1: 1, c1: 0, r2: 3, c2: 0 }],
        }),
      });
      return null;
    }
    render(<Probe />);
    act(() => {
      api!.clipboard.paste("a\nb\nc", { row: 0, col: 0 }, api!.updateSheet, {
        onReject: api!.setStatus,
      });
    });
    expect(api!.getRaw(1, 0)).toBe("");
    expect(api!.status).toContain("partially intersects a merge");
  });

  test("sort that cuts a merge sets status and leaves order", () => {
    let api: ReturnType<typeof useSpreadsheet> | undefined;
    function Probe() {
      api = useSpreadsheet({
        initialWorkbook: workbook({
          cells: {
            "0_0": "b",
            "1_0": "a",
            "2_0": "c",
          } as Record<CellKey, string>,
          // Extends past the used rows the sort covers — partial relative to the
          // sort range.
          merges: [{ r1: 1, c1: 0, r2: 5, c2: 0 }],
        }),
      });
      return null;
    }
    render(<Probe />);
    act(() => {
      api!.select({ r1: 0, c1: 0, r2: 2, c2: 0 });
      api!.sort(0, "asc");
    });
    expect(api!.getRaw(0, 0)).toBe("b");
    expect(api!.status).toContain("Merges must be entirely within");
  });
});
