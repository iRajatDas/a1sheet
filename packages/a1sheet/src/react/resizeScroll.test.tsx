/**
 * Resize must not steal scroll to the selection. Ensure-visible is for
 * navigation; layout edits compensate (or leave) the offset instead.
 */
import { describe, expect, test } from "bun:test";
import { act, fireEvent, render } from "@testing-library/react";
import { createWorkbook } from "../model/workbook.js";
import { DEFAULT_COL_WIDTH, ROW_HEIGHT } from "./constants.js";
import { Sheet, useSheet } from "./index.js";
import type { UseSpreadsheetResult } from "./useSpreadsheet.js";

function mountWideSheet() {
  const wb = createWorkbook(["Sheet1"]);
  Object.assign(wb.sheets[0] as object, { numRows: 80, numCols: 40 });

  let api!: UseSpreadsheetResult;
  function Capture() {
    api = useSheet();
    return null;
  }

  const { container } = render(
    <Sheet.Root defaultWorkbook={wb} height={320}>
      <Capture />
      <Sheet.Grid style={{ height: 280 }} />
    </Sheet.Root>,
  );

  const scroller = container.querySelector(".a1s-scroller") as HTMLElement;
  if (!scroller) throw new Error("no scroller");

  return { api, scroller, container };
}

describe("resize while scrolled", () => {
  test("resizing an earlier column does not jump scroll to the selection", () => {
    const { api, scroller } = mountWideSheet();

    act(() => {
      api.select({ r1: 10, c1: 20, r2: 10, c2: 20 });
    });

    act(() => {
      scroller.scrollLeft = 900;
      scroller.scrollTop = 280;
      fireEvent.scroll(scroller);
    });

    const leftBefore = scroller.scrollLeft;
    const topBefore = scroller.scrollTop;
    expect(leftBefore).toBeGreaterThan(0);
    expect(topBefore).toBeGreaterThan(0);

    // Column 1 sits fully left of scrollLeft=900 (default width ~100).
    act(() => {
      api.setColWidth(1, DEFAULT_COL_WIDTH + 120);
    });

    expect(scroller.scrollLeft).toBe(leftBefore);
    expect(scroller.scrollTop).toBe(topBefore);
  });

  test("resizing an earlier row does not jump scroll to the selection", () => {
    const { api, scroller } = mountWideSheet();

    act(() => {
      api.select({ r1: 40, c1: 5, r2: 40, c2: 5 });
    });

    act(() => {
      scroller.scrollLeft = 200;
      scroller.scrollTop = 600;
      fireEvent.scroll(scroller);
    });

    const leftBefore = scroller.scrollLeft;
    const topBefore = scroller.scrollTop;

    act(() => {
      api.setRowHeight(2, ROW_HEIGHT + 40);
    });

    expect(scroller.scrollLeft).toBe(leftBefore);
    expect(scroller.scrollTop).toBe(topBefore);
  });

  test("selection change still reveals the active cell", () => {
    const { api, scroller } = mountWideSheet();

    Object.defineProperty(scroller, "clientHeight", {
      configurable: true,
      value: 280,
    });
    Object.defineProperty(scroller, "clientWidth", {
      configurable: true,
      value: 400,
    });

    act(() => {
      api.select({ r1: 0, c1: 0, r2: 0, c2: 0 });
    });

    // Same ensure-visible path as keyboard: focusRow / focusCol identity.
    act(() => {
      api.select({ r1: 50, c1: 0, r2: 50, c2: 0 });
    });

    expect(scroller.scrollTop).toBeGreaterThan(0);
  });
});
