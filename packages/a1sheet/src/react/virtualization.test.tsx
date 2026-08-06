/**
 * Virtualization: what gets rendered, and what the scrollbar reports.
 *
 * These are behaviour tests, not render tests — each one scrolls or filters and
 * then asserts on the DOM that results. The two properties they pin down are
 * independent and were each broken on their own:
 *
 *   1. Only cells near the viewport exist, on BOTH axes.
 *   2. The scroll extent describes the whole sheet anyway.
 *
 * Get (1) without (2) and the scrollbar shrinks to the size of the window into
 * the sheet, which is what a user sees as a "fake" scrollbar.
 */
import { describe, expect, test } from "bun:test";
import { fireEvent, render } from "@testing-library/react";
import type { Range } from "../model/types.js";
import { createWorkbook } from "../model/workbook.js";
import {
  DEFAULT_COL_WIDTH,
  HEADER_HEIGHT,
  MIN_ROW_HEIGHT,
  ROW_HEADER_WIDTH,
  ROW_HEIGHT,
} from "./constants.js";
import { Spreadsheet } from "./Spreadsheet.js";

interface SheetOverrides {
  cells?: Record<string, string>;
  numRows?: number;
  numCols?: number;
  frozenCols?: number;
  hiddenRows?: Set<number>;
  filters?: Record<number, Set<string>>;
  merges?: Range[];
  rowHeights?: Record<number, number>;
}

function setup(overrides: SheetOverrides = {}) {
  const wb = createWorkbook(["Sheet1"]);
  Object.assign(wb.sheets[0] as object, overrides);
  const { container } = render(<Spreadsheet defaultWorkbook={wb} />);

  const scroller = container.querySelector(".a1s-scroller") as HTMLElement;
  const grid = scroller.firstElementChild as HTMLElement;

  const cellAt = (row: number, col: number) =>
    container.querySelector(`.a1s-cell[data-row="${row}"][data-col="${col}"]`);

  const renderedCols = () =>
    [...container.querySelectorAll(".a1s-cell[data-row='0']")].map((el) =>
      Number((el as HTMLElement).dataset.col),
    );

  const renderedRows = () =>
    [
      ...new Set(
        [...container.querySelectorAll(".a1s-cell")].map((el) =>
          Number((el as HTMLElement).dataset.row),
        ),
      ),
    ].sort((a, b) => a - b);

  const scrollTo = (position: { top?: number; left?: number }) =>
    fireEvent.scroll(scroller, {
      target: {
        scrollTop: position.top ?? scroller.scrollTop,
        scrollLeft: position.left ?? scroller.scrollLeft,
      },
    });

  const rowHeaderAt = (row: number) =>
    container.querySelector(`.a1s-head[data-row="${row}"]`) as HTMLElement | null;

  const spacer = () => container.querySelector(".a1s-spacer") as HTMLElement;

  return {
    container,
    scroller,
    grid,
    cellAt,
    rowHeaderAt,
    spacer,
    renderedCols,
    renderedRows,
    scrollTo,
  };
}

describe("column virtualization", () => {
  test("a column past the right edge is not in the DOM", () => {
    const { cellAt } = setup({ numCols: 500 });

    expect(cellAt(0, 0)).not.toBeNull();
    expect(cellAt(0, 400)).toBeNull();
  });

  test("cell count does not grow with column count", () => {
    const narrow = setup({ numCols: 26 }).renderedCols().length;
    const wide = setup({ numCols: 500 }).renderedCols().length;

    expect(wide).toBe(narrow);
  });

  test("scrolling right renders the columns that came into view", () => {
    const { cellAt, scrollTo } = setup({ numCols: 500 });
    expect(cellAt(0, 100)).toBeNull();

    scrollTo({ left: 100 * DEFAULT_COL_WIDTH });

    expect(cellAt(0, 100)).not.toBeNull();
    // ...and drops the ones that left. A window that only ever grows is not a
    // window.
    expect(cellAt(0, 0)).toBeNull();
  });

  test("frozen columns stay rendered however far right you scroll", () => {
    const { cellAt, scrollTo } = setup({ numCols: 500, frozenCols: 2 });

    scrollTo({ left: 100 * DEFAULT_COL_WIDTH });

    expect(cellAt(0, 0)).not.toBeNull();
    expect(cellAt(0, 1)).not.toBeNull();
    expect(cellAt(0, 2)).toBeNull();
  });

  test("a merge whose origin scrolled out still renders", () => {
    // The merged block is drawn by its top-left cell. If the window drops that
    // cell the whole block disappears, leaving a hole rather than a wide cell.
    const { cellAt, scrollTo } = setup({
      numCols: 500,
      merges: [{ r1: 0, c1: 0, r2: 0, c2: 200 }],
    });

    scrollTo({ left: 50 * DEFAULT_COL_WIDTH });

    expect(cellAt(0, 0)).not.toBeNull();
  });
});

describe("row virtualization", () => {
  test("cell count does not grow with row count", () => {
    const small = setup({ numRows: 1_000 }).renderedRows().length;
    const huge = setup({ numRows: 100_000 }).renderedRows().length;

    expect(huge).toBe(small);
  });
});

describe("scroll extent", () => {
  test("the grid is as tall as the sheet, not as tall as the window", () => {
    const numRows = 100_000;
    const { grid } = setup({ numRows });

    expect(grid.style.minHeight).toBe(`${HEADER_HEIGHT + numRows * ROW_HEIGHT}px`);
  });

  test("the grid is as wide as the sheet, and stays that wide once scrolled", () => {
    const numCols = 500;
    const expected = `${ROW_HEADER_WIDTH + numCols * DEFAULT_COL_WIDTH}px`;
    const { grid, scrollTo } = setup({ numCols });

    expect(grid.style.minWidth).toBe(expected);

    // The horizontal extent used to be whatever the rendered cells reached, so
    // it followed the window rightward instead of describing the sheet.
    scrollTo({ left: 100 * DEFAULT_COL_WIDTH });
    expect(grid.style.minWidth).toBe(expected);
  });

  test("hidden rows shorten the sheet, because they are not laid out", () => {
    const numRows = 1_000;
    const hiddenRows = new Set([1, 2, 3, 4, 5]);
    const { grid } = setup({ numRows, hiddenRows });

    expect(grid.style.minHeight).toBe(
      `${HEADER_HEIGHT + (numRows - hiddenRows.size) * ROW_HEIGHT}px`,
    );
  });
});

describe("variable row heights", () => {
  test("a resized row is laid out at its own height", () => {
    const { rowHeaderAt } = setup({ rowHeights: { 2: 80 } });

    expect(rowHeaderAt(2)?.style.height).toBe("80px");
    expect(rowHeaderAt(3)?.style.height).toBe(`${ROW_HEIGHT}px`);
  });

  test("the scroll extent counts each row's own height", () => {
    const numRows = 1_000;
    const { grid } = setup({ numRows, rowHeights: { 0: 100, 1: 50 } });

    // Two rows are 100 and 50 instead of the default; the rest are unchanged.
    const expected = HEADER_HEIGHT + (numRows - 2) * ROW_HEIGHT + 100 + 50;
    expect(grid.style.minHeight).toBe(`${expected}px`);
  });

  test("the window is placed by summed heights, not by row index", () => {
    // Row 0 is 500px tall, so a row index is no longer proportional to a pixel
    // offset. The spacer standing in for the skipped rows has to be the sum of
    // their heights; anything derived from `index * ROW_HEIGHT` lands 474px out.
    const { spacer, renderedRows, scrollTo } = setup({
      numRows: 1_000,
      rowHeights: { 0: 500 },
    });

    scrollTo({ top: HEADER_HEIGHT + 500 + 20 * ROW_HEIGHT });

    const first = renderedRows()[0] as number;
    expect(first).toBeGreaterThan(0);
    expect(Number.parseInt(spacer().style.height, 10)).toBe(
      500 + (first - 1) * ROW_HEIGHT,
    );
  });

  test("dragging the row grabber resizes that row", () => {
    const { container, rowHeaderAt } = setup();
    const grabber = rowHeaderAt(1)?.querySelector(".a1s-rowresize") as HTMLElement;

    fireEvent.mouseDown(grabber, { clientY: 100 });
    fireEvent.mouseMove(window, { clientY: 140 });
    fireEvent.mouseUp(window);

    expect(rowHeaderAt(1)?.style.height).toBe(`${ROW_HEIGHT + 40}px`);
    // The drag must not have started a rename on the header it lives in.
    expect(container.querySelector(".a1s-head input")).toBeNull();
  });

  test("a resize below the minimum stops at the minimum", () => {
    const { rowHeaderAt } = setup();
    const grabber = rowHeaderAt(1)?.querySelector(".a1s-rowresize") as HTMLElement;

    fireEvent.mouseDown(grabber, { clientY: 100 });
    fireEvent.mouseMove(window, { clientY: -500 });
    fireEvent.mouseUp(window);

    expect(rowHeaderAt(1)?.style.height).toBe(`${MIN_ROW_HEIGHT}px`);
  });

  test("double-clicking the row grabber restores the default height", () => {
    const { rowHeaderAt } = setup({ rowHeights: { 1: 90 } });
    const grabber = rowHeaderAt(1)?.querySelector(".a1s-rowresize") as HTMLElement;

    fireEvent.doubleClick(grabber);

    expect(rowHeaderAt(1)?.style.height).toBe(`${ROW_HEIGHT}px`);
  });

  test("a merged block is as tall as the rows it covers", () => {
    const { cellAt } = setup({
      rowHeights: { 0: 40, 1: 60 },
      merges: [{ r1: 0, c1: 0, r2: 1, c2: 1 }],
    });

    expect((cellAt(0, 0) as HTMLElement).style.height).toBe("100px");
  });
});

describe("filtering", () => {
  test("rows whose value is not allowed are dropped, empty ones included", () => {
    // The scan resolves an empty cell's display once per column rather than
    // once per row. Empty rows must still be filtered out by it.
    const { cellAt, grid } = setup({
      numRows: 100,
      cells: { "0_0": "keep", "1_0": "drop", "2_0": "keep" },
      filters: { 0: new Set(["keep"]) },
    });

    expect(cellAt(0, 0)).not.toBeNull();
    expect(cellAt(1, 0)).toBeNull();
    expect(cellAt(2, 0)).not.toBeNull();
    expect(grid.style.minHeight).toBe(`${HEADER_HEIGHT + 2 * ROW_HEIGHT}px`);
  });

  test("a filter that allows the empty value keeps the empty rows", () => {
    const { cellAt } = setup({
      numRows: 100,
      cells: { "0_0": "keep", "1_0": "drop" },
      filters: { 0: new Set(["keep", ""]) },
    });

    expect(cellAt(0, 0)).not.toBeNull();
    expect(cellAt(1, 0)).toBeNull();
    expect(cellAt(2, 0)).not.toBeNull();
  });
});
