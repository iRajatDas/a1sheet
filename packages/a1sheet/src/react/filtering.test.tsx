/**
 * Filtering, and the cache that keeps re-filtering cheap.
 *
 * The cache re-tests only rows whose raw text changed, which is sound only
 * while a displayed value cannot change without its raw text changing. These
 * tests are mostly about the cases where that assumption fails — a formula in a
 * filtered column, a change to the filter itself, a change to number formatting
 * — because a stale filter shows the user rows that should not be there and it
 * does so silently.
 */
import { describe, expect, test } from "bun:test";
import { fireEvent, render, screen } from "@testing-library/react";
import { createWorkbook } from "../model/workbook.js";
import { Spreadsheet } from "./Spreadsheet.js";

interface Fixture {
  cells?: Record<string, string>;
  filters?: Record<number, Set<string>>;
  numRows?: number;
}

function setup(fixture: Fixture) {
  const wb = createWorkbook(["Sheet1"]);
  Object.assign(wb.sheets[0] as object, { numRows: 50, ...fixture });
  const { container } = render(<Spreadsheet defaultWorkbook={wb} />);

  const cellAt = (row: number, col: number) =>
    container.querySelector(`.a1s-cell[data-row="${row}"][data-col="${col}"]`);

  const visibleRows = () =>
    [
      ...new Set(
        [...container.querySelectorAll(".a1s-cell")].map((el) =>
          Number((el as HTMLElement).dataset.row),
        ),
      ),
    ].sort((a, b) => a - b);

  const textarea = container.querySelector("textarea") as HTMLTextAreaElement;

  /** Selects a cell and commits `value` into it, the way a user types. */
  const type = (row: number, col: number, value: string) => {
    fireEvent.mouseDown(cellAt(row, col) as HTMLElement);
    fireEvent.keyDown(textarea, { key: "x" });
    const input = container.querySelector(".a1s-cell input") as HTMLInputElement;
    fireEvent.change(input, { target: { value } });
    fireEvent.keyDown(input, { key: "Enter" });
  };

  return { container, cellAt, visibleRows, type };
}

describe("an edit re-filters the row it changed", () => {
  test("a row edited out of the allowed set disappears", () => {
    const { cellAt, type } = setup({
      cells: { "0_0": "keep", "1_0": "keep", "2_0": "keep" },
      filters: { 0: new Set(["keep"]) },
    });
    expect(cellAt(1, 0)).not.toBeNull();

    type(1, 0, "drop");

    expect(cellAt(1, 0)).toBeNull();
    expect(cellAt(0, 0)).not.toBeNull();
    expect(cellAt(2, 0)).not.toBeNull();
  });

  test("a row edited to another allowed value stays", () => {
    const { cellAt, type } = setup({
      cells: { "0_0": "a", "1_0": "b", "2_0": "a" },
      filters: { 0: new Set(["a", "b"]) },
    });

    type(1, 0, "a");

    expect(cellAt(1, 0)).not.toBeNull();
    expect(cellAt(0, 0)).not.toBeNull();
    expect(cellAt(2, 0)).not.toBeNull();
  });

  test("an edit in an unfiltered column leaves the filter alone", () => {
    const { cellAt, visibleRows, type } = setup({
      cells: { "0_0": "keep", "1_0": "drop", "2_0": "keep" },
      filters: { 0: new Set(["keep"]) },
    });
    const before = visibleRows();

    type(0, 3, "anything");

    expect(visibleRows()).toEqual(before);
    expect(cellAt(1, 0)).toBeNull();
  });
});

describe("cases the raw-text shortcut cannot cover", () => {
  // Each formula here points into column B, which carries no filter. That keeps
  // the cell driving the formula visible and clickable however the filter
  // resolves — a hidden row cannot be edited, so the dependency has to live
  // somewhere the filter will not take away.

  test("a formula in a filtered column re-filters when its input changes", () => {
    // A2 holds "=B1". Its display follows B1 while its own raw text never
    // changes, so a cache keyed on raw text alone would leave row 1 showing.
    const { cellAt, type } = setup({
      cells: { "0_0": "5", "0_1": "5", "1_0": "=B1" },
      filters: { 0: new Set(["5"]) },
    });
    expect(cellAt(1, 0)).not.toBeNull();

    type(0, 1, "9");

    expect(cellAt(1, 0)).toBeNull();
    expect(cellAt(0, 0)).not.toBeNull();
  });

  test("a formula that starts passing the filter reappears", () => {
    const { cellAt, type } = setup({
      cells: { "1_0": "=B6", "5_0": "5", "5_1": "9" },
      filters: { 0: new Set(["5"]) },
    });
    expect(cellAt(1, 0)).toBeNull();
    expect(cellAt(5, 1)).not.toBeNull();

    type(5, 1, "5");

    expect(cellAt(1, 0)).not.toBeNull();
  });

  test("typing a formula into a filtered column marks it volatile from then on", () => {
    const { cellAt, type } = setup({
      cells: { "0_0": "5", "0_1": "5", "1_0": "5" },
      filters: { 0: new Set(["5"]) },
    });

    // Until this edit the column is plain text and takes the fast path.
    type(1, 0, "=B1");
    expect(cellAt(1, 0)).not.toBeNull();

    // Now the formula's input changes. Row 1 only re-filters if committing the
    // formula switched the column over to full rescans.
    type(0, 1, "9");
    expect(cellAt(1, 0)).toBeNull();
  });
});

describe("changing the filter itself", () => {
  test("widening the allowed set brings rows back", () => {
    const wb = createWorkbook(["Sheet1"]);
    Object.assign(wb.sheets[0] as object, {
      numRows: 20,
      cells: { "0_0": "a", "1_0": "b" },
      filters: { 0: new Set(["a"]) },
    });
    const { container } = render(<Spreadsheet defaultWorkbook={wb} />);
    const cellAt = (row: number, col: number) =>
      container.querySelector(`.a1s-cell[data-row="${row}"][data-col="${col}"]`);

    expect(cellAt(1, 0)).toBeNull();

    // Clear the filter through the column menu, the way the UI does.
    fireEvent.click(
      screen.getByLabelText("Sort and filter column A") as HTMLElement,
    );
    fireEvent.click(screen.getByText("Clear filter"));

    expect(cellAt(1, 0)).not.toBeNull();
  });
});
