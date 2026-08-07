/**
 * The frozen band, which is `position: sticky` inside the one scroll container
 * rather than a second pane.
 *
 * That makes it cheap and makes it correct under variable row heights, but it
 * also means the band and the rows sliding under it are the same surface — so
 * the two things that tell them apart, the layering and the edge line, are the
 * whole feature and both are asserted here.
 */
import { describe, expect, test } from "bun:test";
import { render } from "@testing-library/react";
import { createWorkbook } from "../model/workbook.js";
import { Spreadsheet } from "./Spreadsheet.js";

function setup(frozen: { frozenRows?: number; frozenCols?: number }) {
  const wb = createWorkbook(["Sheet1"]);
  Object.assign(wb.sheets[0] as object, {
    cells: { "0_0": "head", "3_0": "body" },
    ...frozen,
  });
  const { container } = render(<Spreadsheet defaultWorkbook={wb} />);

  const cellAt = (row: number, col: number) =>
    container.querySelector(
      `.a1s-cell[data-row="${row}"][data-col="${col}"]`,
    ) as HTMLElement;
  const rowHeader = (row: number) =>
    container.querySelector(`.a1s-head[data-row="${row}"]`) as HTMLElement;
  const colHeader = (col: number) =>
    container.querySelector(`.a1s-head[data-col="${col}"]`) as HTMLElement;

  const layer = (el: HTMLElement) => Number(el.style.zIndex || "0");

  return { container, cellAt, rowHeader, colHeader, layer };
}

describe("what covers what", () => {
  test("the frozen row's own number sits above the rows sliding past it", () => {
    // Both are sticky row headers in the same place. At equal layers the
    // scrolled row won, because it comes later in the DOM — so the frozen row
    // was labelled with whatever number happened to be passing.
    const { rowHeader, layer } = setup({ frozenRows: 1 });

    expect(layer(rowHeader(0))).toBeGreaterThan(layer(rowHeader(3)));
  });

  test("an ordinary row header still covers the frozen cells it scrolls past", () => {
    // The other direction of the same overlap: frozen COLUMNS slide under the
    // row header horizontally, and must go under it.
    const { rowHeader, cellAt, layer } = setup({ frozenRows: 1, frozenCols: 1 });

    expect(layer(rowHeader(3))).toBeGreaterThan(layer(cellAt(0, 0)));
  });

  test("the column header covers everything, and the corner covers the column header", () => {
    const { rowHeader, colHeader, layer } = setup({ frozenRows: 1 });
    const corner = document.querySelector(
      ".a1s-head:not([data-row]):not([data-col])",
    );

    expect(layer(colHeader(0))).toBeGreaterThan(layer(rowHeader(0)));
    expect(layer(corner as HTMLElement)).toBeGreaterThan(layer(colHeader(0)));
  });
});

describe("the line marking the edge", () => {
  test("the last frozen row carries it, along its bottom", () => {
    const { cellAt } = setup({ frozenRows: 2 });

    expect(cellAt(1, 0).style.boxShadow).toContain("0 2px 0 0");
    expect(cellAt(0, 0).style.boxShadow).toBe("");
  });

  test("the row header of that row carries it too, so the line is unbroken", () => {
    const { rowHeader } = setup({ frozenRows: 1 });

    expect(rowHeader(0).style.boxShadow).toContain("0 2px 0 0");
  });

  test("the last frozen column carries it down its right-hand side", () => {
    const { cellAt } = setup({ frozenCols: 2 });

    expect(cellAt(3, 1).style.boxShadow).toContain("2px 0 0 0");
    expect(cellAt(3, 0).style.boxShadow).toBe("");
  });

  test("the corner cell of both bands carries both", () => {
    const { cellAt } = setup({ frozenRows: 1, frozenCols: 1 });
    const shadow = cellAt(0, 0).style.boxShadow;

    expect(shadow).toContain("0 2px 0 0");
    expect(shadow).toContain("2px 0 0 0");
  });

  test("an unfrozen sheet draws no line anywhere", () => {
    const { cellAt } = setup({});

    expect(cellAt(0, 0).style.boxShadow).toBe("");
    expect(cellAt(3, 0).style.boxShadow).toBe("");
  });
});
