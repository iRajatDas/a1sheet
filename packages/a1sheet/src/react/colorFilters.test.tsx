/**
 * Colour-criteria filters and named filter views through the public API.
 */
import { describe, expect, test } from "bun:test";
import { act, render } from "@testing-library/react";
import { createRef } from "react";
import { createWorkbook } from "../model/workbook.js";
import { Spreadsheet } from "./Spreadsheet.js";
import { type SheetRootHandle } from "./Root.js";

describe("colour filters", () => {
  test("hides rows whose fill is not in the allowed set", () => {
    const wb = createWorkbook(["Sheet1"]);
    Object.assign(wb.sheets[0] as object, {
      numRows: 20,
      cells: { "0_0": "a", "1_0": "b", "2_0": "c" },
      styles: {
        "0_0": { bg: "#ff0000" },
        "1_0": { bg: "#00ff00" },
        "2_0": { bg: "#ff0000" },
      },
    });
    const ref = createRef<SheetRootHandle>();
    const { container } = render(
      <Spreadsheet defaultWorkbook={wb} ref={ref} />,
    );
    const cellAt = (row: number) =>
      container.querySelector(`.a1s-cell[data-row="${row}"][data-col="0"]`);

    act(() => {
      ref.current?.api.setFilter(0, {
        background: new Set(["#ff0000"]),
      });
    });

    expect(cellAt(0)).not.toBeNull();
    expect(cellAt(1)).toBeNull();
    expect(cellAt(2)).not.toBeNull();
  });

  test("a style change re-filters without a full cache miss on values", () => {
    const wb = createWorkbook(["Sheet1"]);
    Object.assign(wb.sheets[0] as object, {
      numRows: 20,
      cells: { "0_0": "a", "1_0": "b" },
      styles: {
        "0_0": { bg: "#ff0000" },
        "1_0": { bg: "#00ff00" },
      },
      filters: { 0: { background: new Set(["#ff0000"]) } },
    });
    const ref = createRef<SheetRootHandle>();
    const { container } = render(
      <Spreadsheet defaultWorkbook={wb} ref={ref} />,
    );
    const cellAt = (row: number) =>
      container.querySelector(`.a1s-cell[data-row="${row}"][data-col="0"]`);

    expect(cellAt(1)).toBeNull();

    act(() => {
      ref.current?.api.select({ r1: 1, c1: 0, r2: 1, c2: 0 });
    });
    act(() => {
      ref.current?.api.applyStyle({ bg: "#ff0000" });
    });

    expect(ref.current?.api.sheet.styles["1_0"]?.bg).toBe("#ff0000");
    expect(cellAt(1)).not.toBeNull();
  });
});

describe("filter views", () => {
  test("create / activate / missing status", () => {
    const wb = createWorkbook(["Sheet1"]);
    Object.assign(wb.sheets[0] as object, {
      numRows: 20,
      cells: { "0_0": "a", "1_0": "b" },
      filters: { 0: { values: new Set(["a"]) } },
    });
    const ref = createRef<SheetRootHandle>();
    render(<Spreadsheet defaultWorkbook={wb} ref={ref} />);

    act(() => {
      ref.current?.api.createFilterView({ id: "keep-a", name: "Keep A" });
      ref.current?.api.setFilter(0, null);
    });
    expect(ref.current?.api.sheet.filters[0]).toBeUndefined();

    act(() => {
      ref.current?.api.activateFilterView("keep-a");
    });
    expect(ref.current?.api.sheet.filters[0]?.values).toEqual(new Set(["a"]));
    expect(ref.current?.api.sheet.activeFilterViewId).toBe("keep-a");

    act(() => {
      ref.current?.api.activateFilterView("missing");
    });
    expect(ref.current?.api.status).toBe("The view does not exist.");
  });

  test("sortByColor reports status and reorders", () => {
    const wb = createWorkbook(["Sheet1"]);
    Object.assign(wb.sheets[0] as object, {
      cells: { "0_0": "a", "1_0": "b", "2_0": "c" },
      styles: {
        "0_0": { color: "#111" },
        "1_0": { color: "#f00" },
        "2_0": { color: "#f00" },
      },
    });
    const ref = createRef<SheetRootHandle>();
    render(<Spreadsheet defaultWorkbook={wb} ref={ref} />);

    act(() => {
      ref.current?.api.sortByColor({
        col: 0,
        kind: "foreground",
        color: "#f00",
      });
    });

    expect(ref.current?.api.getRaw(0, 0)).toBe("b");
    expect(ref.current?.api.status).toBe(
      "Cells with foreground color #f00 were moved to the top",
    );
  });
});
