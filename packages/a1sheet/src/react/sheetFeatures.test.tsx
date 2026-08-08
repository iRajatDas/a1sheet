/**
 * Find/replace, checkboxes, hyperlinks, text rotation via the public API.
 */
import { describe, expect, test } from "bun:test";
import { act, render } from "@testing-library/react";
import { createRef } from "react";
import { createWorkbook } from "../model/workbook.js";
import type { SheetRootHandle } from "./Root.js";
import { Spreadsheet } from "./Spreadsheet.js";

describe("find and replace", () => {
  test("replaceAll updates cells and status", () => {
    const wb = createWorkbook(["S"]);
    Object.assign(wb.sheets[0] as object, {
      cells: { "0_0": "foo", "1_0": "food", "2_0": "bar" },
    });
    const ref = createRef<SheetRootHandle>();
    render(<Spreadsheet defaultWorkbook={wb} ref={ref} />);

    let count = 0;
    act(() => {
      count = ref.current?.api.replaceAll({ find: "foo", replace: "baz" }) ?? 0;
    });
    expect(count).toBe(2);
    expect(ref.current?.api.getRaw(0, 0)).toBe("baz");
    expect(ref.current?.api.getRaw(1, 0)).toBe("bazd");
    expect(ref.current?.api.status).toBe("Replaced 2 occurrences.");
  });

  test("findNext wraps", () => {
    const wb = createWorkbook(["S"]);
    Object.assign(wb.sheets[0] as object, {
      cells: { "0_0": "x", "2_0": "x" },
    });
    const ref = createRef<SheetRootHandle>();
    render(<Spreadsheet defaultWorkbook={wb} ref={ref} />);

    const first = ref.current?.api.findNext({ find: "x" });
    const second = ref.current?.api.findNext({
      find: "x",
      after: first ?? undefined,
    });
    const wrap = ref.current?.api.findNext({
      find: "x",
      after: { row: 2, col: 0 },
    });
    expect(first?.row).toBe(0);
    expect(second?.row).toBe(2);
    expect(wrap?.row).toBe(0);
  });
});

describe("checkboxes, hyperlinks, rotation", () => {
  test("insertCheckboxes and toggle", () => {
    const ref = createRef<SheetRootHandle>();
    const { container } = render(<Spreadsheet ref={ref} />);

    act(() => {
      ref.current?.api.select({ r1: 0, c1: 0, r2: 0, c2: 0 });
      ref.current?.api.insertCheckboxes();
    });
    expect(ref.current?.api.sheet.styles["0_0"]?.checkbox).toBe(true);
    expect(ref.current?.api.getRaw(0, 0)).toBe("FALSE");

    act(() => {
      ref.current?.api.toggleCheckbox(0, 0);
    });
    expect(ref.current?.api.getRaw(0, 0)).toBe("TRUE");
    expect(container.querySelector('input[type="checkbox"]')).not.toBeNull();
  });

  test("setHyperlink and setTextRotation", () => {
    const ref = createRef<SheetRootHandle>();
    render(<Spreadsheet ref={ref} />);

    act(() => {
      ref.current?.api.selectCell(0, 0);
      ref.current?.api.setCell(0, 0, "Docs");
      ref.current?.api.setHyperlink("https://example.com");
      ref.current?.api.setTextRotation(45);
    });
    expect(ref.current?.api.sheet.styles["0_0"]?.hyperlink).toBe(
      "https://example.com",
    );
    expect(ref.current?.api.sheet.styles["0_0"]?.rotation).toBe(45);
  });

  test("defineName reports status", () => {
    const ref = createRef<SheetRootHandle>();
    render(<Spreadsheet ref={ref} />);
    act(() => {
      ref.current?.api.defineName("Revenue", {
        r1: 0,
        c1: 0,
        r2: 5,
        c2: 0,
      });
    });
    expect(ref.current?.api.status).toBe("Named range REVENUE added");
  });
});
