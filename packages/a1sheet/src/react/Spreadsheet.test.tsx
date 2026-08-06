/**
 * Renders the real component in happy-dom. These are the tests that would have
 * caught "it renders nothing".
 */
import { describe, expect, test } from "bun:test";
import { render, screen } from "@testing-library/react";
import { createWorkbook } from "../model/workbook.js";
import { Sheet } from "./index.js";
import { Spreadsheet } from "./Spreadsheet.js";

function workbookWith(cells: Record<string, string>) {
  const wb = createWorkbook(["Sheet1"]);
  Object.assign((wb.sheets[0] as { cells: object }).cells, cells);
  return wb;
}

describe("rendering", () => {
  test("mounts and paints a grid", () => {
    const { container } = render(<Spreadsheet />);
    const cells = container.querySelectorAll(".a1s-cell");
    expect(cells.length).toBeGreaterThan(0);
  });

  test("renders the column and row headers in view", () => {
    const { container } = render(<Spreadsheet />);
    // Column headers carry a trailing "▾" filter button.
    const heads = [...container.querySelectorAll(".a1s-head")].map((el) =>
      (el.textContent ?? "").replace("▾", ""),
    );
    expect(heads).toContain("A");
    expect(heads).toContain("B");
    expect(heads).toContain("1");
    expect(heads).toContain("2");
    // Columns are virtualized, so a column well off the right edge is absent
    // until it is scrolled to — see "column virtualization" below.
    expect(heads).not.toContain("Z");
  });

  test("shows literal cell values", () => {
    render(<Spreadsheet defaultWorkbook={workbookWith({ "0_0": "hello" })} />);
    expect(screen.getByText("hello")).toBeDefined();
  });

  test("shows the EVALUATED value of a formula, not its source", () => {
    const { container } = render(
      <Spreadsheet
        defaultWorkbook={workbookWith({
          "0_0": "2",
          "1_0": "3",
          "2_0": "=SUM(A1:A2)",
        })}
      />,
    );
    // Scoped to cells: row headers also contain digits.
    const cellText = [...container.querySelectorAll(".a1s-cell")].map(
      (el) => el.textContent,
    );
    expect(cellText).toContain("5");
    expect(cellText).not.toContain("=SUM(A1:A2)");
  });

  test("renders the toolbar, formula bar, tabs and status bar", () => {
    const { container } = render(<Spreadsheet />);
    expect(screen.getByTitle("Undo (Ctrl+Z)")).toBeDefined();
    expect(screen.getByLabelText("Formula")).toBeDefined();
    expect(screen.getByLabelText("Name box")).toBeDefined();
    expect(container.querySelector(`.a1s-tab`)).not.toBeNull();
    expect(container.querySelector(`.a1s-status`)).not.toBeNull();
  });

  test("omitting a part is composition, not a flag", () => {
    // The shadcn contract: chrome you do not render simply is not there. There is
    // no showToolbar prop and there must never be one.
    const { container } = render(
      <Sheet.Root>
        <Sheet.Grid />
      </Sheet.Root>,
    );
    expect(container.querySelector(".a1s-tab")).toBeNull();
    expect(container.querySelector(".a1s-status")).toBeNull();
    expect(container.querySelector(".a1s-btn")).toBeNull();
    expect(container.querySelectorAll(".a1s-cell").length).toBeGreaterThan(0);
  });

  test("injects its CSS so there is nothing to import", () => {
    const { container } = render(<Spreadsheet />);
    const style = container.querySelector("style");
    expect(style?.textContent).toContain(".a1s-cell");
  });

  test("honors a custom class-name prefix", () => {
    const { container } = render(<Spreadsheet classNamePrefix="x-" />);
    expect(container.querySelectorAll(".x-cell").length).toBeGreaterThan(0);
    expect(container.querySelectorAll(".a1s-cell").length).toBe(0);
  });
});

describe("virtualization", () => {
  test("renders far fewer rows than the sheet has", () => {
    const { container } = render(<Spreadsheet height={300} />);
    const rendered = container.querySelectorAll(".a1s-cell").length;
    // 200 rows x 26 cols would be 5200 cells if it drew everything.
    expect(rendered).toBeLessThan(5200);
    expect(rendered).toBeGreaterThan(0);
  });
});

describe("merges", () => {
  test("a merged range renders one cell, not one per covered position", () => {
    const wb = createWorkbook(["S"]);
    const sheet = wb.sheets[0];
    if (sheet) {
      sheet.cells["0_0"] = "merged";
      sheet.merges.push({ r1: 0, c1: 0, r2: 1, c2: 1 });
    }
    render(<Spreadsheet defaultWorkbook={wb} />);
    expect(screen.getAllByText("merged")).toHaveLength(1);
  });
});
