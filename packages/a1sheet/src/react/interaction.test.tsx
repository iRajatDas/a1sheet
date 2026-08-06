/**
 * Interaction tests: typing, navigation, formulas recalculating, undo/redo,
 * formatting, sorting, filtering. These exercise the component the way a user
 * does rather than asserting on internal state.
 */
import { describe, expect, test } from "bun:test";
import { fireEvent, render, screen } from "@testing-library/react";
import { createWorkbook } from "../model/workbook.js";
import { Spreadsheet } from "./Spreadsheet.js";

function setup(cells: Record<string, string> = {}) {
  const wb = createWorkbook(["Sheet1"]);
  Object.assign((wb.sheets[0] as { cells: object }).cells, cells);
  const utils = render(<Spreadsheet defaultWorkbook={wb} />);
  const { container } = utils;

  const cellAt = (row: number, col: number) => {
    // Grid items are ordered; find by the text of the row header sibling is
    // fragile, so index cells by their grid position via the DOM order instead.
    const cellsEls = [...container.querySelectorAll(".a1s-cell")];
    return cellsEls[row * 26 + col] as HTMLElement | undefined;
  };

  const cellText = () =>
    [...container.querySelectorAll(".a1s-cell")].map((el) => el.textContent);

  const hidden = container.querySelector("textarea") as HTMLTextAreaElement;
  const formula = screen.getByLabelText("Formula") as HTMLInputElement;

  return { ...utils, container, cellAt, cellText, hidden, formula };
}

describe("typing into cells", () => {
  test("a printable key starts an edit seeded with that character", () => {
    const { hidden, container } = setup();
    fireEvent.keyDown(hidden, { key: "h" });
    const input = container.querySelector(".a1s-cell input") as HTMLInputElement;
    expect(input).not.toBeNull();
    expect(input.value).toBe("h");
  });

  test("Enter commits and the value appears in the grid", () => {
    const { hidden, container, cellText } = setup();
    fireEvent.keyDown(hidden, { key: "h" });
    const input = container.querySelector(".a1s-cell input") as HTMLInputElement;
    fireEvent.change(input, { target: { value: "hello" } });
    fireEvent.keyDown(input, { key: "Enter" });
    expect(cellText()).toContain("hello");
  });

  test("Escape abandons the edit", () => {
    const { hidden, container, cellText } = setup();
    fireEvent.keyDown(hidden, { key: "x" });
    const input = container.querySelector(".a1s-cell input") as HTMLInputElement;
    fireEvent.change(input, { target: { value: "discard me" } });
    fireEvent.keyDown(input, { key: "Escape" });
    expect(cellText()).not.toContain("discard me");
  });

  test("F2 opens an edit with the existing content", () => {
    const { hidden, container } = setup({ "0_0": "existing" });
    fireEvent.keyDown(hidden, { key: "F2" });
    const input = container.querySelector(".a1s-cell input") as HTMLInputElement;
    expect(input.value).toBe("existing");
  });

  test("Delete clears the selection", () => {
    const { hidden, cellText } = setup({ "0_0": "gone" });
    expect(cellText()).toContain("gone");
    fireEvent.keyDown(hidden, { key: "Delete" });
    expect(cellText()).not.toContain("gone");
  });
});

describe("formulas recalculate after an edit", () => {
  test("editing a dependency updates the dependent cell", () => {
    const { hidden, container, cellText } = setup({
      "0_0": "2",
      "1_0": "=A1*10",
    });
    expect(cellText()).toContain("20");

    // Select A1 is already the default; type a new value.
    fireEvent.keyDown(hidden, { key: "5" });
    const input = container.querySelector(".a1s-cell input") as HTMLInputElement;
    fireEvent.change(input, { target: { value: "5" } });
    fireEvent.keyDown(input, { key: "Enter" });

    expect(cellText()).toContain("50");
    expect(cellText()).not.toContain("20");
  });

  test("a formula typed by the user evaluates immediately", () => {
    const { hidden, container, cellText } = setup({ "0_1": "7" });
    fireEvent.keyDown(hidden, { key: "=" });
    const input = container.querySelector(".a1s-cell input") as HTMLInputElement;
    fireEvent.change(input, { target: { value: "=B1*3" } });
    fireEvent.keyDown(input, { key: "Enter" });
    expect(cellText()).toContain("21");
  });
});

describe("the formula bar", () => {
  test("shows the raw formula, not the evaluated value", () => {
    const { formula } = setup({ "0_0": "=1+1" });
    expect(formula.value).toBe("=1+1");
  });

  test("editing through the formula bar writes the cell", () => {
    const { formula, cellText } = setup();
    fireEvent.change(formula, { target: { value: "=2+3" } });
    fireEvent.keyDown(formula, { key: "Enter" });
    expect(cellText()).toContain("5");
  });

  test("the name box jumps the selection to a typed reference", () => {
    const { container } = setup();
    const nameBox = screen.getByLabelText("Name box") as HTMLInputElement;
    fireEvent.change(nameBox, { target: { value: "C3" } });
    fireEvent.keyDown(nameBox, { key: "Enter" });
    // The active cell carries the active class.
    const active = container.querySelector(".a1s-active");
    expect(active).not.toBeNull();
    expect(nameBox.placeholder).toBe("C3");
  });
});

describe("navigation", () => {
  test("arrow keys move the active cell", () => {
    const { hidden } = setup();
    const nameBox = screen.getByLabelText("Name box") as HTMLInputElement;
    expect(nameBox.placeholder).toBe("A1");
    fireEvent.keyDown(hidden, { key: "ArrowDown" });
    expect(nameBox.placeholder).toBe("A2");
    fireEvent.keyDown(hidden, { key: "ArrowRight" });
    expect(nameBox.placeholder).toBe("B2");
  });

  test("navigation clamps at the sheet edges", () => {
    const { hidden } = setup();
    const nameBox = screen.getByLabelText("Name box") as HTMLInputElement;
    fireEvent.keyDown(hidden, { key: "ArrowUp" });
    fireEvent.keyDown(hidden, { key: "ArrowLeft" });
    expect(nameBox.placeholder).toBe("A1");
  });

  test("Tab moves right, Shift+Tab moves back", () => {
    const { hidden } = setup();
    const nameBox = screen.getByLabelText("Name box") as HTMLInputElement;
    fireEvent.keyDown(hidden, { key: "Tab" });
    expect(nameBox.placeholder).toBe("B1");
    fireEvent.keyDown(hidden, { key: "Tab", shiftKey: true });
    expect(nameBox.placeholder).toBe("A1");
  });
});

describe("undo and redo", () => {
  test("undo restores the previous value, redo reapplies it", () => {
    const { hidden, container, cellText } = setup();
    fireEvent.keyDown(hidden, { key: "a" });
    const input = container.querySelector(".a1s-cell input") as HTMLInputElement;
    fireEvent.change(input, { target: { value: "first" } });
    fireEvent.keyDown(input, { key: "Enter" });
    expect(cellText()).toContain("first");

    fireEvent.click(screen.getByTitle("Undo (Ctrl+Z)"));
    expect(cellText()).not.toContain("first");

    fireEvent.click(screen.getByTitle("Redo (Ctrl+Y)"));
    expect(cellText()).toContain("first");
  });

  test("undo is disabled until something changes", () => {
    setup();
    expect((screen.getByTitle("Undo (Ctrl+Z)") as HTMLButtonElement).disabled).toBe(
      true,
    );
  });

  test("Ctrl+Z undoes", () => {
    const { hidden, container, cellText } = setup();
    fireEvent.keyDown(hidden, { key: "z2" });
    fireEvent.keyDown(hidden, { key: "q" });
    const input = container.querySelector(".a1s-cell input") as HTMLInputElement;
    fireEvent.change(input, { target: { value: "typed" } });
    fireEvent.keyDown(input, { key: "Enter" });
    expect(cellText()).toContain("typed");
    fireEvent.keyDown(hidden, { key: "z", ctrlKey: true });
    expect(cellText()).not.toContain("typed");
  });
});

describe("formatting", () => {
  test("the bold button applies to the selection and reflects state", () => {
    const { container } = setup({ "0_0": "bold me" });
    const boldBtn = screen.getByTitle("Bold (Ctrl+B)");
    expect(boldBtn.className).not.toContain("a1s-on");
    fireEvent.click(boldBtn);
    expect(screen.getByTitle("Bold (Ctrl+B)").className).toContain("a1s-on");
    const cell = [...container.querySelectorAll(".a1s-cell")].find(
      (el) => el.textContent === "bold me",
    ) as HTMLElement;
    expect(cell.style.fontWeight).toBe("700");
  });

  test("a number format changes the displayed value", () => {
    const { container } = setup({ "0_0": "0.5" });
    fireEvent.change(screen.getByLabelText("Number format"), {
      target: { value: "percent" },
    });
    const text = [...container.querySelectorAll(".a1s-cell")].map(
      (el) => el.textContent,
    );
    expect(text).toContain("50.00%");
  });
});

describe("structure", () => {
  test("Insert row shifts existing content down", () => {
    const { cellText, container } = setup({ "0_0": "shifted" });
    const before = [...container.querySelectorAll(".a1s-cell")].findIndex(
      (el) => el.textContent === "shifted",
    );
    fireEvent.click(screen.getByTitle("Insert row"));
    const after = [...container.querySelectorAll(".a1s-cell")].findIndex(
      (el) => el.textContent === "shifted",
    );
    expect(after).toBeGreaterThan(before);
    expect(cellText()).toContain("shifted");
  });

  test("adding a sheet switches to it and clears the view", () => {
    const { cellText } = setup({ "0_0": "on sheet one" });
    expect(cellText()).toContain("on sheet one");
    fireEvent.click(screen.getByTitle("Add sheet"));
    expect(cellText()).not.toContain("on sheet one");
    expect(screen.getByText("Sheet2")).toBeDefined();
  });
});

describe("status bar", () => {
  test("sums the numeric cells in the selection", () => {
    const { container } = setup({ "0_0": "10", "1_0": "20", "2_0": "30" });
    const nameBox = screen.getByLabelText("Name box") as HTMLInputElement;
    fireEvent.change(nameBox, { target: { value: "A1:A3" } });
    fireEvent.keyDown(nameBox, { key: "Enter" });
    const status = container.querySelector(".a1s-status")?.textContent ?? "";
    expect(status).toContain("Sum: 60");
    expect(status).toContain("Avg: 20");
    expect(status).toContain("Count: 3");
  });
});
