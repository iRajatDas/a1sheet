/**
 * Picking references by clicking the grid mid-formula, and the outlines that
 * show what the formula points at.
 *
 * The string mechanics are covered in formula/refEditing.test.ts. These tests
 * are about the parts only the component can get wrong: that a click writes a
 * reference instead of moving the selection, that the open editor survives the
 * click, and that a drag rewrites one reference rather than appending many.
 */
import { describe, expect, test } from "bun:test";
import { fireEvent, render, screen } from "@testing-library/react";
import { createWorkbook } from "../model/workbook.js";
import { Spreadsheet } from "./Spreadsheet.js";

const COLS = 26;

function setup(cells: Record<string, string> = {}) {
  const wb = createWorkbook(["Sheet1"]);
  Object.assign((wb.sheets[0] as { cells: object }).cells, cells);
  const { container } = render(<Spreadsheet defaultWorkbook={wb} />);

  const cellAt = (row: number, col: number) =>
    [...container.querySelectorAll(".a1s-cell")][row * COLS + col] as HTMLElement;

  const editor = () =>
    container.querySelector(".a1s-cell input") as HTMLInputElement | null;

  const textarea = container.querySelector("textarea") as HTMLTextAreaElement;
  const formulaBar = screen.getByLabelText("Formula") as HTMLInputElement;

  /** Opens an edit on A1 seeded with `source`, caret at the end. */
  const startFormula = (source: string) => {
    fireEvent.keyDown(textarea, { key: "=" });
    const input = editor() as HTMLInputElement;
    fireEvent.change(input, {
      target: { value: source, selectionStart: source.length },
    });
    return editor() as HTMLInputElement;
  };

  const outlines = () =>
    [...container.querySelectorAll('[aria-hidden="true"]')].filter((el) =>
      (el as HTMLElement).style.border.startsWith("2px solid"),
    );

  return {
    container,
    cellAt,
    editor,
    textarea,
    formulaBar,
    startFormula,
    outlines,
  };
}

describe("clicking a cell while typing a formula", () => {
  test("writes the reference instead of moving the selection", () => {
    const { cellAt, startFormula, editor } = setup();
    startFormula("=");

    fireEvent.mouseDown(cellAt(2, 1));

    expect(editor()?.value).toBe("=B3");
    // The edit is still open — the click did not commit or move away.
    expect(editor()).not.toBeNull();
  });

  test("the edited cell keeps its editor open", () => {
    const { cellAt, startFormula } = setup();
    startFormula("=");
    fireEvent.mouseDown(cellAt(4, 3));

    // A1 is the cell being edited; its input must still be the one on screen.
    expect(cellAt(0, 0).querySelector("input")).not.toBeNull();
  });

  test("a second click after an operator appends a second reference", () => {
    const { cellAt, startFormula, editor } = setup();
    const input = startFormula("=");
    fireEvent.mouseDown(cellAt(0, 1));
    expect(editor()?.value).toBe("=B1");

    fireEvent.change(input, { target: { value: "=B1+", selectionStart: 4 } });
    fireEvent.mouseDown(cellAt(0, 2));

    expect(editor()?.value).toBe("=B1+C1");
  });

  test("clicking again with the caret on a reference replaces it", () => {
    const { cellAt, startFormula, editor } = setup();
    startFormula("=");
    fireEvent.mouseDown(cellAt(0, 1));
    expect(editor()?.value).toBe("=B1");

    // No operator typed in between, so this is a correction, not an addition.
    fireEvent.mouseDown(cellAt(0, 3));
    expect(editor()?.value).toBe("=D1");
  });

  test("dragging grows one reference into a range", () => {
    const { cellAt, startFormula, editor } = setup();
    startFormula("=SUM(");

    fireEvent.mouseDown(cellAt(1, 1));
    expect(editor()?.value).toBe("=SUM(B2");

    fireEvent.mouseEnter(cellAt(3, 2), { buttons: 1 });
    expect(editor()?.value).toBe("=SUM(B2:C4");

    // Still one reference — a drag rewrites, it does not append.
    fireEvent.mouseEnter(cellAt(4, 4), { buttons: 1 });
    expect(editor()?.value).toBe("=SUM(B2:E5");
  });

  test("a click past a finished operand selects instead of referencing", () => {
    const { cellAt, startFormula, editor } = setup();
    const input = startFormula("=B1");
    fireEvent.change(input, { target: { value: "=B1+2", selectionStart: 5 } });

    fireEvent.mouseDown(cellAt(6, 6));

    // The formula committed and the click behaved like an ordinary selection.
    expect(editor()).toBeNull();
    expect(cellAt(6, 6).className).toContain("a1s-active");
  });

  test("a plain value edit is unaffected — clicks still select", () => {
    const { cellAt, textarea, editor } = setup();
    fireEvent.keyDown(textarea, { key: "h" });
    const input = editor() as HTMLInputElement;
    fireEvent.change(input, { target: { value: "hello", selectionStart: 5 } });

    fireEvent.mouseDown(cellAt(3, 3));

    expect(editor()).toBeNull();
    expect(cellAt(3, 3).className).toContain("a1s-active");
  });
});

describe("reference outlines", () => {
  test("one outline per reference while editing", () => {
    const { startFormula, outlines } = setup();
    expect(outlines()).toHaveLength(0);

    startFormula("=A1+B2");
    expect(outlines()).toHaveLength(2);
  });

  test("a range is one outline covering the whole rectangle", () => {
    const { startFormula, outlines } = setup();
    startFormula("=SUM(B2:C4)");

    const boxes = outlines();
    expect(boxes).toHaveLength(1);
    // Three rows tall at the fixed row height.
    expect((boxes[0] as HTMLElement).style.height).toBe("78px");
  });

  test("repeated references share a color, distinct ones do not", () => {
    const { startFormula, outlines } = setup();
    startFormula("=A1+B2+A1");

    const colors = outlines().map((el) => (el as HTMLElement).style.borderColor);
    expect(colors[0]).toBe(colors[2] as string);
    expect(colors[0]).not.toBe(colors[1] as string);
  });

  test("outlines disappear once the edit is committed", () => {
    const { startFormula, outlines, editor } = setup();
    startFormula("=A1+B2");
    expect(outlines()).toHaveLength(2);

    fireEvent.keyDown(editor() as HTMLInputElement, { key: "Enter" });
    expect(outlines()).toHaveLength(0);
  });

  test("a plain value draws no outlines", () => {
    const { textarea, container } = setup();
    fireEvent.keyDown(textarea, { key: "h" });
    const input = container.querySelector(".a1s-cell input") as HTMLInputElement;
    fireEvent.change(input, { target: { value: "A1 is not a formula" } });

    expect(
      [...container.querySelectorAll('[aria-hidden="true"]')].filter((el) =>
        (el as HTMLElement).style.border.startsWith("2px solid"),
      ),
    ).toHaveLength(0);
  });
});

describe("the formula bar and the cell editor stay in step", () => {
  test("a reference picked in the grid shows in the formula bar too", () => {
    const { cellAt, startFormula, formulaBar } = setup();
    startFormula("=");
    fireEvent.mouseDown(cellAt(2, 2));

    expect(formulaBar.value).toBe("=C3");
  });
});
