/**
 * Regression tests for grid focus and active-cell semantics.
 *
 * The existing interaction tests fire keydown directly at the hidden textarea,
 * which assumes the thing that was actually broken: that clicking a cell leaves
 * focus there. Every test here goes through a real click first and then reads
 * `document.activeElement`, so a grid that is dead to the keyboard fails them.
 */
import { describe, expect, test } from "bun:test";
import { createEvent, fireEvent, render } from "@testing-library/react";
import { createWorkbook } from "../model/workbook.js";
import { Spreadsheet } from "./Spreadsheet.js";

const COLS = 26;

function setup(
  cells: Record<string, string> = {},
  styles: Record<string, object> = {},
) {
  const wb = createWorkbook(["Sheet1"]);
  Object.assign((wb.sheets[0] as { cells: object }).cells, cells);
  Object.assign((wb.sheets[0] as { styles: object }).styles, styles);
  const { container } = render(<Spreadsheet defaultWorkbook={wb} />);

  const cellAt = (row: number, col: number) =>
    [...container.querySelectorAll(".a1s-cell")][row * COLS + col] as HTMLElement;

  const textarea = container.querySelector("textarea") as HTMLTextAreaElement;

  /**
   * Labels of the header cells carrying the "spanned by the selection" class.
   * Column headers also contain the filter caret, which is not part of the label.
   */
  const litHeaders = () =>
    [...container.querySelectorAll(".a1s-head.a1s-headon")].map((el) =>
      (el.textContent ?? "").replace("▾", ""),
    );

  return { container, cellAt, textarea, litHeaders };
}

describe("clicking a cell keeps the keyboard alive", () => {
  test("focus returns to the grid after a click", () => {
    const { cellAt, textarea } = setup();
    textarea.blur();
    expect(document.activeElement).not.toBe(textarea);

    fireEvent.mouseDown(cellAt(1, 1));

    // Without this, focus sits on <body> and every shortcut below is dead.
    expect(document.activeElement).toBe(textarea);
  });

  test("select a cell, then type — the edit starts", () => {
    const { container, cellAt, textarea } = setup();
    textarea.blur();
    fireEvent.mouseDown(cellAt(2, 1));

    // Deliberately keyed at whatever actually has focus, not at the textarea.
    fireEvent.keyDown(document.activeElement as Element, { key: "q" });

    const input = container.querySelector(".a1s-cell input") as HTMLInputElement;
    expect(input).not.toBeNull();
    expect(input.value).toBe("q");
  });

  test("copy reaches the clipboard handler after a click", () => {
    const { cellAt, textarea } = setup({ "0_0": "copy me" });
    textarea.blur();
    fireEvent.mouseDown(cellAt(0, 0));

    let written = "";
    fireEvent.copy(document.activeElement as Element, {
      clipboardData: {
        setData: (_type: string, value: string) => {
          written = value;
        },
      },
    });
    expect(written).toBe("copy me");
  });

  test("clicking a row header also returns focus", () => {
    const { container, textarea } = setup();
    textarea.blur();
    const rowHeader = [...container.querySelectorAll(".a1s-head")].find(
      (el) => el.textContent === "3",
    ) as HTMLElement;

    fireEvent.mouseDown(rowHeader);
    expect(document.activeElement).toBe(textarea);
  });
});

describe("dragging across cells does not select their text", () => {
  test("mousedown on a cell is default-prevented", () => {
    const { cellAt } = setup({ "0_0": "some text" });
    const event = createEvent.mouseDown(cellAt(0, 0));
    fireEvent(cellAt(0, 0), event);

    // The browser's default for mousedown is to start a text selection and move
    // focus off the div. Suppressing it is what fixes both.
    expect(event.defaultPrevented).toBe(true);
  });

  test("a click inside the open editor is left alone", () => {
    const { container, cellAt } = setup({ "0_0": "edit me" });
    fireEvent.doubleClick(cellAt(0, 0));
    const input = container.querySelector(".a1s-cell input") as HTMLInputElement;

    const event = createEvent.mouseDown(input);
    fireEvent(input, event);

    // Suppressing this would stop the caret being placed where you clicked.
    expect(event.defaultPrevented).toBe(false);
  });
});

describe("the active cell is the anchor, not the drag end", () => {
  test("shift-extending leaves the anchor active", () => {
    const { cellAt } = setup();
    fireEvent.mouseDown(cellAt(2, 2));
    fireEvent.mouseDown(cellAt(5, 5), { shiftKey: true });

    expect(cellAt(2, 2).className).toContain("a1s-active");
    expect(cellAt(5, 5).className).not.toContain("a1s-active");
    // Both ends are still inside the selection.
    expect(cellAt(5, 5).className).toContain("a1s-selected");
  });

  test("extending upward and leftward also keeps the anchor active", () => {
    const { cellAt } = setup();
    fireEvent.mouseDown(cellAt(5, 5));
    fireEvent.mouseDown(cellAt(2, 2), { shiftKey: true });

    expect(cellAt(5, 5).className).toContain("a1s-active");
    expect(cellAt(2, 2).className).not.toContain("a1s-active");
  });

  test("typing after a range drag lands in the anchor", () => {
    const { container, cellAt, textarea } = setup();
    textarea.blur();
    fireEvent.mouseDown(cellAt(3, 1));
    fireEvent.mouseDown(cellAt(6, 4), { shiftKey: true });
    fireEvent.keyDown(document.activeElement as Element, { key: "z" });

    const input = container.querySelector(".a1s-cell input") as HTMLInputElement;
    expect(input.value).toBe("z");
    // The editor opened inside the anchor cell, not the cell the drag ended on.
    expect(cellAt(3, 1).contains(input)).toBe(true);
  });

  test("an arrow key after a range drag steps from the anchor", () => {
    const { cellAt, textarea } = setup();
    fireEvent.mouseDown(cellAt(4, 4));
    fireEvent.mouseDown(cellAt(8, 8), { shiftKey: true });
    fireEvent.keyDown(textarea, { key: "ArrowDown" });

    expect(cellAt(5, 4).className).toContain("a1s-active");
    expect(cellAt(8, 8).className).not.toContain("a1s-selected");
  });

  test("an active cell with its own fill keeps that fill", () => {
    // The tint is excluded from the anchor via `:not(.active)` in CSS rather than
    // repainted over it, precisely so this holds. Repainting would show white.
    const { cellAt } = setup({ "1_1": "x" }, { "1_1": { bg: "#ffcc00" } });
    fireEvent.mouseDown(cellAt(1, 1));
    fireEvent.mouseDown(cellAt(3, 3), { shiftKey: true });

    expect(cellAt(1, 1).className).toContain("a1s-active");
    expect(cellAt(1, 1).style.background).toBe("#ffcc00");
  });

  test("every cell resolves to an opaque background, tinted or not", () => {
    // The selection tint is a translucent overlay. If the cell underneath were
    // left transparent, a sticky frozen row would show the scroll through it.
    const { cellAt } = setup();
    fireEvent.mouseDown(cellAt(1, 1));
    fireEvent.mouseDown(cellAt(3, 3), { shiftKey: true });

    expect(cellAt(1, 1).style.background).toBe("#ffffff");
    expect(cellAt(2, 2).style.background).toBe("#ffffff");
  });
});

describe("headers highlight the selected range", () => {
  test("both axes light up for a rectangular selection", () => {
    const { cellAt, litHeaders } = setup();
    fireEvent.mouseDown(cellAt(1, 2));
    fireEvent.mouseDown(cellAt(3, 4), { shiftKey: true });

    const lit = litHeaders();
    expect(lit).toEqual(expect.arrayContaining(["C", "D", "E"]));
    expect(lit).toEqual(expect.arrayContaining(["2", "3", "4"]));
    expect(lit).not.toContain("B");
    expect(lit).not.toContain("1");
  });

  test("a single cell lights exactly one header per axis", () => {
    const { cellAt, litHeaders } = setup();
    fireEvent.mouseDown(cellAt(0, 0));

    expect(litHeaders().sort()).toEqual(["1", "A"]);
  });
});
