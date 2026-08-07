/**
 * Selecting, filling, and copying without a mouse.
 *
 * Everything a mouse can do to a selection has to have a keyboard route, and
 * until now three did not: crossing a long sheet, building a discontiguous
 * selection, and dragging the fill handle. These test the routes, not the
 * underlying rules — `model/navigate.test.ts` covers where Ctrl+Arrow lands.
 */
import { describe, expect, test } from "bun:test";
import { fireEvent, render, screen } from "@testing-library/react";
import { createWorkbook } from "../model/workbook.js";
import {
  DEFAULT_COL_WIDTH,
  HEADER_HEIGHT,
  ROW_HEADER_WIDTH,
  ROW_HEIGHT,
} from "./constants.js";
import { Spreadsheet } from "./Spreadsheet.js";

function setup(cells: Record<string, string> = {}) {
  const wb = createWorkbook(["Sheet1"]);
  Object.assign((wb.sheets[0] as { cells: object }).cells, cells);
  const { container } = render(<Spreadsheet defaultWorkbook={wb} />);
  const textarea = container.querySelector("textarea") as HTMLTextAreaElement;

  const cellAt = (row: number, col: number) =>
    container.querySelector(
      `.a1s-cell[data-row="${row}"][data-col="${col}"]`,
    ) as HTMLElement | null;

  const press = (key: string, modifiers: Record<string, boolean> = {}) =>
    fireEvent.keyDown(textarea, { key, ...modifiers });

  /** Addresses of every selected cell, in DOM order. */
  const selected = () =>
    [...container.querySelectorAll(".a1s-selected, .a1s-active")].map(
      (el) =>
        `${(el as HTMLElement).dataset.row}_${(el as HTMLElement).dataset.col}`,
    );

  const active = () => {
    const el = container.querySelector(".a1s-active") as HTMLElement | null;
    return el ? `${el.dataset.row}_${el.dataset.col}` : null;
  };

  /** One mousemove on window, in client coordinates, as a real drag sends. */
  const dragOver = (row: number, col: number) =>
    fireEvent.mouseMove(window, {
      clientX: ROW_HEADER_WIDTH + col * DEFAULT_COL_WIDTH + 1,
      clientY: HEADER_HEIGHT + row * ROW_HEIGHT + 1,
      buttons: 1,
    });

  /**
   * What a cell shows. Read from the DOM rather than from the workbook object
   * passed in: edits clone-on-write, so that object never changes.
   */
  const shown = (row: number, col: number) => cellAt(row, col)?.textContent ?? "";

  /**
   * The raw content of a cell — formula source included, which is what a fill
   * writes and what the display would hide behind a computed value.
   */
  const rawAt = (row: number, col: number) => {
    fireEvent.mouseDown(cellAt(row, col) as HTMLElement);
    fireEvent.mouseUp(window);
    return (screen.getByLabelText("Formula") as HTMLInputElement).value;
  };

  /** A copy event with the clipboard the browser would provide. */
  const copy = () => {
    const data: Record<string, string> = {};
    fireEvent.copy(textarea, {
      clipboardData: {
        setData: (kind: string, value: string) => {
          data[kind] = value;
        },
      },
    });
    return data;
  };

  return {
    container,
    textarea,
    cellAt,
    press,
    selected,
    active,
    dragOver,
    shown,
    rawAt,
    copy,
  };
}

describe("crossing the sheet", () => {
  test("Ctrl+Arrow jumps to the end of the block", () => {
    const { press, active } = setup({ "0_0": "a", "1_0": "b", "2_0": "c" });
    press("ArrowDown", { ctrlKey: true });
    expect(active()).toBe("2_0");
  });

  test("Ctrl+Shift+Arrow extends to it instead of moving", () => {
    const { press, selected } = setup({ "0_0": "a", "1_0": "b", "2_0": "c" });
    press("ArrowDown", { ctrlKey: true, shiftKey: true });
    expect(selected()).toEqual(["0_0", "1_0", "2_0"]);
  });

  test("Ctrl+A takes the whole sheet", () => {
    const { press, cellAt } = setup();
    press("a", { ctrlKey: true });
    expect(cellAt(0, 0)?.className).toContain("a1s-active");
    expect(cellAt(5, 5)?.className).toContain("a1s-selected");
  });

  test("Ctrl+Space takes the column and Shift+Space the row", () => {
    const { press, cellAt } = setup();
    press(" ", { ctrlKey: true });
    expect(cellAt(9, 0)?.className).toContain("a1s-selected");
    expect(cellAt(9, 1)?.className).not.toContain("a1s-selected");

    press(" ", { shiftKey: true });
    expect(cellAt(0, 4)?.className).toContain("a1s-selected");
    expect(cellAt(1, 4)?.className).not.toContain("a1s-selected");
  });

  test("End runs to the last filled cell in the row", () => {
    const { press, active } = setup({ "0_0": "a", "0_4": "b" });
    press("End");
    expect(active()).toBe("0_4");
  });

  test("Ctrl+End goes to the corner of the used range", () => {
    const { press, active } = setup({ "9_1": "x", "1_4": "y" });
    press("End", { ctrlKey: true });
    // The corner of the two axes, which holds nothing itself.
    expect(active()).toBe("9_4");
  });
});

describe("a discontiguous selection from the keyboard", () => {
  test("Shift+F8 keeps the range behind as the cursor moves on", () => {
    const { press, selected } = setup();

    press("ArrowDown", { shiftKey: true }); // A1:A2
    press("F8", { shiftKey: true }); // add to selection
    press("ArrowRight"); // move away, keeping A1:A2
    press("ArrowRight");
    press("ArrowDown", { shiftKey: true }); // C1:C2

    expect(selected().sort()).toEqual(["0_0", "0_2", "1_0", "1_2"]);
  });

  test("moving twice in add mode does not leave a trail of single cells", () => {
    const { press, selected } = setup();
    press("F8", { shiftKey: true });
    press("ArrowDown");
    press("ArrowDown");
    press("ArrowDown");

    // The one frozen range, plus the cursor's own cell.
    expect(selected().sort()).toEqual(["0_0", "3_0"]);
  });

  test("Escape ends add mode and drops the extra ranges", () => {
    const { press, selected } = setup();
    press("F8", { shiftKey: true });
    press("ArrowDown");
    press("Escape");

    expect(selected()).toEqual(["1_0"]);
  });

  test("Ctrl+click banks each range instead of replacing the last", () => {
    // `addRange` followed by `selectCell` cleared the extras, so a third
    // Ctrl+click used to leave two cells selected rather than three.
    const { cellAt, selected } = setup();
    fireEvent.mouseDown(cellAt(1, 1) as HTMLElement, { ctrlKey: true });
    fireEvent.mouseDown(cellAt(2, 2) as HTMLElement, { ctrlKey: true });

    expect(selected().sort()).toEqual(["0_0", "1_1", "2_2"]);
  });
});

describe("filling from the keyboard", () => {
  test("Ctrl+D copies the top of the selection down", () => {
    const { press, shown } = setup({ "0_0": "x" });
    press("ArrowDown", { shiftKey: true });
    press("ArrowDown", { shiftKey: true });
    press("d", { ctrlKey: true });

    expect(shown(1, 0)).toBe("x");
    expect(shown(2, 0)).toBe("x");
  });

  test("Ctrl+D counts on when the leading cells are a series", () => {
    // The fill handle numbers a column by dragging it. This is that, by
    // keyboard: the filled run at the top is the source, the rest is filled.
    const { press, shown } = setup({ "0_0": "1", "1_0": "2" });
    for (let i = 0; i < 4; i++) press("ArrowDown", { shiftKey: true });
    press("d", { ctrlKey: true });

    expect(shown(2, 0)).toBe("3");
    expect(shown(4, 0)).toBe("5");
  });

  test("Ctrl+D shifts a formula's references down, one row at a time", () => {
    const { press, rawAt } = setup({ "0_0": "=B1*2" });
    press("ArrowDown", { shiftKey: true });
    press("ArrowDown", { shiftKey: true });
    press("d", { ctrlKey: true });

    expect(rawAt(1, 0)).toBe("=B2*2");
    expect(rawAt(2, 0)).toBe("=B3*2");
  });

  test("Ctrl+R fills across, and counts on from a lone number", () => {
    // Deliberately the same rule as the fill handle, which numbers from a
    // single value rather than repeating it. Excel's Ctrl+R copies instead;
    // matching this library's own drag matters more than matching that.
    const { press, shown } = setup({ "0_0": "7" });
    press("ArrowRight", { shiftKey: true });
    press("ArrowRight", { shiftKey: true });
    press("r", { ctrlKey: true });

    expect(shown(0, 1)).toBe("8");
    expect(shown(0, 2)).toBe("9");
  });

  test("with nothing to fill, it says so rather than writing blanks", () => {
    const { press, container, shown } = setup();
    press("ArrowDown", { shiftKey: true });
    press("d", { ctrlKey: true });

    expect(shown(1, 0)).toBe("");
    expect(container.textContent).toContain("Fill down needs a filled cell");
  });
});

describe("dragging past the edge of the grid", () => {
  test("a drag extends the selection from wherever the pointer is", () => {
    // Per-cell mouse events cannot see a pointer held outside the grid, which
    // is exactly where it is when the sheet needs to scroll to follow it.
    const { cellAt, dragOver, selected } = setup();
    fireEvent.mouseDown(cellAt(0, 0) as HTMLElement);
    dragOver(2, 1);

    expect(selected().sort()).toEqual(["0_0", "0_1", "1_0", "1_1", "2_0", "2_1"]);
  });

  test("a pointer above the grid clamps to the first row, not to nothing", () => {
    const { cellAt, selected } = setup();
    fireEvent.mouseDown(cellAt(3, 0) as HTMLElement);
    fireEvent.mouseMove(window, { clientX: -400, clientY: -400, buttons: 1 });

    expect(selected().sort()).toEqual(["0_0", "1_0", "2_0", "3_0"]);
  });

  test("mouseup ends the drag, so later movement does not keep selecting", () => {
    const { cellAt, dragOver, selected } = setup();
    fireEvent.mouseDown(cellAt(0, 0) as HTMLElement);
    dragOver(1, 0);
    fireEvent.mouseUp(window);
    dragOver(5, 0);

    expect(selected().sort()).toEqual(["0_0", "1_0"]);
  });
});

describe("the copy outline", () => {
  test("copying draws a dashed box around what was copied", () => {
    const { container, copy } = setup({ "0_0": "a" });
    expect(container.querySelector(".a1s-marquee")).toBeNull();

    copy();
    expect(container.querySelector(".a1s-marquee")).not.toBeNull();
  });

  test("it marks the source, and stays there when the selection moves", () => {
    const { container, copy, press } = setup({ "0_0": "a" });
    copy();
    press("ArrowDown");
    press("ArrowDown");

    const box = container.querySelector(".a1s-marquee") as HTMLElement;
    expect(box.style.top).toBe(`${HEADER_HEIGHT}px`);
  });

  test("pasting clears it", () => {
    const { container, textarea, copy } = setup({ "0_0": "a" });
    copy();
    fireEvent.paste(textarea, {
      clipboardData: { getData: () => "b" },
    });

    expect(container.querySelector(".a1s-marquee")).toBeNull();
  });

  test("Escape clears it too", () => {
    const { container, copy, press } = setup({ "0_0": "a" });
    copy();
    press("Escape");

    expect(container.querySelector(".a1s-marquee")).toBeNull();
  });
});
