/**
 * A Ctrl+click selection that is actually selected.
 *
 * It used to be shaded and nothing else: the ranges were drawn, and every
 * operation — clear, format, copy — acted on the primary range alone, so the
 * other blocks were decoration. These are the operations, plus the two mouse
 * gestures Excel has that the grid did not: Ctrl+click to deselect, and
 * dragging across the row and column headers.
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
    ) as HTMLElement;

  const colHeader = (col: number) =>
    container.querySelector(`.a1s-head[data-col="${col}"]`) as HTMLElement;
  const rowHeader = (row: number) =>
    container.querySelector(`.a1s-head[data-row="${row}"]`) as HTMLElement;

  /** Ctrl+click, which is what a Mac's Cmd+click arrives as here too. */
  const ctrlClick = (el: HTMLElement) => fireEvent.mouseDown(el, { ctrlKey: true });

  const selected = () =>
    [...container.querySelectorAll(".a1s-selected, .a1s-active")]
      .map(
        (el) =>
          `${(el as HTMLElement).dataset.row}_${(el as HTMLElement).dataset.col}`,
      )
      .sort();

  const shown = (row: number, col: number) => cellAt(row, col)?.textContent ?? "";

  const copy = (): string | null => {
    let text: string | null = null;
    fireEvent.copy(textarea, {
      clipboardData: {
        setData: (_kind: string, value: string) => {
          text = value;
        },
      },
    });
    return text;
  };

  const dragOver = (row: number, col: number) =>
    fireEvent.mouseMove(window, {
      clientX: ROW_HEADER_WIDTH + col * DEFAULT_COL_WIDTH + 1,
      clientY: HEADER_HEIGHT + row * ROW_HEIGHT + 1,
      buttons: 1,
    });

  return {
    container,
    textarea,
    cellAt,
    colHeader,
    rowHeader,
    ctrlClick,
    selected,
    shown,
    copy,
    dragOver,
  };
}

describe("operations cover every selected range", () => {
  test("Delete clears all of them, not just the last", () => {
    const { cellAt, ctrlClick, shown, textarea } = setup({
      "0_0": "a",
      "2_2": "b",
      "4_4": "c",
    });
    ctrlClick(cellAt(2, 2));
    ctrlClick(cellAt(4, 4));
    fireEvent.keyDown(textarea, { key: "Delete" });

    expect(shown(0, 0)).toBe("");
    expect(shown(2, 2)).toBe("");
    expect(shown(4, 4)).toBe("");
  });

  test("bold applies to all of them", () => {
    const { cellAt, ctrlClick, textarea } = setup({ "0_0": "a", "2_2": "b" });
    ctrlClick(cellAt(2, 2));
    fireEvent.keyDown(textarea, { key: "b", ctrlKey: true });

    expect(cellAt(0, 0).style.fontWeight).toBe("700");
    expect(cellAt(2, 2).style.fontWeight).toBe("700");
  });

  test("the status bar counts across them", () => {
    const { cellAt, ctrlClick, container } = setup({ "0_0": "2", "2_2": "3" });
    ctrlClick(cellAt(2, 2));

    expect(container.textContent).toContain("2 ranges");
    expect(container.textContent).toContain("Count: 2");
    expect(container.textContent).toContain("Sum: 5");
  });
});

describe("copying a multiple selection", () => {
  test("stacked blocks in the same columns join into one grid", () => {
    // Excel allows exactly this shape, and it is the useful one: two runs of
    // the same column copied as one column.
    const { cellAt, ctrlClick, copy } = setup({ "0_0": "a", "3_0": "b" });
    ctrlClick(cellAt(3, 0));

    expect(copy()).toBe("a\nb");
  });

  test("blocks sharing their rows join side by side", () => {
    const { cellAt, ctrlClick, copy } = setup({ "0_0": "a", "0_3": "b" });
    ctrlClick(cellAt(0, 3));

    expect(copy()).toBe("a\tb");
  });

  test("ranges that line up in neither direction are refused, with the reason", () => {
    // An L has no flattening, and inventing one pastes a shape nobody selected.
    const { cellAt, ctrlClick, copy, container } = setup({
      "0_0": "a",
      "2_5": "b",
    });
    ctrlClick(cellAt(2, 5));

    expect(copy()).toBeNull();
    expect(container.textContent).toContain(
      "cannot be used on multiple selections",
    );
  });
});

describe("Ctrl+click a second time takes it back out", () => {
  test("clicking a banked range removes it", () => {
    const { cellAt, ctrlClick, selected } = setup();
    ctrlClick(cellAt(2, 2));
    ctrlClick(cellAt(4, 4));
    expect(selected()).toEqual(["0_0", "2_2", "4_4"]);

    ctrlClick(cellAt(2, 2));
    expect(selected()).toEqual(["0_0", "4_4"]);
  });

  test("the last range standing cannot be removed", () => {
    // Something has to be active: there is always a cell the keyboard types in.
    const { cellAt, ctrlClick, selected } = setup();
    ctrlClick(cellAt(0, 0));
    expect(selected()).toEqual(["0_0"]);
  });
});

describe("the row and column headers", () => {
  test("dragging across column headers selects the columns between", () => {
    const { colHeader, cellAt, dragOver } = setup();
    fireEvent.mouseDown(colHeader(1));
    dragOver(0, 3);

    expect(cellAt(5, 2).className).toContain("a1s-selected");
    expect(cellAt(5, 3).className).toContain("a1s-selected");
    expect(cellAt(5, 4).className).not.toContain("a1s-selected");
  });

  test("dragging across row headers selects the rows between", () => {
    const { rowHeader, cellAt, dragOver } = setup();
    fireEvent.mouseDown(rowHeader(1));
    dragOver(3, 0);

    expect(cellAt(2, 5).className).toContain("a1s-selected");
    expect(cellAt(4, 5).className).not.toContain("a1s-selected");
  });

  test("Ctrl+clicking a second column adds it rather than replacing", () => {
    const { colHeader, cellAt, ctrlClick } = setup();
    fireEvent.mouseDown(colHeader(0));
    ctrlClick(colHeader(3));

    expect(cellAt(4, 0).className).toContain("a1s-selected");
    expect(cellAt(4, 3).className).toContain("a1s-selected");
    expect(cellAt(4, 1).className).not.toContain("a1s-selected");
  });

  test("Shift+clicking a column extends from the one already selected", () => {
    const { colHeader, cellAt } = setup();
    fireEvent.mouseDown(colHeader(1));
    fireEvent.mouseDown(colHeader(3), { shiftKey: true });

    expect(cellAt(2, 2).className).toContain("a1s-selected");
  });
});

describe("the selection is outlined, not only shaded", () => {
  test("one box per range", () => {
    const { container, cellAt, ctrlClick } = setup();
    ctrlClick(cellAt(2, 2));

    expect(container.querySelectorAll(".a1s-selbox")).toHaveLength(2);
  });

  test("every box is the same weight — the active range is marked from inside", () => {
    // A heavier border on the primary range stacked with the active cell's own
    // outline and the fill handle: three edges within two pixels of each other.
    const { container, cellAt, ctrlClick } = setup();
    ctrlClick(cellAt(2, 2));

    const classes = [...container.querySelectorAll(".a1s-selbox")].map(
      (el) => el.className,
    );
    expect(new Set(classes).size).toBe(1);
  });

  test("the outlines sit below the sticky headers", () => {
    // They are absolutely positioned over the grid, so at a z-index above the
    // headers a selection scrolled under the header row drew across it.
    const { container } = setup();
    const css = container.querySelector("style")?.textContent ?? "";

    const layer = (selector: string) =>
      Number(
        new RegExp(`\\.${selector} \\{[^}]*z-index: (\\d+)`).exec(css)?.[1] ?? "0",
      );

    expect(layer("a1s-selbox")).toBeLessThan(layer("a1s-dropdown"));
    expect(layer("a1s-marquee")).toBeLessThan(layer("a1s-dropdown"));
  });
});

describe("the copy outline goes away", () => {
  test("when a cell is edited", () => {
    const { container, copy, textarea } = setup({ "0_0": "a" });
    copy();
    expect(container.querySelectorAll(".a1s-marquee")).toHaveLength(1);

    fireEvent.keyDown(textarea, { key: "x" });
    const input = container.querySelector(".a1s-cell input") as HTMLInputElement;
    fireEvent.change(input, { target: { value: "hello" } });
    fireEvent.keyDown(input, { key: "Enter" });

    expect(container.querySelectorAll(".a1s-marquee")).toHaveLength(0);
  });

  test("when the copied cells are cleared", () => {
    const { container, copy, textarea } = setup({ "0_0": "a" });
    copy();
    fireEvent.keyDown(textarea, { key: "Delete" });

    expect(container.querySelectorAll(".a1s-marquee")).toHaveLength(0);
  });
});

describe("pasting", () => {
  test("drops the extra ranges rather than pasting into all of them", () => {
    // Excel refuses a multi-range paste target outright; collapsing to the
    // active range does the thing the user probably meant and says so on screen.
    const { cellAt, ctrlClick, textarea, container } = setup({ "0_0": "a" });
    ctrlClick(cellAt(2, 2));
    ctrlClick(cellAt(4, 4));
    fireEvent.paste(textarea, { clipboardData: { getData: () => "z" } });

    expect(container.textContent).not.toContain("ranges");
    expect(screen.getByText("z")).toBeDefined();
  });
});
