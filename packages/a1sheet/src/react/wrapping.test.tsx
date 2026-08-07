/**
 * A wrapped cell growing its row.
 *
 * `wrap` has always made the text run onto more lines; the row kept the height
 * it had, so every line but the first was clipped. Excel and Sheets grow the
 * row, and stop growing it once you drag it — an explicit height wins.
 *
 * The row header is what carries the height (it is the one item guaranteed to
 * exist in every row), so that is what these read.
 */
import { afterAll, beforeAll, describe, expect, test } from "bun:test";
import { fireEvent, render } from "@testing-library/react";
import type { StyleObject } from "../model/types.js";
import { createWorkbook } from "../model/workbook.js";
import { ROW_HEIGHT } from "./constants.js";
import { Spreadsheet } from "./Spreadsheet.js";

/**
 * A canvas that measures every character as 10px wide.
 *
 * jsdom has no text metrics, so without this `measureText` reports 0, every
 * string fits on one line, and no row ever grows — the tests would pass against
 * a completely broken measurer. The default column is 92px wide with 12px of
 * padding, so a line holds exactly eight of these characters.
 */
const CHAR_PX = 10;
let originalGetContext: typeof HTMLCanvasElement.prototype.getContext;

beforeAll(() => {
  originalGetContext = HTMLCanvasElement.prototype.getContext;
  HTMLCanvasElement.prototype.getContext = (() => ({
    font: "",
    measureText: (text: string) => ({ width: text.length * CHAR_PX }),
  })) as unknown as typeof HTMLCanvasElement.prototype.getContext;
});

afterAll(() => {
  HTMLCanvasElement.prototype.getContext = originalGetContext;
});

function setup(
  cells: Record<string, string>,
  styles: Record<string, StyleObject> = {},
  rowHeights: Record<number, number> = {},
) {
  const wb = createWorkbook(["Sheet1"]);
  Object.assign(wb.sheets[0] as object, { cells, styles, rowHeights });
  const { container } = render(<Spreadsheet defaultWorkbook={wb} />);

  const heightOf = (row: number) =>
    Number.parseFloat(
      (container.querySelector(`.a1s-head[data-row="${row}"]`) as HTMLElement).style
        .height,
    );

  const hidden = container.querySelector("textarea") as HTMLTextAreaElement;
  return { container, heightOf, hidden };
}

/**
 * A default column is 92px wide with 12px of padding, so a line holds eight of
 * these characters. Two four-letter words do not fit together; two three-letter
 * words plus their space do, at exactly seven.
 */
const TWO_LINES = "aaaa bbbb c";
const FOUR_LINES = "aaa bbb ccc ddd eee fff ggg hhh";

describe("growing", () => {
  test("a wrapped cell that needs two lines makes its row taller", () => {
    const { heightOf } = setup({ "0_0": TWO_LINES }, { "0_0": { wrap: true } });

    expect(heightOf(0)).toBeGreaterThan(ROW_HEIGHT);
  });

  test("more lines means more height, in even steps", () => {
    const two = setup({ "0_0": TWO_LINES }, { "0_0": { wrap: true } });
    const four = setup({ "0_0": FOUR_LINES }, { "0_0": { wrap: true } });

    const perLine = two.heightOf(0) - ROW_HEIGHT;
    expect(perLine).toBeGreaterThan(0);
    expect(four.heightOf(0) - ROW_HEIGHT).toBe(perLine * 3);
  });

  test("the same text without wrap leaves the row alone", () => {
    // The cell is what clips, not the row: unwrapped text is one long line.
    const { heightOf } = setup({ "0_0": FOUR_LINES });

    expect(heightOf(0)).toBe(ROW_HEIGHT);
  });

  test("a row grows to its tallest wrapped cell, not its last", () => {
    const { heightOf } = setup(
      { "0_0": FOUR_LINES, "0_1": TWO_LINES },
      { "0_0": { wrap: true }, "0_1": { wrap: true } },
    );
    const tallest = setup({ "0_0": FOUR_LINES }, { "0_0": { wrap: true } });

    expect(heightOf(0)).toBe(tallest.heightOf(0));
  });

  test("only the row with the wrapped cell moves", () => {
    const { heightOf } = setup(
      { "0_0": FOUR_LINES, "1_0": FOUR_LINES },
      { "0_0": { wrap: true } },
    );

    expect(heightOf(0)).toBeGreaterThan(ROW_HEIGHT);
    expect(heightOf(1)).toBe(ROW_HEIGHT);
  });
});

describe("an explicit height wins", () => {
  test("a row that was dragged keeps the height it was given", () => {
    const { heightOf } = setup(
      { "0_0": FOUR_LINES },
      { "0_0": { wrap: true } },
      { 0: 40 },
    );

    expect(heightOf(0)).toBe(40);
  });

  test("auto-fit drops that height and the row springs back to its content", () => {
    // Double-clicking the row divider. Dropping the override IS the measurement
    // now: the fallback is what the wrapped content needs.
    const { container, heightOf } = setup(
      { "0_0": FOUR_LINES },
      { "0_0": { wrap: true } },
      { 0: 40 },
    );
    const grown = setup({ "0_0": FOUR_LINES }, { "0_0": { wrap: true } });

    const grip = container.querySelector(
      '.a1s-head[data-row="0"] .a1s-rowresize',
    ) as HTMLElement;
    fireEvent.doubleClick(grip);

    expect(heightOf(0)).toBe(grown.heightOf(0));
  });
});

describe("keeping up with edits", () => {
  test("typing more text into a wrapped cell grows the row further", () => {
    const { heightOf, hidden, container } = setup(
      { "0_0": TWO_LINES },
      { "0_0": { wrap: true } },
    );
    const before = heightOf(0);

    fireEvent.keyDown(hidden, { key: "x" });
    const input = container.querySelector(".a1s-cell input") as HTMLInputElement;
    fireEvent.change(input, { target: { value: FOUR_LINES } });
    fireEvent.keyDown(input, { key: "Enter" });

    expect(heightOf(0)).toBeGreaterThan(before);
  });
});
