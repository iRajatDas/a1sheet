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
import type { Range, StyleObject } from "../model/types.js";
import { createWorkbook } from "../model/workbook.js";
import { CELL_FONT_SIZE, ROW_HEIGHT } from "./constants.js";
import { Spreadsheet } from "./Spreadsheet.js";
import type { Theme } from "./theme.js";

/**
 * A canvas that measures a character as 10px wide at the default font size, and
 * proportionally wider or narrower at any other.
 *
 * jsdom has no text metrics, so without this `measureText` reports 0, every
 * string fits on one line, and no row ever grows — the tests would pass against
 * a completely broken measurer. The default column is 92px wide with 12px of
 * padding, so a line holds exactly eight characters at the default size.
 *
 * Scaling with the size is what makes the theme-font tests mean anything: a
 * measurer that ignored `ctx.font` would report the same width whatever face the
 * theme asked for, which is precisely the bug.
 */
const CHAR_PX = 10;
const FONT_SIZE = /(\d+(?:\.\d+)?)px/;
let originalGetContext: typeof HTMLCanvasElement.prototype.getContext;

beforeAll(() => {
  originalGetContext = HTMLCanvasElement.prototype.getContext;
  HTMLCanvasElement.prototype.getContext = (() => {
    const ctx = {
      font: "",
      measureText: (text: string) => {
        const match = FONT_SIZE.exec(ctx.font);
        const size = match ? Number(match[1]) : CELL_FONT_SIZE;
        return { width: (text.length * CHAR_PX * size) / CELL_FONT_SIZE };
      },
    };
    return ctx;
  }) as unknown as typeof HTMLCanvasElement.prototype.getContext;
});

afterAll(() => {
  HTMLCanvasElement.prototype.getContext = originalGetContext;
});

interface Setup {
  cells: Record<string, string>;
  styles?: Record<string, StyleObject>;
  rowHeights?: Record<number, number>;
  merges?: Range[];
  theme?: Partial<Theme>;
}

function setup({ cells, styles = {}, rowHeights = {}, merges = [], theme }: Setup) {
  const wb = createWorkbook(["Sheet1"]);
  Object.assign(wb.sheets[0] as object, { cells, styles, rowHeights, merges });
  const { container } = render(
    <Spreadsheet defaultWorkbook={wb} {...(theme ? { theme } : {})} />,
  );

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

const WRAP: Record<string, StyleObject> = { "0_0": { wrap: true } };

describe("growing", () => {
  test("a wrapped cell that needs two lines makes its row taller", () => {
    const { heightOf } = setup({ cells: { "0_0": TWO_LINES }, styles: WRAP });

    expect(heightOf(0)).toBeGreaterThan(ROW_HEIGHT);
  });

  test("more lines means more height, in even steps", () => {
    const two = setup({ cells: { "0_0": TWO_LINES }, styles: WRAP });
    const four = setup({ cells: { "0_0": FOUR_LINES }, styles: WRAP });

    const perLine = two.heightOf(0) - ROW_HEIGHT;
    expect(perLine).toBeGreaterThan(0);
    expect(four.heightOf(0) - ROW_HEIGHT).toBe(perLine * 3);
  });

  test("the same text without wrap leaves the row alone", () => {
    // The cell is what clips, not the row: unwrapped text is one long line.
    const { heightOf } = setup({ cells: { "0_0": FOUR_LINES } });

    expect(heightOf(0)).toBe(ROW_HEIGHT);
  });

  test("a row grows to its tallest wrapped cell, not its last", () => {
    const { heightOf } = setup({
      cells: { "0_0": FOUR_LINES, "0_1": TWO_LINES },
      styles: { "0_0": { wrap: true }, "0_1": { wrap: true } },
    });
    const tallest = setup({ cells: { "0_0": FOUR_LINES }, styles: WRAP });

    expect(heightOf(0)).toBe(tallest.heightOf(0));
  });

  test("only the row with the wrapped cell moves", () => {
    const { heightOf } = setup({
      cells: { "0_0": FOUR_LINES, "1_0": FOUR_LINES },
      styles: WRAP,
    });

    expect(heightOf(0)).toBeGreaterThan(ROW_HEIGHT);
    expect(heightOf(1)).toBe(ROW_HEIGHT);
  });
});

describe("an explicit height wins", () => {
  test("a row that was dragged keeps the height it was given", () => {
    const { heightOf } = setup({
      cells: { "0_0": FOUR_LINES },
      styles: WRAP,
      rowHeights: { 0: 40 },
    });

    expect(heightOf(0)).toBe(40);
  });

  test("auto-fit drops that height and the row springs back to its content", () => {
    // Double-clicking the row divider. Dropping the override IS the measurement
    // now: the fallback is what the wrapped content needs.
    const { container, heightOf } = setup({
      cells: { "0_0": FOUR_LINES },
      styles: WRAP,
      rowHeights: { 0: 40 },
    });
    const grown = setup({ cells: { "0_0": FOUR_LINES }, styles: WRAP });

    const grip = container.querySelector(
      '.a1s-head[data-row="0"] .a1s-rowresize',
    ) as HTMLElement;
    fireEvent.doubleClick(grip);

    expect(heightOf(0)).toBe(grown.heightOf(0));
  });
});

describe("keeping up with edits", () => {
  test("typing more text into a wrapped cell grows the row further", () => {
    const { heightOf, hidden, container } = setup({
      cells: { "0_0": TWO_LINES },
      styles: WRAP,
    });
    const before = heightOf(0);

    fireEvent.keyDown(hidden, { key: "x" });
    const input = container.querySelector(".a1s-cell input") as HTMLInputElement;
    fireEvent.change(input, { target: { value: FOUR_LINES } });
    fireEvent.keyDown(input, { key: "Enter" });

    expect(heightOf(0)).toBeGreaterThan(before);
  });
});

/**
 * The theme's face, not the default one.
 *
 * Measuring in a face the browser is not drawing in is off by however far the
 * two differ — a row a line short, or a line of empty space under the text.
 */
describe("the theme's font is what gets measured", () => {
  test("a larger theme font needs more lines and taller ones", () => {
    const base = setup({ cells: { "0_0": TWO_LINES }, styles: WRAP });
    const large = setup({
      cells: { "0_0": TWO_LINES },
      styles: WRAP,
      theme: { fontSize: "26px" },
    });

    expect(large.heightOf(0)).toBeGreaterThan(base.heightOf(0));
  });

  test("a cell's own fontSize still overrides the theme's", () => {
    const themed = setup({
      cells: { "0_0": TWO_LINES },
      styles: { "0_0": { wrap: true, fontSize: 13 } },
      theme: { fontSize: "26px" },
    });
    const base = setup({ cells: { "0_0": TWO_LINES }, styles: WRAP });

    expect(themed.heightOf(0)).toBe(base.heightOf(0));
  });

  test("a fontSize the measurer cannot read falls back to the default", () => {
    // `em` resolves against an ancestor that does not exist at measuring time.
    // Falling back beats guessing; see docs/LIMITATIONS.md.
    const relative = setup({
      cells: { "0_0": TWO_LINES },
      styles: WRAP,
      theme: { fontSize: "2em" },
    });
    const base = setup({ cells: { "0_0": TWO_LINES }, styles: WRAP });

    expect(relative.heightOf(0)).toBe(base.heightOf(0));
  });
});

/**
 * A merge renders as one box spanning several tracks, so it wraps at the width
 * of all of them and is already as tall as all of them.
 */
describe("merged cells", () => {
  const ACROSS_THREE = [{ r1: 0, c1: 0, r2: 0, c2: 2 }];
  const DOWN_TWO = [{ r1: 0, c1: 0, r2: 1, c2: 0 }];

  test("a merge wraps at its own width, not its first column's", () => {
    const merged = setup({
      cells: { "0_0": FOUR_LINES },
      styles: WRAP,
      merges: ACROSS_THREE,
    });
    const single = setup({ cells: { "0_0": FOUR_LINES }, styles: WRAP });

    // Three columns hold the same text in fewer lines, so the row asks for less.
    expect(merged.heightOf(0)).toBeLessThan(single.heightOf(0));
    expect(merged.heightOf(0)).toBeGreaterThan(ROW_HEIGHT);
  });

  test("text that fits across the merge leaves the row alone", () => {
    const { heightOf } = setup({
      cells: { "0_0": TWO_LINES },
      styles: WRAP,
      merges: ACROSS_THREE,
    });

    expect(heightOf(0)).toBe(ROW_HEIGHT);
  });

  test("a cell the merge covers grows nothing — it renders nothing", () => {
    const { heightOf } = setup({
      cells: { "0_1": FOUR_LINES },
      styles: { "0_1": { wrap: true } },
      merges: ACROSS_THREE,
    });

    expect(heightOf(0)).toBe(ROW_HEIGHT);
  });

  test("a merge down two rows only asks for the height they do not already give", () => {
    const merged = setup({
      cells: { "0_0": FOUR_LINES },
      styles: WRAP,
      merges: DOWN_TWO,
    });
    const single = setup({ cells: { "0_0": FOUR_LINES }, styles: WRAP });

    // Same text, same width, so the same total — spread over the two rows.
    expect(merged.heightOf(0) + merged.heightOf(1)).toBe(single.heightOf(0));
    // And it grows downwards: the first row keeps the height it had.
    expect(merged.heightOf(0)).toBe(ROW_HEIGHT);
  });
});
