/**
 * F9, and what a calculation cycle is in the React layer.
 *
 * A volatile function holds its value for the life of one evaluator, and an
 * edit is what normally replaces the evaluator. Without an explicit trigger a
 * sheet whose only formula is `=RAND()` is frozen: correct within its cycle,
 * and never given another one.
 */
import { describe, expect, test } from "bun:test";
import { fireEvent, render } from "@testing-library/react";
import { createWorkbook } from "../model/workbook.js";
import { Spreadsheet } from "./Spreadsheet.js";

function setup(cells: Record<string, string>) {
  const wb = createWorkbook(["Sheet1"]);
  Object.assign((wb.sheets[0] as { cells: object }).cells, cells);
  const { container } = render(<Spreadsheet defaultWorkbook={wb} />);

  const hidden = container.querySelector("textarea") as HTMLTextAreaElement;
  /** Types into the active cell and commits, which is one calculation cycle. */
  const type = (text: string) => {
    fireEvent.keyDown(hidden, { key: text });
    const input = container.querySelector(".a1s-cell input") as HTMLInputElement;
    fireEvent.keyDown(input, { key: "Enter" });
  };
  const shown = (row: number, col: number) =>
    (
      container.querySelector(
        `.a1s-cell[data-row="${row}"][data-col="${col}"]`,
      ) as HTMLElement
    ).textContent;

  return { container, hidden, shown, type };
}

/** Distinct values A1 took across `times` presses of F9. */
function pressF9(
  hidden: HTMLElement,
  shown: () => string | null,
  times: number,
): Set<string | null> {
  const seen = new Set<string | null>([shown()]);
  for (let i = 0; i < times; i++) {
    fireEvent.keyDown(hidden, { key: "F9" });
    seen.add(shown());
  }
  return seen;
}

describe("F9", () => {
  test("gives RAND a new value", () => {
    const { hidden, shown } = setup({ "0_0": "=RAND()" });

    expect(pressF9(hidden, () => shown(0, 0), 8).size).toBeGreaterThan(1);
  });

  test("leaves an ordinary formula exactly where it was", () => {
    // The point of a recalculation is that it cannot change anything that is a
    // function of the cells, and those have not moved.
    const { hidden, shown } = setup({ "0_0": "2", "0_1": "=A1*21" });

    expect(shown(0, 1)).toBe("42");
    fireEvent.keyDown(hidden, { key: "F9" });
    expect(shown(0, 1)).toBe("42");
  });

  test("is not an edit, so undo reaches past it to the last real change", () => {
    // Asserting that the cell still reads 1 after a bare F9 + undo would pass
    // either way. The recalculation has to sit BETWEEN an edit and the undo for
    // the test to notice it consuming a history entry.
    const { hidden, shown, type } = setup({ "0_0": "1" });

    type("9");
    expect(shown(0, 0)).toBe("9");

    fireEvent.keyDown(hidden, { key: "F9" });
    fireEvent.keyDown(hidden, { key: "z", ctrlKey: true });

    expect(shown(0, 0)).toBe("1");
  });
});

describe("an edit is a calculation cycle too", () => {
  test("editing an unrelated cell moves NOW forward", async () => {
    // What distinguishes an edit that begins a cycle from one that merely
    // rebuilds the evaluator: RAND rerolls either way, because it reads the RNG
    // afresh. NOW reads the cycle's instant, so it only moves if there is one.
    const { hidden, shown, type } = setup({ "0_0": "=NOW()" });

    const before = shown(0, 0);
    await new Promise((resolve) => setTimeout(resolve, 5));
    fireEvent.keyDown(hidden, { key: "ArrowDown" });
    type("1");

    expect(shown(0, 0)).not.toBe(before);
  });

  test("editing an unrelated cell rerolls RAND, as it does in Excel", () => {
    const { hidden, shown, type } = setup({ "0_0": "=RAND()" });

    const seen = new Set([shown(0, 0)]);
    for (let i = 1; i <= 8; i++) {
      fireEvent.keyDown(hidden, { key: "ArrowDown" });
      type(String(i));
      seen.add(shown(0, 0));
    }

    expect(seen.size).toBeGreaterThan(1);
  });
});
