/**
 * Ctrl+Arrow. The rule is small and every case of it is a different rule, which
 * is why this is a table of them rather than a couple of examples.
 */
import { describe, expect, test } from "bun:test";
import { dataEdge, lastUsedCell, lastUsedInRow } from "./navigate.js";

const extent = { numRows: 100, numCols: 26 };
const DOWN = { dRow: 1, dCol: 0 };
const UP = { dRow: -1, dCol: 0 };
const RIGHT = { dRow: 0, dCol: 1 };

/** A column of values starting at row 0, blanks written as "". */
function column(values: string[]): Record<string, string> {
  const cells: Record<string, string> = {};
  values.forEach((v, r) => {
    if (v !== "") cells[`${r}_0`] = v;
  });
  return cells;
}

describe("running to the edge of a block", () => {
  test("from inside a block, stops on its last filled cell", () => {
    const cells = column(["a", "b", "c", "", "", "z"]);
    expect(dataEdge(cells, 0, 0, DOWN, extent)).toEqual({ row: 2, col: 0 });
  });

  test("from the last cell of a block, crosses the gap to the next", () => {
    // Not "stays put": pressing it twice has to get you to the next block, or
    // the shortcut cannot cross a sheet.
    const cells = column(["a", "b", "c", "", "", "z"]);
    expect(dataEdge(cells, 2, 0, DOWN, extent)).toEqual({ row: 5, col: 0 });
  });

  test("from a blank, stops on the FIRST filled cell it reaches", () => {
    const cells = column(["", "", "a", "b", "c"]);
    expect(dataEdge(cells, 0, 0, DOWN, extent)).toEqual({ row: 2, col: 0 });
  });

  test("from a blank with the block starting right below it, stops on top of it", () => {
    // The rule keys on the cell under the cursor as well as the one ahead.
    // Reading only the neighbour sends this to the END of the block instead.
    const cells = column(["", "a", "b", "c"]);
    expect(dataEdge(cells, 0, 0, DOWN, extent)).toEqual({ row: 1, col: 0 });
  });

  test("with nothing ahead, goes to the far edge of the sheet", () => {
    expect(dataEdge({}, 5, 0, DOWN, extent)).toEqual({ row: 99, col: 0 });
    expect(dataEdge({}, 5, 0, UP, extent)).toEqual({ row: 0, col: 0 });
  });

  test("at the edge already, stays where it is", () => {
    expect(dataEdge({}, 0, 0, UP, extent)).toEqual({ row: 0, col: 0 });
  });

  test("the same rules apply across a row", () => {
    const cells = { "0_0": "a", "0_1": "b", "0_5": "c" };
    expect(dataEdge(cells, 0, 0, RIGHT, extent)).toEqual({ row: 0, col: 1 });
    expect(dataEdge(cells, 0, 1, RIGHT, extent)).toEqual({ row: 0, col: 5 });
  });

  test("a step that moves nowhere is a programmer error", () => {
    expect(() => dataEdge({}, 0, 0, { dRow: 0, dCol: 0 }, extent)).toThrow(
      /at least one axis/,
    );
  });
});

describe("the used range", () => {
  test("the last filled cell is the corner of the two axes, not one cell", () => {
    // B10 and E2 filled: Ctrl+End goes to E10, which holds nothing.
    const cells = { "9_1": "x", "1_4": "y" };
    expect(lastUsedCell(cells)).toEqual({ row: 9, col: 4 });
  });

  test("an empty sheet ends where it starts", () => {
    expect(lastUsedCell({})).toEqual({ row: 0, col: 0 });
  });

  test("End travels along the row it is on", () => {
    const cells = { "0_3": "a", "5_9": "b" };
    expect(lastUsedInRow(cells, 0, 26)).toBe(3);
    expect(lastUsedInRow(cells, 5, 26)).toBe(9);
    expect(lastUsedInRow(cells, 1, 26)).toBe(0);
  });
});
