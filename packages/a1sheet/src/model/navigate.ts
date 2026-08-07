/**
 * Where the keyboard jumps to.
 *
 * Ctrl+Arrow in a spreadsheet is not "move a lot" — it is "move to the edge of
 * this block of data", which is what makes a hundred-thousand-row sheet navigable
 * without a mouse. The rule, from Excel and Sheets:
 *
 * - Standing on filled ground with a filled neighbour: run to the LAST filled
 *   cell before the next gap.
 * - Standing next to a gap: skip the gap and stop on the FIRST filled cell after
 *   it.
 * - Nothing filled ahead at all: go to the far edge of the sheet.
 *
 * Framework-agnostic and pure, so the keyboard layer above it stays a switch
 * statement over key names.
 */
import { cellKey } from "./address.js";

export interface GridExtent {
  numRows: number;
  numCols: number;
}

export interface Step {
  /** -1, 0, or 1 on each axis. Both zero is a programmer error, not a no-op. */
  dRow: number;
  dCol: number;
}

const filled = (cells: Record<string, string>, row: number, col: number) =>
  (cells[cellKey(row, col)] ?? "") !== "";

/**
 * The cell Ctrl+Arrow lands on, starting from `row`/`col`.
 *
 * Returns the starting cell when the step would leave the sheet, so a caller can
 * apply the result unconditionally.
 */
export function dataEdge(
  cells: Record<string, string>,
  row: number,
  col: number,
  step: Step,
  extent: GridExtent,
): { row: number; col: number } {
  const { dRow, dCol } = step;
  if (dRow === 0 && dCol === 0) {
    throw new Error("dataEdge: step must move on at least one axis");
  }

  const inside = (r: number, c: number) =>
    r >= 0 && c >= 0 && r < extent.numRows && c < extent.numCols;

  let r = row;
  let c = col;
  if (!inside(r + dRow, c + dCol)) return { row, col };

  // Which of the two rules applies. Both the cell under the cursor and its
  // neighbour matter: standing on the last filled cell of a block the next press
  // must cross the gap rather than stay, and standing on a blank it must stop on
  // the first filled cell it reaches rather than run to the end of that block.
  const crossingGap = !filled(cells, r, c) || !filled(cells, r + dRow, c + dCol);

  while (inside(r + dRow, c + dCol)) {
    const nextFilled = filled(cells, r + dRow, c + dCol);
    if (crossingGap) {
      r += dRow;
      c += dCol;
      // First filled cell after the gap wins; an empty run to the sheet's edge
      // falls out of the loop and lands there.
      if (nextFilled) return { row: r, col: c };
    } else {
      if (!nextFilled) return { row: r, col: c };
      r += dRow;
      c += dCol;
    }
  }
  return { row: r, col: c };
}

/**
 * The bottom-right of the used range: Ctrl+End.
 *
 * Scans the keys rather than the extent, so an empty sheet costs nothing and a
 * sparse one costs its filled cells rather than its declared size.
 */
export function lastUsedCell(cells: Record<string, string>): {
  row: number;
  col: number;
} {
  let row = 0;
  let col = 0;
  for (const [key, value] of Object.entries(cells)) {
    if (value === "") continue;
    const split = key.indexOf("_");
    const r = Number(key.slice(0, split));
    const c = Number(key.slice(split + 1));
    if (r > row) row = r;
    if (c > col) col = c;
  }
  return { row, col };
}

/** The last filled column in one row, for End; the row itself when it is empty. */
export function lastUsedInRow(
  cells: Record<string, string>,
  row: number,
  numCols: number,
): number {
  for (let c = numCols - 1; c >= 0; c--) {
    if (filled(cells, row, c)) return c;
  }
  return 0;
}
