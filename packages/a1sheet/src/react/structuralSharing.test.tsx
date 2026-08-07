/**
 * What an edit does NOT copy.
 *
 * Every sheet operation used to go through a full clone, and a clone copies
 * `cells` — 66 ms of the 70 a clone costs on a workbook of a million filled
 * cells. Column resize fires on every mousemove of a drag, so that was a
 * seventy-millisecond stall per frame for an operation that changes one number.
 *
 * The property that makes it cheap is identity: an operation confined to a
 * row/column-keyed container leaves `cells` and `styles` as the SAME objects.
 * Asserting identity rather than a duration is what makes this a test rather
 * than a benchmark — it fails for the right reason on a slow machine too.
 */
import { describe, expect, test } from "bun:test";
import { act, render } from "@testing-library/react";
import { createRef } from "react";
import type { CellKey, RawCell } from "../model/types.js";
import { createWorkbook } from "../model/workbook.js";
import { Grid } from "./components/Grid.js";
import { Root, type SheetRootHandle } from "./Root.js";

function mounted() {
  const wb = createWorkbook(["Sheet1"]);
  Object.assign(wb.sheets[0]?.cells as Record<CellKey, RawCell>, {
    "0_0": "a",
    "1_0": "b",
  });
  Object.assign(wb.sheets[0]?.styles as object, { "0_0": { bold: true } });

  const ref = createRef<SheetRootHandle>();
  render(
    <Root ref={ref} defaultWorkbook={wb}>
      <Grid />
    </Root>,
  );
  return ref;
}

/** Each case: the operation, and the fields it is allowed to replace. */
const CHEAP = [
  ["setColWidth", (api: SheetRootHandle["api"]) => api.setColWidth(2, 140)],
  ["setRowHeight", (api: SheetRootHandle["api"]) => api.setRowHeight(2, 40)],
  ["setColLabel", (api: SheetRootHandle["api"]) => api.setColLabel(1, "Qty")],
  ["toggleColHidden", (api: SheetRootHandle["api"]) => api.toggleColHidden(1)],
  ["toggleRowHidden", (api: SheetRootHandle["api"]) => api.toggleRowHidden(1)],
  ["freezeToSelection", (api: SheetRootHandle["api"]) => api.freezeToSelection()],
  ["appendRows", (api: SheetRootHandle["api"]) => api.appendRows(100)],
  ["setFilter", (api: SheetRootHandle["api"]) => api.setFilter(0, new Set(["a"]))],
] as const;

describe("an operation that touches no cell copies no cells", () => {
  for (const [name, run] of CHEAP) {
    test(`${name} leaves cells and styles untouched`, () => {
      const ref = mounted();
      const before = ref.current?.api.sheet;
      if (!before) throw new Error("not mounted");

      act(() => {
        const api = ref.current?.api;
        if (api) run(api);
      });

      const after = ref.current?.api.sheet;
      expect(after).not.toBe(before);
      expect(after?.cells).toBe(before.cells);
      expect(after?.styles).toBe(before.styles);
    });
  }
});

describe("surface edits skip formula recalculation", () => {
  test("applyStyle keeps the same evaluator", () => {
    const ref = mounted();
    const before = ref.current?.api.evaluator;
    act(() => ref.current?.api.applyStyle({ bold: true }));
    expect(ref.current?.api.evaluator).toBe(before);
  });

  test("freeze keeps the same evaluator", () => {
    const ref = mounted();
    const before = ref.current?.api.evaluator;
    act(() => ref.current?.api.freezeToSelection());
    expect(ref.current?.api.evaluator).toBe(before);
  });

  test("setCell rebuilds the evaluator", () => {
    const ref = mounted();
    const before = ref.current?.api.evaluator;
    act(() => ref.current?.api.setCell(0, 0, "changed"));
    expect(ref.current?.api.evaluator).not.toBe(before);
  });
});

describe("the operations still do what they say", () => {
  test("a resize is visible on the sheet", () => {
    const ref = mounted();
    act(() => ref.current?.api.setColWidth(2, 140));
    expect(ref.current?.api.sheet.colWidths[2]).toBe(140);
  });

  test("a reset removes the entry rather than writing a default", () => {
    // `useRowWindow` skips its offset table entirely while `rowHeights` is
    // empty, and it can only become empty again this way.
    const ref = mounted();
    act(() => ref.current?.api.setRowHeight(2, 40));
    act(() => ref.current?.api.resetRowHeight(2));
    expect(ref.current?.api.sheet.rowHeights).toEqual({});
  });

  test("toggling hidden twice returns to where it started", () => {
    const ref = mounted();
    act(() => ref.current?.api.toggleColHidden(1));
    expect(ref.current?.api.sheet.hiddenCols.has(1)).toBe(true);
    act(() => ref.current?.api.toggleColHidden(1));
    expect(ref.current?.api.sheet.hiddenCols.has(1)).toBe(false);
  });

  test("clearing a filter drops the entry", () => {
    const ref = mounted();
    act(() => ref.current?.api.setFilter(0, new Set(["a"])));
    expect(ref.current?.api.sheet.filters[0]).toBeDefined();
    act(() => ref.current?.api.setFilter(0, null));
    expect(ref.current?.api.sheet.filters).toEqual({});
  });

  test("a patch does not leak back into the previous sheet", () => {
    // The point of copying the container: undo has to find the old value.
    const ref = mounted();
    const before = ref.current?.api.sheet;
    act(() => ref.current?.api.setColWidth(2, 140));
    expect(before?.colWidths[2]).toBeUndefined();
  });

  test("an edit that does touch cells still replaces them", () => {
    const ref = mounted();
    const before = ref.current?.api.sheet;
    act(() => ref.current?.api.setCell(4, 4, "typed"));
    expect(ref.current?.api.sheet.cells).not.toBe(before?.cells);
    expect(ref.current?.api.sheet.cells["4_4" as CellKey]).toBe("typed");
  });

  test("applyStyle replaces styles but not cells", () => {
    const ref = mounted();
    const before = ref.current?.api.sheet;
    act(() => ref.current?.api.applyStyle({ color: "#ff0000" }));
    const after = ref.current?.api.sheet;
    expect(after?.cells).toBe(before?.cells);
    expect(after?.styles).not.toBe(before?.styles);
    expect(after?.styles["0_0"]?.color).toBe("#ff0000");
  });
});
