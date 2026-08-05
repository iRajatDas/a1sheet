import { describe, expect, test } from "bun:test";
import {
  canRedo,
  canUndo,
  emptyHistory,
  HISTORY_LIMIT,
  push,
  redo,
  undo,
} from "./history.js";
import type { Workbook } from "./types.js";
import { createWorkbook } from "./workbook.js";

function wbNamed(name: string): Workbook {
  const wb = createWorkbook();
  const first = wb.sheets[0];
  if (first) first.name = name;
  return wb;
}

describe("push", () => {
  test("records a snapshot and clears the redo stack", () => {
    let h = push(emptyHistory(), wbNamed("a"));
    h = { ...h, future: [wbNamed("f")] };
    h = push(h, wbNamed("b"));
    expect(h.past).toHaveLength(2);
    expect(h.future).toHaveLength(0);
  });

  test("caps the past stack at HISTORY_LIMIT, dropping the oldest", () => {
    let h = emptyHistory();
    for (let i = 0; i < HISTORY_LIMIT + 10; i++) h = push(h, wbNamed(`s${i}`));
    expect(h.past).toHaveLength(HISTORY_LIMIT);
    expect(h.past[0]?.sheets[0]?.name).toBe("s10");
  });
});

describe("undo / redo", () => {
  test("undo returns the previous snapshot and banks the current one", () => {
    const before = wbNamed("before");
    const current = wbNamed("current");
    const h = push(emptyHistory(), before);

    const r = undo(h, current);
    expect(r).not.toBeNull();
    expect(r?.workbook.sheets[0]?.name).toBe("before");
    expect(r?.history.future[0]?.sheets[0]?.name).toBe("current");
    expect(r?.history.past).toHaveLength(0);
  });

  test("redo reverses an undo exactly", () => {
    const before = wbNamed("before");
    const current = wbNamed("current");

    const undone = undo(push(emptyHistory(), before), current);
    expect(undone).not.toBeNull();
    const redone = redo(undone!.history, undone!.workbook);

    expect(redone?.workbook.sheets[0]?.name).toBe("current");
    expect(redone?.history.past[0]?.sheets[0]?.name).toBe("before");
  });

  test("both return null at the ends of the stack", () => {
    expect(undo(emptyHistory(), wbNamed("x"))).toBeNull();
    expect(redo(emptyHistory(), wbNamed("x"))).toBeNull();
  });
});

describe("canUndo / canRedo", () => {
  test("reflect stack contents", () => {
    expect(canUndo(emptyHistory())).toBe(false);
    expect(canRedo(emptyHistory())).toBe(false);
    expect(canUndo(push(emptyHistory(), wbNamed("a")))).toBe(true);
  });
});
