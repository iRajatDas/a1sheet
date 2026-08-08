/**
 * Paste-special modes — values, formats, formulas, transpose, text.
 */
import { describe, expect, test } from "bun:test";
import { act, render } from "@testing-library/react";
import { makeSheet } from "../model/sheet.js";
import type { CellKey, Sheet, Workbook } from "../model/types.js";
import { useSpreadsheet } from "./useSpreadsheet.js";

function workbook(sheet: Partial<Sheet>): Workbook {
  return {
    sheets: [{ ...makeSheet("S"), ...sheet }],
    activeSheetIndex: 0,
    namedRanges: {},
  };
}

type Api = ReturnType<typeof useSpreadsheet>;

function api(sheet: Partial<Sheet>) {
  let captured: Api | undefined;
  function Probe() {
    captured = useSpreadsheet({ initialWorkbook: workbook(sheet) });
    return null;
  }
  render(<Probe />);
  if (!captured) throw new Error("hook did not run");
  const get = () => captured as Api;
  const run = (fn: (a: Api) => void) => {
    act(() => {
      fn(get());
    });
  };
  return { get, run };
}

function pasteAt(
  a: Api,
  text: string,
  target: { row: number; col: number },
  mode: import("./useClipboard.js").PasteMode,
) {
  a.clipboard.paste(text, target, a.updateSheet, {
    mode,
    evaluator: a.evaluator,
    onReject: a.setStatus,
  });
}

describe("paste modes", () => {
  test("values pastes evaluated results, not formulas", () => {
    const { get, run } = api({
      cells: { "0_0": "10", "0_1": "=A1*2" } as Record<CellKey, string>,
    });
    let text = "";
    run((a) => {
      text = a.clipboard.copy(a.sheet, [{ r1: 0, c1: 0, r2: 0, c2: 1 }]) ?? "";
    });
    run((a) => pasteAt(a, text, { row: 2, col: 0 }, "values"));
    expect(get().getRaw(2, 0)).toBe("10");
    expect(get().getRaw(2, 1)).toBe("20");
    expect(get().status).toContain("Pasted values");
  });

  test("formats pastes styles only and ignores locked for values", () => {
    const { get, run } = api({
      cells: {
        "0_0": "hello",
        "1_0": "keep",
      } as Record<CellKey, string>,
      styles: {
        "0_0": { bold: true, color: "#ff0000" },
        "1_0": { locked: true },
      } as Sheet["styles"],
    });
    let text = "";
    run((a) => {
      text = a.clipboard.copy(a.sheet, [{ r1: 0, c1: 0, r2: 0, c2: 0 }]) ?? "";
    });
    run((a) => pasteAt(a, text, { row: 1, col: 0 }, "formats"));
    expect(get().getRaw(1, 0)).toBe("keep");
    expect(get().sheet.styles["1_0"]?.bold).toBe(true);
    expect(get().sheet.styles["1_0"]?.color).toBe("#ff0000");
    expect(get().sheet.styles["1_0"]?.locked).toBe(true);
    expect(get().status).toContain("Pasted formats");
  });

  test("formulas shift references on internal paste", () => {
    const { get, run } = api({
      cells: { "0_0": "3", "0_1": "=A1+1" } as Record<CellKey, string>,
    });
    let text = "";
    run((a) => {
      text = a.clipboard.copy(a.sheet, [{ r1: 0, c1: 1, r2: 0, c2: 1 }]) ?? "";
    });
    run((a) => pasteAt(a, text, { row: 2, col: 1 }, "formulas"));
    expect(get().getRaw(2, 1)).toBe("=A3+1");
    expect(get().status).toContain("Pasted formulas");
  });

  test("transpose swaps rows and columns", () => {
    const { get, run } = api({
      cells: {
        "0_0": "a",
        "0_1": "b",
        "1_0": "c",
        "1_1": "d",
      } as Record<CellKey, string>,
    });
    let text = "";
    run((a) => {
      text = a.clipboard.copy(a.sheet, [{ r1: 0, c1: 0, r2: 1, c2: 1 }]) ?? "";
    });
    run((a) => pasteAt(a, text, { row: 3, col: 0 }, "transpose"));
    expect(get().getRaw(3, 0)).toBe("a");
    expect(get().getRaw(3, 1)).toBe("c");
    expect(get().getRaw(4, 0)).toBe("b");
    expect(get().getRaw(4, 1)).toBe("d");
    expect(get().status).toContain("Pasted transposed");
  });

  test("text pastes a formula as a literal string", () => {
    const { get, run } = api({});
    run((a) => pasteAt(a, "=1+1", { row: 0, col: 0 }, "text"));
    expect(get().getRaw(0, 0)).toBe("'=1+1");
    expect(get().getDisplay(0, 0)).toBe("=1+1");
    expect(get().status).toContain("Pasted as text");
  });

  test("values mode skips locked cells", () => {
    const { get, run } = api({
      cells: { "0_0": "x", "1_0": "locked" } as Record<CellKey, string>,
      styles: { "1_0": { locked: true } } as Sheet["styles"],
    });
    run((a) => pasteAt(a, "new", { row: 1, col: 0 }, "values"));
    expect(get().getRaw(1, 0)).toBe("locked");
  });
});
