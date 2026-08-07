/**
 * The grid and editing gaps that were documented rather than fixed: filling
 * upward and leftward, hiding a column, and data-validation dropdowns.
 */
import { describe, expect, test } from "bun:test";
import { act, fireEvent, render, screen } from "@testing-library/react";
import { makeSheet } from "../model/sheet.js";
import type { CellKey, DataValidation, Sheet, Workbook } from "../model/types.js";
import { listLiterals } from "../model/validation.js";
import { Spreadsheet } from "./Spreadsheet.js";
import { useSpreadsheet } from "./useSpreadsheet.js";

function workbook(sheet: Partial<Sheet>): Workbook {
  return {
    sheets: [{ ...makeSheet("S"), ...sheet }],
    activeSheetIndex: 0,
    namedRanges: {},
  };
}

type Api = ReturnType<typeof useSpreadsheet>;

/**
 * Drives the headless hook, which is where the fill and hiding logic lives.
 *
 * `run` wraps each call in `act`, so the re-render it causes has flushed before
 * the next read — without it every assertion sees the state from before the
 * call it is testing.
 */
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

describe("filling upward and leftward", () => {
  test("a series dragged upward counts down, not on", () => {
    // 1, 2, 3 filled upward continues 0, -1, -2. Extrapolating forward and
    // placing the result above would give 4, 5, 6 — an answer, and backwards.
    const { get, run } = api({
      cells: { "3_0": "1", "4_0": "2", "5_0": "3" } as Record<CellKey, string>,
    });

    run((a) => a.fill.start({ r1: 3, c1: 0, r2: 5, c2: 0 }));
    run((a) => a.fill.moveTo(0, 0));
    run((a) => a.fill.commit(a.updateSheet));

    // The raw content, which is what a fill writes. A literal cell evaluates to
    // its own text, so the values are strings here and not numbers.
    expect(get().getRaw(2, 0)).toBe("0");
    expect(get().getRaw(1, 0)).toBe("-1");
    expect(get().getRaw(0, 0)).toBe("-2");
  });

  test("a series dragged leftward does the same across", () => {
    const { get, run } = api({
      cells: { "0_3": "10", "0_4": "20" } as Record<CellKey, string>,
    });

    run((a) => a.fill.start({ r1: 0, c1: 3, r2: 0, c2: 4 }));
    run((a) => a.fill.moveTo(0, 1));
    run((a) => a.fill.commit(a.updateSheet));

    expect(get().getRaw(0, 2)).toBe("0");
    expect(get().getRaw(0, 1)).toBe("-10");
  });

  test("a formula dragged upward shifts its references upward", () => {
    const { get, run } = api({
      cells: { "0_0": "5", "1_0": "6", "2_0": "7", "2_1": "=A3*2" } as Record<
        CellKey,
        string
      >,
    });

    run((a) => a.fill.start({ r1: 2, c1: 1, r2: 2, c2: 1 }));
    run((a) => a.fill.moveTo(0, 1));
    run((a) => a.fill.commit(a.updateSheet));

    expect(get().getRaw(1, 1)).toBe("=A2*2");
    expect(get().getRaw(0, 1)).toBe("=A1*2");
    expect(get().getValue(0, 1)).toBe(10);
  });

  test("the preview covers the direction the drag went", () => {
    const { get, run } = api({ cells: { "5_5": "1" } as Record<CellKey, string> });

    run((a) => a.fill.start({ r1: 5, c1: 5, r2: 5, c2: 5 }));
    run((a) => a.fill.moveTo(2, 2));

    expect(get().fill.preview).toEqual({ r1: 2, c1: 2, r2: 5, c2: 5 });
  });

  test("dragging back to the source fills nothing", () => {
    const { get, run } = api({ cells: { "0_0": "1" } as Record<CellKey, string> });

    run((a) => a.fill.start({ r1: 0, c1: 0, r2: 0, c2: 0 }));
    run((a) => a.fill.moveTo(0, 0));

    expect(get().fill.commit(get().updateSheet)).toBeNull();
  });
});

describe("hiding a column", () => {
  test("a hidden column has no width, so everything after it shifts left", () => {
    // Zero width rather than absent: the cumulative offsets, the CSS grid
    // tracks, the sticky freeze positions, and the hit test are all written in
    // terms of a column's width, and a "visible columns" mapping would need a
    // second notion of where each column sits.
    const { get, run } = api({});
    const before = get().colWindow.colOffset(3);

    run((a) => a.toggleColHidden(1));

    expect(get().colWindow.colWidth(1)).toBe(0);
    expect(get().colWindow.colOffset(3)).toBeLessThan(before);
  });

  test("hiding is a toggle", () => {
    const { get, run } = api({});
    run((a) => a.toggleColHidden(2));
    expect(get().sheet.hiddenCols.has(2)).toBe(true);
    run((a) => a.toggleColHidden(2));
    expect(get().sheet.hiddenCols.has(2)).toBe(false);
  });

  test("the context menu offers it", () => {
    render(<Spreadsheet />);
    const cell = document.querySelector('[data-row="0"][data-col="0"]');
    if (!cell) throw new Error("no cell");
    fireEvent.contextMenu(cell);

    expect(screen.getByText("Hide column")).toBeDefined();
  });
});

describe("data-validation dropdowns", () => {
  const listRule: DataValidation = {
    range: { r1: 0, c1: 0, r2: 2, c2: 0 },
    kind: "list",
    formulas: ['"Arsenal,Chelsea,Fulham"'],
  };

  test("a literal list is offered as its items", () => {
    // Excel writes a literal list as ONE quoted comma-separated string, so
    // reading it as a single value offers a dropdown with one nonsense entry.
    expect(listLiterals(listRule)).toEqual(["Arsenal", "Chelsea", "Fulham"]);
  });

  test("a range-backed list is not literal, and resolves through the sheet", () => {
    const ranged: DataValidation = {
      ...listRule,
      formulas: ["$C$1:$C$2"],
    };
    expect(listLiterals(ranged)).toBeNull();

    const { get } = api({
      cells: { "0_2": "yes", "1_2": "no" } as Record<CellKey, string>,
      validations: [ranged],
    });
    expect(get().choicesFor(0, 0)).toEqual(["yes", "no"]);
  });

  test("only the cells the rule covers get choices", () => {
    const { get } = api({ validations: [listRule] });

    expect(get().choicesFor(0, 0)).toEqual(["Arsenal", "Chelsea", "Fulham"]);
    expect(get().choicesFor(9, 9)).toBeUndefined();
  });

  test("the active cell shows a dropdown affordance", () => {
    render(<Spreadsheet defaultWorkbook={workbook({ validations: [listRule] })} />);

    // A1 is active on mount.
    expect(screen.getByLabelText("Choose a value")).toBeDefined();
  });

  test("the editor offers the choices without preventing typing", () => {
    // A <select> would take away typing, which Excel allows. A datalist does not.
    render(<Spreadsheet defaultWorkbook={workbook({ validations: [listRule] })} />);
    const cell = document.querySelector('[data-row="0"][data-col="0"]');
    if (!cell) throw new Error("no cell");
    fireEvent.doubleClick(cell);

    const options = document.querySelectorAll("datalist option");
    expect([...options].map((o) => o.getAttribute("value"))).toEqual([
      "Arsenal",
      "Chelsea",
      "Fulham",
    ]);
  });
});
