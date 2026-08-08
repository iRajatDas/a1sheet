/**
 * Freeze panes: model counts, status strings, coordinates stay absolute.
 */
import { describe, expect, test } from "bun:test";
import { act, render } from "@testing-library/react";
import { createRef } from "react";
import { createWorkbook } from "../model/workbook.js";
import { type SheetRootHandle } from "./Root.js";
import { Spreadsheet } from "./Spreadsheet.js";

describe("freeze panes", () => {
  test("freezeToSelection sets counts and reports status", () => {
    const ref = createRef<SheetRootHandle>();
    render(<Spreadsheet ref={ref} />);

    act(() => {
      ref.current?.api.select({ r1: 0, c1: 0, r2: 1, c2: 0 });
    });
    act(() => {
      ref.current?.api.freezeToSelection();
    });

    expect(ref.current?.api.sheet.frozenRows).toBe(2);
    expect(ref.current?.api.sheet.frozenCols).toBe(1);
    expect(ref.current?.api.status).toBe("Freeze rows. Freeze columns.");
  });

  test("getValue / setCell ignore freeze — addresses stay absolute", () => {
    const wb = createWorkbook(["S"]);
    Object.assign(wb.sheets[0] as object, {
      frozenRows: 2,
      frozenCols: 1,
      cells: { "3_2": "deep" },
    });
    const ref = createRef<SheetRootHandle>();
    render(<Spreadsheet defaultWorkbook={wb} ref={ref} />);

    expect(ref.current?.api.getRaw(3, 2)).toBe("deep");
    act(() => {
      ref.current?.api.setCell(3, 2, "moved");
    });
    expect(ref.current?.api.getRaw(3, 2)).toBe("moved");
    expect(ref.current?.api.sheet.frozenRows).toBe(2);
  });

  test("unfreeze clears panes with plural status", () => {
    const wb = createWorkbook(["S"]);
    Object.assign(wb.sheets[0] as object, { frozenRows: 2, frozenCols: 2 });
    const ref = createRef<SheetRootHandle>();
    render(<Spreadsheet defaultWorkbook={wb} ref={ref} />);

    act(() => {
      ref.current?.api.unfreeze();
    });
    expect(ref.current?.api.sheet.frozenRows).toBe(0);
    expect(ref.current?.api.sheet.frozenCols).toBe(0);
    expect(ref.current?.api.status).toBe("Unfreeze rows. Unfreeze columns.");
  });

  test("frozen rows stay in the row window list", () => {
    const wb = createWorkbook(["S"]);
    Object.assign(wb.sheets[0] as object, {
      frozenRows: 2,
      numRows: 50,
    });
    const ref = createRef<SheetRootHandle>();
    render(<Spreadsheet defaultWorkbook={wb} ref={ref} />);

    const frozen = ref.current?.api.rowWindow.frozenRowsList ?? [];
    expect(frozen.map((r) => r.absRow)).toEqual([0, 1]);
  });
});
