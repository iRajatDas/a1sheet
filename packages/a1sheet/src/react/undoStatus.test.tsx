/**
 * Undo/redo status messages on the public API.
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

describe("undo/redo status", () => {
  test("undo reports Action was undone", () => {
    let api: ReturnType<typeof useSpreadsheet> | undefined;
    function Probe() {
      api = useSpreadsheet({
        initialWorkbook: workbook({
          cells: { "0_0": "a" } as Record<CellKey, string>,
        }),
      });
      return null;
    }
    render(<Probe />);
    act(() => {
      api!.setCell(0, 0, "b");
    });
    act(() => {
      api!.undo();
    });
    expect(api!.getRaw(0, 0)).toBe("a");
    expect(api!.status).toBe("Action was undone.");
  });

  test("redo with nothing to redo reports Couldn't redo", () => {
    let api: ReturnType<typeof useSpreadsheet> | undefined;
    function Probe() {
      api = useSpreadsheet({
        initialWorkbook: workbook({}),
      });
      return null;
    }
    render(<Probe />);
    act(() => {
      api!.redo();
    });
    expect(api!.status).toBe("Couldn't redo. Try again.");
  });

  test("undo with nothing to undo is quiet", () => {
    let api: ReturnType<typeof useSpreadsheet> | undefined;
    function Probe() {
      api = useSpreadsheet({ initialWorkbook: workbook({}) });
      return null;
    }
    render(<Probe />);
    act(() => {
      api!.undo();
    });
    expect(api!.status).toBe("");
  });
});
