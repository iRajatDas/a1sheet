/**
 * Autofill: all directions, typed guards, weekday/month sequences.
 */
import { describe, expect, test } from "bun:test";
import { act, render } from "@testing-library/react";
import { previewFillCheck } from "../model/autofill.js";
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

describe("previewFillCheck", () => {
  test("ok when extending down a column", () => {
    const sheet = makeSheet("S");
    const result = previewFillCheck(sheet, {
      source: { r1: 0, c1: 0, r2: 2, c2: 0 },
      target: { row: 5, col: 0 },
    });
    expect(result.ok).toBe(true);
  });

  test("DEST_EQUALS_SOURCE when target stays inside the source", () => {
    const sheet = makeSheet("S");
    const result = previewFillCheck(sheet, {
      source: { r1: 0, c1: 0, r2: 2, c2: 0 },
      target: { row: 1, col: 0 },
    });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.code).toBe("DEST_EQUALS_SOURCE");
  });

  test("DEST_SOURCE_DIMENSIONS_BOTH_DIFFERENT when both axes expand", () => {
    const sheet = makeSheet("S");
    const result = previewFillCheck(sheet, {
      source: { r1: 0, c1: 0, r2: 1, c2: 1 },
      target: { row: 4, col: 4 },
    });
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.code).toBe("DEST_SOURCE_DIMENSIONS_BOTH_DIFFERENT");
      expect(result.message).toContain("both axes");
    }
  });

  test("PARTIAL_MERGE when the destination cuts through a merge", () => {
    const sheet = makeSheet("S");
    sheet.merges = [{ r1: 3, c1: 0, r2: 5, c2: 0 }];
    const result = previewFillCheck(sheet, {
      source: { r1: 0, c1: 0, r2: 0, c2: 0 },
      target: { row: 4, col: 0 },
    });
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.code).toBe("PARTIAL_MERGE");
      expect(result.message).toContain("partially intersects a merge");
    }
  });

  test("ok when the destination fully covers a merge", () => {
    const sheet = makeSheet("S");
    sheet.merges = [{ r1: 3, c1: 0, r2: 5, c2: 0 }];
    const result = previewFillCheck(sheet, {
      source: { r1: 0, c1: 0, r2: 0, c2: 0 },
      target: { row: 5, col: 0 },
    });
    expect(result.ok).toBe(true);
  });

  test("CONSTRAIN_DEST_FAILED when the target is past the grid", () => {
    const sheet = makeSheet("S");
    sheet.numRows = 10;
    sheet.numCols = 5;
    const result = previewFillCheck(sheet, {
      source: { r1: 0, c1: 0, r2: 0, c2: 0 },
      target: { row: 20, col: 0 },
    });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.code).toBe("CONSTRAIN_DEST_FAILED");
  });
});

describe("fill handle — sequences and guards", () => {
  test("weekday names continue and wrap", () => {
    const { get, run } = api({
      cells: { "0_0": "Monday", "1_0": "Tuesday" } as Record<CellKey, string>,
    });
    run((a) => a.fill.start({ r1: 0, c1: 0, r2: 1, c2: 0 }));
    run((a) => a.fill.moveTo(4, 0));
    run((a) =>
      a.fill.commit(a.updateSheet, { sheet: a.sheet, onReject: a.setStatus }),
    );
    expect(get().getRaw(2, 0)).toBe("Wednesday");
    expect(get().getRaw(3, 0)).toBe("Thursday");
    expect(get().getRaw(4, 0)).toBe("Friday");
  });

  test("month abbreviations continue", () => {
    const { get, run } = api({
      cells: { "0_0": "Jan", "0_1": "Feb" } as Record<CellKey, string>,
    });
    run((a) => a.fill.start({ r1: 0, c1: 0, r2: 0, c2: 1 }));
    run((a) => a.fill.moveTo(0, 3));
    run((a) =>
      a.fill.commit(a.updateSheet, { sheet: a.sheet, onReject: a.setStatus }),
    );
    expect(get().getRaw(0, 2)).toBe("Mar");
    expect(get().getRaw(0, 3)).toBe("Apr");
  });

  test("both-axes conflict sets status and writes nothing", () => {
    const { get, run } = api({
      cells: { "0_0": "1", "0_1": "2", "1_0": "3", "1_1": "4" } as Record<
        CellKey,
        string
      >,
    });
    run((a) => a.fill.start({ r1: 0, c1: 0, r2: 1, c2: 1 }));
    run((a) => a.fill.moveTo(4, 4));
    run((a) =>
      a.fill.commit(a.updateSheet, { sheet: a.sheet, onReject: a.setStatus }),
    );
    expect(get().getRaw(2, 0)).toBe("");
    expect(get().getRaw(0, 2)).toBe("");
    expect(get().status).toContain("both axes");
  });

  test("partial merge sets status and writes nothing", () => {
    const { get, run } = api({
      cells: { "0_0": "x" } as Record<CellKey, string>,
      merges: [{ r1: 2, c1: 0, r2: 4, c2: 0 }],
    });
    run((a) => a.fill.start({ r1: 0, c1: 0, r2: 0, c2: 0 }));
    run((a) => a.fill.moveTo(3, 0));
    run((a) =>
      a.fill.commit(a.updateSheet, { sheet: a.sheet, onReject: a.setStatus }),
    );
    expect(get().getRaw(1, 0)).toBe("");
    expect(get().status).toContain("partially intersects a merge");
  });

  test("filling over a whole merge writes into the top-left", () => {
    const { get, run } = api({
      cells: { "0_0": "1" } as Record<CellKey, string>,
      merges: [{ r1: 2, c1: 0, r2: 4, c2: 0 }],
    });
    run((a) => a.fill.start({ r1: 0, c1: 0, r2: 0, c2: 0 }));
    run((a) => a.fill.moveTo(4, 0));
    run((a) =>
      a.fill.commit(a.updateSheet, { sheet: a.sheet, onReject: a.setStatus }),
    );
    expect(get().getRaw(2, 0)).toBe("3");
    expect(get().status).toBe("");
  });

  test("numeric series still extrapolates linearly", () => {
    const { get, run } = api({
      cells: { "0_0": "1", "1_0": "2" } as Record<CellKey, string>,
    });
    run((a) => a.fill.start({ r1: 0, c1: 0, r2: 1, c2: 0 }));
    run((a) => a.fill.moveTo(3, 0));
    run((a) =>
      a.fill.commit(a.updateSheet, { sheet: a.sheet, onReject: a.setStatus }),
    );
    expect(get().getRaw(2, 0)).toBe("3");
    expect(get().getRaw(3, 0)).toBe("4");
  });
});
