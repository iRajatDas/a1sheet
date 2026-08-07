/**
 * Defined names Excel scoped to one sheet.
 *
 * These used to be dropped on import: a name with a `localSheetId` was skipped
 * outright, so every formula using it read `#NAME?` behind whatever value Excel
 * had cached. Merging them into the workbook's names is not the alternative —
 * two sheets are free to define the same name differently, and one would have
 * silently won everywhere.
 */
import { describe, expect, test } from "bun:test";
import { createEvaluator } from "../../formula/evaluate.js";
import type { CellKey, RawCell } from "../../model/types.js";
import { readXlsx } from "./read.js";
import { writeXlsx } from "./write.js";

/** Two sheets, each defining RATE as a different cell, plus a workbook-wide name. */
function twoScopes() {
  return writeXlsx(
    [
      {
        name: "First",
        cells: { "0_0": "10", "1_0": "=RATE*2" } as Record<CellKey, RawCell>,
        styles: {},
        merges: [],
        namedRanges: { RATE: { r1: 0, c1: 0, r2: 0, c2: 0 } },
      },
      {
        name: "Second",
        cells: { "0_0": "7", "1_0": "=RATE*2" } as Record<CellKey, RawCell>,
        styles: {},
        merges: [],
        namedRanges: { RATE: { r1: 0, c1: 0, r2: 0, c2: 0 } },
      },
    ],
    { namedRanges: { TOTAL: { r1: 0, c1: 0, r2: 1, c2: 0 } } },
  );
}

describe("importing", () => {
  test("a sheet-scoped name lands on its own sheet, not the workbook", async () => {
    const sheets = await readXlsx(twoScopes());

    expect(sheets[0]?.namedRanges.RATE).toEqual({ r1: 0, c1: 0, r2: 0, c2: 0 });
    expect(sheets[1]?.namedRanges.RATE).toEqual({ r1: 0, c1: 0, r2: 0, c2: 0 });
    // The workbook keeps only the unscoped one.
    expect(sheets.namedRanges?.RATE).toBeUndefined();
    expect(sheets.namedRanges?.TOTAL).toEqual({ r1: 0, c1: 0, r2: 1, c2: 0 });
  });

  test("each sheet's copy is its own, so neither overwrites the other", async () => {
    const sheets = await readXlsx(twoScopes());

    expect(sheets[0]?.namedRanges).not.toBe(sheets[1]?.namedRanges);
  });

  test("a name whose localSheetId is not a sheet index is dropped, not misfiled", async () => {
    // Attaching it to a guessed sheet is worse than losing it: the formula would
    // read a real number from the wrong place.
    const bytes = writeXlsx([{ name: "Only", cells: {}, styles: {}, merges: [] }]);
    const sheets = await readXlsx(bytes);

    expect(sheets[0]?.namedRanges).toEqual({});
  });
});

describe("evaluating", () => {
  test("a sheet's own name shadows the workbook's of the same name", () => {
    const cells = { "0_0": "10", "5_0": "99", "1_0": "=RATE" } as Record<
      CellKey,
      RawCell
    >;
    const ev = createEvaluator(
      cells,
      { RATE: { r1: 5, c1: 0, r2: 5, c2: 0 } },
      { sheetNamedRanges: { RATE: { r1: 0, c1: 0, r2: 0, c2: 0 } } },
    );

    expect(ev.getCellDisplay(1, 0)).toBe(10);
  });

  test("a workbook name the sheet does not redefine still resolves", () => {
    const cells = { "5_0": "99", "1_0": "=OTHER" } as Record<CellKey, RawCell>;
    const ev = createEvaluator(
      cells,
      { OTHER: { r1: 5, c1: 0, r2: 5, c2: 0 } },
      { sheetNamedRanges: { RATE: { r1: 0, c1: 0, r2: 0, c2: 0 } } },
    );

    expect(ev.getCellDisplay(1, 0)).toBe(99);
  });

  test("a sheet formula name shadows a workbook RANGE of the same name", () => {
    // The kinds are checked in order — ranges first — so shadowing has to remove
    // the workbook's range, not merely add the sheet's formula beside it.
    const cells = { "5_0": "99", "1_0": "=RATE" } as Record<CellKey, RawCell>;
    const ev = createEvaluator(
      cells,
      { RATE: { r1: 5, c1: 0, r2: 5, c2: 0 } },
      { sheetNamedFormulas: { RATE: "1+1" } },
    );

    expect(ev.getCellDisplay(1, 0)).toBe(2);
  });
});

describe("exporting", () => {
  test("sheet-scoped names survive a round-trip on their own sheet", async () => {
    const once = await readXlsx(twoScopes());
    const again = await readXlsx(
      writeXlsx(
        once.map((s) => ({
          name: s.name,
          cells: s.cells,
          styles: s.styles,
          merges: s.merges,
          namedRanges: s.namedRanges,
          namedFormulas: s.namedFormulas,
        })),
        { namedRanges: once.namedRanges ?? {} },
      ),
    );

    expect(again[0]?.namedRanges.RATE).toEqual({ r1: 0, c1: 0, r2: 0, c2: 0 });
    expect(again[1]?.namedRanges.RATE).toEqual({ r1: 0, c1: 0, r2: 0, c2: 0 });
    expect(again.namedRanges?.TOTAL).toEqual({ r1: 0, c1: 0, r2: 1, c2: 0 });
    expect(again.namedRanges?.RATE).toBeUndefined();
  });
});
