/**
 * Import fidelity: the two ways a file used to arrive silently wrong.
 *
 * Dates landed seventy years out, because a serial is just a number and the two
 * formats count from different epochs. Formulas using anything this engine does
 * not implement landed as `#NAME?`, discarding the value Excel had already
 * computed and stored right next to them — so a workbook built on dynamic arrays
 * or structured references imported as a wall of errors.
 *
 * Both are asserted against a real Excel file, not a fixture we wrote: a
 * self-round-trip would have agreed with itself in either direction.
 */
import { describe, expect, test } from "bun:test";
import { join } from "node:path";
import { createEvaluator } from "../../formula/evaluate.js";
import type { CellKey } from "../../model/types.js";
import { daySerialToExcelSerial, excelSerialToDaySerial } from "./dates.js";
import { readXlsx } from "./read.js";
import { writeXlsx } from "./write.js";

const FIXTURE = join(
  import.meta.dir,
  "..",
  "..",
  "..",
  "..",
  "..",
  "dataset",
  "excel-files",
  "building-an-excel-football-league-table-5-years-on.xlsx",
);

const fixture = await Bun.file(FIXTURE)
  .arrayBuffer()
  .then((b) => new Uint8Array(b))
  .catch(() => undefined);

describe("Excel's day-serial epoch", () => {
  test("a serial reads back as the date Excel shows", () => {
    // 45520 is 2024-08-16 in Excel. Read as a Unix day serial it is 2094-08-18,
    // which is why this failure was invisible: a plausible date, wrong century.
    const day = excelSerialToDaySerial(45520);
    expect(new Date(day * 86400000).toISOString().slice(0, 10)).toBe("2024-08-16");
  });

  test("the phantom 1900 leap day does not shift the dates around it", () => {
    // Excel believes 1900-02-29 existed, so the offset differs either side of it.
    const iso = (serial: number) =>
      new Date(excelSerialToDaySerial(serial) * 86400000)
        .toISOString()
        .slice(0, 10);
    expect(iso(1)).toBe("1900-01-01");
    expect(iso(59)).toBe("1900-02-28");
    expect(iso(61)).toBe("1900-03-01");
  });

  test("every real serial survives a round trip through our epoch", () => {
    for (const serial of [1, 59, 61, 366, 25569, 45520, 60000]) {
      expect(daySerialToExcelSerial(excelSerialToDaySerial(serial))).toBe(serial);
    }
  });

  test("the phantom day collapses onto the real one it precedes", () => {
    // 60 is 1900-02-29, which never happened, so it cannot round trip. It reads
    // as 1900-02-28 and writes back as 59 — the serial that day really has.
    expect(excelSerialToDaySerial(60)).toBe(excelSerialToDaySerial(59));
    expect(daySerialToExcelSerial(excelSerialToDaySerial(60))).toBe(59);
  });
});

describe.skipIf(!fixture)("a workbook Excel wrote", () => {
  const bytes = fixture as Uint8Array;

  test("a date column imports as the dates it displays in Excel", async () => {
    const [data] = await readXlsx(bytes);
    if (!data) throw new Error("no sheets");

    // Data!G2, the first date_time value. Excel shows 16/08/2024 20:00.
    const serial = Number(data.cells["1_6" as CellKey]);
    expect(data.styles["1_6" as CellKey]?.numFmt).toBe("date");
    expect(new Date(serial * 86400000).toISOString().slice(0, 10)).toBe(
      "2024-08-16",
    );
  });

  test("a formula we cannot evaluate shows Excel's value, not #NAME?", async () => {
    const [data] = await readXlsx(bytes);
    if (!data) throw new Error("no sheets");
    const evaluator = createEvaluator(data.cells, {}, data.cachedValues);

    // Data!E2 is IF(COUNTBLANK(tblMatches[[#This Row],...])=0,1,0). Structured
    // references are not implemented, so evaluation fails and the import answers.
    expect(data.cells["1_4" as CellKey]).toContain("tblMatches[");
    expect(evaluator.getCellDisplay(1, 4)).toBe(1);
  });

  test("a LET/LAMBDA defined name shows its value rather than an error", async () => {
    const sheets = await readXlsx(bytes);
    const exampleA = sheets[1];
    if (!exampleA) throw new Error("no second sheet");
    const evaluator = createEvaluator(exampleA.cells, {}, exampleA.cachedValues);

    // Example A!C4 is "=OverallTable", a defined name wrapping LET and LAMBDA.
    expect(exampleA.cells["3_2" as CellKey]).toBe("=OverallTable");
    expect(evaluator.getCellDisplay(3, 2)).toBe("POS");
  });

  test("dates survive the trip back out to a file", async () => {
    const [data] = await readXlsx(bytes);
    if (!data) throw new Error("no sheets");
    const written = writeXlsx([
      {
        name: data.name,
        cells: data.cells,
        styles: data.styles,
        cachedValues: data.cachedValues,
        merges: data.merges,
      },
    ]);
    const [reread] = await readXlsx(written);

    expect(reread?.cells["1_6" as CellKey]).toBe(data.cells["1_6" as CellKey]);
  });
});

describe("a cached value is a fallback, never an override", () => {
  const cells = { "0_0": "1", "1_0": "=A1+1" } as Record<CellKey, string>;

  test("a formula this engine can evaluate wins over the import", () => {
    // Otherwise an imported sheet would be frozen: edits would recompute a
    // result that nothing ever displays.
    const evaluator = createEvaluator(cells, {}, {
      "1_0": 999,
    } as Record<CellKey, number>);
    expect(evaluator.getCellDisplay(1, 0)).toBe(2);
  });

  test("a formula it cannot evaluate falls back", () => {
    const evaluator = createEvaluator(
      { "0_0": "=NOSUCHFUNC(1)" } as Record<CellKey, string>,
      {},
      { "0_0": 42 } as Record<CellKey, number>,
    );
    expect(evaluator.getCellDisplay(0, 0)).toBe(42);
  });

  test("without a fallback the error still surfaces", () => {
    const evaluator = createEvaluator(
      { "0_0": "=NOSUCHFUNC(1)" } as Record<CellKey, string>,
      {},
    );
    expect(evaluator.getCellDisplay(0, 0)).toBe("#NAME?");
  });
});
