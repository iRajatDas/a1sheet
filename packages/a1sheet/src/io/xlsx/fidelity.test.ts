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
import { condStyleFor } from "../../format/condFormat.js";
import { formatValue } from "../../format/numFmt.js";
import { createEvaluator } from "../../formula/evaluate.js";
import { imageUrlIn } from "../../formula/imageCall.js";
import { tableIndex } from "../../formula/tableRefs.js";
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
    const evaluator = createEvaluator(
      data.cells,
      {},
      { cachedValues: data.cachedValues },
    );

    // Data!E2 is IF(COUNTBLANK(tblMatches[[#This Row],...])=0,1,0). Structured
    // references are not implemented, so evaluation fails and the import answers.
    expect(data.cells["1_4" as CellKey]).toContain("tblMatches[");
    expect(evaluator.getCellDisplay(1, 4)).toBe(1);
  });

  test("a LET/LAMBDA defined name shows its value rather than an error", async () => {
    const sheets = await readXlsx(bytes);
    const exampleA = sheets[1];
    if (!exampleA) throw new Error("no second sheet");
    const evaluator = createEvaluator(
      exampleA.cells,
      {},
      { cachedValues: exampleA.cachedValues },
    );

    // Example A!C4 is "=OverallTable", a defined name wrapping LET and LAMBDA.
    expect(exampleA.cells["3_2" as CellKey]).toBe("=OverallTable");
    expect(evaluator.getCellDisplay(3, 2)).toBe("POS");
  });

  test("a table's header takes its colour from the theme accent", async () => {
    // Data!A1 is the header of tblMatches, styled TableStyleMedium4. That resolves
    // to accent3 — a green that appears nowhere in the file, since the style is a
    // built-in Excel expects every reader to know.
    const [data] = await readXlsx(bytes);
    const header = data?.styles["0_0" as CellKey];

    expect(header?.bg).toBe("#196b24");
    expect(header?.bold).toBe(true);
    // The dxf that styles this header sets `<u val="none"/>`, which must not read
    // as "underlined".
    expect(header?.underline).toBeUndefined();
  });

  test("alternate table rows are banded", async () => {
    const [data] = await readXlsx(bytes);

    expect(data?.styles["1_0" as CellKey]?.bg).toBeUndefined();
    expect(data?.styles["2_0" as CellKey]?.bg).toBeDefined();
  });

  test("a heading styled by a conditional format is not left plain", async () => {
    // Example A's purple title bars are a `type="expression"` rule, not cell
    // formatting. A reader that skips conditional formatting shows them unstyled.
    const sheets = await readXlsx(bytes);
    const exampleA = sheets[1];
    if (!exampleA) throw new Error("no second sheet");
    const evaluator = createEvaluator(
      exampleA.cells,
      {},
      { cachedValues: exampleA.cachedValues },
    );

    const style = condStyleFor(
      { condFormats: exampleA.condFormats, evaluator },
      3,
      2,
    );
    expect(style?.gradient?.stops.map((s) => s.color)).toEqual([
      "#3d195b",
      "#612890",
    ]);
  });

  test("fonts, sizes, and borders arrive with the cells", async () => {
    const sheets = await readXlsx(bytes);
    const exampleA = sheets[1];
    const title = exampleA?.styles["1_4" as CellKey];

    expect(title?.fontFamily).toBe("Aptos Narrow");
    expect(title?.fontSize).toBe(20);
    // The purple rule under the "Overall" title, which is a border.
    expect(title?.borders?.bottom).toEqual({ line: "thick", color: "#7030a0" });
  });

  test("a goal difference keeps the sign its format code adds", async () => {
    const sheets = await readXlsx(bytes);
    const exampleA = sheets[1];
    if (!exampleA) throw new Error("no second sheet");
    const evaluator = createEvaluator(
      exampleA.cells,
      {},
      { cachedValues: exampleA.cachedValues },
    );

    // Column L of the league table, formatted "\+0;\-0;0". Bucketed as an
    // integer it rendered 45 and dropped the plus.
    const value = evaluator.getCellDisplay(4, 11);
    expect(formatValue(value as number, exampleA.styles["4_11" as CellKey])).toBe(
      "+45",
    );
  });

  test("a date column shows its time as well as its date", async () => {
    const [data] = await readXlsx(bytes);
    if (!data) throw new Error("no sheets");
    const evaluator = createEvaluator(
      data.cells,
      {},
      { cachedValues: data.cachedValues },
    );

    // numFmtId 22, whose code the file never states.
    expect(
      formatValue(
        evaluator.getCellDisplay(1, 6) as number,
        data.styles["1_6" as CellKey],
      ),
    ).toBe("8/16/24 20:00");
  });

  test("an IMAGE cell yields a picture rather than #VALUE!", async () => {
    // Data!J2 is `=IMAGE("…Arsenal_FC…png")`, which Excel stores as the error
    // value #VALUE! plus a vm pointing at an embedded PNG.
    const [data] = await readXlsx(bytes);
    if (!data) throw new Error("no sheets");

    expect(data.images["1_9" as CellKey]?.src).toStartWith(
      "data:image/png;base64,",
    );
    // …and the formula still evaluates to something meaningful, so a cell whose
    // image cannot be drawn shows where it came from.
    const evaluator = createEvaluator(
      data.cells,
      {},
      { cachedValues: data.cachedValues },
    );
    expect(evaluator.getCellDisplay(1, 9)).toContain("Arsenal_FC");
  });

  test("Excel's modern-function prefix is stripped so the call is recognizable", () => {
    // `_xlfn.IMAGE` and `_xlfn._xlws.SORT` are how Excel writes functions newer
    // than the format. Left in place they guarantee #NAME? even where we have
    // the function.
    expect(imageUrlIn('=IMAGE("x.png")')).toBe("x.png");
  });

  test("a dynamic-array workbook recalculates instead of only displaying", async () => {
    // The whole league table is one formula: a workbook-level defined name
    // wrapping LET, LAMBDA, MAP, SORT, UNIQUE, VSTACK, HSTACK, XLOOKUP,
    // SEQUENCE, and CHOOSECOLS, reading a table on another sheet through
    // structured references, and spilling over 21 rows.
    //
    // Evaluated with NO cached values, so nothing here can come from what Excel
    // stored — this is the engine computing the table from the match results.
    const result = await readXlsx(bytes);
    const exampleA = result[1];
    if (!exampleA) throw new Error("no second sheet");

    const evaluator = createEvaluator(exampleA.cells, result.namedRanges ?? {}, {
      tables: tableIndex(
        result.flatMap((s) => s.tables.map((t) => ({ ...t, sheet: s.name }))),
      ),
      ...(result.namedFormulas ? { namedFormulas: result.namedFormulas } : {}),
      spillRanges: exampleA.spillRanges,
      sheets: result.map((s) => ({ name: s.name, cells: s.cells })),
    });

    // The anchor, C4, holds `=OverallTable`.
    expect(evaluator.getCellDisplay(3, 2)).toBe("POS");
    // …and the rest of the table is its spill, not stored values.
    expect(evaluator.getCellDisplay(4, 4)).toBe("Liverpool");
    expect(evaluator.getCellDisplay(4, 12)).toBe(84);
    expect(evaluator.getCellDisplay(5, 4)).toBe("Arsenal");
  });

  test("an array formula's own output does not block its spill", async () => {
    // Excel writes a dynamic array's result into the sheet as ordinary values so
    // other readers can see it. Without the `<f t="array" ref>` region those
    // cells look like content in the way, and every such formula reports #SPILL!
    // against its own output.
    const result = await readXlsx(bytes);
    expect(Object.keys(result[1]?.spillRanges ?? {}).length).toBeGreaterThan(0);
  });

  test("a whole real workbook survives an export and re-import", async () => {
    // The end-to-end version of everything above. Reading these parts without
    // writing them is a quiet data loss: a workbook imported, edited, and
    // exported would come back with its tables flattened to plain cells and its
    // rules and pictures gone.
    const src = await readXlsx(bytes);
    const written = writeXlsx(
      src.map((s) => ({
        name: s.name,
        cells: s.cells,
        styles: s.styles,
        cachedValues: s.cachedValues,
        merges: s.merges,
        colWidths: s.colWidths,
        rowHeights: s.rowHeights,
        tables: s.tables,
        condFormats: s.condFormats,
        images: s.images,
      })),
    );
    const back = await readXlsx(written);

    for (const [i, sheet] of back.entries()) {
      const before = src[i];
      if (!before) throw new Error("sheet count changed");
      expect(sheet.name).toBe(before.name);
      expect(sheet.tables).toEqual(before.tables);
      expect(sheet.condFormats).toHaveLength(before.condFormats.length);
      expect(Object.keys(sheet.images)).toHaveLength(
        Object.keys(before.images).length,
      );
      expect(Object.keys(sheet.styles)).toHaveLength(
        Object.keys(before.styles).length,
      );
      expect(sheet.merges).toEqual(before.merges);
    }

    // The crest is the same picture, not a re-encoded or swapped one.
    expect(back[0]?.images["1_9" as CellKey]?.src).toBe(
      src[0]?.images["1_9" as CellKey]?.src,
    );
    // …and the purple gradient a conditional format carries came back whole.
    expect(back[1]?.condFormats[0]?.style.gradient).toEqual(
      src[1]?.condFormats[0]?.style.gradient,
    );
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
    const evaluator = createEvaluator(
      cells,
      {},
      {
        cachedValues: {
          "1_0": 999,
        } as Record<CellKey, number>,
      },
    );
    expect(evaluator.getCellDisplay(1, 0)).toBe(2);
  });

  test("a formula it cannot evaluate falls back", () => {
    const evaluator = createEvaluator(
      { "0_0": "=NOSUCHFUNC(1)" } as Record<CellKey, string>,
      {},
      { cachedValues: { "0_0": 42 } as Record<CellKey, number> },
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
