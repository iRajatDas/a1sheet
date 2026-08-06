/**
 * XLSX round-trip and real-file tests.
 *
 * The self-round-trip cases prove writer and reader agree. The `dataset/` cases
 * prove the reader survives files produced by Excel and LibreOffice, which is the
 * part a hand-written OOXML implementation is most likely to get wrong.
 */
import { describe, expect, test } from "bun:test";
import { readdir } from "node:fs/promises";
import { join } from "node:path";
import type { CellKey, RawCell, StyleObject } from "../../model/types.js";
import { readXlsx } from "./read.js";
import { writeXlsx, type XlsxSheetInput } from "./write.js";

function sheet(
  name: string,
  cells: Record<string, string>,
  styles: Record<string, StyleObject> = {},
): XlsxSheetInput {
  return {
    name,
    cells: cells as Record<CellKey, RawCell>,
    styles: styles as Record<CellKey, StyleObject>,
    merges: [],
  };
}

const DATASET = join(import.meta.dir, "..", "..", "..", "..", "..", "dataset");

/**
 * Read at module scope, not inside the describe callback: bun collects tests
 * synchronously, so an `async` describe body has not resolved yet when
 * `test.skipIf` is evaluated — every case would silently skip.
 */
const DATASET_FILES: string[] = await readdir(join(DATASET, "excel-files")).catch(
  () => [],
);

describe("self round-trip", () => {
  test("literal values survive", async () => {
    const bytes = writeXlsx([
      sheet("S", { "0_0": "hello", "1_0": "42", "2_0": "3.14" }),
    ]);
    const [out] = await readXlsx(bytes);
    expect(out?.cells["0_0"]).toBe("hello");
    expect(out?.cells["1_0"]).toBe("42");
    expect(out?.cells["2_0"]).toBe("3.14");
  });

  test("sheet names survive", async () => {
    const bytes = writeXlsx([sheet("Budget", {}), sheet("Notes", {})]);
    const out = await readXlsx(bytes);
    expect(out.map((s) => s.name)).toEqual(["Budget", "Notes"]);
  });

  test("formulas survive as formulas", async () => {
    const bytes = writeXlsx([
      sheet("S", { "0_0": "2", "1_0": "3", "2_0": "=SUM(A1:A2)" }),
    ]);
    const [out] = await readXlsx(bytes);
    expect(out?.cells["2_0"]).toBe("=SUM(A1:A2)");
  });

  test("styles survive", async () => {
    const bytes = writeXlsx([
      sheet(
        "S",
        { "0_0": "styled" },
        {
          "0_0": {
            bold: true,
            italic: true,
            color: "#ff0000",
            bg: "#00ff00",
            align: "center",
          },
        },
      ),
    ]);
    const [out] = await readXlsx(bytes);
    const s = out?.styles["0_0"];
    expect(s?.bold).toBe(true);
    expect(s?.italic).toBe(true);
    expect(s?.color).toBe("#ff0000");
    expect(s?.bg).toBe("#00ff00");
    expect(s?.align).toBe("center");
  });

  test("a styled but empty cell keeps its formatting", async () => {
    const bytes = writeXlsx([sheet("S", {}, { "5_5": { bold: true } })]);
    const [out] = await readXlsx(bytes);
    expect(out?.styles["5_5"]?.bold).toBe(true);
  });

  test("number formats survive", async () => {
    const bytes = writeXlsx([
      sheet(
        "S",
        { "0_0": "0.5", "0_1": "1234.5" },
        { "0_0": { numFmt: "percent" }, "0_1": { numFmt: "currency" } },
      ),
    ]);
    const [out] = await readXlsx(bytes);
    expect(out?.styles["0_0"]?.numFmt).toBe("percent");
    expect(out?.styles["0_1"]?.numFmt).toBe("currency");
  });

  test("merges survive", async () => {
    const bytes = writeXlsx([
      {
        ...sheet("S", { "0_0": "wide" }),
        merges: [{ r1: 0, c1: 0, r2: 2, c2: 3 }],
      },
    ]);
    const [out] = await readXlsx(bytes);
    expect(out?.merges).toEqual([{ r1: 0, c1: 0, r2: 2, c2: 3 }]);
  });

  test("XML-hostile text is escaped and recovered", async () => {
    const nasty = `a<b & c>"d" 'e' é中`;
    const bytes = writeXlsx([sheet("S", { "0_0": nasty })]);
    const [out] = await readXlsx(bytes);
    expect(out?.cells["0_0"]).toBe(nasty);
  });

  test("a sheet name needing escaping survives", async () => {
    const bytes = writeXlsx([sheet(`R&D <2024>`, { "0_0": "x" })]);
    const [out] = await readXlsx(bytes);
    expect(out?.name).toBe(`R&D <2024>`);
  });

  test("reported extent covers the used range", async () => {
    const bytes = writeXlsx([sheet("S", { "9_4": "corner" })]);
    const [out] = await readXlsx(bytes);
    expect(out?.rows).toBe(10);
    expect(out?.cols).toBe(5);
  });

  test("an empty workbook reads back as one empty sheet", async () => {
    const bytes = writeXlsx([sheet("Empty", {})]);
    const out = await readXlsx(bytes);
    expect(out).toHaveLength(1);
    expect(Object.keys(out[0]?.cells ?? {})).toHaveLength(0);
  });

  test("a large sheet round-trips intact", async () => {
    const cells: Record<string, string> = {};
    for (let r = 0; r < 200; r++) {
      for (let c = 0; c < 10; c++) cells[`${r}_${c}`] = `${r}-${c}`;
    }
    const bytes = writeXlsx([sheet("Big", cells)]);
    const [out] = await readXlsx(bytes);
    expect(Object.keys(out?.cells ?? {})).toHaveLength(2000);
    expect(out?.cells["199_9"]).toBe("199-9");
  });
});

describe("rejects unsupported formats clearly", () => {
  test("a ZIP that is not a spreadsheet is rejected by name, not by crash", async () => {
    // Hand-build a ZIP with no workbook part.
    const bytes = writeXlsx([sheet("S", { "0_0": "x" })]);
    // Corrupt the workbook part name so the guard trips.
    const text = new TextDecoder("latin1").decode(bytes);
    const broken = text.replaceAll("xl/workbook.xml", "xl/notabook.xml");
    const brokenBytes = Uint8Array.from(broken, (ch) => ch.charCodeAt(0));
    // Assert on the code, not the message: the code is the stable contract.
    await expect(readXlsx(brokenBytes)).rejects.toMatchObject({
      code: "MALFORMED_FILE",
    });
  });

  test("non-zip input is rejected", async () => {
    await expect(
      readXlsx(new TextEncoder().encode("just text")),
    ).rejects.toMatchObject({ code: "NOT_A_ZIP" });
  });
});

/**
 * Real-world corpus. Skipped when dataset/ is absent so the suite still passes on
 * a fresh clone.
 */
describe("dataset/ real files", () => {
  const xlsx = DATASET_FILES.filter((f) => /\.xlsx$/i.test(f));
  const xlsm = DATASET_FILES.filter((f) => /\.xlsm$/i.test(f));
  const binary = DATASET_FILES.filter((f) => /\.(xlsb|xls)$/i.test(f));

  test.skipIf(xlsx.length === 0)("reads every .xlsx without throwing", async () => {
    for (const name of xlsx) {
      const bytes = new Uint8Array(
        await Bun.file(join(DATASET, "excel-files", name)).arrayBuffer(),
      );
      const sheets = await readXlsx(bytes);
      expect(sheets.length).toBeGreaterThan(0);
      // Every sheet must have a name and a usable extent.
      for (const s of sheets) {
        expect(typeof s.name).toBe("string");
        expect(s.rows).toBeGreaterThan(0);
      }
    }
  });

  test.skipIf(xlsx.length === 0)(
    "recovers actual cell content from a real file",
    async () => {
      let totalCells = 0;
      for (const name of xlsx) {
        const bytes = new Uint8Array(
          await Bun.file(join(DATASET, "excel-files", name)).arrayBuffer(),
        );
        for (const s of await readXlsx(bytes)) {
          totalCells += Object.keys(s.cells).length;
        }
      }
      // If inflate or sharedStrings were broken this would be 0.
      expect(totalCells).toBeGreaterThan(100);
    },
  );

  test.skipIf(xlsm.length === 0)(
    "reads .xlsm, which is the same container",
    async () => {
      for (const name of xlsm) {
        const bytes = new Uint8Array(
          await Bun.file(join(DATASET, "excel-files", name)).arrayBuffer(),
        );
        const sheets = await readXlsx(bytes);
        expect(sheets.length).toBeGreaterThan(0);
      }
    },
  );

  test.skipIf(binary.length === 0)(
    "rejects .xlsb and .xls with an actionable message",
    async () => {
      for (const name of binary) {
        const bytes = new Uint8Array(
          await Bun.file(join(DATASET, "excel-files", name)).arrayBuffer(),
        );
        // Either a ZIP holding binary parts, or not a ZIP at all. Both must throw
        // rather than silently returning an empty sheet.
        await expect(readXlsx(bytes)).rejects.toThrow();
      }
    },
  );

  test.skipIf(xlsx.length === 0)(
    "a real file survives a read/write/read cycle",
    async () => {
      const name = xlsx[0] as string;
      const bytes = new Uint8Array(
        await Bun.file(join(DATASET, "excel-files", name)).arrayBuffer(),
      );
      const first = await readXlsx(bytes);
      const rewritten = writeXlsx(
        first.map((s) => ({
          name: s.name,
          cells: s.cells,
          styles: s.styles,
          merges: s.merges,
        })),
      );
      const second = await readXlsx(rewritten);
      expect(second.map((s) => s.name)).toEqual(first.map((s) => s.name));
      expect(Object.keys(second[0]?.cells ?? {}).length).toBe(
        Object.keys(first[0]?.cells ?? {}).length,
      );
    },
  );
});

describe("row heights and column widths", () => {
  test("a resized row and column survive a round trip", async () => {
    const bytes = writeXlsx([
      {
        ...sheet("S", { "0_0": "a", "5_2": "b" }),
        colWidths: { 0: 180, 2: 60 },
        rowHeights: { 0: 40, 5: 80 },
      },
    ]);
    const [out] = await readXlsx(bytes);

    // The units are characters and points, so a pixel of rounding is expected.
    expect(out?.colWidths[0]).toBeCloseTo(180, -1);
    expect(out?.colWidths[2]).toBeCloseTo(60, -1);
    expect(out?.rowHeights[0]).toBeCloseTo(40, -1);
    expect(out?.rowHeights[5]).toBeCloseTo(80, -1);
  });

  test("untouched rows and columns carry no override", async () => {
    // Writing a default for every row and column would pin the whole sheet to
    // sizes the user never chose, and defeat the fast path in useRowWindow.
    const bytes = writeXlsx([
      { ...sheet("S", { "0_0": "a" }), colWidths: { 0: 180 } },
    ]);
    const [out] = await readXlsx(bytes);

    expect(out?.colWidths[1]).toBeUndefined();
    expect(Object.keys(out?.rowHeights ?? {})).toHaveLength(0);
  });

  test("a resized row with no content is still written", async () => {
    const bytes = writeXlsx([
      { ...sheet("S", { "0_0": "a" }), rowHeights: { 9: 55 } },
    ]);
    const [out] = await readXlsx(bytes);

    expect(out?.rowHeights[9]).toBeCloseTo(55, -1);
  });

  test("a sheet with no sizing writes no <cols> element", async () => {
    const bytes = writeXlsx([sheet("S", { "0_0": "a" })]);
    const [out] = await readXlsx(bytes);

    expect(Object.keys(out?.colWidths ?? {})).toHaveLength(0);
  });

  test.skipIf(DATASET_FILES.length === 0)(
    "a real Excel file's column widths come through as plausible pixels",
    async () => {
      // Self-round-trip only proves the writer and reader agree with each
      // other. Excel emits <col> in runs, mixes custom and default columns, and
      // uses a character-based unit — none of which our own output exercises.
      const { readFile } = await import("node:fs/promises");
      let checked = 0;
      for (const file of DATASET_FILES.filter((f) => f.endsWith(".xlsx"))) {
        const bytes = new Uint8Array(
          await readFile(join(DATASET, "excel-files", file)),
        );
        const sheets = await readXlsx(bytes).catch(() => []);
        for (const s of sheets) {
          for (const px of Object.values(s.colWidths)) {
            // A width outside this range means the unit conversion is wrong,
            // not that someone chose an unusual column.
            expect(px).toBeGreaterThan(4);
            expect(px).toBeLessThan(2000);
            checked++;
          }
        }
      }
      expect(checked).toBeGreaterThan(0);
    },
  );
});
