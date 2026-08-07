/**
 * Exporting the parts that live outside `<sheetData>`.
 *
 * Reading these without writing them is a quiet data loss: a workbook imported,
 * edited, and exported would come back with its tables flattened to plain cells
 * and its rules and pictures gone. So the assertions here are round trips —
 * write, read back, compare — rather than checks on the XML we happen to emit.
 */
import { describe, expect, test } from "bun:test";
import type {
  CellKey,
  CondFormat,
  RawCell,
  SheetTable,
  StyleObject,
} from "../../model/types.js";
import { readXlsx } from "./read.js";
import { writeXlsx, type XlsxSheetInput } from "./write.js";

const CELLS = {
  "0_0": "team",
  "0_1": "won",
  "1_0": "Arsenal",
  "1_1": "20",
  "2_0": "Chelsea",
  "2_1": "15",
} as Record<CellKey, RawCell>;

function sheet(extra: Partial<XlsxSheetInput> = {}): XlsxSheetInput {
  return { name: "Sheet1", cells: CELLS, styles: {}, merges: [], ...extra };
}

/** Writes one sheet and reads it straight back. */
async function roundTrip(input: XlsxSheetInput) {
  const back = await readXlsx(writeXlsx([input]));
  const first = back[0];
  if (!first) throw new Error("no sheet came back");
  return first;
}

describe("tables", () => {
  const table: SheetTable = {
    name: "tblTeams",
    range: { r1: 0, c1: 0, r2: 2, c2: 1 },
    columns: ["team", "won"],
    headerRow: true,
  };

  test("a table survives with its range and column names", async () => {
    // The names are what a structured reference resolves against, so losing them
    // breaks every formula that used the table even though the cells are intact.
    const back = await roundTrip(sheet({ tables: [table] }));

    expect(back.tables).toEqual([table]);
  });

  test("two tables get distinct ids", async () => {
    // Excel rejects the whole file on a duplicate table id, with an error that
    // names nothing useful.
    const second: SheetTable = {
      ...table,
      name: "tblOther",
      range: { r1: 5, c1: 0, r2: 7, c2: 1 },
    };
    const back = await roundTrip(sheet({ tables: [table, second] }));

    expect(back.tables.map((t) => t.name).sort()).toEqual(["tblOther", "tblTeams"]);
  });

  test("a sheet with no tables writes no table parts", async () => {
    const back = await roundTrip(sheet());
    expect(back.tables).toEqual([]);
  });
});

describe("conditional formats", () => {
  const purple: StyleObject = {
    color: "#ffffff",
    gradient: {
      degree: 90,
      stops: [
        { position: 0, color: "#3d195b" },
        { position: 1, color: "#612890" },
      ],
    },
  };

  test("an expression rule keeps its range, formula, and style", async () => {
    const format: CondFormat = {
      range: { r1: 0, c1: 0, r2: 0, c2: 1 },
      priority: 1,
      rule: { type: "expression", formula: "ISBLANK($A$1)=FALSE" },
      style: purple,
    };
    const back = await roundTrip(sheet({ condFormats: [format] }));

    expect(back.condFormats).toHaveLength(1);
    expect(back.condFormats[0]?.range).toEqual(format.range);
    expect(back.condFormats[0]?.rule).toEqual(format.rule);
    expect(back.condFormats[0]?.style.gradient).toEqual(purple.gradient);
  });

  test("a gradient-only style is not lost", async () => {
    // A dxf that writes only its font and solid fill comes back EMPTY for a
    // gradient rule — and an empty dxf is indistinguishable from no style, so
    // the rule pointing at it is dropped on re-read.
    const format: CondFormat = {
      range: { r1: 0, c1: 0, r2: 0, c2: 0 },
      priority: 1,
      rule: { type: "expression", formula: "TRUE" },
      style: { gradient: purple.gradient as NonNullable<StyleObject["gradient"]> },
    };
    const back = await roundTrip(sheet({ condFormats: [format] }));

    expect(back.condFormats).toHaveLength(1);
  });

  test("a solid fill survives, despite dxf naming it the other way round", async () => {
    // A cellXf's solid fill states its colour as fgColor; a dxf's states it as
    // bgColor. Writing fgColor makes the fill vanish on re-read.
    const format: CondFormat = {
      range: { r1: 0, c1: 0, r2: 0, c2: 0 },
      priority: 1,
      rule: { type: "expression", formula: "TRUE" },
      style: { bg: "#ff0000" },
    };
    const back = await roundTrip(sheet({ condFormats: [format] }));

    expect(back.condFormats[0]?.style.bg).toBe("#ff0000");
  });

  test("a cellIs rule keeps its operator and operands", async () => {
    const format: CondFormat = {
      range: { r1: 0, c1: 1, r2: 2, c2: 1 },
      priority: 2,
      rule: { type: "cellIs", operator: "between", operands: ["10", "30"] },
      style: { bold: true },
      stopIfTrue: true,
    };
    const back = await roundTrip(sheet({ condFormats: [format] }));

    expect(back.condFormats[0]?.rule).toEqual(format.rule);
    expect(back.condFormats[0]?.stopIfTrue).toBe(true);
    expect(back.condFormats[0]?.priority).toBe(2);
  });

  test("several rules keep their own styles rather than sharing one", async () => {
    const rules: CondFormat[] = [
      {
        range: { r1: 0, c1: 0, r2: 0, c2: 0 },
        priority: 1,
        rule: { type: "expression", formula: "TRUE" },
        style: { bg: "#ff0000" },
      },
      {
        range: { r1: 1, c1: 0, r2: 1, c2: 0 },
        priority: 2,
        rule: { type: "expression", formula: "TRUE" },
        style: { bg: "#00ff00" },
      },
    ];
    const back = await roundTrip(sheet({ condFormats: rules }));

    expect(back.condFormats.map((f) => f.style.bg)).toEqual(["#ff0000", "#00ff00"]);
  });
});

describe("in-cell images", () => {
  /** A one-pixel PNG, as a data URI. */
  const PNG =
    "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

  test("an embedded picture survives byte for byte", async () => {
    const back = await roundTrip(
      sheet({
        cells: { ...CELLS, "1_2": '=IMAGE("https://example.test/a.png")' },
        images: {
          "1_2": { src: PNG, alt: "https://example.test/a.png" },
        } as Record<CellKey, { src: string; alt?: string }>,
      }),
    );

    expect(back.images["1_2" as CellKey]?.src).toBe(PNG);
    expect(back.images["1_2" as CellKey]?.alt).toBe("https://example.test/a.png");
  });

  test("a URL-only picture keeps its URL", async () => {
    const back = await roundTrip(
      sheet({
        cells: { ...CELLS, "1_2": '=IMAGE("https://example.test/b.png")' },
        images: {
          "1_2": { src: "https://example.test/b.png" },
        } as Record<CellKey, { src: string }>,
      }),
    );

    expect(back.images["1_2" as CellKey]?.src).toBe("https://example.test/b.png");
  });

  test("one picture used by many cells is stored once", async () => {
    // A workbook typically points many cells at the same image — a hundred and
    // forty at twenty in the sample file — and a media part each multiplies the
    // file size several times over.
    const images = Object.fromEntries(
      Array.from({ length: 20 }, (_, i) => [`${i}_5`, { src: PNG }]),
    ) as Record<CellKey, { src: string }>;
    const shared = writeXlsx([sheet({ images })]);
    const single = writeXlsx([
      sheet({
        images: { "0_5": { src: PNG } } as Record<CellKey, { src: string }>,
      }),
    ]);

    // Twenty references cost barely more than one, rather than twenty times more.
    expect(shared.length).toBeLessThan(single.length * 2);

    const back = await readXlsx(shared);
    expect(Object.keys(back[0]?.images ?? {})).toHaveLength(20);
  });

  test("pictures on different sheets do not swap", async () => {
    // A cell key is unique only within a sheet. Indexing the value metadata by it
    // alone gave every sheet the first sheet's pictures.
    const other =
      "data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7";
    const bytes = writeXlsx([
      sheet({
        name: "One",
        images: { "0_0": { src: PNG } } as Record<CellKey, { src: string }>,
      }),
      sheet({
        name: "Two",
        images: { "0_0": { src: other } } as Record<CellKey, { src: string }>,
      }),
    ]);
    const back = await readXlsx(bytes);

    expect(back[0]?.images["0_0" as CellKey]?.src).toBe(PNG);
    expect(back[1]?.images["0_0" as CellKey]?.src).toBe(other);
  });

  test("a cell holding only a picture is still written", async () => {
    const back = await roundTrip(
      sheet({
        images: { "9_9": { src: PNG } } as Record<CellKey, { src: string }>,
      }),
    );
    expect(back.images["9_9" as CellKey]?.src).toBe(PNG);
  });
});
