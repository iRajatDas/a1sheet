/**
 * How the grid treats a formula it cannot evaluate but the file already had a
 * value for, and where the monospace font is allowed to appear.
 */
import { describe, expect, test } from "bun:test";
import { fireEvent, render, screen } from "@testing-library/react";
import { makeSheet } from "../model/sheet.js";
import type { CellKey, Sheet, Workbook } from "../model/types.js";
import { Spreadsheet } from "./Spreadsheet.js";
import { buildCss } from "./styles.js";
import { resolveTheme } from "./theme.js";

/** A sheet as an import leaves it: a formula we cannot parse, plus its value. */
function importedSheet(): Workbook {
  const sheet: Sheet = {
    ...makeSheet("Imported"),
    cells: {
      "0_0": "=SUM(A2:A3)",
      "0_1": "=NOSUCHFUNC(A1)",
      "1_0": "2",
      "2_0": "3",
    } as Record<CellKey, string>,
    cachedValues: { "0_1": 41 } as Record<CellKey, number>,
  };
  return { sheets: [sheet], activeSheetIndex: 0, namedRanges: {} };
}

function cellText(row: number, col: number): string {
  const el = document.querySelector(`[data-row="${row}"][data-col="${col}"]`);
  if (!el) throw new Error(`no cell at ${row},${col}`);
  return el.textContent ?? "";
}

describe("a formula the engine does not implement", () => {
  test("displays the value the file was imported with", () => {
    render(<Spreadsheet defaultWorkbook={importedSheet()} />);
    expect(cellText(0, 1)).toBe("41");
  });

  test("stops standing in once the user replaces the formula", () => {
    // The imported value describes the formula that WAS there. Keeping it would
    // show a number with no relationship to the cell's contents.
    render(<Spreadsheet defaultWorkbook={importedSheet()} />);

    const cell = document.querySelector('[data-row="0"][data-col="1"]');
    if (!cell) throw new Error("no cell");
    fireEvent.mouseDown(cell);
    fireEvent.doubleClick(cell);
    const editor = screen.getByRole("textbox", { name: /formula/i });
    fireEvent.change(editor, { target: { value: "=OTHERBADFUNC(1)" } });
    fireEvent.keyDown(editor, { key: "Enter" });

    expect(cellText(0, 1)).toBe("#NAME?");
  });

  test("does not shadow a formula the engine can evaluate", () => {
    render(<Spreadsheet defaultWorkbook={importedSheet()} />);
    expect(cellText(0, 0)).toBe("5");
  });
});

describe("a cell holding an IMAGE formula", () => {
  const withImage = (): Workbook => {
    const sheet: Sheet = {
      ...makeSheet("Pictures"),
      cells: { "0_0": '=IMAGE("https://example.test/crest.png")' } as Record<
        CellKey,
        string
      >,
      images: {
        "0_0": { src: "data:image/png;base64,AAAA", alt: "a crest" },
      } as Record<CellKey, { src: string; alt?: string }>,
    };
    return { sheets: [sheet], activeSheetIndex: 0, namedRanges: {} };
  };

  test("draws the picture instead of the URL", () => {
    render(<Spreadsheet defaultWorkbook={withImage()} />);

    const img = document.querySelector('[data-row="0"][data-col="0"] img');
    expect(img?.getAttribute("src")).toBe("data:image/png;base64,AAAA");
    // The value is still a URL, and showing it beside the image would be noise.
    expect(cellText(0, 0)).toBe("");
  });

  test("the source URL is the alternative text", () => {
    render(<Spreadsheet defaultWorkbook={withImage()} />);
    const img = document.querySelector('[data-row="0"][data-col="0"] img');
    expect(img?.getAttribute("alt")).toBe("a crest");
  });

  test("replacing the formula removes the picture", () => {
    // Otherwise the image outlives the formula that asked for it.
    render(<Spreadsheet defaultWorkbook={withImage()} />);

    const cell = document.querySelector('[data-row="0"][data-col="0"]');
    if (!cell) throw new Error("no cell");
    fireEvent.mouseDown(cell);
    fireEvent.doubleClick(cell);
    const editor = screen.getByRole("textbox", { name: /formula/i });
    fireEvent.change(editor, { target: { value: "just text" } });
    fireEvent.keyDown(editor, { key: "Enter" });

    expect(document.querySelector('[data-row="0"][data-col="0"] img')).toBeNull();
    expect(cellText(0, 0)).toBe("just text");
  });

  test("typing an IMAGE formula draws one", () => {
    render(<Spreadsheet />);

    const cell = document.querySelector('[data-row="0"][data-col="0"]');
    if (!cell) throw new Error("no cell");
    fireEvent.mouseDown(cell);
    fireEvent.doubleClick(cell);
    const editor = screen.getByRole("textbox", { name: /formula/i });
    fireEvent.change(editor, {
      target: { value: '=IMAGE("https://example.test/x.png")' },
    });
    fireEvent.keyDown(editor, { key: "Enter" });

    const img = document.querySelector('[data-row="0"][data-col="0"] img');
    expect(img?.getAttribute("src")).toBe("https://example.test/x.png");
  });
});

describe("the stylesheet", () => {
  test("sets cells in the body font, not the monospace one", () => {
    // A grid of names and totals in a monospace face reads as a terminal. The
    // mono font is for formula editing, where column alignment carries meaning.
    const theme = resolveTheme({
      fontFamily: "BodyFace",
      monoFontFamily: "MonoFace",
    });
    const css = buildCss("a1s-", theme);

    const cellRule = css.slice(css.indexOf(".a1s-cell {"));
    expect(cellRule.slice(0, cellRule.indexOf("}"))).toContain(
      "font-family: BodyFace",
    );
    expect(css).not.toContain("font-family: MonoFace");
  });
});
