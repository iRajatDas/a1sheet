/**
 * Format-detecting entry point for reading a spreadsheet file.
 *
 * Dispatch is by ZIP magic first and file extension second, so a .xlsx that has
 * been renamed still reads correctly.
 */

import type { Range } from "../model/types.js";
import { csvToCells } from "./csv/read.js";
import type { AsyncReadOptions } from "./progress.js";
import type { XlsxSheetData } from "./xlsx/read.js";
import { readXlsx } from "./xlsx/read.js";
import { isZip } from "./zip/zip.js";

export type WorkbookFormat = "xlsx" | "csv";

export interface ReadResult {
  format: WorkbookFormat;
  sheets: XlsxSheetData[];
  /** Defined names that resolve to a range. Empty for a CSV. */
  namedRanges: Record<string, Range>;
  /**
   * Defined names that hold a formula rather than a range. Empty for a CSV.
   * Bodies, without the leading `=`.
   */
  namedFormulas: Record<string, string>;
}

/**
 * Reads a .xlsx or .csv into sheet data. A CSV becomes a single sheet named
 * after the file, with no styles and no merges.
 *
 * Cancellable and progress-reporting via `options`:
 *
 *   const controller = new AbortController();
 *   const result = await readWorkbookFile(file, {
 *     signal: controller.signal,
 *     onProgress: ({ ratio, detail }) => setBar(ratio, detail),
 *   });
 *
 * An aborted read rejects with `AbortedError` (`code: "ABORTED"`) and leaves
 * nothing behind — the workbook is only built once the whole file has parsed, so
 * cancelling can never half-import.
 */
export async function readWorkbookFile(
  file: File | Blob,
  options: AsyncReadOptions = {},
): Promise<ReadResult> {
  const bytes = new Uint8Array(await file.arrayBuffer());
  const name = "name" in file && file.name ? file.name : "Sheet1";

  // ZIP magic alone does not prove this is an XLSX — .xlsb is also a ZIP, holding
  // binary BIFF12 parts. `readXlsx` distinguishes them by looking for
  // xl/workbook.xml vs xl/workbookBin.bin and throws a clear "unsupported format"
  // error for the latter.
  if (isZip(bytes)) {
    const sheets = await readXlsx(bytes, options);
    return {
      format: "xlsx",
      sheets,
      namedRanges: sheets.namedRanges ?? {},
      namedFormulas: sheets.namedFormulas ?? {},
    };
  }

  const text = new TextDecoder().decode(bytes);
  const { cells, rows, cols } = await csvToCells(text, options);
  return {
    format: "csv",
    // A CSV has no names of any kind.
    namedRanges: {},
    namedFormulas: {},
    sheets: [
      {
        name: name.replace(/\.[^.]+$/, "") || "Sheet1",
        cells,
        styles: {},
        cachedValues: {},
        condFormats: [],
        images: {},
        tables: [],
        spillRanges: {},
        merges: [],
        rows,
        cols,
        // CSV carries no layout information at all.
        colWidths: {},
        rowHeights: {},
        hiddenCols: [],
        hiddenRows: [],
        // A CSV has no names of any kind, sheet-scoped included.
        namedRanges: {},
        namedFormulas: {},
        validations: [],
      },
    ],
  };
}
