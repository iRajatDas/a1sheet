/**
 * Format-detecting entry point for reading a spreadsheet file.
 *
 * Dispatch is by ZIP magic first and file extension second, so a .xlsx that has
 * been renamed still reads correctly.
 */
import { csvToCells } from "./csv/read.js";
import type { AsyncReadOptions } from "./progress.js";
import type { XlsxSheetData } from "./xlsx/read.js";
import { readXlsx } from "./xlsx/read.js";
import { isZip } from "./zip/zip.js";

export type WorkbookFormat = "xlsx" | "csv";

export interface ReadResult {
  format: WorkbookFormat;
  sheets: XlsxSheetData[];
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
    return { format: "xlsx", sheets: await readXlsx(bytes, options) };
  }

  const text = new TextDecoder().decode(bytes);
  const { cells, rows, cols } = await csvToCells(text, options);
  return {
    format: "csv",
    sheets: [
      {
        name: name.replace(/\.[^.]+$/, "") || "Sheet1",
        cells,
        styles: {},
        cachedValues: {},
        condFormats: [],
        merges: [],
        rows,
        cols,
        // CSV carries no layout information at all.
        colWidths: {},
        rowHeights: {},
      },
    ],
  };
}
