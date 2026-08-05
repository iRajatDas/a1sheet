/**
 * Format-detecting entry point for reading a spreadsheet file.
 *
 * Dispatch is by ZIP magic first and file extension second, so a .xlsx that has
 * been renamed still reads correctly.
 */
import { csvToCells } from "./csv/read.js";
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
 */
export async function readWorkbookFile(file: File | Blob): Promise<ReadResult> {
  const bytes = new Uint8Array(await file.arrayBuffer());
  const name = "name" in file && file.name ? file.name : "Sheet1";

  // ZIP magic alone does not prove this is an XLSX — .xlsb is also a ZIP, holding
  // binary BIFF12 parts. `readXlsx` distinguishes them by looking for
  // xl/workbook.xml vs xl/workbookBin.bin and throws a clear "unsupported format"
  // error for the latter.
  if (isZip(bytes)) {
    return { format: "xlsx", sheets: await readXlsx(bytes) };
  }

  const text = new TextDecoder().decode(bytes);
  const { cells, rows, cols } = csvToCells(text);
  return {
    format: "csv",
    sheets: [
      {
        name: name.replace(/\.[^.]+$/, "") || "Sheet1",
        cells,
        styles: {},
        merges: [],
        rows,
        cols,
      },
    ],
  };
}
