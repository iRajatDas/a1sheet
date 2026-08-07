/**
 * Browser download helpers.
 *
 * These are the only functions in the non-React entrypoint that touch the DOM.
 * They throw a clear error outside a browser rather than failing on an undefined
 * global, so the same import works in Node and in a Worker as long as you call
 * the byte-producing functions instead.
 */
import type { Evaluator } from "../formula/evaluate.js";
import type { CellKey, RawCell } from "../model/types.js";
import { safeFilename } from "./csv/sanitize.js";
import { cellsToCSV } from "./csv/write.js";
import type { XlsxSheetInput } from "./xlsx/write.js";
import { writeXlsx } from "./xlsx/write.js";

function saveBlob(blob: Blob, filename: string): void {
  if (typeof document === "undefined") {
    throw new Error(
      "download helpers require a browser; use writeXlsx/cellsToCSV to get bytes instead",
    );
  }
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

export function downloadXlsx(
  sheets: XlsxSheetInput[],
  filename = "spreadsheet.xlsx",
): void {
  const bytes = writeXlsx(sheets);
  saveBlob(
    new Blob([bytes as unknown as BlobPart], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }),
    safeFilename(filename, "spreadsheet.xlsx"),
  );
}

export function downloadCsv(
  cells: Record<CellKey, RawCell>,
  evaluator: Evaluator,
  filename = "spreadsheet.csv",
): void {
  saveBlob(
    new Blob([cellsToCSV(cells, evaluator)], { type: "text/csv" }),
    safeFilename(filename, "spreadsheet.csv"),
  );
}
