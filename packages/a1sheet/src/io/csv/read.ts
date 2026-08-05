/**
 * CSV reader. Ported from ref/xlsxIO.js:405-421 and 442-448.
 *
 * RFC 4180 quoting: `""` inside a quoted field is a literal quote. Bare `\r` is
 * skipped, so CRLF and LF inputs both parse.
 */
import type { CellKey, RawCell } from "../../model/types.js";

/** Splits CSV text into a row-major array of raw field strings. */
export function parseCSV(text: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = "";
  let inQuotes = false;

  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (inQuotes) {
      if (c === '"') {
        if (text[i + 1] === '"') {
          field += '"';
          i++;
        } else {
          inQuotes = false;
        }
      } else {
        field += c;
      }
    } else if (c === '"') {
      inQuotes = true;
    } else if (c === ",") {
      row.push(field);
      field = "";
    } else if (c === "\n") {
      row.push(field);
      rows.push(row);
      row = [];
      field = "";
    } else if (c === "\r") {
      // skip
    } else {
      field += c;
    }
  }

  row.push(field);
  rows.push(row);
  return rows;
}

export interface CsvCells {
  cells: Record<CellKey, RawCell>;
  rows: number;
  cols: number;
}

/** Parses CSV text into the internal sparse cell map. Empty fields are omitted. */
export function csvToCells(text: string): CsvCells {
  const rows = parseCSV(text.replace(/\r\n/g, "\n"));
  const cells: Record<CellKey, RawCell> = {};
  let maxCols = 0;

  rows.forEach((row, r) => {
    maxCols = Math.max(maxCols, row.length);
    row.forEach((val, c) => {
      if (val !== "") cells[`${r}_${c}`] = val;
    });
  });

  return { cells, rows: rows.length, cols: maxCols };
}
