/**
 * CSV reader.
 *
 * RFC 4180 quoting: `""` inside a quoted field is a literal quote. Bare `\r` is
 * skipped, so CRLF and LF inputs both parse.
 */
import type { CellKey, RawCell } from "../../model/types.js";
import {
  type AsyncReadOptions,
  countOccurrences,
  createPacer,
} from "../progress.js";
import { denormalizeCsvValue } from "./sanitize.js";

/**
 * Yields CSV rows one at a time as fields are scanned.
 *
 * Lazy so `csvToCells` can hand control back to the browser between rows. A
 * 38 MB CSV is a single string of ~38 million characters; scanning it to
 * completion before the caller sees row one is exactly the multi-second freeze
 * §6 forbids.
 */
export function* iterCsvRows(text: string): Generator<string[]> {
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
      } else if (c === "\r" && text[i + 1] === "\n") {
        // Normalize the CRLF Excel writes inside a multi-line quoted cell. Done
        // here rather than by replacing over the whole input first, which would
        // copy the entire file — 38 MB of string — before parsing even starts.
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
      yield row;
      row = [];
      field = "";
    } else if (c === "\r") {
      // skip
    } else {
      field += c;
    }
  }

  row.push(field);
  yield row;
}

/**
 * Splits CSV text into a row-major array of raw field strings.
 *
 * Synchronous, so it holds the thread for the length of the input. Fine for
 * clipboard-sized text; for a file, prefer `csvToCells`, which paces itself.
 */
export function parseCSV(text: string): string[][] {
  return [...iterCsvRows(text)];
}

export interface CsvCells {
  cells: Record<CellKey, RawCell>;
  rows: number;
  cols: number;
}

/** Rows converted between checkpoints. See the note in ../progress.ts on cost. */
const ROWS_PER_CHECKPOINT = 512;

/**
 * Parses CSV text into the internal sparse cell map. Empty fields are omitted.
 *
 * Paced: pass `signal` to cancel and `onProgress` to drive a progress bar. Text
 * short enough to parse within a frame never yields, so pasted content is as fast
 * as it ever was.
 */
export async function csvToCells(
  text: string,
  options: AsyncReadOptions = {},
): Promise<CsvCells> {
  const pacer = createPacer(options);
  // Before the pre-scan, so a signal that is already aborted costs nothing.
  await pacer.checkpoint("parsing", 0, "start");

  const cells: Record<CellKey, RawCell> = {};
  let maxCols = 0;
  let r = 0;

  // Newlines are an upper bound on rows — one inside a quoted field is counted but
  // does not end a row — so the bar can lag slightly. It never overshoots or
  // reverses, which is what matters.
  const estimatedRows = Math.max(1, countOccurrences(text, "\n") + 1);

  for (const row of iterCsvRows(text)) {
    maxCols = Math.max(maxCols, row.length);
    row.forEach((val, c) => {
      // Undo the export-side injection guard so our own files round-trip exactly.
      const clean = denormalizeCsvValue(val);
      if (clean !== "") cells[`${r}_${c}`] = clean;
    });
    r++;
    if (r % ROWS_PER_CHECKPOINT === 0) {
      await pacer.checkpoint("parsing", r / estimatedRows, `row ${r}`);
    }
  }

  pacer.finish(`${r} rows`);
  return { cells, rows: r, cols: maxCols };
}
