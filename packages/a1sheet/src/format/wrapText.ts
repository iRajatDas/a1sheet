/**
 * How many lines a string takes when wrapped to a width.
 *
 * This exists to answer "how tall does this row need to be", which has to be
 * known before the cell is laid out — so it cannot be read back off the DOM, and
 * it has to reproduce the browser's line breaking closely enough that the height
 * it predicts is the height the text turns out to need. Being one line short
 * clips the last line; being one line over leaves a visible gap.
 *
 * What the grid's cells declare is `white-space: normal` plus
 * `overflow-wrap: anywhere`, and that is exactly the algorithm here: break at
 * spaces, and break *inside* a word that cannot fit on a line of its own rather
 * than letting it overflow. `anywhere` rather than `break-word` is also why a
 * long word contributes several lines instead of one long one.
 *
 * Measurement is injected. The caller owns the canvas (see `useTextMeasurer`),
 * this module owns the line breaking, and neither needs the other to be testable.
 */

/** Measures one run of text in whatever font the caller has already selected. */
export type MeasureRun = (text: string) => number;

/** A newline the user typed. Always a break, whatever the width. */
const HARD_BREAK = /\r\n|\r|\n/;

/**
 * Lines `text` occupies at `width` pixels. Always at least 1, even for `""` — an
 * empty cell still has a line box.
 *
 * A width at or below zero means the column is too narrow to lay anything out
 * in; the answer is 1 rather than a division by zero or an unbounded loop.
 */
export function wrappedLineCount(
  text: string,
  width: number,
  measure: MeasureRun,
): number {
  if (text === "") return 1;
  if (width <= 0) return 1;

  let lines = 0;
  for (const paragraph of text.split(HARD_BREAK)) {
    lines += paragraphLines(paragraph, width, measure);
  }
  return lines;
}

/** Greedy line filling over one hard-break-free run. */
function paragraphLines(
  paragraph: string,
  width: number,
  measure: MeasureRun,
): number {
  if (paragraph === "") return 1;

  let lines = 1;
  // What is on the current line, kept as text rather than a width: measuring the
  // whole line each time is what accounts for kerning between the words, which
  // summing individual word widths does not.
  let line = "";

  for (const word of splitWords(paragraph)) {
    const candidate = line === "" ? word : `${line} ${word}`;
    if (measure(candidate) <= width) {
      line = candidate;
      continue;
    }

    // The word does not fit after what is already there. Move to a fresh line
    // first, then deal with a word too long for even an empty one.
    if (line !== "") {
      lines++;
      line = "";
    }
    if (measure(word) <= width) {
      line = word;
      continue;
    }
    const broken = breakLongWord(word, width, measure);
    lines += broken.lines - 1;
    line = broken.tail;
  }

  return lines;
}

/**
 * Splits on whitespace, dropping it. The space between words is re-added when a
 * candidate line is assembled, so a run of several spaces collapses — which is
 * what `white-space: normal` does too.
 */
function splitWords(paragraph: string): string[] {
  return paragraph.split(/\s+/).filter((w) => w !== "");
}

/**
 * A word wider than the whole line, broken character by character as
 * `overflow-wrap: anywhere` does.
 *
 * Returns the lines it fills and what is left on the last of them, so the caller
 * can keep packing words after it.
 */
function breakLongWord(
  word: string,
  width: number,
  measure: MeasureRun,
): { lines: number; tail: string } {
  let lines = 1;
  let tail = "";

  for (const char of word) {
    const candidate = tail + char;
    if (tail !== "" && measure(candidate) > width) {
      lines++;
      tail = char;
      continue;
    }
    tail = candidate;
  }

  return { lines, tail };
}
