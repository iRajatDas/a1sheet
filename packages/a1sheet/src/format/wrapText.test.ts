/**
 * Line counting for wrapped cells.
 *
 * The number this produces becomes a row's height before the text is laid out,
 * so it has to match what the browser will do closely enough that the row is
 * neither short (clipping the last line) nor tall (a visible gap). The cases
 * here are the ones where a greedy filler and a browser could disagree.
 *
 * A monospace measurer keeps the arithmetic legible: every character is 10px, so
 * a width of 100 fits exactly ten of them.
 */
import { describe, expect, test } from "bun:test";
import { wrappedLineCount } from "./wrapText.js";

const CHAR = 10;
const mono = (text: string) => text.length * CHAR;
const lines = (text: string, chars: number) =>
  wrappedLineCount(text, chars * CHAR, mono);

describe("filling lines", () => {
  test("text that fits is one line", () => {
    expect(lines("abc", 10)).toBe(1);
  });

  test("exactly filling the width does not spill onto a second line", () => {
    // The off-by-one that shows up as a stray empty line under every full row.
    expect(lines("abcdefghij", 10)).toBe(1);
    expect(lines("abcdefghijk", 10)).toBe(2);
  });

  test("words pack greedily, and the space between them counts", () => {
    // "aaa bbb" is 7 characters including the space, so it fits in 7 but not 6.
    expect(lines("aaa bbb", 7)).toBe(1);
    expect(lines("aaa bbb", 6)).toBe(2);
  });

  test("a run of spaces collapses, as white-space: normal makes it", () => {
    expect(lines("aaa     bbb", 7)).toBe(1);
  });

  test("leading and trailing whitespace does not create empty lines", () => {
    expect(lines("   abc   ", 10)).toBe(1);
  });
});

describe("words too long for a line", () => {
  test("a long word breaks mid-word rather than overflowing", () => {
    // overflow-wrap: anywhere. Without it this is one line that runs outside the
    // cell, which is the clipping the row growth exists to prevent.
    expect(lines("abcdefghijklmnopqrst", 10)).toBe(2);
  });

  test("a broken word leaves its tail available to the next word", () => {
    // "abcdefghijkl" fills one line and leaves "kl"; "mn" then joins it rather
    // than starting a third line.
    expect(lines("abcdefghijkl mn", 10)).toBe(2);
  });

  test("a word starting mid-line moves down before it breaks", () => {
    expect(lines("ab cdefghijklmnopqrst", 10)).toBe(3);
  });
});

describe("breaks the user typed", () => {
  test("a newline always breaks, however much room is left", () => {
    expect(lines("a\nb", 10)).toBe(2);
  });

  test("every newline convention counts once", () => {
    expect(lines("a\r\nb", 10)).toBe(2);
    expect(lines("a\rb", 10)).toBe(2);
  });

  test("a blank line is a line", () => {
    expect(lines("a\n\nb", 10)).toBe(3);
  });

  test("each paragraph wraps on its own", () => {
    expect(lines("abcdefghijk\nabcdefghijk", 10)).toBe(4);
  });
});

describe("degenerate input", () => {
  test("empty text still occupies its line box", () => {
    expect(lines("", 10)).toBe(1);
  });

  test("a width of zero answers one line rather than looping forever", () => {
    // What a hidden column reports. Not a crash and not an unbounded loop.
    expect(wrappedLineCount("abc", 0, mono)).toBe(1);
    expect(wrappedLineCount("abc", -5, mono)).toBe(1);
  });

  test("a measurer that measures nothing yields one line", () => {
    // No canvas — SSR, or a test environment without one. Everything fits, so
    // no row grows, which is the same behaviour as before wrapping existed.
    expect(wrappedLineCount("a very long sentence indeed", 10, () => 0)).toBe(1);
  });
});
