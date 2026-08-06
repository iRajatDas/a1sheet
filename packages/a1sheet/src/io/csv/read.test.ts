/**
 * CSV parsing, through the public API.
 *
 * The CRLF cases matter disproportionately: normalization used to happen by
 * replacing over the whole input, which copied the entire file before parsing
 * started. It now happens inline in the quoted-field branch, and these tests pin
 * the behavior that change had to preserve exactly.
 */
import { describe, expect, test } from "bun:test";
import { csvToCells, parseCSV } from "./read.js";

describe("parseCSV", () => {
  test("splits plain rows and fields", () => {
    expect(parseCSV("a,b\nc,d")).toEqual([
      ["a", "b"],
      ["c", "d"],
    ]);
  });

  test("a quoted field keeps its commas", () => {
    expect(parseCSV('"Smith, John",42')).toEqual([["Smith, John", "42"]]);
  });

  test('"" inside a quoted field is one literal quote', () => {
    expect(parseCSV('"say ""hi""",x')).toEqual([['say "hi"', "x"]]);
  });

  test("CRLF between rows parses the same as LF", () => {
    expect(parseCSV("a,b\r\nc,d")).toEqual(parseCSV("a,b\nc,d"));
  });

  test("a quoted field spanning lines keeps its newline", () => {
    expect(parseCSV('"line1\nline2",x')).toEqual([["line1\nline2", "x"]]);
  });

  test("CRLF inside a quoted field is normalized to LF", () => {
    // Excel writes CRLF inside multi-line cells. A stray \r would surface in the
    // cell value and then round-trip back out into every export.
    expect(parseCSV('"line1\r\nline2",x')).toEqual([["line1\nline2", "x"]]);
  });

  test("a lone CR inside a quoted field is preserved", () => {
    expect(parseCSV('"a\rb"')).toEqual([["a\rb"]]);
  });

  test("empty input is one empty field, not zero rows", () => {
    expect(parseCSV("")).toEqual([[""]]);
  });

  test("a trailing newline does not invent a row of content", () => {
    expect(parseCSV("a\n")).toEqual([["a"], [""]]);
  });
});

describe("csvToCells", () => {
  test("omits empty fields and reports the extent", async () => {
    const { cells, rows, cols } = await csvToCells("a,,c\n,b");
    expect(cells).toEqual({ "0_0": "a", "0_2": "c", "1_1": "b" });
    expect(rows).toBe(2);
    expect(cols).toBe(3);
  });

  test("cols is the widest row, not the first", async () => {
    const { cols } = await csvToCells("a\na,b,c\na,b");
    expect(cols).toBe(3);
  });
});
