/**
 * The day serial, which is now Excel's day serial everywhere.
 *
 * The engine used to count days from the Unix epoch and rebase on the way in and
 * out of a file. That round-tripped, so it looked fine — but it meant the number
 * in the cell was seventy years from the number Excel puts there, and every
 * consumer that did not go through the importer saw the wrong one: a serial typed
 * straight into a cell, a serial pasted from Excel, `DATE()` compared against a
 * literal. These tests pin the serials themselves, not just the round trip.
 */
import { describe, expect, test } from "bun:test";
import { formatValue } from "./format/numFmt.js";
import { createEvaluator } from "./formula/evaluate.js";
import type { CellKey, RawCell } from "./model/types.js";
import { msToSerial, serialToMs, serialToParts } from "./serial.js";

const iso = (serial: number) =>
  new Date(serialToMs(serial)).toISOString().slice(0, 10);

function evaluate(formula: string): unknown {
  const cells: Record<CellKey, RawCell> = { "0_0": formula };
  return createEvaluator(cells, {}).getCellDisplay(0, 0);
}

describe("the serials themselves", () => {
  test("a serial means what Excel says it means", () => {
    // The whole bug in one assertion. Under the old Unix epoch this serial was
    // 2094-08-18 — a plausible date, wrong century, no error anywhere.
    expect(iso(45520)).toBe("2024-08-16");
    expect(iso(1)).toBe("1900-01-01");
    expect(iso(25569)).toBe("1970-01-01");
  });

  test("the fraction is the time of day", () => {
    expect(serialToParts(45520.5)?.hours).toBe(12);
  });

  test("the phantom 1900 leap day does not shift the dates around it", () => {
    // Excel believes 1900-02-29 existed, so the offset differs either side of it.
    expect(iso(59)).toBe("1900-02-28");
    expect(iso(61)).toBe("1900-03-01");
  });

  test("the phantom day collapses onto the real one it precedes", () => {
    // 60 is 1900-02-29, which never happened, so it cannot round trip. It reads
    // as 1900-02-28 and comes back as 59 — the serial that day really has.
    expect(iso(60)).toBe(iso(59));
    expect(msToSerial(serialToMs(60))).toBe(59);
  });

  test("every real serial survives a round trip", () => {
    for (const serial of [1, 59, 61, 366, 25569, 45520, 60000]) {
      expect(msToSerial(serialToMs(serial))).toBe(serial);
    }
  });

  test("a serial no date can be built from has no parts rather than NaN ones", () => {
    expect(serialToParts(Number.POSITIVE_INFINITY)).toBeUndefined();
    expect(serialToParts(1e15)).toBeUndefined();
  });
});

describe("what the rest of the engine does with them", () => {
  test("DATE produces the serial Excel produces", () => {
    expect(evaluate("=DATE(2024,8,16)")).toBe(45520);
  });

  test("adding one to a date gives the next day, across the month end", () => {
    // Day arithmetic worked under the old epoch too; it is asserted here because
    // moving the epoch is exactly the change that could have broken it.
    expect(evaluate("=YEAR(DATE(2024,2,29)+1)")).toBe(2024);
    expect(evaluate("=MONTH(DATE(2024,2,29)+1)")).toBe(3);
    expect(evaluate("=DAY(DATE(2024,2,29)+1)")).toBe(1);
  });

  test("a serial typed straight into a cell formats as the date Excel shows", () => {
    expect(formatValue(45520, { numFmt: "date" })).toBe("2024-08-16");
  });

  test("a format code renders the same serial the same way", () => {
    expect(formatValue(45520, { numFmtCode: "yyyy-mm-dd" })).toBe("2024-08-16");
  });

  test("TODAY is today, and NOW is inside it", () => {
    const today = evaluate("=TODAY()") as number;
    const now = evaluate("=NOW()") as number;

    expect(iso(today)).toBe(new Date().toISOString().slice(0, 10));
    expect(now - today).toBeGreaterThanOrEqual(0);
    expect(now - today).toBeLessThan(1);
  });

  test("a date part of an unusable serial is #NUM! rather than NaN", () => {
    // Beyond the ±8.64e15 ms a JS Date can hold, so there is no year to return.
    expect(evaluate("=YEAR(99999999999)")).toBe("#NUM!");
  });
});
