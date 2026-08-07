/**
 * Volatile functions: the ones whose result is not a function of any cell.
 *
 * Two properties define them, and they pull in opposite directions. WITHIN one
 * calculation a volatile must hold still, or `=RAND()` in A1 and `=A1*2` in B1
 * disagree about what A1 is. ACROSS calculations it must move, or it is a
 * constant with extra steps. The evaluator is the calculation: it memoizes for
 * its own lifetime, and a new one is a new cycle.
 */
import { describe, expect, test } from "bun:test";
import type { CellKey, RawCell } from "../model/types.js";
import { msToSerial } from "../serial.js";
import { createEvaluator } from "./evaluate.js";

function sheet(cells: Record<string, string>, now?: number) {
  return createEvaluator(
    cells as Record<CellKey, RawCell>,
    {},
    now === undefined ? {} : { now },
  );
}

describe("holding still within one calculation", () => {
  test("a cell reading RAND sees the same number RAND returned", () => {
    const e = sheet({ "0_0": "=RAND()", "0_1": "=A1", "0_2": "=A1*2" });

    const a = e.getCellDisplay(0, 0) as number;
    expect(e.getCellDisplay(0, 1)).toBe(a);
    expect(e.getCellDisplay(0, 2)).toBe(a * 2);
  });

  test("asking the same cell twice gives the same answer", () => {
    const e = sheet({ "0_0": "=RAND()" });

    expect(e.getCellDisplay(0, 0)).toBe(e.getCellDisplay(0, 0));
  });

  test("two NOW cells report the same instant", () => {
    // Not guaranteed by memoization: these are different cells, and calling
    // Date.now() in each could land them either side of a millisecond.
    const e = sheet({ "0_0": "=NOW()", "0_1": "=NOW()", "0_2": "=TODAY()" });

    expect(e.getCellDisplay(0, 1)).toBe(e.getCellDisplay(0, 0));
    expect(e.getCellDisplay(0, 2)).toBe(
      Math.floor(e.getCellDisplay(0, 0) as number),
    );
  });

  test("NOW reports the instant the cycle was given, not the wall clock", () => {
    const e = sheet({ "0_0": "=NOW()" }, Date.UTC(2024, 7, 16, 20, 0));

    expect(e.getCellDisplay(0, 0)).toBe(msToSerial(Date.UTC(2024, 7, 16, 20, 0)));
  });
});

describe("moving across calculations", () => {
  test("a new calculation gives RAND a new number", () => {
    const cells = { "0_0": "=RAND()" };
    const seen = new Set<unknown>();
    for (let i = 0; i < 8; i++) seen.add(sheet(cells).getCellDisplay(0, 0));

    // Eight independent draws colliding on one value would mean RAND is frozen.
    expect(seen.size).toBeGreaterThan(1);
  });

  test("a new calculation moves NOW forward", () => {
    const cells = { "0_0": "=NOW()" };
    const earlier = sheet(cells, 1_000_000).getCellDisplay(0, 0) as number;
    const later = sheet(cells, 2_000_000).getCellDisplay(0, 0) as number;

    expect(later).toBeGreaterThan(earlier);
  });
});

describe("RANDBETWEEN", () => {
  test("stays inside the interval, both ends included", () => {
    const e = sheet({ "0_0": "=RANDBETWEEN(1,6)" });
    const roll = e.getCellDisplay(0, 0) as number;

    expect(Number.isInteger(roll)).toBe(true);
    expect(roll).toBeGreaterThanOrEqual(1);
    expect(roll).toBeLessThanOrEqual(6);
  });

  test("can reach both ends", () => {
    const cells = { "0_0": "=RANDBETWEEN(0,1)" };
    const seen = new Set<unknown>();
    for (let i = 0; i < 40; i++) seen.add(sheet(cells).getCellDisplay(0, 0));

    // An off-by-one in the interval shows up as one of these never appearing.
    expect([...seen].sort()).toEqual([0, 1]);
  });

  test("equal bounds are that number, not an error", () => {
    expect(sheet({ "0_0": "=RANDBETWEEN(7,7)" }).getCellDisplay(0, 0)).toBe(7);
  });

  test("fractional bounds move inward, so the result is inside them", () => {
    const cells = { "0_0": "=RANDBETWEEN(1.2, 2.9)" };
    for (let i = 0; i < 20; i++) {
      expect(sheet(cells).getCellDisplay(0, 0)).toBe(2);
    }
  });

  test("an interval containing no integer is #NUM!", () => {
    expect(sheet({ "0_0": "=RANDBETWEEN(1.2, 1.8)" }).getCellDisplay(0, 0)).toBe(
      "#NUM!",
    );
    expect(sheet({ "0_0": "=RANDBETWEEN(6,1)" }).getCellDisplay(0, 0)).toBe(
      "#NUM!",
    );
  });
});
