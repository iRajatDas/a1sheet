/**
 * Date functions. Ported from ref/formulaEngine.js:216-221.
 *
 * These speak day serials, which are Excel's day serials — see `src/serial.ts`
 * for what that means and why the 1900 leap year is a lie.
 */
import { msToSerial, type SerialParts, serialToParts } from "../../serial.js";
import { toNumber } from "../values.js";
import type { FormulaFunction } from "./registry.js";

/** Excel's error for a serial no calendar date corresponds to. */
const NUM_ERROR = "#NUM!";

/** Lifts a UTC part accessor into a function that fails on an unusable serial. */
function part(read: (parts: SerialParts) => number): FormulaFunction {
  return (a) => {
    const parts = serialToParts(toNumber(a[0]));
    return parts === undefined ? NUM_ERROR : read(parts);
  };
}

export const dateFunctions: Record<string, FormulaFunction> = {
  // Volatile: `host.now` is the cycle's instant, so every TODAY and NOW on the
  // sheet agrees, and all of them move together on the next cycle.
  TODAY: (_args, host) => Math.floor(msToSerial(host.now)),
  NOW: (_args, host) => msToSerial(host.now),

  /** Month is 1-indexed on the way in, matching Excel. */
  DATE: (a) => {
    const ms = Date.UTC(toNumber(a[0]), toNumber(a[1]) - 1, toNumber(a[2]));
    return Number.isNaN(ms) ? NUM_ERROR : Math.floor(msToSerial(ms));
  },

  YEAR: part((p) => p.year),
  MONTH: part((p) => p.month),
  DAY: part((p) => p.day),
};
