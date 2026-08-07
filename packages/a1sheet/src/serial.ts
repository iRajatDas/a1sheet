/**
 * The day serial: the number a cell actually holds when it displays a date.
 *
 * A serial here is Excel's serial. `45520` is 2024-08-16 in this engine, in
 * Excel, in Sheets, and in the XLSX file on disk — one number, one meaning. That
 * is the point of this module and the reason it sits at the root rather than
 * under `io/`: an epoch is not a file format concern. It is what `=A1+1`,
 * `YEAR(A1)`, the date formatter, and the importer must all agree on, and the
 * moment any one of them counts from somewhere else the disagreement shows up as
 * a plausible date in the wrong century rather than as an error.
 *
 * Serial 1 is 1900-01-01 and fractions are the time of day, so 45520.5 is noon.
 *
 * Excel inherited a Lotus 1-2-3 bug: it believes 1900 was a leap year. Serial 60
 * is that phantom 1900-02-29, which means serials at or below it sit one day
 * further from the real calendar than arithmetic suggests. The phantom itself has
 * no real date and resolves to 1900-02-28 rather than throwing — a file
 * containing it is not a file worth refusing. That is the one asymmetry:
 * `msToSerial(serialToMs(60))` is 59, the serial that day really has.
 *
 * Everything is UTC. A serial has no timezone, and reading one back through a
 * local-time `Date` would move dates across midnight for half the world.
 */

/** Milliseconds in a day. Serials are days, so this is the only unit here. */
export const DAY_MS = 86400000;

/** Days from serial 0 (1899-12-30) to the Unix epoch, counting the phantom. */
const UNIX_EPOCH_SERIAL = 25569;

/** Excel's 1900-02-29. At and below it, a serial is one day nearer the epoch. */
const PHANTOM_LEAP_SERIAL = 60;

/** A serial as milliseconds since the Unix epoch, for `new Date`. */
export function serialToMs(serial: number): number {
  const offset =
    serial < PHANTOM_LEAP_SERIAL ? UNIX_EPOCH_SERIAL - 1 : UNIX_EPOCH_SERIAL;
  return (serial - offset) * DAY_MS;
}

/** The inverse. The phantom day is not representable and comes back as 59. */
export function msToSerial(ms: number): number {
  const serial = ms / DAY_MS + UNIX_EPOCH_SERIAL;
  return serial > PHANTOM_LEAP_SERIAL ? serial : serial - 1;
}

/** A serial split into UTC calendar parts. Invalid serials give `undefined`. */
export interface SerialParts {
  readonly year: number;
  readonly month: number;
  readonly day: number;
  readonly hours: number;
  readonly minutes: number;
  readonly seconds: number;
  /** 0 is Sunday, matching `Date.prototype.getUTCDay`. */
  readonly weekday: number;
}

export function serialToParts(serial: number): SerialParts | undefined {
  const d = new Date(Math.round(serialToMs(serial)));
  if (Number.isNaN(d.getTime())) return undefined;
  return {
    year: d.getUTCFullYear(),
    month: d.getUTCMonth() + 1,
    day: d.getUTCDate(),
    hours: d.getUTCHours(),
    minutes: d.getUTCMinutes(),
    seconds: d.getUTCSeconds(),
    weekday: d.getUTCDay(),
  };
}
