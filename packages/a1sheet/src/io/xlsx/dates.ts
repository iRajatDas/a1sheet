/**
 * Converting between Excel's day serial and ours.
 *
 * Both are "days since an epoch", and the epochs differ by seventy years:
 * SpreadsheetML counts from 1899-12-30 and this engine counts from the Unix
 * epoch (see `formula/values.ts`). Importing a serial without converting it does
 * not look broken, which is what makes it dangerous — 2024-08-16 reads back as
 * 2094-08-18, a plausible date in the wrong century.
 *
 * The offset is not a single constant because Excel inherited a Lotus 1-2-3 bug:
 * it believes 1900 was a leap year, so serials 1..59 are one day further from the
 * epoch than arithmetic suggests. Serial 60 is the phantom 1900-02-29 itself,
 * which has no real date; it maps to 1900-02-28 rather than throwing, because a
 * file containing it is not a file worth refusing.
 */

/** Days from 1899-12-30 to 1970-01-01, counting Excel's phantom 1900-02-29. */
const EPOCH_OFFSET_DAYS = 25569;
/** Excel's 1900-02-29. Serials below it are one day nearer the epoch than this. */
const PHANTOM_LEAP_SERIAL = 60;

export function excelSerialToDaySerial(serial: number): number {
  return serial >= PHANTOM_LEAP_SERIAL
    ? serial - EPOCH_OFFSET_DAYS
    : serial - EPOCH_OFFSET_DAYS + 1;
}

/**
 * The inverse, except for the phantom day itself: serials 59 and 60 both denote
 * 1900-02-28 here, and this returns the real one. Nothing else is asymmetric.
 */
export function daySerialToExcelSerial(serial: number): number {
  const shifted = serial + EPOCH_OFFSET_DAYS;
  return shifted > PHANTOM_LEAP_SERIAL ? shifted : shifted - 1;
}
