/**
 * Number-format codes. Six buckets could not express what a real workbook
 * contains: a goal difference of `+45`, a kick-off at `8/16/24 20:00`, a negative
 * shown in parentheses.
 */
import { describe, expect, test } from "bun:test";
import { msToSerial } from "../serial.js";
import { applyFormatCode, isDateFormat } from "./formatCode.js";

describe("signs and sections", () => {
  test("a positive section's literal sign is emitted", () => {
    // The football workbook's goal-difference column. Bucketing this as "integer"
    // rendered 45, losing the sign the format exists to add.
    expect(applyFormatCode("+0;-0;0", 45)).toBe("+45");
  });

  test("the negative section supplies its own sign", () => {
    expect(applyFormatCode("+0;-0;0", -47)).toBe("-47");
  });

  test("the zero section is used for zero", () => {
    expect(applyFormatCode("+0;-0;0", 0)).toBe("0");
  });

  test("a negative with no section of its own still gets a minus", () => {
    expect(applyFormatCode("0.00", -3.5)).toBe("-3.50");
  });

  test("a negative section can wrap in parentheses instead of signing", () => {
    expect(applyFormatCode("#,##0;(#,##0)", -1234)).toBe("(1,234)");
  });

  test("an empty section hides the value", () => {
    expect(applyFormatCode("0;;", 0)).toBe("");
  });
});

describe("numbers", () => {
  test("decimal places come from the pattern", () => {
    expect(applyFormatCode("0.00", Math.PI)).toBe("3.14");
    expect(applyFormatCode("0", 3.6)).toBe("4");
  });

  test("thousands group only when the pattern says so", () => {
    expect(applyFormatCode("#,##0", 1234567)).toBe("1,234,567");
    expect(applyFormatCode("0", 1234567)).toBe("1234567");
  });

  test("a percentage scales as well as marking", () => {
    expect(applyFormatCode("0%", 0.25)).toBe("25%");
    expect(applyFormatCode("0.00%", 0.1234)).toBe("12.34%");
  });

  test("literals surround the digits where they sit", () => {
    expect(applyFormatCode('"$"#,##0.00', 1234.5)).toBe("$1,234.50");
    expect(applyFormatCode('0" kg"', 12)).toBe("12 kg");
  });

  test("an escaped character is a literal", () => {
    expect(applyFormatCode("\\+0;\\-0;0", 45)).toBe("+45");
  });

  test("a colour directive is stripped from the output", () => {
    expect(applyFormatCode("[Red]#,##0", 1234)).toBe("1,234");
  });
});

describe("dates and times", () => {
  /** 2024-08-16T20:00Z as a day serial, which is 45520.833… in Excel too. */
  const KICKOFF = msToSerial(Date.UTC(2024, 7, 16, 20, 0));

  test("a date-time format shows both halves", () => {
    // numFmtId 22, which the file states by id and carries no code of its own.
    expect(applyFormatCode("m/d/yy h:mm", KICKOFF)).toBe("8/16/24 20:00");
  });

  test("padded and unpadded parts differ", () => {
    expect(applyFormatCode("mm-dd-yy", KICKOFF)).toBe("08-16-24");
    expect(applyFormatCode("m-d-yyyy", KICKOFF)).toBe("8-16-2024");
  });

  test("m means month or minute depending on its neighbours", () => {
    // The whole difficulty of date patterns. After an hour or before a second it
    // is minutes; anywhere else it is the month.
    expect(applyFormatCode("h:mm", KICKOFF)).toBe("20:00");
    expect(applyFormatCode("mm/yyyy", KICKOFF)).toBe("08/2024");
    expect(applyFormatCode("mm:ss", KICKOFF)).toBe("00:00");
  });

  test("month and day names come from their run lengths", () => {
    expect(applyFormatCode("d-mmm-yy", KICKOFF)).toBe("16-Aug-24");
    expect(applyFormatCode("mmmm", KICKOFF)).toBe("August");
    expect(applyFormatCode("dddd", KICKOFF)).toBe("Friday");
    expect(applyFormatCode("ddd", KICKOFF)).toBe("Fri");
  });

  test("a 12-hour pattern wraps the hour and names the half", () => {
    expect(applyFormatCode("h:mm AM/PM", KICKOFF)).toBe("8:00 PM");
  });

  test("an elapsed-hours marker counts past 24", () => {
    // [h] is total hours, not the hour of the day: 1.5 days is 36 hours.
    expect(applyFormatCode("[h]:mm:ss", 1.5)).toBe("36:00:00");
  });
});

describe("text", () => {
  test("the fourth section renders a string through @", () => {
    expect(applyFormatCode('0;0;0;"<"@">"', "hello")).toBe("<hello>");
  });

  test("a string against a numeric-only code declines rather than guesses", () => {
    // Returning null lets the caller fall back instead of printing something wrong.
    expect(applyFormatCode("0.00", "hello")).toBeNull();
  });
});

describe("telling a date format from a numeric one", () => {
  test("date letters in the body count", () => {
    expect(isDateFormat("m/d/yy")).toBe(true);
    expect(isDateFormat("h:mm:ss")).toBe(true);
  });

  test("date letters inside a directive or a literal do not", () => {
    // [Red] contains a d, and "days" contains three date letters.
    expect(isDateFormat("[Red]#,##0")).toBe(false);
    expect(isDateFormat('0" days"')).toBe(false);
  });
});
