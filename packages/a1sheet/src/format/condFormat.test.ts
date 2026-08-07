/**
 * Applying conditional formats. Kept out of the renderer so the rules can be
 * tested as rules — a render test would only show that some colour arrived.
 */
import { describe, expect, test } from "bun:test";
import { createEvaluator } from "../formula/evaluate.js";
import type { CellKey, CondFormat, RawCell } from "../model/types.js";
import { condStyleFor } from "./condFormat.js";

const cells = {
  "0_0": "5",
  "0_1": "",
  "1_0": "50",
  "2_0": "hello",
} as Record<CellKey, RawCell>;

const evaluator = createEvaluator(cells, {});
const at = (condFormats: readonly CondFormat[], row: number, col: number) =>
  condStyleFor({ condFormats, evaluator }, row, col);

const RED = { bg: "#ff0000" } as const;
const BLUE = { bg: "#0000ff" } as const;

describe("an expression rule", () => {
  test("applies where its formula is true", () => {
    const formats: CondFormat[] = [
      {
        range: { r1: 0, c1: 0, r2: 0, c2: 0 },
        priority: 1,
        rule: { type: "expression", formula: "A1>1" },
        style: RED,
      },
    ];
    expect(at(formats, 0, 0)).toEqual(RED);
  });

  test("does not apply where its formula is false", () => {
    const formats: CondFormat[] = [
      {
        range: { r1: 0, c1: 0, r2: 0, c2: 0 },
        priority: 1,
        rule: { type: "expression", formula: "A1>100" },
        style: RED,
      },
    ];
    expect(at(formats, 0, 0)).toBeUndefined();
  });

  test("shifts relative references per cell, and leaves absolute ones alone", () => {
    // A rule's formula is written for the top-left of its range and copied across
    // the rest, exactly as a dragged formula is. This is why `$C$4` tests one cell
    // for a whole range and `A1` tests each cell against itself.
    const relative: CondFormat[] = [
      {
        range: { r1: 0, c1: 0, r2: 1, c2: 0 },
        priority: 1,
        rule: { type: "expression", formula: "A1>10" },
        style: RED,
      },
    ];
    // A1 is 5, A2 is 50 — so only the second row matches.
    expect(at(relative, 0, 0)).toBeUndefined();
    expect(at(relative, 1, 0)).toEqual(RED);

    const absolute: CondFormat[] = [
      {
        range: { r1: 0, c1: 0, r2: 1, c2: 0 },
        priority: 1,
        rule: { type: "expression", formula: "$A$2>10" },
        style: RED,
      },
    ];
    expect(at(absolute, 0, 0)).toEqual(RED);
    expect(at(absolute, 1, 0)).toEqual(RED);
  });

  test("a formula that errors does not match", () => {
    const formats: CondFormat[] = [
      {
        range: { r1: 0, c1: 0, r2: 0, c2: 0 },
        priority: 1,
        rule: { type: "expression", formula: "NOSUCHFUNC(1)" },
        style: RED,
      },
    ];
    expect(at(formats, 0, 0)).toBeUndefined();
  });

  test("a cell outside every range gets nothing", () => {
    const formats: CondFormat[] = [
      {
        range: { r1: 0, c1: 0, r2: 0, c2: 0 },
        priority: 1,
        rule: { type: "expression", formula: "1" },
        style: RED,
      },
    ];
    expect(at(formats, 9, 9)).toBeUndefined();
  });
});

describe("value rules", () => {
  test("cellIs compares the cell against its operands", () => {
    const between: CondFormat[] = [
      {
        range: { r1: 0, c1: 0, r2: 2, c2: 0 },
        priority: 1,
        rule: { type: "cellIs", operator: "between", operands: ["1", "10"] },
        style: RED,
      },
    ];
    expect(at(between, 0, 0)).toEqual(RED);
    expect(at(between, 1, 0)).toBeUndefined();
  });

  test("containsBlanks and its negation split empty from filled", () => {
    const blanks: CondFormat[] = [
      {
        range: { r1: 0, c1: 0, r2: 0, c2: 1 },
        priority: 1,
        rule: { type: "containsBlanks", negate: false },
        style: RED,
      },
    ];
    expect(at(blanks, 0, 1)).toEqual(RED);
    expect(at(blanks, 0, 0)).toBeUndefined();
  });

  test("containsText matches a substring of the displayed value", () => {
    const text: CondFormat[] = [
      {
        range: { r1: 2, c1: 0, r2: 2, c2: 0 },
        priority: 1,
        rule: { type: "containsText", text: "ell", negate: false },
        style: RED,
      },
    ];
    expect(at(text, 2, 0)).toEqual(RED);
  });
});

describe("several rules on one cell", () => {
  const range = { r1: 0, c1: 0, r2: 0, c2: 0 };

  test("the higher-priority rule wins where they collide", () => {
    // Excel counts priority 1 as most important, which is the reverse of the
    // order they have to be applied in.
    const formats: CondFormat[] = [
      {
        range,
        priority: 2,
        rule: { type: "expression", formula: "1" },
        style: BLUE,
      },
      {
        range,
        priority: 1,
        rule: { type: "expression", formula: "1" },
        style: RED,
      },
    ];
    expect(at(formats, 0, 0)?.bg).toBe("#ff0000");
  });

  test("keys only one rule sets are all kept", () => {
    const formats: CondFormat[] = [
      {
        range,
        priority: 2,
        rule: { type: "expression", formula: "1" },
        style: BLUE,
      },
      {
        range,
        priority: 1,
        rule: { type: "expression", formula: "1" },
        style: { bold: true },
      },
    ];
    expect(at(formats, 0, 0)).toEqual({ bg: "#0000ff", bold: true });
  });

  test("stopIfTrue prevents lower-priority rules from being considered", () => {
    const formats: CondFormat[] = [
      {
        range,
        priority: 2,
        rule: { type: "expression", formula: "1" },
        style: BLUE,
      },
      {
        range,
        priority: 1,
        rule: { type: "expression", formula: "1" },
        style: { bold: true },
        stopIfTrue: true,
      },
    ];
    expect(at(formats, 0, 0)).toEqual({ bold: true });
  });
});

describe("statistical rules", () => {
  // 10, 20, 30, 40 down column A of a sheet of its own.
  const ranked = createEvaluator(
    {
      "0_0": "10",
      "1_0": "20",
      "2_0": "30",
      "3_0": "40",
    } as Record<CellKey, RawCell>,
    {},
  );
  const range = { r1: 0, c1: 0, r2: 3, c2: 0 };
  const style = { bg: "#ff0000" } as const;
  const over = (rule: CondFormat["rule"], row: number) =>
    condStyleFor(
      { condFormats: [{ range, priority: 1, rule, style }], evaluator: ranked },
      row,
      0,
    );

  test("top10 marks the highest N", () => {
    const rule = { type: "top10", rank: 2, bottom: false, percent: false } as const;
    expect(over(rule, 3)).toEqual(style);
    expect(over(rule, 2)).toEqual(style);
    expect(over(rule, 1)).toBeUndefined();
  });

  test("bottom marks the lowest N instead", () => {
    const rule = { type: "top10", rank: 1, bottom: true, percent: false } as const;
    expect(over(rule, 0)).toEqual(style);
    expect(over(rule, 3)).toBeUndefined();
  });

  test("percent makes the rank a share of the range", () => {
    // 50% of four values is the top two.
    const rule = { type: "top10", rank: 50, bottom: false, percent: true } as const;
    expect(over(rule, 2)).toEqual(style);
    expect(over(rule, 1)).toBeUndefined();
  });

  test("aboveAverage compares against the range's mean", () => {
    // The mean of 10, 20, 30, 40 is 25.
    const rule = { type: "aboveAverage", below: false, orEqual: false } as const;
    expect(over(rule, 2)).toEqual(style);
    expect(over(rule, 1)).toBeUndefined();
  });

  test("below and orEqual flip and widen it", () => {
    expect(over({ type: "aboveAverage", below: true, orEqual: false }, 1)).toEqual(
      style,
    );
    expect(
      over({ type: "aboveAverage", below: true, orEqual: false }, 2),
    ).toBeUndefined();
  });

  test("a non-numeric cell is never marked", () => {
    const text = createEvaluator(
      { "0_0": "hello" } as Record<CellKey, RawCell>,
      {},
    );
    expect(
      condStyleFor(
        {
          condFormats: [
            {
              range,
              priority: 1,
              rule: { type: "aboveAverage", below: false, orEqual: false },
              style,
            },
          ],
          evaluator: text,
        },
        0,
        0,
      ),
    ).toBeUndefined();
  });
});
