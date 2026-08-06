import { describe, expect, test } from "bun:test";
import { explainErrorValue } from "./errorText.js";
import { createEvaluator, ERROR_VALUES } from "./evaluate.js";

describe("explainErrorValue", () => {
  test("every error sentinel has an explanation", () => {
    // A new sentinel without a sentence is a cell the user cannot act on.
    for (const sentinel of ERROR_VALUES) {
      expect(explainErrorValue(sentinel)).toBeString();
    }
  });

  test("a normal value has none", () => {
    expect(explainErrorValue("hello")).toBeNull();
    expect(explainErrorValue(42)).toBeNull();
    expect(explainErrorValue("")).toBeNull();
    expect(explainErrorValue(undefined)).toBeNull();
  });

  test("text that merely starts with # is not an error", () => {
    expect(explainErrorValue("#hashtag")).toBeNull();
  });

  test("the circular-reference wording says what to do about it", () => {
    const text = explainErrorValue("#CYCLE!") as string;
    expect(text).toContain("Circular reference");
    expect(text.length).toBeGreaterThan(40);
  });
});

describe("explanations reach real evaluated errors", () => {
  test("a self-referencing SUM explains itself", () => {
    // C1 sums a range that includes C1 — the case from the Sheets screenshot.
    const evaluator = createEvaluator(
      { "0_0": "1", "0_1": "2", "0_2": "=SUM(A1:C1)" },
      {},
    );
    expect(explainErrorValue(evaluator.getCellDisplay(0, 2))).toContain(
      "Circular reference",
    );
  });

  test("division by zero explains itself", () => {
    const evaluator = createEvaluator({ "0_0": "=1/0" }, {});
    expect(explainErrorValue(evaluator.getCellDisplay(0, 0))).toContain(
      "Division by zero",
    );
  });

  test("a working formula explains nothing", () => {
    const evaluator = createEvaluator({ "0_0": "1", "0_1": "=A1+1" }, {});
    expect(explainErrorValue(evaluator.getCellDisplay(0, 1))).toBeNull();
  });
});
