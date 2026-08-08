import { describe, expect, test } from "bun:test";
import { shiftFormulaRefs } from "./refs.js";

describe("shiftFormulaRefs", () => {
  test("shifts relative refs like Excel copy-paste", () => {
    expect(shiftFormulaRefs("A1+1", 2, 1)).toBe("B3+1");
    expect(shiftFormulaRefs("$A1+B$2", 1, 1)).toBe("$A2+C$2");
    expect(shiftFormulaRefs("$A$1", 5, 5)).toBe("$A$1");
  });

  test("preserves sheet qualifiers while shifting the address", () => {
    expect(shiftFormulaRefs("Sheet2!A1+1", 1, 0)).toBe("Sheet2!A2+1");
    expect(shiftFormulaRefs("'My Sheet'!B2", 0, 1)).toBe("'My Sheet'!C2");
    expect(shiftFormulaRefs("Sheet2!$A$1", 3, 3)).toBe("Sheet2!$A$1");
  });

  test("off-grid relative refs become #REF!", () => {
    expect(shiftFormulaRefs("A1", 0, -1)).toBe("#REF!");
  });
});
