import { describe, expect, test } from "bun:test";
import { tokenize } from "./tokenize.js";

describe("refs vs names", () => {
  test("SUM stays a name, A1 becomes a ref", () => {
    expect(tokenize("SUM(A1:A3)")).toEqual([
      { type: "name", value: "SUM" },
      { type: "(" },
      { type: "ref", value: "A1", colAbs: false, rowAbs: false },
      { type: ":" },
      { type: "ref", value: "A3", colAbs: false, rowAbs: false },
      { type: ")" },
    ]);
  });

  test("tracks $ markers per axis", () => {
    expect(tokenize("$A$1")).toEqual([
      { type: "ref", value: "A1", colAbs: true, rowAbs: true },
    ]);
    expect(tokenize("$A1")).toEqual([
      { type: "ref", value: "A1", colAbs: true, rowAbs: false },
    ]);
    expect(tokenize("A$1")).toEqual([
      { type: "ref", value: "A1", colAbs: false, rowAbs: true },
    ]);
  });

  test("uppercases refs and names", () => {
    expect(tokenize("sum(a1)")).toEqual([
      { type: "name", value: "SUM" },
      { type: "(" },
      { type: "ref", value: "A1", colAbs: false, rowAbs: false },
      { type: ")" },
    ]);
  });

  test("does not split an identifier that merely starts like a ref", () => {
    // The negative lookahead in REF_RE is what prevents ref "A1" + name "B".
    expect(tokenize("A1B")).toEqual([{ type: "name", value: "A1B" }]);
  });
});

describe("comparison operators", () => {
  test.each(["=", "<", ">", "<=", ">=", "<>"] as const)(
    "lexes %s as a single cmp token",
    (op) => {
      expect(tokenize(`1${op}2`)).toEqual([
        { type: "num", value: 1 },
        { type: "cmp", value: op },
        { type: "num", value: 2 },
      ]);
    },
  );

  test("<> is one token, not < followed by > (regression)", () => {
    // Appending only "=" to a comparison operator lexes "<>" as two tokens, and
    // parsePrimary swallows the ">" as a literal 0.
    const tokens = tokenize("A1<>B1");
    expect(tokens.filter((t) => t.type === "cmp")).toHaveLength(1);
    expect(tokens[1]).toEqual({ type: "cmp", value: "<>" });
  });
});

describe("literals", () => {
  test("reads numbers including decimals", () => {
    expect(tokenize("3.14")).toEqual([{ type: "num", value: 3.14 }]);
  });

  test("reads quoted strings", () => {
    expect(tokenize('"hi there"')).toEqual([{ type: "str", value: "hi there" }]);
  });

  test("skips whitespace", () => {
    expect(tokenize("  1  +  2  ")).toHaveLength(3);
  });

  test("skips unrecognized characters instead of failing", () => {
    expect(tokenize("1 @ 2")).toEqual([
      { type: "num", value: 1 },
      { type: "num", value: 2 },
    ]);
  });
});
