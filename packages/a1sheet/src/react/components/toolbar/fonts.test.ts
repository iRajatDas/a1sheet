import { describe, expect, test } from "bun:test";
import { fontById, noteRecentFont, recentFonts, SHEET_FONTS } from "./fonts.js";

describe("toolbar fonts", () => {
  test("lists built-in faces without duplicates", () => {
    const ids = SHEET_FONTS.map((f) => f.id);
    expect(new Set(ids).size).toBe(ids.length);
  });

  test("tracks recent picks in order", () => {
    noteRecentFont("Georgia");
    noteRecentFont("Arial");
    noteRecentFont("Georgia");
    expect(recentFonts().map((f) => f.id)).toEqual(["Georgia", "Arial"]);
  });

  test("resolves unknown ids to theme default", () => {
    expect(fontById("NoSuchFont").label).toBe("Default (Theme)");
  });
});
