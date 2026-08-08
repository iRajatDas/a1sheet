import { describe, expect, test } from "bun:test";
import { resolveTheme, sheetsTheme } from "./theme.js";

describe("sheetsTheme", () => {
  test("uses blue accent and light chrome", () => {
    const t = resolveTheme(sheetsTheme);
    expect(t.accent).toBe("#1a73e8");
    expect(t.selectedBg).toBe("rgba(26, 115, 232, 0.12)");
    expect(t.headerBg).toBe("#f8f9fa");
    expect(t.cellBg).toBe("#ffffff");
    expect(t.cellText).toBe("#202124");
    expect(t.toolbarBg).toBe("#ffffff");
    expect(t.border).toBe("#e0e0e0");
    expect(t.headerBorder).toBe("#c0c0c0");
    expect(t.fontFamily.toLowerCase()).toContain("arial");
  });
});
