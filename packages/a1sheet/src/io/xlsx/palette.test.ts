/**
 * Theme colour resolution — the reason a themed workbook used to import almost
 * monochrome. Excel writes `<color theme="4" tint="-0.25"/>` for most colours,
 * and a reader that looks only at `rgb` finds nothing there.
 */
import { describe, expect, test } from "bun:test";
import {
  applyTint,
  EMPTY_PALETTE,
  mixColors,
  parseThemePalette,
  resolveColorAttrs,
} from "./palette.js";

/** A cut-down theme1.xml in the shape Excel writes: sysClr for the first pair. */
const THEME_XML = `<?xml version="1.0"?><a:theme xmlns:a="x"><a:themeElements>
<a:clrScheme name="Office">
<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
<a:dk2><a:srgbClr val="0E2841"/></a:dk2>
<a:lt2><a:srgbClr val="E8E8E8"/></a:lt2>
<a:accent1><a:srgbClr val="156082"/></a:accent1>
<a:accent2><a:srgbClr val="E97132"/></a:accent2>
<a:accent3><a:srgbClr val="196B24"/></a:accent3>
<a:accent4><a:srgbClr val="0F9ED5"/></a:accent4>
<a:accent5><a:srgbClr val="A02B93"/></a:accent5>
<a:accent6><a:srgbClr val="4EA72E"/></a:accent6>
<a:hlink><a:srgbClr val="467886"/></a:hlink>
<a:folHlink><a:srgbClr val="96607D"/></a:folHlink>
</a:clrScheme></a:themeElements></a:theme>`;

describe("the theme palette", () => {
  test("indexes light before dark, not the order the file lists", () => {
    // theme1.xml writes dk1 then lt1; Excel indexes them lt1 then dk1. Reading
    // them in file order inverts every document — black text on a black fill.
    const palette = parseThemePalette(THEME_XML);

    expect(palette.slots[0]).toBe("#ffffff");
    expect(palette.slots[1]).toBe("#000000");
    expect(palette.slots[2]).toBe("#e8e8e8");
    expect(palette.slots[3]).toBe("#0e2841");
  });

  test("reads a system colour from its cached last value", () => {
    // dk1 and lt1 are references to OS colours. `lastClr` is the value they last
    // resolved to, and the only sensible answer away from that OS.
    expect(parseThemePalette(THEME_XML).slots[1]).toBe("#000000");
  });

  test("accents land where the table-style names expect them", () => {
    const palette = parseThemePalette(THEME_XML);
    expect(palette.slots[4]).toBe("#156082");
    expect(palette.slots[6]).toBe("#196b24");
  });

  test("a missing theme part yields an empty palette rather than throwing", () => {
    expect(parseThemePalette(undefined).slots).toEqual([]);
    expect(parseThemePalette("<not-a-theme/>").slots).toEqual([]);
  });
});

describe("resolving a colour element", () => {
  const palette = parseThemePalette(THEME_XML);

  test("a literal ARGB drops its alpha", () => {
    expect(resolveColorAttrs({ rgb: "FF612890" }, palette)).toBe("#612890");
  });

  test("a theme index resolves through the palette", () => {
    expect(resolveColorAttrs({ theme: "6" }, palette)).toBe("#196b24");
  });

  test("a tint lightens toward white and darkens toward black", () => {
    const lighter = resolveColorAttrs({ theme: "6", tint: "0.5" }, palette);
    const darker = resolveColorAttrs({ theme: "6", tint: "-0.5" }, palette);
    expect(lighter).not.toBe("#196b24");
    expect(darker).not.toBe("#196b24");
    // Lightening must raise every channel and darkening must lower it.
    expect(Number.parseInt((lighter as string).slice(1, 3), 16)).toBeGreaterThan(
      0x19,
    );
    expect(Number.parseInt((darker as string).slice(1, 3), 16)).toBeLessThan(0x19);
  });

  test("the legacy indexed palette still resolves", () => {
    // LibreOffice and anything writing pre-2007 styles use these.
    expect(resolveColorAttrs({ indexed: "2" }, palette)).toBe("#ff0000");
  });

  test("auto and an unknown index mean inherit, not black", () => {
    // Defaulting to black is how a missing border colour becomes a black grid.
    expect(resolveColorAttrs({ auto: "1" }, palette)).toBeNull();
    expect(resolveColorAttrs({ theme: "99" }, palette)).toBeNull();
    expect(resolveColorAttrs({}, palette)).toBeNull();
  });

  test("a theme index with no palette resolves to nothing", () => {
    expect(resolveColorAttrs({ theme: "4" }, EMPTY_PALETTE)).toBeNull();
  });
});

describe("colour arithmetic", () => {
  test("a zero tint is the identity", () => {
    expect(applyTint("#123456", 0)).toBe("#123456");
  });

  test("a full tint reaches the extremes", () => {
    expect(applyTint("#808080", 1)).toBe("#ffffff");
    expect(applyTint("#808080", -1)).toBe("#000000");
  });

  test("mixing at the endpoints returns each side unchanged", () => {
    expect(mixColors("#ffffff", "#196b24", 0)).toBe("#ffffff");
    expect(mixColors("#ffffff", "#196b24", 1)).toBe("#196b24");
  });
});
