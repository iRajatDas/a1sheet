/**
 * The icon set, as an invariant rather than a convention.
 *
 * The controls used to be typographic glyphs and one emoji. Both look fine on
 * the machine they were written on and wrong elsewhere: a glyph resolves
 * through the host page's font stack, and an emoji renders as a colour bitmap
 * that ignores `color`, so the lock stayed black on an active teal button.
 * Nothing stops the next contributor from reaching for `✂` — except this.
 */
import { describe, expect, test } from "bun:test";
import { render } from "@testing-library/react";
import { Spreadsheet } from "./Spreadsheet.js";

/**
 * Emoji, dingbats, arrows, geometric shapes, and the miscellaneous-symbol
 * blocks — the ranges a decorative glyph comes from. Ordinary punctuation is
 * deliberately outside this: an em dash in a status message is prose.
 */
const DECORATIVE_GLYPH = new RegExp(
  [
    "[",
    "\\u2190-\\u21FF", // arrows
    "\\u2300-\\u23FF", // miscellaneous technical
    "\\u25A0-\\u25FF", // geometric shapes
    "\\u2600-\\u26FF", // miscellaneous symbols
    "\\u2700-\\u27BF", // dingbats
    "\\u2B00-\\u2BFF", // supplemental arrows and shapes
    "\\u{1F000}-\\u{1FAFF}", // emoji planes
    "\\u{FE0F}", // variation selector 16
    "]",
  ].join(""),
  "u",
);

function textNodesOf(root: Element): string[] {
  const found: string[] = [];
  const walk = (node: Node) => {
    if (node.nodeType === Node.TEXT_NODE) {
      const text = node.textContent ?? "";
      if (text.trim()) found.push(text);
      return;
    }
    for (const child of node.childNodes) walk(child);
  };
  walk(root);
  return found;
}

describe("the built-in primitives draw with icons, not glyphs", () => {
  test("no rendered text contains a decorative glyph or emoji", () => {
    const { container } = render(<Spreadsheet />);

    const offenders = textNodesOf(container).filter((text) =>
      DECORATIVE_GLYPH.test(text),
    );

    expect(offenders).toEqual([]);
  });

  test("every icon is an inline SVG that inherits the button's color", () => {
    const { container } = render(<Spreadsheet />);
    const svgs = [...container.querySelectorAll("svg")];

    expect(svgs.length).toBeGreaterThan(0);
    for (const svg of svgs) {
      // currentColor is the whole point: an active button flips to white text
      // and the icon has to come with it.
      expect(svg.getAttribute("stroke")).toBe("currentColor");
      // The button carries the accessible name; a duplicate on the icon would
      // make a screen reader read it twice.
      expect(svg.getAttribute("aria-hidden")).toBe("true");
    }
  });

  test("every button still has an accessible name without its text", () => {
    const { container } = render(<Spreadsheet />);
    const buttons = [...container.querySelectorAll("button")];

    expect(buttons.length).toBeGreaterThan(0);
    const unnamed = buttons.filter(
      (b) => !(b.getAttribute("aria-label") ?? b.textContent ?? "").trim(),
    );

    expect(unnamed.map((b) => b.outerHTML)).toEqual([]);
  });
});
