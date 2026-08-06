/**
 * The style-to-CSS mapping. Pure, so it is tested directly: a border weight or a
 * gradient angle is easy to get wrong and hard to notice in a render test, which
 * would only prove that something was set.
 */
import { describe, expect, test } from "bun:test";
import { cellCss } from "./cellStyle.js";
import { defaultTheme, resolveTheme } from "./theme.js";

describe("borders", () => {
  test("each Excel line kind maps to a CSS width and style", () => {
    expect(
      cellCss({ borders: { top: { line: "thin" } } }, defaultTheme).borderTop,
    ).toContain("1px solid");
    expect(
      cellCss({ borders: { top: { line: "medium" } } }, defaultTheme).borderTop,
    ).toContain("2px solid");
    expect(
      cellCss({ borders: { top: { line: "thick" } } }, defaultTheme).borderTop,
    ).toContain("3px solid");
    expect(
      cellCss({ borders: { top: { line: "double" } } }, defaultTheme).borderTop,
    ).toContain("double");
    expect(
      cellCss({ borders: { top: { line: "dashed" } } }, defaultTheme).borderTop,
    ).toContain("dashed");
  });

  test("an edge with no colour of its own takes the text colour", () => {
    // Excel's `auto`. Falling back to black instead draws a black grid over a
    // dark theme.
    const css = cellCss(
      { color: "#ff0000", borders: { bottom: { line: "thin" } } },
      defaultTheme,
    );
    expect(css.borderBottom).toBe("1px solid #ff0000");
  });

  test("only the edges the style sets are written", () => {
    // The stylesheet already draws the grid line on the right and bottom of every
    // cell. Writing all four edges unconditionally doubles every internal line.
    const css = cellCss({ borders: { bottom: { line: "thin" } } }, defaultTheme);
    expect(css.borderTop).toBeUndefined();
    expect(css.borderLeft).toBeUndefined();
    expect(css.borderRight).toBeUndefined();
    expect(css.borderBottom).toBeDefined();
  });
});

describe("fills", () => {
  test("a solid fill becomes the background", () => {
    expect(cellCss({ bg: "#612890" }, defaultTheme).background).toBe("#612890");
  });

  test("a gradient becomes a linear-gradient with its stops in order", () => {
    const css = cellCss(
      {
        gradient: {
          degree: 90,
          stops: [
            { position: 1, color: "#612890" },
            { position: 0, color: "#3d195b" },
          ],
        },
      },
      defaultTheme,
    );

    // OOXML measures the sweep; CSS measures where the sweep points. A 90-degree
    // OOXML gradient runs top to bottom, which CSS calls 180deg.
    expect(css.background).toBe(
      "linear-gradient(180deg, #3d195b 0%, #612890 100%)",
    );
  });

  test("a gradient wins over a flat fill", () => {
    const css = cellCss(
      {
        bg: "#3d195b",
        gradient: { degree: 0, stops: [{ position: 0, color: "#ff0000" }] },
      },
      defaultTheme,
    );
    expect(css.background).toContain("linear-gradient");
  });

  test("no fill falls back to the theme, never to transparent", () => {
    // A frozen cell is sticky, and a transparent one shows the scrolled content
    // through it.
    expect(cellCss({}, defaultTheme).background).toBe(defaultTheme.cellBg);
  });
});

describe("fonts", () => {
  test("a family from a file keeps the theme stack behind it", () => {
    // A file names one family. Without a fallback, a reader that lacks it drops to
    // the browser default rather than to the sheet's own face.
    const theme = resolveTheme({ fontFamily: "BodyFace" });
    const css = cellCss({ fontFamily: "Aptos Narrow" }, theme);

    expect(css.fontFamily).toBe('"Aptos Narrow", BodyFace');
  });

  test("a single-word family needs no quoting", () => {
    const theme = resolveTheme({ fontFamily: "BodyFace" });
    expect(cellCss({ fontFamily: "Calibri" }, theme).fontFamily).toBe(
      "Calibri, BodyFace",
    );
  });

  test("a size is passed through in pixels", () => {
    expect(cellCss({ fontSize: 20 }, defaultTheme).fontSize).toBe(20);
  });
});

describe("alignment", () => {
  test("horizontal alignment sets both the text and the flex axis", () => {
    // The cell is a flex container, so textAlign alone does not move its content.
    const css = cellCss({ align: "right" }, defaultTheme);
    expect(css.textAlign).toBe("right");
    expect(css.justifyContent).toBe("flex-end");
  });

  test("no alignment leaves the default, so numbers and text can differ", () => {
    // Excel's "general" is left for text and right for numbers, which a style
    // cannot decide on its own.
    const css = cellCss({}, defaultTheme);
    expect(css.textAlign).toBeUndefined();
    expect(css.justifyContent).toBeUndefined();
  });

  test("vertical alignment maps onto the cross axis", () => {
    expect(cellCss({ valign: "top" }, defaultTheme).alignItems).toBe("flex-start");
    expect(cellCss({ valign: "bottom" }, defaultTheme).alignItems).toBe("flex-end");
  });

  test("wrapping opts out of the stylesheet's nowrap", () => {
    const css = cellCss({ wrap: true }, defaultTheme);
    expect(css.whiteSpace).toBe("normal");
  });
});
