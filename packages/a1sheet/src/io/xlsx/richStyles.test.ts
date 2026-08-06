/**
 * The parts of a style Excel keeps somewhere other than the cell: named style
 * inheritance, differential formats, table styles, and conditional formatting.
 *
 * Each of these was a reason an imported workbook looked plainer than it did in
 * Excel while every individual cell read "correctly".
 */
import { describe, expect, test } from "bun:test";
import type { CellKey, StyleObject } from "../../model/types.js";
import { parseCondFormats } from "./condFormat.js";
import { parseThemePalette } from "./palette.js";
import { parseDifferentialStyle, parseStylesXml } from "./styles.js";
import { applyTableStyles, parseDxfs, parseTableXml } from "./tables.js";

const THEME = parseThemePalette(`<a:theme xmlns:a="x"><a:clrScheme name="Office">
<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
<a:dk2><a:srgbClr val="0E2841"/></a:dk2><a:lt2><a:srgbClr val="E8E8E8"/></a:lt2>
<a:accent1><a:srgbClr val="156082"/></a:accent1>
<a:accent2><a:srgbClr val="E97132"/></a:accent2>
<a:accent3><a:srgbClr val="196B24"/></a:accent3>
<a:accent4><a:srgbClr val="0F9ED5"/></a:accent4>
<a:accent5><a:srgbClr val="A02B93"/></a:accent5>
<a:accent6><a:srgbClr val="4EA72E"/></a:accent6>
</a:clrScheme></a:theme>`);

function styles(body: string): (StyleObject | null)[] {
  return parseStylesXml(`<styleSheet>${body}</styleSheet>`, THEME);
}

describe("a font", () => {
  test("carries its family and its size in pixels", () => {
    const [style] = styles(
      `<fonts><font><sz val="15"/><name val="Aptos Narrow"/></font></fonts>` +
        `<cellXfs><xf fontId="0"/></cellXfs>`,
    );

    expect(style?.fontFamily).toBe("Aptos Narrow");
    // 15pt at 96dpi is 20px. Points are the file's unit; the DOM wants pixels.
    expect(style?.fontSize).toBe(20);
  });

  test("takes its colour from the theme", () => {
    const [style] = styles(
      `<fonts><font><color theme="0"/></font></fonts><cellXfs><xf fontId="0"/></cellXfs>`,
    );
    expect(style?.color).toBe("#ffffff");
  });

  test("is not underlined by a differential format switching underline off", () => {
    // `<u val="none"/>` is how a dxf states "not underlined". Treating the tag's
    // presence as truth underlined every table header in the workbook.
    const style = parseDifferentialStyle(
      `<font><b/><i val="0"/><u val="none"/><color theme="0"/></font>`,
      THEME,
    );

    expect(style?.bold).toBe(true);
    expect(style?.italic).toBeUndefined();
    expect(style?.underline).toBeUndefined();
  });
});

describe("a fill", () => {
  test("reads a solid pattern's foreground", () => {
    const [style] = styles(
      `<fills><fill><patternFill patternType="solid"><fgColor rgb="FF612890"/></patternFill></fill></fills>` +
        `<cellXfs><xf fillId="0"/></cellXfs>`,
    );
    expect(style?.bg).toBe("#612890");
  });

  test("ignores the two placeholder fills every workbook carries", () => {
    // patternType="none" and "gray125" are structural, not colours. Reading them
    // as fills would give every cell in the file a background.
    const parsed = styles(
      `<fills><fill><patternFill patternType="none"/></fill>` +
        `<fill><patternFill patternType="gray125"/></fill></fills>` +
        `<cellXfs><xf fillId="0"/><xf fillId="1"/></cellXfs>`,
    );
    expect(parsed[0]).toBeNull();
    expect(parsed[1]?.bg).toBeUndefined();
  });

  test("reads a gradient as stops, keeping the first as a flat fallback", () => {
    // The sample workbook's purple headings are a gradient, not a solid fill.
    const style = parseDifferentialStyle(
      `<fill><gradientFill degree="90">` +
        `<stop position="0"><color rgb="FF3D195B"/></stop>` +
        `<stop position="1"><color rgb="FF612890"/></stop>` +
        `</gradientFill></fill>`,
      THEME,
    );

    expect(style?.gradient?.degree).toBe(90);
    expect(style?.gradient?.stops).toEqual([
      { position: 0, color: "#3d195b" },
      { position: 1, color: "#612890" },
    ]);
    expect(style?.bg).toBe("#3d195b");
  });
});

describe("a border", () => {
  test("reads each edge's line and colour", () => {
    const [style] = styles(
      `<borders><border><left/><right/>` +
        `<top style="thin"><color rgb="FF95CA82"/></top>` +
        `<bottom style="thick"><color rgb="FF7030A0"/></bottom>` +
        `</border></borders><cellXfs><xf borderId="0"/></cellXfs>`,
    );

    expect(style?.borders).toEqual({
      top: { line: "thin", color: "#95ca82" },
      bottom: { line: "thick", color: "#7030a0" },
    });
  });

  test("an auto colour leaves the edge to inherit", () => {
    // Defaulting an unspecified border colour to black draws a black grid.
    const [style] = styles(
      `<borders><border><bottom style="thin"><color auto="1"/></bottom></border></borders>` +
        `<cellXfs><xf borderId="0"/></cellXfs>`,
    );
    expect(style?.borders?.bottom).toEqual({ line: "thin" });
  });
});

describe("a named cell style", () => {
  test("supplies what the cell's own xf leaves at zero", () => {
    // Excel usually omits the applyBorder/applyAlignment flags that are meant to
    // signal this, so inheritance has to be inferred from a zero id. Without it,
    // titles and headings lose the borders and centring their style carries.
    const [style] = styles(
      `<fonts><font/><font><b/></font></fonts>` +
        `<borders><border/><border><bottom style="thick"><color rgb="FF7030A0"/></bottom></border></borders>` +
        `<cellStyleXfs><xf fontId="0" borderId="0"/>` +
        `<xf fontId="1" borderId="1"><alignment horizontal="center"/></xf></cellStyleXfs>` +
        `<cellXfs><xf fontId="0" fillId="0" borderId="0" xfId="1"/></cellXfs>`,
    );

    expect(style?.borders?.bottom?.line).toBe("thick");
    expect(style?.align).toBe("center");
  });

  test("does not override what the cell's own xf states", () => {
    const [style] = styles(
      `<fonts><font/><font><b/></font><font><i/></font></fonts>` +
        `<cellStyleXfs><xf fontId="0"/><xf fontId="1"/></cellStyleXfs>` +
        `<cellXfs><xf fontId="2" xfId="1"/></cellXfs>`,
    );

    expect(style?.italic).toBe(true);
    expect(style?.bold).toBeUndefined();
  });
});

describe("a number format", () => {
  test("keeps the file's literal code alongside the bucket", () => {
    const [style] = styles(
      `<numFmts><numFmt numFmtId="164" formatCode="\\+0;\\-0;0"/></numFmts>` +
        `<cellXfs><xf numFmtId="164"/></cellXfs>`,
    );
    expect(style?.numFmtCode).toBe("\\+0;\\-0;0");
  });

  test("supplies a built-in id's code, which the file never states", () => {
    const [style] = styles(`<cellXfs><xf numFmtId="22"/></cellXfs>`);
    expect(style?.numFmtCode).toBe("m/d/yy h:mm");
    expect(style?.numFmt).toBe("date");
  });

  test("a differential format's General does not format every plain cell", () => {
    // <dxfs> holds <numFmt> elements of its own. Scanning the whole document for
    // them picked up numFmtId="0" as "General" and applied it workbook-wide.
    const [style] = styles(
      `<dxfs><dxf><numFmt numFmtId="0" formatCode="General"/></dxf></dxfs>` +
        `<cellXfs><xf numFmtId="0"/></cellXfs>`,
    );
    expect(style).toBeNull();
  });
});

describe("a table", () => {
  const TABLE_XML =
    `<table ref="A1:G381" headerRowDxfId="1" dataDxfId="0">` +
    `<tableStyleInfo name="TableStyleMedium4" showRowStripes="1" showColumnStripes="0"/>` +
    `</table>`;

  test("reads its range, header, and banding", () => {
    const table = parseTableXml(TABLE_XML);

    expect(table?.range).toEqual({ r1: 0, c1: 0, r2: 380, c2: 6 });
    expect(table?.headerRow).toBe(true);
    expect(table?.bandedRows).toBe(true);
    expect(table?.bandedCols).toBe(false);
    expect(table?.styleName).toBe("TableStyleMedium4");
  });

  test("a built-in style name resolves to a theme accent", () => {
    // TableStyleMedium4 is accent3 — the green header in the sample workbook,
    // a colour that appears nowhere in the file.
    const table = parseTableXml(TABLE_XML);
    const cellStyles: Record<CellKey, StyleObject> = {};
    applyTableStyles({
      tables: table ? [table] : [],
      dxfs: [],
      palette: THEME,
      styles: cellStyles,
    });

    expect(cellStyles["0_0" as CellKey]?.bg).toBe("#196b24");
    expect(cellStyles["0_0" as CellKey]?.bold).toBe(true);
  });

  test("banding starts from the first data row and alternates", () => {
    const table = parseTableXml(TABLE_XML);
    const cellStyles: Record<CellKey, StyleObject> = {};
    applyTableStyles({
      tables: table ? [table] : [],
      dxfs: [],
      palette: THEME,
      styles: cellStyles,
    });

    expect(cellStyles["1_0" as CellKey]?.bg).toBeUndefined();
    expect(cellStyles["2_0" as CellKey]?.bg).toBeDefined();
    expect(cellStyles["3_0" as CellKey]?.bg).toBeUndefined();
  });

  test("a cell's own formatting wins over its table's", () => {
    // Table styling is the least specific layer; anything set directly beats it.
    const table = parseTableXml(TABLE_XML);
    const cellStyles: Record<CellKey, StyleObject> = {
      ["0_0" as CellKey]: { bg: "#ff0000" },
    };
    applyTableStyles({
      tables: table ? [table] : [],
      dxfs: [],
      palette: THEME,
      styles: cellStyles,
    });

    expect(cellStyles["0_0" as CellKey]?.bg).toBe("#ff0000");
    // …while still picking up what the cell did not state.
    expect(cellStyles["0_0" as CellKey]?.bold).toBe(true);
  });

  test("a dxf the file states is applied over the built-in recipe", () => {
    const dxfs = parseDxfs(
      `<styleSheet><dxfs><dxf/><dxf><font><color rgb="FF00FF00"/></font></dxf></dxfs></styleSheet>`,
      THEME,
      parseDifferentialStyle,
    );
    const table = parseTableXml(TABLE_XML);
    const cellStyles: Record<CellKey, StyleObject> = {};
    applyTableStyles({
      tables: table ? [table] : [],
      dxfs,
      palette: THEME,
      styles: cellStyles,
    });

    expect(cellStyles["0_0" as CellKey]?.color).toBe("#00ff00");
  });
});

describe("conditional formatting", () => {
  const dxfs = parseDxfs(
    `<styleSheet><dxfs>` +
      `<dxf><fill><patternFill patternType="solid"><fgColor rgb="FF612890"/></patternFill></fill></dxf>` +
      `<dxf><font><b/></font></dxf>` +
      `</dxfs></styleSheet>`,
    THEME,
    parseDifferentialStyle,
  );

  test("an expression rule keeps its range, formula, and style", () => {
    const formats = parseCondFormats({
      sheetXml:
        `<conditionalFormatting sqref="C4:M4">` +
        `<cfRule type="expression" dxfId="0" priority="1">` +
        `<formula>ISBLANK($C$4)=FALSE</formula></cfRule></conditionalFormatting>`,
      dxfs,
    });

    expect(formats).toHaveLength(1);
    expect(formats[0]?.range).toEqual({ r1: 3, c1: 2, r2: 3, c2: 12 });
    expect(formats[0]?.rule).toEqual({
      type: "expression",
      formula: "ISBLANK($C$4)=FALSE",
    });
    expect(formats[0]?.style.bg).toBe("#612890");
    expect(formats[0]?.priority).toBe(1);
  });

  test("one rule over several ranges becomes one format per range", () => {
    const formats = parseCondFormats({
      sheetXml:
        `<conditionalFormatting sqref="A1:A2 C1:C2">` +
        `<cfRule type="expression" dxfId="1" priority="1"><formula>1</formula></cfRule>` +
        `</conditionalFormatting>`,
      dxfs,
    });
    expect(formats).toHaveLength(2);
  });

  test("a cellIs rule keeps its operator and operands", () => {
    const formats = parseCondFormats({
      sheetXml:
        `<conditionalFormatting sqref="A1:A9">` +
        `<cfRule type="cellIs" operator="between" dxfId="1" priority="2">` +
        `<formula>1</formula><formula>10</formula></cfRule></conditionalFormatting>`,
      dxfs,
    });

    expect(formats[0]?.rule).toEqual({
      type: "cellIs",
      operator: "between",
      operands: ["1", "10"],
    });
  });

  test("a graphical rule is dropped rather than half-applied", () => {
    // A colour scale or data bar is a drawing, not a style. There is nothing to
    // return for it, and failing on it would reject the whole file.
    const formats = parseCondFormats({
      sheetXml:
        `<conditionalFormatting sqref="A1:A9">` +
        `<cfRule type="colorScale" priority="1"><colorScale/></cfRule>` +
        `</conditionalFormatting>`,
      dxfs,
    });
    expect(formats).toEqual([]);
  });

  test("a rule pointing at a dxf that is not there is dropped", () => {
    const formats = parseCondFormats({
      sheetXml:
        `<conditionalFormatting sqref="A1"><cfRule type="expression" dxfId="99"` +
        ` priority="1"><formula>1</formula></cfRule></conditionalFormatting>`,
      dxfs,
    });
    expect(formats).toEqual([]);
  });
});
