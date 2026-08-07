/**
 * Table styling: `xl/tables/*.xml` plus the `<dxfs>` block of styles.xml.
 *
 * A table's appearance does not live in the cells it covers. `Format as Table`
 * records a range and a style name, and Excel paints the header, the banding, and
 * the borders from that — so a table imported by cell style alone arrives
 * completely unstyled even though every cell looked coloured in Excel.
 *
 * Two mechanisms, and a file uses both:
 *
 *   - **Differential formats.** `headerRowDxfId="26"` indexes `<dxfs>`, a list of
 *     partial styles. These are in the file, so they can be read exactly.
 *   - **Built-in style names.** `TableStyleMedium4` is not in the file at all;
 *     every reader is expected to know the recipe. The colour comes from the
 *     workbook theme's accents, which is why the name carries a number.
 *
 * The result is flattened onto the cells at import: a table becomes ordinary cell
 * styles covering its range. That loses the live "it is a table" relationship —
 * adding a row will not extend the banding — and gains the whole of the rest of
 * the library working on it unchanged. The alternative is a parallel styling
 * layer resolved on every render, for a feature nothing else in the model knows
 * about.
 */

import { lettersToCol } from "../../model/address.js";
import type { CellKey, HexColor, Range, StyleObject } from "../../model/types.js";
import { applyTint, mixColors, type ThemePalette } from "./palette.js";
import { findElement, findElements } from "./xml.js";

/** One table, as the file describes it. */
export interface XlsxTable {
  /** The name formulas use, which is `name` not `displayName`. */
  name: string;
  range: Range;
  /** Column names in order, from `<tableColumn>`. */
  columns: readonly string[];
  /** True when the first row of `range` holds column names. */
  headerRow: boolean;
  /** Built-in style name, e.g. `"TableStyleMedium4"`. Absent for no style. */
  styleName?: string;
  bandedRows: boolean;
  bandedCols: boolean;
  /** Indices into `<dxfs>`, where the file states the formatting explicitly. */
  headerDxf?: number;
  dataDxf?: number;
}

const REF_RE = /^([A-Za-z]+)(\d+)$/;

function parseRef(ref: string): { row: number; col: number } | null {
  const m = ref.match(REF_RE);
  if (!m?.[1] || !m[2]) return null;
  return {
    col: lettersToCol(m[1].toUpperCase()),
    row: Number.parseInt(m[2], 10) - 1,
  };
}

function parseRange(ref: string): Range | null {
  const [from, to] = ref.split(":");
  if (!from) return null;
  const a = parseRef(from);
  if (!a) return null;
  const b = to ? parseRef(to) : a;
  if (!b) return null;
  return { r1: a.row, c1: a.col, r2: b.row, c2: b.col };
}

function optionalIndex(value: string | undefined): number | undefined {
  if (value === undefined) return undefined;
  const n = Number.parseInt(value, 10);
  return Number.isFinite(n) ? n : undefined;
}

/** Reads one `xl/tables/tableN.xml`. */
export function parseTableXml(xml: string): XlsxTable | null {
  const table = findElement(xml, "table");
  if (!table?.attrs.ref) return null;
  const range = parseRange(table.attrs.ref);
  if (!range) return null;

  const info = findElement(xml, "tableStyleInfo");
  const styleName = info?.attrs.name;

  return {
    name: table.attrs.name ?? table.attrs.displayName ?? "",
    range,
    columns: findElements(xml, "tableColumn").map(
      (c, i) => c.attrs.name ?? `Column${i + 1}`,
    ),
    // headerRowCount defaults to 1; only an explicit "0" means there is none.
    headerRow: table.attrs.headerRowCount !== "0",
    ...(styleName ? { styleName } : {}),
    bandedRows: info?.attrs.showRowStripes === "1",
    bandedCols: info?.attrs.showColumnStripes === "1",
    ...(optionalIndex(table.attrs.headerRowDxfId) !== undefined
      ? { headerDxf: optionalIndex(table.attrs.headerRowDxfId) as number }
      : {}),
    ...(optionalIndex(table.attrs.dataDxfId) !== undefined
      ? { dataDxf: optionalIndex(table.attrs.dataDxfId) as number }
      : {}),
  };
}

/**
 * The accent a built-in table style is built from.
 *
 * `TableStyleLight1`, `Medium1`, and `Dark1` are the accent-free variants. From 2
 * upward the names cycle through the six accents, so Medium4 is accent3 and
 * Medium5 is accent4 — the green and blue that a workbook's two tables show
 * without either colour appearing anywhere in the file.
 */
const ACCENT_COUNT = 6;
/** Index of accent1 in the theme palette. See THEME_SLOTS in palette.ts. */
const FIRST_ACCENT_SLOT = 4;

type StyleFamily = "light" | "medium" | "dark";

interface BuiltinStyle {
  family: StyleFamily;
  accent: HexColor | undefined;
}

const BUILTIN_NAME_RE = /^TableStyle(Light|Medium|Dark)(\d+)$/;

function parseBuiltinName(
  name: string,
  palette: ThemePalette,
): BuiltinStyle | null {
  const m = name.match(BUILTIN_NAME_RE);
  if (!m?.[1] || !m[2]) return null;
  const family = m[1].toLowerCase() as StyleFamily;
  const n = Number.parseInt(m[2], 10);
  if (n <= 1) return { family, accent: undefined };
  const accentIndex = (n - 2) % ACCENT_COUNT;
  return { family, accent: palette.slots[FIRST_ACCENT_SLOT + accentIndex] };
}

/** How much of the accent a banded row keeps. Excel's stripe is a pale wash. */
const BAND_WEIGHT = 0.2;
/** Light styles tint their header instead of filling it. */
const LIGHT_HEADER_TINT = 0.6;
/** Dark styles deepen the accent for the header. */
const DARK_HEADER_TINT = -0.25;
const WHITE = "#ffffff" as HexColor;

interface BuiltinRecipe {
  header: StyleObject;
  band: StyleObject | null;
}

/**
 * The recipe for a built-in style.
 *
 * Approximate, and deliberately so: Excel's real definitions are a table of
 * several hundred entries in its own binary resources, differing in border weight
 * and stripe opacity per family. What matters visually is that the header is the
 * accent colour and the body is striped with a wash of it, which is what this
 * reproduces. Documented in docs/LIMITATIONS.md.
 */
function builtinRecipe(style: BuiltinStyle): BuiltinRecipe {
  const accent = style.accent;
  if (!accent) {
    return {
      header: { bold: true, borders: { bottom: { line: "medium" } } },
      band: null,
    };
  }

  if (style.family === "light") {
    return {
      header: {
        bold: true,
        borders: { bottom: { line: "medium", color: accent } },
      },
      band: { bg: applyTint(accent, LIGHT_HEADER_TINT) },
    };
  }

  const headerBg =
    style.family === "dark" ? applyTint(accent, DARK_HEADER_TINT) : accent;
  return {
    header: { bold: true, color: WHITE, bg: headerBg },
    band: { bg: mixColors(WHITE, accent, BAND_WEIGHT) },
  };
}

/** A `<dxf>` entry: a partial style, read with the same rules as a full one. */
export function parseDxfs(
  stylesXml: string | undefined,
  palette: ThemePalette,
  parseOne: (inner: string, palette: ThemePalette) => StyleObject | null,
): (StyleObject | null)[] {
  if (!stylesXml) return [];
  const block = findElement(stylesXml, "dxfs");
  if (!block) return [];
  return findElements(block.inner, "dxf").map((dxf) =>
    parseOne(dxf.inner, palette),
  );
}

export interface ApplyTablesOptions {
  tables: readonly XlsxTable[];
  dxfs: readonly (StyleObject | null)[];
  palette: ThemePalette;
  /** Cell styles read from `cellXfs`, which a table's own styling sits under. */
  styles: Record<CellKey, StyleObject>;
}

/**
 * Writes each table's styling into the cell style map.
 *
 * Layered under whatever the cell already had: a cell's own formatting is more
 * specific than its table's, and in Excel it wins. So the table supplies the
 * fill and the header weight, and anything the user set directly overrides it.
 */
export function applyTableStyles(opts: ApplyTablesOptions): void {
  const { tables, dxfs, palette, styles } = opts;

  for (const table of tables) {
    const builtin = table.styleName
      ? parseBuiltinName(table.styleName, palette)
      : null;
    const recipe = builtin ? builtinRecipe(builtin) : null;

    const headerFromDxf =
      table.headerDxf === undefined ? null : (dxfs[table.headerDxf] ?? null);
    const dataFromDxf =
      table.dataDxf === undefined ? null : (dxfs[table.dataDxf] ?? null);

    const { r1, c1, r2, c2 } = table.range;
    const firstDataRow = table.headerRow ? r1 + 1 : r1;

    for (let row = r1; row <= r2; row++) {
      const isHeader = table.headerRow && row === r1;
      // The stripe alternates from the first data row, so which rows are shaded
      // does not shift when a table has no header.
      const banded = table.bandedRows && (row - firstDataRow) % 2 === 1;

      for (let col = c1; col <= c2; col++) {
        const layers: (StyleObject | null)[] = isHeader
          ? [recipe?.header ?? null, headerFromDxf]
          : [dataFromDxf, banded ? (recipe?.band ?? null) : null];

        const beneath = layers.filter((l): l is StyleObject => l !== null);
        if (beneath.length === 0) continue;

        const key = `${row}_${col}` as CellKey;
        const own = styles[key];
        styles[key] = Object.assign({}, ...beneath, own ?? {});
      }
    }
  }
}
