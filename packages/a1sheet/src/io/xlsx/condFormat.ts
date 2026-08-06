/**
 * Reading `<conditionalFormatting>` out of a worksheet.
 *
 * Easy to dismiss as a nice-to-have, and then a real workbook turns out to use it
 * for its headings: the purple title bars in the sample football workbook are a
 * `type="expression"` rule with the formula `ISBLANK($C$4)=FALSE`, not cell
 * formatting at all. A reader that skips this shows those rows unstyled.
 *
 * The style comes from `<dxfs>`, indexed by `dxfId`, which the caller has already
 * parsed — this module only decides which cells a rule covers and what test it
 * applies.
 */
import { lettersToCol } from "../../model/address.js";
import type {
  CondFormat,
  CondOperator,
  CondRule,
  Range,
  StyleObject,
} from "../../model/types.js";
import { findElements, textOf } from "./xml.js";

const REF_RE = /^\$?([A-Za-z]+)\$?(\d+)$/;

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
  return {
    r1: Math.min(a.row, b.row),
    c1: Math.min(a.col, b.col),
    r2: Math.max(a.row, b.row),
    c2: Math.max(a.col, b.col),
  };
}

const OPERATORS: readonly CondOperator[] = [
  "lessThan",
  "lessThanOrEqual",
  "equal",
  "notEqual",
  "greaterThanOrEqual",
  "greaterThan",
  "between",
  "notBetween",
];

/** Rules whose effect is a drawing rather than a style. Read, then dropped. */
const GRAPHICAL_TYPES = new Set(["colorScale", "dataBar", "iconSet"]);

function parseRule(
  type: string | undefined,
  attrs: Record<string, string>,
  formulas: readonly string[],
): CondRule | null {
  if (!type || GRAPHICAL_TYPES.has(type)) return null;

  if (type === "expression") {
    const formula = formulas[0];
    return formula ? { type: "expression", formula } : null;
  }

  if (type === "cellIs") {
    const operator = OPERATORS.find((o) => o === attrs.operator);
    if (!operator || formulas.length === 0) return null;
    return { type: "cellIs", operator, operands: formulas };
  }

  if (type === "containsText" || type === "notContainsText") {
    const text = attrs.text;
    return text === undefined
      ? null
      : { type: "containsText", text, negate: type === "notContainsText" };
  }

  if (type === "containsBlanks" || type === "notContainsBlanks") {
    return { type: "containsBlanks", negate: type === "notContainsBlanks" };
  }

  return null;
}

/** Excel's default when a rule omits `priority`; every real file states one. */
const LOWEST_PRIORITY = Number.MAX_SAFE_INTEGER;

export interface ParseCondFormatsOptions {
  sheetXml: string;
  /** Differential formats from styles.xml, indexed by `dxfId`. */
  dxfs: readonly (StyleObject | null)[];
}

export function parseCondFormats(opts: ParseCondFormatsOptions): CondFormat[] {
  const { sheetXml, dxfs } = opts;
  const out: CondFormat[] = [];

  for (const block of findElements(sheetXml, "conditionalFormatting")) {
    const sqref = block.attrs.sqref;
    if (!sqref) continue;

    // sqref is a space-separated list of ranges; one rule can cover several.
    const ranges = sqref
      .split(/\s+/)
      .map(parseRange)
      .filter((r): r is Range => r !== null);
    if (ranges.length === 0) continue;

    for (const cfRule of findElements(block.inner, "cfRule")) {
      const formulas = findElements(cfRule.inner, "formula").map((f) =>
        textOf(f.inner),
      );
      const rule = parseRule(cfRule.attrs.type, cfRule.attrs, formulas);
      if (!rule) continue;

      const dxfId = Number.parseInt(cfRule.attrs.dxfId ?? "", 10);
      const style = Number.isFinite(dxfId) ? dxfs[dxfId] : null;
      // A rule with no style has nothing to contribute — that is a graphical
      // rule we dropped, or a file referencing a dxf that is not there.
      if (!style) continue;

      const priority = Number.parseInt(cfRule.attrs.priority ?? "", 10);
      for (const range of ranges) {
        out.push({
          range,
          priority: Number.isFinite(priority) ? priority : LOWEST_PRIORITY,
          rule,
          style,
          ...(cfRule.attrs.stopIfTrue === "1" ? { stopIfTrue: true } : {}),
        });
      }
    }
  }

  return out;
}
