/**
 * Autocomplete metadata for the formula library.
 *
 * Categories and argument-type names match Sheets' help content
 * (`m=ritzfunctionhelpcontent.js`). This is intentionally incomplete relative
 * to every Excel function — only entries we evaluate (or special-case) are
 * listed, so a consumer can drive suggestions without promising `#NAME?`-free
 * coverage of the whole catalog.
 */
import { FUNCTIONS } from "./functions/registry.js";

export const FORMULA_CATEGORIES = [
  "DATE",
  "ENGINEERING",
  "FILTER",
  "FINANCIAL",
  "GOOGLE",
  "INFO",
  "LOGICAL",
  "LOOKUP",
  "MATH",
  "OPERATOR",
  "STATISTICAL",
  "TEXT",
  "DATABASE",
  "PARSER",
  "ARRAY",
  "WEB",
] as const;

export type FormulaCategory = (typeof FORMULA_CATEGORIES)[number];

export const FORMULA_ARG_TYPES = [
  "NUMBER",
  "PERCENT",
  "DATE",
  "TIME",
  "BOOLEAN",
  "SPARKLINE_SECOND_ARGUMENT",
  "RANGE",
  "CANONICAL_FUNCTION_NAME",
  "FORMULA",
  "LITERAL",
  "TRANSLATED",
] as const;

export type FormulaArgType = (typeof FORMULA_ARG_TYPES)[number];

export interface FormulaArgMeta {
  readonly name: string;
  readonly type: FormulaArgType;
  /** When true, the argument may be omitted. */
  readonly optional: boolean;
  /** When true, the argument may repeat (Sheets varargs). */
  readonly repeating: boolean;
}

export interface FormulaMeta {
  readonly name: string;
  readonly category: FormulaCategory;
  readonly shortDescription: string;
  readonly args: readonly FormulaArgMeta[];
}

function arg(
  name: string,
  type: FormulaArgType,
  options: { optional?: boolean; repeating?: boolean } = {},
): FormulaArgMeta {
  return {
    name,
    type,
    optional: options.optional === true,
    repeating: options.repeating === true,
  };
}

/**
 * Hand-curated metadata for the highest-value / autocomplete-facing functions.
 * Names are uppercase to match the tokenizer.
 */
export const FORMULA_CATALOG: readonly FormulaMeta[] = [
  {
    name: "XLOOKUP",
    category: "LOOKUP",
    shortDescription: "Returns values from a match in a lookup range",
    args: [
      arg("search_key", "NUMBER"),
      arg("lookup_range", "RANGE"),
      arg("result_range", "RANGE"),
      arg("missing_value", "TRANSLATED", { optional: true }),
      arg("match_mode", "NUMBER", { optional: true }),
      arg("search_mode", "NUMBER", { optional: true }),
    ],
  },
  {
    name: "XMATCH",
    category: "LOOKUP",
    shortDescription: "Returns the relative position of an item in a range",
    args: [
      arg("search_key", "NUMBER"),
      arg("lookup_range", "RANGE"),
      arg("match_mode", "NUMBER", { optional: true }),
      arg("search_mode", "NUMBER", { optional: true }),
    ],
  },
  {
    name: "SEQUENCE",
    category: "ARRAY",
    shortDescription: "Returns a sequence of numbers as an array",
    args: [
      arg("rows", "NUMBER"),
      arg("columns", "NUMBER", { optional: true }),
      arg("start", "NUMBER", { optional: true }),
      arg("step", "NUMBER", { optional: true }),
    ],
  },
  {
    name: "TOCOL",
    category: "ARRAY",
    shortDescription: "Returns the array as a single column",
    args: [arg("array", "RANGE")],
  },
  {
    name: "FLATTEN",
    category: "ARRAY",
    shortDescription: "Flattens values into a single column",
    args: [
      arg("range", "RANGE"),
      arg("range2", "RANGE", { optional: true, repeating: true }),
    ],
  },
  {
    name: "TEXTJOIN",
    category: "TEXT",
    shortDescription: "Combines text with a delimiter",
    args: [
      arg("delimiter", "LITERAL"),
      arg("ignore_empty", "BOOLEAN"),
      arg("text1", "RANGE", { repeating: true }),
    ],
  },
  {
    name: "IFS",
    category: "LOGICAL",
    shortDescription: "Evaluates multiple conditions",
    args: [
      arg("condition1", "BOOLEAN"),
      arg("value1", "FORMULA"),
      arg("condition2", "BOOLEAN", { optional: true, repeating: true }),
    ],
  },
  {
    name: "SWITCH",
    category: "LOGICAL",
    shortDescription: "Matches an expression against cases",
    args: [
      arg("expression", "FORMULA"),
      arg("case1", "FORMULA"),
      arg("value1", "FORMULA"),
      arg("default", "FORMULA", { optional: true }),
    ],
  },
  {
    name: "FILTER",
    category: "FILTER",
    shortDescription: "Filters a range by condition",
    args: [
      arg("range", "RANGE"),
      arg("condition1", "RANGE"),
      arg("condition2", "RANGE", { optional: true, repeating: true }),
    ],
  },
  {
    name: "SORT",
    category: "ARRAY",
    shortDescription: "Sorts a range or array",
    args: [
      arg("range", "RANGE"),
      arg("sort_column", "NUMBER", { optional: true }),
      arg("is_ascending", "BOOLEAN", { optional: true }),
    ],
  },
  {
    name: "UNIQUE",
    category: "ARRAY",
    shortDescription: "Returns unique rows from a range",
    args: [arg("range", "RANGE")],
  },
] as const;

const byName = new Map(FORMULA_CATALOG.map((m) => [m.name, m]));

/** Metadata for one function, or undefined when we have none. */
export function formulaMeta(name: string): FormulaMeta | undefined {
  return byName.get(name.toUpperCase());
}

/**
 * Prefix suggestions for autocomplete. Only names that both appear in the
 * catalog and are registered for evaluation (or are lazy special forms).
 */
const LAZY_FORMS = new Set([
  "IFS",
  "SWITCH",
  "IF",
  "IFERROR",
  "IFNA",
  "LET",
  "LAMBDA",
]);

export function suggestFormulas(prefix: string): readonly FormulaMeta[] {
  const needle = prefix.trim().toUpperCase();
  if (!needle) return FORMULA_CATALOG;
  return FORMULA_CATALOG.filter((m) => {
    if (!m.name.startsWith(needle)) return false;
    return m.name in FUNCTIONS || LAZY_FORMS.has(m.name);
  });
}

/** True when the engine can evaluate `name` (registry or lazy form). */
export function isImplementedFormula(name: string): boolean {
  const upper = name.toUpperCase();
  return upper in FUNCTIONS || LAZY_FORMS.has(upper);
}
