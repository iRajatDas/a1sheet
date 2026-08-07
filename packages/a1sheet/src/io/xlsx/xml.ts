/**
 * XML helpers for OOXML: escaping, colors, and a hand-written scanner.
 *
 * Why a scanner and not DOMParser: DOMParser exists only in browsers, and this
 * package's "." entrypoint must work in Node and Web Workers too. Depending on
 * it would also make `readXlsx` untestable outside a DOM.
 *
 * A regex scanner is sufficient here and nowhere near a general XML parser: the
 * documents are machine-generated SpreadsheetML with shallow, predictable
 * structure. It does NOT handle CDATA, namespaced attribute lookup, or comments
 * containing tag-like text. If a1sheet ever needs to read arbitrary XML, this is
 * the wrong tool.
 */
import type { HexColor } from "../../model/types.js";

export function xmlEscape(s: unknown): string {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

/** Reverses `xmlEscape`, plus numeric character references. */
export function xmlUnescape(s: string): string {
  return s
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(Number(d)))
    .replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/&amp;/g, "&");
}

/** "#rrggbb" -> "FFrrggbb" (OOXML ARGB). Falls back to opaque black. */
export function colorToRgb(hex?: string): string {
  if (!hex) return "FF000000";
  const h = hex.replace("#", "").toUpperCase();
  return `FF${h.length === 6 ? h : "000000"}`;
}

/** "FFrrggbb" -> "#rrggbb". Returns null for missing or malformed input. */
export function rgbToColor(rgb?: string | null): HexColor | null {
  if (!rgb || rgb.length < 6) return null;
  return `#${rgb.slice(-6).toLowerCase()}` as HexColor;
}

export interface XmlElement {
  /** Attribute values, already unescaped. Local names only — prefixes are kept. */
  attrs: Record<string, string>;
  /** Raw inner XML. Empty string for a self-closing tag. */
  inner: string;
}

const ATTR_RE = /([\w:.-]+)\s*=\s*"([^"]*)"/g;

function parseAttrs(source: string): Record<string, string> {
  const attrs: Record<string, string> = {};
  ATTR_RE.lastIndex = 0;
  let m = ATTR_RE.exec(source);
  while (m) {
    attrs[m[1] as string] = xmlUnescape(m[2] as string);
    m = ATTR_RE.exec(source);
  }
  return attrs;
}

/**
 * Yields every `<tag …>…</tag>` and `<tag …/>` at any depth, lazily.
 *
 * Lazy because a worksheet in a large workbook holds millions of `<c>` elements:
 * collecting them into an array costs the memory of the whole sheet before the
 * first cell can be looked at, and gives the reader no place to yield to the event
 * loop. Consuming this one element at a time lets `readXlsx` checkpoint mid-sheet.
 *
 * Nesting caveat: for a tag that can contain itself this yields the OUTERMOST
 * match's inner content spanning to the first matching close tag, which is wrong
 * for recursive structures. No SpreadsheetML element we read is self-nesting.
 */
export function* iterElements(xml: string, tag: string): Generator<XmlElement> {
  const open = new RegExp(`<${tag}(\\s[^>]*?)?(/?)>`, "g");
  let m = open.exec(xml);

  while (m) {
    const attrs = parseAttrs(m[1] ?? "");
    if (m[2] === "/") {
      yield { attrs, inner: "" };
      m = open.exec(xml);
      continue;
    }
    const contentStart = m.index + m[0].length;
    const closeIdx = xml.indexOf(`</${tag}>`, contentStart);
    const inner = closeIdx === -1 ? "" : xml.slice(contentStart, closeIdx);
    yield { attrs, inner };
    open.lastIndex = closeIdx === -1 ? contentStart : closeIdx + tag.length + 3;
    m = open.exec(xml);
  }
}

/** Every match, collected. Use `iterElements` when the count may be unbounded. */
export function findElements(xml: string, tag: string): XmlElement[] {
  return [...iterElements(xml, tag)];
}

/**
 * The first matching element, or null. Stops at the first match — this runs three
 * times per cell, so scanning the whole fragment would be waste at scale.
 */
export function findElement(xml: string, tag: string): XmlElement | null {
  for (const el of iterElements(xml, tag)) return el;
  return null;
}

/** True when a (possibly self-closing) tag appears anywhere in the fragment. */
export function hasElement(xml: string, tag: string): boolean {
  return new RegExp(`<${tag}(\\s[^>]*?)?/?>`).test(xml);
}

/** Concatenated text of a fragment with all tags removed, unescaped. */
export function textOf(xml: string): string {
  return xmlUnescape(xml.replace(/<[^>]*>/g, ""));
}
