/**
 * Theme tokens.
 *
 * The component ships its CSS as a string injected into one `<style>` tag rather
 * than a `.css` file — that is what keeps the drop-in requirement true (no CSS
 * loader needed in the consuming app). This object is the restyling surface.
 *
 * Each resolved theme also maps to `--a1s-*` custom properties on the root so a
 * host app (Tailwind and other utility CSS stacks) can reference the same tokens without reaching
 * into the JS object.
 */
import type { CSSProperties } from "react";
import {
  CELL_FONT_SIZE,
  CELL_FONT_STACK,
  type CellFont,
  DEFAULT_CELL_FONT,
} from "./constants.js";

export interface Theme {
  accent: string;
  border: string;
  headerBorder: string;
  /**
   * The line marking the edge of a frozen band. Deliberately darker than
   * `border`: it separates two surfaces that scroll independently, which a line
   * of the same weight as an ordinary grid line cannot say.
   */
  freezeLine: string;
  buttonBorder: string;
  headerBg: string;
  headerText: string;
  cellBg: string;
  cellText: string;
  selectedBg: string;
  toolbarBg: string;
  /**
   * Outline colors for references in a formula being edited, cycled by the
   * reference's group. Give it as many entries as you like; a formula with more
   * distinct references than colors reuses them from the start.
   */
  refColors: readonly string[];
  /** The channel the scrollbar thumb runs in. Sits beside the cells, never over them. */
  scrollbarTrack: string;
  scrollbarThumb: string;
  scrollbarThumbHover: string;
  fontFamily: string;
  monoFontFamily: string;
  fontSize: string;
}

/** Alias for the built-in palette. */
export const lightTheme: Theme = {
  accent: "#0d9488",
  border: "#e2e8f0",
  headerBorder: "#cbd5e1",
  freezeLine: "#9aa3af",
  buttonBorder: "#d1d5db",
  headerBg: "#f8fafc",
  headerText: "#475569",
  cellBg: "#ffffff",
  cellText: "#1e293b",
  selectedBg: "rgba(13, 148, 136, 0.10)",
  toolbarBg: "#ffffff",
  refColors: ["#2563eb", "#dc2626", "#7c3aed", "#ea580c", "#0891b2"] as const,
  scrollbarTrack: "#f1f3f4",
  scrollbarThumb: "#c4c7c5",
  scrollbarThumbHover: "#9aa0a6",
  fontFamily: CELL_FONT_STACK,
  monoFontFamily: "ui-monospace, SFMono-Regular, Menlo, Consolas, monospace",
  fontSize: `${CELL_FONT_SIZE}px`,
};

/** @deprecated Use `lightTheme`. Kept for existing imports. */
export const defaultTheme: Theme = lightTheme;

/**
 * Dark palette as a partial — pass to `theme` on Root or merge with
 * `resolveTheme`. There is no dark-mode boolean; a theme is values.
 */
export const darkTheme: Partial<Theme> = {
  accent: "#2dd4bf",
  border: "#1e293b",
  headerBorder: "#334155",
  freezeLine: "#64748b",
  buttonBorder: "#334155",
  headerBg: "#0f172a",
  headerText: "#94a3b8",
  cellBg: "#0b1220",
  cellText: "#e2e8f0",
  selectedBg: "rgba(45, 212, 191, 0.16)",
  toolbarBg: "#0f172a",
  scrollbarTrack: "#0f172a",
  scrollbarThumb: "#334155",
  scrollbarThumbHover: "#475569",
};

export function resolveTheme(partial?: Partial<Theme>): Theme {
  return partial ? { ...lightTheme, ...partial } : lightTheme;
}

/** Whether native controls (color wells, selects) should use dark chrome. */
export function themeColorScheme(theme: Theme): "light" | "dark" {
  const hex = theme.cellBg.trim();
  const match = /^#?([0-9a-f]{6})$/i.exec(hex);
  if (!match) return "light";
  const n = Number.parseInt(match[1]!, 16);
  const r = (n >> 16) & 255;
  const g = (n >> 8) & 255;
  const b = n & 255;
  const luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
  return luminance < 0.45 ? "dark" : "light";
}

/** CSS custom properties for the resolved theme, for inline style on the root. */
export function themeCssVars(theme: Theme): CSSProperties {
  const vars: Record<string, string> = {
    "--a1s-accent": theme.accent,
    "--a1s-border": theme.border,
    "--a1s-header-border": theme.headerBorder,
    "--a1s-freeze-line": theme.freezeLine,
    "--a1s-button-border": theme.buttonBorder,
    "--a1s-header-bg": theme.headerBg,
    "--a1s-header-text": theme.headerText,
    "--a1s-cell-bg": theme.cellBg,
    "--a1s-cell-text": theme.cellText,
    "--a1s-selected-bg": theme.selectedBg,
    "--a1s-toolbar-bg": theme.toolbarBg,
    "--a1s-scrollbar-track": theme.scrollbarTrack,
    "--a1s-scrollbar-thumb": theme.scrollbarThumb,
    "--a1s-scrollbar-thumb-hover": theme.scrollbarThumbHover,
    "--a1s-font-family": theme.fontFamily,
    "--a1s-mono-font-family": theme.monoFontFamily,
    "--a1s-font-size": theme.fontSize,
  };
  theme.refColors.forEach((color, index) => {
    vars[`--a1s-ref-color-${index}`] = color;
  });
  return vars as CSSProperties;
}

const PX_SIZE = /^\s*(\d+(?:\.\d+)?)px\s*$/;

/**
 * The face this theme draws cells in, as the wrapped-row measurer needs it.
 *
 * Only a px `fontSize` can be honoured: `em`, `rem`, and `%` resolve against an
 * ancestor that does not exist yet at measuring time, so those fall back to the
 * default size — a wrapped row can then be a line off. Give `fontSize` in px and
 * the measurement matches what the browser draws.
 */
export function themeCellFont(theme: Theme): CellFont {
  const match = PX_SIZE.exec(theme.fontSize);
  const size = match ? Number(match[1]) : DEFAULT_CELL_FONT.size;
  return {
    family: theme.fontFamily || DEFAULT_CELL_FONT.family,
    size: size > 0 ? size : DEFAULT_CELL_FONT.size,
  };
}
