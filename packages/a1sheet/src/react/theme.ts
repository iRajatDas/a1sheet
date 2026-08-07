/**
 * Theme tokens.
 *
 * The component ships its CSS as a string injected into one `<style>` tag rather
 * than a `.css` file — that is what keeps the drop-in requirement true (no CSS
 * loader needed in the consuming app). This object is the restyling surface.
 */
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

export const defaultTheme: Theme = {
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

export function resolveTheme(partial?: Partial<Theme>): Theme {
  return partial ? { ...defaultTheme, ...partial } : defaultTheme;
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
