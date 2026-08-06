/**
 * Theme tokens.
 *
 * The component ships its CSS as a string injected into one `<style>` tag rather
 * than a `.css` file — that is what keeps the drop-in requirement true (no CSS
 * loader needed in the consuming app). This object is the restyling surface.
 */
export interface Theme {
  accent: string;
  border: string;
  headerBorder: string;
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
  fontFamily:
    'ui-sans-serif, system-ui, -apple-system, "Segoe UI", Roboto, sans-serif',
  monoFontFamily: "ui-monospace, SFMono-Regular, Menlo, Consolas, monospace",
  fontSize: "13px",
};

export function resolveTheme(partial?: Partial<Theme>): Theme {
  return partial ? { ...defaultTheme, ...partial } : defaultTheme;
}
