/**
 * Built-in font faces for the toolbar picker.
 *
 * System stacks only — no network fetch, no runtime dependency. Names match what
 * Excel and Sheets label in their menus; the browser resolves them locally.
 */
export interface FontFace {
  /** Stored on `StyleObject.fontFamily`. Empty means inherit the sheet theme. */
  id: string;
  /** Shown in the menu, rendered in the face when possible. */
  label: string;
}

export const THEME_FONT: FontFace = { id: "", label: "Default (Theme)" };

export const SHEET_FONTS: readonly FontFace[] = [
  THEME_FONT,
  { id: "Arial", label: "Arial" },
  { id: "Calibri", label: "Calibri" },
  { id: "Cambria", label: "Cambria" },
  { id: "Comic Sans MS", label: "Comic Sans MS" },
  { id: "Courier New", label: "Courier New" },
  { id: "Georgia", label: "Georgia" },
  { id: "Helvetica", label: "Helvetica" },
  { id: "Impact", label: "Impact" },
  { id: "Lucida Console", label: "Lucida Console" },
  { id: "Palatino Linotype", label: "Palatino Linotype" },
  { id: "Segoe UI", label: "Segoe UI" },
  { id: "Tahoma", label: "Tahoma" },
  { id: "Times New Roman", label: "Times New Roman" },
  { id: "Trebuchet MS", label: "Trebuchet MS" },
  { id: "Verdana", label: "Verdana" },
  { id: "Roboto", label: "Roboto" },
  { id: "Roboto Mono", label: "Roboto Mono" },
] as const;

const RECENT_LIMIT = 6;
let recentFontIds: string[] = [];

export function noteRecentFont(id: string): void {
  if (!id) return;
  recentFontIds = [id, ...recentFontIds.filter((f) => f !== id)].slice(0, RECENT_LIMIT);
}

export function recentFonts(): readonly FontFace[] {
  const byId = new Map(SHEET_FONTS.map((f) => [f.id, f]));
  return recentFontIds
    .map((id) => byId.get(id))
    .filter((f): f is FontFace => f !== undefined);
}

export function fontById(id: string | undefined): FontFace {
  return SHEET_FONTS.find((f) => f.id === (id ?? "")) ?? THEME_FONT;
}
