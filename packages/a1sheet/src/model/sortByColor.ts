/**
 * Reorders rows so cells matching a fill or text colour sit at the top of the
 * used range. Data mutation (like `sortByColumn`); undo via history.
 */
import { cellKey } from "./address.js";
import type { CellImage, CellKey, CellValue, RawCell, Sheet, StyleObject } from "./types.js";

export interface SortByColorOptions {
  col: number;
  kind: "background" | "foreground";
  color: string;
}

export function sortByColor(sheet: Sheet, options: SortByColorOptions): Sheet {
  const { col, kind, color } = options;
  let maxRow = -1;
  for (const key of Object.keys(sheet.cells) as CellKey[]) {
    if (sheet.cells[key] === "") continue;
    const r = Number(key.slice(0, key.indexOf("_")));
    if (r > maxRow) maxRow = r;
  }
  if (maxRow < 1) return sheet;

  const colorAt = (r: number): string | undefined => {
    const style = sheet.styles[cellKey(r, col)];
    return kind === "background" ? style?.bg : style?.color;
  };

  const order = Array.from({ length: maxRow + 1 }, (_, r) => r);
  order.sort((ra, rb) => {
    const aMatch = colorAt(ra) === color ? 0 : 1;
    const bMatch = colorAt(rb) === color ? 0 : 1;
    if (aMatch !== bMatch) return aMatch - bMatch;
    return ra - rb;
  });

  const cells: Record<CellKey, RawCell> = {};
  const styles: Record<CellKey, StyleObject> = {};
  const cachedValues: Record<CellKey, CellValue> = {};
  const images: Record<CellKey, CellImage> = {};
  for (const [newRow, oldRow] of order.entries()) {
    for (let c = 0; c < sheet.numCols; c++) {
      const from = cellKey(oldRow, c);
      const to = cellKey(newRow, c);
      const cell = sheet.cells[from];
      if (cell !== undefined) cells[to] = cell;
      const style = sheet.styles[from];
      if (style !== undefined) styles[to] = style;
      const cached = sheet.cachedValues[from];
      if (cached !== undefined) cachedValues[to] = cached;
      const image = sheet.images[from];
      if (image !== undefined) images[to] = image;
    }
  }
  const keepBelow = <T>(from: Record<CellKey, T>, into: Record<CellKey, T>) => {
    for (const key of Object.keys(from) as CellKey[]) {
      if (Number(key.slice(0, key.indexOf("_"))) > maxRow) {
        into[key] = from[key] as T;
      }
    }
  };
  keepBelow(sheet.cells, cells);
  keepBelow(sheet.styles, styles);
  keepBelow(sheet.cachedValues, cachedValues);
  keepBelow(sheet.images, images);

  return { ...sheet, cells, styles, cachedValues, images };
}
