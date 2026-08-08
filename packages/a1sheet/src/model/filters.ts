/**
 * Column filters and named filter views.
 *
 * Value filters stay view-layer only (hide rows). Colour criteria match
 * `StyleObject.bg` / `.color`. Filter views snapshot `filters` without touching
 * cells — activate copies the snapshot onto the live map.
 */
import { filterIdExists } from "./gridErrors.js";
import type { ColumnFilter, FilterView, Sheet } from "./types.js";

export type FilterInput = ColumnFilter | Set<string>;

/** Accepts a bare value-set (legacy) or a full criteria object. */
export function normalizeColumnFilter(input: FilterInput): ColumnFilter {
  if (input instanceof Set) return { values: input };
  return input;
}

export function cloneColumnFilter(filter: ColumnFilter): ColumnFilter {
  const next: ColumnFilter = {};
  if (filter.values) next.values = new Set(filter.values);
  if (filter.background) next.background = new Set(filter.background);
  if (filter.foreground) next.foreground = new Set(filter.foreground);
  return next;
}

export function cloneFilters(
  filters: Record<number, ColumnFilter>,
): Record<number, ColumnFilter> {
  const next: Record<number, ColumnFilter> = {};
  for (const key of Object.keys(filters)) {
    const col = Number(key);
    const filter = filters[col];
    if (filter) next[col] = cloneColumnFilter(filter);
  }
  return next;
}

export function isColumnFilterEmpty(filter: ColumnFilter): boolean {
  return (
    (filter.values === undefined || filter.values.size === 0) &&
    (filter.background === undefined || filter.background.size === 0) &&
    (filter.foreground === undefined || filter.foreground.size === 0)
  );
}

/**
 * Whether a cell passes one column's criteria. Colour checks use exact hex
 * strings as stored on the style (callers normalize if they need case-folding).
 */
export function rowMatchesColumnFilter(options: {
  filter: ColumnFilter;
  display: string;
  background: string | undefined;
  foreground: string | undefined;
}): boolean {
  const { filter, display, background, foreground } = options;
  if (filter.values && !filter.values.has(display)) return false;
  if (filter.background) {
    if (background === undefined || !filter.background.has(background)) {
      return false;
    }
  }
  if (filter.foreground) {
    if (foreground === undefined || !filter.foreground.has(foreground)) {
      return false;
    }
  }
  return true;
}

export function createFilterView(
  sheet: Sheet,
  options: { id: string; name: string },
): Sheet {
  const views = sheet.filterViews ?? {};
  if (options.id in views) {
    throw filterIdExists(options.id);
  }
  const view: FilterView = {
    id: options.id,
    name: options.name,
    filters: cloneFilters(sheet.filters ?? {}),
  };
  return {
    ...sheet,
    filterViews: { ...views, [options.id]: view },
    activeFilterViewId: options.id,
  };
}

export function activateFilterView(sheet: Sheet, id: string): Sheet | null {
  const view = sheet.filterViews?.[id];
  if (!view) return null;
  return {
    ...sheet,
    filters: cloneFilters(view.filters),
    activeFilterViewId: id,
  };
}

export function deleteFilterView(sheet: Sheet, id: string): Sheet {
  const views = sheet.filterViews ?? {};
  if (!(id in views)) return sheet;
  const { [id]: _removed, ...rest } = views;
  return {
    ...sheet,
    filterViews: rest,
    activeFilterViewId:
      sheet.activeFilterViewId === id ? null : sheet.activeFilterViewId,
  };
}

export const FILTER_VIEW_MISSING = "The view does not exist.";

export function colorMovedToTopMessage(options: {
  kind: "background" | "foreground";
  color: string;
}): string {
  const label =
    options.kind === "background" ? "background color" : "foreground color";
  return `Cells with ${label} ${options.color} were moved to the top`;
}

export function filteredColumnSortedMessage(options: {
  colLabel: string;
  ascending: boolean;
}): string {
  const dir = options.ascending ? "ascending" : "descending";
  return `Filtered column ${options.colLabel} sorted in ${dir} order`;
}
