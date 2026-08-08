"use client";

/**
 * The one context every primitive reads from.
 *
 * Primitives take no `api`/`theme`/`prefix` props — that was prop drilling and it
 * made it impossible for a consumer to write their own primitive without threading
 * three arguments they should not have to know about.
 *
 * Transient UI state (which menu is open, what is being renamed) lives here rather
 * than in the root component, because `Grid` opens menus that `Root` renders. That
 * coordination is exactly what context is for.
 */
import {
  createContext,
  type ReactNode,
  type RefObject,
  useContext,
  useMemo,
  useState,
} from "react";
import { MissingProviderError } from "../errors.js";
import type { CellContentProps, SheetComponents } from "./primitives/types.js";
import type { Theme } from "./theme.js";
import type { ColumnMenuState, ContextMenuState, RenamingState } from "./types.js";
import type { UseSpreadsheetResult } from "./useSpreadsheet.js";

/** Transient view state — never part of the workbook, so undo does not restore it. */
export interface SheetUiState {
  renaming: RenamingState | null;
  setRenaming(next: RenamingState | null): void;
  contextMenu: ContextMenuState | null;
  setContextMenu(next: ContextMenuState | null): void;
  columnMenu: ColumnMenuState | null;
  setColumnMenu(next: ColumnMenuState | null): void;
  closeMenus(): void;
}

export interface SheetContextValue {
  api: UseSpreadsheetResult;
  theme: Theme;
  /** Class-name prefix for the injected CSS, e.g. "a1s-". */
  prefix: string;
  ui: SheetUiState;
  /** Optional render overrides from `<Sheet.Root components={…}>`. */
  components: SheetComponents;
  /**
   * The hidden textarea that owns keyboard focus for the grid. Primitives call
   * `.focus()` on it after an interaction that should return focus to the sheet.
   */
  focusRef: RefObject<HTMLTextAreaElement | null>;
}

/** Per-grid overrides; `Grid.renderCellContent` wins over `components.CellContent`. */
export interface GridRenderOverrides {
  renderCellContent?: (props: CellContentProps) => ReactNode;
}

const GridRenderContext = createContext<GridRenderOverrides>({});

export function GridRenderProvider({
  value,
  children,
}: {
  value: GridRenderOverrides;
  children: ReactNode;
}) {
  return (
    <GridRenderContext.Provider value={value}>
      {children}
    </GridRenderContext.Provider>
  );
}

export function useGridRender(): GridRenderOverrides {
  return useContext(GridRenderContext);
}

const SheetContext = createContext<SheetContextValue | null>(null);

export interface SheetContextProviderProps extends SheetContextValue {
  children: ReactNode;
}

export function SheetContextProvider({
  children,
  ...value
}: SheetContextProviderProps) {
  return <SheetContext.Provider value={value}>{children}</SheetContext.Provider>;
}

/**
 * Reads the sheet context. Throws — rather than returning null — when used outside
 * `<Sheet.Root>`: that is a programmer error and failing fast beats a confusing
 * downstream crash. `component` names the caller so the message is actionable.
 */
export function useSheetContext(component = "Sheet primitive"): SheetContextValue {
  const ctx = useContext(SheetContext);
  if (!ctx) throw new MissingProviderError(component, "Sheet.Root");
  return ctx;
}

/** The headless API alone — the common case for a consumer's own component. */
export function useSheet(): UseSpreadsheetResult {
  return useSheetContext("useSheet()").api;
}

/** Builds the transient UI state. Used by `Sheet.Root`; not part of the API. */
export function useSheetUiState(): SheetUiState {
  const [renaming, setRenaming] = useState<RenamingState | null>(null);
  const [contextMenu, setContextMenu] = useState<ContextMenuState | null>(null);
  const [columnMenu, setColumnMenu] = useState<ColumnMenuState | null>(null);

  return useMemo(
    () => ({
      renaming,
      setRenaming,
      contextMenu,
      setContextMenu,
      columnMenu,
      setColumnMenu,
      closeMenus: () => {
        setContextMenu(null);
        setColumnMenu(null);
      },
    }),
    [renaming, contextMenu, columnMenu],
  );
}
