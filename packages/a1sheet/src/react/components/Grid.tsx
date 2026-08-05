/**
 * The cell grid. Ported from the grid JSX in ref/Spreadsheet.jsx:568-698.
 *
 * Non-negotiable mechanics:
 * - CSS Grid, not flexbox. Merged cells need gridColumn/gridRow spans, and sticky
 *   positioning composes cleanly inside a single scroll container.
 * - `gridTemplateColumns` is computed: ROW_HEADER_WIDTH then each column's width.
 *   `gridTemplateRows` explicitly sizes the header and the frozen-row band;
 *   `gridAutoRows` sizes everything else implicitly. That is what lets
 *   virtualization work with no spacer divs — unrendered rows simply are not grid
 *   items, and the browser still allocates track space up to the highest
 *   referenced gridRow line.
 * - Freeze panes are `position: sticky` inside ONE scrolling container. A
 *   4-quadrant split with synced scrollLeft was considered and rejected in the POC.
 *   `stickyStyleFor` computes top/left/zIndex; the corner takes the highest z.
 *   This only works because row height is fixed.
 */
import {
  type CSSProperties,
  type ReactNode,
  useCallback,
  useEffect,
  useMemo,
  useRef,
} from "react";
import { colToLetters } from "../../model/address.js";
import {
  DEFAULT_COL_WIDTH,
  HEADER_HEIGHT,
  ROW_HEADER_WIDTH,
  ROW_HEIGHT,
} from "../constants.js";
import { Cell } from "./Cell.js";
import type {
  BaseProps,
  ColumnMenuState,
  ContextMenuState,
  RenamingState,
} from "./props.js";

export interface GridProps extends BaseProps {
  renaming: RenamingState | null;
  setRenaming(r: RenamingState | null): void;
  onContextMenu(state: ContextMenuState): void;
  onColumnMenu(state: ColumnMenuState): void;
}

export function Grid({
  api,
  theme,
  prefix,
  renaming,
  setRenaming,
  onContextMenu,
  onColumnMenu,
}: GridProps): ReactNode {
  const { sheet, rowWindow, fill } = api;
  const containerRef = useRef<HTMLDivElement>(null);
  const resizeRef = useRef<{ col: number; startX: number; startW: number } | null>(
    null,
  );

  const frozenRows = sheet.frozenRows || 0;
  const frozenCols = sheet.frozenCols || 0;

  const colWidth = useCallback(
    (c: number) => sheet.colWidths[c] ?? DEFAULT_COL_WIDTH,
    [sheet.colWidths],
  );

  const cols = useMemo(
    () => Array.from({ length: sheet.numCols }, (_, i) => i),
    [sheet.numCols],
  );

  const colOffset = useCallback(
    (c: number) => {
      let x = 0;
      for (let i = 0; i < c; i++) x += colWidth(i);
      return x;
    },
    [colWidth],
  );

  const gridTemplateColumns = useMemo(
    () => `${ROW_HEADER_WIDTH}px ${cols.map((c) => `${colWidth(c)}px`).join(" ")}`,
    [cols, colWidth],
  );

  const setViewportHeight = api.setViewportHeight;
  const setColWidth = api.setColWidth;

  /** Measure the scroll container so virtualization knows how many rows to draw. */
  useEffect(() => {
    const el = containerRef.current;
    if (!el) return;
    setViewportHeight(el.clientHeight);
    if (typeof ResizeObserver === "undefined") return;
    const ro = new ResizeObserver((entries) => {
      for (const e of entries) setViewportHeight(e.contentRect.height);
    });
    ro.observe(el);
    return () => ro.disconnect();
  }, [setViewportHeight]);

  /** Column resize drag, tracked on window so it survives leaving the header. */
  useEffect(() => {
    function onMove(e: MouseEvent) {
      const st = resizeRef.current;
      if (!st) return;
      setColWidth(st.col, st.startW + (e.clientX - st.startX));
    }
    function onUp() {
      resizeRef.current = null;
    }
    window.addEventListener("mousemove", onMove);
    window.addEventListener("mouseup", onUp);
    return () => {
      window.removeEventListener("mousemove", onMove);
      window.removeEventListener("mouseup", onUp);
    };
  }, [setColWidth]);

  function stickyStyleFor(
    isHeaderRow: boolean,
    isRowHeaderCol: boolean,
    frozenRowIdx?: number,
    frozenColIdx?: number,
  ): CSSProperties {
    const style: CSSProperties = {};
    let z = 0;
    if (isHeaderRow) {
      style.top = 0;
      z = Math.max(z, 4);
    } else if (frozenRowIdx !== undefined) {
      style.top = HEADER_HEIGHT + frozenRowIdx * ROW_HEIGHT;
      z = Math.max(z, 2);
    }
    if (isRowHeaderCol) {
      style.left = 0;
      z = Math.max(z, isHeaderRow ? 5 : 3);
    } else if (frozenColIdx !== undefined) {
      style.left = ROW_HEADER_WIDTH + colOffset(frozenColIdx);
      z = Math.max(z, isHeaderRow ? 4 : 2);
    }
    if (style.top !== undefined || style.left !== undefined) {
      style.position = "sticky";
      style.zIndex = z;
      style.background = theme.cellBg;
    }
    return style;
  }

  /**
   * Fill-drag hit testing from raw mouse coordinates. Per-cell onMouseEnter was
   * unreliable during fast drags in the POC, so the overlay maps clientX/Y to a
   * row and column itself.
   */
  function hitTest(clientX: number, clientY: number) {
    const el = containerRef.current;
    if (!el) return null;
    const box = el.getBoundingClientRect();
    const x = clientX - box.left + el.scrollLeft - ROW_HEADER_WIDTH;
    const y = clientY - box.top + el.scrollTop - HEADER_HEIGHT;

    let col = 0;
    let acc = 0;
    while (col < sheet.numCols - 1 && acc + colWidth(col) <= x) {
      acc += colWidth(col);
      col++;
    }

    const visualIndex = Math.floor(y / ROW_HEIGHT);
    const row =
      visualIndex < frozenRows
        ? visualIndex
        : (rowWindow.visibleRows[visualIndex - frozenRows] ?? sheet.numRows - 1);

    return { row: Math.max(0, row), col: Math.max(0, col) };
  }

  function renderRowHeader(r: number, gridRow: number, frozenRowIdx?: number) {
    const isRenaming = renaming?.type === "row" && renaming.index === r;
    return (
      <div
        key="rh"
        className={`${prefix}head`}
        style={{
          gridColumn: 1,
          gridRow,
          height: ROW_HEIGHT,
          ...stickyStyleFor(false, true, frozenRowIdx, undefined),
        }}
        onMouseDown={() =>
          api.select({ r1: r, c1: 0, r2: r, c2: sheet.numCols - 1 })
        }
        onDoubleClick={() =>
          setRenaming({
            type: "row",
            index: r,
            value: sheet.rowLabels[r] ?? String(r + 1),
          })
        }
      >
        {isRenaming ? (
          <input
            // biome-ignore lint/a11y/noAutofocus: inline rename must take the caret
            autoFocus
            value={renaming.value}
            style={{ width: "100%" }}
            onChange={(e) => setRenaming({ ...renaming, value: e.target.value })}
            onKeyDown={(e) => {
              if (e.key === "Enter") {
                api.setRowLabel(r, renaming.value);
                setRenaming(null);
              }
              if (e.key === "Escape") setRenaming(null);
            }}
            onBlur={() => setRenaming(null)}
          />
        ) : (
          (sheet.rowLabels[r] ?? r + 1)
        )}
      </div>
    );
  }

  /**
   * `display: contents` lets one JSX element per row keep the row headers and
   * cells as direct grid items, so the grid tracks still line up.
   */
  function renderRow(absRow: number, gridRow: number, frozen: boolean) {
    return (
      <div key={`r${absRow}`} style={{ display: "contents" }}>
        {renderRowHeader(absRow, gridRow, frozen ? absRow : undefined)}
        {cols.map((c) => (
          <Cell
            key={c}
            api={api}
            theme={theme}
            prefix={prefix}
            row={absRow}
            col={c}
            gridRow={gridRow}
            stickyStyle={stickyStyleFor(
              false,
              false,
              frozen ? absRow : undefined,
              c < frozenCols ? c : undefined,
            )}
            onContextMenu={onContextMenu}
          />
        ))}
      </div>
    );
  }

  return (
    <div
      ref={containerRef}
      onScroll={(e) => api.setScrollTop((e.target as HTMLDivElement).scrollTop)}
      onMouseUp={() => {
        if (fill.dragging) fill.commit(api.updateSheet);
      }}
      style={{ flex: 1, minHeight: 0, overflow: "auto", position: "relative" }}
    >
      <div
        style={{
          display: "grid",
          gridTemplateColumns,
          gridTemplateRows: `${HEADER_HEIGHT}px repeat(${frozenRows}, ${ROW_HEIGHT}px)`,
          gridAutoRows: `${ROW_HEIGHT}px`,
          position: "relative",
        }}
      >
        {/* corner */}
        <div
          className={`${prefix}head`}
          style={{
            gridColumn: 1,
            gridRow: 1,
            height: HEADER_HEIGHT,
            ...stickyStyleFor(true, true),
          }}
        />

        {/* column headers */}
        {cols.map((c) => {
          const isRenaming = renaming?.type === "col" && renaming.index === c;
          const filtered = sheet.filters[c] !== undefined;
          return (
            <div
              key={`ch${c}`}
              className={`${prefix}head`}
              style={{
                gridColumn: c + 2,
                gridRow: 1,
                height: HEADER_HEIGHT,
                ...stickyStyleFor(
                  true,
                  false,
                  undefined,
                  c < frozenCols ? c : undefined,
                ),
              }}
              onMouseDown={() =>
                api.select({ r1: 0, c1: c, r2: sheet.numRows - 1, c2: c })
              }
              onDoubleClick={() =>
                setRenaming({
                  type: "col",
                  index: c,
                  value: sheet.colLabels[c] ?? colToLetters(c),
                })
              }
            >
              {isRenaming ? (
                <input
                  // biome-ignore lint/a11y/noAutofocus: inline rename must take the caret
                  autoFocus
                  value={renaming.value}
                  style={{ width: "100%" }}
                  onChange={(e) =>
                    setRenaming({ ...renaming, value: e.target.value })
                  }
                  onKeyDown={(e) => {
                    if (e.key === "Enter") {
                      api.setColLabel(c, renaming.value);
                      setRenaming(null);
                    }
                    if (e.key === "Escape") setRenaming(null);
                  }}
                  onBlur={() => setRenaming(null)}
                />
              ) : (
                <>
                  {sheet.colLabels[c] ?? colToLetters(c)}
                  <button
                    type="button"
                    aria-label={`Sort and filter column ${colToLetters(c)}`}
                    style={{
                      position: "absolute",
                      right: 6,
                      border: "none",
                      background: "none",
                      cursor: "pointer",
                      padding: 0,
                      fontSize: 10,
                      color: filtered ? theme.accent : theme.headerText,
                    }}
                    onMouseDown={(e) => e.stopPropagation()}
                    onClick={(e) => {
                      e.stopPropagation();
                      const box = (e.target as HTMLElement).getBoundingClientRect();
                      onColumnMenu({ col: c, x: box.left, y: box.bottom });
                    }}
                  >
                    ▾
                  </button>
                  <div
                    className={`${prefix}resize`}
                    onMouseDown={(e) => {
                      e.stopPropagation();
                      resizeRef.current = {
                        col: c,
                        startX: e.clientX,
                        startW: colWidth(c),
                      };
                    }}
                  />
                </>
              )}
            </div>
          );
        })}

        {/* frozen band, always rendered */}
        {rowWindow.frozenRowsList.map(({ absRow, gridRow }) =>
          renderRow(absRow, gridRow, true),
        )}

        {/* virtualized band */}
        {rowWindow.windowRows.map(({ absRow, gridRow }) =>
          renderRow(absRow, gridRow, false),
        )}
      </div>

      {/* Fill-drag overlay: hit-tests raw mouse coordinates. */}
      {fill.dragging && (
        <div
          style={{
            position: "absolute",
            inset: 0,
            zIndex: 20,
            cursor: "crosshair",
          }}
          onMouseMove={(e) => {
            const hit = hitTest(e.clientX, e.clientY);
            if (hit) fill.moveTo(hit.row, hit.col);
          }}
          onMouseUp={() => fill.commit(api.updateSheet)}
        />
      )}
    </div>
  );
}
