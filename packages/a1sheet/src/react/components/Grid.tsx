"use client";

/**
 * The cell grid. Ported from the grid JSX in ref/Spreadsheet.jsx:568-698.
 *
 * Non-negotiable mechanics:
 * - CSS Grid, not flexbox. Merged cells need gridColumn/gridRow spans, and sticky
 *   positioning composes cleanly inside a single scroll container.
 * - `gridTemplateColumns` names EVERY column, virtualized or not. A track
 *   definition costs a string entry; a cell costs a DOM node. Keeping the tracks
 *   explicit is what gives the container its true width and keeps sticky offsets
 *   honest, while `colWindow` decides which cells exist.
 * - `gridTemplateRows` explicitly sizes the header and the frozen-row band.
 *   Everything below is an implicit track sized by its contents, because rows
 *   can differ in height. Windowed rows take CONSECUTIVE grid lines and a
 *   single spacer item ahead of them covers the rows scrolled off the top —
 *   absolute lines would require the browser to size tracks for rows that are
 *   not in the DOM, which it cannot do once heights vary.
 * - The grid is therefore only as large as what is drawn. `minHeight` and
 *   `minWidth` restore the real extent so the scrollbar describes the sheet
 *   rather than the window into it.
 * - The scrollbars are ours, in channels of their own to the right of and below
 *   the scroll container, which hides its native ones. See `Scrollbar` for why.
 * - Freeze panes are `position: sticky` inside ONE scrolling container. A
 *   4-quadrant split with synced scrollLeft was considered and rejected in the POC.
 *   `stickyStyleFor` computes top/left/zIndex; the corner takes the highest z.
 *   Offsets come from the cumulative tables in `useRowWindow`/`useColWindow`,
 *   so they hold whatever the row heights and column widths are.
 */
import {
  type CSSProperties,
  type ReactNode,
  useCallback,
  useEffect,
  useId,
  useMemo,
  useRef,
} from "react";
import { colToLetters } from "../../model/address.js";
import type { Range } from "../../model/types.js";
import {
  AUTOFIT_SAMPLE_LIMIT,
  HEADER_HEIGHT,
  MAX_AUTOFIT_COL_WIDTH,
  MIN_COL_WIDTH,
  ROW_HEADER_WIDTH,
  SCROLLBAR_SIZE,
} from "../constants.js";
import { useSheetContext } from "../context.js";
import { useTextMeasurer } from "../useTextMeasurer.js";
import { Cell } from "./Cell.js";
import { ChevronDownIcon } from "./icons.js";
import { Scrollbar } from "./Scrollbar.js";

/** Left and right padding on a cell, from the stylesheet. Auto-fit must clear it. */
const CELL_PADDING_X = 12;

export interface GridProps {
  /**
   * Rendered inside the scroll container, after the last row. Put anything that
   * belongs at the end of the sheet here — `<Sheet.AddRows />`, a totals
   * banner — and it scrolls with the content instead of sitting in a fixed bar
   * below it.
   */
  children?: ReactNode;
}

export function Grid({ children }: GridProps = {}): ReactNode {
  const { api, theme, prefix, ui, focusRef } = useSheetContext("Sheet.Grid");
  const { renaming, setRenaming } = ui;
  const { sheet, rowWindow, colWindow, fill, bounds } = api;
  const containerRef = useRef<HTMLDivElement>(null);
  const resizeRef = useRef<{ col: number; startX: number; startW: number } | null>(
    null,
  );
  const resizeRowRef = useRef<{
    row: number;
    startY: number;
    startH: number;
  } | null>(null);
  const measureText = useTextMeasurer();
  const scrollerId = useId();

  /**
   * Move the scroll container, not the virtualization state.
   *
   * The two are kept in step by the container's own `scroll` event: writing
   * `scrollTop` here fires it, and the handler above is what tells the window
   * where to draw. Setting `api.setScrollTop` directly instead would move the
   * drawing without moving the element — the bug this pairing exists to avoid.
   * The browser clamps out-of-range values, so a drag past either end settles
   * at the end rather than needing clamping here.
   */
  const scrollToTop = useCallback((offset: number) => {
    const el = containerRef.current;
    if (el) el.scrollTop = offset;
  }, []);
  const scrollToLeft = useCallback((offset: number) => {
    const el = containerRef.current;
    if (el) el.scrollLeft = offset;
  }, []);

  const frozenRows = sheet.frozenRows || 0;
  const frozenCols = sheet.frozenCols || 0;

  const { colWidth, colOffset, colAt, windowCols } = colWindow;

  const gridTemplateColumns = useMemo(() => {
    const tracks: string[] = [`${ROW_HEADER_WIDTH}px`];
    for (let c = 0; c < sheet.numCols; c++) tracks.push(`${colWidth(c)}px`);
    return tracks.join(" ");
  }, [sheet.numCols, colWidth]);

  // Only the header and the frozen band are explicit. Everything below is an
  // implicit track sized by whatever lands in it — the spacer, then the window.
  const rowHeight = rowWindow.rowHeight;
  const gridTemplateRows = useMemo(() => {
    const tracks: string[] = [`${HEADER_HEIGHT}px`];
    for (let r = 0; r < frozenRows; r++) tracks.push(`${rowHeight(r)}px`);
    return tracks.join(" ");
  }, [frozenRows, rowHeight]);

  const setViewportHeight = api.setViewportHeight;
  const setViewportWidth = api.setViewportWidth;
  const setColWidth = api.setColWidth;
  const setRowHeight = api.setRowHeight;
  const setScrollTop = api.setScrollTop;
  const setScrollLeft = api.setScrollLeft;

  /**
   * Measure the scroll container so virtualization knows how much to draw.
   *
   * A zero measurement is discarded rather than believed. A container that is
   * `display: none`, or that has not been laid out yet, reports 0 — and taking
   * that at face value windows the grid down to nothing, so it stays empty even
   * after it becomes visible if no resize follows.
   */
  useEffect(() => {
    const el = containerRef.current;
    if (!el) return;
    const measure = (height: number, width: number) => {
      if (height > 0) setViewportHeight(height);
      if (width > 0) setViewportWidth(width);
    };
    measure(el.clientHeight, el.clientWidth);

    // The container does not always start at the top. A browser restoring
    // scroll on back-navigation, a bfcache restore, or a consumer scrolling it
    // before we mount all set scrollTop without an onScroll we can hear, so
    // virtualization would go on drawing the top of the sheet into a container
    // showing the middle of it — rows in the wrong place, and blank space where
    // content should be. Reading it once on mount is what keeps the two agreed.
    if (el.scrollTop > 0) setScrollTop(el.scrollTop);
    if (el.scrollLeft > 0) setScrollLeft(el.scrollLeft);
    if (typeof ResizeObserver === "undefined") return;
    const ro = new ResizeObserver((entries) => {
      for (const e of entries) measure(e.contentRect.height, e.contentRect.width);
    });
    ro.observe(el);
    return () => ro.disconnect();
  }, [setViewportHeight, setViewportWidth, setScrollTop, setScrollLeft]);

  /** Resize drags, tracked on window so they survive leaving the header. */
  useEffect(() => {
    function onMove(e: MouseEvent) {
      const col = resizeRef.current;
      if (col) setColWidth(col.col, col.startW + (e.clientX - col.startX));
      const row = resizeRowRef.current;
      if (row) setRowHeight(row.row, row.startH + (e.clientY - row.startY));
    }
    function onUp() {
      resizeRef.current = null;
      resizeRowRef.current = null;
    }
    window.addEventListener("mousemove", onMove);
    window.addEventListener("mouseup", onUp);
    return () => {
      window.removeEventListener("mousemove", onMove);
      window.removeEventListener("mouseup", onUp);
    };
  }, [setColWidth, setRowHeight]);

  /**
   * Auto-fit a column to the widest thing in it, as double-clicking the divider
   * does in Excel and Sheets.
   *
   * Candidate rows come from `sheet.cells`, so a sparse column costs its filled
   * cells rather than its length — and the count is capped anyway
   * (AUTOFIT_SAMPLE_LIMIT), because a double-click must not stall. A column with
   * more values than that is fitted to the widest of the ones measured, which
   * can leave a later, longer value clipped.
   */
  const autoFitCol = useCallback(
    (col: number) => {
      const sample = containerRef.current?.querySelector(`.${prefix}cell`);
      if (!sample || typeof getComputedStyle === "undefined") return;
      const font = getComputedStyle(sample).font;

      let widest = measureText(sheet.colLabels[col] ?? colToLetters(col), font);
      let measured = 0;
      const suffix = `_${col}`;
      for (const key of Object.keys(sheet.cells)) {
        if (!key.endsWith(suffix)) continue;
        if (measured >= AUTOFIT_SAMPLE_LIMIT) break;
        measured++;
        const row = Number(key.slice(0, key.length - suffix.length));
        const width = measureText(api.getDisplay(row, col), font);
        if (width > widest) widest = width;
      }

      // Zero means the environment gave us no canvas to measure with. Leaving
      // the column alone beats collapsing it to the minimum.
      if (widest === 0) return;
      api.setColWidth(
        col,
        Math.min(
          MAX_AUTOFIT_COL_WIDTH,
          Math.max(MIN_COL_WIDTH, widest + CELL_PADDING_X),
        ),
      );
    },
    [prefix, measureText, sheet.cells, sheet.colLabels, api],
  );

  /**
   * Auto-fit a row. Cells are single-line — `white-space: nowrap` — so the
   * height that hugs the content is the default height, whatever the content
   * is. Dropping the override is therefore the whole operation, and it stays
   * correct if wrapping is added later only if this is revisited then.
   */
  const autoFitRow = api.resetRowHeight;

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
      style.top = HEADER_HEIGHT + (rowWindow.rowTop(frozenRowIdx) ?? 0);
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
      // Deliberately no background here. A sticky element must be opaque or
      // scrolled content shows through, but both kinds already are: headers from
      // the `head` class, cells from the background `Cell` computes. Setting one
      // inline would outrank the class that highlights a header in the selection.
    }
    return style;
  }

  /**
   * Pixel box for a range, in the grid element's coordinates.
   *
   * The inverse of `hitTest`: a row's y is its position in the visible order, so
   * hidden and filtered-out rows collapse exactly as they do on screen. Returns
   * null when no part of the range is currently visible, which is why this
   * cannot be a plain multiply — a filtered sheet has no linear row mapping.
   */
  function rectFor(range: Range): CSSProperties | null {
    const top = rowWindow.rowTop(range.r1);
    const bottom = rowWindow.rowTop(range.r2);
    if (top === null || bottom === null) return null;

    const left = colOffset(range.c1);
    const right = colOffset(range.c2) + colWidth(range.c2);

    return {
      left: ROW_HEADER_WIDTH + left,
      top: HEADER_HEIGHT + top,
      width: right - left,
      height: bottom - top + rowWindow.rowHeight(range.r2),
    };
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

    return { row: Math.max(0, rowWindow.rowAt(y)), col: Math.max(0, colAt(x)) };
  }

  function renderRowHeader(r: number, gridRow: number, frozenRowIdx?: number) {
    const isRenaming = renaming?.type === "row" && renaming.index === r;
    const inSelection = r >= bounds.r1 && r <= bounds.r2;
    return (
      <div
        key="rh"
        className={`${prefix}head${inSelection ? ` ${prefix}headon` : ""}`}
        data-row={r}
        style={{
          gridColumn: 1,
          gridRow,
          // The row header is the one item guaranteed to exist in every row, so
          // its height is what pins the track down when every cell in the row
          // is covered by a merge.
          height: rowWindow.rowHeight(r),
          ...stickyStyleFor(false, true, frozenRowIdx, undefined),
        }}
        onMouseDown={(e) => {
          if ((e.target as HTMLElement).tagName === "INPUT") return;
          e.preventDefault();
          focusRef.current?.focus();
          api.select({ r1: r, c1: 0, r2: r, c2: sheet.numCols - 1 });
        }}
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
          <>
            {sheet.rowLabels[r] ?? r + 1}
            <div
              className={`${prefix}rowresize`}
              onMouseDown={(e) => {
                e.stopPropagation();
                resizeRowRef.current = {
                  row: r,
                  startY: e.clientY,
                  startH: rowWindow.rowHeight(r),
                };
              }}
              onDoubleClick={(e) => {
                e.stopPropagation();
                autoFitRow(r);
              }}
            />
          </>
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
        {windowCols.map((c) => (
          <Cell
            key={c}
            row={absRow}
            col={c}
            gridRow={gridRow}
            stickyStyle={stickyStyleFor(
              false,
              false,
              frozen ? absRow : undefined,
              c < frozenCols ? c : undefined,
            )}
          />
        ))}
      </div>
    );
  }

  return (
    <div
      // Four areas: the sheet, a channel per axis, and the corner between them.
      // The bars are siblings of the scroll container rather than inside it, so
      // they can never be scrolled away from or drawn over the cells.
      className={`${prefix}frame`}
      style={{
        flex: 1,
        minHeight: 0,
        display: "grid",
        gridTemplateColumns: `minmax(0, 1fr) ${SCROLLBAR_SIZE}px`,
        gridTemplateRows: `minmax(0, 1fr) ${SCROLLBAR_SIZE}px`,
      }}
    >
      <div
        ref={containerRef}
        id={scrollerId}
        // Named so it can be found. Scrolling the sheet from outside means
        // moving this element — setting `api.setScrollTop` alone only tells
        // virtualization where to draw, and the container would stay put.
        className={`${prefix}scroller`}
        onScroll={(e) => {
          const el = e.target as HTMLDivElement;
          setScrollTop(el.scrollTop);
          setScrollLeft(el.scrollLeft);
        }}
        onMouseUp={() => {
          if (fill.dragging) fill.commit(api.updateSheet);
        }}
        style={{ overflow: "auto", position: "relative" }}
      >
        <div
          style={{
            display: "grid",
            gridTemplateColumns,
            gridTemplateRows,
            // `min-content`, not a fixed height: each row now carries its own,
            // and the track has to take it from the items in it.
            gridAutoRows: "min-content",
            // The real extent of the sheet on both axes, not of the window into
            // it. Vertically, implicit tracks stop at the last rendered row, so
            // the scrollbar would report a few hundred pixels for a hundred
            // thousand rows. Horizontally the cause differs: this div is
            // block-level, so it takes the scroller's width and the column tracks
            // overflow it — which makes the scrollable width follow whichever
            // cells happen to be rendered. Both tracks are fixed-size, so the
            // surplus collects at the edges instead of stretching them.
            minHeight: rowWindow.contentHeight,
            minWidth: ROW_HEADER_WIDTH + colWindow.totalWidth,
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
          {windowCols.map((c) => {
            const isRenaming = renaming?.type === "col" && renaming.index === c;
            const filtered = sheet.filters[c] !== undefined;
            const inSelection = c >= bounds.c1 && c <= bounds.c2;
            return (
              <div
                key={`ch${c}`}
                className={`${prefix}head${inSelection ? ` ${prefix}headon` : ""}`}
                data-col={c}
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
                onMouseDown={(e) => {
                  if ((e.target as HTMLElement).tagName === "INPUT") return;
                  e.preventDefault();
                  focusRef.current?.focus();
                  api.select({ r1: 0, c1: c, r2: sheet.numRows - 1, c2: c });
                }}
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
                        display: "flex",
                        color: filtered ? theme.accent : theme.headerText,
                      }}
                      onMouseDown={(e) => e.stopPropagation()}
                      onClick={(e) => {
                        e.stopPropagation();
                        const box = (
                          e.target as HTMLElement
                        ).getBoundingClientRect();
                        ui.setColumnMenu({ col: c, x: box.left, y: box.bottom });
                      }}
                    >
                      <ChevronDownIcon />
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
                      onDoubleClick={(e) => {
                        // Without this the header's own double-click starts a
                        // rename, which is not what aiming at the divider meant.
                        e.stopPropagation();
                        autoFitCol(c);
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

          {/* One item standing in for every row scrolled off the top. Rows can
            differ in height, so the window cannot be placed by grid line — an
            absent row has no track and nothing to size it. This carries the
            whole distance in a single track instead. */}
          <div
            aria-hidden="true"
            className={`${prefix}spacer`}
            style={{
              gridColumn: "1 / -1",
              gridRow: frozenRows + 2,
              height: rowWindow.leadingSpace,
              pointerEvents: "none",
            }}
          />

          {/* virtualized band */}
          {rowWindow.windowRows.map(({ absRow, gridRow }) =>
            renderRow(absRow, gridRow, false),
          )}

          {/* Reference outlines for the formula being typed. Rendered inside the
            grid element so they scroll with it, and as siblings of the cells so
            a cell's own borders are not disturbed. */}
          {api.formulaRefs.spans.map((span) => {
            const box = rectFor(span.range);
            if (!box) return null;
            const color = theme.refColors[
              span.group % Math.max(1, theme.refColors.length)
            ] as string;
            return (
              <div
                key={`${span.start}-${span.end}`}
                aria-hidden="true"
                style={{
                  position: "absolute",
                  ...box,
                  border: `2px solid ${color}`,
                  background: `${color}14`,
                  pointerEvents: "none",
                  zIndex: 6,
                }}
              />
            );
          })}
        </div>

        {/* End-of-sheet slot. Outside the grid element so it is not a grid item
            competing for a track, but inside the scroller so it sits after the
            last row and scrolls with it. */}
        {children}

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

      <Scrollbar
        orientation="vertical"
        viewport={api.viewportHeight}
        content={rowWindow.contentHeight}
        offset={api.scrollTop}
        controls={scrollerId}
        onScrollTo={scrollToTop}
        prefix={prefix}
      />
      <Scrollbar
        orientation="horizontal"
        viewport={api.viewportWidth}
        content={ROW_HEADER_WIDTH + colWindow.totalWidth}
        offset={api.scrollLeft}
        controls={scrollerId}
        onScrollTo={scrollToLeft}
        prefix={prefix}
      />
      <div className={`${prefix}sbcorner`} />
    </div>
  );
}
