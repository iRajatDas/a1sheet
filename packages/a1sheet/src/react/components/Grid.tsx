"use client";

/**
 * The cell grid.
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
 *   4-quadrant split with synced scrollLeft was considered and rejected.
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
import { cellKey, colToLetters, normalizeRange } from "../../model/address.js";
import type { Range } from "../../model/types.js";
import {
  AUTOFIT_SAMPLE_LIMIT,
  CELL_PADDING_X,
  HEADER_HEIGHT,
  MAX_AUTOFIT_COL_WIDTH,
  MIN_COL_WIDTH,
  ROW_HEADER_WIDTH,
  SCROLLBAR_SIZE,
} from "../constants.js";
import { GridRenderProvider, useSheetContext } from "../context.js";
import { mergeClass } from "../primitives/mergeClass.js";
import type { CellContentProps, PrimitiveProps } from "../primitives/types.js";
import { revealOffset } from "../reveal.js";
import { useStylePreview } from "../stylePreview.js";
import { useAutoScroll } from "../useAutoScroll.js";
import { useTextMeasurer } from "../useTextMeasurer.js";
import { Cell } from "./Cell.js";
import { ChevronDownIcon } from "./icons.js";
import { Scrollbar } from "./Scrollbar.js";

/**
 * The room the sort/filter button holds in a column header: a 12px icon plus its
 * 2px margin. Auto-fit adds it to the header's own width, so fitting a column to
 * its title does not leave the title ellipsed by the button beside it.
 */
const HEADER_MENU_WIDTH = 14;
/** Thickness of the line marking the edge of a frozen band, as Sheets draws it. */
const FREEZE_LINE_PX = 2;

export interface GridProps extends PrimitiveProps {
  /**
   * Rendered inside the scroll container, after the last row. Put anything that
   * belongs at the end of the sheet here — `<Sheet.AddRows />`, a totals
   * banner — and it scrolls with the content instead of sitting in a fixed bar
   * below it.
   */
  children?: ReactNode;
  /**
   * Replaces the default cell display for this grid. Wins over
   * `<Sheet.Root components={{ CellContent }}>`. Editing, selection, and merges
   * stay in `Cell`.
   */
  renderCellContent?: (props: CellContentProps) => ReactNode;
  /** Applied to the scroll container inside the frame. */
  scrollerClassName?: string;
  scrollerStyle?: CSSProperties;
}

export function Grid({
  children,
  className,
  style,
  renderCellContent,
  scrollerClassName,
  scrollerStyle,
}: GridProps = {}): ReactNode {
  const { api, theme, prefix, ui, focusRef } = useSheetContext("Sheet.Grid");
  const { patch: stylePreviewPatch } = useStylePreview();
  const hasStylePreview = Object.keys(stylePreviewPatch).length > 0;
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
   * Which drag is in progress, if any. Not state: no render depends on it.
   *
   * The three selecting kinds differ only in what the pointer's position means —
   * a cell, a whole row, or a whole column — which is why they are one drag with
   * a mode rather than three.
   */
  const dragRef = useRef<"select" | "rows" | "cols" | "fill" | null>(null);
  const dragMoveRef = useRef<(clientX: number, clientY: number) => void>(() => {});
  const dragEndRef = useRef<() => void>(() => {});

  /**
   * Hold a drag near an edge and the sheet scrolls, faster the further past it
   * you push — the only way to select or fill past the bottom of the screen.
   * Each frame re-runs the drag handler at the pointer's last position, because
   * the content moved underneath it.
   */
  const autoScroll = useAutoScroll(containerRef, (x, y) =>
    dragMoveRef.current(x, y),
  );
  const autoScrollRef = useRef(autoScroll);
  autoScrollRef.current = autoScroll;

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

  /**
   * Selection and fill drags, tracked on window so they survive the pointer
   * leaving the grid entirely. Mounted once: the handlers read everything they
   * need through refs, so re-subscribing per render would only churn.
   */
  useEffect(() => {
    function onMove(e: MouseEvent) {
      if (!dragRef.current) return;
      // A drag that started on a cell must not turn into a text selection of
      // whatever the pointer passes over on its way.
      e.preventDefault();
      dragMoveRef.current(e.clientX, e.clientY);
      autoScrollRef.current.track(e.clientX, e.clientY);
    }
    function onUp() {
      if (!dragRef.current) return;
      dragEndRef.current();
      dragRef.current = null;
      autoScrollRef.current.stop();
    }
    window.addEventListener("mousemove", onMove);
    window.addEventListener("mouseup", onUp);
    return () => {
      window.removeEventListener("mousemove", onMove);
      window.removeEventListener("mouseup", onUp);
    };
  }, []);

  /**
   * Keep the moving end of the selection on screen.
   *
   * Without this the keyboard is unusable past the first screenful: arrow down
   * from the last visible row and the selection moves somewhere you cannot see,
   * with nothing scrolling to follow it. It tracks `selection.r2`/`c2` rather
   * than the active cell, because during a Shift+arrow extension that is the end
   * doing the moving.
   *
   * Only the minimum scroll needed, and only when the cell is actually outside —
   * a cell already in view must never be re-centred, or every keystroke would
   * shove the sheet around.
   */
  const focusRow = api.selection.r2;
  const focusCol = api.selection.c2;
  const rowTop = rowWindow.rowTop;
  useEffect(() => {
    const el = containerRef.current;
    if (!el || dragRef.current) return;

    // An unmeasured container — display:none, or not laid out yet — reports a
    // zero client size, and every cell in it reads as below the fold. Believing
    // that scrolls a sheet nobody is looking at to the bottom of the selection.
    if (el.clientHeight === 0 || el.clientWidth === 0) return;

    const top = rowTop(focusRow);
    if (top !== null) {
      el.scrollTop = revealOffset({
        offset: el.scrollTop,
        viewport: el.clientHeight,
        start: top,
        size: rowWindow.rowHeight(focusRow),
        lead: HEADER_HEIGHT,
      });
    }

    el.scrollLeft = revealOffset({
      offset: el.scrollLeft,
      viewport: el.clientWidth,
      start: colOffset(focusCol),
      size: colWidth(focusCol),
      lead: ROW_HEADER_WIDTH,
    });
  }, [focusRow, focusCol, rowTop, rowWindow.rowHeight, colOffset, colWidth]);

  /**
   * A fill drag starts on the handle inside a cell, whose mousedown stops
   * propagating before the grid can see it. The hook's own state is the signal.
   */
  useEffect(() => {
    if (fill.dragging) dragRef.current = "fill";
  }, [fill.dragging]);

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

      let widest =
        measureText(sheet.colLabels[col] ?? colToLetters(col), font) +
        HEADER_MENU_WIDTH;
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
   * Auto-fit a row, which is still just dropping its explicit height — but for
   * a different reason than it used to be. `rowHeight` falls back to what the
   * row's wrapped cells need, so removing the override IS the measurement; an
   * unwrapped row hugs its single line at the default height, and a wrapped one
   * springs to however many lines it takes.
   */
  const autoFitRow = api.resetRowHeight;

  /**
   * Sticky placement and layer for one item. See the layer note in styles.ts;
   * these are the four that pin themselves, and they must be ordered by what
   * scrolls underneath what.
   *
   * The frozen band's ROW HEADER is the subtle one. It is sticky on both axes,
   * so it sits where an ordinary row header sits and where a frozen cell sits —
   * and at the same layer as an ordinary row header, the scrolled rows won the
   * tie by coming later in the DOM. That is why the frozen row's number was
   * being painted over by whatever row happened to scroll past it.
   */
  const LAYER = {
    frozenCell: 3,
    rowHeader: 4,
    /** The frozen band's own row header, over the ordinary ones it overlaps. */
    frozenRowHeader: 5,
    colHeader: 6,
    corner: 7,
  } as const;

  function stickyStyleFor(
    isHeaderRow: boolean,
    isRowHeaderCol: boolean,
    frozenRowIdx?: number,
    frozenColIdx?: number,
  ): CSSProperties {
    const style: CSSProperties = {};
    let z = 0;
    const inFrozenRow = !isHeaderRow && frozenRowIdx !== undefined;

    if (isHeaderRow) {
      style.top = 0;
      z = Math.max(z, LAYER.colHeader);
    } else if (frozenRowIdx !== undefined) {
      style.top = HEADER_HEIGHT + (rowWindow.rowTop(frozenRowIdx) ?? 0);
      z = Math.max(z, LAYER.frozenCell);
    }
    if (isRowHeaderCol) {
      style.left = 0;
      z = Math.max(
        z,
        isHeaderRow
          ? LAYER.corner
          : inFrozenRow
            ? LAYER.frozenRowHeader
            : LAYER.rowHeader,
      );
    } else if (frozenColIdx !== undefined) {
      style.left = ROW_HEADER_WIDTH + colOffset(frozenColIdx);
      z = Math.max(z, isHeaderRow ? LAYER.colHeader : LAYER.frozenCell);
    }

    // The line that says where the freeze is. Without it the frozen band and
    // the rows sliding under it are the same surface with the same grid lines,
    // and the sheet reads as if row 1 simply repeats — which is exactly how it
    // looked. A shadow rather than a border: it costs no layout, it cannot be
    // overridden by a cell's own borders, and it draws over the neighbour
    // instead of pushing it.
    const edges: string[] = [];
    if (inFrozenRow && frozenRowIdx === frozenRows - 1) {
      edges.push(`0 ${FREEZE_LINE_PX}px 0 0 ${theme.freezeLine}`);
    }
    if (frozenColIdx !== undefined && frozenColIdx === frozenCols - 1) {
      edges.push(`${FREEZE_LINE_PX}px 0 0 0 ${theme.freezeLine}`);
    }
    if (edges.length > 0) style.boxShadow = edges.join(", ");

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
   * The cell under a pointer position, from raw client coordinates.
   *
   * Both drags run on this rather than on per-cell mouse events: those miss
   * cells during a fast drag, and they see nothing at all once the pointer
   * leaves the grid — which is exactly when auto-scroll has to keep working.
   * Coordinates are clamped, so a pointer held off the left edge reads as
   * column 0 rather than as nothing.
   */
  function hitTest(clientX: number, clientY: number) {
    const el = containerRef.current;
    if (!el) return null;
    const box = el.getBoundingClientRect();
    const x = clientX - box.left + el.scrollLeft - ROW_HEADER_WIDTH;
    const y = clientY - box.top + el.scrollTop - HEADER_HEIGHT;

    return { row: Math.max(0, rowWindow.rowAt(y)), col: Math.max(0, colAt(x)) };
  }

  // Latest-render copies of everything the window-level drag handlers need. The
  // listeners are mounted once and outlive any single render, so reading them
  // through refs is what keeps a long drag from extending against a stale sheet.
  dragMoveRef.current = (clientX: number, clientY: number) => {
    const hit = hitTest(clientX, clientY);
    if (!hit) return;
    if (dragRef.current === "fill") {
      fill.moveTo(hit.row, hit.col);
      return;
    }
    // Dragging across headers keeps the band full-width or full-height and only
    // moves the end along that axis.
    if (dragRef.current === "rows") {
      api.extendTo(hit.row, sheet.numCols - 1);
      return;
    }
    if (dragRef.current === "cols") {
      api.extendTo(sheet.numRows - 1, hit.col);
      return;
    }
    // Dragging after a reference click grows that reference into a range.
    if (api.formulaRefs.active) api.formulaRefs.extendPickTo(hit.row, hit.col);
    else api.extendTo(hit.row, hit.col);
  };
  dragEndRef.current = () => {
    if (dragRef.current === "fill") {
      fill.commit(api.updateSheet, {
        sheet: api.sheet,
        onReject: api.setStatus,
      });
    }
  };

  /**
   * Clicking a row or column header, with the modifiers Excel gives it: plain
   * replaces the selection, Shift extends across to the header clicked, Ctrl
   * banks the current selection and adds this band — or removes it, if it is
   * already selected. The drag that may follow is picked up by the window
   * listeners, which extend the band across the headers the pointer crosses.
   */
  function selectBand(
    axis: "row" | "col",
    index: number,
    e: { shiftKey: boolean; ctrlKey: boolean; metaKey: boolean },
  ) {
    const band =
      axis === "row"
        ? { r1: index, c1: 0, r2: index, c2: sheet.numCols - 1 }
        : { r1: 0, c1: index, r2: sheet.numRows - 1, c2: index };

    if (e.ctrlKey || e.metaKey) {
      if (!api.removeRangeAt(band.r1, band.c1)) api.startNewRange(band);
    } else if (e.shiftKey) {
      api.extendTo(band.r2, band.c2);
    } else {
      api.select(band);
    }
    dragRef.current = axis === "row" ? "rows" : "cols";
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
          if (e.button !== 0) return;
          e.preventDefault();
          focusRef.current?.focus();
          selectBand("row", r, e);
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
        {windowCols.map((c) => {
          const key = cellKey(absRow, c);
          const selected = api.isSelected(absRow, c);
          return (
            <Cell
              key={c}
              row={absRow}
              col={c}
              gridRow={gridRow}
              styleRef={sheet.styles[key]}
              stylePreview={
                hasStylePreview && selected ? stylePreviewPatch : undefined
              }
              stickyStyle={stickyStyleFor(
                false,
                false,
                frozen ? absRow : undefined,
                c < frozenCols ? c : undefined,
              )}
            />
          );
        })}
      </div>
    );
  }

  const gridRender = useMemo(
    () => (renderCellContent ? { renderCellContent } : {}),
    [renderCellContent],
  );

  return (
    <GridRenderProvider value={gridRender}>
      <div
        // Four areas: the sheet, a channel per axis, and the corner between them.
        // The bars are siblings of the scroll container rather than inside it, so
        // they can never be scrolled away from or drawn over the cells.
        className={mergeClass(`${prefix}frame`, className)}
        style={{
          flex: 1,
          minHeight: 0,
          display: "grid",
          gridTemplateColumns: `minmax(0, 1fr) ${SCROLLBAR_SIZE}px`,
          gridTemplateRows: `minmax(0, 1fr) ${SCROLLBAR_SIZE}px`,
          ...style,
        }}
      >
        <div
          ref={containerRef}
          id={scrollerId}
          role="grid"
          aria-rowcount={sheet.numRows}
          aria-colcount={sheet.numCols}
          aria-label={sheet.name}
          className={mergeClass(`${prefix}scroller`, scrollerClassName)}
          style={{
            overflow: "auto",
            position: "relative",
            cursor: fill.dragging ? "crosshair" : undefined,
            ...scrollerStyle,
          }}
          onScroll={(e) => {
            const el = e.target as HTMLDivElement;
            setScrollTop(el.scrollTop);
            setScrollLeft(el.scrollLeft);
          }}
          // A mousedown that landed on a cell begins a selection drag. The cell
          // itself has already set the anchor by the time this bubbles up; all
          // that is left is to say a drag is running, so the window listeners
          // start extending. The fill handle stops propagation, so it never
          // arrives here — that drag announces itself through `fill.dragging`.
          onMouseDown={(e) => {
            if (e.button !== 0) return;
            const cell = (e.target as HTMLElement).closest?.(
              "[data-row][data-col]",
            );
            if (cell) dragRef.current = "select";
          }}
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
                    if (e.button !== 0) return;
                    e.preventDefault();
                    focusRef.current?.focus();
                    selectBand("col", c, e);
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
                      <span className={`${prefix}headlabel`}>
                        {sheet.colLabels[c] ?? colToLetters(c)}
                      </span>
                      <button
                        type="button"
                        aria-label={`Sort and filter column ${colToLetters(c)}`}
                        className={`${prefix}headmenu${filtered ? ` ${prefix}headfiltered` : ""}`}
                        onMouseDown={(e) => e.stopPropagation()}
                        onClick={(e) => {
                          e.stopPropagation();
                          // currentTarget, not target: a click landing on the SVG
                          // would otherwise place the menu against the icon's box.
                          const box = e.currentTarget.getBoundingClientRect();
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

            {/* An outline around each selected range. The tint on the cells says
            which cells; this says where each range begins and ends, which is the
            difference between a multi-range selection and a scattering of
            shaded cells. The primary range is drawn heavier. */}
            {api.ranges.map((range, index) => {
              const box = rectFor(normalizeRange(range));
              if (!box) return null;
              return (
                <div
                  // The identity of a range IS its position in the list: two of
                  // them can cover the same cells, and reordering is not a thing
                  // that happens — a range is appended or the lot is cleared.
                  // biome-ignore lint/suspicious/noArrayIndexKey: see above
                  key={index}
                  aria-hidden="true"
                  className={`${prefix}selbox`}
                  style={{ position: "absolute", ...box }}
                />
              );
            })}

            {/* What was copied, dashed, until it is pasted or dismissed. The only
            thing on screen that says a copy happened at all — and, once the
            selection has moved to the paste target, the only thing saying where
            the content is coming from. */}
            {api.clipboard.copiedRanges.map((range, index) => {
              const box = rectFor(normalizeRange(range));
              if (!box) return null;
              return (
                <div
                  // biome-ignore lint/suspicious/noArrayIndexKey: as above
                  key={index}
                  aria-hidden="true"
                  className={`${prefix}marquee`}
                  style={{ position: "absolute", ...box }}
                />
              );
            })}

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
                    // Below the sticky headers and the frozen bands, like every
                    // other overlay — see the layer note in styles.ts.
                    zIndex: 1,
                  }}
                />
              );
            })}
          </div>

          {/* End-of-sheet slot. Outside the grid element so it is not a grid item
            competing for a track, but inside the scroller so it sits after the
            last row and scrolls with it. */}
          {children}
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
    </GridRenderProvider>
  );
}
