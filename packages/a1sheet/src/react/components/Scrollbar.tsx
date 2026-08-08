"use client";

/**
 * A scrollbar drawn by us, in a channel of its own beside the grid.
 *
 * The native scrollbar was the obvious thing to use and it does not work here.
 * macOS and mobile default to *overlay* scrollbars: they fade in over the
 * content while you scroll and take no layout space, so they cover the
 * rightmost column and the bottom row exactly while you are moving through
 * them, and the grid has no way to lay itself out around something with no
 * width. Styling them does not fix it either — `::-webkit-scrollbar` is
 * non-standard and Firefox ignores it, `scrollbar-width` offers `thin` and
 * `auto` and nothing else, and neither engine has a way to say "always
 * visible". The result is a control that looks and behaves differently on
 * every platform, in the one component where the scrollbar is a primary
 * instrument rather than an afterthought.
 *
 * So the scroll container hides its native bars and these are laid out
 * alongside it, one per axis, identical to each other and always present.
 * Sheets and Excel both do the same.
 *
 * Scrolling itself is still the browser's: this writes `scrollLeft`/
 * `scrollTop` and the resulting `scroll` event feeds back in as `offset`. The
 * wheel, trackpad, keyboard, and `scrollIntoView` therefore keep working
 * untouched, and the thumb follows them because it is a view of the scroll
 * position rather than the owner of it.
 */
import { type ReactNode, useCallback, useEffect, useRef, useState } from "react";
import { SCROLLBAR_MIN_THUMB } from "../constants.js";

export interface ScrollbarProps {
  orientation: "vertical" | "horizontal";
  /** Length of the visible window along this axis, which is the track length. */
  viewport: number;
  /** Length of the whole sheet along this axis. */
  content: number;
  /** Current scroll offset along this axis. */
  offset: number;
  /** Id of the element being scrolled, for `aria-controls`. */
  controls: string;
  /** Called with the offset to scroll to. Clamping is the caller's business. */
  onScrollTo(offset: number): void;
  prefix: string;
}

interface Drag {
  /** Pointer position when the drag started, along the scrolled axis. */
  start: number;
  /** Scroll offset when the drag started. */
  offset: number;
  /** Scroll pixels per pointer pixel, fixed for the life of the drag. */
  scale: number;
}

export function Scrollbar({
  orientation,
  viewport,
  content,
  offset,
  controls,
  onScrollTo,
  prefix,
}: ScrollbarProps): ReactNode {
  const vertical = orientation === "vertical";
  const dragRef = useRef<Drag | null>(null);
  const [dragging, setDragging] = useState(false);

  const scrollable = Math.max(0, content - viewport);
  // Proportional to how much of the sheet is on screen, then floored so it
  // stays grabbable — at 100k rows the honest proportion is under a pixel —
  // and capped at the track, because a sheet smaller than its window has a
  // proportion above 1 and would otherwise overhang the channel.
  const thumb =
    content > 0
      ? Math.min(
          viewport,
          Math.max(
            SCROLLBAR_MIN_THUMB,
            Math.round((viewport / content) * viewport),
          ),
        )
      : viewport;
  const travel = Math.max(0, viewport - thumb);
  const position = scrollable > 0 ? (offset / scrollable) * travel : 0;

  const startDrag = useCallback(
    (e: React.MouseEvent) => {
      // Only the primary button, and never the track's own handler as well.
      if (e.button !== 0) return;
      e.preventDefault();
      e.stopPropagation();
      dragRef.current = {
        start: vertical ? e.clientY : e.clientX,
        offset,
        // Zero travel means the sheet fits; the drag then moves nothing, which
        // is correct and avoids dividing by zero.
        scale: travel > 0 ? scrollable / travel : 0,
      };
      setDragging(true);
    },
    [vertical, offset, travel, scrollable],
  );

  /** Tracked on window so the drag survives the pointer leaving the bar. */
  useEffect(() => {
    if (!dragging) return;
    function onMove(e: MouseEvent) {
      const drag = dragRef.current;
      if (!drag) return;
      const moved = (vertical ? e.clientY : e.clientX) - drag.start;
      onScrollTo(drag.offset + moved * drag.scale);
    }
    function onUp() {
      dragRef.current = null;
      setDragging(false);
    }
    window.addEventListener("mousemove", onMove);
    window.addEventListener("mouseup", onUp);
    return () => {
      window.removeEventListener("mousemove", onMove);
      window.removeEventListener("mouseup", onUp);
    };
  }, [dragging, vertical, onScrollTo]);

  /** Clicking the track pages toward the click, as every platform scrollbar does. */
  const pageToward = useCallback(
    (e: React.MouseEvent<HTMLDivElement>) => {
      if (e.button !== 0) return;
      const box = e.currentTarget.getBoundingClientRect();
      const at = vertical ? e.clientY - box.top : e.clientX - box.left;
      onScrollTo(at < position ? offset - viewport : offset + viewport);
    },
    [vertical, position, offset, viewport, onScrollTo],
  );

  const length = `${thumb}px`;
  const shift = `translate${vertical ? "Y" : "X"}(${position}px)`;

  return (
    <div
      role="scrollbar"
      aria-orientation={orientation}
      aria-controls={controls}
      aria-valuemin={0}
      aria-valuemax={Math.round(scrollable)}
      aria-valuenow={Math.round(offset)}
      className={`${prefix}sbar ${prefix}sbar-${orientation}`}
      onMouseDown={pageToward}
    >
      <div
        className={`${prefix}sbthumb${dragging ? ` ${prefix}on` : ""}`}
        style={
          vertical
            ? { height: length, transform: shift }
            : { width: length, transform: shift }
        }
        onMouseDown={startDrag}
      />
    </div>
  );
}
