"use client";

/**
 * Scrolling while a drag is held against the edge of the grid.
 *
 * Selecting a range taller than the screen is otherwise impossible with the
 * mouse: the pointer reaches the last visible row and the selection stops there.
 * Every spreadsheet solves it the same way — hold near an edge and the sheet
 * comes to you, faster the further past the edge you push.
 *
 * The loop calls back on every frame with the pointer's last position, even
 * though the pointer has not moved. That is the point: the content moved under
 * it, so the cell beneath the pointer is a different one and the caller must
 * re-run its hit test to keep extending.
 */
import { type RefObject, useCallback, useEffect, useRef } from "react";

/** How close to an edge the pointer must come before the sheet starts moving. */
const EDGE_ZONE_PX = 24;
/** Pixels per frame at the point the zone is fully entered, and at its fastest. */
const MIN_SPEED_PX = 2;
const MAX_SPEED_PX = 28;
/** How far past the edge the pointer must go for full speed. */
const RAMP_PX = 120;

/**
 * Scroll speed for one axis: negative to scroll back, positive forward, zero
 * when the pointer is comfortably inside.
 *
 * Pure, and exported because the ramp is the whole feel of the interaction and a
 * feel is not testable through a DOM event.
 */
export function edgeVelocity(pos: number, min: number, max: number): number {
  const past =
    pos < min + EDGE_ZONE_PX
      ? pos - (min + EDGE_ZONE_PX)
      : pos > max - EDGE_ZONE_PX
        ? pos - (max - EDGE_ZONE_PX)
        : 0;
  if (past === 0) return 0;

  const ramp = Math.min(1, Math.abs(past) / RAMP_PX);
  const speed = MIN_SPEED_PX + (MAX_SPEED_PX - MIN_SPEED_PX) * ramp;
  return Math.sign(past) * speed;
}

export interface UseAutoScrollResult {
  /** Report the pointer. Starts the loop if the pointer is near an edge. */
  track(clientX: number, clientY: number): void;
  /** End the drag. Safe to call when nothing is running. */
  stop(): void;
}

/**
 * @param scroller The element to scroll — the grid's scroll container.
 * @param onFrame Called after each scroll step with the pointer's last position.
 */
export function useAutoScroll(
  scroller: RefObject<HTMLElement | null>,
  onFrame: (clientX: number, clientY: number) => void,
): UseAutoScrollResult {
  const pointer = useRef<{ x: number; y: number } | null>(null);
  const frame = useRef<number | null>(null);
  // Read through a ref so a caller does not have to memoize the callback: the
  // loop outlives any one render, and capturing the first one would leave it
  // extending the selection with a stale hit test.
  const latest = useRef(onFrame);
  latest.current = onFrame;

  const stop = useCallback(() => {
    pointer.current = null;
    if (frame.current !== null && typeof cancelAnimationFrame !== "undefined") {
      cancelAnimationFrame(frame.current);
    }
    frame.current = null;
  }, []);

  const step = useCallback(() => {
    frame.current = null;
    const el = scroller.current;
    const at = pointer.current;
    if (!el || !at) return;

    const box = el.getBoundingClientRect();
    const dx = edgeVelocity(at.x, box.left, box.right);
    const dy = edgeVelocity(at.y, box.top, box.bottom);
    if (dx === 0 && dy === 0) return;

    el.scrollLeft += dx;
    el.scrollTop += dy;
    latest.current(at.x, at.y);

    if (typeof requestAnimationFrame !== "undefined") {
      frame.current = requestAnimationFrame(step);
    }
  }, [scroller]);

  const track = useCallback(
    (clientX: number, clientY: number) => {
      pointer.current = { x: clientX, y: clientY };
      if (frame.current !== null || typeof requestAnimationFrame === "undefined") {
        return;
      }
      frame.current = requestAnimationFrame(step);
    },
    [step],
  );

  // A drag interrupted by unmounting must not leave a frame scheduled against a
  // detached element.
  useEffect(() => stop, [stop]);

  return { track, stop };
}
