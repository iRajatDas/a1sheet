/**
 * Keeping the cell the keyboard is moving on screen.
 *
 * One axis at a time, and pure, because the interesting part is arithmetic: the
 * frozen headers float over the content rather than sitting beside it, so the
 * first `lead` pixels of the visible box are covered and a cell is only clear of
 * them when the scroll offset is at or above the cell's own start.
 */
export interface RevealInput {
  /** Current scroll offset of the container. */
  offset: number;
  /** Visible length of the container along this axis. */
  viewport: number;
  /** Start of the cell, measured from the beginning of the content. */
  start: number;
  /** Length of the cell along this axis. */
  size: number;
  /** Thickness of the header floating over the start of the box. */
  lead: number;
}

/**
 * The offset the container should have. Equal to `offset` when the cell is
 * already visible — scrolling a visible cell into view would drag the sheet
 * around on every keystroke.
 */
export function revealOffset({
  offset,
  viewport,
  start,
  size,
  lead,
}: RevealInput): number {
  if (start < offset) return start;
  const end = start + size + lead;
  if (end > offset + viewport) return end - viewport;
  return offset;
}

export interface TrackResizeScrollInput {
  /** Current scroll offset along this axis. */
  offset: number;
  /** Content start of the resized track. */
  trackStart: number;
  /** Track size before the edit. */
  prevSize: number;
  /** Track size after the edit. */
  nextSize: number;
}

/**
 * Keep the same content under the viewport when a track fully before the
 * scroll origin changes size (Excel / Sheets behaviour). Tracks that intersect
 * or sit after the viewport leave the offset alone — only the max scroll extent
 * may shrink, which the browser clamps.
 */
export function compensateScrollForTrackResize({
  offset,
  trackStart,
  prevSize,
  nextSize,
}: TrackResizeScrollInput): number {
  const delta = nextSize - prevSize;
  if (delta === 0) return offset;
  if (trackStart + prevSize <= offset) {
    return Math.max(0, offset + delta);
  }
  return offset;
}
