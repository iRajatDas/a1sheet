/**
 * Keep a fixed-position menu inside the viewport.
 *
 * Prefer the click point as the top-left. If that would overflow the bottom,
 * flip above the cursor; if that still overflows, clamp. Same idea horizontally
 * (slide left rather than flip, matching common desktop context menus).
 */

export interface MenuAnchorPoint {
  x: number;
  y: number;
}

export interface ClampMenuPositionOptions extends MenuAnchorPoint {
  width: number;
  height: number;
  viewportWidth: number;
  viewportHeight: number;
  /** Inset from the viewport edges. */
  margin?: number;
}

export interface ClampedMenuPosition {
  left: number;
  top: number;
}

export function clampMenuPosition(
  options: ClampMenuPositionOptions,
): ClampedMenuPosition {
  const margin = options.margin ?? 8;
  const maxLeft = Math.max(margin, options.viewportWidth - options.width - margin);
  const maxTop = Math.max(margin, options.viewportHeight - options.height - margin);

  let left = options.x;
  let top = options.y;

  if (left + options.width > options.viewportWidth - margin) {
    left = options.x - options.width;
  }
  if (left < margin) left = margin;
  if (left > maxLeft) left = maxLeft;

  if (top + options.height > options.viewportHeight - margin) {
    top = options.y - options.height;
  }
  if (top < margin) top = margin;
  if (top > maxTop) top = maxTop;

  return { left, top };
}
