/**
 * Keep a fixed-position menu inside collision bounds (usually the sheet root
 * intersected with the visual viewport — not the bare window).
 *
 * Prefer the click point as the top-left. If that would overflow the bottom of
 * the bounds, flip above the cursor; if that still overflows, clamp. Horizontally
 * slide left when the right edge would overflow.
 */

export interface MenuAnchorPoint {
  x: number;
  y: number;
}

/** Axis-aligned box in viewport (client) coordinates. */
export interface MenuBounds {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

export interface ClampMenuPositionOptions extends MenuAnchorPoint {
  width: number;
  height: number;
  bounds: MenuBounds;
  /** Inset from the bounds edges. */
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
  const { bounds, width, height } = options;
  const maxLeft = Math.max(bounds.left + margin, bounds.right - width - margin);
  const maxTop = Math.max(bounds.top + margin, bounds.bottom - height - margin);

  let left = options.x;
  let top = options.y;

  if (left + width > bounds.right - margin) {
    left = options.x - width;
  }
  if (left < bounds.left + margin) left = bounds.left + margin;
  if (left > maxLeft) left = maxLeft;

  if (top + height > bounds.bottom - margin) {
    top = options.y - height;
  }
  if (top < bounds.top + margin) top = bounds.top + margin;
  if (top > maxTop) top = maxTop;

  return { left, top };
}

/**
 * Collision box for menus: the sheet root's client rect clipped to the visual
 * viewport so a short split pane flips correctly even when the window still has
 * space below.
 */
export function collisionBoundsFromElement(el: Element | null): MenuBounds {
  if (typeof window === "undefined") {
    return { left: 0, top: 0, right: 0, bottom: 0 };
  }
  const vv = window.visualViewport;
  const winW = vv?.width ?? window.innerWidth;
  const winH = vv?.height ?? window.innerHeight;
  const winLeft = vv?.offsetLeft ?? 0;
  const winTop = vv?.offsetTop ?? 0;
  const winRight = winLeft + winW;
  const winBottom = winTop + winH;
  const hasWindow = winW > 0 && winH > 0;

  if (!el) {
    return hasWindow
      ? { left: winLeft, top: winTop, right: winRight, bottom: winBottom }
      : { left: 0, top: 0, right: 0, bottom: 0 };
  }

  const r = el.getBoundingClientRect();
  if (!hasWindow) {
    return { left: r.left, top: r.top, right: r.right, bottom: r.bottom };
  }
  return {
    left: Math.max(r.left, winLeft),
    top: Math.max(r.top, winTop),
    right: Math.min(r.right, winRight),
    bottom: Math.min(r.bottom, winBottom),
  };
}
