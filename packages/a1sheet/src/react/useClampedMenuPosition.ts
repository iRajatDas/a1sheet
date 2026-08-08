"use client";

/**
 * Positions a fixed menu at an anchor, then clamps it into the viewport after
 * layout so tall menus flip above the cursor and wide ones slide left.
 */
import {
  type CSSProperties,
  type RefObject,
  useLayoutEffect,
  useState,
} from "react";
import { clampMenuPosition, type MenuAnchorPoint } from "./menuPosition.js";

export function useClampedMenuPosition(
  anchor: MenuAnchorPoint | null,
  ref: RefObject<HTMLElement | null>,
  style?: CSSProperties,
): CSSProperties {
  const x = anchor?.x;
  const y = anchor?.y;
  const [pos, setPos] = useState({ left: x ?? 0, top: y ?? 0 });

  useLayoutEffect(() => {
    if (x === undefined || y === undefined) return;
    const el = ref.current;
    if (!el || typeof window === "undefined") {
      setPos({ left: x, top: y });
      return;
    }
    const rect = el.getBoundingClientRect();
    const next = clampMenuPosition({
      x,
      y,
      width: Math.max(rect.width, el.offsetWidth),
      height: Math.max(rect.height, el.offsetHeight),
      viewportWidth: window.innerWidth,
      viewportHeight: window.innerHeight,
    });
    setPos((prev) =>
      prev.left === next.left && prev.top === next.top ? prev : next,
    );
  }, [x, y, ref]);

  return { ...style, left: pos.left, top: pos.top };
}
