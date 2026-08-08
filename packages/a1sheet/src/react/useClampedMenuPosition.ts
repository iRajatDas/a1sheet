"use client";

/**
 * Positions a fixed menu at an anchor inside the sheet's collision bounds, then
 * re-clamps when the menu's size changes (ResizeObserver). Portaled menus use
 * this so host `overflow: hidden` cannot clip them.
 */
import {
  type CSSProperties,
  type RefObject,
  useLayoutEffect,
  useState,
} from "react";
import { useSheetContext } from "./context.js";
import {
  clampMenuPosition,
  collisionBoundsFromElement,
  type MenuAnchorPoint,
} from "./menuPosition.js";

export function useClampedMenuPosition(
  anchor: MenuAnchorPoint | null,
  ref: RefObject<HTMLElement | null>,
  style?: CSSProperties,
): CSSProperties {
  const { rootRef } = useSheetContext("useClampedMenuPosition");
  const x = anchor?.x;
  const y = anchor?.y;
  const [pos, setPos] = useState({ left: x ?? 0, top: y ?? 0 });
  const [ready, setReady] = useState(false);

  useLayoutEffect(() => {
    if (x === undefined || y === undefined) {
      setReady(false);
      return;
    }

    const el = ref.current;
    const apply = () => {
      const node = ref.current;
      if (!node || typeof window === "undefined") {
        setPos({ left: x, top: y });
        setReady(true);
        return;
      }
      const rect = node.getBoundingClientRect();
      const next = clampMenuPosition({
        x,
        y,
        width: Math.max(rect.width, node.offsetWidth),
        height: Math.max(rect.height, node.offsetHeight),
        bounds: collisionBoundsFromElement(rootRef.current),
      });
      setPos((prev) =>
        prev.left === next.left && prev.top === next.top ? prev : next,
      );
      setReady(true);
    };

    apply();

    const node = el;
    if (!node || typeof ResizeObserver === "undefined") return;

    const ro = new ResizeObserver(() => apply());
    ro.observe(node);
    return () => ro.disconnect();
  }, [x, y, ref, rootRef]);

  return {
    ...style,
    position: "fixed",
    left: pos.left,
    top: pos.top,
    zIndex: (style?.zIndex as number | undefined) ?? 1000,
    visibility: ready ? "visible" : "hidden",
  };
}
