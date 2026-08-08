"use client";

import { type ReactNode, useEffect } from "react";
import { mergeClass } from "../../primitives/mergeClass.js";
import type { PrimitiveProps } from "../../primitives/types.js";
import { Slot } from "../../Slot.js";

export interface MenuItemProps extends PrimitiveProps {
  disabled?: boolean;
  onSelect: () => void;
  children: ReactNode;
}

export function MenuItem({
  asChild,
  className,
  style,
  disabled,
  onSelect,
  children,
}: MenuItemProps): ReactNode {
  const shared = {
    className: mergeClass(className),
    style,
    role: "menuitem" as const,
    disabled,
    onClick: onSelect,
  };

  if (asChild) {
    return <Slot {...shared}>{children}</Slot>;
  }

  return (
    <button type="button" {...shared}>
      {children}
    </button>
  );
}

export interface MenuSeparatorProps extends PrimitiveProps {}

export function MenuSeparator({
  className,
  style,
}: MenuSeparatorProps = {}): ReactNode {
  return <hr className={className} style={style} />;
}

export function useMenuKeyboard(
  open: boolean,
  onClose: () => void,
  containerRef: React.RefObject<HTMLElement | null>,
): void {
  useEffect(() => {
    if (!open) return;
    const root = containerRef.current;
    if (!root) return;

    const items = () => [
      ...root.querySelectorAll<HTMLElement>('[role="menuitem"]:not([disabled])'),
    ];

    const focusAt = (index: number) => {
      const list = items();
      if (list.length === 0) return;
      const i = ((index % list.length) + list.length) % list.length;
      list[i]?.focus();
    };

    focusAt(0);

    const onKeyDown = (e: KeyboardEvent) => {
      const list = items();
      const current = list.indexOf(document.activeElement as HTMLElement);
      switch (e.key) {
        case "ArrowDown":
          e.preventDefault();
          focusAt(current + 1);
          break;
        case "ArrowUp":
          e.preventDefault();
          focusAt(current - 1);
          break;
        case "Home":
          e.preventDefault();
          focusAt(0);
          break;
        case "End":
          e.preventDefault();
          focusAt(list.length - 1);
          break;
        case "Escape":
          e.preventDefault();
          onClose();
          break;
        default:
          break;
      }
    };

    root.addEventListener("keydown", onKeyDown);
    return () => root.removeEventListener("keydown", onKeyDown);
  }, [open, onClose, containerRef]);
}
