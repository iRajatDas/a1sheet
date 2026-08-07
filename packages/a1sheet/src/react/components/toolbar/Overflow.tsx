"use client";

import { type ReactNode, useEffect, useState } from "react";
import { useSheetContext } from "../../context.js";
import { mergeClass } from "../../primitives/mergeClass.js";
import type { PrimitiveProps } from "../../primitives/types.js";

const NARROW_QUERY = "(max-width: 640px)";

export interface ToolbarOverflowProps extends PrimitiveProps {
  /** Label for the collapsed menu trigger. */
  menuLabel?: string;
  children?: ReactNode;
}

/**
 * On wide viewports renders children inline. On narrow viewports collapses them
 * behind a single menu button — compose secondary tool groups here.
 */
export function ToolbarOverflow({
  children,
  className,
  style,
  menuLabel = "More tools",
}: ToolbarOverflowProps): ReactNode {
  const { prefix, ui } = useSheetContext("Sheet.Toolbar.Overflow");
  const [narrow, setNarrow] = useState(false);
  const [open, setOpen] = useState(false);

  useEffect(() => {
    if (typeof window === "undefined" || !window.matchMedia) return;
    const mq = window.matchMedia(NARROW_QUERY);
    const update = () => setNarrow(mq.matches);
    update();
    mq.addEventListener("change", update);
    return () => mq.removeEventListener("change", update);
  }, []);

  useEffect(() => {
    if (!open) return;
    const close = () => {
      setOpen(false);
      ui.closeMenus();
    };
    window.addEventListener("click", close);
    return () => window.removeEventListener("click", close);
  }, [open, ui]);

  if (!narrow) {
    return (
      <div
        className={mergeClass(`${prefix}toolbar-overflow`, className)}
        style={{ display: "inline-flex", alignItems: "center", gap: 4, ...style }}
      >
        {children}
      </div>
    );
  }

  return (
    <div
      className={mergeClass(`${prefix}toolbar-overflow`, className)}
      style={{ display: "inline-flex", position: "relative", ...style }}
    >
      <button
        type="button"
        className={`${prefix}btn ${prefix}iconbtn`}
        aria-label={menuLabel}
        aria-haspopup="menu"
        aria-expanded={open}
        onClick={(e) => {
          e.stopPropagation();
          setOpen((v) => !v);
        }}
      >
        ⋯
      </button>
      {open && (
        <div
          className={`${prefix}menu`}
          role="menu"
          style={{ position: "absolute", top: "100%", right: 0, marginTop: 4, zIndex: 1000 }}
          onClick={(e) => e.stopPropagation()}
          onKeyDown={(e) => {
            if (e.key === "Escape") setOpen(false);
          }}
        >
          <div
            style={{
              display: "flex",
              flexDirection: "column",
              gap: 4,
              padding: 4,
            }}
          >
            {children}
          </div>
        </div>
      )}
    </div>
  );
}
