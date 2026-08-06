"use client";

/**
 * Text width measurement, for auto-fitting a column to its contents.
 *
 * A canvas 2D context rather than a hidden DOM node on purpose: `measureText`
 * costs roughly a microsecond and triggers no layout, where inserting a element
 * per candidate string and reading `offsetWidth` forces a synchronous reflow
 * each time. Auto-fit measures thousands of strings on one double-click, so the
 * difference is the difference between instant and a visible stall.
 *
 * The context lives in a ref rather than at module scope: a module-level canvas
 * is global mutable state, and creating one at import time breaks SSR.
 */
import { useCallback, useRef } from "react";

/** Measures `text` in `font` (any CSS `font` shorthand). Returns 0 where canvas is unavailable. */
export type MeasureText = (text: string, font: string) => number;

export function useTextMeasurer(): MeasureText {
  const ctxRef = useRef<CanvasRenderingContext2D | null | undefined>(undefined);

  return useCallback((text: string, font: string) => {
    if (ctxRef.current === undefined) {
      ctxRef.current =
        typeof document === "undefined"
          ? null
          : document.createElement("canvas").getContext("2d");
    }
    const ctx = ctxRef.current;
    // No canvas means no measurement: SSR, or a test environment that stubs it.
    // Callers fall back to leaving the size alone rather than collapsing it.
    if (!ctx || typeof ctx.measureText !== "function") return 0;
    if (font && ctx.font !== font) ctx.font = font;
    return ctx.measureText(text).width;
  }, []);
}
