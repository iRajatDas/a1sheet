"use client";

/**
 * Shared view-layer types.
 *
 * These used to live in `components/props.ts` alongside a `BaseProps` bag that
 * every component destructured. That bag is gone — primitives read from context —
 * so what remains is the transient UI state shapes and the `asChild` convention.
 */

/** Anchor for a floating menu, in viewport coordinates. */
export interface MenuAnchor {
  x: number;
  y: number;
}

export interface ColumnMenuState extends MenuAnchor {
  col: number;
}

export interface ContextMenuState extends MenuAnchor {
  row: number;
  col: number;
}

/** Inline rename target — column header, row header, or sheet tab. */
export interface RenamingState {
  type: "col" | "row" | "sheet";
  index: number;
  value: string;
}

/**
 * Mixed into any primitive that renders exactly one DOM element.
 *
 * With `asChild`, the primitive merges its props onto the child you provide instead
 * of rendering its own element — no `as` prop, no wrapper div.
 */
export interface AsChildProps {
  asChild?: boolean;
}
