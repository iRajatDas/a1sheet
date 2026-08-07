"use client";

/**
 * Renders a DOM element or merges onto a single child when `asChild`.
 * Shared by toolbar buttons, menu items, and other leaf primitives.
 */
import {
  type CSSProperties,
  type ElementType,
  createElement,
  forwardRef,
  type HTMLAttributes,
  type ReactNode,
  type Ref,
} from "react";
import { Slot } from "../Slot.js";
import type { PrimitiveProps } from "./types.js";

export type PartProps<T extends ElementType = "div"> = PrimitiveProps &
  Omit<HTMLAttributes<HTMLElement>, "className" | "style"> & {
    children?: ReactNode;
    /** Used when `asChild` is false. Defaults to `"div"`. */
    defaultElement?: T;
  };

export const Part = forwardRef(function Part<T extends ElementType = "div">(
  {
    asChild,
    className,
    style,
    children,
    defaultElement = "div" as T,
    ...props
  }: PartProps<T>,
  ref: React.Ref<HTMLElement>,
) {
  if (asChild) {
    return (
      <Slot ref={ref} className={className} style={style} {...props}>
        {children}
      </Slot>
    );
  }

  return createElement(
    defaultElement,
    { ref, className, style, ...props },
    children,
  );
}) as <T extends ElementType = "div">(
  props: PartProps<T> & { ref?: React.Ref<HTMLElement> },
) => ReactNode;
