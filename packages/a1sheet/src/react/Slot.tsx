"use client";

/**
 * `asChild` support, in the Radix sense: instead of rendering its own
 * element, a primitive merges its props onto the single child element it is given.
 *
 * That is what lets a consumer swap the underlying element or drop in their own
 * component without the library growing an `as` prop or a wrapper div:
 *
 *   <Sheet.ExportButton asChild>
 *     <MyFancyButton>Download</MyFancyButton>
 *   </Sheet.ExportButton>
 *
 * Merge rules, chosen to be unsurprising:
 * - event handlers: ours runs first, then the child's
 * - className: concatenated, child last so it can override
 * - style: shallow-merged, child wins
 * - everything else: child wins (it is the more specific intent)
 * - refs: both are attached
 */
import {
  Children,
  type CSSProperties,
  cloneElement,
  forwardRef,
  isValidElement,
  type ReactNode,
  type Ref,
} from "react";
import { InvalidArgumentError } from "../errors.js";

export type SlotProps = {
  children?: ReactNode;
} & Record<string, unknown>;

type AnyProps = Record<string, unknown>;

function composeRefs<T>(...refs: (Ref<T> | undefined)[]) {
  return (node: T | null) => {
    for (const ref of refs) {
      if (!ref) continue;
      if (typeof ref === "function") ref(node);
      else (ref as { current: T | null }).current = node;
    }
  };
}

function mergeProps(ours: AnyProps, theirs: AnyProps): AnyProps {
  const merged: AnyProps = { ...ours, ...theirs };

  for (const key of Object.keys(theirs)) {
    const ourValue = ours[key];
    const theirValue = theirs[key];

    // Event handlers: call ours, then theirs.
    if (/^on[A-Z]/.test(key)) {
      if (typeof ourValue === "function" && typeof theirValue === "function") {
        merged[key] = (...args: unknown[]) => {
          (ourValue as (...a: unknown[]) => void)(...args);
          (theirValue as (...a: unknown[]) => void)(...args);
        };
      }
      continue;
    }

    if (key === "className") {
      merged[key] = [ourValue, theirValue].filter(Boolean).join(" ");
      continue;
    }

    if (key === "style") {
      merged[key] = {
        ...(ourValue as CSSProperties),
        ...(theirValue as CSSProperties),
      };
    }
  }

  return merged;
}

/**
 * Renders its single child with `props` merged in.
 *
 * Throws on zero or multiple children: `asChild` with anything other than exactly
 * one element is a programmer error, and silently rendering the first child would
 * hide it.
 */
export const Slot = forwardRef<HTMLElement, SlotProps>(
  function Slot(props, forwardedRef) {
    // Destructuring off a type with an index signature widens `children` to unknown,
    // so it is read explicitly rather than pulled out in the parameter list.
    const { children, ...rest } = props;
    const ours = rest as AnyProps;
    const child = Children.toArray(children as ReactNode).filter(isValidElement);

    if (child.length !== 1) {
      throw new InvalidArgumentError(
        "asChild children",
        "exactly one React element",
        `${child.length} elements`,
      );
    }

    const only = child[0] as React.ReactElement<AnyProps> & {
      ref?: Ref<HTMLElement>;
    };

    return cloneElement(only, {
      ...mergeProps(ours, only.props),
      ref: composeRefs(forwardedRef, only.ref),
    } as AnyProps);
  },
);
