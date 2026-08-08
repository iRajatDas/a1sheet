"use client";

import { forwardRef, type ReactNode } from "react";
import { useSheetContext } from "../../context.js";
import { mergeClass } from "../../primitives/mergeClass.js";
import type { PrimitiveProps } from "../../primitives/types.js";
import { Slot } from "../../Slot.js";

export interface IconButtonProps extends PrimitiveProps {
  label: string;
  pressed?: boolean;
  disabled?: boolean;
  onClick: () => void;
  children: ReactNode;
}

export const IconButton = forwardRef<HTMLButtonElement, IconButtonProps>(
  function IconButton(
    { asChild, className, style, label, pressed, disabled, onClick, children },
    ref,
  ) {
    const { prefix } = useSheetContext("Sheet.Toolbar.IconButton");
    const cls = mergeClass(
      `${prefix}btn`,
      `${prefix}iconbtn`,
      pressed ? `${prefix}on` : undefined,
      className,
    );
    const shared = {
      className: cls,
      style,
      title: label,
      "aria-label": label,
      "aria-pressed": pressed,
      disabled,
      onClick,
    };

    if (asChild) {
      return (
        <Slot ref={ref} {...shared}>
          {children}
        </Slot>
      );
    }

    return (
      <button ref={ref} type="button" {...shared}>
        {children}
      </button>
    );
  },
);
