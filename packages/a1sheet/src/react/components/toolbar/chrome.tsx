"use client";

import type { CSSProperties, ReactNode } from "react";
import { useSheetContext } from "../../context.js";
import { mergeClass } from "../../primitives/mergeClass.js";
import type { PrimitiveProps } from "../../primitives/types.js";

export interface ToolbarSeparatorProps extends PrimitiveProps {}

export function ToolbarSeparator({
  className,
  style,
}: ToolbarSeparatorProps = {}): ReactNode {
  const { prefix } = useSheetContext("Sheet.Toolbar.Separator");
  return (
    <span
      className={mergeClass(`${prefix}sep`, className)}
      style={style}
      aria-hidden="true"
    />
  );
}

export interface ToolbarGroupProps extends PrimitiveProps {
  children?: ReactNode;
}

/** Keeps a set of controls on one line without wrapping between them. */
export function ToolbarGroup({
  children,
  className,
  style,
}: ToolbarGroupProps): ReactNode {
  return (
    <div
      className={mergeClass(className)}
      style={{
        display: "inline-flex",
        alignItems: "center",
        gap: 4,
        flexWrap: "nowrap",
        ...style,
      }}
      role="group"
    >
      {children}
    </div>
  );
}

export interface ToolbarStatusProps extends PrimitiveProps {}

export function ToolbarStatus({
  className,
  style,
}: ToolbarStatusProps = {}): ReactNode {
  const { api, theme } = useSheetContext("Sheet.Toolbar.Status");
  if (!api.status) return null;
  return (
    <span
      className={className}
      style={{ fontSize: 12, color: theme.headerText, ...style }}
      role="status"
    >
      {api.status}
    </span>
  );
}
