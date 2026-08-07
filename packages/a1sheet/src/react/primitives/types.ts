"use client";

import type { CSSProperties, ReactNode } from "react";
import type { FormulaValue } from "../../formula/values.js";
import type { StyleObject } from "../../model/types.js";
import type { AsChildProps } from "../types.js";

/** Mixed into any primitive that exposes a root DOM element. */
export interface PrimitiveProps extends AsChildProps {
  className?: string;
  style?: CSSProperties;
}

/** Props for a custom cell body renderer. Interaction stays in `Cell`. */
export interface CellContentProps {
  row: number;
  col: number;
  display: string;
  raw: string;
  value: FormulaValue;
  style: StyleObject;
  isEditing: boolean;
  isSelected: boolean;
  isActive: boolean;
  isLocked: boolean;
}

/** Injectable renderers, set on `<Sheet.Root components={…}>`. */
export interface SheetComponents {
  CellContent?: (props: CellContentProps) => ReactNode;
}
