"use client";

/**
 * Keeps a text input's caret and the edit state's `caret` in agreement.
 *
 * Needed in both directions:
 *
 * - input -> state, because clicking a cell mid-formula has to know where the
 *   caret was, and by then the click has already happened;
 * - state -> input, because picking a reference rewrites the text from outside
 *   the input, and React restores the caret to the end of a controlled value.
 *   Without this, `=SUM(A1|,B2)` would jump the caret past `B2` on every pick.
 *
 * Used by the in-cell editor and the formula bar, which are two views of one
 * edit and must not disagree about where the caret is.
 */
import { type RefObject, useCallback, useEffect, useRef } from "react";

export interface CaretBinding {
  ref: RefObject<HTMLInputElement | null>;
  /** Wire to the input's `onSelect`; fires for clicks, arrows, and typing. */
  onSelect(event: { currentTarget: HTMLInputElement }): void;
}

export function useCaretBinding(
  caret: number | undefined,
  setCaret: (caret: number) => void,
): CaretBinding {
  const ref = useRef<HTMLInputElement | null>(null);

  useEffect(() => {
    const el = ref.current;
    if (!el || caret === undefined) return;
    // Guarded: assigning unconditionally would fight the user's own selection
    // and collapse a drag-select inside the input on every render.
    if (el.selectionStart === caret && el.selectionEnd === caret) return;
    if (document.activeElement !== el) return;
    el.setSelectionRange(caret, caret);
  }, [caret]);

  const onSelect = useCallback(
    (event: { currentTarget: HTMLInputElement }) => {
      const position = event.currentTarget.selectionStart;
      if (position !== null) setCaret(position);
    },
    [setCaret],
  );

  return { ref, onSelect };
}
