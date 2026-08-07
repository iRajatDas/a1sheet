"use client";

import {
  createContext,
  type ReactNode,
  useCallback,
  useContext,
  useMemo,
  useState,
} from "react";
import type { StyleObject } from "../model/types.js";

export interface StylePreviewState {
  patch: Partial<StyleObject>;
  setPreview(patch: Partial<StyleObject> | null): void;
  clearPreview(): void;
}

const StylePreviewContext = createContext<StylePreviewState>({
  patch: {},
  setPreview: () => {},
  clearPreview: () => {},
});

export function StylePreviewProvider({ children }: { children: ReactNode }) {
  const [patch, setPatch] = useState<Partial<StyleObject>>({});

  const setPreview = useCallback((next: Partial<StyleObject> | null) => {
    setPatch(next ?? {});
  }, []);

  const clearPreview = useCallback(() => {
    setPatch({});
  }, []);

  const value = useMemo(
    () => ({ patch, setPreview, clearPreview }),
    [patch, setPreview, clearPreview],
  );

  return (
    <StylePreviewContext.Provider value={value}>{children}</StylePreviewContext.Provider>
  );
}

export function useStylePreview(): StylePreviewState {
  return useContext(StylePreviewContext);
}
