"use client";

import { type ReactNode, useEffect, useMemo, useRef, useState } from "react";
import { useSheetContext } from "../../context.js";
import { ChevronDownIcon } from "../icons.js";
import { useMenuKeyboard } from "../menu/primitives.js";
import {
  type FontFace,
  fontById,
  noteRecentFont,
  recentFonts,
  SHEET_FONTS,
  THEME_FONT,
} from "./fonts.js";

function quoteFamily(name: string): string {
  return /^[\w-]+$/.test(name) ? name : `"${name.replace(/"/g, "")}"`;
}

export function FontFamilyMenu(): ReactNode {
  const { api, theme, prefix } = useSheetContext("Sheet.Toolbar.FontFamily");
  const [open, setOpen] = useState(false);
  const [query, setQuery] = useState("");
  const menuRef = useRef<HTMLDivElement>(null);
  const current = api.activeStyle.fontFamily ?? "";

  useMenuKeyboard(open, () => setOpen(false), menuRef);

  useEffect(() => {
    if (!open) return;
    const close = () => setOpen(false);
    window.addEventListener("click", close);
    return () => window.removeEventListener("click", close);
  }, [open]);

  const recent = recentFonts();
  const q = query.trim().toLowerCase();
  const matches = useMemo(() => {
    if (!q) return SHEET_FONTS;
    return SHEET_FONTS.filter((f) => f.label.toLowerCase().includes(q));
  }, [q]);

  const pick = (id: string) => {
    if (id) noteRecentFont(id);
    if (id) api.applyStyle({ fontFamily: id });
    else api.unsetStyle("fontFamily");
    setOpen(false);
    setQuery("");
  };

  const label = fontById(current).label;

  return (
    <div style={{ position: "relative", display: "inline-flex" }}>
      <button
        type="button"
        className={`${prefix}btn ${prefix}fontbtn`}
        aria-haspopup="listbox"
        aria-expanded={open}
        title="Font"
        onClick={(e) => {
          e.stopPropagation();
          setOpen((v) => !v);
        }}
      >
        <span className={`${prefix}fontbtn-label`}>{label}</span>
        <ChevronDownIcon />
      </button>
      {open && (
        <div
          ref={menuRef}
          className={`${prefix}menu ${prefix}fontmenu`}
          role="listbox"
          aria-label="Font family"
          onClick={(e) => e.stopPropagation()}
        >
          <input
            className={`${prefix}input ${prefix}fontmenu-search`}
            placeholder="Filter fonts"
            aria-label="Filter fonts"
            value={query}
            onChange={(e) => setQuery(e.target.value)}
          />
          {!q && recent.length > 0 && (
            <>
              <div className={`${prefix}fontmenu-heading`}>Recent</div>
              {recent.map((face) => (
                <FontMenuItem
                  key={`recent-${face.id}`}
                  prefix={prefix}
                  face={face}
                  themeFont={theme.fontFamily}
                  selected={face.id === current}
                  onPick={() => pick(face.id)}
                />
              ))}
              <div className={`${prefix}fontmenu-heading`}>All fonts</div>
            </>
          )}
          {matches.map((face) => (
            <FontMenuItem
              key={face.id || "theme"}
              prefix={prefix}
              face={face}
              themeFont={theme.fontFamily}
              selected={face.id === current}
              onPick={() => pick(face.id)}
            />
          ))}
          {matches.length === 0 && (
            <div className={`${prefix}fontmenu-empty`}>No matching fonts</div>
          )}
        </div>
      )}
    </div>
  );
}

function FontMenuItem({
  prefix,
  face,
  themeFont,
  selected,
  onPick,
}: {
  prefix: string;
  face: FontFace;
  themeFont: string;
  selected: boolean;
  onPick: () => void;
}) {
  return (
    <button
      type="button"
      role="option"
      aria-selected={selected}
      className={`${prefix}fontmenu-item${selected ? ` ${prefix}on` : ""}`}
      style={{
        fontFamily: face.id ? `${quoteFamily(face.id)}, ${themeFont}` : themeFont,
      }}
      onClick={onPick}
    >
      {face.label}
    </button>
  );
}
