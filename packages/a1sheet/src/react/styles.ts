/**
 * The single stylesheet, emitted as a string into one `<style>` tag by
 * `<Spreadsheet />`. Ported from the `<style>` block in ref/Spreadsheet.jsx.
 *
 * Not a `.css` file on purpose: a CSS import would force consumers to have a CSS
 * loader configured, breaking the drop-in requirement. Class names are prefixed
 * so they cannot collide with the host app.
 */
import { SCROLLBAR_SIZE } from "./constants.js";
import type { Theme } from "./theme.js";

/** Transparent border around the thumb, leaving a 6px bar in a 14px channel. */
const SCROLLBAR_THUMB_INSET = 4;

export function buildCss(prefix: string, t: Theme): string {
  const p = prefix;
  return `
.${p}cell { border-right: 1px solid ${t.border}; border-bottom: 1px solid ${t.border};
  overflow: hidden; white-space: nowrap; display: flex; align-items: center;
  padding: 0 6px; font-size: ${t.fontSize}; font-family: ${t.fontFamily};
  cursor: cell; position: relative; box-sizing: border-box; user-select: none; }
/* Tints are painted as an overlay, never with the background property.
   selectedBg is deliberately translucent, and a background declaration REPLACES
   the base color rather than compositing over it — so setting it directly turned
   every tinted element 90% transparent, which on a sticky header or a frozen row
   means the scrolled content shows straight through. An ::after layer composites
   instead, which also makes the tint work over a cell's own fill and over the
   locked stripes.
   Excluding the active cell keeps the anchor untinted inside the range, as Excel
   and Sheets do, so you can always see where typing will land. */
.${p}cell.${p}selected:not(.${p}active)::after,
.${p}head.${p}headon::after { content: ""; position: absolute; inset: 0;
  background: ${t.selectedBg}; pointer-events: none; }
.${p}cell.${p}active { outline: 2px solid ${t.accent}; outline-offset: -2px; z-index: 1; }
.${p}cell.${p}locked { background-image: repeating-linear-gradient(45deg,
  rgba(0,0,0,0.03) 0 4px, transparent 4px 8px); }
/* The editor is the one place inside a cell where selecting text is the point,
   so it opts back out of the cell's user-select: none. It matches the cell's font
   rather than the formula bar's, so text does not reflow the moment editing starts.
   monoFontFamily is for formula editing, where alignment carries meaning; a sheet
   of names and totals set in a monospace face just looks like a terminal. */
.${p}cell input { border: none; outline: none; width: 100%; height: 100%;
  font-size: ${t.fontSize}; font-family: ${t.fontFamily}; background: ${t.cellBg};
  color: inherit; padding: 0; user-select: text; }
.${p}head { background: ${t.headerBg}; border-right: 1px solid ${t.border};
  border-bottom: 1px solid ${t.headerBorder}; font-size: 12px; font-weight: 600;
  color: ${t.headerText}; display: flex; align-items: center; justify-content: center;
  user-select: none; position: relative; box-sizing: border-box; cursor: default; }
/* Headers spanned by the selection, so the range is readable from the edges.
   The tint itself is the shared ::after rule above; only the text changes here. */
.${p}head.${p}headon { color: ${t.accent}; }
/* Resize grabbers straddle the divider they move: half inside the header, half
   over the neighbour, so the target is the line the user is aiming at rather
   than the few pixels to one side of it. */
.${p}resize { position: absolute; right: -3px; top: 0; width: 7px; height: 100%;
  cursor: col-resize; z-index: 6; }
.${p}rowresize { position: absolute; left: 0; bottom: -3px; width: 100%; height: 7px;
  cursor: row-resize; z-index: 6; }
/* inline-flex so an icon and its label sit on one baseline-free centre line.
   Buttons hold both — icon-only in the toolbar, icon plus text in the file
   menu — and a plain inline-block would drop the SVG onto the text baseline. */
.${p}btn { border: 1px solid ${t.buttonBorder}; background: ${t.toolbarBg};
  padding: 5px 9px; border-radius: 6px; font-size: ${t.fontSize}; cursor: pointer;
  color: ${t.cellText}; line-height: 1.2; display: inline-flex; align-items: center;
  justify-content: center; gap: 6px; }
.${p}btn:hover:not(:disabled) { background: ${t.headerBg}; }
.${p}btn:disabled { opacity: 0.4; cursor: default; }
.${p}btn.${p}on { background: ${t.accent}; color: #fff; border-color: ${t.accent}; }
/* Square, so a row of icon buttons reads as a row rather than as a ragged
   set of differently-proportioned boxes. */
.${p}iconbtn { padding: 6px; }
/* The dropdown affordance on a cell with a data-validation list. Sits at the
   right edge, inside the cell, and never over the fill handle's corner. */
.${p}dropdown { position: absolute; right: 1px; top: 50%;
  transform: translateY(-50%); display: flex; align-items: center;
  justify-content: center; width: 16px; height: 16px; padding: 0;
  border: 1px solid ${t.buttonBorder}; border-radius: 3px;
  background: ${t.headerBg}; color: ${t.headerText}; cursor: pointer; z-index: 4; }
.${p}dropdown:hover { background: ${t.toolbarBg}; }
.${p}fillhandle { position: absolute; right: -4px; bottom: -4px; width: 7px; height: 7px;
  background: ${t.accent}; border: 1px solid #fff; cursor: crosshair; z-index: 5; }
.${p}menu { position: fixed; background: ${t.toolbarBg}; border: 1px solid ${t.buttonBorder};
  border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.12); padding: 4px;
  z-index: 1000; min-width: 180px; }
.${p}menu button { display: block; width: 100%; text-align: left; border: none;
  background: none; padding: 6px 10px; font-size: ${t.fontSize}; border-radius: 4px;
  cursor: pointer; color: ${t.cellText}; }
.${p}menu button:hover { background: ${t.headerBg}; }
.${p}menu hr { border: none; border-top: 1px solid ${t.border}; margin: 4px 0; }
.${p}tab { padding: 6px 12px; font-size: ${t.fontSize}; border-radius: 6px 6px 0 0;
  cursor: pointer; border: 1px solid transparent; background: none; color: ${t.cellText}; }
.${p}tab.${p}on { background: ${t.cellBg}; border-color: ${t.border};
  border-bottom-color: ${t.cellBg}; font-weight: 600; }
.${p}sep { display: inline-block; width: 1px; height: 20px; background: ${t.border}; }
.${p}input { font-size: ${t.fontSize}; padding: 4px 8px; border: 1px solid ${t.buttonBorder};
  border-radius: 6px; background: ${t.cellBg}; color: ${t.cellText}; }

/* The native scrollbars are hidden and replaced. See Scrollbar.tsx for why:
   in short, they are overlay bars on macOS and mobile, so they sit on top of
   the rightmost column and the bottom row exactly while you scroll through
   them, take no layout space for the grid to work around, and cannot be made
   always-visible or given a consistent look across engines.

   Hidden, not disabled — the container still scrolls, so the wheel, the
   trackpad, the keyboard, and scrollIntoView all behave normally. */
.${p}scroller { scrollbar-width: none; -ms-overflow-style: none; }
.${p}scroller::-webkit-scrollbar { display: none; }

/* Both bars are the same control on different axes, so everything except the
   direction is shared. The track is always present, whether or not the sheet
   overflows: a scrollbar that comes and goes shifts the layout under the
   pointer, and an empty channel reads as "nothing to scroll" perfectly well. */
.${p}sbar { position: relative; background: ${t.scrollbarTrack}; }
.${p}sbar-vertical { grid-column: 2; grid-row: 1; border-left: 1px solid ${t.border}; }
.${p}sbar-horizontal { grid-column: 1; grid-row: 2; border-top: 1px solid ${t.border}; }
.${p}sbcorner { grid-column: 2; grid-row: 2; background: ${t.scrollbarTrack};
  border-left: 1px solid ${t.border}; border-top: 1px solid ${t.border}; }
/* Inset from the channel by a transparent border rather than by being
   narrower, so it stays centred and can be rounded. Positioned with a
   transform, which the compositor can move without a layout pass — this runs
   on every scroll event. */
.${p}sbthumb { position: absolute; top: 0; left: 0;
  background: ${t.scrollbarThumb}; background-clip: content-box;
  border: ${SCROLLBAR_THUMB_INSET}px solid transparent;
  border-radius: ${SCROLLBAR_SIZE}px; box-sizing: border-box; }
.${p}sbar-vertical .${p}sbthumb { width: 100%; }
.${p}sbar-horizontal .${p}sbthumb { height: 100%; }
.${p}sbthumb:hover, .${p}sbthumb.${p}on { background: ${t.scrollbarThumbHover}; }
`.trim();
}
