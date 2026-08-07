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

/**
 * Layers, low to high. Everything drawn over the grid has to sit below the
 * sticky headers and the frozen bands, or it draws on top of them the moment the
 * cell it describes scrolls underneath — a selection outline printed across the
 * column headers, which is what happened before these were pinned down:
 *
 *   0  ordinary cells
 *   1  the active cell, selection outlines, the copy marquee, reference outlines
 *   2  the fill handle, frozen-band cells
 *   3  the row header      4  the column header      5  the corner
 *
 * Sticky elements take a z-index and therefore their own stacking context, so a
 * resize grabber at 6 inside a header is scoped to that header and does not
 * compete with anything here.
 */
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
/* The active cell inside a range: a 1px outline, not 2px. At the weight the grid
   lines are drawn, 2px reads as a heavy box drawn ON a cell rather than as the
   cursor sitting in one — and doubled up with the range border and the copy
   outline on the same cell it was the loudest thing on screen. */
.${p}cell.${p}active { outline: 1px solid ${t.accent}; outline-offset: -1px; z-index: 1; }
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
/* The label truncates rather than running under the menu button. min-width: 0 is
   what allows that at all — a flex item's default minimum is its content, so
   without it the label pushes the button out of the header instead of ellipsing. */
.${p}headlabel { overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
  min-width: 0; }
/* The sort/filter button sits IN the flow, not over the label: absolutely
   positioned it collided with the letters in any column narrow enough that the
   two wanted the same pixels. It is hidden rather than unmounted when idle, so
   its space stays reserved and the label does not shift on hover. */
.${p}headmenu { flex: none; display: flex; align-items: center; margin-left: 2px;
  border: none; background: none; padding: 0; cursor: pointer; color: inherit;
  visibility: hidden; }
.${p}head:hover .${p}headmenu,
.${p}head.${p}headon .${p}headmenu,
.${p}headmenu.${p}headfiltered { visibility: visible; }
.${p}headmenu.${p}headfiltered { color: ${t.accent}; }
/* The outline around each selected range, drawn as one box rather than as
   per-cell borders: a border on every cell would draw the internal grid lines in
   the accent colour too, which is not what a selection looks like anywhere.
   1px, and the same 1px on every range: the active one is already marked by the
   cell outline and the fill handle inside it, and a heavier line there stacked
   with those two on a single-cell selection came to three overlapping edges. */
.${p}selbox { border: 1px solid ${t.accent}; pointer-events: none; z-index: 1;
  box-sizing: border-box; }
/* The copied range. Dashed and thin — it marks something that already happened,
   so it must not compete with the selection, which marks what happens next. */
.${p}marquee { border: 1px dashed ${t.accent}; pointer-events: none; z-index: 1;
  box-sizing: border-box; }
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
  background: ${t.headerBg}; color: ${t.headerText}; cursor: pointer; z-index: 2; }
.${p}dropdown:hover { background: ${t.toolbarBg}; }
/* Round, and a little larger than the square it replaces: at 7px square it read
   as a rendering artefact of the corner rather than as something to grab. */
.${p}fillhandle { position: absolute; right: -4px; bottom: -4px; width: 8px; height: 8px;
  background: ${t.accent}; border: 1px solid ${t.cellBg}; border-radius: 50%;
  cursor: crosshair; z-index: 2; }
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
