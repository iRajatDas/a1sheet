/**
 * The single stylesheet, emitted as a string into one `<style>` tag by
 * `<Spreadsheet />`. Ported from the `<style>` block in ref/Spreadsheet.jsx.
 *
 * Not a `.css` file on purpose: a CSS import would force consumers to have a CSS
 * loader configured, breaking the drop-in requirement. Class names are prefixed
 * so they cannot collide with the host app.
 */
import type { Theme } from "./theme.js";

export function buildCss(prefix: string, t: Theme): string {
  const p = prefix;
  return `
.${p}cell { border-right: 1px solid ${t.border}; border-bottom: 1px solid ${t.border};
  overflow: hidden; white-space: nowrap; display: flex; align-items: center;
  padding: 0 6px; font-size: ${t.fontSize}; font-family: ${t.monoFontFamily};
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
   so it opts back out of the cell's user-select: none. */
.${p}cell input { border: none; outline: none; width: 100%; height: 100%;
  font-size: ${t.fontSize}; font-family: ${t.monoFontFamily}; background: ${t.cellBg};
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
.${p}btn { border: 1px solid ${t.buttonBorder}; background: ${t.toolbarBg};
  padding: 5px 9px; border-radius: 6px; font-size: ${t.fontSize}; cursor: pointer;
  color: ${t.cellText}; line-height: 1.2; }
.${p}btn:hover:not(:disabled) { background: ${t.headerBg}; }
.${p}btn:disabled { opacity: 0.4; cursor: default; }
.${p}btn.${p}on { background: ${t.accent}; color: #fff; border-color: ${t.accent}; }
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
`.trim();
}
