/**
 * Scrolling the keyboard's cursor back into view.
 *
 * The failure modes are opposites and both are bad: not scrolling at all leaves
 * the selection somewhere off screen, and scrolling too eagerly shoves the sheet
 * around under a cell that was perfectly visible.
 */
import { describe, expect, test } from "bun:test";
import { compensateScrollForTrackResize, revealOffset } from "./reveal.js";

/** A 500px viewport with a 26px header floating over the top of it. */
const view = { viewport: 500, size: 26, lead: 26 };

describe("a visible cell is left alone", () => {
  test("in the middle of the viewport", () => {
    expect(revealOffset({ ...view, offset: 1000, start: 1200 })).toBe(1000);
  });

  test("flush against the last visible pixel", () => {
    // start + size + lead === offset + viewport: the last position that still
    // clears the header. One pixel of slack here and every keystroke scrolls.
    expect(revealOffset({ ...view, offset: 1000, start: 1448 })).toBe(1000);
  });
});

describe("a cell out of view is brought back by the minimum", () => {
  test("above the top: the row's own start becomes the offset", () => {
    expect(revealOffset({ ...view, offset: 1000, start: 800 })).toBe(800);
  });

  test("below the bottom: only far enough for the row to clear", () => {
    // 1600 + 26 + 26 - 500. Not centred — the cell arrives at the edge it was
    // approaching, which is where a spreadsheet user expects to find it.
    expect(revealOffset({ ...view, offset: 1000, start: 1600 })).toBe(1152);
  });

  test("a cell taller than the viewport pins its top edge", () => {
    const tall = revealOffset({
      offset: 0,
      viewport: 100,
      start: 40,
      size: 400,
      lead: 26,
    });
    // It cannot all fit; what matters is that the scroll lands somewhere the
    // cell occupies rather than past it.
    expect(tall).toBeGreaterThan(40);
    expect(tall).toBeLessThan(440);
  });
});

describe("compensateScrollForTrackResize", () => {
  test("shifts scroll when the track is fully before the viewport", () => {
    expect(
      compensateScrollForTrackResize({
        offset: 500,
        trackStart: 100,
        prevSize: 80,
        nextSize: 180,
      }),
    ).toBe(600);
  });

  test("leaves scroll alone when the track intersects the viewport", () => {
    expect(
      compensateScrollForTrackResize({
        offset: 500,
        trackStart: 480,
        prevSize: 80,
        nextSize: 180,
      }),
    ).toBe(500);
  });

  test("leaves scroll alone when the track is after the viewport origin", () => {
    expect(
      compensateScrollForTrackResize({
        offset: 500,
        trackStart: 600,
        prevSize: 80,
        nextSize: 40,
      }),
    ).toBe(500);
  });

  test("clamps at zero when shrinking a leading track", () => {
    expect(
      compensateScrollForTrackResize({
        offset: 40,
        trackStart: 0,
        prevSize: 40,
        nextSize: 0,
      }),
    ).toBe(0);
  });

  test("compares in content-space when a sticky lead is present", () => {
    // Scroller offset 40 with a 44px row header → content offset 0. Column 0
    // (start 0, width 100) intersects the content origin — do not compensate.
    expect(
      compensateScrollForTrackResize({
        offset: 40,
        trackStart: 0,
        prevSize: 100,
        nextSize: 180,
        lead: 44,
      }),
    ).toBe(40);
  });

  test("still compensates a leading track past the sticky header", () => {
    // scrollLeft 200, lead 44 → content 156. Track ending at 100 is fully left.
    expect(
      compensateScrollForTrackResize({
        offset: 200,
        trackStart: 0,
        prevSize: 100,
        nextSize: 180,
        lead: 44,
      }),
    ).toBe(280);
  });
});
