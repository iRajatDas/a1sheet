/**
 * The edge-scroll ramp.
 *
 * A drag held against the bottom of the grid has to scroll, and how fast is the
 * whole feel of it: too slow and a thousand-row selection takes a minute, too
 * fast and you overshoot every time. The shape is a pure function so it can be
 * checked without a pointer, a clock, or a layout.
 */
import { describe, expect, test } from "bun:test";
import { edgeVelocity } from "./useAutoScroll.js";

// A 600px-tall element from y=100 to y=700.
const TOP = 100;
const BOTTOM = 700;

describe("nothing moves while the pointer is inside", () => {
  test("the middle is still", () => {
    expect(edgeVelocity(400, TOP, BOTTOM)).toBe(0);
  });

  test("so is the area just short of the edge zone", () => {
    expect(edgeVelocity(BOTTOM - 25, TOP, BOTTOM)).toBe(0);
  });
});

describe("past the edge it moves, and away from the pointer", () => {
  test("near the bottom scrolls forward, near the top back", () => {
    expect(edgeVelocity(BOTTOM - 10, TOP, BOTTOM)).toBeGreaterThan(0);
    expect(edgeVelocity(TOP + 10, TOP, BOTTOM)).toBeLessThan(0);
  });

  test("further out is faster, up to a ceiling", () => {
    const near = edgeVelocity(BOTTOM + 10, TOP, BOTTOM);
    const far = edgeVelocity(BOTTOM + 100, TOP, BOTTOM);
    const absurd = edgeVelocity(BOTTOM + 5000, TOP, BOTTOM);

    expect(far).toBeGreaterThan(near);
    // A pointer flung to the far side of the screen must not scroll a thousand
    // rows a frame — past the ramp the speed stops growing.
    expect(absurd).toBe(edgeVelocity(BOTTOM + 200, TOP, BOTTOM));
  });

  test("the two directions are mirror images", () => {
    expect(edgeVelocity(TOP - 50, TOP, BOTTOM)).toBe(
      -edgeVelocity(BOTTOM + 50, TOP, BOTTOM),
    );
  });

  test("entering the zone starts gently rather than jumping", () => {
    // The step from stationary to moving is what makes an auto-scroll feel
    // uncontrollable; the first frame inside the zone must be a crawl.
    expect(Math.abs(edgeVelocity(BOTTOM - 23, TOP, BOTTOM))).toBeLessThan(4);
  });
});
