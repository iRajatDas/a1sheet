import { describe, expect, test } from "bun:test";
import { clampMenuPosition, collisionBoundsFromElement } from "./menuPosition.js";

const windowBounds = {
  left: 0,
  top: 0,
  right: 800,
  bottom: 600,
};

describe("clampMenuPosition", () => {
  test("keeps the click point when the menu fits", () => {
    expect(
      clampMenuPosition({
        x: 100,
        y: 80,
        width: 200,
        height: 160,
        bounds: windowBounds,
      }),
    ).toEqual({ left: 100, top: 80 });
  });

  test("flips above when there is not enough space below", () => {
    expect(
      clampMenuPosition({
        x: 100,
        y: 550,
        width: 200,
        height: 160,
        bounds: windowBounds,
      }),
    ).toEqual({ left: 100, top: 390 });
  });

  test("slides left when there is not enough space on the right", () => {
    expect(
      clampMenuPosition({
        x: 750,
        y: 80,
        width: 200,
        height: 160,
        bounds: windowBounds,
      }),
    ).toEqual({ left: 550, top: 80 });
  });

  test("clamps into the bounds when the menu is taller than the box", () => {
    expect(
      clampMenuPosition({
        x: 40,
        y: 40,
        width: 200,
        height: 700,
        bounds: windowBounds,
        margin: 8,
      }),
    ).toEqual({ left: 40, top: 8 });
  });

  test("flips against a short sheet pane even when the window is tall", () => {
    // Cursor near the bottom of a 280px-tall host pane sitting in a tall window.
    expect(
      clampMenuPosition({
        x: 40,
        y: 250,
        width: 180,
        height: 200,
        bounds: { left: 0, top: 0, right: 400, bottom: 280 },
        margin: 8,
      }),
    ).toEqual({ left: 40, top: 50 });
  });

  test("clamps to the top-left of tiny bounds when the menu is larger", () => {
    expect(
      clampMenuPosition({
        x: 20,
        y: 20,
        width: 200,
        height: 200,
        bounds: { left: 10, top: 10, right: 110, bottom: 90 },
        margin: 4,
      }),
    ).toEqual({ left: 14, top: 14 });
  });

  test("clamps to the top edge when flipping still overflows", () => {
    expect(
      clampMenuPosition({
        x: 40,
        y: 40,
        width: 100,
        height: 100,
        bounds: { left: 0, top: 100, right: 400, bottom: 180 },
        margin: 8,
      }),
    ).toEqual({ left: 40, top: 108 });
  });
});

describe("collisionBoundsFromElement", () => {
  test("falls back to the window when the element is missing", () => {
    const bounds = collisionBoundsFromElement(null);
    expect(bounds.left).toBeLessThanOrEqual(0);
    expect(bounds.top).toBeLessThanOrEqual(0);
    // happy-dom may report 0×0; real browsers use innerWidth/innerHeight.
    expect(bounds.right).toBeGreaterThanOrEqual(bounds.left);
    expect(bounds.bottom).toBeGreaterThanOrEqual(bounds.top);
  });

  test("intersects the element rect with the window", () => {
    const el = {
      getBoundingClientRect: () =>
        ({
          left: 50,
          top: 100,
          right: 350,
          bottom: 300,
          width: 300,
          height: 200,
          x: 50,
          y: 100,
          toJSON() {},
        }) as DOMRect,
    } as Element;
    const bounds = collisionBoundsFromElement(el);
    expect(bounds.left).toBe(50);
    expect(bounds.top).toBe(100);
    expect(bounds.right).toBe(350);
    expect(bounds.bottom).toBe(300);
  });
});
