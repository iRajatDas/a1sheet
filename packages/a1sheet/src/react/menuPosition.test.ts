import { describe, expect, test } from "bun:test";
import { clampMenuPosition } from "./menuPosition.js";

describe("clampMenuPosition", () => {
  test("keeps the click point when the menu fits", () => {
    expect(
      clampMenuPosition({
        x: 100,
        y: 80,
        width: 200,
        height: 160,
        viewportWidth: 800,
        viewportHeight: 600,
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
        viewportWidth: 800,
        viewportHeight: 600,
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
        viewportWidth: 800,
        viewportHeight: 600,
      }),
    ).toEqual({ left: 550, top: 80 });
  });

  test("clamps into the viewport when the menu is taller than the window", () => {
    expect(
      clampMenuPosition({
        x: 40,
        y: 40,
        width: 200,
        height: 700,
        viewportWidth: 800,
        viewportHeight: 600,
        margin: 8,
      }),
    ).toEqual({ left: 40, top: 8 });
  });
});
