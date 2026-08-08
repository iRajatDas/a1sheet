/**
 * Context / column menus: dismiss on scroll, clamp into the viewport.
 */
import { describe, expect, test } from "bun:test";
import { act, fireEvent, render, screen } from "@testing-library/react";
import { Sheet } from "./index.js";
import { Spreadsheet } from "./Spreadsheet.js";

describe("menus dismiss on scroll", () => {
  test("scrolling the grid closes the context menu", () => {
    render(
      <Sheet.Root height={240}>
        <Sheet.Grid style={{ height: 200 }} />
        <Sheet.ContextMenu />
      </Sheet.Root>,
    );

    const cell = document.querySelector('[data-row="0"][data-col="0"]');
    if (!cell) throw new Error("no cell");
    fireEvent.contextMenu(cell);
    expect(screen.getByRole("menu")).toBeDefined();

    const scroller = document.querySelector(".a1s-scroller");
    if (!scroller) throw new Error("no scroller");
    act(() => {
      scroller.scrollTop = 80;
      fireEvent.scroll(scroller);
    });

    expect(screen.queryByRole("menu")).toBeNull();
  });
});

describe("menus clamp into the viewport", () => {
  test("a menu opened near the bottom flips above the cursor", () => {
    const original = HTMLElement.prototype.getBoundingClientRect;
    HTMLElement.prototype.getBoundingClientRect = function mockRect() {
      if (this.getAttribute("role") === "menu") {
        return {
          width: 180,
          height: 220,
          top: 0,
          left: 0,
          right: 180,
          bottom: 220,
          x: 0,
          y: 0,
          toJSON() {},
        } as DOMRect;
      }
      return original.call(this);
    };

    try {
      render(<Spreadsheet />);
      const cell = document.querySelector('[data-row="0"][data-col="0"]');
      if (!cell) throw new Error("no cell");

      const nearBottom = window.innerHeight - 12;
      fireEvent.contextMenu(cell, { clientX: 40, clientY: nearBottom });

      const menu = screen.getByRole("menu");
      const top = Number.parseFloat(menu.style.top);
      expect(top).toBe(nearBottom - 220);
    } finally {
      HTMLElement.prototype.getBoundingClientRect = original;
    }
  });
});
