/**
 * Context / column menus: dismiss on scroll, clamp into sheet bounds, portal.
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

describe("menus clamp into sheet bounds", () => {
  test("menus render in a document.body portal", () => {
    render(
      <Sheet.Root height={240}>
        <Sheet.Grid style={{ height: 200 }} />
        <Sheet.ContextMenu />
      </Sheet.Root>,
    );
    const cell = document.querySelector('[data-row="0"][data-col="0"]');
    if (!cell) throw new Error("no cell");
    fireEvent.contextMenu(cell);
    const menu = screen.getByRole("menu");
    expect(menu.parentElement).toBe(document.body);
  });

  test("a menu opened near the bottom of a short pane flips above the cursor", () => {
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
      if (this.classList?.contains("a1s-root")) {
        return {
          width: 400,
          height: 280,
          top: 0,
          left: 0,
          right: 400,
          bottom: 280,
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

      // Near the bottom of the 280px sheet pane — window height is irrelevant.
      fireEvent.contextMenu(cell, { clientX: 40, clientY: 250 });

      const menu = screen.getByRole("menu");
      const top = Number.parseFloat(menu.style.top);
      expect(top).toBe(250 - 220);
    } finally {
      HTMLElement.prototype.getBoundingClientRect = original;
    }
  });
});
