/**
 * The grid's scrollbars.
 *
 * They are ours rather than the browser's because a native scrollbar is an
 * overlay bar on macOS and mobile: it draws on top of the rightmost column and
 * the bottom row exactly while you are scrolling through them, takes no layout
 * space for the grid to work around, and cannot be made always-visible or
 * given a consistent look across engines. See `components/Scrollbar.tsx`.
 *
 * What is asserted here is behaviour — a drag moves the container, a track
 * click pages — plus the two structural facts that would silently undo the
 * change: the native bars staying hidden, and the thumb keeping a floor.
 */
import { describe, expect, test } from "bun:test";
import { fireEvent, render } from "@testing-library/react";
import { Scrollbar } from "./components/Scrollbar.js";
import { SCROLLBAR_MIN_THUMB } from "./constants.js";
import { Spreadsheet } from "./Spreadsheet.js";

function setup() {
  const { container } = render(<Spreadsheet />);
  const scroller = container.querySelector(".a1s-scroller");
  if (!scroller) throw new Error("no scroll container");

  const bar = (orientation: "vertical" | "horizontal") => {
    const el = container.querySelector(`.a1s-sbar-${orientation}`);
    if (!el) throw new Error(`no ${orientation} scrollbar`);
    const thumb = el.querySelector(".a1s-sbthumb");
    if (!thumb) throw new Error(`no ${orientation} thumb`);
    return { el, thumb };
  };

  return { container, scroller: scroller as HTMLElement, bar };
}

describe("both axes get the same always-present bar", () => {
  test("a bar per axis, plus the corner between them", () => {
    const { container } = setup();

    const bars = container.querySelectorAll('[role="scrollbar"]');
    expect(bars.length).toBe(2);
    expect([...bars].map((b) => b.getAttribute("aria-orientation")).sort()).toEqual(
      ["horizontal", "vertical"],
    );
    expect(container.querySelector(".a1s-sbcorner")).not.toBeNull();
  });

  test("each bar says which element it scrolls", () => {
    const { container, scroller } = setup();

    for (const bar of container.querySelectorAll('[role="scrollbar"]')) {
      expect(bar.getAttribute("aria-controls")).toBe(scroller.id);
    }
    expect(scroller.id).not.toBe("");
  });

  test("the native scrollbars are hidden, not merely restyled", () => {
    const { container } = setup();
    const css = container.querySelector("style")?.textContent ?? "";

    expect(css).toContain(".a1s-scroller { scrollbar-width: none;");
    expect(css).toContain(".a1s-scroller::-webkit-scrollbar { display: none; }");
  });
});

describe("the thumb", () => {
  test("never shrinks below the grabbable floor", () => {
    // The default sheet is far taller than the viewport, so the honest
    // proportional thumb is small; the floor is what keeps it usable.
    const { bar } = setup();

    const height = (bar("vertical").thumb as HTMLElement).style.height;
    expect(Number.parseFloat(height)).toBeGreaterThanOrEqual(SCROLLBAR_MIN_THUMB);
  });

  test("never overhangs its channel when the sheet fits the window", () => {
    // The proportional size is viewport/content, which is above 1 whenever the
    // sheet is smaller than the window showing it. Driven directly, because
    // the sizes that produce it are a property of the geometry rather than of
    // any particular workbook.
    const { container } = render(
      <Scrollbar
        orientation="vertical"
        viewport={200}
        content={80}
        offset={0}
        controls="x"
        onScrollTo={() => {}}
        prefix="a1s-"
      />,
    );
    const thumb = container.querySelector(".a1s-sbthumb") as HTMLElement;

    expect(Number.parseFloat(thumb.style.height)).toBe(200);
  });

  test("dragging it scrolls the container", () => {
    const { scroller, bar } = setup();
    const { thumb } = bar("vertical");

    expect(scroller.scrollTop).toBe(0);
    fireEvent.mouseDown(thumb, { button: 0, clientY: 0 });
    fireEvent.mouseMove(window, { clientY: 40 });

    // Scaled up: the thumb crosses a short track while the sheet crosses its
    // whole height, so 40px of pointer travel is far more than 40px of sheet.
    expect(scroller.scrollTop).toBeGreaterThan(40);

    fireEvent.mouseUp(window);
    const settled = scroller.scrollTop;
    fireEvent.mouseMove(window, { clientY: 400 });
    expect(scroller.scrollTop).toBe(settled);
  });

  test("the horizontal bar drags on the other axis, and only that one", () => {
    const { scroller, bar } = setup();

    fireEvent.mouseDown(bar("horizontal").thumb, { button: 0, clientX: 0 });
    fireEvent.mouseMove(window, { clientX: 30, clientY: 30 });

    expect(scroller.scrollLeft).toBeGreaterThan(0);
    expect(scroller.scrollTop).toBe(0);
  });

  test("a non-primary button does not start a drag", () => {
    const { scroller, bar } = setup();

    fireEvent.mouseDown(bar("vertical").thumb, { button: 2, clientY: 0 });
    fireEvent.mouseMove(window, { clientY: 100 });

    expect(scroller.scrollTop).toBe(0);
  });
});

describe("clicking the track", () => {
  test("pages toward the click", () => {
    const { scroller, bar } = setup();

    // Well below the thumb, which sits at the top while scrollTop is 0.
    fireEvent.mouseDown(bar("vertical").el, { button: 0, clientY: 400 });

    expect(scroller.scrollTop).toBeGreaterThan(0);
  });
});
