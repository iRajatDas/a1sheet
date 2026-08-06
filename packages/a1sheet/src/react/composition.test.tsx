/**
 * The composition contract: context, asChild, imperative handle, controlled mode,
 * and failing fast outside the provider.
 *
 * These exist because the previous API used `show*` booleans, and nothing stopped
 * that from coming back except a test that asserts the contract.
 */
import { describe, expect, test } from "bun:test";
import { act, fireEvent, render, screen } from "@testing-library/react";
import { createRef } from "react";
import { MissingProviderError } from "../errors.js";
import type { Workbook } from "../model/types.js";
import { createWorkbook } from "../model/workbook.js";
import { ADD_ROWS_DEFAULT, HEADER_HEIGHT, ROW_HEIGHT } from "./constants.js";
import { Root, Sheet, type SheetRootHandle, Slot, useSheet } from "./index.js";

function workbookWith(cells: Record<string, string>) {
  const wb = createWorkbook(["Sheet1"]);
  Object.assign((wb.sheets[0] as { cells: object }).cells, cells);
  return wb;
}

describe("primitives compose freely", () => {
  test("grid alone renders", () => {
    const { container } = render(
      <Sheet.Root>
        <Sheet.Grid />
      </Sheet.Root>,
    );
    expect(container.querySelectorAll(".a1s-cell").length).toBeGreaterThan(0);
  });

  test("order is the consumer's choice, not the library's", () => {
    const { container } = render(
      <Sheet.Root>
        <Sheet.StatusBar />
        <Sheet.Grid />
        <Sheet.Toolbar />
      </Sheet.Root>,
    );
    const kids = [...(container.querySelector(".a1s-root")?.children ?? [])];
    const statusIdx = kids.findIndex((el) => el.classList.contains("a1s-status"));
    const toolbarIdx = kids.findIndex((el) => el.querySelector(".a1s-btn"));
    // Status bar precedes the toolbar because that is how it was written.
    expect(statusIdx).toBeLessThan(toolbarIdx);
  });

  test("a consumer's own component sits alongside the primitives", () => {
    function MyFooter() {
      const api = useSheet();
      return <div data-testid="mine">rows: {api.sheet.numRows}</div>;
    }
    render(
      <Sheet.Root>
        <Sheet.Grid />
        <MyFooter />
      </Sheet.Root>,
    );
    expect(screen.getByTestId("mine").textContent).toBe("rows: 200");
  });

  test("no primitive accepts a show* prop", () => {
    // A compile-time guarantee, asserted at runtime as documentation: passing an
    // unknown prop must not change what renders.
    const { container } = render(
      <Sheet.Root>
        {/* @ts-expect-error showToolbar is not, and must never be, a prop */}
        <Sheet.Grid showToolbar={false} />
      </Sheet.Root>,
    );
    expect(container.querySelectorAll(".a1s-cell").length).toBeGreaterThan(0);
  });
});

describe("provider boundary", () => {
  test("a primitive outside Root throws a named, actionable error", () => {
    // React logs the error too; the assertion is on the throw.
    expect(() => render(<Sheet.Grid />)).toThrow(MissingProviderError);
    expect(() => render(<Sheet.Grid />)).toThrow(
      /must be rendered inside <Sheet.Root>/,
    );
  });

  test("useSheet outside Root throws", () => {
    function Orphan() {
      useSheet();
      return null;
    }
    expect(() => render(<Orphan />)).toThrow(MissingProviderError);
  });

  test("the error carries a stable code", () => {
    try {
      render(<Sheet.StatusBar />);
      throw new Error("should have thrown");
    } catch (e) {
      expect((e as MissingProviderError).code).toBe("MISSING_PROVIDER");
    }
  });
});

describe("imperative handle", () => {
  test("focus() returns keyboard focus to the grid", () => {
    const ref = createRef<SheetRootHandle>();
    const { container } = render(
      <Sheet.Root ref={ref}>
        <Sheet.Grid />
      </Sheet.Root>,
    );
    const outside = document.createElement("input");
    document.body.appendChild(outside);
    outside.focus();
    expect(document.activeElement).toBe(outside);

    ref.current?.focus();
    expect(document.activeElement).toBe(container.querySelector("textarea"));
    outside.remove();
  });

  test("api exposes live state, not a snapshot", () => {
    const ref = createRef<SheetRootHandle>();
    render(
      <Sheet.Root defaultWorkbook={workbookWith({ "0_0": "before" })} ref={ref}>
        <Sheet.Grid />
      </Sheet.Root>,
    );
    expect(ref.current?.api.getRaw(0, 0)).toBe("before");
    // act() so the state update flushes and the handle is rebuilt — `api` is the
    // live object from the latest render, not a snapshot taken at mount.
    act(() => ref.current?.api.setCell(0, 0, "after"));
    expect(ref.current?.api.getRaw(0, 0)).toBe("after");
  });
});

describe("controlled and uncontrolled", () => {
  test("uncontrolled holds its own state", () => {
    const { container } = render(
      <Sheet.Root defaultWorkbook={workbookWith({ "0_0": "x" })}>
        <Sheet.Grid />
      </Sheet.Root>,
    );
    const text = [...container.querySelectorAll(".a1s-cell")].map(
      (el) => el.textContent,
    );
    expect(text).toContain("x");
  });

  test("controlled reports changes and does not self-update", () => {
    const wb = workbookWith({ "0_0": "held" });
    let reported: Workbook | null = null;

    const ref = createRef<SheetRootHandle>();
    const { container } = render(
      <Sheet.Root
        workbook={wb}
        onWorkbookChange={(next) => {
          reported = next;
        }}
        ref={ref}
      >
        <Sheet.Grid />
      </Sheet.Root>,
    );

    act(() => ref.current?.api.setCell(0, 0, "changed"));

    // The change was reported upward...
    expect(reported).not.toBeNull();
    expect((reported as unknown as Workbook).sheets[0]?.cells["0_0"]).toBe(
      "changed",
    );

    // ...and NOT applied locally, because the parent owns the value.
    const text = [...container.querySelectorAll(".a1s-cell")].map(
      (el) => el.textContent,
    );
    expect(text).toContain("held");
    expect(text).not.toContain("changed");
  });
});

describe("asChild", () => {
  test("Slot merges props onto the child instead of adding an element", () => {
    let clicked = 0;
    const { container } = render(
      <Slot className="ours" onClick={() => clicked++}>
        <button type="button" className="theirs">
          Go
        </button>
      </Slot>,
    );
    const btn = container.querySelector("button") as HTMLButtonElement;
    // One element, not a wrapper plus a button.
    expect(container.children).toHaveLength(1);
    expect(btn.className).toBe("ours theirs");
    fireEvent.click(btn);
    expect(clicked).toBe(1);
  });

  test("both handlers run, ours first", () => {
    const order: string[] = [];
    const { container } = render(
      <Slot onClick={() => order.push("ours")}>
        <button type="button" onClick={() => order.push("theirs")}>
          Go
        </button>
      </Slot>,
    );
    fireEvent.click(container.querySelector("button") as HTMLButtonElement);
    expect(order).toEqual(["ours", "theirs"]);
  });

  test("style merges with the child winning", () => {
    const { container } = render(
      <Slot style={{ color: "red", margin: "1px" }}>
        <div style={{ color: "blue" }} />
      </Slot>,
    );
    const el = container.firstElementChild as HTMLElement;
    expect(el.style.color).toBe("blue");
    expect(el.style.margin).toBe("1px");
  });

  test("more than one child is a programmer error", () => {
    expect(() =>
      render(
        <Slot>
          <span />
          <span />
        </Slot>,
      ),
    ).toThrow(/exactly one React element/);
  });
});

describe("Root layout props", () => {
  test("className and style are merged, not replaced", () => {
    const { container } = render(
      <Root className="mine" style={{ opacity: 0.5 }}>
        <Sheet.Grid />
      </Root>,
    );
    const root = container.firstElementChild as HTMLElement;
    expect(root.className).toContain("a1s-root");
    expect(root.className).toContain("mine");
    expect(root.style.opacity).toBe("0.5");
  });
});

describe("file I/O is a primitive, not a preset privilege", () => {
  // The regression: import and export used to arrive as three optional
  // callbacks that only <Spreadsheet /> passed, so a hand-composed
  // <Sheet.Toolbar /> rendered no Import button and the feature looked deleted.

  test("Sheet.FileMenu brings its own handlers, with no props", () => {
    render(
      <Sheet.Root>
        <Sheet.Toolbar>
          <Sheet.FileMenu />
        </Sheet.Toolbar>
        <Sheet.Grid />
      </Sheet.Root>,
    );

    expect(screen.getByText("Import")).toBeDefined();
    expect(screen.getByText("Export CSV")).toBeDefined();
    expect(screen.getByText("Export XLSX")).toBeDefined();
    expect(screen.getByLabelText("Import a spreadsheet")).toBeDefined();
  });

  test("a toolbar without one has no file buttons and is otherwise intact", () => {
    // Not a `show*` flag — the part is simply absent from the tree, which is
    // what keeps the XLSX writer out of a bundle that does not want it.
    render(
      <Sheet.Root>
        <Sheet.Toolbar />
      </Sheet.Root>,
    );

    expect(screen.queryByText("Import")).toBeNull();
    expect(screen.queryByText("Export XLSX")).toBeNull();
    expect(screen.getByTitle("Freeze up through the selection")).toBeDefined();
  });

  test("arbitrary children go into the toolbar too", () => {
    render(
      <Sheet.Root>
        <Sheet.Toolbar>
          <button type="button">Save to my server</button>
        </Sheet.Toolbar>
      </Sheet.Root>,
    );

    expect(screen.getByText("Save to my server")).toBeDefined();
  });
});

describe("adding rows at the bottom", () => {
  function setupAddRows() {
    const wb = createWorkbook(["Sheet1"]);
    Object.assign(wb.sheets[0] as object, { numRows: 20 });
    const { container } = render(
      <Sheet.Root defaultWorkbook={wb}>
        <Sheet.Grid>
          <Sheet.AddRows />
        </Sheet.Grid>
      </Sheet.Root>,
    );
    const scroller = container.querySelector(".a1s-scroller") as HTMLElement;
    return { container, scroller, grid: scroller.firstElementChild as HTMLElement };
  }

  test("the control sits inside the scroll container, not in a bar below it", () => {
    // Google Sheets puts it at the end of the content, so scrolling to the
    // bottom of the sheet is what reveals it.
    const { scroller } = setupAddRows();

    expect(scroller.contains(screen.getByText("more rows at the bottom"))).toBe(
      true,
    );
  });

  test("Add grows the sheet, and the new rows are real scroll extent", () => {
    const { grid } = setupAddRows();

    fireEvent.click(screen.getByText("Add"));

    expect(Number.parseInt(grid.style.minHeight, 10)).toBe(
      HEADER_HEIGHT + (20 + ADD_ROWS_DEFAULT) * ROW_HEIGHT,
    );
  });

  test("the count is editable and Enter commits it", () => {
    const { grid } = setupAddRows();
    const input = screen.getByLabelText("Number of rows to add");

    fireEvent.change(input, { target: { value: "5" } });
    fireEvent.keyDown(input, { key: "Enter" });

    expect(Number.parseInt(grid.style.minHeight, 10)).toBe(
      HEADER_HEIGHT + 25 * ROW_HEIGHT,
    );
  });

  test("a blank count cannot be submitted", () => {
    const { grid } = setupAddRows();
    const before = grid.style.minHeight;

    fireEvent.change(screen.getByLabelText("Number of rows to add"), {
      target: { value: "" },
    });
    fireEvent.click(screen.getByText("Add"));

    expect(grid.style.minHeight).toBe(before);
  });

  test("adding rows is undoable", () => {
    const { container, grid } = setupAddRows();
    const before = grid.style.minHeight;

    fireEvent.click(screen.getByText("Add"));
    expect(grid.style.minHeight).not.toBe(before);

    const textarea = container.querySelector("textarea") as HTMLTextAreaElement;
    fireEvent.keyDown(textarea, { key: "z", ctrlKey: true });

    expect(grid.style.minHeight).toBe(before);
  });
});
