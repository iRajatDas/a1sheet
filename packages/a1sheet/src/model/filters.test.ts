import { describe, expect, test } from "bun:test";
import {
  activateFilterView,
  colorMovedToTopMessage,
  createFilterView,
  deleteFilterView,
  FILTER_VIEW_MISSING,
  normalizeColumnFilter,
  rowMatchesColumnFilter,
} from "./filters.js";
import { type GridError, isGridError } from "./gridErrors.js";
import { makeSheet } from "./sheet.js";
import { sortByColor } from "./sortByColor.js";

describe("normalizeColumnFilter", () => {
  test("wraps a bare value set", () => {
    const set = new Set(["a"]);
    expect(normalizeColumnFilter(set)).toEqual({ values: set });
  });

  test("passes through a criteria object", () => {
    const filter = { background: new Set(["#ff0000"]) };
    expect(normalizeColumnFilter(filter)).toBe(filter);
  });
});

describe("rowMatchesColumnFilter", () => {
  test("values alone", () => {
    expect(
      rowMatchesColumnFilter({
        filter: { values: new Set(["keep"]) },
        display: "keep",
        background: undefined,
        foreground: undefined,
      }),
    ).toBe(true);
    expect(
      rowMatchesColumnFilter({
        filter: { values: new Set(["keep"]) },
        display: "drop",
        background: undefined,
        foreground: undefined,
      }),
    ).toBe(false);
  });

  test("background colour must match when set", () => {
    expect(
      rowMatchesColumnFilter({
        filter: { background: new Set(["#abc"]) },
        display: "x",
        background: "#abc",
        foreground: undefined,
      }),
    ).toBe(true);
    expect(
      rowMatchesColumnFilter({
        filter: { background: new Set(["#abc"]) },
        display: "x",
        background: "#def",
        foreground: undefined,
      }),
    ).toBe(false);
  });

  test("values AND colour", () => {
    expect(
      rowMatchesColumnFilter({
        filter: {
          values: new Set(["ok"]),
          foreground: new Set(["#111"]),
        },
        display: "ok",
        background: undefined,
        foreground: "#111",
      }),
    ).toBe(true);
    expect(
      rowMatchesColumnFilter({
        filter: {
          values: new Set(["ok"]),
          foreground: new Set(["#111"]),
        },
        display: "ok",
        background: undefined,
        foreground: "#222",
      }),
    ).toBe(false);
  });
});

describe("filter views", () => {
  test("create snapshots current filters and activates the view", () => {
    const sheet = makeSheet("S");
    sheet.filters[0] = { values: new Set(["a"]) };
    const next = createFilterView(sheet, { id: "v1", name: "Alpha" });
    expect(next.filterViews.v1?.name).toBe("Alpha");
    expect(next.filterViews.v1?.filters[0]?.values).toEqual(new Set(["a"]));
    expect(next.activeFilterViewId).toBe("v1");
    // Snapshot is a clone — mutating live filters later must not alter the view.
    next.filters[0] = { values: new Set(["b"]) };
    expect(next.filterViews.v1?.filters[0]?.values).toEqual(new Set(["a"]));
  });

  test("duplicate id throws FILTER_ID_EXISTS", () => {
    let sheet = makeSheet("S");
    sheet = createFilterView(sheet, { id: "v1", name: "A" });
    try {
      createFilterView(sheet, { id: "v1", name: "B" });
      throw new Error("should have thrown");
    } catch (e) {
      expect(isGridError(e)).toBe(true);
      expect((e as GridError).code).toBe("FILTER_ID_EXISTS");
    }
  });

  test("activate missing view returns null (caller reports status)", () => {
    const sheet = makeSheet("S");
    expect(activateFilterView(sheet, "nope")).toBeNull();
    expect(FILTER_VIEW_MISSING).toBe("The view does not exist.");
  });

  test("activate restores the view's filters", () => {
    let sheet = makeSheet("S");
    sheet.filters[0] = { values: new Set(["a"]) };
    sheet = createFilterView(sheet, { id: "v1", name: "A" });
    sheet = { ...sheet, filters: { 0: { values: new Set(["z"]) } } };
    const activated = activateFilterView(sheet, "v1");
    expect(activated?.filters[0]?.values).toEqual(new Set(["a"]));
    expect(activated?.activeFilterViewId).toBe("v1");
  });

  test("delete drops the view and clears active id when it was active", () => {
    let sheet = makeSheet("S");
    sheet = createFilterView(sheet, { id: "v1", name: "A" });
    sheet = deleteFilterView(sheet, "v1");
    expect(sheet.filterViews.v1).toBeUndefined();
    expect(sheet.activeFilterViewId).toBeNull();
  });
});

describe("sortByColor", () => {
  test("matching background colours rise to the top", () => {
    let sheet = makeSheet("S");
    sheet.cells["0_0"] = "a";
    sheet.cells["1_0"] = "b";
    sheet.cells["2_0"] = "c";
    sheet.styles["0_0"] = { bg: "#fff" };
    sheet.styles["1_0"] = { bg: "#f00" };
    sheet.styles["2_0"] = { bg: "#f00" };
    sheet = sortByColor(sheet, {
      col: 0,
      kind: "background",
      color: "#f00",
    });
    expect(sheet.cells["0_0"]).toBe("b");
    expect(sheet.cells["1_0"]).toBe("c");
    expect(sheet.cells["2_0"]).toBe("a");
    expect(colorMovedToTopMessage({ kind: "background", color: "#f00" })).toBe(
      "Cells with background color #f00 were moved to the top",
    );
  });
});
