/**
 * Composing primitives — the supported extension point.
 *
 * These stories are the argument for the API: every arrangement below is a
 * different tree, and none of them needed a new prop on anything.
 */
import type { Meta, StoryObj } from "@storybook/react-vite";
import { Sheet, useSheet } from "a1sheet/react";
import { useState } from "react";
import * as fixtures from "./fixtures.js";

const meta = {
  title: "Composition/Primitives",
  tags: ["autodocs"],
  parameters: {
    docs: {
      description: {
        component:
          "Sheet.Root provides context; every other primitive reads from it and throws a named error outside it. Reorder them, drop them, wrap them, or mix in your own components.",
      },
    },
  },
} satisfies Meta;

export default meta;
type Story = StoryObj<typeof meta>;

export const GridOnly: Story = {
  name: "Grid only",
  render: () => (
    <Sheet.Root defaultWorkbook={fixtures.budget()} height={360}>
      <Sheet.Grid />
    </Sheet.Root>
  ),
  parameters: {
    docs: {
      description: {
        story:
          "No toolbar, no tabs, no status bar — not because a flag turned them off, but because they are not in the tree. Editing, selection, and the clipboard all still work.",
      },
    },
  },
};

export const ReorderedParts: Story = {
  name: "Parts in a different order",
  render: () => (
    <Sheet.Root defaultWorkbook={fixtures.budget()} height={420}>
      <Sheet.StatusBar />
      <Sheet.Grid />
      <Sheet.FormulaBar />
      <Sheet.Toolbar />
      <Sheet.ContextMenu />
      <Sheet.ColumnMenu />
    </Sheet.Root>
  ),
  parameters: {
    docs: {
      description: {
        story:
          "The status bar on top and the toolbar at the bottom. The library never owns your layout, so this needs no support from it.",
      },
    },
  },
};

/** A consumer component reading the same context the primitives read. */
function SelectionSummary() {
  const api = useSheet();
  const { r1, c1, r2, c2 } = api.bounds;
  const count = (r2 - r1 + 1) * (c2 - c1 + 1);
  return (
    <footer
      style={{
        padding: "8px 12px",
        font: "500 13px/1.4 system-ui",
        background: "#f1f5f9",
        borderTop: "1px solid #cbd5e1",
      }}
    >
      {count} cell{count === 1 ? "" : "s"} selected · active value{" "}
      <code>{api.getDisplay(api.active.row, api.active.col) || "—"}</code>
    </footer>
  );
}

export const YourOwnComponent: Story = {
  name: "Your component, our data",
  render: () => (
    <Sheet.Root defaultWorkbook={fixtures.budget()} height={420}>
      <Sheet.Grid />
      <SelectionSummary />
    </Sheet.Root>
  ),
  parameters: {
    docs: {
      description: {
        story:
          "useSheet() returns the same headless API the primitives use, so a component you write is not a second-class citizen.",
      },
    },
  },
};

export const ToolbarWithYourButtons: Story = {
  name: "Extra buttons in the toolbar",
  render: () => (
    <Sheet.Root defaultWorkbook={fixtures.budget()} height={420}>
      <Sheet.Toolbar>
        <Sheet.FileMenu />
        <SaveButton />
      </Sheet.Toolbar>
      <Sheet.Grid />
      <Sheet.StatusBar />
    </Sheet.Root>
  ),
  parameters: {
    docs: {
      description: {
        story:
          "Toolbar renders its children after a separator. File I/O is itself just a child — omit Sheet.FileMenu and the XLSX writer never enters your bundle.",
      },
    },
  },
};

function SaveButton() {
  const api = useSheet();
  const [saved, setSaved] = useState(0);
  return (
    <button
      type="button"
      className="a1s-btn"
      onClick={() => {
        // Whatever your app does with a workbook.
        void api.workbook;
        setSaved((n) => n + 1);
      }}
    >
      {saved === 0 ? "Save" : `Saved ×${saved}`}
    </button>
  );
}

export const TwoSheetsOnePage: Story = {
  name: "Two independent sheets",
  render: () => (
    <div style={{ display: "grid", gap: 16, gridTemplateColumns: "1fr 1fr" }}>
      <Sheet.Root defaultWorkbook={fixtures.budget()} height={320}>
        <Sheet.FormulaBar />
        <Sheet.Grid />
      </Sheet.Root>
      <Sheet.Root defaultWorkbook={fixtures.formulas()} height={320}>
        <Sheet.FormulaBar />
        <Sheet.Grid />
      </Sheet.Root>
    </div>
  ),
  parameters: {
    docs: {
      description: {
        story:
          "No singletons and no global state, so two roots on one page do not interfere.",
      },
    },
  },
};
