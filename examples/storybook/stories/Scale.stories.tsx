/**
 * Virtualization, sizing, and file I/O — the parts that are hard to believe
 * without touching them.
 */
import type { Meta, StoryObj } from "@storybook/react-vite";
import { Sheet } from "a1sheet/react";
import * as fixtures from "./fixtures.js";

const meta = {
  title: "Scale/Large sheets",
  tags: ["autodocs"],
  parameters: {
    docs: {
      description: {
        component: [
          "Both axes are virtualized: only cells near the viewport exist in the DOM, and the cost of a frame is flat in the size of the sheet.",
          "",
          "| sheet | cells in DOM | mount | per keystroke |",
          "|---|---|---|---|",
          "| 1k rows x 26 cols | 286 | 83 ms | 7.1 ms |",
          "| 100k rows x 26 cols | 286 | 30 ms | 6.6 ms |",
          "| 1k rows x 500 cols | 286 | 53 ms | 5.1 ms |",
          "| 100k rows x 500 cols | 286 | 125 ms | 11.9 ms |",
          "",
          "Measured under Bun with happy-dom, so treat them as relative rather than browser-accurate. What they show is the shape: adding rows or columns does not add work.",
          "",
          "**What does still scale with the sheet:** committing an edit copies the cell map, so an edit costs about 26 ms at 10k filled cells and about 390 ms at 1M. That is the write path, not the render path, and it is the next thing to fix.",
        ].join("\n"),
      },
    },
  },
} satisfies Meta;

export default meta;
type Story = StoryObj<typeof meta>;

export const HundredThousandRows: Story = {
  name: "100,000 rows",
  render: () => (
    <Sheet.Root defaultWorkbook={fixtures.large(100_000)} height={520}>
      <Sheet.Toolbar />
      <Sheet.FormulaBar />
      <Sheet.Grid>
        <Sheet.AddRows />
      </Sheet.Grid>
      <Sheet.StatusBar />
      <Sheet.ContextMenu />
      <Sheet.ColumnMenu />
    </Sheet.Root>
  ),
  parameters: {
    docs: {
      description: {
        story:
          "About 290 cells exist in the DOM at any moment, whatever the sheet's size. Drag the scrollbar: its length is the real 2,600,026px extent, not the height of the rendered window.",
      },
    },
  },
};

export const ManyColumns: Story = {
  name: "500 columns",
  render: () => (
    <Sheet.Root defaultWorkbook={fixtures.large(2_000, 500)} height={520}>
      <Sheet.Toolbar />
      <Sheet.Grid />
      <Sheet.StatusBar />
      <Sheet.ColumnMenu />
    </Sheet.Root>
  ),
  parameters: {
    docs: {
      description: {
        story:
          "Columns are virtualized too. Scroll sideways — the horizontal extent stays the width of the sheet rather than following whatever is rendered.",
      },
    },
  },
};

export const Filtering: Story = {
  name: "Filtering 5,000 rows",
  render: () => (
    <Sheet.Root defaultWorkbook={fixtures.filtered()} height={520}>
      <Sheet.Toolbar />
      <Sheet.Grid />
      <Sheet.StatusBar />
      <Sheet.ColumnMenu />
    </Sheet.Root>
  ),
  parameters: {
    docs: {
      description: {
        story:
          "Column C is filtered to Blocked and Done. Use the ▾ on a column header to change it. Editing a filtered column re-tests only the row you changed, not the sheet.",
      },
    },
  },
};

export const Sizing: Story = {
  name: "Resizing and auto-fit",
  render: () => (
    <Sheet.Root defaultWorkbook={fixtures.layout()} height={440}>
      <Sheet.Toolbar />
      <Sheet.Grid>
        <Sheet.AddRows />
      </Sheet.Grid>
      <Sheet.StatusBar />
      <Sheet.ColumnMenu />
    </Sheet.Root>
  ),
  parameters: {
    docs: {
      description: {
        story:
          "Drag the divider on a column header, or the one under a row header. Double-click a column divider to fit it to its widest value; double-click a row divider to reset the height. Row 1 and column A start out resized.",
      },
    },
  },
};

export const FileIO: Story = {
  name: "Import and export",
  render: () => (
    <Sheet.Root defaultWorkbook={fixtures.budget()} height={480}>
      <Sheet.Toolbar>
        <Sheet.FileMenu />
      </Sheet.Toolbar>
      <Sheet.FormulaBar />
      <Sheet.Grid>
        <Sheet.AddRows />
      </Sheet.Grid>
      <Sheet.Tabs />
      <Sheet.StatusBar />
      <Sheet.ContextMenu />
      <Sheet.ColumnMenu />
    </Sheet.Root>
  ),
  parameters: {
    docs: {
      description: {
        story:
          "Import a real .xlsx or .csv. Large reads yield to the browser and report progress in the status bar; choosing a second file cancels the first. Export writes a real workbook — no dependencies involved in either direction.",
      },
    },
  },
};
