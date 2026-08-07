/**
 * The flagship composition demo — realistic data, every primitive in the tree,
 * custom chrome on top, and a separate scale story that proves virtualization.
 */
import type { Meta, StoryObj } from "@storybook/react-vite";
import { toA1 } from "a1sheet";
import {
  darkTheme,
  type CellContentProps,
  Sheet,
  useSheet,
} from "a1sheet/react";
import { type ReactNode, useState } from "react";
import { budget, large, salesReport } from "./fixtures.js";

const meta = {
  title: "Composition/Showcase",
  tags: ["autodocs"],
  parameters: {
    layout: "fullscreen",
    docs: {
      description: {
        component:
          "A full application shell built only from composed primitives — no `show*` props, no preset shortcuts. " +
          "Edit cells, scroll a 25k-row sheet in the scale story, and inspect live state from your own components via `useSheet()`.",
      },
    },
  },
} satisfies Meta;

export default meta;
type Story = StoryObj<typeof meta>;

// ------------------------------------------------------------------ helpers

function AttainmentBadge({ row, col, display }: CellContentProps): ReactNode {
  if (col !== 4 || row < 3 || row > 8 || !display.includes("%")) {
    return display;
  }
  const pct = Number.parseFloat(display);
  const hit = !Number.isNaN(pct) && pct >= 100;
  const color = hit ? "var(--a1s-accent)" : "#f87171";
  return (
    <span
      style={{
        display: "inline-flex",
        alignItems: "center",
        padding: "2px 10px",
        borderRadius: 999,
        fontSize: 12,
        fontWeight: 600,
        background: hit ? "rgba(45,212,191,0.15)" : "rgba(248,113,113,0.12)",
        color,
        border: `1px solid ${color}44`,
      }}
    >
      {display}
    </span>
  );
}

function SaveButton() {
  const api = useSheet();
  const [saved, setSaved] = useState(0);
  return (
    <button
      type="button"
      className="a1s-btn"
      onClick={() => {
        void api.workbook;
        setSaved((n) => n + 1);
      }}
    >
      {saved === 0 ? "Save workbook" : `Saved ×${saved}`}
    </button>
  );
}

/** Your own panel — reads the same context as every primitive. */
function LiveInspector() {
  const api = useSheet();
  const { row, col } = api.active;
  const { r1, c1, r2, c2 } = api.bounds;
  const count = (r2 - r1 + 1) * (c2 - c1 + 1);

  return (
    <aside
      style={{
        font: "13px/1.5 system-ui, sans-serif",
        background: "var(--a1s-header-bg)",
        borderLeft: "1px solid var(--a1s-border)",
        padding: "12px 14px",
        overflow: "auto",
        color: "var(--a1s-cell-text)",
      }}
    >
      <h3
        style={{
          margin: "0 0 10px",
          fontSize: 11,
          letterSpacing: "0.08em",
          textTransform: "uppercase",
          color: "var(--a1s-header-text)",
        }}
      >
        useSheet() inspector
      </h3>
      <dl style={{ margin: 0, display: "grid", gap: 8 }}>
        <InspectorRow label="Sheet" value={api.sheet.name} />
        <InspectorRow label="Dimensions" value={`${api.sheet.numRows} × ${api.sheet.numCols}`} />
        <InspectorRow label="Active" value={toA1(row, col)} />
        <InspectorRow label="Selection" value={`${count} cell${count === 1 ? "" : "s"}`} />
        <InspectorRow label="Display" value={api.getDisplay(row, col) || "—"} />
        <InspectorRow label="Raw" value={api.getRaw(row, col) || "—"} mono />
        <InspectorRow
          label="History"
          value={`${api.canUndo ? "undo" : "—"} / ${api.canRedo ? "redo" : "—"}`}
        />
        <InspectorRow label="Filled cells" value={String(Object.keys(api.sheet.cells).length)} />
      </dl>
      <p style={{ margin: "14px 0 0", fontSize: 12, color: "var(--a1s-header-text)" }}>
        This panel is not a library primitive — it is a child of <code>Sheet.Root</code> calling{" "}
        <code>useSheet()</code>, the same hook the toolbar uses.
      </p>
    </aside>
  );
}

function InspectorRow({
  label,
  value,
  mono,
}: {
  label: string;
  value: string;
  mono?: boolean;
}) {
  return (
    <div style={{ display: "grid", gridTemplateColumns: "88px 1fr", gap: 8 }}>
      <dt style={{ margin: 0, color: "var(--a1s-header-text)" }}>{label}</dt>
      <dd
        style={{
          margin: 0,
          fontFamily: mono ? "var(--a1s-mono-font-family)" : undefined,
          overflow: "hidden",
          textOverflow: "ellipsis",
        }}
      >
        {value}
      </dd>
    </div>
  );
}

function FeatureCallout({ children }: { children: ReactNode }) {
  return (
    <p
      style={{
        margin: 0,
        padding: "8px 12px",
        font: "12px/1.45 system-ui",
        color: "var(--a1s-header-text)",
        borderBottom: "1px solid var(--a1s-border)",
        background: "var(--a1s-toolbar-bg)",
      }}
    >
      {children}
    </p>
  );
}

// ------------------------------------------------------------------ stories

export const FullApplication: Story = {
  name: "Full application shell",
  render: () => (
    <div
      style={{
        height: "100vh",
        padding: 16,
        boxSizing: "border-box",
        background: "#0b1220",
        display: "flex",
        flexDirection: "column",
        gap: 12,
      }}
    >
      <header style={{ font: "600 15px/1.4 system-ui", color: "#e2e8f0" }}>
        Composed primitives · Q3 report · conditional formatting · frozen panes ·
        formulas · custom cell renderer · your own sidebar
      </header>

      <Sheet.Root
        defaultWorkbook={salesReport()}
        theme={darkTheme}
        height="100%"
        className="rounded-xl shadow-2xl overflow-hidden flex-1 min-h-0"
        style={{ border: "1px solid var(--a1s-border)" }}
      >
        <FeatureCallout>
          <strong>Toolbar</strong> is composed atom-by-atom (<code>Sheet.Toolbar.Undo</code>,{" "}
          <code>.Bold</code>, <code>.Overflow</code>, …) plus <code>Sheet.FileMenu</code> and a
          custom Save button. Try formatting, freeze, merge, import/export.
        </FeatureCallout>

        <Sheet.Toolbar style={{ flexWrap: "wrap" }}>
          <Sheet.Toolbar.Group>
            <Sheet.Toolbar.Undo />
            <Sheet.Toolbar.Redo />
          </Sheet.Toolbar.Group>
          <Sheet.Toolbar.Separator />
          <Sheet.Toolbar.Group>
            <Sheet.Toolbar.FontFamily />
            <Sheet.Toolbar.Bold />
            <Sheet.Toolbar.Italic />
            <Sheet.Toolbar.Underline />
          </Sheet.Toolbar.Group>
          <Sheet.Toolbar.Separator />
          <Sheet.Toolbar.Group>
            <Sheet.Toolbar.AlignLeft />
            <Sheet.Toolbar.AlignCenter />
            <Sheet.Toolbar.AlignRight />
          </Sheet.Toolbar.Group>
          <Sheet.Toolbar.TextColor />
          <Sheet.Toolbar.FillColor />
          <Sheet.Toolbar.NumFmt />
          <Sheet.Toolbar.Lock />
          <Sheet.Toolbar.Separator />
          <Sheet.Toolbar.Overflow menuLabel="Structure tools">
            <Sheet.Toolbar.Group>
              <Sheet.Toolbar.InsertRow />
              <Sheet.Toolbar.DeleteRow />
              <Sheet.Toolbar.InsertCol />
              <Sheet.Toolbar.DeleteCol />
            </Sheet.Toolbar.Group>
            <Sheet.Toolbar.Separator />
            <Sheet.Toolbar.Group>
              <Sheet.Toolbar.Merge />
              <Sheet.Toolbar.Unmerge />
              <Sheet.Toolbar.Freeze />
              <Sheet.Toolbar.Unfreeze />
            </Sheet.Toolbar.Group>
          </Sheet.Toolbar.Overflow>
          <Sheet.Toolbar.Separator />
          <Sheet.FileMenu />
          <SaveButton />
          <Sheet.Toolbar.Status />
        </Sheet.Toolbar>

        <Sheet.FormulaBar />

        <div style={{ display: "flex", flex: 1, minHeight: 0 }}>
          <div style={{ display: "flex", flexDirection: "column", flex: 1, minWidth: 0 }}>
            <FeatureCallout>
              <strong>Grid</strong> — scroll frozen Q3 data; attainment column uses{" "}
              <code>renderCellContent</code>. Column header menu sorts/filters. Right-click or
              long-press for context menu.
            </FeatureCallout>
            <Sheet.Grid
              className="flex-1 min-h-0"
              renderCellContent={(props) => <AttainmentBadge {...props} />}
            >
              <Sheet.AddRows />
            </Sheet.Grid>
            <Sheet.Tabs />
            <Sheet.StatusBar />
          </div>
          <LiveInspector />
        </div>

        <Sheet.ContextMenu>
          <Sheet.ContextMenu.Copy />
          <Sheet.ContextMenu.Paste />
          <Sheet.ContextMenu.Separator />
          <Sheet.ContextMenu.ClearContents />
          <Sheet.ContextMenu.Separator />
          <Sheet.ContextMenu.Item onSelect={() => {}}>
            Push to data warehouse (your code)
          </Sheet.ContextMenu.Item>
        </Sheet.ContextMenu>
        <Sheet.ColumnMenu />
      </Sheet.Root>
    </div>
  ),
  parameters: {
    docs: {
      description: {
        story:
          "Everything a product team would assemble: formatting toolbar, formula bar, virtualized grid " +
          "with frozen rows and conditional colours, sheet tabs, status aggregates, file I/O, " +
          "column filter menu, context menu with a custom item, add-rows, and a sidebar driven by useSheet().",
      },
    },
  },
};

export const AtScale: Story = {
  name: "25,000 rows — virtualization",
  render: () => {
    const rows = 25_000;
    const wb = large(rows);
    return (
      <div
        style={{
          height: "100vh",
          padding: 16,
          boxSizing: "border-box",
          background: "#f8fafc",
          display: "flex",
          flexDirection: "column",
          gap: 12,
        }}
      >
        <header style={{ font: "600 15px/1.4 system-ui", color: "#0f172a" }}>
          {rows.toLocaleString()} rows × {wb.sheets[0]?.numCols ?? 0} columns — only ~300 cells
          exist in the DOM at once
        </header>
        <Sheet.Root
          defaultWorkbook={wb}
          height="100%"
          className="rounded-xl shadow-lg overflow-hidden flex-1 min-h-0"
        >
          <Sheet.Toolbar>
            <Sheet.Toolbar.Default />
            <Sheet.Toolbar.Separator />
            <Sheet.Toolbar.Status />
          </Sheet.Toolbar>
          <Sheet.FormulaBar />
          <Sheet.Grid className="flex-1 min-h-0">
            <Sheet.AddRows />
          </Sheet.Grid>
          <Sheet.StatusBar />
          <Sheet.ContextMenu />
          <Sheet.ColumnMenu />
        </Sheet.Root>
      </div>
    );
  },
  parameters: {
    docs: {
      description: {
        story:
          "Scroll to row 20,000 — cost stays flat because both axes are virtualized. " +
          "Frozen header row, filters, and editing all still work at this size.",
      },
    },
  },
};

export const GridOnlyMinimal: Story = {
  name: "Grid only (minimal chrome)",
  render: () => (
    <Sheet.Root defaultWorkbook={salesReport()} height={480} className="rounded-lg shadow">
      <Sheet.Grid renderCellContent={(p) => <AttainmentBadge {...p} />} />
    </Sheet.Root>
  ),
};

export const CustomToolbarOnly: Story = {
  name: "Toolbar atoms + overflow",
  render: () => (
    <Sheet.Root defaultWorkbook={budget()} height={420}>
      <Sheet.Toolbar>
        <Sheet.Toolbar.Undo />
        <Sheet.Toolbar.Bold />
        <Sheet.Toolbar.Overflow>
          <Sheet.Toolbar.Merge />
          <Sheet.Toolbar.Freeze />
        </Sheet.Toolbar.Overflow>
        <Sheet.FileMenu />
      </Sheet.Toolbar>
      <Sheet.Grid />
    </Sheet.Root>
  ),
};
