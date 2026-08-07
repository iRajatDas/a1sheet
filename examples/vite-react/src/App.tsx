import { toA1 } from "a1sheet";
import {
  darkTheme,
  type CellContentProps,
  Sheet,
  useSheet,
} from "a1sheet/react";
import { type ReactNode, useMemo, useState } from "react";
import { large, salesReport } from "./fixtures.js";

type Mode = "full" | "scale";

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

function LiveInspector() {
  const api = useSheet();
  const { row, col } = api.active;
  const { r1, c1, r2, c2 } = api.bounds;
  const count = (r2 - r1 + 1) * (c2 - c1 + 1);

  return (
    <aside
      style={{
        width: 260,
        flexShrink: 0,
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
        <InspectorRow
          label="Dimensions"
          value={`${api.sheet.numRows} × ${api.sheet.numCols}`}
        />
        <InspectorRow label="Active" value={toA1(row, col)} />
        <InspectorRow
          label="Selection"
          value={`${count} cell${count === 1 ? "" : "s"}`}
        />
        <InspectorRow label="Display" value={api.getDisplay(row, col) || "—"} />
        <InspectorRow label="Raw" value={api.getRaw(row, col) || "—"} mono />
        <InspectorRow
          label="History"
          value={`${api.canUndo ? "undo" : "—"} / ${api.canRedo ? "redo" : "—"}`}
        />
        <InspectorRow
          label="Filled cells"
          value={String(Object.keys(api.sheet.cells).length)}
        />
      </dl>
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

function ModeSwitch({
  mode,
  onModeChange,
}: {
  mode: Mode;
  onModeChange: (mode: Mode) => void;
}) {
  return (
    <nav
      style={{
        display: "flex",
        gap: 8,
        alignItems: "center",
        font: "13px system-ui",
      }}
    >
      <span style={{ fontWeight: 600 }}>a1sheet playground</span>
      <button
        type="button"
        onClick={() => onModeChange("full")}
        style={tabStyle(mode === "full")}
      >
        Full shell
      </button>
      <button
        type="button"
        onClick={() => onModeChange("scale")}
        style={tabStyle(mode === "scale")}
      >
        25k rows
      </button>
    </nav>
  );
}

function tabStyle(active: boolean): React.CSSProperties {
  return {
    padding: "6px 12px",
    borderRadius: 6,
    border: "1px solid",
    borderColor: active ? "#2dd4bf" : "#334155",
    background: active ? "rgba(45,212,191,0.12)" : "transparent",
    color: active ? "#99f6e4" : "#94a3b8",
    cursor: "pointer",
    font: "inherit",
  };
}

function FullShell() {
  return (
    <Sheet.Root
      defaultWorkbook={salesReport()}
      theme={darkTheme}
      height="100%"
      style={{
        border: "1px solid var(--a1s-border)",
        borderRadius: 12,
        overflow: "hidden",
        display: "flex",
        flexDirection: "column",
        flex: 1,
        minHeight: 0,
      }}
    >
      <FeatureCallout>
        Composed toolbar atoms, <code>renderCellContent</code>, file I/O, context
        menu, column filters, and a custom <code>useSheet()</code> sidebar.
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
        <div
          style={{
            display: "flex",
            flexDirection: "column",
            flex: 1,
            minWidth: 0,
          }}
        >
          <Sheet.Grid
            style={{ flex: 1, minHeight: 0 }}
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
  );
}

function AtScaleShell() {
  const rows = 25_000;
  const workbook = useMemo(() => large(rows), [rows]);
  const sheet = workbook.sheets[0];

  return (
    <>
      <p style={{ margin: 0, font: "13px/1.4 system-ui", color: "#94a3b8" }}>
        {rows.toLocaleString()} rows × {sheet?.numCols ?? 0} columns — only ~300
        cells in the DOM. Scroll to row 20,000 to verify virtualization.
      </p>
      <Sheet.Root
        defaultWorkbook={workbook}
        height="100%"
        style={{
          border: "1px solid #e2e8f0",
          borderRadius: 12,
          overflow: "hidden",
          display: "flex",
          flexDirection: "column",
          flex: 1,
          minHeight: 0,
        }}
      >
        <Sheet.Toolbar>
          <Sheet.Toolbar.Default />
          <Sheet.Toolbar.Separator />
          <Sheet.Toolbar.Status />
        </Sheet.Toolbar>
        <Sheet.FormulaBar />
        <Sheet.Grid style={{ flex: 1, minHeight: 0 }}>
          <Sheet.AddRows />
        </Sheet.Grid>
        <Sheet.StatusBar />
        <Sheet.ContextMenu />
        <Sheet.ColumnMenu />
      </Sheet.Root>
    </>
  );
}

function readMode(): Mode {
  const value = new URLSearchParams(window.location.search).get("mode");
  return value === "scale" ? "scale" : "full";
}

export function App() {
  const [mode, setMode] = useState<Mode>(readMode);

  const selectMode = (next: Mode) => {
    setMode(next);
    const url = new URL(window.location.href);
    if (next === "full") {
      url.searchParams.delete("mode");
    } else {
      url.searchParams.set("mode", next);
    }
    window.history.replaceState(null, "", url);
  };

  const isFull = mode === "full";

  return (
    <div
      style={{
        height: "100vh",
        boxSizing: "border-box",
        padding: 16,
        display: "flex",
        flexDirection: "column",
        gap: 12,
        background: isFull ? "#0b1220" : "#f8fafc",
        color: isFull ? "#e2e8f0" : "#0f172a",
      }}
    >
      <ModeSwitch mode={mode} onModeChange={selectMode} />
      {isFull ? <FullShell /> : <AtScaleShell />}
    </div>
  );
}
