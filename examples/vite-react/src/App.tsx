import { suggestFormulas, toA1 } from "a1sheet";
import {
  type CellContentProps,
  darkTheme,
  Sheet,
  sheetsTheme,
  useSheet,
} from "a1sheet/react";
import { type CSSProperties, type ReactNode, useMemo, useState } from "react";
import {
  formatActive,
  LabButton,
  LabField,
  LabIntro,
  LabMeta,
  LabRow,
  LabSection,
  LabSegmented,
  LabStatusBar,
  LabSteps,
  LabSwatch,
  LabToolbar,
  scrollHitIntoView,
} from "../../storybook/stories/labChrome.js";
import {
  colourFilterDemo,
  findReplaceDemo,
  interactiveCellsDemo,
  large,
  salesReport,
  sequencesDemo,
} from "./fixtures.js";

type Mode = "full" | "scale" | "lab";
type LabSheet = "colour" | "find" | "interactive" | "sequences";

function AttainmentBadge({ row, col, display }: CellContentProps): ReactNode {
  if (col !== 4 || row < 3 || row > 8 || !display.includes("%")) {
    return display;
  }
  const pct = Number.parseFloat(display);
  const hit = !Number.isNaN(pct) && pct >= 100;
  const color = hit ? "#188038" : "#d93025";
  return (
    <span
      style={{
        display: "inline-flex",
        alignItems: "center",
        padding: "1px 8px",
        borderRadius: 4,
        fontSize: 12,
        fontWeight: 600,
        background: hit ? "rgba(24,128,56,0.12)" : "rgba(217,48,37,0.1)",
        color,
        border: `1px solid ${color}33`,
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
        <InspectorRow label="Status" value={api.status || "—"} />
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

function ModeSwitch({
  mode,
  onModeChange,
}: {
  mode: Mode;
  onModeChange: (mode: Mode) => void;
}) {
  const light = mode === "full" || mode === "scale";
  return (
    <nav
      style={{
        display: "flex",
        gap: 8,
        alignItems: "center",
        font: light
          ? '13px/1.4 Arial, "Helvetica Neue", Helvetica, sans-serif'
          : "13px system-ui",
        flexWrap: "wrap",
        color: light ? "#202124" : undefined,
      }}
    >
      <span style={{ fontWeight: 600 }}>
        {mode === "full" ? "Untitled spreadsheet" : "a1sheet playground"}
      </span>
      {(
        [
          ["full", "Full shell"],
          ["lab", "Feature lab"],
          ["scale", "25k rows"],
        ] as const
      ).map(([id, label]) => (
        <button
          key={id}
          type="button"
          onClick={() => onModeChange(id)}
          style={tabStyle(mode === id, light)}
        >
          {label}
        </button>
      ))}
    </nav>
  );
}

function tabStyle(active: boolean, light: boolean): CSSProperties {
  if (light) {
    return {
      padding: "6px 14px",
      borderRadius: 18,
      border: "1px solid",
      borderColor: active ? "#1a73e8" : "#dadce0",
      background: active ? "rgba(26, 115, 232, 0.08)" : "#fff",
      color: active ? "#1967d2" : "#3c4043",
      cursor: "pointer",
      font: "inherit",
      fontWeight: active ? 600 : 500,
    };
  }
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

function DocTitleBar({
  showInspector,
  onToggleInspector,
}: {
  showInspector: boolean;
  onToggleInspector: () => void;
}) {
  return (
    <div
      style={{
        display: "flex",
        alignItems: "center",
        gap: 12,
        padding: "8px 12px 6px",
        borderBottom: "1px solid #dadce0",
        background: "#fff",
        font: '14px/1.3 Arial, "Helvetica Neue", Helvetica, sans-serif',
        color: "#202124",
      }}
    >
      <div
        style={{
          width: 28,
          height: 28,
          borderRadius: 4,
          background: "#0f9d58",
          color: "#fff",
          display: "grid",
          placeItems: "center",
          fontWeight: 700,
          fontSize: 15,
          flexShrink: 0,
        }}
        aria-hidden
      >
        S
      </div>
      <div style={{ flex: 1, minWidth: 0 }}>
        <div style={{ fontWeight: 500, fontSize: 18, letterSpacing: "-0.01em" }}>
          Sales report
        </div>
        <div style={{ fontSize: 12, color: "#5f6368" }}>
          Toolbar · formula bar · grid · tabs — `sheetsTheme`
        </div>
      </div>
      <button
        type="button"
        onClick={onToggleInspector}
        style={{
          padding: "6px 12px",
          borderRadius: 4,
          border: "1px solid #dadce0",
          background: showInspector ? "rgba(26, 115, 232, 0.08)" : "#fff",
          color: showInspector ? "#1967d2" : "#3c4043",
          cursor: "pointer",
          font: "inherit",
          fontSize: 13,
          fontWeight: 500,
        }}
      >
        {showInspector ? "Hide inspector" : "Inspector"}
      </button>
    </div>
  );
}

function FullShell() {
  const [showInspector, setShowInspector] = useState(false);

  return (
    <Sheet.Root
      defaultWorkbook={salesReport()}
      theme={sheetsTheme}
      height="100%"
      style={{
        border: "1px solid #dadce0",
        borderRadius: 8,
        overflow: "hidden",
        display: "flex",
        flexDirection: "column",
        flex: 1,
        minHeight: 0,
        background: "#fff",
        boxShadow:
          "0 1px 2px rgba(60,64,67,0.15), 0 1px 3px 1px rgba(60,64,67,0.08)",
      }}
    >
      <DocTitleBar
        showInspector={showInspector}
        onToggleInspector={() => setShowInspector((v) => !v)}
      />

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
        {showInspector ? <LiveInspector /> : null}
      </div>

      <Sheet.ContextMenu>
        <Sheet.ContextMenu.Copy />
        <Sheet.ContextMenu.Paste />
        <Sheet.ContextMenu.PasteValues />
        <Sheet.ContextMenu.PasteFormats />
        <Sheet.ContextMenu.PasteFormulas />
        <Sheet.ContextMenu.PasteTranspose />
        <Sheet.ContextMenu.PasteText />
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
  const workbook = useMemo(() => large(rows), []);
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

function LabPanel({ sheetKind }: { sheetKind: LabSheet }): ReactNode {
  const api = useSheet();
  const tone = "dark" as const;
  const [needle, setNeedle] = useState("alpha");
  const [replacement, setReplacement] = useState("OMEGA");
  const [prefix, setPrefix] = useState("XL");
  const [last, setLast] = useState("");
  const catalogHits = suggestFormulas(prefix);
  const green = "#bbf7d0";
  const redText = "#7f1d1d";

  if (sheetKind === "colour") {
    return (
      <>
        <LabIntro
          tone={tone}
          title="Colour filters & views"
          body="Filter by fill or text colour. Save a named view, reopen it, or sort matching fills to the top — status confirms each step."
        />
        <LabToolbar tone={tone}>
          <LabSection tone={tone} label="Filter" hint="Column A">
            <LabRow>
              <LabSwatch
                tone={tone}
                color={green}
                label="Green fills"
                selected={api.sheet.filters[0]?.background?.has(green) === true}
                onClick={() => api.setFilter(0, { background: new Set([green]) })}
              />
              <LabSwatch
                tone={tone}
                color={redText}
                label="Red text"
                selected={api.sheet.filters[0]?.foreground?.has(redText) === true}
                onClick={() => api.setFilter(0, { foreground: new Set([redText]) })}
              />
              <LabButton
                tone={tone}
                kind="ghost"
                onClick={() => api.setFilter(0, null)}
              >
                Clear
              </LabButton>
            </LabRow>
          </LabSection>
          <LabSection tone={tone} label="Views & sort">
            <LabRow>
              <LabButton
                tone={tone}
                kind="primary"
                onClick={() =>
                  api.createFilterView({ id: "greens", name: "Greens" })
                }
              >
                Save “Greens”
              </LabButton>
              <LabButton
                tone={tone}
                onClick={() => api.activateFilterView("greens")}
              >
                Open “Greens”
              </LabButton>
              <LabButton
                tone={tone}
                onClick={() =>
                  api.sortByColor({
                    col: 0,
                    kind: "background",
                    color: green,
                  })
                }
              >
                Greens to top
              </LabButton>
            </LabRow>
          </LabSection>
          <LabStatusBar tone={tone} />
        </LabToolbar>
      </>
    );
  }

  if (sheetKind === "find") {
    return (
      <>
        <LabIntro
          tone={tone}
          title="Find & replace"
          body="Find next wraps the sheet. Replace all rewrites raw cells and formulas, then reports a count."
        />
        <LabToolbar tone={tone}>
          <LabSection tone={tone} label="Search">
            <LabRow>
              <LabField
                tone={tone}
                label="Find"
                value={needle}
                onChange={setNeedle}
                width={180}
              />
              <LabField
                tone={tone}
                label="Replace with"
                value={replacement}
                onChange={setReplacement}
                width={140}
              />
              <div style={{ display: "flex", gap: 8, paddingTop: 16 }}>
                <LabButton
                  tone={tone}
                  kind="primary"
                  onClick={() => {
                    const hit = api.findNext({
                      find: needle,
                      after: api.active,
                    });
                    if (!hit) {
                      setLast("No match");
                      return;
                    }
                    api.selectCell(hit.row, hit.col);
                    scrollHitIntoView(api, hit.row);
                    setLast(`At ${toA1(hit.row, hit.col)}`);
                  }}
                >
                  Find next
                </LabButton>
                <LabButton
                  tone={tone}
                  onClick={() => {
                    const n = api.replaceAll({
                      find: needle,
                      replace: replacement,
                    });
                    setLast(`Replaced ${n}`);
                  }}
                >
                  Replace all
                </LabButton>
              </div>
            </LabRow>
          </LabSection>
          <LabStatusBar tone={tone} />
          <LabMeta
            tone={tone}
            items={[
              { label: "Cursor", value: formatActive(api) },
              { label: "Last", value: last || "—" },
            ]}
          />
        </LabToolbar>
      </>
    );
  }

  if (sheetKind === "interactive") {
    const i = api.workbook.activeSheetIndex;
    return (
      <>
        <LabIntro
          tone={tone}
          title="Tabs, checkboxes, links"
          body="Duplicate sheets, insert checkboxes, attach links, rotate text — status and meta update as you go."
        />
        <LabToolbar tone={tone}>
          <LabSection tone={tone} label="Sheet tabs">
            <LabRow>
              <LabButton
                tone={tone}
                kind="primary"
                onClick={() => api.duplicateSheetAt(i)}
              >
                Duplicate “{api.sheet.name}”
              </LabButton>
              <LabButton
                tone={tone}
                disabled={i === 0}
                onClick={() => api.moveSheetAt(i, i - 1)}
              >
                Move left
              </LabButton>
              <LabButton
                tone={tone}
                disabled={i >= api.workbook.sheets.length - 1}
                onClick={() => api.moveSheetAt(i, i + 1)}
              >
                Move right
              </LabButton>
            </LabRow>
          </LabSection>
          <LabSection tone={tone} label="Cell chrome">
            <LabRow>
              <LabButton
                tone={tone}
                onClick={() => {
                  api.select({ r1: 3, c1: 1, r2: 3, c2: 1 });
                  api.insertCheckboxes();
                }}
              >
                Checkbox B4
              </LabButton>
              <LabButton
                tone={tone}
                onClick={() => {
                  api.selectCell(2, 2);
                  api.setCell(2, 2, "Home");
                  api.setHyperlink("https://example.com");
                }}
              >
                Link C3
              </LabButton>
              <LabButton
                tone={tone}
                onClick={() => {
                  api.selectCell(2, 3);
                  api.setCell(2, 3, "Tilt");
                  api.setTextRotation(-45);
                }}
              >
                Rotate D3
              </LabButton>
            </LabRow>
          </LabSection>
          <LabSection tone={tone} label="Catalog">
            <LabRow>
              <LabField
                tone={tone}
                label="Prefix"
                value={prefix}
                onChange={setPrefix}
                width={100}
              />
              <span style={{ color: "#94a3b8", paddingTop: 16 }}>
                {catalogHits.map((h) => h.name).join(" · ") || "No matches"}
              </span>
            </LabRow>
          </LabSection>
          <LabStatusBar tone={tone} />
          <LabMeta
            tone={tone}
            items={[
              {
                label: "Sheets",
                value: api.workbook.sheets.map((s) => s.name).join(" → "),
              },
            ]}
          />
        </LabToolbar>
      </>
    );
  }

  return (
    <>
      <LabIntro
        tone={tone}
        title="Fill & paste special"
        body="Drag the fill handle on weekday or month series. Right-click for paste values, formats, formulas, transposed, or as text."
      />
      <LabToolbar tone={tone}>
        <LabSteps
          tone={tone}
          steps={[
            "Select A2:A3 (Mon → Tue) and drag the fill handle down.",
            "Copy a block, select a destination, right-click → Paste transposed.",
            "Watch the status line for paste-mode confirmation.",
          ]}
        />
        <LabStatusBar tone={tone} />
      </LabToolbar>
    </>
  );
}

function LabShell() {
  const [sheetKind, setSheetKind] = useState<LabSheet>("colour");
  const workbook = useMemo(() => {
    switch (sheetKind) {
      case "colour":
        return colourFilterDemo();
      case "find":
        return findReplaceDemo();
      case "interactive":
        return interactiveCellsDemo();
      case "sequences":
        return sequencesDemo();
    }
  }, [sheetKind]);

  return (
    <Sheet.Root
      key={sheetKind}
      defaultWorkbook={workbook}
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
      <div
        style={{
          padding: "10px 16px",
          borderBottom: "1px solid var(--a1s-border)",
          background: "var(--a1s-header-bg)",
          display: "flex",
          alignItems: "center",
          gap: 12,
          flexWrap: "wrap",
        }}
      >
        <LabSegmented
          tone="dark"
          value={sheetKind}
          onChange={setSheetKind}
          options={[
            { id: "colour", label: "Colour filters" },
            { id: "find", label: "Find / replace" },
            { id: "interactive", label: "Tabs & cells" },
            { id: "sequences", label: "Fill & paste" },
          ]}
        />
      </div>
      <LabPanel sheetKind={sheetKind} />
      <Sheet.Toolbar>
        <Sheet.Toolbar.Undo />
        <Sheet.Toolbar.Redo />
        <Sheet.Toolbar.Separator />
        <Sheet.Toolbar.Freeze />
        <Sheet.Toolbar.Unfreeze />
        <Sheet.Toolbar.Separator />
        <Sheet.Toolbar.Status />
      </Sheet.Toolbar>
      <Sheet.FormulaBar />
      <Sheet.Grid style={{ flex: 1, minHeight: 0 }} />
      <Sheet.Tabs />
      <Sheet.StatusBar />
      <Sheet.ContextMenu>
        <Sheet.ContextMenu.Copy />
        <Sheet.ContextMenu.Paste />
        <Sheet.ContextMenu.PasteValues />
        <Sheet.ContextMenu.PasteFormats />
        <Sheet.ContextMenu.PasteFormulas />
        <Sheet.ContextMenu.PasteTranspose />
        <Sheet.ContextMenu.PasteText />
      </Sheet.ContextMenu>
      <Sheet.ColumnMenu />
    </Sheet.Root>
  );
}

function readMode(): Mode {
  const value = new URLSearchParams(window.location.search).get("mode");
  if (value === "scale" || value === "lab") return value;
  return "full";
}

export function App() {
  const [mode, setMode] = useState<Mode>(readMode);

  const selectMode = (next: Mode) => {
    setMode(next);
    const url = new URL(window.location.href);
    if (next === "full") url.searchParams.delete("mode");
    else url.searchParams.set("mode", next);
    window.history.replaceState(null, "", url);
  };

  const dark = mode === "lab";

  return (
    <div
      style={{
        height: "100vh",
        boxSizing: "border-box",
        padding: mode === "full" ? 12 : 16,
        display: "flex",
        flexDirection: "column",
        gap: mode === "full" ? 8 : 12,
        background: dark ? "#0b1220" : "#f8f9fa",
        color: dark ? "#e2e8f0" : "#202124",
      }}
    >
      <ModeSwitch mode={mode} onModeChange={selectMode} />
      {mode === "full" ? (
        <FullShell />
      ) : mode === "lab" ? (
        <LabShell />
      ) : (
        <AtScaleShell />
      )}
    </div>
  );
}
