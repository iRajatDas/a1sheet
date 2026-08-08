/**
 * Recent API demos with product-shaped lab chrome (grouped tools, primary
 * actions, live status) — not a flat button dump.
 */
import type { Meta, StoryObj } from "@storybook/react-vite";
import { suggestFormulas, toA1 } from "a1sheet";
import { Sheet, useSheet } from "a1sheet/react";
import { type ReactNode, useState } from "react";
import {
  colourFilterDemo,
  findReplaceDemo,
  interactiveCellsDemo,
  sequencesDemo,
} from "./fixtures.js";
import {
  formatActive,
  LabButton,
  LabField,
  LabIntro,
  LabMeta,
  LabRow,
  LabSection,
  LabStatusBar,
  LabSteps,
  LabSwatch,
  LabToolbar,
  scrollHitIntoView,
} from "./labChrome.js";

const meta = {
  title: "Features/Recent APIs",
  tags: ["autodocs"],
  parameters: {
    docs: {
      description: {
        component:
          "Task 1–10 surfaces with a real control layout: colour filters, find/replace, sheet tabs, checkboxes/links, paste special, and formula catalog.",
      },
    },
  },
} satisfies Meta;

export default meta;
type Story = StoryObj<typeof meta>;

const shell = {
  height: 480,
  style: {
    border: "1px solid #e2e8f0",
    borderRadius: 12,
    overflow: "hidden",
    display: "flex",
    flexDirection: "column",
  } as const,
};

// ------------------------------------------------------------------ paste

export const PasteSpecial: Story = {
  name: "Paste special",
  render: () => (
    <Sheet.Root defaultWorkbook={sequencesDemo()} {...shell}>
      <LabIntro
        title="Paste special"
        body="Copy a block, then choose how it lands — values only, formats, formulas, transposed, or as text. Watch the status line for confirmation."
      />
      <LabToolbar>
        <LabSteps
          steps={[
            "Select A2:A3 (Mon, Tue) and Copy from the context menu.",
            "Select an empty cell and right-click → Paste transposed or Paste values.",
            "Status reports the mode, e.g. Pasted values into …",
          ]}
        />
        <LabStatusBar />
      </LabToolbar>
      <Sheet.Grid style={{ flex: 1, minHeight: 0 }} />
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
    </Sheet.Root>
  ),
};

// --------------------------------------------------------------- fill seq

export const FillSequences: Story = {
  name: "Fill sequences and guards",
  render: () => (
    <Sheet.Root defaultWorkbook={sequencesDemo()} {...shell}>
      <LabIntro
        title="Fill handle"
        body="Weekdays and months continue with wraparound. A fill that stretches both axes at once is rejected — the status line explains why."
      />
      <LabToolbar>
        <LabSteps
          steps={[
            "Select A2:A3 (Mon → Tue).",
            "Drag the square at the bottom-right corner down several rows.",
            "Try a diagonal drag that changes both row and column — it should refuse.",
          ]}
        />
        <LabStatusBar />
      </LabToolbar>
      <Sheet.Grid style={{ flex: 1, minHeight: 0 }} />
      <Sheet.StatusBar />
    </Sheet.Root>
  ),
};

// ----------------------------------------------------------- colour filter

function ColourFilterPanel(): ReactNode {
  const api = useSheet();
  const green = "#bbf7d0";
  const redText = "#7f1d1d";
  const viewId = "greens";

  return (
    <>
      <LabIntro
        title="Colour filters & views"
        body="Filter by fill or text colour without touching cell values. Save a named view, switch back later, or sort matching colours to the top."
      />
      <LabToolbar>
        <LabSection label="Filter by colour" hint="Column A">
          <LabRow>
            <LabSwatch
              color={green}
              label="Green fills"
              selected={api.sheet.filters[0]?.background?.has(green) === true}
              onClick={() =>
                api.setFilter(0, { background: new Set([green]) })
              }
            />
            <LabSwatch
              color={redText}
              label="Red text"
              selected={api.sheet.filters[0]?.foreground?.has(redText) === true}
              onClick={() =>
                api.setFilter(0, { foreground: new Set([redText]) })
              }
            />
            <LabButton kind="ghost" onClick={() => api.setFilter(0, null)}>
              Clear
            </LabButton>
          </LabRow>
        </LabSection>
        <LabSection label="Named views" hint="Snapshots of the live filters">
          <LabRow>
            <LabButton
              kind="primary"
              onClick={() =>
                api.createFilterView({ id: viewId, name: "Greens" })
              }
            >
              Save as “Greens”
            </LabButton>
            <LabButton onClick={() => api.activateFilterView(viewId)}>
              Open “Greens”
            </LabButton>
            <LabButton
              kind="ghost"
              onClick={() => api.activateFilterView("missing")}
              title="Demonstrates The view does not exist."
            >
              Open missing…
            </LabButton>
          </LabRow>
        </LabSection>
        <LabSection label="Sort">
          <LabRow>
            <LabButton
              onClick={() =>
                api.sortByColor({
                  col: 0,
                  kind: "background",
                  color: green,
                })
              }
            >
              Move green fills to top
            </LabButton>
          </LabRow>
        </LabSection>
        <LabStatusBar />
        <LabMeta
          items={[
            {
              label: "Active view",
              value: api.sheet.activeFilterViewId ?? "none",
            },
            {
              label: "Filtered cols",
              value: String(Object.keys(api.sheet.filters).length),
            },
          ]}
        />
      </LabToolbar>
    </>
  );
}

export const ColourFilters: Story = {
  name: "Colour filters and filter views",
  render: () => (
    <Sheet.Root defaultWorkbook={colourFilterDemo()} {...shell}>
      <ColourFilterPanel />
      <Sheet.Grid style={{ flex: 1, minHeight: 0 }} />
      <Sheet.ColumnMenu />
      <Sheet.StatusBar />
    </Sheet.Root>
  ),
};

// ------------------------------------------------------------ find/replace

function FindReplacePanel(): ReactNode {
  const api = useSheet();
  const [needle, setNeedle] = useState("alpha");
  const [replacement, setReplacement] = useState("OMEGA");
  const [last, setLast] = useState("");

  const findNext = () => {
    const hit = api.findNext({ find: needle, after: api.active });
    if (!hit) {
      setLast("No more matches");
      return;
    }
    api.selectCell(hit.row, hit.col);
    scrollHitIntoView(api, hit.row);
    setLast(`Selected ${toA1(hit.row, hit.col)}`);
  };

  const replaceAll = () => {
    const n = api.replaceAll({ find: needle, replace: replacement });
    setLast(n === 0 ? "Nothing to replace" : `Updated ${n} cell${n === 1 ? "" : "s"}`);
  };

  return (
    <>
      <LabIntro
        title="Find & replace"
        body="Headless search over raw cell text and formulas. Find next wraps the sheet; replace all reports a count in the status line."
      />
      <LabToolbar>
        <LabSection label="Search">
          <LabRow>
            <LabField
              label="Find"
              value={needle}
              onChange={setNeedle}
              placeholder="Text or formula fragment"
              width={200}
            />
            <LabField
              label="Replace with"
              value={replacement}
              onChange={setReplacement}
              width={160}
            />
            <div style={{ display: "flex", gap: 8, paddingTop: 16 }}>
              <LabButton kind="primary" onClick={findNext}>
                Find next
              </LabButton>
              <LabButton onClick={replaceAll}>Replace all</LabButton>
            </div>
          </LabRow>
        </LabSection>
        <LabStatusBar />
        <LabMeta
          items={[
            { label: "Cursor", value: formatActive(api) },
            { label: "Last action", value: last || "—" },
            {
              label: "Hits now",
              value: String(api.findAll({ find: needle }).length),
            },
          ]}
        />
      </LabToolbar>
    </>
  );
}

export const FindReplace: Story = {
  name: "Find and replace",
  render: () => (
    <Sheet.Root defaultWorkbook={findReplaceDemo()} {...shell}>
      <FindReplacePanel />
      <Sheet.Grid style={{ flex: 1, minHeight: 0 }} />
      <Sheet.StatusBar />
    </Sheet.Root>
  ),
};

// --------------------------------------------------------------- sheet tabs

function SheetTabsPanel(): ReactNode {
  const api = useSheet();
  const i = api.workbook.activeSheetIndex;
  const names = api.workbook.sheets.map((s) => s.name);

  return (
    <>
      <LabIntro
        title="Sheet tabs"
        body="Duplicate deep-copies filters and styles. Reorder with move; naming a range reports in status."
      />
      <LabToolbar>
        <LabSection label="Tabs">
          <LabRow>
            <LabButton
              kind="primary"
              onClick={() => api.duplicateSheetAt(i)}
            >
              Duplicate “{api.sheet.name}”
            </LabButton>
            <LabButton
              disabled={i === 0}
              onClick={() => api.moveSheetAt(i, i - 1)}
            >
              Move left
            </LabButton>
            <LabButton
              disabled={i >= api.workbook.sheets.length - 1}
              onClick={() => api.moveSheetAt(i, i + 1)}
            >
              Move right
            </LabButton>
            <LabButton
              kind="ghost"
              onClick={() =>
                api.defineName("Tasks", { r1: 1, c1: 0, r2: 2, c2: 0 })
              }
            >
              Name A2:A3 “Tasks”
            </LabButton>
          </LabRow>
        </LabSection>
        <LabStatusBar />
        <LabMeta
          items={[
            { label: "Order", value: names.join(" → ") },
            { label: "Active", value: api.sheet.name },
          ]}
        />
      </LabToolbar>
    </>
  );
}

export const SheetTabOps: Story = {
  name: "Duplicate and move sheets",
  render: () => (
    <Sheet.Root defaultWorkbook={interactiveCellsDemo()} {...shell}>
      <SheetTabsPanel />
      <Sheet.Grid style={{ flex: 1, minHeight: 0 }} />
      <Sheet.Tabs />
      <Sheet.StatusBar />
    </Sheet.Root>
  ),
};

// ---------------------------------------------------- checkbox / link / rot

function InteractivePanel(): ReactNode {
  const api = useSheet();

  return (
    <>
      <LabIntro
        title="Checkboxes, links, rotation"
        body="Toggle B2 with the checkbox. Actions below target empty cells so you can see hyperlinks and rotation land on the grid."
      />
      <LabToolbar>
        <LabSection label="Insert">
          <LabRow>
            <LabButton
              kind="primary"
              onClick={() => {
                api.select({ r1: 3, c1: 1, r2: 3, c2: 1 });
                api.insertCheckboxes();
              }}
            >
              Checkbox in B4
            </LabButton>
            <LabButton
              onClick={() => {
                api.selectCell(2, 2);
                api.setCell(2, 2, "Home");
                api.setHyperlink("https://example.com");
              }}
            >
              Link C3 → example.com
            </LabButton>
            <LabButton
              onClick={() => {
                api.selectCell(2, 3);
                api.setCell(2, 3, "Tilt");
                api.setTextRotation(-45);
              }}
            >
              Rotate D3 −45°
            </LabButton>
          </LabRow>
        </LabSection>
        <LabSection label="Freeze panes">
          <LabRow>
            <LabButton
              onClick={() => {
                api.select({ r1: 0, c1: 0, r2: 1, c2: 0 });
                api.freezeToSelection();
              }}
            >
              Freeze through A2
            </LabButton>
            <LabButton kind="ghost" onClick={() => api.unfreeze()}>
              Unfreeze
            </LabButton>
          </LabRow>
        </LabSection>
        <LabStatusBar />
        <LabMeta
          items={[
            {
              label: "B2",
              value: api.sheet.styles["1_1"]?.checkbox
                ? `checkbox · ${api.getRaw(1, 1)}`
                : "—",
            },
            {
              label: "C2 link",
              value: api.sheet.styles["1_2"]?.hyperlink ?? "—",
            },
            {
              label: "Freeze",
              value: `${api.sheet.frozenRows}r × ${api.sheet.frozenCols}c`,
            },
          ]}
        />
      </LabToolbar>
    </>
  );
}

export const InteractiveCells: Story = {
  name: "Checkboxes, links, rotation, freeze status",
  render: () => (
    <Sheet.Root defaultWorkbook={interactiveCellsDemo()} {...shell}>
      <InteractivePanel />
      <Sheet.Grid style={{ flex: 1, minHeight: 0 }} />
      <Sheet.StatusBar />
    </Sheet.Root>
  ),
};

// ---------------------------------------------------------- formula catalog

function CatalogPanel(): ReactNode {
  const [prefix, setPrefix] = useState("XL");
  const hits = suggestFormulas(prefix);

  return (
    <>
      <LabIntro
        title="Formula catalog"
        body="Autocomplete metadata from the Sheets help categories. Type a prefix to see implementable matches — B4 already holds =FLATTEN(A2:A3)."
      />
      <LabToolbar>
        <LabSection label="Suggest">
          <LabRow>
            <LabField
              label="Prefix"
              value={prefix}
              onChange={setPrefix}
              placeholder="e.g. XL, SEQ, TEX"
              width={140}
            />
            <div
              style={{
                display: "flex",
                flexWrap: "wrap",
                gap: 6,
                paddingTop: 16,
                maxWidth: 480,
              }}
            >
              {hits.length === 0 ? (
                <span style={{ color: "#64748b" }}>No matches</span>
              ) : (
                hits.map((h) => (
                  <span
                    key={h.name}
                    title={h.shortDescription}
                    style={{
                      font: "12px/1 ui-monospace, Menlo, monospace",
                      padding: "5px 9px",
                      borderRadius: 999,
                      background: "rgba(13, 148, 136, 0.12)",
                      color: "#0f766e",
                    }}
                  >
                    {h.name}
                    <span style={{ opacity: 0.55, marginLeft: 6 }}>
                      {h.category}
                    </span>
                  </span>
                ))
              )}
            </div>
          </LabRow>
        </LabSection>
        <LabMeta
          items={[
            {
              label: "Top hit",
              value: hits[0]
                ? `${hits[0].name} · ${hits[0].args.length} args`
                : "—",
            },
          ]}
        />
      </LabToolbar>
    </>
  );
}

export const FormulaCatalog: Story = {
  name: "Formula catalog and FLATTEN",
  render: () => (
    <Sheet.Root defaultWorkbook={interactiveCellsDemo()} {...shell}>
      <CatalogPanel />
      <Sheet.FormulaBar />
      <Sheet.Grid style={{ flex: 1, minHeight: 0 }} />
      <Sheet.StatusBar />
    </Sheet.Root>
  ),
};
