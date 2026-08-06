/**
 * The headless layer. Behaviour lives in hooks; a component is a thin renderer
 * over one. These stories use no a1sheet component at all.
 */
import type { Meta, StoryObj } from "@storybook/react-vite";
import { cellKey, findRefSpans, toA1 } from "a1sheet";
import { Sheet, useSheet, useSpreadsheet } from "a1sheet/react";
import { useState } from "react";
import { budget, large } from "./fixtures.js";
import { Panel, Row } from "./ui.js";

const meta = {
  title: "Hooks/Headless",
  tags: ["autodocs"],
  parameters: {
    docs: {
      description: {
        component: [
          "Every behaviour is a hook, and the components are thin renderers over them. If a hook cannot run without rendering, it is in the wrong place.",
          "",
          "| Hook | What it owns |",
          "|---|---|",
          "| `useSpreadsheet` | everything, composed — the whole API in one object |",
          "| `useWorkbook` | the workbook, undo history, and the only clone-on-write |",
          "| `useSelection` | selection, the anchor, extra ranges |",
          "| `useEditing` | the open editor and its caret |",
          "| `useSheetOps` | formatting, structure, sort, filter, sizing |",
          "| `useClipboard` | copy, cut, paste, TSV interchange |",
          "| `useFillHandle` | fill-handle drag and series extrapolation |",
          "| `useFormulaRefs` | reference picking while a formula is typed |",
          "| `useRowWindow` / `useColWindow` | virtualization on each axis |",
          "| `useCaretBinding` | caret sync for a controlled editor |",
          "",
          "`useSheet()` is different: it reads the API out of context, for a component rendered *inside* `Sheet.Root`.",
        ].join("\n"),
      },
    },
  },
} satisfies Meta;

export default meta;
type Story = StoryObj<typeof meta>;

export const NoComponentsAtAll: Story = {
  name: "useSpreadsheet — your own UI",
  render: function HeadlessStory() {
    const api = useSpreadsheet({ initialWorkbook: budget() });
    const { r1, c1 } = api.bounds;

    return (
      <div style={{ font: "13px/1.5 system-ui", padding: 12 }}>
        <Panel title="A grid rendered entirely by this story">
          <Row label="active">{toA1(api.active.row, api.active.col)}</Row>
          <Row label="value">
            <code>{api.getDisplay(r1, c1) || "empty"}</code>
          </Row>
          <Row label="raw">
            <code>{api.getRaw(r1, c1) || "empty"}</code>
          </Row>
        </Panel>

        <table style={{ borderCollapse: "collapse" }}>
          <tbody>
            {Array.from({ length: 8 }, (_, row) => (
              <tr key={cellKey(row, 0)}>
                {Array.from({ length: 5 }, (_, col) => {
                  const selected = api.isSelected(row, col);
                  return (
                    <td
                      key={cellKey(row, col)}
                      onClick={() => api.selectCell(row, col)}
                      onKeyDown={() => undefined}
                      style={{
                        border: "1px solid #cbd5e1",
                        padding: "4px 8px",
                        minWidth: 88,
                        cursor: "pointer",
                        background: selected ? "#ccfbf1" : "#fff",
                        font: "12px/1.4 ui-monospace, monospace",
                      }}
                    >
                      {api.getDisplay(row, col)}
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>

        <div style={{ marginTop: 12, display: "flex", gap: 8 }}>
          <button
            type="button"
            className="a1s-btn"
            onClick={() => api.setCell(r1, c1, `set at ${toA1(r1, c1)}`)}
          >
            setCell
          </button>
          <button
            type="button"
            className="a1s-btn"
            disabled={!api.canUndo}
            onClick={api.undo}
          >
            undo
          </button>
          <button
            type="button"
            className="a1s-btn"
            onClick={() => api.applyStyle({ bold: !api.activeStyle.bold })}
          >
            bold
          </button>
        </div>
      </div>
    );
  },
  parameters: {
    docs: {
      description: {
        story:
          "No `Sheet.Root`, no `Sheet.Grid` — a plain `<table>` driven by `useSpreadsheet()`. Formulas, undo, and formatting all work, because none of that lives in the components.",
      },
    },
  },
};

export const UseSheetInsideRoot: Story = {
  name: "useSheet — your part, our context",
  render: function UseSheetStory() {
    return (
      <Sheet.Root defaultWorkbook={budget()} height={380}>
        <Sheet.Toolbar />
        <Sheet.Grid />
        <Stats />
      </Sheet.Root>
    );
  },
  parameters: {
    docs: {
      description: {
        story:
          "`useSheet()` returns the same object the built-in primitives read. A component you write is not a second-class citizen, and it needs no props threaded to it. Outside `Sheet.Root` it throws `MissingProviderError` naming the fix.",
      },
    },
  },
};

/** Aggregates the current selection — the kind of part every app ends up wanting. */
function Stats() {
  const api = useSheet();
  const { r1, c1, r2, c2 } = api.bounds;
  let sum = 0;
  let numeric = 0;
  let filled = 0;
  for (let r = r1; r <= r2; r++) {
    for (let c = c1; c <= c2; c++) {
      const value = api.getValue(r, c);
      if (value !== "" && value !== undefined) filled++;
      if (typeof value === "number") {
        sum += value;
        numeric++;
      }
    }
  }
  return (
    <Panel title="Your own status bar">
      <Row label="filled">{filled}</Row>
      <Row label="numeric">{numeric}</Row>
      <Row label="sum">{numeric ? sum.toFixed(2) : "—"}</Row>
      <Row label="average">{numeric ? (sum / numeric).toFixed(2) : "—"}</Row>
    </Panel>
  );
}

export const Virtualization: Story = {
  name: "useRowWindow / useColWindow",
  render: function WindowStory() {
    return (
      <Sheet.Root defaultWorkbook={large(100_000, 200)} height={420}>
        <WindowReadout />
        <Sheet.Grid />
      </Sheet.Root>
    );
  },
  parameters: {
    docs: {
      description: {
        story:
          "Scroll and watch the numbers. The window moves; its size does not. `rowTop`, `rowAt`, and `colAt` are the coordinate mapping — binary searches over cumulative offset tables, so they hold when rows and columns differ in size.",
      },
    },
  },
};

function WindowReadout() {
  const api = useSheet();
  const { rowWindow, colWindow } = api;
  const first = rowWindow.windowRows[0]?.absRow ?? 0;
  const last = rowWindow.windowRows.at(-1)?.absRow ?? 0;
  return (
    <Panel title="What is actually in the DOM">
      <Row label="sheet">
        {api.sheet.numRows.toLocaleString()} rows × {api.sheet.numCols} cols
      </Row>
      <Row label="rows rendered">
        {rowWindow.windowRows.length} (rows {first}–{last})
      </Row>
      <Row label="columns rendered">{colWindow.windowCols.length}</Row>
      <Row label="cells rendered">
        {rowWindow.windowRows.length * colWindow.windowCols.length}
      </Row>
      <Row label="true scroll extent">
        {rowWindow.contentHeight.toLocaleString()}px ×{" "}
        {colWindow.totalWidth.toLocaleString()}px
      </Row>
    </Panel>
  );
}

export const FormulaRefEditing: Story = {
  name: "findRefSpans — formula mechanics, no React",
  render: function RefStory() {
    const [source, setSource] = useState("=SUM(B2:C4)+A1*LOG10(B2)");
    const spans = findRefSpans(source);
    const colors = ["#2563eb", "#dc2626", "#7c3aed", "#ea580c", "#0891b2"];
    return (
      <div style={{ padding: 12, font: "13px/1.5 system-ui" }}>
        <input
          className="a1s-input"
          aria-label="Formula source"
          value={source}
          onChange={(e) => setSource(e.target.value)}
          style={{ width: "100%", maxWidth: 520, fontFamily: "ui-monospace" }}
        />
        <Panel title={`findRefSpans() → ${spans.length} reference(s)`}>
          {spans.length === 0 && <Row label="none">no references found</Row>}
          {spans.map((span) => (
            <Row
              key={`${span.start}-${span.end}`}
              label={`chars ${span.start}–${span.end}`}
            >
              <code
                style={{
                  color: colors[span.group % colors.length],
                  fontWeight: 600,
                }}
              >
                {span.text}
              </code>{" "}
              → rows {span.range.r1}–{span.range.r2}, cols {span.range.c1}–
              {span.range.c2} (group {span.group})
            </Row>
          ))}
        </Panel>
        <p style={{ color: "#475569", maxWidth: 620 }}>
          Note what it does <em>not</em> match: <code>LOG10</code> is a function
          call, not a reference to cell LOG10, and text inside quotes is skipped.
          Repeated references share a group, which is how the grid colour-matches
          their outlines. This is a plain string function exported from{" "}
          <code>"a1sheet"</code> — no React involved, so your own editor gets the
          same behaviour.
        </p>
      </div>
    );
  },
};
