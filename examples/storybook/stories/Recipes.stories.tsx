/**
 * Problem statements and their solutions — the questions people actually arrive
 * with, each answered by a working sheet rather than a paragraph.
 */
import type { Meta, StoryObj } from "@storybook/react-vite";
import {
  cellKey,
  FUNCTIONS,
  parseRangeRef,
  type Range,
  registerFunction,
  toA1,
  toNumber,
} from "a1sheet";
import { Sheet, useSheet } from "a1sheet/react";
import { useMemo, useState } from "react";
import { budget, large } from "./fixtures.js";
import { Panel, Problem, Row } from "./ui.js";

const meta = {
  title: "Recipes/Recipes",
  tags: ["autodocs"],
  parameters: {
    docs: {
      description: {
        component:
          "Each recipe is a problem someone has actually hit, the shortest honest solution, and a sheet you can drive.",
      },
    },
  },
} satisfies Meta;

export default meta;
type Story = StoryObj<typeof meta>;

// --------------------------------------------------------------- read-only

function ReadOnlyGuard() {
  const api = useSheet();
  // Locking every cell is the supported way to make a sheet read-only: locked
  // cells refuse edits and clears but stay selectable and copyable, which is
  // what "read-only" should mean.
  useMemo(() => {
    api.updateSheet((s) => {
      for (let r = 0; r < 12; r++) {
        for (let c = 0; c < 6; c++) {
          const key = cellKey(r, c);
          s.styles[key] = { ...(s.styles[key] ?? {}), locked: true };
        }
      }
      return s;
    }, false);
  }, [api.updateSheet]);
  return null;
}

export const ReadOnly: Story = {
  name: "Make it read-only",
  render: () => (
    <div>
      <Problem
        problem="There is no readOnly prop, and adding one would be a boolean gating behaviour — the pattern this library refuses."
        solution="Lock the cells. Locked cells refuse edits and clears, and say so, while staying selectable and copyable."
      />
      <Sheet.Root defaultWorkbook={budget()} height={340}>
        <ReadOnlyGuard />
        <Sheet.Grid />
        <Sheet.StatusBar />
      </Sheet.Root>
    </div>
  ),
  parameters: {
    docs: {
      description: {
        story:
          "Try to type. If you want it visually read-only too, leave out `Sheet.Toolbar` — the parts you do not render do not exist.",
      },
    },
  },
};

// ------------------------------------------------------------ custom function

export const CustomFunction: Story = {
  name: "Add my own formula function",
  render: function CustomFnStory() {
    useMemo(() => {
      // Idempotent: the story can re-render without redefining.
      if (!FUNCTIONS.MARKUP) {
        registerFunction("MARKUP", (args) => {
          const base = toNumber(args[0]);
          const pct = toNumber(args[1]);
          return base * (1 + pct / 100);
        });
      }
    }, []);

    const wb = budget();
    const sheet = wb.sheets[0];
    if (sheet) {
      sheet.cells["9_0"] = "MARKUP demo";
      sheet.cells["9_1"] = "100";
      sheet.cells["9_2"] = "15";
      sheet.cells["9_3"] = "=MARKUP(B10,C10)";
      sheet.styles["9_0"] = { bold: true };
    }

    return (
      <div>
        <Problem
          problem="Our domain has a calculation the 30 built-in functions do not cover."
          solution="registerFunction(name, fn). It joins the registry and works in every sheet, including in files you write."
        />
        <Panel title="The whole implementation">
          <Row label="">
            <code style={{ fontSize: 11 }}>
              {`registerFunction("MARKUP", (args) => toNumber(args[0]) * (1 + toNumber(args[1]) / 100))`}
            </code>
          </Row>
          <Row label="D10">
            <code>=MARKUP(B10,C10)</code> → 115
          </Row>
        </Panel>
        <Sheet.Root defaultWorkbook={wb} height={340}>
          <Sheet.FormulaBar />
          <Sheet.Grid />
        </Sheet.Root>
      </div>
    );
  },
};

// -------------------------------------------------------------- named ranges

export const NamedRanges: Story = {
  name: "Use named ranges",
  render: function NamedStory() {
    const wb = budget();
    wb.namedRanges = { LineTotals: parseRangeRef("D2:D4") as Range };
    const sheet = wb.sheets[0];
    if (sheet) {
      sheet.cells["9_0"] = "Named";
      sheet.cells["9_3"] = "=SUM(LineTotals)";
      sheet.styles["9_0"] = { bold: true };
    }
    return (
      <div>
        <Problem
          problem="=SUM(D2:D4) breaks silently when someone inserts a row, and reads like nothing."
          solution="Define a name once on the workbook and use it in formulas."
        />
        <Panel title="workbook.namedRanges">
          <Row label="LineTotals">D2:D4</Row>
          <Row label="D10">
            <code>=SUM(LineTotals)</code>
          </Row>
        </Panel>
        <Sheet.Root defaultWorkbook={wb} height={340}>
          <Sheet.FormulaBar />
          <Sheet.Grid />
        </Sheet.Root>
      </div>
    );
  },
  parameters: {
    docs: {
      description: {
        story:
          "Names are workbook-level and resolve against the active sheet. `api.defineName(name, range)` adds one at runtime. There is no per-sheet scoping yet — that is in LIMITATIONS.",
      },
    },
  },
};

// ------------------------------------------------------------------- toolbar

function ValidationBanner() {
  const api = useSheet();
  const problems: string[] = [];
  for (let r = 1; r <= 4; r++) {
    const qty = api.getValue(r, 1);
    if (qty !== "" && typeof qty === "number" && qty > 10) {
      problems.push(`${toA1(r, 1)} is ${qty} — over the limit of 10`);
    }
  }
  if (problems.length === 0) return null;
  return (
    <div
      role="status"
      style={{
        font: "13px/1.5 system-ui",
        background: "#fef3c7",
        borderBottom: "1px solid #fcd34d",
        padding: "8px 12px",
      }}
    >
      {problems.join(" · ")}
    </div>
  );
}

export const Validation: Story = {
  name: "Validate as the user types",
  render: () => (
    <div>
      <Problem
        problem="There is no onBeforeEdit hook to reject a value, and rejecting keystrokes is hostile anyway."
        solution="Render your own component that reads the sheet through useSheet() and reports. It re-renders on every change, so validation is just derived state."
      />
      <Sheet.Root defaultWorkbook={budget()} height={340}>
        <ValidationBanner />
        <Sheet.Grid />
      </Sheet.Root>
    </div>
  ),
  parameters: {
    docs: {
      description: {
        story:
          "Set B2 to something over 10 and the banner appears. To *block* the write rather than report it, use controlled mode and reject in `onWorkbookChange` by not storing the new workbook.",
      },
    },
  },
};

// -------------------------------------------------------------------- search

export const FindAndSelect: Story = {
  name: "Search a large sheet",
  render: function SearchStory() {
    const [wb] = useState(() => large(20_000));
    const [needle, setNeedle] = useState("ROW-19999");
    return (
      <div>
        <Problem
          problem="There is no built-in find and replace, and a 20,000-row sheet is not something to scroll through."
          solution="cells is a plain object. Scan it yourself and call selectCell — a full pass over a million cells is about 30ms."
        />
        <Sheet.Root defaultWorkbook={wb} height={420}>
          <SearchBar needle={needle} onNeedle={setNeedle} />
          <Sheet.Grid />
          <Sheet.StatusBar />
        </Sheet.Root>
      </div>
    );
  },
  parameters: {
    docs: {
      description: {
        story:
          "Press Find and the sheet scrolls to the hit and selects it. Search is a genuine gap in the library rather than something hidden — but the data is plain and the API is public, so it is about fifteen lines.",
      },
    },
  },
};

function SearchBar({
  needle,
  onNeedle,
}: {
  needle: string;
  onNeedle: (value: string) => void;
}) {
  const api = useSheet();
  const [result, setResult] = useState("");

  const find = () => {
    const target = needle.toLowerCase();
    for (const [key, raw] of Object.entries(api.sheet.cells)) {
      if (!raw.toLowerCase().includes(target)) continue;
      const [row, col] = key.split("_").map(Number) as [number, number];
      api.selectCell(row, col);

      // Scroll the container itself, not `api.setScrollTop`. That setter only
      // tells virtualization where to draw; moving the element is what makes
      // its own onScroll fire and keeps the two in step. `rowTop` is the same
      // coordinate mapping the grid lays out with, so it lands exactly.
      const top = api.rowWindow.rowTop(row);
      const scroller = document.querySelector(".a1s-scroller");
      if (top !== null && scroller) scroller.scrollTop = top;

      setResult(`found at ${toA1(row, col)}`);
      return;
    }
    setResult("no match");
  };

  return (
    <Panel title="Find">
      <div style={{ display: "flex", gap: 8 }}>
        <input
          className="a1s-input"
          aria-label="Search term"
          value={needle}
          onChange={(e) => onNeedle(e.target.value)}
        />
        <button type="button" className="a1s-btn" onClick={find}>
          Find
        </button>
        <span style={{ alignSelf: "center", color: "#475569" }}>{result}</span>
      </div>
    </Panel>
  );
}
