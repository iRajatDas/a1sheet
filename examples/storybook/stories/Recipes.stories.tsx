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
import {
  LabButton,
  LabField,
  LabMeta,
  LabRow,
  LabSection,
  LabStatusBar,
  LabToolbar,
  scrollHitIntoView,
} from "./labChrome.js";
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
  name: "Find and replace",
  render: function SearchStory() {
    const [wb] = useState(() => large(20_000));
    const [needle, setNeedle] = useState("ROW-19999");
    const [replacement, setReplacement] = useState("FOUND");
    return (
      <div>
        <Problem
          problem="You need find-next and replace-all on a large sheet without shipping a dialog."
          solution="useSheet().findNext / replaceAll — headless, with match-case and entire-cell options. Scroll with rowWindow.rowTop on the scroller."
        />
        <Sheet.Root defaultWorkbook={wb} height={420}>
          <FindReplaceBar
            needle={needle}
            onNeedle={setNeedle}
            replacement={replacement}
            onReplacement={setReplacement}
          />
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
          "Find next walks then wraps. Replace all rewrites raw cells (including formulas) and sets status to `Replaced N occurrence(s).`",
      },
    },
  },
};

function FindReplaceBar({
  needle,
  onNeedle,
  replacement,
  onReplacement,
}: {
  needle: string;
  onNeedle: (value: string) => void;
  replacement: string;
  onReplacement: (value: string) => void;
}) {
  const api = useSheet();
  const [result, setResult] = useState("");

  const find = () => {
    const hit = api.findNext({ find: needle, after: api.active });
    if (!hit) {
      setResult("no match");
      return;
    }
    api.selectCell(hit.row, hit.col);
    scrollHitIntoView(api, hit.row);
    setResult(`found at ${toA1(hit.row, hit.col)}`);
  };

  const replace = () => {
    const n = api.replaceAll({ find: needle, replace: replacement });
    setResult(`replaced ${n}`);
  };

  return (
    <LabToolbar>
      <LabSection label="Find & replace" hint="20,000-row sheet">
        <LabRow>
          <LabField
            label="Find"
            value={needle}
            onChange={onNeedle}
            width={200}
          />
          <LabField
            label="Replace with"
            value={replacement}
            onChange={onReplacement}
            width={140}
          />
          <div style={{ display: "flex", gap: 8, paddingTop: 16 }}>
            <LabButton kind="primary" onClick={find}>
              Find next
            </LabButton>
            <LabButton onClick={replace}>Replace all</LabButton>
          </div>
        </LabRow>
      </LabSection>
      <LabStatusBar />
      <LabMeta
        items={[
          { label: "Last", value: result || "—" },
          {
            label: "Hits",
            value: String(api.findAll({ find: needle }).length),
          },
        ]}
      />
    </LabToolbar>
  );
}
