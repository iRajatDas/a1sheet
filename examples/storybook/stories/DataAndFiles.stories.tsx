/**
 * Reading and writing real files, with the parts that only matter at size:
 * progress, cancellation, typed errors, and the CSV injection guard.
 */
import type { Meta, StoryObj } from "@storybook/react-vite";
import {
  cellsToCSV,
  createEvaluator,
  isA1SheetError,
  readWorkbookFile,
  writeXlsx,
} from "a1sheet";
import { Sheet } from "a1sheet/react";
import { useRef, useState } from "react";
import { budget } from "./fixtures.js";
import { Panel, Problem, Row } from "./ui.js";

const meta = {
  title: "Data and files/Import and export",
  tags: ["autodocs"],
  parameters: {
    docs: {
      description: {
        component: [
          "XLSX and CSV, read and written by hand — a DEFLATE decoder, a ZIP reader, an OOXML parser, and a CSV parser, all in this package, because `dependencies` is `{}`.",
          "",
          "`.xls` and `.xlsb` are rejected with an actionable message rather than failing somewhere deep in the parser. Both are binary formats and out of scope.",
        ].join("\n"),
      },
    },
  },
} satisfies Meta;

export default meta;
type Story = StoryObj<typeof meta>;

export const TheEasyWay: Story = {
  name: "Sheet.FileMenu",
  render: () => (
    <Sheet.Root defaultWorkbook={budget()} height={420}>
      <Sheet.Toolbar>
        <Sheet.FileMenu />
      </Sheet.Toolbar>
      <Sheet.Grid />
      <Sheet.Tabs />
      <Sheet.StatusBar />
    </Sheet.Root>
  ),
  parameters: {
    docs: {
      description: {
        story:
          "Drop a real `.xlsx` or `.csv` in. Progress appears in the status bar; choosing a second file cancels the first. Not rendering `Sheet.FileMenu` keeps the XLSX writer and ZIP reader out of your bundle entirely.",
      },
    },
  },
};

export const ProgressAndCancellation: Story = {
  name: "Progress, cancellation, typed errors",
  render: function ImportStory() {
    const [log, setLog] = useState<string[]>([]);
    const [ratio, setRatio] = useState(0);
    const controllerRef = useRef<AbortController | null>(null);
    const say = (line: string) => setLog((prev) => [...prev.slice(-6), line]);

    async function onPick(file: File) {
      controllerRef.current?.abort();
      const controller = new AbortController();
      controllerRef.current = controller;
      setLog([]);
      setRatio(0);

      try {
        const result = await readWorkbookFile(file, {
          signal: controller.signal,
          onProgress: ({ phase, ratio: r, detail }) => {
            setRatio(r);
            say(`${phase} ${Math.round(r * 100)}% — ${detail}`);
          },
        });
        say(
          `done: ${result.format}, ${result.sheets.length} sheet(s), ` +
            `${result.sheets.reduce((n, s) => n + Object.keys(s.cells).length, 0)} cells`,
        );
      } catch (err) {
        // Branch on the code. Never on the message — codes are stable, messages
        // are free to improve.
        if (!isA1SheetError(err)) throw err;
        switch (err.code) {
          case "ABORTED":
            return say("aborted — nothing was applied");
          case "UNSUPPORTED_FORMAT":
            return say(`unsupported: ${err.message}`);
          case "NOT_A_ZIP":
          case "MALFORMED_FILE":
            return say("that file is not a spreadsheet");
          default:
            throw err;
        }
      }
    }

    return (
      <div style={{ padding: 12 }}>
        <Problem
          problem="A 38 MB workbook parsed on the main thread freezes the tab, with no progress and no way out."
          solution="readWorkbookFile takes an AbortSignal and an onProgress callback, and yields to the browser between chunks."
        />
        <Panel title="Pick a file — a big one, ideally">
          <input
            type="file"
            accept=".csv,.xlsx,.xlsm,.xls,.xlsb"
            aria-label="File to read"
            onChange={(e) => {
              const file = e.target.files?.[0];
              if (file) void onPick(file);
              e.target.value = "";
            }}
          />
          <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
            <progress value={ratio} max={1} style={{ width: 260 }} />
            <button
              type="button"
              className="a1s-btn"
              onClick={() => controllerRef.current?.abort()}
            >
              Cancel
            </button>
          </div>
          {log.map((line) => (
            <Row key={line} label="">
              <code style={{ fontSize: 11 }}>{line}</code>
            </Row>
          ))}
        </Panel>
        <p style={{ font: "13px/1.6 system-ui", color: "#475569", maxWidth: 640 }}>
          <code>onProgress</code> is rate-limited to about once per frame and once
          per 1%, so you can call <code>setState</code> from it directly.{" "}
          <code>ratio</code> never decreases and a successful read always ends at
          exactly 1. Aborting rejects with <code>AbortedError</code> and applies
          nothing — the workbook is built only after the whole file has parsed, so a
          cancelled import cannot half-replace what you had. Try a{" "}
          <code>.xlsb</code> to see the typed rejection.
        </p>
      </div>
    );
  },
};

export const Writing: Story = {
  name: "Writing files, and the CSV injection guard",
  render: function WriteStory() {
    const wb = budget();
    const sheet = wb.sheets[0];
    // A value a hostile source might put in a cell.
    const hostile = { ...(sheet?.cells ?? {}), "9_0": "=cmd|'/c calc'!A1" };
    const csv = cellsToCSV(hostile, createEvaluator(hostile, {}));
    const bytes = writeXlsx([
      { name: "Budget", cells: hostile, styles: sheet?.styles ?? {}, merges: [] },
    ]);

    return (
      <div style={{ padding: 12 }}>
        <Problem
          problem="A cell reading =cmd|'/c calc'!A1 is a formula the moment Excel opens the CSV. That is a real remote-code-execution vector, not a theoretical one."
          solution="Export neutralises any value starting with = + - @, tab, or CR. It is on by default and not configurable."
        />
        <Panel title="cellsToCSV output">
          <Row label="cell A10 contains">
            <code>=cmd|'/c calc'!A1</code>
          </Row>
          <Row label="last CSV line">
            <code>{csv.trim().split("\n").at(-1)}</code>
          </Row>
          <Row label="writeXlsx produced">
            {bytes.length.toLocaleString()} bytes
          </Row>
        </Panel>
        <p style={{ font: "13px/1.6 system-ui", color: "#475569", maxWidth: 640 }}>
          Download filenames are sanitised too — no separators, no traversal.
        </p>
      </div>
    );
  },
};

export const RoundTrip: Story = {
  name: "What survives a round trip",
  render: () => (
    <div style={{ padding: 12 }}>
      <Panel title="XLSX">
        <Row label="survives">
          values, formulas, shared strings, bold/italic/underline, text and fill
          colour, alignment, number formats, merges, column widths, row heights
        </Row>
        <Row label="does not">
          charts, pivot tables, conditional formatting, images, comments, defined
          styles beyond the above
        </Row>
      </Panel>
      <Panel title="CSV">
        <Row label="survives">displayed values only</Row>
        <Row label="does not">
          formulas, styles, merges, sizing — a CSV has nowhere to put them
        </Row>
      </Panel>
      <p style={{ font: "13px/1.6 system-ui", color: "#475569", maxWidth: 640 }}>
        Column widths are stored in multiples of the widest digit of the normal
        font, assumed to be Calibri 11, and row heights in points — so both are
        accurate to about a pixel rather than exactly. Only entries a file marks{" "}
        <code>customWidth</code>/<code>customHeight</code> are read, so an
        Excel-laid-out sheet does not arrive with every column pinned to a size
        nobody chose.
      </p>
    </div>
  ),
};
