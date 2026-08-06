/**
 * Controlled and uncontrolled state, and the imperative handle.
 *
 * Each story shows the state living *outside* the sheet, so you can watch it
 * move rather than take the claim on trust.
 */
import type { Meta, StoryObj } from "@storybook/react-vite";
import { cellKey, toA1, type Workbook } from "a1sheet";
import { Sheet, type SheetRootHandle, useSheet } from "a1sheet/react";
import { useRef, useState } from "react";
import { budget } from "./fixtures.js";
import { Panel, Row } from "./ui.js";

const meta = {
  title: "State/Controlled and uncontrolled",
  tags: ["autodocs"],
  parameters: {
    docs: {
      description: {
        component: [
          "Two modes, and the component never silently switches between them:",
          "",
          "| Mode | Props | Who owns the workbook |",
          "|---|---|---|",
          "| Uncontrolled | `defaultWorkbook` | the component |",
          "| Controlled | `workbook` + `onWorkbookChange` | you |",
          "",
          "Selection, editing, and scroll position are always internal — they are view state, not document state. Reach them through `useSheet()` or the imperative handle.",
        ].join("\n"),
      },
    },
  },
} satisfies Meta;

export default meta;
type Story = StoryObj<typeof meta>;

export const Uncontrolled: Story = {
  render: () => (
    <Sheet.Root defaultWorkbook={budget()} height={360}>
      <Sheet.Toolbar />
      <Sheet.Grid />
    </Sheet.Root>
  ),
  parameters: {
    docs: {
      description: {
        story:
          "`defaultWorkbook` is read once. Changing it afterwards does nothing — that is what makes it uncontrolled, and it is why the prop is named `default`.",
      },
    },
  },
};

export const Controlled: Story = {
  render: function ControlledStory() {
    const [wb, setWb] = useState<Workbook>(budget);
    const [changes, setChanges] = useState(0);
    const sheet = wb.sheets[wb.activeSheetIndex];

    return (
      <div>
        <Panel title="Your state">
          <Row label="onWorkbookChange calls">{changes}</Row>
          <Row label="active sheet">{sheet?.name}</Row>
          <Row label="filled cells">
            {sheet ? Object.keys(sheet.cells).length : 0}
          </Row>
          <Row label="D2">
            <code>{sheet?.cells[cellKey(1, 3)] ?? "—"}</code>
          </Row>
          <button
            type="button"
            className="a1s-btn"
            onClick={() =>
              setWb((prev) => {
                // Ordinary immutable update — the workbook is plain data.
                const next = structuredClone(prev);
                const s = next.sheets[next.activeSheetIndex];
                if (s)
                  s.cells[cellKey(1, 1)] = String(Math.ceil(Math.random() * 9));
                return next;
              })
            }
          >
            Set B2 from outside
          </button>
        </Panel>
        <Sheet.Root
          workbook={wb}
          onWorkbookChange={(next) => {
            setWb(next);
            setChanges((n) => n + 1);
          }}
          height={340}
        >
          <Sheet.Toolbar />
          <Sheet.Grid />
        </Sheet.Root>
      </div>
    );
  },
  parameters: {
    docs: {
      description: {
        story:
          "The workbook lives in the story's `useState`. Edit a cell and the counter moves; press the button and the grid follows. Writes go both ways because the workbook is plain serialisable data, not an opaque instance.",
      },
    },
  },
};

export const PersistToStorage: Story = {
  name: "Persisting every change",
  render: function PersistStory() {
    const [saved, setSaved] = useState<string>("nothing saved yet");
    return (
      <div>
        <Panel title="What you would send to your backend">
          <Row label="last write">{saved}</Row>
        </Panel>
        <Sheet.Root
          defaultWorkbook={budget()}
          onWorkbookChange={(wb) => {
            // In a real app: debounce, then POST. Sets do not survive JSON, so
            // hiddenRows and filters need converting if you serialise this way.
            const sheet = wb.sheets[0];
            setSaved(
              `${Object.keys(sheet?.cells ?? {}).length} cells at ${new Date().toLocaleTimeString()}`,
            );
          }}
          height={340}
        >
          <Sheet.Toolbar />
          <Sheet.Grid />
        </Sheet.Root>
      </div>
    );
  },
  parameters: {
    docs: {
      description: {
        story:
          "`onWorkbookChange` fires in uncontrolled mode too, so you can persist without taking ownership of the state. Note that `hiddenRows` and `filters` are `Set`s and do not survive `JSON.stringify` — convert them if you serialise.",
      },
    },
  },
};

/** Reads the live selection out of context, to prove it is view state. */
function SelectionReadout() {
  const api = useSheet();
  const { r1, c1, r2, c2 } = api.bounds;
  const cells = (r2 - r1 + 1) * (c2 - c1 + 1);
  return (
    <Panel title="View state (never in the workbook)">
      <Row label="range">
        {toA1(r1, c1)}:{toA1(r2, c2)} — {cells} cell{cells === 1 ? "" : "s"}
      </Row>
      <Row label="anchor">
        {toA1(api.active.row, api.active.col)} — arrows step from here
      </Row>
      <Row label="value">
        <code>{api.getDisplay(api.active.row, api.active.col) || "empty"}</code>
      </Row>
      <Row label="editing">{api.editing ? "yes" : "no"}</Row>
      <Row label="undo / redo">
        {api.canUndo ? "can undo" : "—"} / {api.canRedo ? "can redo" : "—"}
      </Row>
    </Panel>
  );
}

export const ViewState: Story = {
  name: "Selection is view state",
  render: () => (
    <Sheet.Root defaultWorkbook={budget()} height={360}>
      <SelectionReadout />
      <Sheet.Grid />
    </Sheet.Root>
  ),
  parameters: {
    docs: {
      description: {
        story:
          "Drag a range. The active cell stays at the **anchor** where the drag started, as Excel and Sheets do, rather than following the mouse — so you can always see where typing will land.",
      },
    },
  },
};

export const ImperativeHandle: Story = {
  name: "Driving it from outside",
  render: function HandleStory() {
    const ref = useRef<SheetRootHandle>(null);
    return (
      <div>
        <Panel title="ref.current">
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            <button
              type="button"
              className="a1s-btn"
              onClick={() => ref.current?.api.setCell(0, 0, "written from a ref")}
            >
              api.setCell(0, 0, …)
            </button>
            <button
              type="button"
              className="a1s-btn"
              onClick={() => ref.current?.api.selectCell(3, 2)}
            >
              api.selectCell(3, 2)
            </button>
            <button
              type="button"
              className="a1s-btn"
              onClick={() => ref.current?.api.undo()}
            >
              api.undo()
            </button>
            <button
              type="button"
              className="a1s-btn"
              onClick={() => ref.current?.focus()}
            >
              focus()
            </button>
          </div>
        </Panel>
        <Sheet.Root ref={ref} defaultWorkbook={budget()} height={340}>
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
          "`SheetRootHandle` exposes `focus()` and the live `api`. Use it when the sheet has to react to something outside React's tree — a keyboard shortcut owned by your shell, or a websocket push.",
      },
    },
  },
};
