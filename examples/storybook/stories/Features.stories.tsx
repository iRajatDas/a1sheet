/**
 * Each feature, in isolation, with the interaction that exercises it scripted
 * as a play function — so the story demonstrates itself and fails loudly if the
 * behaviour regresses.
 */
import type { Meta, StoryObj } from "@storybook/react-vite";
import { explainErrorValue, FUNCTIONS, toA1 } from "a1sheet";
import { Sheet, useSheet } from "a1sheet/react";
import { expect, userEvent } from "storybook/test";
import { budget, formats, formulas, layout, protectedSheet } from "./fixtures.js";
import { Panel, Row } from "./ui.js";

const meta = {
  title: "Features/Editing",
  tags: ["autodocs"],
} satisfies Meta;

export default meta;
type Story = StoryObj<typeof meta>;

/**
 * Finds a cell by address rather than by position among the rendered nodes.
 * Both axes are virtualized, so the nth cell in the DOM is not row n / COLS —
 * `data-row`/`data-col` is the stable way in, for a test or for your own code.
 */
function cellAt(root: HTMLElement, row: number, col: number): HTMLElement {
  const el = root.querySelector(`.a1s-cell[data-row="${row}"][data-col="${col}"]`);
  if (!el)
    throw new Error(`no cell at row ${row}, col ${col} — is it in the window?`);
  return el as HTMLElement;
}

export const TypeToEdit: Story = {
  name: "Select a cell and type",
  render: () => (
    <Sheet.Root defaultWorkbook={budget()} height={360}>
      <Sheet.FormulaBar />
      <Sheet.Grid />
    </Sheet.Root>
  ),
  play: async ({ canvasElement }) => {
    await userEvent.click(cellAt(canvasElement, 2, 0));

    const textarea = canvasElement.querySelector("textarea") as HTMLTextAreaElement;
    await userEvent.type(textarea, "Sprocket");

    const editor = canvasElement.querySelector(
      ".a1s-cell input",
    ) as HTMLInputElement;
    await expect(editor).not.toBeNull();
    await expect(editor.value).toBe("Sprocket");
  },
  parameters: {
    docs: {
      description: {
        story:
          "Click a cell, then type — no double-click needed. The whole grid's keyboard and clipboard live on one hidden textarea, and a cell mousedown returns focus to it. Getting that wrong is what once made the grid silently dead to the keyboard after the first click.",
      },
    },
  },
};

export const KeyboardNavigation: Story = {
  render: () => (
    <Sheet.Root defaultWorkbook={budget()} height={360}>
      <ActiveCellReadout />
      <Sheet.Grid />
    </Sheet.Root>
  ),
  play: async ({ canvasElement }) => {
    await userEvent.click(cellAt(canvasElement, 1, 1));
    const textarea = canvasElement.querySelector("textarea") as HTMLTextAreaElement;
    await userEvent.keyboard("{ArrowDown}{ArrowDown}{ArrowRight}");
    await expect(textarea).toBe(document.activeElement as HTMLTextAreaElement);
  },
  parameters: {
    docs: {
      description: {
        story:
          "Arrows move, Shift+arrows extend, Tab and Enter commit and step, Ctrl/Cmd+arrow jumps. Try Shift+drag too: the anchor stays put and the readout above shows it.",
      },
    },
  },
};

function ActiveCellReadout() {
  const api = useSheet();
  return (
    <Panel title="Selection">
      <Row label="anchor (arrows step from here)">
        {toA1(api.active.row, api.active.col)}
      </Row>
      <Row label="range">
        {toA1(api.bounds.r1, api.bounds.c1)}:{toA1(api.bounds.r2, api.bounds.c2)}
      </Row>
    </Panel>
  );
}

export const Formulas: Story = {
  name: "Formula engine",
  render: () => (
    <div>
      <Panel title={`${Object.keys(FUNCTIONS).length} functions available`}>
        <Row label="registered">
          <code style={{ fontSize: 11 }}>
            {Object.keys(FUNCTIONS).sort().join(", ")}
          </code>
        </Row>
      </Panel>
      <Sheet.Root defaultWorkbook={formulas()} height={420}>
        <Sheet.FormulaBar />
        <Sheet.Grid />
        <Sheet.StatusBar />
      </Sheet.Root>
    </div>
  ),
  parameters: {
    docs: {
      description: {
        story:
          "Lazy, memoised evaluation with cycle detection. `$` anchoring and named ranges are supported. Registering your own function is one call to `registerFunction`.",
      },
    },
  },
};

export const ReferencePicking: Story = {
  name: "Picking references while typing",
  render: () => (
    <Sheet.Root defaultWorkbook={budget()} height={380}>
      <Sheet.FormulaBar />
      <Sheet.Grid />
      <Sheet.StatusBar />
    </Sheet.Root>
  ),
  play: async ({ canvasElement }) => {
    await userEvent.click(cellAt(canvasElement, 9, 3));
    const textarea = canvasElement.querySelector("textarea") as HTMLTextAreaElement;
    await userEvent.type(textarea, "=SUM(");
    // Mid-formula, a click writes a reference instead of moving the selection.
    await userEvent.click(cellAt(canvasElement, 1, 3));
    const editor = canvasElement.querySelector(
      ".a1s-cell input",
    ) as HTMLInputElement;
    await expect(editor.value).toBe("=SUM(D2");
  },
  parameters: {
    docs: {
      description: {
        story:
          "The play function typed `=SUM(` into D10 and clicked D2. Carry on from there: drag to grow that one reference into a range, and watch every reference get a colour-matched outline in the grid. A click *after* a finished operand — `=A1+2` with the caret at the end — means select again, as in Sheets.",
      },
    },
  },
};

export const ErrorValues: Story = {
  name: "Errors explain themselves",
  render: () => (
    <div>
      <Panel title="explainErrorValue()">
        {["#DIV/0!", "#CYCLE!", "#NAME?", "#VALUE!", "#REF!"].map((sentinel) => (
          <Row key={sentinel} label={sentinel}>
            {explainErrorValue(sentinel)}
          </Row>
        ))}
      </Panel>
      <Sheet.Root defaultWorkbook={formulas()} height={380}>
        <Sheet.Grid />
        <Sheet.StatusBar />
      </Sheet.Root>
    </div>
  ),
  parameters: {
    docs: {
      description: {
        story:
          "Select A9, A10, or A11 and the status bar says what to do about it. A bad formula degrades to an error value in that one cell and never breaks the sheet.",
      },
    },
  },
};

export const Formatting: Story = {
  name: "Number formats and styling",
  render: () => (
    <Sheet.Root defaultWorkbook={formats()} height={340}>
      <Sheet.Toolbar />
      <Sheet.Grid />
    </Sheet.Root>
  ),
  parameters: {
    docs: {
      description: {
        story:
          "Column B is the raw value, column C the same value with a format applied. Formats are style, not data — the underlying text is untouched, so a formatted cell still computes.",
      },
    },
  },
};

export const StructureAndMerges: Story = {
  name: "Rows, columns, merges, freeze",
  render: () => (
    <Sheet.Root defaultWorkbook={layout()} height={420}>
      <Sheet.Toolbar />
      <Sheet.Grid />
      <Sheet.ContextMenu />
      <Sheet.ColumnMenu />
    </Sheet.Root>
  ),
  parameters: {
    docs: {
      description: {
        story:
          "Right-click for insert, delete, merge, and hide. Metadata moves with the structure: insert a row above a resized one and the height travels with its row rather than staying on the index.",
      },
    },
  },
};

export const SortAndFilter: Story = {
  render: () => (
    <Sheet.Root defaultWorkbook={budget()} height={400}>
      <Sheet.Toolbar />
      <Sheet.Grid />
      <Sheet.ColumnMenu />
      <Sheet.StatusBar />
    </Sheet.Root>
  ),
  parameters: {
    docs: {
      description: {
        story:
          "Use the ▾ on a column header. **Sort rewrites the data** — it is a real reordering, undoable through history but not by clearing anything. **Filter is a view** — clearing it always restores every row, because it never touches `cells`.",
      },
    },
  },
};

export const FillHandle: Story = {
  name: "Fill handle and series",
  render: () => (
    <Sheet.Root defaultWorkbook={budget()} height={380}>
      <Sheet.Grid />
      <Sheet.StatusBar />
    </Sheet.Root>
  ),
  parameters: {
    docs: {
      description: {
        story:
          "Select B2:B4 and drag the small square at the bottom-right corner. Numbers extrapolate as a linear series; formulas shift their relative references per cell, and `$`-anchored parts stay put.",
      },
    },
  },
};

export const LockedAndHidden: Story = {
  name: "Locked cells and hidden rows",
  render: () => (
    <Sheet.Root defaultWorkbook={protectedSheet()} height={320}>
      <Sheet.Toolbar />
      <Sheet.Grid />
      <Sheet.StatusBar />
      <Sheet.ContextMenu />
    </Sheet.Root>
  ),
  parameters: {
    docs: {
      description: {
        story:
          "B2 is locked: editing and clearing refuse and the status bar says why, but it stays selectable and copyable by design. Row 4 is hidden — the row numbers skip it, and it is genuinely absent from the layout rather than zero-height.",
      },
    },
  },
};

export const UndoRedo: Story = {
  render: () => (
    <Sheet.Root defaultWorkbook={budget()} height={360}>
      <Sheet.Toolbar />
      <Sheet.Grid />
      <HistoryReadout />
    </Sheet.Root>
  ),
  parameters: {
    docs: {
      description: {
        story:
          "Fifty levels, Ctrl+Z and Ctrl+Y. A resize drag is deliberately excluded from history — it would otherwise push one entry per mousemove and bury everything else.",
      },
    },
  },
};

function HistoryReadout() {
  const api = useSheet();
  return (
    <Panel title="History">
      <Row label="can undo">{api.canUndo ? "yes" : "no"}</Row>
      <Row label="can redo">{api.canRedo ? "yes" : "no"}</Row>
    </Panel>
  );
}
