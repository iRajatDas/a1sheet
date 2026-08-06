/**
 * `<Spreadsheet />` — the preset, for when the default arrangement is fine.
 */
import type { Meta, StoryObj } from "@storybook/react-vite";
import { Spreadsheet } from "a1sheet/react";
import { budget, DARK, formulas, layout } from "./fixtures.js";

const meta = {
  title: "Preset/Spreadsheet",
  component: Spreadsheet,
  tags: ["autodocs"],
  argTypes: {
    defaultWorkbook: {
      description:
        "Uncontrolled starting workbook. The component owns it from then on.",
      table: { category: "State", type: { summary: "Workbook" } },
      control: false,
    },
    workbook: {
      description:
        "Controlled workbook. Pair with onWorkbookChange. Never mix with defaultWorkbook — the component does not silently switch modes.",
      table: { category: "State", type: { summary: "Workbook" } },
      control: false,
    },
    onWorkbookChange: {
      description: "Fires with the next workbook after every change.",
      table: { category: "State", type: { summary: "(wb: Workbook) => void" } },
    },
    theme: {
      description:
        "Partial theme. Only the keys you pass are overridden. There is no dark-mode boolean — a theme is values.",
      table: { category: "Appearance", type: { summary: "Partial<Theme>" } },
      control: "object",
    },
    classNamePrefix: {
      description:
        'Prefix for every injected class name. Change it and the CSS cannot collide with your app. Defaults to "a1s-".',
      table: {
        category: "Appearance",
        defaultValue: { summary: '"a1s-"' },
        type: { summary: "string" },
      },
      control: "text",
    },
    height: {
      description: "Height of the container. Anything CSS accepts.",
      table: { category: "Appearance", type: { summary: "string | number" } },
      control: "text",
    },
    className: {
      description: "Applied to the outermost element.",
      table: { category: "Appearance" },
      control: "text",
    },
    style: {
      description: "Applied to the outermost element.",
      table: { category: "Appearance" },
      control: false,
    },
    children: { table: { disable: true } },
  },
  parameters: {
    docs: {
      description: {
        component: [
          "A preset that composes the primitives in a default arrangement.",
          "",
          "It takes **no layout props** — no `showToolbar`, no `showStatusBar`, no `toolbarPosition`. Every prop below is about the workbook or the skin, never about which parts exist. A part you do not want is a child you do not render, which means composing the primitives yourself.",
          "",
          "`<Spreadsheet />` and `<Sheet.Root>` take exactly the same props, because the preset is nothing but `Sheet.Root` with a fixed set of children.",
        ].join("\n"),
      },
    },
  },
} satisfies Meta<typeof Spreadsheet>;

export default meta;
type Story = StoryObj<typeof meta>;

export const Basic: Story = {
  args: { defaultWorkbook: budget(), height: 480 },
  parameters: {
    docs: {
      description: {
        story:
          "Type to edit. Enter and Tab commit and move, Escape cancels, F2 edits in place, Delete clears. Arrows navigate, Shift+arrows extend, Ctrl/Cmd+click adds a range. Ctrl+B/I/U format. Ctrl+Z and Ctrl+Y undo and redo, fifty deep.",
      },
    },
  },
};

export const Formulas: Story = {
  name: "Formulas and error values",
  args: { defaultWorkbook: formulas(), height: 520 },
  parameters: {
    docs: {
      description: {
        story:
          "Select an error cell: the status bar explains what to do about it rather than only showing the sentinel. Circular references report themselves in every cell of the cycle, not just the one that closed it.",
      },
    },
  },
};

export const FrozenAndMerged: Story = {
  name: "Frozen panes, merges, and fills",
  args: { defaultWorkbook: layout(), height: 420 },
  parameters: {
    docs: {
      description: {
        story:
          "Two frozen rows and one frozen column, a merge across the title, per-cell fills, and a pre-set row height and column width. Scroll in both directions — freeze panes are `position: sticky` inside a single scroll container, not four synced ones.",
      },
    },
  },
};

export const DarkTheme: Story = {
  args: { defaultWorkbook: budget(), height: 480, theme: DARK },
  parameters: {
    docs: {
      description: {
        story:
          "Theming is a partial `Theme` object. Edit it in the Controls panel below and the grid updates live.",
      },
    },
  },
};

export const Empty: Story = {
  name: "No workbook at all",
  args: { height: 420 },
  parameters: {
    docs: {
      description: {
        story:
          "Every prop is optional. With no workbook it creates an empty one, which is the fastest way to check the component mounts in your app.",
      },
    },
  },
};
