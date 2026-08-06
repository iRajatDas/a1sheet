/**
 * `<Spreadsheet />` — the preset, for when the default arrangement is fine.
 *
 * Every story here is one line of JSX plus a workbook. If you find yourself
 * wanting a prop that does not exist, that is the signal to compose the
 * primitives instead — see the Composition stories.
 */
import type { Meta, StoryObj } from "@storybook/react-vite";
import { Spreadsheet } from "a1sheet/react";
import * as fixtures from "./fixtures.js";

const meta = {
  title: "Preset/Spreadsheet",
  component: Spreadsheet,
  parameters: {
    docs: {
      description: {
        component:
          "A preset that composes the primitives in a default arrangement. It takes no layout props — no showToolbar, no showStatusBar — because an absent part is a child you do not render.",
      },
    },
  },
} satisfies Meta<typeof Spreadsheet>;

export default meta;
type Story = StoryObj<typeof meta>;

export const Basic: Story = {
  args: { defaultWorkbook: fixtures.budget(), height: 480 },
};

export const Formulas: Story = {
  name: "Formulas and error values",
  args: { defaultWorkbook: fixtures.formulas(), height: 480 },
  parameters: {
    docs: {
      description: {
        story:
          "Select an error cell and the status bar explains it. Start typing a formula and clicking the grid writes references instead of moving the selection.",
      },
    },
  },
};

export const FrozenAndMerged: Story = {
  name: "Frozen panes, merges, and fills",
  args: { defaultWorkbook: fixtures.layout(), height: 420 },
};

export const DarkTheme: Story = {
  args: {
    defaultWorkbook: fixtures.budget(),
    height: 480,
    theme: {
      accent: "#2dd4bf",
      border: "#1e293b",
      headerBorder: "#334155",
      buttonBorder: "#334155",
      headerBg: "#0f172a",
      headerText: "#94a3b8",
      cellBg: "#0b1220",
      cellText: "#e2e8f0",
      selectedBg: "rgba(45,212,191,0.16)",
      toolbarBg: "#0f172a",
    },
  },
  parameters: {
    docs: {
      description: {
        story:
          "Theming is a partial Theme object. There is no dark-mode boolean — a theme is values, and you supply the ones you want to change.",
      },
    },
  },
};

export const Branded: Story = {
  name: "Custom prefix and typography",
  args: {
    defaultWorkbook: fixtures.budget(),
    height: 420,
    classNamePrefix: "acme-",
    theme: {
      accent: "#7c3aed",
      fontSize: "14px",
      monoFontFamily: "ui-monospace, SFMono-Regular, monospace",
      refColors: ["#7c3aed", "#db2777", "#ea580c"],
    },
  },
  parameters: {
    docs: {
      description: {
        story:
          "classNamePrefix renames every injected class, so the CSS cannot collide with a host application's stylesheet.",
      },
    },
  },
};
