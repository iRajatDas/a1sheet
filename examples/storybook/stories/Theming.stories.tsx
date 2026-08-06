/**
 * Theming: values, not flags. Every story here is the same grid with a
 * different `Partial<Theme>`.
 */
import type { Meta, StoryObj } from "@storybook/react-vite";
import {
  buildCss,
  defaultTheme,
  resolveTheme,
  Sheet,
  type Theme,
} from "a1sheet/react";
import { budget, DARK } from "./fixtures.js";
import { Panel, Problem, Row } from "./ui.js";

const meta = {
  title: "Theming/Themes",
  tags: ["autodocs"],
  parameters: {
    docs: {
      description: {
        component: [
          "Styling is inline plus one injected `<style>` tag. There is **no CSS file to import**, so dropping this into an app needs no loader, no build-config change, and no bundler plugin.",
          "",
          "A theme is a `Partial<Theme>`: pass the keys you want changed and the rest fall back to `defaultTheme`. There is no `darkMode` boolean — dark is a set of values, and a boolean would only cover the one dark palette we happened to pick.",
        ].join("\n"),
      },
    },
  },
} satisfies Meta;

export default meta;
type Story = StoryObj<typeof meta>;

const sheetWith = (theme: Partial<Theme>, prefix?: string) => (
  <Sheet.Root
    defaultWorkbook={budget()}
    height={360}
    theme={theme}
    {...(prefix ? { classNamePrefix: prefix } : {})}
  >
    <Sheet.Toolbar />
    <Sheet.FormulaBar />
    <Sheet.Grid />
    <Sheet.StatusBar />
  </Sheet.Root>
);

export const Default: Story = {
  render: () => sheetWith({}),
  parameters: {
    docs: { description: { story: "`defaultTheme`, untouched." } },
  },
};

export const Dark: Story = {
  render: () => sheetWith(DARK),
  parameters: {
    docs: {
      description: {
        story:
          "Thirteen values. Note `selectedBg` is deliberately translucent — the selection tint is painted as an `::after` overlay so it composites over a cell's own fill instead of replacing it. The three `scrollbar*` keys matter more than they look: the grid draws its own scrollbars in channels beside the cells rather than letting the platform float them over the content, so those channels are a real surface and stay light unless you colour them.",
      },
    },
  },
};

export const HighContrast: Story = {
  render: () =>
    sheetWith({
      accent: "#000000",
      border: "#000000",
      headerBorder: "#000000",
      buttonBorder: "#000000",
      headerBg: "#ffffff",
      headerText: "#000000",
      cellBg: "#ffffff",
      cellText: "#000000",
      selectedBg: "rgba(0,0,0,0.18)",
      toolbarBg: "#ffffff",
      scrollbarTrack: "#ffffff",
      scrollbarThumb: "#000000",
      scrollbarThumbHover: "#000000",
      fontSize: "15px",
    }),
};

export const Branded: Story = {
  name: "A different product's palette",
  render: () =>
    sheetWith(
      {
        accent: "#7c3aed",
        border: "#ede9fe",
        headerBorder: "#ddd6fe",
        buttonBorder: "#ddd6fe",
        headerBg: "#faf5ff",
        headerText: "#6d28d9",
        selectedBg: "rgba(124,58,237,0.12)",
        toolbarBg: "#faf5ff",
        fontFamily: "Georgia, serif",
        fontSize: "14px",
        refColors: ["#7c3aed", "#db2777", "#ea580c"],
      },
      "acme-",
    ),
  parameters: {
    docs: {
      description: {
        story:
          "Also renamed every class to `acme-*` via `classNamePrefix`, so the injected CSS cannot collide with the host app's stylesheet. `refColors` is the palette used to outline references in a formula being edited.",
      },
    },
  },
};

export const TheThemeItself: Story = {
  name: "Every key",
  render: () => (
    <div style={{ padding: 12 }}>
      <Problem
        problem="A component that ships a CSS file forces every consumer to configure a loader for it."
        solution="One <style> tag built from a plain object, with a prefix you control."
      />
      <Panel title="defaultTheme">
        {Object.entries(defaultTheme).map(([key, value]) => (
          <Row key={key} label={key}>
            <code>{Array.isArray(value) ? value.join(", ") : String(value)}</code>
          </Row>
        ))}
      </Panel>
      <Panel title="Escape hatches">
        <Row label="resolveTheme(partial)">
          merges your keys over the defaults — the same call the component makes
        </Row>
        <Row label="buildCss(prefix, theme)">
          returns the stylesheet as a string, if you would rather inject it yourself
          or ship it statically
        </Row>
        <Row label="stylesheet size">
          {buildCss("a1s-", resolveTheme({})).length.toLocaleString()} characters
        </Row>
      </Panel>
    </div>
  ),
};
