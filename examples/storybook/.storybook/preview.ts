import type { Preview } from "@storybook/react-vite";

const preview: Preview = {
  // Storybook's built-in Outline and Measure tools are off unless someone asks
  // for them. Outline paints `outline: 1px solid !important` over every
  // element, which on a grid means blue lines on all four sides of every cell —
  // it reads as a bug in the component rather than as a tool being on, and the
  // setting follows you between stories once enabled.
  initialGlobals: { outline: false, measureEnabled: false },
  parameters: {
    controls: { expanded: true, sort: "requiredFirst" },
    layout: "fullscreen",
    docs: { toc: true },
    options: {
      // Read in order, this is a tutorial: what it is, how to compose it, how
      // to control it, what it can do, then the hard parts.
      storySort: {
        order: [
          "Start here",
          ["Introduction", "Quick start", "Composition over configuration"],
          "Preset",
          "Composition",
          "State",
          "Hooks",
          "Features",
          "Data and files",
          "Scale",
          "Theming",
          "Recipes",
        ],
      },
    },
  },
};

export default preview;
