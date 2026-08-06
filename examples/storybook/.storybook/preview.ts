import type { Preview } from "@storybook/react-vite";

const preview: Preview = {
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
