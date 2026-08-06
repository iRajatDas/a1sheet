import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import type { StorybookConfig } from "@storybook/react-vite";

const here = dirname(fileURLToPath(import.meta.url));
const src = resolve(here, "..", "..", "..", "packages", "a1sheet", "src");

const config: StorybookConfig = {
  stories: ["../stories/**/*.mdx", "../stories/**/*.stories.tsx"],
  addons: ["@storybook/addon-docs"],
  framework: { name: "@storybook/react-vite", options: {} },
  typescript: {
    // `react-docgen`, not `react-docgen-typescript`: the latter reaches into
    // the TypeScript compiler's internals and throws against the TypeScript 7
    // this repo builds with. This one parses the source, which is enough for
    // prop names and JSDoc; the precise types and prose come from the explicit
    // `argTypes` in the stories.
    reactDocgen: "react-docgen",
  },
  viteFinal(cfg) {
    // Alias to source, not to dist: a story is the fastest feedback loop in the
    // repo and should not need a build step between an edit and the browser.
    cfg.resolve ??= {};
    cfg.resolve.alias = {
      ...cfg.resolve.alias,
      "a1sheet/react": join(src, "react", "index.ts"),
      a1sheet: join(src, "index.ts"),
    };
    return cfg;
  },
};

export default config;
