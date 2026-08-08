/**
 * Every story actually renders, and every play function actually passes.
 *
 * A Storybook that builds is not a Storybook that works: the build only proves
 * the modules parse. These tests import each story module, render each story,
 * and run its `play` where it has one — so a broken demo fails here rather than
 * in front of whoever opened the docs to learn what the library does.
 *
 * This runs under happy-dom rather than a browser, so it catches render errors,
 * bad fixtures, and broken interactions, but not layout or styling.
 */
import { describe, expect, test } from "bun:test";
import { cleanup, render } from "@testing-library/react";
import { createElement, type ReactElement } from "react";
import * as Composition from "./Composition.stories.js";
import * as DataAndFiles from "./DataAndFiles.stories.js";
import * as Features from "./Features.stories.js";
import * as FeaturesRecent from "./FeaturesRecent.stories.js";
import * as Hooks from "./Hooks.stories.js";
import * as Preset from "./Preset.stories.js";
import * as Recipes from "./Recipes.stories.js";
import * as Scale from "./Scale.stories.js";
import * as Showcase from "./Showcase.stories.js";
import * as State from "./State.stories.js";
import * as Theming from "./Theming.stories.js";

/** The shape a story exposes, narrowed to what this file drives. */
interface Story {
  name?: string;
  render?: (args: Record<string, unknown>) => ReactElement;
  args?: Record<string, unknown>;
  play?: (context: { canvasElement: HTMLElement }) => Promise<void> | void;
}

interface StoryModule {
  default: { title: string; component?: unknown };
  [exportName: string]: unknown;
}

const MODULES: StoryModule[] = [
  Showcase as unknown as StoryModule,
  Preset as unknown as StoryModule,
  Composition as unknown as StoryModule,
  State as unknown as StoryModule,
  Hooks as unknown as StoryModule,
  Features as unknown as StoryModule,
  FeaturesRecent as unknown as StoryModule,
  DataAndFiles as unknown as StoryModule,
  Scale as unknown as StoryModule,
  Theming as unknown as StoryModule,
  Recipes as unknown as StoryModule,
];

function storiesOf(mod: StoryModule): [string, Story][] {
  return Object.entries(mod)
    .filter(([name]) => name !== "default")
    .map(([name, value]) => [name, value as Story]);
}

for (const mod of MODULES) {
  const title = mod.default.title;

  describe(title, () => {
    for (const [exportName, story] of storiesOf(mod)) {
      const label = story.name ?? exportName;

      test(`${label} renders`, () => {
        const element = elementFor(mod, story);

        const { container } = render(element);
        expect(container.firstElementChild).not.toBeNull();
        cleanup();
      });
    }
  });
}

/**
 * A story's `render` is mounted as a component, not called as a function —
 * Storybook does the same, and it is what lets a render function use hooks.
 * A story without one falls back to `meta.component` with its args.
 */
function elementFor(mod: StoryModule, story: Story): ReactElement {
  const Component = (story.render ?? mod.default.component) as
    | ((props: Record<string, unknown>) => ReactElement)
    | undefined;
  if (!Component) {
    throw new Error("story has neither `render` nor a `component` on its meta");
  }
  return createElement(Component, story.args ?? {});
}

describe("play functions", () => {
  // Run separately from the render pass: a play function needs the story
  // mounted and left mounted while it drives it.
  const withPlay = MODULES.flatMap((mod) =>
    storiesOf(mod)
      .filter(([, story]) => typeof story.play === "function")
      .map(([exportName, story]) => ({
        title: `${mod.default.title} / ${story.name ?? exportName}`,
        mod,
        story,
      })),
  );

  test("there are some", () => {
    expect(withPlay.length).toBeGreaterThan(0);
  });

  for (const { title, mod, story } of withPlay) {
    test(title, async () => {
      const { container } = render(elementFor(mod, story));
      await story.play?.({ canvasElement: container as HTMLElement });
      cleanup();
    });
  }
});
