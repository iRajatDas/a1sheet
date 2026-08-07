/**
 * Storybook's Vite builder declares which Vite majors it supports, and a
 * package manager installs a mismatch without saying anything. Vite 8 split the
 * docs chunk so it evaluated before Storybook's preview runtime had defined the
 * globals it reads, and the published Storybook showed a loader forever — a
 * green build, a broken site. This asserts the installed Vite is one the
 * builder claims to support.
 */
import { dirname } from "node:path";
import { expect, test } from "bun:test";

const CARET_MAJOR = /^\^(\d+)\./;

/** The majors a `^a.b.c || ^d.e.f` range allows — the only form the builder publishes. */
function caretMajors(range: string): number[] {
  const parts = range.split("||").map((p) => p.trim());
  const majors = parts.map((part) => {
    const match = CARET_MAJOR.exec(part);
    if (!match) {
      throw new Error(
        `peer range "${range}" is not a union of carets any more; teach this test the new form`,
      );
    }
    return Number(match[1]);
  });
  return majors;
}

test("the installed Vite is a major @storybook/builder-vite supports", async () => {
  // The builder is a transitive dependency of the framework, so it resolves
  // from there rather than from here.
  const framework = dirname(
    Bun.resolveSync("@storybook/react-vite/package.json", import.meta.dir),
  );
  const builder = await Bun.file(
    Bun.resolveSync("@storybook/builder-vite/package.json", framework),
  ).json();
  const vite = await Bun.file(
    Bun.resolveSync("vite/package.json", import.meta.dir),
  ).json();

  const range = builder.peerDependencies?.vite as string | undefined;
  expect(range).toBeString();

  const installed = Number(vite.version.split(".")[0]);
  expect(caretMajors(range as string)).toContain(installed);
});
