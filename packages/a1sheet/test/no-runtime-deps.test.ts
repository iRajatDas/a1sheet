/**
 * Architectural guards. These enforce constraints that are easy to violate
 * accidentally and expensive to discover later:
 *
 *   1. Zero runtime dependencies.
 *   2. No React outside src/react/, so the "." entrypoint stays usable in plain
 *      JS, Node, and Workers.
 *   3. The published React bundle declares "use client" as its first statement.
 */
import { describe, expect, test } from "bun:test";
import { readdir, readFile } from "node:fs/promises";
import { join } from "node:path";
import pkg from "../package.json" with { type: "json" };

const SRC = join(import.meta.dir, "..", "src");
const DIST = join(import.meta.dir, "..", "dist");

async function sourceFiles(dir: string): Promise<string[]> {
  const out: string[] = [];
  for (const entry of await readdir(dir, { withFileTypes: true })) {
    const full = join(dir, entry.name);
    if (entry.isDirectory()) out.push(...(await sourceFiles(full)));
    else if (/\.tsx?$/.test(entry.name) && !entry.name.endsWith(".test.ts")) {
      out.push(full);
    }
  }
  return out;
}

describe("zero runtime dependencies", () => {
  test("package.json declares no dependencies", () => {
    expect(pkg.dependencies).toEqual({});
  });

  test("react is a peer dependency, not a real one", () => {
    expect(pkg.peerDependencies).toHaveProperty("react");
    expect(Object.keys(pkg.dependencies ?? {})).not.toContain("react");
  });
});

describe("React containment", () => {
  test("nothing outside src/react/ imports react", async () => {
    const files = await sourceFiles(SRC);
    const offenders: string[] = [];

    for (const file of files) {
      if (file.includes(`${join("src", "react")}`)) continue;
      const text = await readFile(file, "utf8");
      // Match imports only — the word "React" in a comment is fine.
      if (/^\s*import[^;]*from\s+["']react(-dom)?["']/m.test(text)) {
        offenders.push(file.slice(SRC.length + 1));
      }
    }

    expect(offenders).toEqual([]);
  });

  test("the framework-agnostic barrel does not re-export the react entrypoint", async () => {
    const text = await readFile(join(SRC, "index.ts"), "utf8");
    expect(text).not.toMatch(/from\s+["']\.\/react/);
  });
});

describe("exports map", () => {
  test("every subpath points at a file that the build will emit", () => {
    const exports = pkg.exports as Record<string, unknown>;
    expect(Object.keys(exports)).toEqual([".", "./react", "./package.json"]);
  });
});

/**
 * Only meaningful once `bun run build` has run, so these skip rather than fail
 * on a clean checkout. Bundling buries the per-module directives mid-file, where
 * a directive is just a string expression and does nothing — the build hoists
 * one to the top, and that is what an RSC consumer reads.
 */
describe("the built bundles", () => {
  async function built(path: string): Promise<string | null> {
    try {
      return await readFile(join(DIST, path), "utf8");
    } catch {
      return null;
    }
  }

  test("the React entry declares 'use client' first", async () => {
    const js = await built(join("react", "index.js"));
    if (js === null) return;
    expect(js.split("\n")[0]).toBe('"use client";');
  });

  test("the framework-agnostic entry does not", async () => {
    // It has to stay importable from a server component and a Worker.
    const js = await built("index.js");
    if (js === null) return;
    expect(js.split("\n")[0]).not.toBe('"use client";');
  });

  test("the root entry ships real bindings, not a bare export list", async () => {
    const js = await built("index.js");
    if (js === null) return;
    const trim = js.trimStart();
    const bareExportOnly =
      trim.startsWith("export {") &&
      !/^\s*import\b/m.test(js) &&
      !/\bfunction\b|\bclass\b|\bconst\b|\blet\b|\bvar\b/.test(js);
    expect(bareExportOnly).toBe(false);

    const mod = await import(join(DIST, "index.js"));
    expect(typeof mod.A1SheetError).toBe("function");
    expect(typeof mod.readWorkbookFile).toBe("function");
    expect(typeof mod.makeSheet).toBe("function");
    expect(typeof mod.findAll).toBe("function");
  });
});
