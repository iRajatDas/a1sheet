/**
 * Builds the GitHub Pages site into `site/dist`.
 *
 *   /            a static landing page — real HTML, so it is indexable
 *   /assets/     the bundled live demo
 *   /storybook/  the full component documentation
 *
 * The landing page is hand-written HTML rather than another React app on
 * purpose: a crawler reads the headings and prose without running anything, and
 * the demo is progressive enhancement on top.
 */
import { cp, mkdir, rm } from "node:fs/promises";
import { join } from "node:path";

const here = import.meta.dir;
const OUT = join(here, "dist");
const SRC = join(here, "..", "packages", "a1sheet", "src");

await rm(OUT, { recursive: true, force: true });
await mkdir(join(OUT, "assets"), { recursive: true });

const built = await Bun.build({
  entrypoints: [join(here, "demo.tsx")],
  outdir: join(OUT, "assets"),
  target: "browser",
  format: "esm",
  minify: true,
  define: { "process.env.NODE_ENV": '"production"' },
  // Aliased to source so the demo always shows the commit it was built from,
  // not whatever version happens to be on npm.
  plugins: [
    {
      name: "a1sheet-source",
      setup(build) {
        build.onResolve({ filter: /^a1sheet(\/react)?$/ }, (args) => ({
          path:
            args.path === "a1sheet"
              ? join(SRC, "index.ts")
              : join(SRC, "react", "index.ts"),
        }));
      },
    },
  ],
});

if (!built.success) {
  for (const log of built.logs) console.error(log);
  process.exit(1);
}

await cp(join(here, "index.html"), join(OUT, "index.html"));
await cp(join(here, "404.html"), join(OUT, "404.html"));
await cp(join(here, "robots.txt"), join(OUT, "robots.txt"));
await cp(join(here, "sitemap.xml"), join(OUT, "sitemap.xml"));

const storybook = join(here, "..", "examples", "storybook", "storybook-static");
if (await Bun.file(join(storybook, "index.html")).exists()) {
  await cp(storybook, join(OUT, "storybook"), { recursive: true });

  // Storybook hard-codes `<title>storybook - Storybook</title>` into its shell
  // ahead of anything `managerHead` injects, and the first title in a document
  // is the one that counts. Rewriting it here beats fighting the template.
  const shell = join(OUT, "storybook", "index.html");
  const html = await Bun.file(shell).text();
  await Bun.write(shell, html.replace(/<title>[^<]*<\/title>/, ""));
} else {
  console.warn("no storybook-static — run `bun run storybook:build` first");
}

const bytes = built.outputs.reduce((n, o) => n + o.size, 0);
console.log(`site built, demo bundle ${(bytes / 1024).toFixed(0)} kB`);
