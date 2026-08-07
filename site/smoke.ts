/**
 * Gates the Pages deploy without needing a browser.
 *
 * The React test suite can stay green while a Storybook built against an
 * unsupported Vite ships a blank docs page — that is what happened with Vite 8.
 * `peerRange.test.ts` catches the peer mismatch; this file catches a site build
 * that forgot Storybook or dropped the demo / preview bundles.
 *
 * A headless Chrome dump-dom pass is desirable but hangs on GitHub Actions'
 * Chrome build (SPA timers / sockets never settle under `--virtual-time-budget`),
 * so it is not the gate.
 */
import { existsSync, readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";

const DIST = join(import.meta.dir, "dist");
const failures: string[] = [];

function fail(message: string): void {
  failures.push(message);
}

function mustExist(rel: string): void {
  if (!existsSync(join(DIST, rel))) fail(`missing ${rel}`);
}

mustExist("index.html");
mustExist("assets/demo.js");
mustExist("storybook/index.html");
mustExist("storybook/iframe.html");
mustExist("storybook/index.json");

const demo = await Bun.file(join(DIST, "assets/demo.js")).text();
// Minifiers may split `"a1s-" + "root"`, but the prefix must appear.
if (!demo.includes("a1s-")) {
  fail("assets/demo.js does not contain the a1s- class prefix");
}

const iframe = await Bun.file(join(DIST, "storybook/iframe.html")).text();
const assetRefs = [...iframe.matchAll(/(?:src|href)="(\.\/assets\/[^"]+)"/g)].map(
  (m) => m[1]!.replace(/^\.\//, "storybook/"),
);
if (assetRefs.length === 0) {
  fail("storybook/iframe.html references no ./assets/* files");
}
for (const rel of assetRefs) {
  mustExist(rel);
}

// Library code is code-split into a chunk (often named fixtures-*); anywhere
// under storybook/assets with the class prefix proves it shipped.
const storyAssets = join(DIST, "storybook/assets");
const libraryChunks = readdirSync(storyAssets).filter(
  (name) =>
    name.endsWith(".js") &&
    readFileSync(join(storyAssets, name), "utf8").includes("a1s-"),
);
if (libraryChunks.length === 0) {
  fail("no storybook/assets/*.js contains the a1s- class prefix");
}

const index = await Bun.file(join(DIST, "storybook/index.json")).text();
if (!index.includes("preset-spreadsheet--basic")) {
  fail("storybook/index.json is missing preset-spreadsheet--basic");
}
if (!index.includes("start-here-introduction--docs")) {
  fail("storybook/index.json is missing start-here-introduction--docs");
}

if (failures.length > 0) {
  for (const failure of failures) console.error(`FAIL ${failure}`);
  process.exit(1);
}

console.log(
  `ok  site/dist static smoke (${libraryChunks.length} storybook chunks with a1s-)`,
);
