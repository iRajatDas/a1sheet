/**
 * Build script: ESM bundles via Bun, declarations via tsc.
 *
 * Two entrypoints matching the `exports` map in package.json:
 *   src/index.ts       -> dist/index.js        (model + formula + io + format)
 *   src/react/index.ts -> dist/react/index.js  (hooks + <Spreadsheet />)
 *
 * `splitting` keeps shared model/formula code in one chunk so a consumer
 * importing both entrypoints does not get two copies.
 *
 * It also copies the repository README and LICENSE into the package, because npm
 * reads the ones beside package.json. Kept generated rather than checked in: two
 * copies of the same document drift, and the stale one is the one that ships.
 */
import { copyFile, rm } from "node:fs/promises";

const OUT = "dist";
const COPIED = ["README.md", "LICENSE"];

/**
 * Bun's transpiler picks react/jsx-dev-runtime unless NODE_ENV is production,
 * and it reads that at process start — assigning process.env.NODE_ENV here is
 * too late. So re-exec once with the env set. A published library must reference
 * the production JSX runtime; the dev one carries extra validation and is not
 * guaranteed to resolve in a consumer's production build.
 *
 * Done here rather than as a shell prefix in package.json so it works on Windows.
 */
if (process.env.NODE_ENV !== "production") {
  const child = Bun.spawnSync({
    cmd: [process.execPath, "run", import.meta.path],
    env: { ...process.env, NODE_ENV: "production" },
    stdout: "inherit",
    stderr: "inherit",
  });
  process.exit(child.exitCode ?? 1);
}

await rm(OUT, { recursive: true, force: true });

const result = await Bun.build({
  entrypoints: ["src/index.ts", "src/react/index.ts"],
  outdir: OUT,
  root: "src",
  target: "browser",
  format: "esm",
  splitting: true,
  sourcemap: "linked",
  minify: false,
  external: ["react", "react-dom", "react/jsx-runtime", "react/jsx-dev-runtime"],
});

if (!result.success) {
  for (const log of result.logs) console.error(log);
  process.exit(1);
}

// `bun x` rather than `bunx` — the latter is not always on PATH even when bun is.
const tsc = Bun.spawnSync({
  cmd: ["bun", "x", "tsc", "-p", "tsconfig.build.json"],
  stdout: "inherit",
  stderr: "inherit",
});

if (tsc.exitCode !== 0) process.exit(tsc.exitCode ?? 1);

for (const name of COPIED) await copyFile(`../../${name}`, name);

/**
 * Hoist `"use client"` to the top of the React bundle.
 *
 * Every source module under `src/react` declares it, but bundling leaves those
 * copies buried mid-file, where a directive is just a string expression and does
 * nothing. React Server Components read only the FIRST statement of a module, so
 * without this an App Router consumer importing `a1sheet/react` from a server
 * component gets hooks treated as server code.
 *
 * Only this entry gets it. The shared chunk holds the framework-agnostic model,
 * which the "." entry also imports and which must stay usable on a server.
 */
const REACT_ENTRY = `${OUT}/react/index.js`;
const entry = await Bun.file(REACT_ENTRY).text();
await Bun.write(REACT_ENTRY, `"use client";\n${entry}`);

const bytes = result.outputs
  .filter((o) => o.path.endsWith(".js"))
  .reduce((n, o) => n + o.size, 0);

console.log(
  `built ${result.outputs.length} files, ${(bytes / 1024).toFixed(1)} kB of JS`,
);
