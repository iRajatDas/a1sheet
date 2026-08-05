/**
 * Build script: ESM bundles via Bun, declarations via tsc.
 *
 * Two entrypoints matching the `exports` map in package.json:
 *   src/index.ts       -> dist/index.js        (model + formula + io + format)
 *   src/react/index.ts -> dist/react/index.js  (hooks + <Spreadsheet />)
 *
 * `splitting` keeps shared model/formula code in one chunk so a consumer
 * importing both entrypoints does not get two copies.
 */
import { rm } from "node:fs/promises";

const OUT = "dist";

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

const bytes = result.outputs
  .filter((o) => o.path.endsWith(".js"))
  .reduce((n, o) => n + o.size, 0);

console.log(
  `built ${result.outputs.length} files, ${(bytes / 1024).toFixed(1)} kB of JS`,
);
