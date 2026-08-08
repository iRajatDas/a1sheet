/**
 * Build script: ESM bundles via Bun, declarations via tsc.
 *
 * Two entrypoints matching the `exports` map in package.json:
 *   src/index.ts       -> dist/index.js        (model + formula + io + format)
 *   src/react/index.ts -> dist/react/index.js  (hooks + <Spreadsheet />)
 *
 * Splitting is OFF on purpose. With `splitting: true`, some Bun versions emit a
 * root `dist/index.js` that is only `export { … }` with no imports and omit the
 * shared chunk from the publishable tree — which made a1sheet@0.3.0's "." entry
 * unloadable. Self-contained entry bundles are slightly larger when both are
 * imported, but the root package must work alone.
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
  splitting: false,
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
 * Only this entry gets it. The root entry stays free of the directive so it can
 * load on a server or in a Worker.
 */
const REACT_ENTRY = `${OUT}/react/index.js`;
const entry = await Bun.file(REACT_ENTRY).text();
await Bun.write(REACT_ENTRY, `"use client";\n${entry}`);

/**
 * Refuse to ship a root barrel that only re-exports unbound names — that is the
 * a1sheet@0.3.0 failure mode (`Export 'A1SheetError' is not defined in module`).
 */
const rootJs = await Bun.file(`${OUT}/index.js`).text();
const rootTrim = rootJs.trimStart();
if (
  rootTrim.startsWith("export {") &&
  !/^\s*import\b/m.test(rootJs) &&
  !/\bfunction\b|\bclass\b|\bconst\b|\blet\b|\bvar\b/.test(rootJs)
) {
  console.error(
    "dist/index.js is a bare export list with no implementations — refusing to finish the build",
  );
  process.exit(1);
}

const probe = await import(`./${OUT}/index.js`);
for (const name of ["A1SheetError", "readWorkbookFile", "makeSheet"] as const) {
  if (typeof probe[name] !== "function" && typeof probe[name] !== "object") {
    console.error(`dist/index.js is missing a usable export: ${name}`);
    process.exit(1);
  }
}

const bytes = result.outputs
  .filter((o) => o.path.endsWith(".js"))
  .reduce((n, o) => n + o.size, 0);

console.log(
  `built ${result.outputs.length} files, ${(bytes / 1024).toFixed(1)} kB of JS`,
);
