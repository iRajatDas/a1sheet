/**
 * Boots the built site in a headless browser and fails if nothing rendered.
 *
 * The test suite renders every story through the React entrypoint, which proves
 * the library works and proves nothing about the bundle the site ships: a
 * Storybook built against an unsupported Vite spun a loader forever while the
 * suite stayed green. This serves `site/dist` exactly as Pages does and asserts
 * that the landing demo and a story both put a real grid in the DOM.
 */
import { spawnSync } from "node:child_process";
import { existsSync } from "node:fs";
import { join } from "node:path";

/** The grid's root element. Present only if the library actually mounted. */
const MOUNTED = 'class="a1s-root"';

const CHROME_CANDIDATES = [
  process.env.CHROME_PATH,
  "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
  "/usr/bin/google-chrome",
  "/usr/bin/chromium",
  "/usr/bin/chromium-browser",
];

function findChrome(): string | null {
  for (const path of CHROME_CANDIDATES) {
    if (path && existsSync(path)) return path;
  }
  return null;
}

const chrome = findChrome();
if (!chrome) {
  const message =
    "no Chrome found; set CHROME_PATH to smoke-test the built site";
  // Skipping in CI would make the check decorative, which is worse than absent.
  if (process.env.CI) throw new Error(message);
  console.warn(message);
  process.exit(0);
}

const DIST = join(import.meta.dir, "dist");
if (!existsSync(join(DIST, "index.html"))) {
  throw new Error("no site/dist — run `bun run site` first");
}

const server = Bun.serve({
  port: 0,
  async fetch(request) {
    const { pathname } = new URL(request.url);
    const path = pathname.endsWith("/") ? `${pathname}index.html` : pathname;
    const file = Bun.file(join(DIST, path));
    if (!(await file.exists())) return new Response("Not found", { status: 404 });
    return new Response(file);
  },
});

const PAGES = [
  { name: "landing demo", path: "/" },
  {
    name: "Storybook story",
    path: "/storybook/iframe.html?id=preset-spreadsheet--basic&viewMode=story",
  },
];

// Long enough for the bundle to parse, mount, and paint on a cold CI runner.
const VIRTUAL_TIME_MS = 20_000;
/** Wall-clock ceiling — virtual time does not always stop Chrome from hanging. */
const CHROME_TIMEOUT_MS = VIRTUAL_TIME_MS + 15_000;

const failures: string[] = [];
for (const page of PAGES) {
  const url = `http://127.0.0.1:${server.port}${page.path}`;
  const { stdout, status, error, signal } = spawnSync(
    chrome,
    [
      // Classic headless: `--headless=new` hangs under dump-dom + virtual time
      // on the GitHub Actions Chrome build we install.
      "--headless",
      "--disable-gpu",
      "--no-sandbox",
      "--disable-dev-shm-usage",
      "--hide-scrollbars",
      `--virtual-time-budget=${VIRTUAL_TIME_MS}`,
      "--dump-dom",
      url,
    ],
    {
      encoding: "utf8",
      maxBuffer: 64 * 1024 * 1024,
      timeout: CHROME_TIMEOUT_MS,
    },
  );

  if (error) {
    failures.push(`${page.name}: ${error.message}`);
  } else if (status !== 0) {
    failures.push(
      `${page.name}: chrome exited ${status}${signal ? ` (${signal})` : ""}`,
    );
  } else if (!stdout.includes(MOUNTED)) {
    failures.push(`${page.name}: no ${MOUNTED} in the rendered DOM`);
  } else {
    console.log(`ok  ${page.name}`);
  }
}

server.stop(true);

if (failures.length > 0) {
  for (const failure of failures) console.error(`FAIL ${failure}`);
  process.exit(1);
}
