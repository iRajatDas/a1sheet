/**
 * Registers a DOM globally so component tests can run under `bun test`, and
 * unmounts between tests.
 *
 * Preloaded via bunfig.toml, so no per-file import is needed.
 *
 * The cleanup registration matters: @testing-library/react auto-registers it
 * under Jest and Vitest but NOT under bun test, so without this every render
 * accumulates in the same document and queries start matching elements from
 * earlier tests.
 */

import { afterEach } from "bun:test";
import { GlobalRegistrator } from "@happy-dom/global-registrator";

if (typeof document === "undefined") {
  GlobalRegistrator.register();
}

const { cleanup } = await import("@testing-library/react");
afterEach(cleanup);
