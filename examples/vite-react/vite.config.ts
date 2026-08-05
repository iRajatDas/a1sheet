import { fileURLToPath } from "node:url";
import react from "@vitejs/plugin-react";
import { defineConfig } from "vite";

/**
 * Aliases a1sheet to source rather than dist, so the playground picks up edits
 * without a rebuild.
 */
const src = fileURLToPath(new URL("../../packages/a1sheet/src", import.meta.url));

export default defineConfig({
  plugins: [react()],
  resolve: {
    alias: [
      { find: /^a1sheet\/react$/, replacement: `${src}/react/index.ts` },
      { find: /^a1sheet$/, replacement: `${src}/index.ts` },
    ],
  },
});
