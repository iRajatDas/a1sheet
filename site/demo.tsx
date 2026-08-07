/**
 * The live demo on the landing page.
 *
 * Bundled with React included, because the page is a static HTML file on GitHub
 * Pages with no build step of its own and no import map. The whole point is that
 * a visitor edits a real cell and watches a real formula recalculate, so a
 * screenshot would not do.
 */

import { Spreadsheet } from "a1sheet/react";
import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import { salesReport } from "../examples/storybook/stories/fixtures.js";

const mount = document.getElementById("demo");
if (mount) {
  // The page renders a static placeholder for crawlers and for the moment
  // before this script arrives; replacing it is the signal that we are live.
  mount.innerHTML = "";
  createRoot(mount).render(
    <StrictMode>
      <Spreadsheet defaultWorkbook={salesReport()} height="100%" />
    </StrictMode>,
  );
}
