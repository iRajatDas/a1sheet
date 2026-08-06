import { createWorkbook } from "a1sheet";
import { Sheet } from "a1sheet/react";
import { StrictMode } from "react";
import { createRoot } from "react-dom/client";

const workbook = createWorkbook(["Budget", "Notes"]);
const budget = workbook.sheets[0];
if (budget) {
  Object.assign(budget.cells, {
    "0_0": "Item",
    "0_1": "Qty",
    "0_2": "Price",
    "0_3": "Total",
    "1_0": "Widget",
    "1_1": "3",
    "1_2": "4.50",
    "1_3": "=B2*C2",
    "2_0": "Gadget",
    "2_1": "7",
    "2_2": "12.00",
    "2_3": "=B3*C3",
    "4_0": "Sum",
    "4_3": "=SUM(D2:D3)",
  });
  budget.styles["0_0"] = { bold: true };
  budget.styles["0_1"] = { bold: true };
  budget.styles["0_2"] = { bold: true };
  budget.styles["0_3"] = { bold: true };
  budget.frozenRows = 1;
}

const root = document.getElementById("root");
if (!root) throw new Error("#root missing");

createRoot(root).render(
  <StrictMode>
    <div style={{ maxWidth: 1100, margin: "0 auto" }}>
      <h1 style={{ font: "600 20px/1.3 system-ui", marginBottom: 16 }}>
        a1sheet playground
      </h1>
      {/* Composition, not configuration: rearrange or omit any of these. */}
      <Sheet.Root defaultWorkbook={workbook} height={520}>
        <Sheet.Toolbar />
        <Sheet.FormulaBar />
        <Sheet.Grid />
        <Sheet.Tabs />
        <Sheet.StatusBar />
        <Sheet.ContextMenu />
        <Sheet.ColumnMenu />
      </Sheet.Root>
    </div>
  </StrictMode>,
);
