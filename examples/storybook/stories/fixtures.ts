/** Workbooks the stories are built from. Kept out of the stories themselves so
 *  each story reads as the thing it is demonstrating. */
import { createWorkbook, type Workbook } from "a1sheet";

type SheetPatch = Partial<Workbook["sheets"][number]>;

function build(names: string[], patches: SheetPatch[]): Workbook {
  const wb = createWorkbook(names);
  patches.forEach((patch, i) => {
    const sheet = wb.sheets[i];
    if (sheet) Object.assign(sheet, patch);
  });
  return wb;
}

export function budget(): Workbook {
  return build(
    ["Budget", "Notes"],
    [
      {
        cells: {
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
          "2_3": "=C3*B3",
          "3_0": "Doohickey",
          "3_1": "12",
          "3_2": "0.99",
          "3_3": "=B4*C4",
          "5_0": "Subtotal",
          "5_3": "=SUM(D2:D4)",
          "6_0": "Tax",
          "6_3": "=D6*0.2",
          "7_0": "Total",
          "7_3": "=D6+D7",
        },
        styles: {
          "0_0": { bold: true },
          "0_1": { bold: true },
          "0_2": { bold: true },
          "0_3": { bold: true },
          "1_2": { numFmt: "currency" },
          "2_2": { numFmt: "currency" },
          "3_2": { numFmt: "currency" },
          "1_3": { numFmt: "currency" },
          "2_3": { numFmt: "currency" },
          "3_3": { numFmt: "currency" },
          "5_3": { numFmt: "currency", bold: true },
          "6_3": { numFmt: "currency" },
          "7_0": { bold: true },
          "7_3": { numFmt: "currency", bold: true },
        },
        frozenRows: 1,
      },
    ],
  );
}

/** Every formula feature worth seeing at once, including the error values. */
export function formulas(): Workbook {
  return build(
    ["Formulas"],
    [
      {
        cells: {
          "0_0": "Input",
          "0_1": "Formula",
          "0_2": "Result",
          "1_0": "10",
          "1_1": "=A2*2",
          "1_2": "=B2",
          "2_0": "20",
          "2_1": "=SUM(A2:A4)",
          "3_0": "30",
          "3_1": "=AVERAGE(A2:A4)",
          "4_1": '=IF(A2>5,"big","small")',
          "5_1": '=CONCAT(A2," and ",A3)',
          "6_1": "=ROUND(A4/7,2)",
          "7_0": "Errors",
          "8_0": "=1/0",
          "8_1": "divide by zero",
          "9_0": "=SUM(A10:B10)",
          "9_1": "circular — it includes itself",
          "10_0": "=NOSUCHFN(1)",
          "10_1": "unknown function",
        },
        styles: {
          "0_0": { bold: true },
          "0_1": { bold: true },
          "0_2": { bold: true },
          "7_0": { bold: true },
        },
        frozenRows: 1,
      },
    ],
  );
}

/** Frozen panes, merges, and per-cell fills — the layout features together. */
export function layout(): Workbook {
  return build(
    ["Layout"],
    [
      {
        cells: {
          "0_0": "Q1 Regional Report",
          "1_0": "Region",
          "1_1": "Jan",
          "1_2": "Feb",
          "1_3": "Mar",
          "1_4": "Total",
          "2_0": "North",
          "2_1": "120",
          "2_2": "140",
          "2_3": "155",
          "2_4": "=SUM(B3:D3)",
          "3_0": "South",
          "3_1": "90",
          "3_2": "95",
          "3_3": "130",
          "3_4": "=SUM(B4:D4)",
          "4_0": "East",
          "4_1": "200",
          "4_2": "185",
          "4_3": "210",
          "4_4": "=SUM(B5:D5)",
          "5_0": "West",
          "5_1": "75",
          "5_2": "88",
          "5_3": "92",
          "5_4": "=SUM(B6:D6)",
        },
        styles: {
          "0_0": { bold: true, align: "center", bg: "#e0f2fe" },
          "1_0": { bold: true, bg: "#f1f5f9" },
          "1_1": { bold: true, bg: "#f1f5f9" },
          "1_2": { bold: true, bg: "#f1f5f9" },
          "1_3": { bold: true, bg: "#f1f5f9" },
          "1_4": { bold: true, bg: "#f1f5f9" },
          "2_4": { bold: true },
          "3_4": { bold: true },
          "4_4": { bold: true },
          "5_4": { bold: true },
        },
        merges: [{ r1: 0, c1: 0, r2: 0, c2: 4 }],
        frozenRows: 2,
        frozenCols: 1,
        rowHeights: { 0: 44 },
        colWidths: { 0: 130 },
      },
    ],
  );
}

const STATUSES = ["Open", "In progress", "Blocked", "Done"] as const;
const OWNERS = ["Ada", "Grace", "Alan", "Edsger", "Barbara"] as const;

/**
 * A sheet big enough that virtualization is the only reason it renders. Cells
 * are generated rather than listed so the story stays readable.
 */
export function large(rows: number, cols = 40): Workbook {
  const cells: Record<string, string> = {
    "0_0": "ID",
    "0_1": "Owner",
    "0_2": "Status",
    "0_3": "Score",
  };
  const styles: Record<string, { bold: boolean }> = {
    "0_0": { bold: true },
    "0_1": { bold: true },
    "0_2": { bold: true },
    "0_3": { bold: true },
  };
  for (let r = 1; r <= rows; r++) {
    cells[`${r}_0`] = `ROW-${r}`;
    cells[`${r}_1`] = OWNERS[r % OWNERS.length] as string;
    cells[`${r}_2`] = STATUSES[r % STATUSES.length] as string;
    cells[`${r}_3`] = String((r * 7919) % 1000);
  }
  return build(
    ["Large"],
    [{ cells, styles, numRows: rows + 1, numCols: cols, frozenRows: 1 }],
  );
}

/** Pre-filtered, to show the filter interacting with virtualization. */
export function filtered(): Workbook {
  const wb = large(5_000);
  const sheet = wb.sheets[0];
  if (sheet) sheet.filters = { 2: new Set(["Blocked", "Done"]) };
  return wb;
}
