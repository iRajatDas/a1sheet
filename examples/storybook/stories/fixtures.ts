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

/** A dark palette, as a partial Theme. There is no dark-mode boolean. */
export const DARK = {
  accent: "#2dd4bf",
  border: "#1e293b",
  headerBorder: "#334155",
  buttonBorder: "#334155",
  headerBg: "#0f172a",
  headerText: "#94a3b8",
  cellBg: "#0b1220",
  cellText: "#e2e8f0",
  selectedBg: "rgba(45,212,191,0.16)",
  toolbarBg: "#0f172a",
  // The scrollbar gutter is a real surface, so a dark theme has to colour it
  // too — left at the default it is a light stripe down the right-hand edge.
  scrollbarTrack: "#0f172a",
  scrollbarThumb: "#334155",
  scrollbarThumbHover: "#475569",
} as const;

/** Numbers in every supported format, for the formatting story. */
export function formats(): Workbook {
  return build(
    ["Formats"],
    [
      {
        cells: {
          "0_0": "Format",
          "0_1": "Raw",
          "0_2": "Shown",
          "1_0": "General",
          "1_1": "1234.5678",
          "1_2": "1234.5678",
          "2_0": "Integer",
          "2_1": "1234.5678",
          "2_2": "1234.5678",
          "3_0": "Number",
          "3_1": "1234.5678",
          "3_2": "1234.5678",
          "4_0": "Percent",
          "4_1": "0.4267",
          "4_2": "0.4267",
          "5_0": "Currency",
          "5_1": "1234.5678",
          "5_2": "1234.5678",
          "6_0": "Date",
          "6_1": "45000",
          "6_2": "45000",
        },
        styles: {
          "0_0": { bold: true },
          "0_1": { bold: true },
          "0_2": { bold: true },
          "2_2": { numFmt: "integer" },
          "3_2": { numFmt: "number" },
          "4_2": { numFmt: "percent" },
          "5_2": { numFmt: "currency" },
          "6_2": { numFmt: "date" },
        },
        frozenRows: 1,
        colWidths: { 0: 110 },
      },
    ],
  );
}

/** A locked cell and a hidden row, for the protection story. */
export function protectedSheet(): Workbook {
  return build(
    ["Protected"],
    [
      {
        cells: {
          "0_0": "Editable",
          "0_1": "try me",
          "1_0": "Locked",
          "1_1": "cannot change",
          "2_0": "Row 3 below is hidden",
          "3_0": "you cannot see this",
        },
        styles: {
          "0_0": { bold: true },
          "1_0": { bold: true },
          "1_1": { locked: true, bg: "#fee2e2" },
        },
        hiddenRows: new Set([3]),
      },
    ],
  );
}

/**
 * A quarter of regional numbers, as a real report would be laid out: a title,
 * a frozen header, currency and percent formats, a total row that sums the
 * ones above it, and a dynamic array that re-ranks the whole block.
 *
 * Deliberately not "Widget, Gadget, Doohickey" — this is the fixture the
 * showcase and the landing page use, and a demo that looks like a placeholder
 * reads as one.
 */
export function salesReport(): Workbook {
  return build(
    ["Q3", "Notes"],
    [
      {
        cells: {
          "0_0": "Q3 regional performance",
          "2_0": "Region",
          "2_1": "Owner",
          "2_2": "Bookings",
          "2_3": "Target",
          "2_4": "Attainment",
          "2_5": "Variance",
          "3_0": "North America",
          "3_1": "Priya Raman",
          "3_2": "1284300",
          "3_3": "1192400",
          "3_4": "=C4/D4",
          "3_5": "=C4-D4",
          "4_0": "EMEA",
          "4_1": "Tomás Herrera",
          "4_2": "968750",
          "4_3": "1040000",
          "4_4": "=C5/D5",
          "4_5": "=C5-D5",
          "5_0": "APAC",
          "5_1": "Wei Chen",
          "5_2": "742100",
          "5_3": "690000",
          "5_4": "=C6/D6",
          "5_5": "=C6-D6",
          "6_0": "LATAM",
          "6_1": "Ana Beatriz Lima",
          "6_2": "318400",
          "6_3": "402500",
          "6_4": "=C7/D7",
          "6_5": "=C7-D7",
          "7_0": "India",
          "7_1": "Devika Nair",
          "7_2": "596900",
          "7_3": "480000",
          "7_4": "=C8/D8",
          "7_5": "=C8-D8",
          "8_0": "Japan",
          "8_1": "Kenji Sato",
          "8_2": "233050",
          "8_3": "265000",
          "8_4": "=C9/D9",
          "8_5": "=C9-D9",
          "9_0": "Total",
          "9_2": "=SUM(C4:C9)",
          "9_3": "=SUM(D4:D9)",
          "9_4": "=C10/D10",
          "9_5": "=C10-D10",
          "11_0": "Ranked by bookings",
          "12_0": "=SORT(A4:C9, 3, -1)",
        },
        styles: {
          "0_0": {
            bold: true,
            fontSize: 16,
          },
          "2_0": {
            bold: true,
            bg: "#f1f5f9",
          },
          "2_1": {
            bold: true,
            bg: "#f1f5f9",
          },
          "2_2": {
            bold: true,
            bg: "#f1f5f9",
          },
          "2_3": {
            bold: true,
            bg: "#f1f5f9",
          },
          "2_4": {
            bold: true,
            bg: "#f1f5f9",
          },
          "2_5": {
            bold: true,
            bg: "#f1f5f9",
          },
          "3_2": {
            numFmt: "currency",
          },
          "3_3": {
            numFmt: "currency",
          },
          "3_4": {
            numFmt: "percent",
          },
          "3_5": {
            numFmt: "currency",
          },
          "4_2": {
            numFmt: "currency",
          },
          "4_3": {
            numFmt: "currency",
          },
          "4_4": {
            numFmt: "percent",
          },
          "4_5": {
            numFmt: "currency",
          },
          "5_2": {
            numFmt: "currency",
          },
          "5_3": {
            numFmt: "currency",
          },
          "5_4": {
            numFmt: "percent",
          },
          "5_5": {
            numFmt: "currency",
          },
          "6_2": {
            numFmt: "currency",
          },
          "6_3": {
            numFmt: "currency",
          },
          "6_4": {
            numFmt: "percent",
          },
          "6_5": {
            numFmt: "currency",
          },
          "7_2": {
            numFmt: "currency",
          },
          "7_3": {
            numFmt: "currency",
          },
          "7_4": {
            numFmt: "percent",
          },
          "7_5": {
            numFmt: "currency",
          },
          "8_2": {
            numFmt: "currency",
          },
          "8_3": {
            numFmt: "currency",
          },
          "8_4": {
            numFmt: "percent",
          },
          "8_5": {
            numFmt: "currency",
          },
          "9_0": {
            bold: true,
          },
          "9_2": {
            bold: true,
            numFmt: "currency",
          },
          "9_3": {
            bold: true,
            numFmt: "currency",
          },
          "9_4": {
            bold: true,
            numFmt: "percent",
          },
          "9_5": {
            bold: true,
            numFmt: "currency",
          },
          "11_0": {
            bold: true,
          },
        },
        colWidths: { 0: 150, 1: 168, 2: 120, 3: 120, 4: 110, 5: 120 },
        frozenRows: 3,
        condFormats: [
          {
            range: { r1: 3, c1: 4, r2: 8, c2: 4 },
            priority: 1,
            rule: { type: "cellIs", operator: "lessThan", operands: ["1"] },
            style: { color: "#b91c1c" },
          },
          {
            range: { r1: 3, c1: 4, r2: 8, c2: 4 },
            priority: 2,
            rule: {
              type: "cellIs",
              operator: "greaterThanOrEqual",
              operands: ["1"],
            },
            style: { color: "#15803d" },
          },
        ],
        numRows: 200,
        numCols: 12,
      },
    ],
  );
}
