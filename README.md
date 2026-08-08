# a1sheet — a React spreadsheet component with Excel and CSV built in

[![npm](https://img.shields.io/npm/v/a1sheet.svg)](https://www.npmjs.com/package/a1sheet)
[![npm bundle size](https://img.shields.io/bundlephobia/minzip/a1sheet)](https://bundlephobia.com/package/a1sheet)
[![dependencies](https://img.shields.io/badge/dependencies-0-brightgreen)](https://www.npmjs.com/package/a1sheet?activeTab=dependencies)
[![license](https://img.shields.io/npm/l/a1sheet.svg)](LICENSE)

Read and write **Excel (.xlsx)** and **CSV** in the browser, evaluate
**spreadsheet formulas**, and render an editable, virtualized **data grid** —
from one package with **no runtime dependencies**.

No grid library. No parsing library. No date library. `dependencies` is empty and
a test enforces that it stays empty.

```tsx
import { Spreadsheet } from "a1sheet/react";

<Spreadsheet defaultWorkbook={wb} />;
```

**[Live demo and documentation →](https://irajatdas.github.io/a1sheet/)**

## Install

```sh
bun add a1sheet    # npm/pnpm/yarn all fine — React is a peer dependency
```

**ESM only.** There is no CommonJS build, and there will not be one: shipping
both is how a package ends up with two copies of its own module state in one
process, and the grid keeps state. Every current bundler and Node 18+ handle
this natively. From a CommonJS file, `await import("a1sheet")` works.

## Use

Two entrypoints. The root one has no React in it at all, so it works in plain JS,
Node, and Web Workers:

```ts
import { readWorkbookFile, writeXlsx, cellsToCSV } from "a1sheet";

const { sheets } = await readWorkbookFile(file);   // .xlsx or .csv, auto-detected
const bytes = writeXlsx(sheets);
```

The React layer is **composition-first**. You assemble the primitives; the library
never owns your layout:

```tsx
import { Sheet } from "a1sheet/react";

<Sheet.Root defaultWorkbook={wb} onWorkbookChange={save}>
  <Sheet.Toolbar>
    <Sheet.FileMenu />          {/* import/export — omit it and the XLSX
                                    writer never enters your bundle */}
  </Sheet.Toolbar>
  <Sheet.FormulaBar />
  <Sheet.Grid>
    <Sheet.AddRows />           {/* children of the grid sit at the end of
                                    the scrollable content */}
  </Sheet.Grid>
  <Sheet.Tabs />
  <Sheet.StatusBar />
  <Sheet.ContextMenu />
  <Sheet.ColumnMenu />
</Sheet.Root>
```

Reorder them, drop the ones you do not want, wrap them in your own layout, or mix
in your own components — they read the same context:

```tsx
import { Sheet, useSheet } from "a1sheet/react";

function SelectionSummary() {
  const api = useSheet();                      // same state the primitives use
  return <footer>{api.getDisplay(api.active.row, api.active.col)}</footer>;
}

<Sheet.Root>
  <Sheet.Grid />
  <SelectionSummary />
</Sheet.Root>
```

There are no `showToolbar`-style props and there never will be — a part you do not
want is a child you do not render.

If the default arrangement is fine, there is a preset:

```tsx
import { Spreadsheet } from "a1sheet/react";

<Spreadsheet defaultWorkbook={wb} />       // composes the primitives above
```

Or skip the components entirely and drive your own UI from the headless hook:

```tsx
import { useSpreadsheet } from "a1sheet/react";

const api = useSpreadsheet();
// api.sheet, api.selection, api.getDisplay(r, c), api.setCell(r, c, raw), …
```

### Controlled or uncontrolled

```tsx
<Sheet.Root defaultWorkbook={wb} />                        // uncontrolled
<Sheet.Root workbook={wb} onWorkbookChange={setWb} />      // controlled
```

### Imperative handle

```tsx
const ref = useRef<SheetRootHandle>(null);
<Sheet.Root ref={ref}>…</Sheet.Root>
ref.current?.focus();
ref.current?.api.setCell(0, 0, "hi");
```

Styling is inline with a `theme` prop and a `classNamePrefix` — no stylesheet to
import, so dropping this into an existing app needs no build-config change. A
dark theme should set `scrollbarTrack`, `scrollbarThumb`, and
`scrollbarThumbHover` along with the rest: the grid draws its own scrollbars in
channels beside the cells rather than letting the platform float them over the
content, so those channels are a real surface and stay light if you leave them
alone. `freezeLine` is the other one worth setting: it draws the edge of a
frozen band, and it has to read as heavier than an ordinary grid line or the
band and the rows scrolling under it look like one surface.

Icons are inline SVG from [Tabler Icons](https://tabler.io/icons) (MIT), copied
into the source rather than installed so that `dependencies` stays empty. See
`packages/a1sheet/THIRD-PARTY-NOTICES.md`.

### Integrating with Tailwind

The root element exposes `--a1s-accent`, `--a1s-border`, `--a1s-cell-bg`, and
the rest of the theme as CSS custom properties, so host styles can reference
them directly:

```tsx
import { Sheet, darkTheme, sheetsTheme } from "a1sheet/react";

// Light blue accent (`sheetsTheme`), or darkTheme. lightTheme stays teal.
<Sheet.Root theme={sheetsTheme} className="rounded-xl border shadow-sm">
  <Sheet.Toolbar className="border-b border-[var(--a1s-border)]">
    <Sheet.Toolbar.Undo />
    <Sheet.Toolbar.Bold />
    <Sheet.Toolbar.Separator />
    <Sheet.FileMenu />
  </Sheet.Toolbar>
  <Sheet.Grid
    className="flex-1 min-h-0"
    renderCellContent={({ display }) => <span className="font-medium">{display}</span>}
  />
</Sheet.Root>

// <Sheet.Root theme={darkTheme}>…</Sheet.Root>
```

Every chrome primitive accepts `className` and `style`. Toolbar and context-menu
controls are atoms you compose — there are no `show*` props. Use `asChild` on
`Sheet.Toolbar.IconButton` (or any atom built on it) to slot in your own button
component. Context and column menus portal to `document.body` and clamp against
the sheet root’s client rect (clipped to the visual viewport), so split-pane
hosts with `overflow: hidden` do not need custom positioning.

See **Composition → Showcase → Full application shell** in Storybook for a
complete example: Q3 report data, composed toolbar atoms, `renderCellContent`,
`useSheet()` sidebar, file I/O, and a separate **25,000-row** virtualization story.

### Keyboard

Everything the mouse can do to a selection has a keyboard route, because a grid
that needs a pointer is a grid some people cannot use.

| Keys | What they do |
| --- | --- |
| Arrows, Shift+Arrows | Move, extend |
| Ctrl/Cmd+Arrow | Run to the edge of the block of data |
| Ctrl/Cmd+Shift+Arrow | Extend to that edge |
| Home / End | Start of the row / its last filled cell |
| Ctrl+Home / Ctrl+End | A1 / the corner of the used range |
| PageUp / PageDown | By a screenful of rows, Shift to extend |
| Ctrl+A | The whole sheet |
| Ctrl+Space / Shift+Space | The column / the row |
| Shift+F8 | Add to selection — keep this range and start another |
| Ctrl+D / Ctrl+R | Fill down / right |
| F9 | Recalculate — refreshes `RAND`, `TODAY`, `NOW` |
| Escape | Drop the extra ranges, the copy outline, and add mode |

With the mouse: drag to select, Shift+click to extend, Ctrl/Cmd+click to add a
range — or to remove one already selected. The row and column headers take the
same three, and a drag across them selects the bands between.

A multiple selection is a real selection: clearing, formatting, and the status
bar's count all cover every range of it. Copying joins the ranges when they line
up into one block — same columns stacked, or same rows side by side — and
refuses anything else, as Excel does, rather than inventing a shape.

**Shift+F8** is how a discontiguous selection gets made without Ctrl+click: the
range you leave behind is kept as the cursor moves on, so Shift+F8, arrow away,
Shift+arrow to size the next one builds the same thing a mouse would.

**Ctrl+D and Ctrl+R** run the fill handle's engine rather than a plain copy. The
leading filled lines of the selection are the source and the rest is the
destination, so `1, 2` at the top of a selected column counts on to its end —
the thing you would otherwise have to drag for. Excel's Ctrl+D always copies its
top row; matching this library's own fill matters more than matching that.

A drag held near an edge scrolls the sheet, faster the further past the edge you
push, and keeps extending while you hold it there.

### Formula editing

Typing a formula puts the grid into reference-picking mode, as Excel and Sheets
do. Clicking a cell writes its reference at the caret instead of moving the
selection; dragging grows that one reference into a range; and every reference
in the formula is outlined in the grid, colour-matched so repeated references
read as the same thing.

A click only picks when the caret is somewhere a reference can go. After a
finished operand — `=A1+2` with the caret at the end — a click means what it
usually means, and selects.

The string mechanics are exported and framework-agnostic, so a custom editor
gets the same behaviour:

```ts
import { findRefSpans, insertRefAtCaret } from "a1sheet";

findRefSpans("=SUM(B2:C4)+A1");
// [{ start: 5, end: 10, text: "B2:C4", range: {…}, group: 0 },
//  { start: 12, end: 14, text: "A1",    range: {…}, group: 1 }]

insertRefAtCaret("=SUM(", 5, "B2");   // → { value: "=SUM(B2", caret: 7, span: {…} }
insertRefAtCaret("=A1+2", 5, "B2");   // → null: not a reference position
```

Error values carry an explanation rather than only a sentinel — a circular
reference reports itself in every cell of the cycle, not just the one that
closed it:

```ts
import { explainErrorValue } from "a1sheet";

explainErrorValue("#CYCLE!");
// "Circular reference: this formula depends on its own result. …"
```

### Sizing rows and columns

Drag the divider on a column header to resize the column, or the one below a row
header to resize the row. Double-click a divider to auto-fit: a column sizes to
its widest value, a row to the tallest of its wrapped cells.

**A row grows to fit wrapped text on its own**, as in Excel and Sheets — and
stops the moment you drag it, because an explicit height always wins over a
measured one. Double-clicking the row divider drops that height and hands the
row back to its content. Narrowing a column re-wraps and re-measures, so the
rows below it move too. A merged cell wraps at the merge's full width and grows
the last row it spans.

Wrapped text is measured in the face your theme draws it in, so restyling the
grid resizes its rows to match. Give `theme.fontSize` in px: the measurement
happens during render, where a relative unit has nothing to resolve against, and
a non-px value falls back to the default size.

Sizes live on the sheet, so they are yours to set and to persist:

```ts
sheet.colWidths[2] = 180;   // px; absent means the default
sheet.rowHeights[7] = 64;
```

Or through the API, which is what the grid itself uses:

```ts
api.setColWidth(2, 180);
api.setRowHeight(7, 64);
api.resetRowHeight(7);      // back to the default
```

### Performance

Both axes are virtualized: only cells near the viewport exist in the DOM. Cost
is flat in sheet size, not proportional to it. Measured under Bun with happy-dom
— relative numbers, not browser-accurate:

| sheet | cells in DOM | mount | per keystroke |
|---|---|---|---|
| 1k rows × 26 cols | 286 | 83 ms | 7.1 ms |
| 100k rows × 26 cols | 286 | 30 ms | 6.6 ms |
| 1k rows × 500 cols | 286 | 53 ms | 5.1 ms |
| 100k rows × 500 cols | 286 | 125 ms | 11.9 ms |

The scroll extent is the real one — a 100k-row sheet scrolls through
2,600,026 px — even though only about a dozen rows exist at any moment.

Two things do still scale with the sheet, and are worth knowing before you load
a million cells into one:

- **Committing a cell edit copies the cell map.** About 0.4 ms at 10k filled
  cells, 4 ms at 100k, 70 ms at 1M. This is the write path, not the render path,
  and it is the next thing to fix. Operations that touch no cell — resize,
  freeze, hide, filter, relabel — do not pay it and cost nothing measurable at
  any size.
- **A filter over a column of formulas rescans every row on every edit**
  (about 60 ms at 100k). Columns of plain values are re-tested incrementally and
  cost nothing measurable.

### Reading large files

`readWorkbookFile` (and `readXlsx`, and `csvToCells`) take an optional second
argument. A read yields to the browser between chunks, so a big file no longer
freezes the tab:

```ts
const controller = new AbortController();

const { sheets } = await readWorkbookFile(file, {
  signal: controller.signal,
  onProgress: ({ ratio, phase, detail }) => setBar(ratio, `${phase}: ${detail}`),
});
```

`onProgress` fires at most once per frame and once per 1% — a rate a React
consumer can call `setState` from directly. `ratio` never decreases and the last
report of a successful read is always `1`.

`controller.abort()` rejects the read with `AbortedError` (`code: "ABORTED"`) at
its next checkpoint. Nothing is half-applied: the workbook is built only after the
whole file has parsed, so a cancelled import leaves your existing sheet untouched.

Small inputs never yield — the reader only hands the thread back once it has
actually held it for a frame — so pasted text and small files cost what they
always did.

Large ones trade throughput for responsiveness, and the trade is not free. On a
37 MB, 600k-cell workbook, measured under Bun:

| | wall clock | longest uninterrupted block |
|---|---|---|
| before (blocking) | ~1.1 s | the entire read |
| now (paced) | ~1.7 s | ~60 ms |

The extra time is not the yielding itself — that is 4 ms across ~135 yields — it
is the runtime finally getting to collect garbage, which a read that never
releases the thread simply defers. A 1.7 s import you can watch and cancel beats
a 1.1 s frozen tab, so pacing is not opt-in: it applies whether or not you pass
`signal` or `onProgress`.

### Errors

Every throw is a typed class with a stable `code`. Branch on the code, never on
message text:

```ts
import { isA1SheetError } from "a1sheet";

try {
  await readWorkbookFile(file, { signal });
} catch (err) {
  if (!isA1SheetError(err)) throw err;
  switch (err.code) {
    case "ABORTED":            return;                    // user changed their mind
    case "UNSUPPORTED_FORMAT": return toast(err.message);  // .xlsb, .xls
    case "NOT_A_ZIP":
    case "MALFORMED_FILE":     return toast("That file isn't a spreadsheet.");
    default:                   throw err;
  }
}
```

## Common tasks

### Read an .xlsx or .csv file in the browser

`readWorkbookFile` detects the format by ZIP magic first and extension second, so
a renamed file still reads. It runs off the main thread's critical path for large
files — see [Reading large files](#reading-large-files).

```ts
import { readWorkbookFile } from "a1sheet";

const { format, sheets, namedRanges } = await readWorkbookFile(file);
const first = sheets[0];              // cells, styles, merges, sizing, tables…
```

### Export to Excel from React

```ts
import { downloadXlsx } from "a1sheet";

downloadXlsx(
  workbook.sheets.map((s) => ({
    name: s.name,
    cells: s.cells,
    styles: s.styles,
    merges: s.merges,
  })),
  { namedRanges: workbook.namedRanges, filename: "report.xlsx" },
);
```

### Export to CSV safely

Values beginning with `=`, `+`, `-`, `@`, tab, or CR are neutralized on the way
out, because a CSV opened in Excel will otherwise execute them. This is on by
default and documented under [What an import keeps](#what-an-import-keeps-and-what-it-does-not).

```ts
import { downloadCsv } from "a1sheet";

downloadCsv(sheet.cells, api.evaluator, "report.csv");
```

### Evaluate spreadsheet formulas without any UI

The `"."` entrypoint never imports React, so this runs in Node and in a Web
Worker as well as the browser.

```ts
import { createEvaluator } from "a1sheet";

const cells = { "0_0": "6", "1_0": "=A1*7" };
const ev = createEvaluator(cells, {});
ev.getCellDisplay(1, 0);              // 42
```

### Parse CSV text into cells

```ts
import { csvToCells } from "a1sheet";

const { cells, rows, cols } = await csvToCells("a,b\n1,2");
```

## Alternatives

Most of this ground is covered by more than one package, and the split is usually
file-handling *or* a grid *or* a formula engine. a1sheet is the three together,
which is worth it when they have to agree with each other and not worth it when
you only need one:

| If you need | Consider |
|---|---|
| Only to parse or write spreadsheet files, no UI | `xlsx` (SheetJS), `exceljs`, `papaparse` for CSV |
| Only a grid, with your own data layer | `react-data-grid`, `ag-grid`, `handsontable`, TanStack Table |
| Only a formula engine | `hyperformula`, `formulajs` |
| A spreadsheet: files, formulas, and an editable grid that agree | a1sheet |

Check each one's license and bundle size against your own constraints — they
differ, and some grids are commercially licensed. a1sheet is MIT with an empty
`dependencies`, which is the whole reason it exists.

## Repository

```
packages/a1sheet/     the library
examples/vite-react/  dev playground, aliased to src (no rebuild needed)
examples/storybook/   every use case as a story, also aliased to src
docs/                 LIMITATIONS.md
```

## Development

```sh
bun install
bun test           # bun's runner; happy-dom is preloaded for component tests
bun run typecheck
bun run lint       # biome
bun run build      # ESM via bun, .d.ts via tsc
bun run dev        # example app
bun run storybook  # every use case, on :6006
```

[The published Storybook](https://irajatdas.github.io/a1sheet/storybook/) is the
documentation, and `bun run storybook` is the same thing aliased straight to
`src` so an edit shows up without a build:

| Section | What is in it |
|---|---|
| Start here | Introduction, quick start, and why composition over configuration |
| Preset | `<Spreadsheet />` with a live props table |
| Composition | Grid only, reordered parts, your own component, custom toolbar |
| State | Controlled, uncontrolled, persistence, view state, the imperative handle |
| Hooks | A grid built from `useSpreadsheet` with no a1sheet component at all |
| Features | Editing, formulas, reference picking, formats, structure, fill, undo |
| Data and files | Import with progress and cancellation, typed errors, the CSV injection guard |
| Scale | 100,000 rows, 500 columns, filtering, resizing |
| Theming | Five palettes and every theme key |
| Recipes | Read-only, custom functions, named ranges, validation, search |

Stories are tested, not just built: `bun test` renders every one of them and
runs their play functions, so a broken demo fails in CI rather than in front of
whoever opened the docs.

## Architecture

The parts worth knowing before editing:

- **Cell addressing is `"${row}_${col}"`,** zero-indexed. `cells` (raw text,
  `"="`-prefixed for formulas) and `styles` are separate parallel maps and are
  never merged. Changing that means touching every read site in `model/`,
  `formula/`, `io/`, and `react/`.
- **`useWorkbook` is the only module that may write.** It clones on write and
  pushes history; every other hook receives `updateSheet`/`updateWorkbook` as
  arguments and cannot reach the workbook setter.
- **The evaluator is lazy and disposable.** A new one is built whenever `cells` or
  `namedRanges` change. That is intentional and cheap — do not try to persist or
  incrementally mutate one.
- **The grid is CSS Grid with sticky freeze panes** in a single scroll container.
  `gridAutoRows` is what lets row virtualization work with no spacer divs.

## What works

- **Grid** — CSS Grid with row and column virtualization, sticky freeze panes,
  merged cells, row and column resize with double-click auto-fit,
  row/column/sheet renaming, hidden rows.
- **Editing** — type to edit, F2, Enter/Tab/Escape, arrow and Shift+arrow
  navigation, Delete to clear, drag-select that auto-scrolls at the edges,
  Ctrl+click multi-select, and a keyboard route to all of it (see
  [Keyboard](#keyboard)).
- **Formulas** — over 100 functions including the dynamic-array set (`LET`, `LAMBDA`,
  `MAP`, `SORT`, `UNIQUE`, `FILTER`, `HSTACK`, `XLOOKUP`), array values that
  spill, elementwise operators, structured table references, cross-sheet
  references, defined names holding a formula, `$` anchoring, lazy memoized
  evaluation, cycle detection, live recalculation on edit.
- **Formatting** — bold/italic/underline, font family and size, horizontal and
  vertical alignment, wrapping, text and fill colour, gradients, borders, number
  formats by literal format code, conditional formatting, cell locking,
  Ctrl+B/I/U.
- **Structure** — insert/delete row and column, merge/unmerge, freeze, sort,
  per-column value and colour filters (plus named filter views), find/replace,
  checkboxes/hyperlinks, sheet-tab duplicate/move, hidden rows and columns, data-validation dropdowns,
  multiple sheets, undo/redo (50 levels).
- **Clipboard** — copy/cut/paste with relative-reference shifting on internal
  paste, TSV interchange with Excel and Sheets.
- **Fill handle** — drag in any direction, linear series extrapolation, weekday
  and month sequences, per-cell formula reference shifting. Rejects diagonal
  both-axis fills, partial merges, and out-of-bounds destinations via
  `previewFillCheck` (stable `AutofillCode`).
- **File I/O** — XLSX read and write (hand-written DEFLATE codec, ZIP, OOXML,
  shared strings, styles, themes, tables, conditional formats, images), CSV read
  and write. Verified against real files from Excel and LibreOffice in
  `dataset/`.

`.xlsb` and `.xls` are rejected with an actionable message rather than failing
obscurely — both are binary formats, out of scope.

### What an import keeps, and what it does not

Worth knowing before you point this at a real workbook.

**Formatting.** Values, formula text, merges, and custom column widths and row
heights. Fonts — family, size, bold, italic, underline, colour. Fills, including
gradients. Borders, with their line kinds and colours. Horizontal and vertical
alignment and text wrapping. Number formats by their literal format code, so `+45`
and `8/16/24 20:00` survive rather than collapsing to `45` and a bare date.

Colours are resolved through the workbook theme, which matters more than it
sounds: Excel writes most of them as `<color theme="4" tint="-0.25"/>`, not as
RGB. Formatting that lives outside the cell is read too — the borders and centring
a cell inherits from a named style, a table's header and banding from
`xl/tables/`, and conditional formats, which real workbooks use for headings and
not only for highlighting. In-cell `=IMAGE("…")` draws the picture the file
embeds.

Colour scales, data bars, and icon sets are drawn too, and data-validation lists
become dropdowns. Not read: charts, pivot tables, floating pictures and shapes.

**Formulas.** Kept as text and re-evaluated. Arrays are values, so a formula
returning one spills into the cells beside it; `LET`, `LAMBDA`, `MAP`, `SORT`,
`UNIQUE`, `FILTER`, `XLOOKUP` and the rest of the modern set work, as do
structured table references, cross-sheet references, and defined names that hold
a formula. `OFFSET` and `INDIRECT` build a reference rather than read one, so
they work too. Over 100 functions against Excel's several hundred.

The volatile functions — `RAND`, `RANDBETWEEN`, `TODAY`, `NOW` — depend on no
cell, so nothing on the sheet can say when they are stale. As in Excel they are
recomputed once per calculation cycle: every edit begins one, and **F9** begins
one without changing the document. Within a cycle they hold still, so `=RAND()`
in A1 and `=A1*2` in B1 always agree, and every `NOW()` reports the same instant.
`useSheet().recalculate()` is the same trigger for your own button.

Where evaluation still fails, the cell shows the value Excel last computed for
it, from `Sheet.cachedValues`. Such a cell is effectively read-only: nothing here
can recalculate a formula it could not parse, so editing its inputs will not
change it. Editing the cell itself drops the imported value and reveals the
error.

**Export** writes everything import reads: cells, styles, merges, sizing, tables,
conditional formats, and in-cell images. A table comes back with a neutral style
name — its appearance is flattened onto the cells at import, so re-declaring one
would paint it twice — and colours are written as literal RGB rather than through
a theme.

## Limitations

Every gap is an intentional scope cut, listed with the extension point where the
work would start: [`docs/LIMITATIONS.md`](https://github.com/iRajatDas/a1sheet/blob/main/docs/LIMITATIONS.md).

## License

MIT
