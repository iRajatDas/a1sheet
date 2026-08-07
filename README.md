# a1sheet

Zero-dependency Excel (XLSX) and CSV handling for the browser, plus an editable
React spreadsheet component with a formula engine.

No grid library. No parsing library. No date library. `dependencies` is empty and
a test enforces that it stays empty.

## Install

```sh
bun add a1sheet    # npm/pnpm/yarn all fine — React is a peer dependency
```

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
alone.

Icons are inline SVG from [Tabler Icons](https://tabler.io/icons) (MIT), copied
into the source rather than installed so that `dependencies` stays empty. See
`packages/a1sheet/THIRD-PARTY-NOTICES.md`.

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
its widest value, a row returns to the default height — a wrapped cell does not
yet grow its row, so there is no taller content for a row to hug.

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

- **Committing an edit copies the cell map.** `useWorkbook` clones on write, so
  an edit costs about 26 ms at 10k filled cells and about 390 ms at 1M. This is
  the write path, not the render path, and it is the next thing to fix.
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

## Repository

```
packages/a1sheet/     the library
examples/vite-react/  dev playground, aliased to src (no rebuild needed)
examples/storybook/   every use case as a story, also aliased to src
ref/                  the original POC — frozen, reference only, never imported
docs/                 LIMITATIONS.md and the architecture spec
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

`bun run storybook` is the documentation — 12 docs pages and 49 interactive
stories, aliased straight to `src` so an edit shows up without a build:

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

Full design in
[`docs/superpowers/specs/2026-08-05-a1sheet-architecture-design.md`](docs/superpowers/specs/2026-08-05-a1sheet-architecture-design.md).

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

Everything from the POC is ported, typed, and under test.

- **Grid** — CSS Grid with row and column virtualization, sticky freeze panes,
  merged cells, row and column resize with double-click auto-fit,
  row/column/sheet renaming, hidden rows.
- **Editing** — type to edit, F2, Enter/Tab/Escape, arrow and Shift+arrow
  navigation, Delete to clear, drag-select, Ctrl+click multi-select.
- **Formulas** — ~90 functions including the dynamic-array set (`LET`, `LAMBDA`,
  `MAP`, `SORT`, `UNIQUE`, `FILTER`, `HSTACK`, `XLOOKUP`), array values that
  spill, elementwise operators, structured table references, cross-sheet
  references, defined names holding a formula, `$` anchoring, lazy memoized
  evaluation, cycle detection, live recalculation on edit.
- **Formatting** — bold/italic/underline, font family and size, horizontal and
  vertical alignment, wrapping, text and fill colour, gradients, borders, number
  formats by literal format code, conditional formatting, cell locking,
  Ctrl+B/I/U.
- **Structure** — insert/delete row and column, merge/unmerge, freeze, sort,
  per-column value filters, hidden rows and columns, data-validation dropdowns,
  multiple sheets, undo/redo (50 levels).
- **Clipboard** — copy/cut/paste with relative-reference shifting on internal
  paste, TSV interchange with Excel and Sheets.
- **Fill handle** — drag in any direction, linear series extrapolation, per-cell
  formula reference shifting.
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
a formula. About 90 functions against Excel's several hundred — no `OFFSET`,
`INDIRECT`, or volatile functions.

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

Every gap is listed with the extension point where the work would start:
[`docs/LIMITATIONS.md`](docs/LIMITATIONS.md).

## Limitations

Intentional scope cuts, with the extension point for each:
[`docs/LIMITATIONS.md`](docs/LIMITATIONS.md).

## License

MIT
