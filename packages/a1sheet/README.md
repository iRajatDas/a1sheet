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

Two entrypoints. The root one contains no React at all, so it works in plain JS,
Node, and Web Workers:

```ts
import { readWorkbookFile, writeXlsx, createEvaluator } from "a1sheet";

const { sheets } = await readWorkbookFile(file);   // .xlsx or .csv, auto-detected
const bytes = writeXlsx(sheets);
```

The React layer is separate:

```tsx
import { Spreadsheet } from "a1sheet/react";

<Spreadsheet height={520} onChange={(wb) => save(wb)} />
```

Or drive your own UI from the headless hook:

```tsx
import { useSpreadsheet } from "a1sheet/react";

const api = useSpreadsheet();
// api.sheet, api.selection, api.getDisplay(r, c), api.setCell(r, c, raw), …
```

Styling is inline with a `theme` prop — no stylesheet to import, so dropping this
into an existing app needs no build-config change.

## Repository

```
packages/a1sheet/     the library
examples/vite-react/  dev playground, aliased to src (no rebuild needed)
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
```

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

Everything from the POC is ported, typed, and under test — 156 tests.

- **Grid** — CSS Grid with row virtualization, sticky freeze panes, merged cells,
  column resize, row/column/sheet renaming, hidden rows.
- **Editing** — type to edit, F2, Enter/Tab/Escape, arrow and Shift+arrow
  navigation, Delete to clear, drag-select, Ctrl+click multi-select.
- **Formulas** — 30 functions, `$` anchoring, named ranges, lazy memoized
  evaluation, cycle detection, live recalculation on edit.
- **Formatting** — bold/italic/underline, alignment, text and fill color, six
  number formats, cell locking, Ctrl+B/I/U.
- **Structure** — insert/delete row and column, merge/unmerge, freeze, sort,
  per-column value filters, multiple sheets, undo/redo (50 levels).
- **Clipboard** — copy/cut/paste with relative-reference shifting on internal
  paste, TSV interchange with Excel and Sheets.
- **Fill handle** — drag down or right, linear series extrapolation, per-cell
  formula reference shifting.
- **File I/O** — XLSX read and write (hand-written DEFLATE decoder, ZIP, OOXML,
  shared strings, styles), CSV read and write. Verified against real files from
  Excel and LibreOffice in `dataset/`.

`.xlsb` and `.xls` are rejected with an actionable message rather than failing
obscurely — both are binary formats, out of scope.

## Limitations

Intentional scope cuts, with the extension point for each:
[`docs/LIMITATIONS.md`](docs/LIMITATIONS.md).

## License

MIT
