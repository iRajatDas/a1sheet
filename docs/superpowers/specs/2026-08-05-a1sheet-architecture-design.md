# a1sheet — Architecture Design

**Date:** 2026-08-05
**Status:** Approved
**Supersedes:** nothing (first spec; `ref/` POC is the input)

## Purpose

Turn the `ref/` proof of concept — a zero-dependency React spreadsheet in three
files, ~1650 lines — into a publishable library named `a1sheet`. The library reads
and writes Excel (XLSX) and CSV in the browser, evaluates formulas, and ships an
editable grid component.

The POC works. This is a productionization, not a redesign. Every design
constraint below is inherited from it deliberately.

## Non-negotiable constraints

Carried from the POC. Violating any of these is a design change requiring a new spec.

1. **Zero runtime dependencies.** No grid libraries, no parsers, no date
   libraries. A feature that appears to need one gets a minimal hand-written
   implementation, the way DEFLATE, ZIP, and OOXML already were. Enforced by a
   test, not a convention.
2. **`"r_c"` cell addressing.** Zero-indexed `` `${row}_${col}` `` string keys into
   flat objects. `cells` (raw text, `"="`-prefixed for formulas) and `styles`
   (style objects) stay **separate parallel maps**, never merged into a cell object.
3. **Fixed row height** (`ROW_HEIGHT = 26`). Only column width is resizable. The
   sticky-freeze-pane offset math depends on this.
4. **Single-sheet formulas.** No `Sheet2!A1`. The extension point is documented
   but not built.
5. **Drop-in consumption.** Any React project with JSX/ESM configured can render
   `<Spreadsheet />` with no build-config change. This forbids a CSS import in
   the public path.

## Naming

`a1sheet`. Chosen over `react-excel-editor` (better npm keyword match, worse
brand and inaccurate about the React-free core) and `excelkit` (Microsoft holds
the "Excel" trademark; acceptable for descriptive use, a poor foundation for a
product brand).

npm discoverability comes from the `description` and `keywords` fields rather
than the name. This was measured, not assumed: `apollo-spreadsheet` ranks page 1
for npm's "react excel" query on 16 weekly downloads, matching on its description
alone.

```json
{
  "name": "a1sheet",
  "description": "Zero-dependency React spreadsheet component with Excel (XLSX) and CSV read/write, formulas, and an editable grid.",
  "keywords": ["react","excel","xlsx","csv","spreadsheet","grid",
               "editor","viewer","parser","formula","sheet","workbook","ooxml"]
}
```

### Positioning

Measured weekly downloads of the incumbents, and what each actually does:

| Package | Weekly | Scope |
|---|---|---|
| `xlsx` (SheetJS) | 12,037,971 | parse/write, no UI |
| `react-csv` | 1,107,450 | write CSV only |
| `react-spreadsheet` | 147,062 | editable grid, no formulas, no XLSX I/O |
| `@progress/kendo-react-excel-export` | 63,791 | export only, commercial |
| `react-spreadsheet-import` | 19,167 | import wizard only |
| `react-excel-renderer` | 6,177 | read-only display |

Each owns exactly one verb: export, import, render, or parse. None combines
editing, formulas, and XLSX round-trip in one zero-dependency package. That
combination is the product.

## Repository shape

Bun workspaces. The POC stays frozen at `ref/` as reference and is never imported
by the package.

```
react-spreadsheet/
├─ package.json            # workspaces: ["packages/*", "examples/*"]
├─ bunfig.toml
├─ tsconfig.base.json
├─ biome.json
├─ ref/                    # frozen POC — reference only
├─ docs/
│  ├─ LIMITATIONS.md
│  └─ superpowers/specs/
├─ packages/a1sheet/       # the published library
└─ examples/vite-react/    # dev playground and drop-in smoke test
```

## Package: one package, two entrypoints

```
"."       → model + formula + io + format.  No React in this graph.
"./react" → hooks + <Spreadsheet />.        React is a peerDependency.
```

Not two packages. The I/O half must stand alone — usable from plain JS, Node, or
a Web Worker — but both halves always ship together, so separate packages would
buy version skew for nothing. Subpath exports give the same tree-shaking and the
same "no React in my bundle" guarantee.

`dependencies` stays permanently empty.

## Source layout

```
packages/a1sheet/src/
├─ index.ts                 # public barrel for "."
├─ model/                   # framework-agnostic
│  ├─ types.ts              # Workbook, Sheet, StyleObject, CellKey
│  ├─ address.ts            # colToLetters, lettersToCol, parseCellRef, cellKey
│  ├─ sheet.ts              # makeSheet, cloneSheet, insert/deleteRow|Col, sortByColumn
│  ├─ workbook.ts           # createWorkbook, sheet CRUD, namedRanges
│  └─ history.ts            # snapshot stack, cap 50
├─ formula/
│  ├─ tokenize.ts           # REF_RE runs before identifier scanning
│  ├─ parse.ts              # recursive descent, precedence preserved
│  ├─ evaluate.ts           # evalNode, createEvaluator, lazy memo, #CYCLE!
│  ├─ functions/            # math|text|logic|lookup|date + registry
│  ├─ refs.ts               # shiftFormulaRefs (token-level)
│  └─ series.ts             # extrapolateSeries
├─ io/
│  ├─ zip/                  # crc32, inflate, zip   (deflate = documented gap)
│  ├─ xlsx/                 # xml, read, write, styles, sharedStrings
│  ├─ csv/                  # read, write
│  └─ index.ts              # readWorkbookFile — dispatch on extension + zip magic
├─ format/numFmt.ts         # formatValue, numFmt enum
└─ react/
   ├─ useWorkbook.ts        # the only module that may clone-or-write
   ├─ useSelection.ts
   ├─ useEditing.ts
   ├─ useClipboard.ts
   ├─ useFillHandle.ts
   ├─ useRowWindow.ts
   ├─ useSpreadsheet.ts     # composes the hooks into one headless API
   ├─ Spreadsheet.tsx       # thin shell
   └─ components/           # Toolbar, FormulaBar, Grid, Cell, headers, tabs, menus
```

### Why `Spreadsheet.jsx` gets split

At 738 lines it is past the size where it can be held in context reliably, which
makes edits to it error-prone for both humans and agents. The split is along
seams the POC already has implicitly — selection, editing, clipboard, fill,
virtualization are each self-contained state machines that happen to live in one
function.

Three invariants govern the split:

- **`useWorkbook` owns all mutation.** It exposes `updateSheet(updater, addHistory)`
  and `updateWorkbook(updater, addHistory)`, clones on write via `cloneSheet`, and
  pushes history. Every other hook receives these as arguments and cannot reach
  the workbook directly. This makes "never mutate in place" structural rather
  than a comment.
- **`useRowWindow` owns the two-stage row mapping.** It unions manual `hiddenRows`
  with active `filters` into `effectiveHiddenRows`, builds the compacted
  `normalRowMapping` of visible absolute row indices, and derives the virtual
  window from `scrollTop`. `Grid` consumes its output and computes
  `gridRow = frozenRowsCount + visualIndex + 2`. Frozen rows remain a separate
  always-rendered band.
- **`Grid` stays CSS Grid with sticky freeze panes.** No 4-quadrant split with
  synced `scrollLeft` — that was considered and rejected in the POC.
  `gridAutoRows` is what makes virtualization work without spacer divs.

### The one behavioral cleanup

Clipboard logic is genuinely duplicated in the POC: the `onCopy`/`onPaste`
handlers on the hidden textarea, and a separate `navigator.clipboard` path for
the context menu. Both implement the internal-vs-external paste decision, and
they can drift. `useClipboard` consolidates them so `shiftFormulaRefs` is
applied in exactly one place, with the two transports as thin callers.

The `lastCopiedRef` text-comparison heuristic for detecting an internal paste is
preserved as-is, including its known failure mode (copying identical text from
another app between an internal copy and paste).

## TypeScript

Full port, not JS plus hand-written declarations. For a library the types are the
API surface, and the data model is precise enough to type cleanly. The formula
AST and `StyleObject` benefit immediately.

TypeScript is a devDependency. Runtime dependencies stay at zero.

This makes the port a real pass over all ~1650 lines rather than a file move.

## Build

`bun build` produces ESM; `tsc --emitDeclarationOnly` produces declarations; a
`build.ts` drives both. ESM only — no CJS. Using Bun's bundler avoids a
rollup/tsup devDependency.

## Styling

Inline styles, as in the POC, plus a `theme` prop covering the ~15 colors and
sizes and an optional `classNamePrefix`. No stylesheet in the public path, which
is what keeps the drop-in requirement literally true.

## Testing

`bun test`. No test-framework dependency.

- **`formula/`** — table-driven: expression → expected display. Every function,
  plus `#CYCLE!` detection, `$`-anchoring, and `shiftFormulaRefs` deltas.
- **`io/`** — round-trip: build workbook → `writeXlsx` → `readXlsx` → deep-equal.
  Plus committed `fixtures/` exported from real Excel, Sheets, and LibreOffice,
  asserting import does not throw and silently ignores charts and pivot tables.
- **`model/`** — row/column insert and delete key shifting, `sortByColumn` key
  rewriting, history cap at 50.
- **`react/`** — `@testing-library/react` with `happy-dom` (devDependencies):
  keyboard navigation, edit commit, fill handle, and the filter/hidden-row
  mapping.
- **`no-runtime-deps.test.ts`** — asserts `dependencies` is empty and that no
  module under `src/` outside `src/react/` imports `react`.

## Deferred work

Reproduced in `docs/LIMITATIONS.md` and the README. Each is an intentional scope
cut, not a bug.

| Limitation | Extension point |
|---|---|
| No cross-sheet formulas (`Sheet2!A1`) | `tokenize.ts` needs a `!`-anchored lookahead before the ref pattern — a bare regex collides with sheet names ending in letters+digits. `createEvaluator` moves from per-sheet to per-workbook to expose `getSheet(name)`. |
| ZIP export is STORE-only | No DEFLATE *encoder* exists, only the decoder. Add one at `io/zip/deflate.ts` and switch `makeZip`'s method field to 8. |
| No borders | Add fields to `StyleObject`, render in `Cell`, add a `<borders>` section to the `styles.xml` writer and reader (currently a single empty `<border/>` placeholder). |
| No conditional formatting | New per-sheet `condFormats: [{range, rule, style}]`, resolved in `Cell` after the base style and before inline style fields are read. |
| No column hiding | Needs the `normalRowMapping` visual-index compaction mirrored for columns, interacting with `frozenCols` the way `frozenRows` interacts with `effectiveHiddenRows`. |
| Fixed row height | Variable heights break the sticky `top` offset math for frozen rows. Requires a cumulative offset table. |
| Fill handle down/right only | `commitFillDrag` guards on `targetR > rMax` / `targetC > cMax`. |
| Multi-range selection is status-bar only | `extraRanges` is deliberately ignored by copy, fill, and paste. |
| Paste aligns from top-left, no shape validation | — |
| Date serials are Unix-epoch, not Excel's 1899-12-30 | Internally consistent for arithmetic; a serial copied out will not match Excel's for the same date. |
| XLSX import ignores charts, pivot tables, data validation | Ignored rather than erroring. |

## Test corpus

`dataset/` holds 23 real-world files (38 MB), which resolves the fixture question.
Format breakdown and what each is usable for:

| Format | Count | Status |
|---|---|---|
| `.xlsx` | 4 | In scope. Primary round-trip fixtures. |
| `.xlsm` | 5 | Readable — macro-enabled XLSX, same ZIP+OOXML container. Read-only fixtures; we do not preserve the macro parts on write. |
| `.xlsb` | 12 | **Out of scope.** Binary BIFF12, not XML. Would need an entirely separate parser. |
| `.xls` | 1 | **Out of scope.** Legacy BIFF8 compound-file format. |
| `.csv` | 1 | In scope (10,000 rows — also a useful perf fixture). |

So only 9 of the 23 files exercise the reader, and 4 exercise the writer. The 13
`.xlsb`/`.xls` files are still valuable as *negative* fixtures: importing one must
fail with a clear "unsupported format" error rather than crashing or silently
producing an empty sheet. `readWorkbookFile` currently detects ZIP magic, which
correctly rejects `.xls` but will wrongly accept `.xlsb` — that needs a
content-type check against `[Content_Types].xml` before it can report accurately.

Adding `.xlsb` or `.xls` support is a separate spec, not a scope creep of this one.
