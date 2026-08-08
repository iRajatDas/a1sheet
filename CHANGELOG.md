# Changelog

Notable user-visible changes. This project follows [semver](https://semver.org)
honestly: breaking means major.

## Unreleased

## 0.3.3

### Fixed

- **Context and column menus use the sheet root (clipped to the visual viewport)
  as collision bounds**, not only `window.innerHeight`/`innerWidth`. In a short
  split pane the menu flips above the cursor when there is no room below the
  host edge. Menus also **portal to `document.body`** so host `overflow: hidden`
  cannot clip them, and **re-clamp via ResizeObserver** when the menu size
  changes.

## 0.3.2

### Fixed

- **Context and column menus dismiss on grid scroll** and **clamp into the
  viewport** (flip above / slide left when there is not enough room), so they no
  longer float over the wrong cells or spill off-screen.

## 0.3.1

### Fixed

- **Root `"a1sheet"` entry is loadable again.** `0.3.0` shipped `dist/index.js` as a
  bare `export { … }` list with no module bodies (and no shared chunk), so
  `import { readWorkbookFile, makeSheet, … } from "a1sheet"` threw
  `Export '…' is not defined in module`. The build now emits self-contained
  entry bundles and refuses to finish if the root barrel is empty.

## 0.3.0

### Breaking

- **`Sheet.filters` values are `ColumnFilter` objects**, not bare `Set<string>`.
  A value-only filter is `{ values: new Set([...]) }`. Colour filters use
  `background` / `foreground` sets on the same object.

### Added

- **Composable toolbar atoms** — `Sheet.Toolbar.Undo`, `.Bold`, `.Merge`, and the
  rest, each with optional `asChild` via `Sheet.Toolbar.IconButton`. The shell
  renders the default set when given no children; `<Spreadsheet />` composes the
  full preset explicitly.
- **`Sheet.ContextMenu.*` items** — copy, paste, insert/delete, and a generic
  `.Item` / `.Separator` for custom entries. Keyboard navigation with arrow keys
  inside an open menu.
- **`Sheet.Toolbar.Overflow`** — collapses secondary controls behind a menu on
  narrow viewports.
- **`className` / `style` on every chrome primitive**, plus `--a1s-*` CSS custom
  properties on the root for Tailwind-compatible theming.
- **`darkTheme` / `lightTheme` presets** exported from `a1sheet/react`.
- **`sheetsTheme`** — opt-in light palette with `#1a73e8` accent. `lightTheme`
  stays teal; pass `theme={sheetsTheme}` on `Sheet.Root`.
- **`renderCellContent` and `components.CellContent`** — custom cell display
  without replacing selection or editing.
- **`Part` and `mergeClass`** exported for consumers building their own primitives.
- **Data-validation enforcement** for list, numeric, and text-length rules on
  edit and paste; paste shape mismatch warns when the selection does not match
  the clipboard block.
- **Touch** — larger resize/fill handles on coarse pointers; long-press opens the
  context menu.
- **Accessibility** — `role="grid"` / `gridcell`, toolbar `role="toolbar"`, live
  status region on the status bar.
- **Autofill guards** — `previewFillCheck` rejects diagonal both-axis fills,
  destinations past the grid edge, and ranges that partially cut a merge, with
  stable `AutofillCode` values and Sheets-compatible status messages.
- **Fill sequences** — weekdays and months (full and abbreviated) continue with
  wraparound via `SEQUENCES` / `extrapolateSeries`.
- **Paste special** — `PasteMode` on clipboard paste: values, formats, formulas,
  transpose, and text. Context menu exposes each; formats ignore cell locks;
  values skip locked cells. Status reports `Pasted values into A1` etc.
- **Merge guards** — paste, sort, and filter refuse ranges that partially cut a
  merge (`PASTE_PARTIAL_MERGE`, `SORT_PARTIAL_MERGE`, `FILTER_PARTIAL_MERGE`).
- **Undo/redo status** — successful undo reports `Action was undone.`; redo with
  an empty stack reports `Couldn't redo. Try again.`
- **`GridError` codes** — `MERGE_RANGE_SINGLETON`, `GRID_RANGE_UNBOUNDED`,
  `FILTER_ID_EXISTS`, and related codes for fail-fast grid boundaries. Branch on
  `error.code`.
- **Colour filters and filter views** — `ColumnFilter` accepts `values`,
  `background`, and `foreground` sets; `createFilterView` / `activateFilterView`
  / `deleteFilterView` snapshot and restore them. Missing views report
  `The view does not exist.`; `sortByColor` moves matching fill/text colours to
  the top with Sheets-compatible status text.
- **Formula catalog** — `FORMULA_CATALOG` / `suggestFormulas` expose category and
  argument-type metadata (Sheets help-content enums) for autocomplete. `FLATTEN`
  stacks ranges into a column.
- **Freeze status** — `freezeToSelection` / `unfreeze` report Freeze/Unfreeze
  row(s)/column(s) status strings.
- **Find & replace** — headless `findAll` / `findNext` / `replaceAll` with
  match-case, entire-cell, and formula search options; status
  `Replaced N occurrence(s).`
- **Sheet tabs** — `duplicateSheet` / `moveSheet` (last sheet cannot be deleted).
- **Checkboxes, hyperlinks, text rotation** — `insertCheckboxes`, `setHyperlink`,
  `setTextRotation` on the public API; checkbox cells toggle TRUE/FALSE.

### Changed

- **Cell edits use `patchSheet`** — shallow-copy only `cells`, `cachedValues`, and
  `images` instead of cloning the whole sheet. Full persistent cell maps remain
  future work for million-cell sheets.

### Fixed

- **Formula paste matches Excel relative-ref rules more closely** — internal
  paste still shifts after a trailing clipboard newline; sheet-qualified refs
  keep their sheet name while the address moves; keyboard paste lands on the
  active cell (not the selection’s far corner).
- **Storybook on the docs site loads again.** Vite 8 was installed under Storybook
  even though `@storybook/builder-vite` only peers `^5 || ^6 || ^7`. The Vite 8
  build split the docs chunk so it ran before Storybook's preview globals existed
  (`__STORYBOOK_MODULE_CORE_EVENTS_PREVIEW_ERRORS__`), and the published docs
  page stayed blank. Storybook now pins Vite 7; a peer-range test and a static
  smoke of `site/dist` gate the Pages deploy so a green suite cannot ship a dead
  Storybook again.

## 0.2.0

### Breaking

- **`downloadXlsx(sheets, options)`.** The filename moved into the options object
  alongside `namedRanges`, rather than being a third positional argument.
- **`writeXlsx(sheets, options)`** takes workbook-level defined names in
  `options`; a sheet's own names travel on its entry. Additive, but the meaning
  of `XlsxSheetInput.namedRanges` changed from "the workbook's" to "this
  sheet's".
- **`Sheet` gained `namedRanges` and `namedFormulas`.** Build sheets with
  `makeSheet`, not object literals.
- **`useRowWindow` takes one options object** instead of four positional
  arguments. It is exported for consumers writing their own grid; the primitives
  and `useSpreadsheet` are unaffected.

### Fixed

- **Arrays chain.** `=SORT(A1:A3)` now reads what `=SEQUENCE(3)` spilled into A2
  and A3, across sheets as well as within one. The index of where arrays land
  used to be built in a single prepass, and a formula asking about a cell
  mid-pass got a blank — so the second array in a chain sorted `[1, "", ""]`.
- **Sheet-scoped defined names survive an import.** A name Excel scoped to one
  sheet was skipped outright, leaving every formula using it reading `#NAME?`
  behind the cached value. They now live on the sheet, shadow a workbook name of
  the same spelling, and export again with their `localSheetId`.

### Performance

- **Edits that touch no cell no longer copy the cell map.** Resize, freeze, hide,
  filter, relabel and merge go through the new `patchSheet`, which replaces named
  fields and leaves `cells` and `styles` as the same objects. Column resize fires
  on every mousemove, so at a million filled cells that path went from a 70 ms
  stall per frame to unmeasurable. A cell edit still copies the map — 0.4 ms at
  10k filled cells, 70 ms at 1M — and that needs `cells` to stop being a plain
  object, which is not a patch release.

### Added

- **`patchSheet`** on the headless API, for consumers writing operations that
  touch a row/column-keyed container.
- **A published site**: a landing page and the full Storybook at
  <https://irajatdas.github.io/a1sheet/>, deployed from `main`.

## 0.1.0

First release. Everything before it is development history, kept in git rather
than here — nobody can have upgraded from a version that was never published.

**Zero runtime dependencies**, and a test that fails if one is ever added. No
grid library, no XLSX library, no date library: the ZIP reader, the DEFLATE
codec both ways, the OOXML scanner, the formula engine, and the virtualized grid
are all in this package.

**Two entrypoints, permanently.** `a1sheet` is framework-agnostic and never
reaches React, so it runs in Node and in a Web Worker; `a1sheet/react` is the
only place React appears. A test enforces the boundary.

**Composition, not configuration.** `Sheet.Root` provides the state and every
other part is a child you place, replace, or leave out — there is no
`showToolbar` prop, because an absent toolbar is an unrendered
`<Sheet.Toolbar/>`. `<Spreadsheet />` is a preset that composes the default
arrangement and takes no layout props of its own. Leaf primitives take `asChild`.

### The grid

Row and column virtualization over a single CSS Grid with sticky freeze panes,
variable row heights and column widths, merged cells, hidden rows, column
filters, resize with double-click auto-fit, and rows that grow to fit wrapped
text. Type to edit, F2, Enter/Tab/Escape, arrow and Shift+arrow navigation,
Ctrl+arrow to the edge of the data, drag-select that auto-scrolls past the edge,
Ctrl+click multi-select, a fill handle in four directions, and a keyboard route
to all of it. Formula editing highlights the references you are typing and lets
you pick them by clicking the grid.

### Formulas

Over 100 functions, including the dynamic-array set — `LET`, `LAMBDA`, `MAP`,
`SORT`, `UNIQUE`, `FILTER`, `HSTACK`, `XLOOKUP`, `SEQUENCE` — with array values
that spill, elementwise operators, structured table references, cross-sheet
references, defined names, `$` anchoring, and `$`-aware reference rewriting on
insert, delete, fill, and paste. Evaluation is lazy and memoized per calculation
cycle, so `RAND()` holds still within a cycle and moves across one; F9 (or
`recalculate()`) begins a new one. Cycles are detected and reported as `#CYCLE!`
in every participating cell rather than silently reading zero.

Arrays chain: `=SORT(A1:A3)` reads what `=SEQUENCE(3)` spilled into A2 and A3,
across sheets as well as within one. Anchors resolve on demand rather than in a
prepass, which is what makes the values exist by the time the second formula
asks for them.

### Files

XLSX and CSV, read and write. Import keeps number formats, fonts, fills,
borders, alignment, merges, freeze panes, column widths and row heights, tables,
conditional formats, data validation, defined names, and in-cell images.
Anything it cannot evaluate keeps the value Excel cached, so an unsupported
formula shows its number instead of an error. Reads above a size threshold are
async, take an `AbortSignal`, and report progress.

CSV export neutralizes values beginning with `=`, `+`, `-`, `@`, tab, or CR,
which is a real code-execution vector when the file is opened in Excel.
Download filenames are sanitized against separators and traversal.

### Performance

Rows and columns are both virtualized, over real scroll extents — a 100k-row
sheet scrolls through 2,600,026 px with about a dozen rows mounted. Row offsets
skip their lookup table entirely while every row is the default height.

Edits that touch no cell — resize, freeze, hide, filter, relabel — replace only
the field they change and leave the cell map alone, so they cost nothing
measurable at any size. That matters most for column resize, which fires on
every mousemove. A cell edit still copies the map: 0.4 ms at 10k filled cells,
70 ms at 1M.

### Types and errors

`strict` with `noUncheckedIndexedAccess`, no `any`, `readonly` on what consumers
should not mutate, and branded `CellKey`. Every throw is a named class with a
stable `code` to branch on, never a message string to match. Bad user data
degrades — one cell reads `#VALUE!` — while a programmer error throws
immediately and says what to fix.

### Sheet-scoped defined names

A name Excel scoped to one sheet imports onto that sheet, shadows a workbook
name spelled the same, and exports again with its `localSheetId`. They used to
be dropped: merging them into the workbook's names would have let one sheet's
definition win everywhere, so skipping was the safe half of a choice with no
good half.

### Known limitations

Documented individually, each with the extension point where the work would
start: [docs/LIMITATIONS.md](docs/LIMITATIONS.md).
