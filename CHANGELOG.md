# Changelog

Notable user-visible changes. This project follows [semver](https://semver.org)
honestly: breaking means major.

Started at the cancellation work; earlier commits predate the file and are not
backfilled.

## Unreleased

### Added (dynamic arrays)

The formula engine now has a value kind it did not have: an array. That is what a
modern Excel workbook is made of, and without it such a file was not partly
readable — it was a grid of `#NAME?` with its numbers thrown away.

- **Arrays are values.** A range evaluates to a matrix, always two-dimensional
  because Excel has no one-dimensional array. Collapsing a column to a flat list
  is what made `TRANSPOSE`, `HSTACK`, and `INDEX(t, 0, 2)` impossible to express.
- **Every operator broadcasts.** `(teams = team) * played` over two columns is a
  column of products; it used to take the top-left of each side and return one
  number where the formula meant many — an answer, and the wrong one. A 1x1
  stretches to anything, a row stretches down, a column stretches across, and a
  mismatch pads with `#N/A`.
- **Formulas spill.** An array formula fills the cells below and beside its
  anchor. A cell already holding content blocks it and the anchor reports
  `#SPILL!` rather than overwriting. `Sheet.spillRanges` carries the region an
  imported array formula declares — Excel writes a dynamic array's result into
  the sheet as ordinary values so other readers can see it, and without knowing
  the region a formula's own output looks like an obstruction to its spill.
- **`LET` and `LAMBDA`,** with the functions that consume a lambda: `MAP`,
  `MAKEARRAY`, `BYROW`, `BYCOL`, `REDUCE`, `SCAN`. A lambda closes over the scope
  it was defined in, so one returned from a `LET` still sees that `LET`'s
  bindings.
- **The array library**: `SEQUENCE`, `UNIQUE`, `SORT`, `SORTBY`, `FILTER`,
  `HSTACK`, `VSTACK`, `TRANSPOSE`, `CHOOSECOLS`, `CHOOSEROWS`, `TAKE`, `DROP`,
  `TOROW`, `TOCOL`, `ROWS`, `COLUMNS`, `SUMPRODUCT`, `TEXTSPLIT`.
- **The rest of the modern set**: `XLOOKUP`, `XMATCH`, `LOOKUP`, `HLOOKUP`,
  `IFS`, `SWITCH`, `IFERROR`, `IFNA`, `TEXTJOIN`, `TEXT`, `SUBSTITUTE`,
  `REPLACE`, `FIND`, `SEARCH`, `REPT`, `VALUE`, `CHAR`, `CODE`, `PROPER`,
  `COUNTIF`, `SUMIF`, `AVERAGEIF`, `COUNTBLANK`, `ISBLANK`, `ISERROR`, `ISNA`,
  `ISNUMBER`, `ISTEXT`, `ISLOGICAL`, `NA`, `MEDIAN`, `STDEV`, `PRODUCT`, `MOD`,
  `POWER`, `SQRT`, `SIGN`, `INT`, `TRUNC`, `CEILING`, `FLOOR`, `EXP`, `LN`,
  `LOG10`, `PI`, `XOR`. About 30 functions became about 90.
- **Array literals** (`{1,2;3,4}`), the `&` concatenation operator, and the `%`
  postfix operator, none of which the tokenizer previously recognized.

### Added (references)

- **Structured references** — `tblMatches[home_goal]`,
  `tbl[[#This Row],[a]:[b]]`, `tbl[@col]`, `tbl[#Headers]`, `tbl[#All]`.
  `Sheet.tables` carries the definitions the reader already had; a table's
  styling was being applied while its formulas were not.
- **Cross-sheet references** — `Sheet2!A1` and `'My Sheet'!A1:B9`. Unquoted, a
  name with a space used to lex as two names and a ref, which made such a formula
  silently wrong rather than merely refused.
- **Defined names that hold a formula.** `Workbook.namedFormulas`, alongside
  `namedRanges` — a modern workbook names a computed table this way, and
  `<definedNames>` was not read at all.
- **`Evaluator.getCellValue` and `evaluateArray`** return an array result whole
  rather than collapsed to its top-left.

### Changed (formula engine, breaking)

- **`VLOOKUP` and `MATCH` approximate by default,** as Excel does — the fourth
  argument of `VLOOKUP` and the third of `MATCH` default to "or the next
  smaller". Previously both matched exactly. A well-known footgun, and adopted
  anyway: an imported formula has to compute what Excel computes. Pass `FALSE` or
  `0` for the old behaviour.
- **A numeric cell evaluates to a number.** `=A1` on a cell reading "0.75" gave
  the STRING "0.75"; anything non-numeric gave 0. The displayed text is
  unchanged — "1.50" still shows as "1.50".
- **`AND`, `OR`, `NOT`, and `XOR` return booleans,** not 1 and 0. The grid renders
  a boolean as TRUE/FALSE, which is what Excel shows.
- **Errors propagate through operators.** `IF(COUNTBLANK(bad)=0, 1, 0)` used to
  compare `#NAME?` against 0, find them unequal, and return 0 — a guard turning an
  error into a plausible answer.
- **`IF`, `IFS`, `SWITCH`, `IFERROR`, and `IFNA` are lazy.** The branch not taken
  is not evaluated, so a guard against division by zero now actually guards.
- **`createEvaluator`'s third argument is an options object.** It briefly
  accepted either that or a bare cached-value map, told apart by its keys; that
  cleverness silently mistook `{ tables }` for cached values. Pass
  `{ cachedValues }`.
- **`FormulaValue` no longer includes arrays**; `FormulaArg` is
  `FormulaValue | Matrix | LambdaValue`. A custom function receives matrices, not
  flat lists.
- **`SHAPE_SENSITIVE` is gone.** VLOOKUP, INDEX, and MATCH needed the raw range
  AST node to recover a table's 2D shape; a range is a matrix now, so they are
  ordinary functions — and work on a computed table, which the special case could
  not.

### Fixed

- **A flaky test that raced the clock.** "A large read hands the thread back"
  scheduled a timer and read a big CSV, so it proved something only while the
  machine was slow enough for the read to overrun a frame. It is now driven by an
  injected clock and asserts on ordering, since the yielder uses a MessageChannel
  which the event loop runs ahead of a `setTimeout(0)` anyway.

### Added (XLSX import fidelity)

An imported workbook used to read as a plainer version of itself. None of the
reasons were in the cells.

- **Theme colours resolve.** Excel writes most colours as
  `<color theme="4" tint="-0.25"/>`, against the palette in
  `xl/theme/theme1.xml`; the reader looked only at `rgb`, so almost every colour
  in a file Excel wrote silently became no colour at all. This was the single
  biggest cause. The legacy `indexed` palette resolves too, and `auto` means
  inherit rather than black — defaulting to black is how a missing border colour
  becomes a black grid.
- **Borders.** `StyleObject.borders`, thirteen OOXML line kinds mapped onto CSS,
  read and written. An edge Excel marks `auto` takes the cell's text colour.
- **Font family and size.** `StyleObject.fontFamily` and `fontSize` (pixels). A
  family from a file keeps the theme's stack behind it, so a reader without Aptos
  Narrow falls back to the sheet's own face rather than the browser's.
- **Vertical alignment, text wrapping, and gradient fills.**
  `StyleObject.valign`, `wrap`, and `gradient` — the last read as stops and drawn
  as a CSS `linear-gradient`.
- **Named cell styles are inherited.** A `<cellXfs>` entry can take its font,
  fill, border, and alignment from a named style through its `xfId`, and Excel
  usually omits the `applyFont`/`applyBorder` flags meant to signal it. Titles and
  headings carry their borders and centring that way.
- **`StyleObject.numFmtCode`** holds the file's literal format code and is
  preferred over the six-value `numFmt` bucket, which stays for the format
  dropdown. A new renderer covers sections, literal signs, decimals, grouping,
  percentages, currency, dates, times, elapsed time, and colour hints — so a goal
  difference formatted `+0;-0;0` shows `+45` instead of `45`, and a kick-off shows
  `8/16/24 20:00` instead of a bare date. Built-in format ids carry no code in the
  file, so those are a table here.
- **Table styling.** `Format as Table` keeps a table's appearance in
  `xl/tables/*.xml`, not in its cells. Both mechanisms are read: the differential
  formats the file states, and the built-in style names it does not —
  `TableStyleMedium4` means the third theme accent, a colour that appears nowhere
  in the workbook that shows it. Applied under the cells' own formatting, which
  stays more specific.
- **Conditional formatting.** `Sheet.condFormats`, evaluated at render time rather
  than flattened, so a rule keeps reacting to edits. `expression`, `cellIs`,
  `containsText`, and `containsBlanks` rules, with priority order and
  `stopIfTrue`; an expression's formula is shifted per cell the way a copied one
  is. Colour scales, data bars, and icon sets are read and dropped. Worth having
  even to display a static file: the sample workbook styles its headings this way,
  not with cell formatting.
- **In-cell images.** `Sheet.images` and `CellImage`. `=IMAGE("…")` stores
  `#VALUE!` in the cell and the picture five parts away, so the cell used to
  import as the error that is literally in it. The embedded copy is preferred over
  the source URL — no request, works offline, and cannot be changed by whoever
  owns the address — with one data URI shared per media part, a 16 MiB budget per
  workbook, and a raster-only allow-list. Editing such a cell removes the image;
  typing an `IMAGE()` call adds one.
- **`IMAGE` is a formula function,** returning its URL, so a cell whose picture
  cannot be drawn shows where it was meant to come from.
- **Excel's `_xlfn.` and `_xlws.` prefixes are stripped on import.** They mark
  functions newer than the file format, and left in place they guarantee `#NAME?`
  even for functions this engine does implement.
- **`Evaluator.evaluate(formulaBody)`** evaluates a formula belonging to no cell.
  Conditional-format rules were the first thing to need it.

Two bugs the tests caught rather than I did: `<u val="none"/>` is how a
differential format turns underline *off*, and treating the tag's presence as
truth underlined every table header; and scanning the whole document for
`<numFmt>` picks up the ones inside `<dxfs>`, where a `numFmtId="0"` of "General"
gave every plain cell in the workbook a format code.

Export still writes cells, styles, merges, and sizing — not tables, conditional
formats, or images. See `docs/LIMITATIONS.md`.

### Fixed (XLSX import)

- **Imported dates were seventy years out.** SpreadsheetML counts days from
  1899-12-30 and this engine counts from the Unix epoch, and a serial is only a
  number — nothing about it says which. 2024-08-16 read back as 2094-08-18: a
  plausible date, in the wrong century, which is what kept it invisible. Import
  and export now convert, including Excel's inherited Lotus belief that 1900 was
  a leap year, so the offset is right either side of the day that never existed.
  See `src/io/xlsx/dates.ts`.
- **A formula using anything unimplemented no longer imports as `#NAME?`.** The
  reader was taking each formula's text and discarding the value Excel had
  already computed and stored beside it, so a workbook built on dynamic arrays,
  `LET`, `LAMBDA`, or structured table references arrived as a grid of errors
  with its numbers thrown away. Those values are kept now, in a new
  `Sheet.cachedValues`, and shown whenever evaluation fails.

  A displayed cached value is a snapshot: editing a cell it depends on does not
  update it, because nothing here can recalculate a formula it could not parse.
  Editing the formula cell itself drops the entry and the real error appears. A
  formula the engine *can* evaluate always wins over the import, or an imported
  sheet would be frozen.
- **A formula whose result is text exports as text.** `writeXlsx` wrote every
  formula's cached value as a bare number, so Excel read `"POS"` as `0`.

### Added (XLSX import)

- **`Sheet.cachedValues`**, `XlsxSheetData.cachedValues`, and the optional
  `XlsxSheetInput.cachedValues`. Pass the last of these when exporting to keep an
  unevaluable formula's `<v>` truthful.
- **`createEvaluator` takes an optional third argument,** the cached values to
  fall back on. Existing two-argument calls are unaffected.

### Changed (appearance)

- **Cells are set in `fontFamily`, not `monoFontFamily`.** A sheet of names and
  totals in a monospace face reads as a terminal rather than as a spreadsheet.
  The mono font now appears only in the formula input, where column alignment
  inside an expression carries meaning. Both theme keys are unchanged; if you
  want the old look, set `fontFamily` to a monospace stack.

- **The controls are drawn with icons instead of typographic glyphs and one
  emoji.** `↶ ↷ ⯇ ≡ ⯈ ▾ − 🔒` were resolving through the host page's font stack,
  so they changed shape, weight, and baseline from machine to machine, and
  several have no coverage in the common UI fonts at all — they landed as a
  tofu box. The emoji was worse: a colour bitmap on most platforms, so it
  ignored `color` and the lock stayed black when its button went active teal.
  They are now inline SVG from [Tabler Icons](https://tabler.io/icons) (MIT),
  copied into the source rather than installed — `dependencies` stays empty.
  See `packages/a1sheet/THIRD-PARTY-NOTICES.md`. A test renders the preset and
  fails on any decorative glyph in the output, so this cannot quietly come back.
- **Toolbar buttons are icon-only** and carry a `title` and a matching
  `aria-label`. If you were finding them by text — `getByText("+Row")`,
  `getByText("Freeze")` — use the label: "Insert row", "Delete row",
  "Insert column", "Delete column", "Merge cells", "Unmerge cells", "Freeze up
  through the selection", "Unfreeze". `Sheet.FileMenu` keeps its text, since
  "CSV" and "XLSX" are not drawable.
- **The grid draws its own scrollbars, one per axis, in channels of their own.**
  macOS and mobile default to overlay scrollbars: they fade in over the content
  while you scroll, so they cover the rightmost column and the bottom row
  exactly while you are moving through them, and taking no layout space they
  leave the grid nothing to lay itself out around. Styling the native ones does
  not fix it — `::-webkit-scrollbar` is non-standard and Firefox ignores it,
  `scrollbar-width` offers only `thin` and `auto`, and neither engine can be
  told "always visible". The scroll container now hides its native bars and two
  identical ones sit beside it, always present, as in Sheets and Excel.
  Scrolling is still the browser's — the bars write `scrollTop`/`scrollLeft`
  and follow the resulting `scroll` event — so the wheel, trackpad, keyboard,
  and `scrollIntoView` are untouched. Each bar is a `role="scrollbar"` with
  `aria-controls` pointing at the container. Dragging the thumb and clicking
  the track to page both work.
- **`Theme` gained `scrollbarTrack`, `scrollbarThumb`, and
  `scrollbarThumbHover`.** A custom dark theme should set them — the channel is
  a real surface, and left at the defaults it is a light stripe down the
  right-hand edge. `Partial<Theme>` still fills the rest in, so nothing breaks.
- **`useSpreadsheet` returns `scrollTop`, `scrollLeft`, `viewportHeight`, and
  `viewportWidth`.** The setters were already public; the values were not, so
  nothing outside the hook could describe the window into the sheet. The
  grid's own scrollbars were the first thing to need them.

### Changed (breaking: file I/O is a primitive)

- **`Sheet.Toolbar` no longer takes `onImport`, `onExportCsv`, or
  `onExportXlsx`.** Render `<Sheet.FileMenu />` inside it instead. Those
  callbacks were only ever passed by the `<Spreadsheet />` preset, so a
  hand-composed toolbar rendered no Import button and the feature looked as
  though it had been removed — a part of the library depending on the preset to
  work is exactly what composition is supposed to prevent. Omitting
  `Sheet.FileMenu` still keeps the XLSX writer and ZIP reader out of the bundle,
  which is the property those callbacks existed to protect.
- **`Sheet.Toolbar` takes `children`,** rendered after a separator. Your own
  buttons go there, using the `a1s-btn` class.
- **`Sheet.Grid` takes `children`,** rendered inside the scroll container after
  the last row — an end-of-sheet slot that scrolls with the content.

### Added (grid)

- **`Sheet.AddRows`** — "Add N more rows at the bottom", where Google Sheets
  puts it: at the end of the scrollable content, sticky to the left edge so it
  stays readable when the sheet is scrolled sideways. Backed by
  `api.appendRows(count)`, which is undoable.

### Added (repository)

- **A Storybook** at `examples/storybook`, run with `bun run storybook` — 12
  docs pages and 49 stories, meant as the documentation rather than a gallery.
  Prose pages for the quick start and for composition-over-configuration; a
  props table on the preset; interactive stories for controlled and uncontrolled
  state, the imperative handle, the headless hooks, every feature, import with
  progress and cancellation, 100k rows, theming, and a recipes section answering
  the questions the API deliberately has no prop for (read-only, validation,
  search). Stories alias to `src`, so there is no build step between an edit and
  the browser, and `bun test` renders every story and runs its play function.
- **The grid reads its container's scroll position on mount.** A browser
  restoring scroll on back-navigation, a bfcache restore, or a consumer
  scrolling the container before the grid mounts all set `scrollTop` without an
  `onScroll` to hear, so virtualization went on drawing the top of the sheet
  into a container showing the middle of it — rows in the wrong place, blank
  space where content should be.
- **`.a1s-scroller`** on the grid's scroll container. Scrolling the sheet from
  outside means moving that element; `api.setScrollTop` alone only tells
  virtualization where to draw, and the two would fall out of step.

### Added (XLSX sizing)

- **Row heights and column widths survive an XLSX round trip.** `readXlsx`
  reads `<col width>` and `<row ht>`, and `writeXlsx` emits both;
  `XlsxSheetData` and `XlsxSheetInput` gained `colWidths` and `rowHeights`, in
  pixels. Only entries the file marks `customWidth`/`customHeight` are read, so
  an Excel-laid-out sheet does not import with every column pinned to a size
  nobody chose.

### Changed (filtering)

- **An active filter no longer rescans the sheet on every edit.** The verdict is
  cached and only rows whose raw content in a filtered column actually changed
  are re-tested — about 60 ms per committed edit at 100k rows. A filtered column
  containing a formula still rescans in full, because a formula's display can
  change while its raw text does not; that case is detected, not assumed.

### Added (grid performance and sizing)

- **Columns are virtualized.** Only the columns near the viewport exist in the
  DOM, as rows already did. A 500-column sheet rendered 6,000 cells and cost
  1,351 ms to mount and 73 ms per keystroke; it now renders the same ~290 cells
  as a 26-column sheet and mounts in ~30 ms. Column track definitions stay
  explicit, so widths, freeze offsets, and the scroll extent are unaffected.
- **The scrollbar describes the sheet, not the window into it.** Both extents
  were previously bounded by whatever happened to be rendered — 338 px of a
  2,600,000 px sheet at 100k rows, and a horizontal extent that followed the
  window rightward. Vertically the browser stops allocating implicit tracks at
  the last drawn row; horizontally the grid is block-level and the column tracks
  overflow it. `minHeight`/`minWidth` fix each.
- **Rows are resizable, and heights may vary.** `Sheet.rowHeights` mirrors
  `colWidths`. Drag the divider below a row header to resize; double-click it to
  return the row to the default. Virtualization places rows through a cumulative
  offset table with a binary search, and a single spacer item stands in for the
  rows above the window — while no row has been resized the table is skipped
  entirely, so an unresized sheet pays nothing.
- **Double-clicking a column divider auto-fits the column** to its widest value,
  measured with canvas `measureText`. Capped at `AUTOFIT_SAMPLE_LIMIT` (2,000)
  cells and `MAX_AUTOFIT_COL_WIDTH` (600 px) so a double-click cannot stall or
  size a column off the screen.
- `setRowHeight`, `resetRowHeight`, and `resetColWidth` on the sheet API;
  `setScrollLeft` and `setViewportWidth` for wiring a custom scroll container.
- Cells and headers carry `data-row`/`data-col`. With both axes virtualized,
  position in the DOM no longer identifies a cell; this does.

### Fixed (model)

- **Row and column metadata now moves with an insert or delete.** `rowHeights`,
  `rowLabels`, `hiddenRows`, `colWidths`, `colLabels`, and `filters` are keyed by
  a bare index and were left in place while the cells around them shifted, so
  inserting a row above a hidden one hid the wrong row and a column filter
  applied to the wrong column.

### Changed (grid performance and sizing)

- The filter scan resolves an empty cell's display once per column instead of
  once per row, which is most of the cost on a sparse sheet.
- `rectFor` and the fill-drag hit test use binary searches over the offset
  tables rather than `indexOf` and a running sum from column zero.
- A zero viewport measurement is ignored rather than believed: a grid inside a
  `display: none` container no longer windows itself down to nothing and stay
  empty after it becomes visible.

### Added

- **Reference picking while typing a formula.** Clicking a cell mid-formula
  writes its reference at the caret instead of moving the selection; dragging
  grows that reference into a range; every reference is outlined in the grid and
  coloured by group, so repeated references match. A click past a finished
  operand still selects, as in Sheets. `findRefSpans`, `insertRefAtCaret`, and
  `isFormulaSource` are exported for custom editors, along with the
  `useFormulaRefs` and `useCaretBinding` hooks.
- **`explainErrorValue`** turns an error sentinel into a sentence saying what to
  do about it. The status bar shows it for the active cell.
- `Theme.refColors`, the palette for reference outlines.
- `EditingState.caret`, so both editors and the grid agree on where the caret is.

### Fixed (grid interaction)

- **Clicking a cell no longer kills the keyboard.** A cell is a plain `<div>`,
  so mousedown moved focus to `<body>` — and every shortcut, the clipboard
  handlers, and "select a cell and start typing" live on the hidden textarea.
  The grid only worked if you never clicked it.
- **Dragging across cells no longer selects their text.** Same unsuppressed
  mousedown default.
- **The active cell is the anchor, not the drag end.** Drag D5→F13 and D5 stays
  active and untinted, as Excel and Sheets do. Arrow keys step from it, typing
  lands in it, and the formula bar no longer flickers through every cell a drag
  passes over.
- Row and column headers highlight across the selected range.
- Translucent tints are overlays rather than `background`, which had left sticky
  headers and frozen rows see-through.
- The fill handle sits at the corner of the selection regardless of the anchor.

### Added (file reading)

- **Cancellation and progress for file reads.** `readWorkbookFile`, `readXlsx`,
  and `csvToCells` take an optional options object with `signal` and
  `onProgress`. Reads yield to the event loop between chunks, so a large file no
  longer freezes the tab. `onProgress` is rate-limited to about once per frame
  and once per 1% of progress, `ratio` never decreases, and a successful read's
  final report is always `1`. Aborting rejects with `AbortedError`
  (`code: "ABORTED"`) and applies nothing — a cancelled import cannot
  half-replace your workbook. Pacing is unconditional and costs throughput: on a
  37 MB, 600k-cell workbook, ~1.1 s blocking becomes ~1.7 s paced, with no block
  longer than ~60 ms. See the README for the measurements.
- **Typed errors are exported.** `A1SheetError`, `isA1SheetError`, `ERROR_CODES`,
  and the individual classes are now part of the public API. They existed but
  were unreachable, so `code`-based branching was impossible.
- `iterCsvRows`, a lazy row-at-a-time CSV scanner. `parseCSV` is now a thin
  collecting wrapper over it.
- `READ_PHASES`, and the `ReadProgress`, `ReadPhase`, and `AsyncReadOptions`
  types.

### Changed

- **Breaking: `csvToCells` returns a `Promise`.** It is paced now, and pacing
  requires being async. `parseCSV` remains synchronous for clipboard-sized text.
- `readXlsx` throws `UnsupportedFormatError`, `MalformedFileError`, and
  `NotAZipError` instead of bare `Error`s. Messages changed; if you were matching
  on message text, switch to `error.code`.
- The `<Spreadsheet />` preset reports import progress in the status bar, and
  cancels an in-flight import when a second file is chosen or the component
  unmounts.

### Fixed

- A CSV import no longer copies the entire input to normalize CRLF before
  parsing. Line endings inside quoted fields are normalized inline instead —
  same result, one fewer full-size string allocation.
- `findElement` stops at the first match rather than scanning the whole fragment.
  It runs three times per cell, so this was real waste at scale.
