# Changelog

Notable user-visible changes. This project follows [semver](https://semver.org)
honestly: breaking means major.

Started at the cancellation work; earlier commits predate the file and are not
backfilled.

## Unreleased

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
