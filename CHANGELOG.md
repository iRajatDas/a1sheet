# Changelog

Notable user-visible changes. This project follows [semver](https://semver.org)
honestly: breaking means major.

Started at the cancellation work; earlier commits predate the file and are not
backfilled.

## Unreleased

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
