# Changelog

Notable user-visible changes. This project follows [semver](https://semver.org)
honestly: breaking means major.

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
