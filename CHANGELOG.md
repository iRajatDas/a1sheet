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

### Types and errors

`strict` with `noUncheckedIndexedAccess`, no `any`, `readonly` on what consumers
should not mutate, and branded `CellKey`. Every throw is a named class with a
stable `code` to branch on, never a message string to match. Bad user data
degrades — one cell reads `#VALUE!` — while a programmer error throws
immediately and says what to fix.

### Known limitations

Documented individually, each with the extension point where the work would
start: [docs/LIMITATIONS.md](docs/LIMITATIONS.md).
