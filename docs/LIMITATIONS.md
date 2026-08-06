# Limitations

Every item here is an intentional scope cut inherited from the `ref/` POC, not a
bug. Each names the extension point where the work would start.

## Formulas

**No cross-sheet references (`Sheet2!A1`).** Formulas resolve against the active
sheet only.
→ `src/formula/tokenize.ts` needs a `!`-anchored lookahead before the ref
pattern. A bare regex is not enough: sheet names ending in letters followed by
digits collide with the ref pattern. `createEvaluator` then moves from per-sheet
to per-workbook so `EvalContext` can expose `getSheet(name)`.

**Date serials use the Unix epoch, not Excel's 1899-12-30.** Internally
consistent and arithmetic-friendly, but a serial number copied out of a1sheet
will not match Excel's serial for the same date. XLSX import and export convert
between the two, so a file is unaffected; only a serial you read off the screen
and paste into Excel by hand will disagree.
→ `src/formula/values.ts` (`DAY_MS`) and `src/formula/functions/date.ts`; the
conversion is `src/io/xlsx/dates.ts`.

**A formula this engine cannot evaluate displays the value it was imported
with, and that value is never recalculated.** `Sheet.cachedValues` holds what
Excel last computed for each formula cell, and it is shown whenever evaluation
fails — otherwise a workbook using dynamic arrays, `LET`, `LAMBDA`, or structured
table references imports as a grid of `#NAME?`. The consequence is that such a
cell is effectively read-only: editing a cell it depends on cannot change it,
because nothing here can recalculate a formula it could not parse. Editing the
formula cell itself drops its imported value, and the error appears.
→ `orImported` in `src/formula/evaluate.ts`. Removing the caveat means
implementing the functions, not changing the fallback.

**Structured table references (`tblMatches[column]`) do not evaluate.** The table
definitions in `xl/tables/*.xml` ARE read — a table's styling is applied — but the
reference syntax is not, so a formula using one falls back to its imported value.
→ `src/formula/tokenize.ts` for the syntax; `src/io/xlsx/tables.ts` already has
the range and column names to resolve against.

**No dynamic arrays.** A formula returning a range does not spill into its
neighbours, and `LET`, `LAMBDA`, `MAP`, `MAKEARRAY`, `HSTACK`, `VSTACK`,
`SEQUENCE`, `UNIQUE`, `SORT`, `CHOOSECOLS`, and `XLOOKUP` are not implemented. An
imported workbook built on them displays, from its cached values, but does not
recalculate.
→ `FormulaValue` in `src/formula/values.ts` would need a 2D array kind, and the
sheet a map of spilled ranges so a spill can be invalidated as a unit.

**`IF` is not lazy.** Both branches are evaluated before dispatch. No correctness
impact (there are no side effects), but the unused branch still computes.
→ `src/formula/functions/logic.ts`.

**Lookups are exact-match only.** No approximate/sorted mode for `VLOOKUP` or
`MATCH`.
→ the shape-sensitive branch of `src/formula/evaluate.ts`.

## Grid

**A wrapped cell does not grow its row.** `wrap` on a style makes text run onto
more lines, but the row keeps whatever height it has, so the extra lines are
clipped — and auto-fitting a row still means returning it to the default height
rather than measuring. Row heights are otherwise free, and the offset tables
handle whatever you set.
→ `autoFitRow` in `Grid` needs to measure wrapped text; automatic growth would
mean the row window asking the DOM for a height it currently computes.

**Auto-fitting a column samples at most `AUTOFIT_SAMPLE_LIMIT` cells** (2,000)
and stops. A longer value further down the column stays clipped. The cap exists
because a double-click must not stall on a hundred thousand measurements.
→ `autoFitCol` in `src/react/components/Grid.tsx`.

**XLSX sizing is approximate, and only what the file marks as custom.** Widths
are stored in multiples of the widest digit of the normal font, assumed to be
Calibri 11; a workbook using another font reads a few pixels off. A `<col>` or
`ht=` without `customWidth`/`customHeight` is ignored, so a file laid out by
Excel does not arrive with every column pinned.
→ `src/io/xlsx/units.ts`.

**A filter over a column containing a formula rescans every row on every edit.**
A formula's displayed value can change while its raw text does not, so the
incremental cache cannot see that the row needs re-testing and the column falls
back to a full scan — about 60 ms per edit at 100k rows. Columns of plain values
are incremental.
→ `volatile` in `src/react/useFilterHidden.ts`; needs the evaluator to report
which cells its last recalculation changed.

**Committing an edit copies the whole cell map.** `useWorkbook` clones on write,
so edit cost grows with the number of filled cells — about 26 ms at 10k, about
390 ms at 1M. Virtualization does not help: this is the write path, and it is
now the largest cost in the grid by a wide margin.
→ `cloneSheet` in `src/model/sheet.ts` and `updateSheet` in
`src/react/useWorkbook.ts` would need structural sharing rather than a spread.

**No column hiding.** Rows can be hidden; columns cannot.
→ Mirror the `visibleRows` compaction for columns, and make it interact with
`frozenCols` the way `effectiveHiddenRows` interacts with `frozenRows`.

**Frozen rows are assumed never hidden.** They render as a separate always-on
band outside the virtualized mapping.
→ `src/react/useRowWindow.ts`.

## Editing

**Fill handle extends down and right only.** Dragging up or left is a no-op.
→ `commit` in `src/react/useFillHandle.ts`.

**Multi-range selection (Ctrl+click) feeds the status bar only.** Copy, fill, and
paste all act on the primary selection.
→ `extraRanges` in `src/react/useSelection.ts`; `StatusBar` is its only consumer.

**Paste aligns from the target's top-left corner with no shape validation**
against the source range.
→ `src/react/useClipboard.ts`.

**Internal-paste detection is a text comparison.** Copying identical text from
another application between an internal copy and paste is misread as internal, so
relative refs shift when they should not.
→ `lastCopied` in `src/react/useClipboard.ts`.

## Formatting

**No indent or text rotation.** A style carries bold, italic, underline, colour,
fill, gradient, borders, font family, font size, horizontal and vertical
alignment, wrapping, and a number-format code.
→ `StyleObject` in `src/model/types.ts`, then `cellCss` in
`src/react/cellStyle.ts` and both directions of `src/io/xlsx/styles.ts`.

**Tint is approximated in RGB, not HSL.** OOXML specifies `tint` against HSL
luminance; the RGB form every other implementation uses differs by a shade or two
on saturated colours and not at all on greys.
→ `applyTint` in `src/io/xlsx/palette.ts`.

**Built-in table style recipes are approximated.** `TableStyleMedium4` resolves to
the right theme accent, and the header is filled with it while the body is striped
with a wash of it. Excel's actual definitions are several hundred entries in its
own resources, differing in border weight and stripe opacity per family.
→ `builtinRecipe` in `src/io/xlsx/tables.ts`.

**Graphical conditional-format rules are dropped.** Colour scales, data bars, and
icon sets are drawings rather than styles, and `top10`/`aboveAverage` need
statistics over the whole range. Expression, `cellIs`, `containsText`, and
`containsBlanks` rules all work.
→ `parseRule` in `src/io/xlsx/condFormat.ts` and `matches` in
`src/format/condFormat.ts`; a data bar also needs something drawn behind the text.

**Conditional formats are re-evaluated on every render** of the cells they cover.
Each rule's formula is memoized per evaluator, so a rule over a thousand cells
with absolute references costs one evaluation — but a rule with relative
references costs one per cell.
→ `condStyleFor` in `src/react/useSpreadsheet.ts`.

**No data validation dropdowns.**

## File I/O

**ZIP export is STORE-only.** There is no DEFLATE *encoder*, only a decoder.
Exported `.xlsx` files are valid and open everywhere; they are just larger than
necessary.
→ Implement `src/io/zip/deflate.ts`, then switch the `method` field in
`src/io/zip/zip.ts` from 0 to 8 for the XML entries. Fixed-Huffman is sufficient.

**XLSX import ignores charts, pivot tables, and data validation** rather than
erroring on them. It trusts well-formed output from Excel, Sheets, and LibreOffice
and does not handle every OOXML edge case.
→ `src/io/xlsx/read.ts`.

**`StyleObject.numFmt` is heuristic; `numFmtCode` is exact.** Arbitrary format
codes collapse into the nearest of six enum values for the format dropdown to
show, while the literal code travels alongside and is what renders.
→ `numFmtToKey` in `src/io/xlsx/styles.ts`.

**Merged-cell-specific styling is not read.** Merges themselves round-trip.

**Only in-cell images are read.** `=IMAGE("…")` resolves to the PNG the file
embeds, or to its source URL. Floating pictures, shapes, and charts anchored over
the grid (`xl/drawings/`) are not read at all.
→ `src/io/xlsx/images.ts` covers the rich-value chain; drawings are a separate
part with their own anchor model.

**Embedded images are capped at 16 MiB of distinct bytes per workbook.** Past that
an image falls back to its source URL, which costs a request instead of memory. An
image with no URL either is not drawn.
→ `EMBED_BUDGET_BYTES` in `src/io/xlsx/images.ts`. A Blob URL per image would lift
the cap at the price of a lifecycle to manage.

**Only raster images are embedded** — PNG, JPEG, GIF, WebP, BMP. Anything else
falls back to its URL. An allow-list rather than a block-list, so a format nobody
has vetted never becomes a `data:` URI.
→ `EMBEDDABLE` in `src/io/xlsx/images.ts`.

**Exporting does not write images, tables, or conditional formats.** They are read
and rendered; `writeXlsx` emits cells, styles, merges, and sizing. A round trip
through export therefore flattens a table to its cell colours and drops the rest.
→ `src/io/xlsx/write.ts` and `src/io/xlsx/writeStyles.ts`.

**Reads are paced, not off-thread.** `readWorkbookFile` yields to the event loop
between chunks and honors an `AbortSignal`, so the tab stays responsive and an
import can be cancelled — but the parse still runs on whichever thread called it.
Pacing costs throughput: on a 37 MB, 600k-cell workbook, ~1.1 s blocking becomes
~1.7 s paced. Only 4 ms of that is the yielding; the rest is garbage collection
the blocking version deferred until it was finished. Pacing is unconditional, so
a Worker caller pays it too, for a responsiveness it does not need.
→ The pacing seam is already the right shape for a Worker: `src/io/progress.ts`
defines the only place a read yields. Move `readXlsx`/`csvToCells` behind a
Worker, post `ReadProgress` back over `postMessage`, and drive the existing
`AbortSignal` from a cancel message. Nothing in the readers changes. Once that
exists, `createPacer` is also the one place to make the frame budget adaptive so
an off-thread read can stop yielding.

**A few steps cannot be interrupted.** Inflating one large ZIP member and
decoding it to a string are single native calls; together they block for ~50 ms
on a 37 MB file regardless of the frame budget, which is why the budget is one
frame rather than React's ~5 ms.
→ Streaming both would mean an incremental inflater in `src/io/zip/inflate.ts`
and a chunked `TextDecoder` with `{ stream: true }` in `src/io/xlsx/read.ts`.

**Progress ratios are estimates.** The denominator comes from counting `<c ` and
`<si` occurrences (xlsx) or newlines (csv) before parsing, none of which is
exact — a `<c/>` with no attributes is missed, a newline inside a quoted CSV
field is over-counted. The pacer clamps progress monotonically into 0..1 and
forces a final `1`, so the bar never reverses or overshoots; it can just move
unevenly.
→ `totalElements` in `src/io/xlsx/read.ts`, `estimatedRows` in
`src/io/csv/read.ts`.

## Fixed relative to the POC

Three defects were found while porting and fixed rather than carried forward.
Each has a regression test.

**`<>` never worked.** `ref/formulaEngine.js:80-86` only ever appended `"="` when
lexing a comparison, so it could not produce the `<>` token. `A1<>B1` lexed as
`cmp "<"` followed by `cmp ">"`, and `parsePrimary` swallowed the `>` as a literal
`0` — silently wrong, never an error, while `evalCompare` had a `"<>"` case waiting.

**VLOOKUP and MATCH always matched the first row for text keys.**
`ref/formulaEngine.js:277` compared with `toNumber(a) === toNumber(b)`, and
`toNumber` coerces non-numeric text to `0` — so any two text keys compared equal.
Numeric comparison now requires both sides to actually parse as numbers.

**Cycle detection only caught direct self-reference.** `ctx.getValue` coerces
through `toNumber`, which flattened the `#CYCLE!` sentinel to `0` one frame up, so
`A1=A2, A2=A1` silently evaluated to `0`. Cycles now raise an internal signal that
unwinds through every frame, and each participating cell reports `#CYCLE!`.

Two further deliberate improvements:

**Errors propagate through function arguments.** `=SUM(UNDEFINED_NAME)` returned
`0` in the POC because `flattenNums` dropped the `#NAME?` string. Error sentinels
are now checked before a function is called.

**No `DOMParser`.** The POC's XLSX reader used `DOMParser`, which exists only in
browsers. `src/io/xlsx/xml.ts` is a small scanner instead, so the framework-agnostic
entrypoint genuinely works in Node and Web Workers — and `readXlsx` is testable
without a DOM.
