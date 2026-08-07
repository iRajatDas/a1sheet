# Limitations

Every item here is an intentional scope cut inherited from the `ref/` POC, not a
bug. Each names the extension point where the work would start.

## Formulas

**A date typed as text stays text.** Only a number is a date; entering
`2024-08-16` in a cell stores the string, and `YEAR` of it is `1905` — the text
coerces to the leading `2024`, which is a serial in 1905. Excel and Sheets parse
the entry against the locale's date formats and store the serial.
→ the entry path in `src/react/useSpreadsheet.ts`; a parser would sit beside
`src/serial.ts` and produce a serial before the value reaches the model.

**Only whole-day arithmetic is Excel-exact.** A serial carries the time of day as
a fraction, so `NOW()` differences accumulate floating-point error at the second
level — visible if you format a difference as `[ss]`. Excel has the same
representation and the same error; the two are not guaranteed to round the same
way.
→ `src/serial.ts`.

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

**A spill is not seen by another spill's tail.** The index of where array
formulas land is built in one pass, and while it is being built an empty cell
reads as empty rather than recursing. So a formula spilling onto cells that a
SECOND formula spills onto in turn sees blanks. Chained spills through an
anchor work; chained spills through a spilled cell do not.
→ `buildSpills` in `src/formula/evaluate.ts` would need to iterate to a fixed
point rather than making a single pass.

**Cross-sheet spilling is not indexed.** A qualified reference reads another
sheet's cells and formulas, but not the cells another sheet's array formula
spills onto — there is one spill index, for the sheet the evaluator was built
for.
→ `spilledInto` in `src/formula/evaluate.ts`; it would need an index per sheet.

**Defined names are workbook-scoped only.** A name Excel scoped to one sheet
(`localSheetId`) is skipped on import rather than imported as a global one,
which would let one sheet's definition win everywhere.
→ `parseDefinedNames` in `src/io/xlsx/read.ts`, and `NamedRanges` would need to
be keyed by sheet as well as name.

**Approximate lookups scan rather than bisect.** `VLOOKUP` and `MATCH` default to
approximate matching, as Excel does, and find the nearest value on the requested
side by scanning. Excel binary-searches, which is faster and gives undefined
results on unsorted data; scanning is slower and gives the right answer. On a
sorted vector the two agree.
→ `findMatch` in `src/formula/functions/lookup.ts`.

**`INDIRECT` is not volatile.** It builds a reference from text, but the result
is memoized like any other formula, so a change to the text it was built from
recalculates it while a change to a cell it *now* points at does not until
something else invalidates the evaluator. In practice an edit rebuilds the
evaluator, so this shows only within a single render.
→ `compute` in `src/formula/evaluate.ts` would need a set of cells excluded from
the cache.

**No `RAND`, `RANDBETWEEN`, or `NOW`-driven recalculation.** `NOW` and `TODAY`
exist but are evaluated once per evaluator, so they do not tick.
→ The same volatility mechanism as `INDIRECT`.

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

**Frozen rows are assumed never hidden.** They render as a separate always-on
band outside the virtualized mapping.
→ `src/react/useRowWindow.ts`.

## Editing

**The fill handle acts on the primary range only.** Clearing, formatting,
copying, and the stats all cover every range of a Ctrl+click selection; dragging
the handle fills from the active one. Excel does the same.
→ `useFillHandle` in `src/react/useFillHandle.ts`.

**Ctrl+clicking a selected range removes the whole range, not the cell.** Excel
splits the range around the cell you click, turning one range into up to four.
Dropping the range is the predictable half of that.
→ `removeRangeAt` in `src/react/useSelection.ts`.

**Paste aligns from the target's top-left corner with no shape validation**
against the source range.
→ `src/react/useClipboard.ts`.

**Internal-paste detection is a text comparison.** Copying identical text from
another application between an internal copy and paste is misread as internal, so
relative refs shift when they should not.
→ `lastCopied` in `src/react/useClipboard.ts`.

## Formatting

**A rotated cell rotates its whole box, not just its text.** `transform` turns
the element, so a rotated cell's borders and fill turn with it and it may overlap
its neighbours. Excel rotates only the text within an upright box.
→ `cellCss` in `src/react/cellStyle.ts` would need the text in an inner element,
which means every cell gaining a wrapper.

**Stacked-vertical text (`textRotation="255"`) is read as no rotation.** It is a
layout mode rather than an angle — one character per line — and has no CSS
equivalent short of `writing-mode`, which would also turn the box.
→ `parseFont`'s alignment branch in `src/io/xlsx/styles.ts`.

**Built-in table style recipes are approximated.** `TableStyleMedium4` resolves to
the right theme accent, and the header is filled with it while the body is striped
with a wash of it. Excel's actual definitions are several hundred entries in its
own resources, differing in border weight and stripe opacity per family.
→ `builtinRecipe` in `src/io/xlsx/tables.ts`.

**Icon sets are drawn from a shape family, not Excel's own glyphs.** The twenty
or so named sets are grouped by meaning — circle, arrow, flag, triangle, star —
and coloured from one three-colour ramp, because what an icon set conveys is
which band a value fell into.
→ `CondIcon` in `src/react/components/CondIcon.tsx`.

**Only `list` data validation is enforced in the UI.** The numeric, date, text
length, and custom kinds are read, written, and exposed on `Sheet.validations`,
but nothing rejects a value that breaks them.
→ `setCell` in `src/react/useSpreadsheet.ts` would test the rule before writing,
and needs somewhere to report the rejection.

**Conditional formats are re-evaluated on every render** of the cells they cover.
Each rule's formula is memoized per evaluator, so a rule over a thousand cells
with absolute references costs one evaluation — but a rule with relative
references costs one per cell.
→ `condStyleFor` in `src/react/useSpreadsheet.ts`.

## File I/O

**The DEFLATE encoder uses fixed Huffman tables, not dynamic ones.** They are
defined by the format and need no header, and on XML they cost a few percent over
optimal tables. `MAX_CHAIN` bounds the match search, trading a little more size
for speed.
→ `deflateRaw` in `src/io/zip/deflate.ts`.

**XLSX import ignores charts and pivot tables** rather than erroring on them. It trusts well-formed output from Excel, Sheets, and LibreOffice
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

**An exported table is written with a neutral style.** The range, the column
names, and the header row survive, but the built-in style name does not: a
table's appearance is flattened onto its cells at import, so re-declaring a style
would paint it twice. The cells keep their colours.
→ `tableXml` in `src/io/xlsx/writeParts.ts`.

**An exported workbook has no theme.** Colours are written as literal RGB, since
`ThemePalette` is resolved on the way in and not kept. A file we wrote therefore
does not follow a reader's theme the way one Excel wrote does.
→ `src/io/xlsx/writeStyles.ts` would need to emit `xl/theme/theme1.xml` and
`Sheet` a place to keep the palette.

**Exported images embed as they were read.** An image that fell back to its URL
on import is written as a URL, not fetched and embedded.
→ `imageParts` in `src/io/xlsx/writeParts.ts`.

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
