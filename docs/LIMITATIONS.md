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
will not match Excel's serial for the same date.
→ `src/formula/values.ts` (`DAY_MS`) and `src/formula/functions/date.ts`.

**`IF` is not lazy.** Both branches are evaluated before dispatch. No correctness
impact (there are no side effects), but the unused branch still computes.
→ `src/formula/functions/logic.ts`.

**Lookups are exact-match only.** No approximate/sorted mode for `VLOOKUP` or
`MATCH`.
→ the shape-sensitive branch of `src/formula/evaluate.ts`.

## Grid

**Fixed row height** (`ROW_HEIGHT = 26`). Only column width is resizable. Both
the sticky `top` offsets for frozen rows and the virtualization index math assume
uniform height.
→ `src/react/useRowWindow.ts` would need a cumulative offset table, and
`Grid`'s `stickyStyleFor` would need to consume it.

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

**No borders.** Bold, italic, underline, color, background, and number format are
supported.
→ Add fields to `StyleObject`, render them in `Cell`, and add a `<borders>`
section to `src/io/xlsx/styles.ts`, which currently writes a single empty
`<border/>` placeholder.

**No conditional formatting.**
→ A new per-sheet `condFormats: [{range, rule, style}]`, resolved in `Cell` after
the base style and before the inline style fields are read.

**No data validation dropdowns.**

## File I/O

**ZIP export is STORE-only.** There is no DEFLATE *encoder*, only a decoder.
Exported `.xlsx` files are valid and open everywhere; they are just larger than
necessary.
→ Implement `src/io/zip/deflate.ts`, then switch the `method` field in
`src/io/zip/zip.ts` from 0 to 8 for the XML entries. Fixed-Huffman is sufficient.

**XLSX import ignores charts, pivot tables, data validation, and conditional
formatting** rather than erroring on them. It trusts well-formed output from
Excel, Sheets, and LibreOffice and does not handle every OOXML edge case.
→ `src/io/xlsx/read.ts`.

**`numFmt` round-tripping is heuristic.** Arbitrary OOXML format codes collapse
into the nearest of our six enum values rather than surviving exactly.
→ `numFmtToKey` in `src/io/xlsx/styles.ts`.

**Merged-cell-specific styling is not read.** Merges themselves round-trip.

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
