# Zero-dependency React Spreadsheet — v2

Three files, no npm packages:

- **formulaEngine.js** — tokenizer, parser, evaluator, function library, named ranges, relative-reference shifting.
- **xlsxIO.js** — ZIP reader/writer, a from-scratch DEFLATE decoder, and OOXML read/write (multi-sheet, with basic styles) + CSV.
- **Spreadsheet.jsx** — the component: workbook state, toolbar, formula bar, grid (CSS Grid, virtualized), sheet tabs, context menu, status bar.

Drop the folder into any React project that already handles JSX (Vite, Next, CRA, etc.) and render `<Spreadsheet />`. No build config changes needed beyond what you already have for JSX/ESM.

## What's in this pass

**Editing UX**
- Formula bar with name box (type a ref to jump, an existing name to select it, or a new name + Enter to define a named range from the current selection)
- Fill handle — drag the corner square down or right; numeric/linear series extrapolate, formulas get their relative refs shifted
- Multi-range selection via Ctrl+click (view/status-bar stats only — see limitations)
- Right-click context menu (insert/delete row/col, clear contents/formatting, lock/unlock)
- Freeze rows/cols (toolbar buttons freeze up through the current selection)
- Copy/paste shifts formula references correctly when pasting inside the app; plain values when pasting from/to elsewhere

**Formatting**
- Bold/italic/underline, alignment, text & fill color, number formats (general/integer/0.00/percent/currency/date)
- Merge cells
- Cell locking (`locked` cells can't be typed into or cleared, but remain selectable/copyable)
- Round-trips through XLSX export/import (styles.xml is written and read)

**Structure**
- Multiple sheets — add, rename (double-click tab), delete, switch
- Sort ascending/descending and checkbox-filter any column (▾ on the header)
- Named ranges
- Custom row/column labels — double-click a header to rename it (display only; internal addressing is still A1-style)

**Formulas**
- `+ - * / ^`, comparisons (`= < > <= >= <>`), parentheses
- Refs & ranges with absolute markers: `A1`, `$A$1`, `A$1:B10`
- SUM AVERAGE MIN MAX COUNT COUNTA ABS ROUND IF AND OR NOT
- Text: CONCAT/CONCATENATE LEFT RIGHT MID TRIM UPPER LOWER LEN
- Lookup: VLOOKUP INDEX MATCH (exact match only)
- Date: TODAY NOW DATE YEAR MONTH DAY (day-serial numbers, arithmetic-friendly)

## Known limitations (by design, to keep this a readable ~1500 lines instead of an unreviewable 10,000)

- **No cross-sheet formulas** (`Sheet2!A1`) — formulas resolve against the active sheet only.
- **Fixed row height** — only column width is resizable.
- **Fill handle** only extends down or right, not up/left.
- **Multi-range selection** (Ctrl+click) only feeds the status bar; copy, fill, and paste act on the primary range.
- **Copy/paste** does not adjust relative references when the source and pasted-to ranges have different shapes (it aligns from the paste's top-left corner).
- **Borders and conditional formatting** are not implemented (bold/italic/color/numFmt are).
- **Dates** use a day-serial epoch anchored to Unix time, not Excel's 1899-12-30 epoch — internally consistent for arithmetic in this app, but a serial number copied out won't match Excel's serial for the same date.
- **XLSX import** trusts well-formed files from Excel/Sheets/LibreOffice; it doesn't handle every OOXML edge case (charts, pivot tables, data validation, conditional formatting are ignored rather than crashing).

## Extending it

- New formula functions go in `FUNCTIONS` in `formulaEngine.js` — most are one-liners.
- New cell style properties: add the field to the style object, read it in `formatValue`/`renderCell` in `Spreadsheet.jsx`, and add read/write support for it in `xlsxIO.js`'s style section if you want it to survive an XLSX round-trip.
- Cross-sheet formulas would go in the tokenizer (`Sheet2!A1` → a ref token with a `sheet` field) and the evaluator context (`ctx` would need access to all sheets, not just one).
