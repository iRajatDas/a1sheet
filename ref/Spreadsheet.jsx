import React, { useState, useRef, useEffect, useMemo } from "react";
import {
  colToLetters, lettersToCol, parseCellRef, createEvaluator, shiftFormulaRefs, extrapolateSeries,
} from "./formulaEngine.js";
import {
  readXlsxFile, downloadXlsx, csvToCells, downloadCsv, getUsedBounds,
} from "./xlsxIO.js";

/* ---------------------------------------------------------------------------
   SCOPE NOTES (read this before extending):
   - Row height is fixed (only column width is resizable) to keep the grid
     math tractable — see ROW_HEIGHT.
   - Formulas are single-sheet only; no Sheet2!A1 cross-sheet refs.
   - Fill handle drag supports extending down or right (the common case),
     not up/left.
   - Multi-range selection (Ctrl+click) is for viewing/status-bar stats;
     copy/fill/paste act on the primary (most recent) range only.
   - Named ranges are stored at workbook level but resolved against whichever
     sheet is currently active when a formula uses them.
   --------------------------------------------------------------------------- */

const ROW_HEIGHT = 26;
const HEADER_HEIGHT = 26;
const ROW_HEADER_WIDTH = 44;
const DEFAULT_COL_WIDTH = 92;
const BUFFER_ROWS = 6;
const ACCENT = "#0d9488";

function clamp(v, lo, hi) { return Math.max(lo, Math.min(hi, v)); }
function uid() { return Math.random().toString(36).slice(2, 9); }

function makeSheet(name) {
  return {
    id: uid(), name, cells: {}, styles: {}, colWidths: {}, merges: [],
    frozenRows: 0, frozenCols: 0, hiddenRows: new Set(), colLabels: {}, rowLabels: {},
    filters: {}, numRows: 200, numCols: 26,
  };
}
function cloneSheet(sheet) {
  return {
    ...sheet, cells: { ...sheet.cells }, styles: { ...sheet.styles }, colWidths: { ...sheet.colWidths },
    merges: sheet.merges.map((m) => ({ ...m })), hiddenRows: new Set(sheet.hiddenRows),
    colLabels: { ...sheet.colLabels }, rowLabels: { ...sheet.rowLabels }, filters: { ...sheet.filters },
  };
}

/* --------------------------- row/col insert & delete --------------------------- */
function shiftKeys(obj, axis, at, delta) {
  const next = {};
  for (const key in obj) {
    const [r, c] = key.split("_").map(Number);
    const v = axis === "row" ? r : c;
    if (v < at) next[key] = obj[key];
    else if (delta < 0 && v === at) continue;
    else next[`${axis === "row" ? v + delta : r}_${axis === "row" ? c : v + delta}`] = obj[key];
  }
  return next;
}
function insertRow(sheet, at) { return { ...sheet, cells: shiftKeys(sheet.cells, "row", at, 1), styles: shiftKeys(sheet.styles, "row", at, 1), numRows: sheet.numRows + 1 }; }
function deleteRow(sheet, at) { return { ...sheet, cells: shiftKeys(sheet.cells, "row", at, -1), styles: shiftKeys(sheet.styles, "row", at, -1), numRows: Math.max(1, sheet.numRows - 1) }; }
function insertCol(sheet, at) { return { ...sheet, cells: shiftKeys(sheet.cells, "col", at, 1), styles: shiftKeys(sheet.styles, "col", at, 1), numCols: sheet.numCols + 1 }; }
function deleteCol(sheet, at) { return { ...sheet, cells: shiftKeys(sheet.cells, "col", at, -1), styles: shiftKeys(sheet.styles, "col", at, -1), numCols: Math.max(1, sheet.numCols - 1) }; }

function formatValue(raw, style) {
  if (raw === undefined || raw === "") return "";
  const numFmt = style && style.numFmt;
  if (!numFmt || numFmt === "general") return String(raw);
  const n = typeof raw === "number" ? raw : parseFloat(raw);
  if (isNaN(n)) return String(raw);
  switch (numFmt) {
    case "integer": return Math.round(n).toLocaleString();
    case "number": return n.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    case "percent": return (n * 100).toFixed(2) + "%";
    case "currency": return n.toLocaleString(undefined, { style: "currency", currency: "USD" });
    case "date": { const d = new Date(n * 86400000); return isNaN(d) ? String(raw) : d.toISOString().slice(0, 10); }
    default: return String(raw);
  }
}

export default function Spreadsheet() {
  const [workbook, setWorkbook] = useState(() => ({ sheets: [makeSheet("Sheet1")], activeSheetIndex: 0, namedRanges: {} }));
  const [selection, setSelection] = useState({ r1: 0, c1: 0, r2: 0, c2: 0 });
  const [extraRanges, setExtraRanges] = useState([]);
  const [editing, setEditing] = useState(null);
  const [history, setHistory] = useState([]);
  const [future, setFuture] = useState([]);
  const [scrollTop, setScrollTop] = useState(0);
  const [containerHeight, setContainerHeight] = useState(500);
  const [status, setStatus] = useState("");
  const [contextMenu, setContextMenu] = useState(null);
  const [colMenu, setColMenu] = useState(null);
  const [renaming, setRenaming] = useState(null); // {type:'col'|'row'|'sheet', index, value}
  const [nameBoxValue, setNameBoxValue] = useState("");
  const [fillDrag, setFillDrag] = useState(null);

  const containerRef = useRef(null);
  const hiddenRef = useRef(null);
  const fileInputRef = useRef(null);
  const colDragRef = useRef(null);

  const sheet = workbook.sheets[workbook.activeSheetIndex];
  const evaluator = useMemo(() => createEvaluator(sheet.cells, workbook.namedRanges), [sheet.cells, workbook.namedRanges]);

  useEffect(() => {
    const el = containerRef.current;
    if (!el || typeof ResizeObserver === "undefined") return;
    const ro = new ResizeObserver((entries) => { for (const e of entries) setContainerHeight(e.contentRect.height); });
    ro.observe(el);
    return () => ro.disconnect();
  }, []);
  useEffect(() => { if (!editing) hiddenRef.current && hiddenRef.current.focus(); }, [editing, workbook.activeSheetIndex]);
  useEffect(() => {
    const onClick = () => { setContextMenu(null); setColMenu(null); };
    window.addEventListener("click", onClick);
    return () => window.removeEventListener("click", onClick);
  }, []);

  /* ------------------------------- history / mutation ------------------------------- */
  function pushHistory() { setHistory((h) => [...h.slice(-49), workbook]); setFuture([]); }
  function undo() { if (!history.length) return; setFuture((f) => [workbook, ...f]); setWorkbook(history[history.length - 1]); setHistory((h) => h.slice(0, -1)); }
  function redo() { if (!future.length) return; setHistory((h) => [...h, workbook]); setWorkbook(future[0]); setFuture((f) => f.slice(1)); }

  function updateSheet(updater, addHistory = true) {
    if (addHistory) pushHistory();
    setWorkbook((wb) => {
      const sheets = wb.sheets.slice();
      sheets[wb.activeSheetIndex] = updater(cloneSheet(sheets[wb.activeSheetIndex]));
      return { ...wb, sheets };
    });
  }
  function updateWorkbook(updater, addHistory = true) {
    if (addHistory) pushHistory();
    setWorkbook((wb) => updater({ ...wb, sheets: wb.sheets.slice() }));
  }

  function selBounds(sel = selection) {
    return { rMin: Math.min(sel.r1, sel.r2), rMax: Math.max(sel.r1, sel.r2), cMin: Math.min(sel.c1, sel.c2), cMax: Math.max(sel.c1, sel.c2) };
  }
  function isLocked(r, c) { const s = sheet.styles[`${r}_${c}`]; return !!(s && s.locked); }
  function anyLockedInSelection() {
    const { rMin, rMax, cMin, cMax } = selBounds();
    for (let r = rMin; r <= rMax; r++) for (let c = cMin; c <= cMax; c++) if (isLocked(r, c)) return true;
    return false;
  }

  /* ------------------------------------ editing ------------------------------------ */
  function startEdit(r, c, initial) {
    if (isLocked(r, c)) { setStatus("That cell is locked."); return; }
    const key = `${r}_${c}`;
    setEditing({ r, c, value: initial !== undefined ? initial : sheet.cells[key] || "" });
  }
  function commitEdit(moveDelta) {
    if (!editing) return;
    const { r, c, value } = editing;
    updateSheet((s) => { const cells = { ...s.cells }; if (value === "") delete cells[`${r}_${c}`]; else cells[`${r}_${c}`] = value; return { ...s, cells }; });
    setEditing(null);
    if (moveDelta) moveActive(moveDelta[0], moveDelta[1], false);
  }
  function cancelEdit() { setEditing(null); }

  function clearSelectionCells() {
    if (anyLockedInSelection()) { setStatus("Selection contains locked cells — unlock to clear."); return; }
    const { rMin, rMax, cMin, cMax } = selBounds();
    updateSheet((s) => {
      const cells = { ...s.cells };
      for (let r = rMin; r <= rMax; r++) for (let c = cMin; c <= cMax; c++) delete cells[`${r}_${c}`];
      return { ...s, cells };
    });
  }
  function clearSelectionFormatting() {
    const { rMin, rMax, cMin, cMax } = selBounds();
    updateSheet((s) => {
      const styles = { ...s.styles };
      for (let r = rMin; r <= rMax; r++) for (let c = cMin; c <= cMax; c++) delete styles[`${r}_${c}`];
      return { ...s, styles };
    });
  }

  function applyStyle(patch) {
    const { rMin, rMax, cMin, cMax } = selBounds();
    updateSheet((s) => {
      const styles = { ...s.styles };
      for (let r = rMin; r <= rMax; r++) for (let c = cMin; c <= cMax; c++) { const key = `${r}_${c}`; styles[key] = { ...(styles[key] || {}), ...patch }; }
      return { ...s, styles };
    });
  }
  const activeStyle = sheet.styles[`${selection.r2}_${selection.c2}`] || {};

  /* --------------------------------- selection / nav --------------------------------- */
  function moveActive(dr, dc, extend) {
    setSelection((sel) => {
      const nr = clamp(sel.r2 + dr, 0, sheet.numRows - 1);
      const nc = clamp(sel.c2 + dc, 0, sheet.numCols - 1);
      return extend ? { ...sel, r2: nr, c2: nc } : { r1: nr, c1: nc, r2: nr, c2: nc };
    });
  }
  function handleCellMouseDown(r, c, e) {
    if (editing) commitEdit(null);
    if (e.ctrlKey || e.metaKey) { setExtraRanges((ex) => [...ex, selection]); setSelection({ r1: r, c1: c, r2: r, c2: c }); }
    else if (e.shiftKey) setSelection((sel) => ({ ...sel, r2: r, c2: c }));
    else { setExtraRanges([]); setSelection({ r1: r, c1: c, r2: r, c2: c }); }
    hiddenRef.current && hiddenRef.current.focus();
  }
  function handleCellMouseEnter(r, c, e) { if (e.buttons === 1 && !fillDrag) setSelection((sel) => ({ ...sel, r2: r, c2: c })); }

  /* ------------------------------------ clipboard ------------------------------------ */
  function handleHiddenKeyDown(e) {
    if (editing) return;
    const key = e.key;
    if (key === "ArrowUp") { moveActive(-1, 0, e.shiftKey); e.preventDefault(); return; }
    if (key === "ArrowDown") { moveActive(1, 0, e.shiftKey); e.preventDefault(); return; }
    if (key === "ArrowLeft") { moveActive(0, -1, e.shiftKey); e.preventDefault(); return; }
    if (key === "ArrowRight") { moveActive(0, 1, e.shiftKey); e.preventDefault(); return; }
    if (key === "Tab") { moveActive(0, e.shiftKey ? -1 : 1, false); e.preventDefault(); return; }
    if (key === "Enter" || key === "F2") { startEdit(selection.r2, selection.c2); e.preventDefault(); return; }
    if (key === "Backspace" || key === "Delete") { clearSelectionCells(); e.preventDefault(); return; }
    if ((e.ctrlKey || e.metaKey) && key.toLowerCase() === "b") { applyStyle({ bold: !activeStyle.bold }); e.preventDefault(); return; }
    if ((e.ctrlKey || e.metaKey) && key.toLowerCase() === "i") { applyStyle({ italic: !activeStyle.italic }); e.preventDefault(); return; }
    if ((e.ctrlKey || e.metaKey) && key.toLowerCase() === "z" && !e.shiftKey) { undo(); e.preventDefault(); return; }
    if ((e.ctrlKey || e.metaKey) && (key.toLowerCase() === "y" || (key.toLowerCase() === "z" && e.shiftKey))) { redo(); e.preventDefault(); return; }
    if (key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) { startEdit(selection.r2, selection.c2, key); e.preventDefault(); return; }
  }
  const lastCopiedRef = useRef(null);
  function handleCopy(e) {
    if (editing) return;
    const { rMin, rMax, cMin, cMax } = selBounds();
    const grid = [];
    for (let r = rMin; r <= rMax; r++) { const row = []; for (let c = cMin; c <= cMax; c++) row.push(sheet.cells[`${r}_${c}`] || ""); grid.push(row); }
    const text = grid.map((row) => row.join("\t")).join("\n");
    lastCopiedRef.current = { originR: rMin, originC: cMin, grid, text };
    e.clipboardData.setData("text/plain", text);
    e.preventDefault();
  }
  function handlePaste(e) {
    if (editing) return;
    const text = e.clipboardData.getData("text/plain");
    const destR = selection.r2, destC = selection.c2;
    const internal = lastCopiedRef.current && lastCopiedRef.current.text === text ? lastCopiedRef.current : null;
    let grid;
    if (internal) {
      const dRow = destR - internal.originR, dCol = destC - internal.originC;
      grid = internal.grid.map((row) => row.map((val) => (val.startsWith("=") ? "=" + shiftFormulaRefs(val.slice(1), dRow, dCol) : val)));
    } else {
      const rows = text.replace(/\r/g, "").split("\n");
      if (rows[rows.length - 1] === "") rows.pop();
      grid = rows.map((r) => r.split("\t"));
    }
    const needRows = destR + grid.length, needCols = destC + Math.max(...grid.map((r) => r.length), 0);
    updateSheet((s) => {
      const cells = { ...s.cells };
      grid.forEach((row, i) => row.forEach((val, j) => { const k = `${destR + i}_${destC + j}`; if (val === "") delete cells[k]; else cells[k] = val; }));
      return { ...s, cells, numRows: Math.max(s.numRows, needRows + 20), numCols: Math.max(s.numCols, needCols + 5) };
    });
    e.preventDefault();
  }

  /* ------------------------------------- fill handle ------------------------------------- */
  function startFillDrag() { const { rMin, rMax, cMin, cMax } = selBounds(); setFillDrag({ rMin, rMax, cMin, cMax, targetR: rMax, targetC: cMax }); }
  function updateFillDrag(r, c) { setFillDrag((fd) => (fd ? { ...fd, targetR: Math.max(fd.rMax, r), targetC: Math.max(fd.cMax, c) } : fd)); }
  function commitFillDrag() {
    if (!fillDrag) return;
    const { rMin, rMax, cMin, cMax, targetR, targetC } = fillDrag;
    const goingDown = targetR > rMax, goingRight = targetC > cMax;
    updateSheet((s) => {
      const cells = { ...s.cells };
      if (goingDown) {
        for (let c = cMin; c <= cMax; c++) {
          const src = []; for (let r = rMin; r <= rMax; r++) src.push(s.cells[`${r}_${c}`] || "");
          const isFormula = src.some((v) => v.startsWith("="));
          if (isFormula) {
            for (let r = rMax + 1; r <= targetR; r++) { const idx = (r - rMin) % src.length; const srcRow = rMin + idx; const v = src[idx]; cells[`${r}_${c}`] = v.startsWith("=") ? "=" + shiftFormulaRefs(v.slice(1), r - srcRow, 0) : v; }
          } else {
            const extra = extrapolateSeries(src, targetR - rMax);
            extra.forEach((v, i) => { cells[`${rMax + 1 + i}_${c}`] = v; });
          }
        }
      } else if (goingRight) {
        for (let r = rMin; r <= rMax; r++) {
          const src = []; for (let c = cMin; c <= cMax; c++) src.push(s.cells[`${r}_${c}`] || "");
          const isFormula = src.some((v) => v.startsWith("="));
          if (isFormula) {
            for (let c = cMax + 1; c <= targetC; c++) { const idx = (c - cMin) % src.length; const srcCol = cMin + idx; const v = src[idx]; cells[`${r}_${c}`] = v.startsWith("=") ? "=" + shiftFormulaRefs(v.slice(1), 0, c - srcCol) : v; }
          } else {
            const extra = extrapolateSeries(src, targetC - cMax);
            extra.forEach((v, i) => { cells[`${r}_${cMax + 1 + i}`] = v; });
          }
        }
      }
      return { ...s, cells };
    });
    setSelection({ r1: rMin, c1: cMin, r2: Math.max(targetR, rMax), c2: Math.max(targetC, cMax) });
    setFillDrag(null);
  }

  /* ---------------------------------- sheet management ---------------------------------- */
  function addSheet() { updateWorkbook((wb) => { const sheets = [...wb.sheets, makeSheet(`Sheet${wb.sheets.length + 1}`)]; return { ...wb, sheets, activeSheetIndex: sheets.length - 1 }; }); setSelection({ r1: 0, c1: 0, r2: 0, c2: 0 }); }
  function deleteSheet(i) { if (workbook.sheets.length <= 1) return; updateWorkbook((wb) => { const sheets = wb.sheets.filter((_, idx) => idx !== i); return { ...wb, sheets, activeSheetIndex: clamp(wb.activeSheetIndex, 0, sheets.length - 1) }; }); }
  function switchSheet(i) { setWorkbook((wb) => ({ ...wb, activeSheetIndex: i })); setSelection({ r1: 0, c1: 0, r2: 0, c2: 0 }); setExtraRanges([]); }

  /* --------------------------------------- sort / filter -------------------------------------- */
  function sortByColumn(colIdx, ascending) {
    updateSheet((s) => {
      const bounds = getUsedBounds(s.cells);
      const ev = createEvaluator(s.cells, workbook.namedRanges);
      const idxs = Array.from({ length: bounds.rows }, (_, i) => i);
      idxs.sort((a, b) => {
        const va = ev.getCellDisplay(a, colIdx), vb = ev.getCellDisplay(b, colIdx);
        const na = parseFloat(va), nb = parseFloat(vb);
        const cmp = (!isNaN(na) && !isNaN(nb)) ? na - nb : String(va).localeCompare(String(vb));
        return ascending ? cmp : -cmp;
      });
      const cells = {}, styles = {};
      idxs.forEach((origRow, newRow) => { for (let c = 0; c < bounds.cols; c++) { const ok = `${origRow}_${c}`, nk = `${newRow}_${c}`; if (s.cells[ok] !== undefined) cells[nk] = s.cells[ok]; if (s.styles[ok] !== undefined) styles[nk] = s.styles[ok]; } });
      return { ...s, cells, styles };
    });
    setColMenu(null);
  }
  function setColumnFilter(colIdx, allowedSet) { updateSheet((s) => ({ ...s, filters: { ...s.filters, [colIdx]: allowedSet } }), false); }
  function clearColumnFilter(colIdx) { updateSheet((s) => { const filters = { ...s.filters }; delete filters[colIdx]; return { ...s, filters }; }, false); }

  const effectiveHiddenRows = useMemo(() => {
    const hidden = new Set(sheet.hiddenRows);
    const entries = Object.entries(sheet.filters || {});
    if (entries.length) {
      const bounds = getUsedBounds(sheet.cells);
      for (let r = 0; r < bounds.rows; r++) for (const [colIdx, allowed] of entries) { const v = evaluator.getCellDisplay(r, Number(colIdx)); if (!allowed.has(String(v))) { hidden.add(r); break; } }
    }
    return hidden;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [sheet.hiddenRows, sheet.filters, sheet.cells, evaluator]);

  /* -------------------------------------- named ranges -------------------------------------- */
  function handleNameBoxEnter() {
    const val = nameBoxValue.trim();
    if (!val) return;
    if (/^[A-Za-z]+[0-9]+(:[A-Za-z]+[0-9]+)?$/.test(val)) {
      const [a, b] = val.split(":");
      const pa = parseCellRef(a), pb = b ? parseCellRef(b) : pa;
      setSelection({ r1: pa.row, c1: pa.col, r2: pb.row, c2: pb.col });
    } else if (workbook.namedRanges[val.toUpperCase()]) {
      const nr = workbook.namedRanges[val.toUpperCase()];
      setSelection({ r1: nr.r1, c1: nr.c1, r2: nr.r2, c2: nr.c2 });
    } else {
      const { rMin, rMax, cMin, cMax } = selBounds();
      updateWorkbook((wb) => ({ ...wb, namedRanges: { ...wb.namedRanges, [val.toUpperCase()]: { r1: rMin, c1: cMin, r2: rMax, c2: cMax } } }));
      setStatus(`Named range "${val.toUpperCase()}" created.`);
    }
    setNameBoxValue("");
  }

  /* ------------------------------------- import / export ------------------------------------- */
  function handleImportFile(e) {
    const file = e.target.files[0];
    if (!file) return;
    if (/\.xlsx$/i.test(file.name)) {
      readXlsxFile(file).then((importedSheets) => {
        pushHistory();
        setWorkbook((wb) => ({
          ...wb,
          sheets: importedSheets.map((s) => ({ ...makeSheet(s.name), cells: s.cells, styles: s.styles, merges: s.merges, numRows: Math.max(s.rows + 20, 200), numCols: Math.max(s.cols + 5, 26) })),
          activeSheetIndex: 0,
        }));
        setStatus(`Imported ${file.name} (${importedSheets.length} sheet${importedSheets.length > 1 ? "s" : ""})`);
      }).catch((err) => setStatus("Import failed: " + err.message));
    } else {
      file.text().then((text) => {
        const { cells, rows, cols } = csvToCells(text);
        pushHistory();
        updateSheet((s) => ({ ...s, cells, styles: {}, numRows: Math.max(rows + 20, 200), numCols: Math.max(cols + 5, 26) }), false);
        setStatus(`Imported ${file.name}`);
      });
    }
    e.target.value = "";
  }
  function handleExportXlsx() { downloadXlsx(workbook.sheets.map((s) => ({ name: s.name, cells: s.cells, styles: s.styles, merges: s.merges, namedRanges: workbook.namedRanges }))); }
  function handleExportCsv() { downloadCsv(sheet.cells, evaluator, `${sheet.name}.csv`); }

  /* ---------------------------------------- rendering ---------------------------------------- */
  const frozenRowsCount = sheet.frozenRows || 0;
  const frozenColsCount = sheet.frozenCols || 0;
  const cols = Array.from({ length: sheet.numCols }, (_, i) => i);
  const colWidth = (c) => sheet.colWidths[c] || DEFAULT_COL_WIDTH;
  const colOffset = (c) => { let x = 0; for (let i = 0; i < c; i++) x += colWidth(i); return x; };
  const gridTemplateColumns = `${ROW_HEADER_WIDTH}px ${cols.map((c) => `${colWidth(c)}px`).join(" ")}`;

  const normalRowMapping = useMemo(() => {
    const arr = [];
    for (let r = frozenRowsCount; r < sheet.numRows; r++) if (!effectiveHiddenRows.has(r)) arr.push(r);
    return arr;
  }, [frozenRowsCount, sheet.numRows, effectiveHiddenRows]);

  const bandScrollTop = Math.max(0, scrollTop - (HEADER_HEIGHT + frozenRowsCount * ROW_HEIGHT));
  const startVisual = clamp(Math.floor(bandScrollTop / ROW_HEIGHT) - BUFFER_ROWS, 0, normalRowMapping.length);
  const visibleCount = Math.ceil(containerHeight / ROW_HEIGHT) + BUFFER_ROWS * 2;
  const endVisual = clamp(startVisual + visibleCount, 0, normalRowMapping.length);
  const visibleNormalRows = normalRowMapping.slice(startVisual, endVisual).map((absRow, i) => ({ absRow, gridRow: frozenRowsCount + startVisual + i + 2 }));
  const frozenRowsList = Array.from({ length: frozenRowsCount }, (_, i) => ({ absRow: i, gridRow: i + 2 }));

  const { rMin, rMax, cMin, cMax } = selBounds();
  function inSelRect(r, c, rect) { return r >= rect.rMin && r <= rect.rMax && c >= rect.cMin && c <= rect.cMax; }
  function isSelected(r, c) {
    if (inSelRect(r, c, { rMin, rMax, cMin, cMax })) return true;
    return extraRanges.some((rect) => inSelRect(r, c, selBounds(rect)));
  }
  function getMergeAt(r, c) { return sheet.merges.find((m) => r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2); }

  function stickyStyleFor(isHeaderRow, isRowHeaderCol, frozenRowIdx, frozenColIdx) {
    const style = {};
    let z = 0;
    if (isHeaderRow) { style.top = 0; z = Math.max(z, 4); }
    else if (frozenRowIdx !== undefined) { style.top = HEADER_HEIGHT + frozenRowIdx * ROW_HEIGHT; z = Math.max(z, 2); }
    if (isRowHeaderCol) { style.left = 0; z = Math.max(z, isHeaderRow ? 5 : 3); }
    else if (frozenColIdx !== undefined) { style.left = ROW_HEADER_WIDTH + colOffset(frozenColIdx); z = Math.max(z, isHeaderRow ? 4 : 2); }
    if (Object.keys(style).length) { style.position = "sticky"; style.zIndex = z; style.background = "#fff"; }
    return style;
  }

  function renderCell(r, c, gridRow) {
    const merge = getMergeAt(r, c);
    if (merge && (merge.r1 !== r || merge.c1 !== c)) return null; // covered by a merge, skip
    const key = `${r}_${c}`;
    const raw = sheet.cells[key];
    const style = sheet.styles[key] || {};
    const selected = isSelected(r, c);
    const active = r === selection.r2 && c === selection.c2;
    const isEditing = editing && editing.r === r && editing.c === c;
    const frozenColIdx = c < frozenColsCount ? c : undefined;
    const gridColumn = c + 2;
    const span = merge ? `${gridColumn} / span ${merge.c2 - merge.c1 + 1}` : gridColumn;
    const rowSpan = merge ? `${gridRow} / span ${merge.r2 - merge.r1 + 1}` : gridRow;
    const displayVal = formatValue(evaluator.getCellDisplay(r, c), style);
    return (
      <div
        key={c}
        className={`ss-cell${selected ? " selected" : ""}${active ? " active" : ""}${style.locked ? " locked" : ""}`}
        style={{
          gridColumn: span, gridRow: rowSpan, height: ROW_HEIGHT, fontWeight: style.bold ? 700 : 400,
          fontStyle: style.italic ? "italic" : "normal", textDecoration: style.underline ? "underline" : "none",
          textAlign: style.align || "left", justifyContent: style.align === "center" ? "center" : style.align === "right" ? "flex-end" : "flex-start",
          color: style.color || "#1e293b", background: isEditing ? "#fff" : (style.bg || (selected ? undefined : "#fff")),
          ...stickyStyleFor(false, false, undefined, frozenColIdx),
        }}
        onMouseDown={(e) => handleCellMouseDown(r, c, e)}
        onMouseEnter={(e) => handleCellMouseEnter(r, c, e)}
        onDoubleClick={() => startEdit(r, c)}
        onContextMenu={(e) => { e.preventDefault(); if (!isSelected(r, c)) setSelection({ r1: r, c1: c, r2: r, c2: c }); setContextMenu({ x: e.clientX, y: e.clientY, r, c }); }}
        title={style.locked ? "Locked cell" : undefined}
      >
        {isEditing ? (
          <input
            autoFocus value={editing.value}
            onChange={(e) => setEditing((ed) => ({ ...ed, value: e.target.value }))}
            onKeyDown={(e) => { if (e.key === "Enter") commitEdit([1, 0]); else if (e.key === "Tab") { e.preventDefault(); commitEdit([0, e.shiftKey ? -1 : 1]); } else if (e.key === "Escape") cancelEdit(); }}
            onBlur={() => commitEdit(null)}
          />
        ) : displayVal}
        {active && !editing && r === rMax && c === cMax && (
          <div
            className="ss-fillhandle"
            onMouseDown={(e) => { e.stopPropagation(); e.preventDefault(); startFillDrag(); }}
          />
        )}
      </div>
    );
  }

  function renderRowHeader(r, gridRow, frozenRowIdx) {
    const isRenaming = renaming && renaming.type === "row" && renaming.index === r;
    return (
      <div
        key="rh" className="ss-head"
        style={{ gridColumn: 1, gridRow, height: ROW_HEIGHT, ...stickyStyleFor(false, true, frozenRowIdx, undefined) }}
        onMouseDown={() => setSelection({ r1: r, c1: 0, r2: r, c2: sheet.numCols - 1 })}
        onDoubleClick={() => setRenaming({ type: "row", index: r, value: sheet.rowLabels[r] || String(r + 1) })}
      >
        {isRenaming ? (
          <input autoFocus value={renaming.value} style={{ width: "100%" }}
            onChange={(e) => setRenaming((rn) => ({ ...rn, value: e.target.value }))}
            onKeyDown={(e) => { if (e.key === "Enter") { updateSheet((s) => ({ ...s, rowLabels: { ...s.rowLabels, [r]: renaming.value } }), false); setRenaming(null); } if (e.key === "Escape") setRenaming(null); }}
            onBlur={() => setRenaming(null)} />
        ) : (sheet.rowLabels[r] || r + 1)}
      </div>
    );
  }

  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100vh", fontFamily: "system-ui, -apple-system, sans-serif", color: "#1e293b", background: "#fff" }}>
      <style>{`
        .ss-cell { border-right: 1px solid #e2e8f0; border-bottom: 1px solid #e2e8f0; overflow: hidden; white-space: nowrap; display:flex; align-items:center; padding: 0 6px; font-size: 13px; font-family: ui-monospace, SFMono-Regular, Menlo, monospace; cursor: cell; position: relative; }
        .ss-cell.selected { background: rgba(13,148,136,0.10) !important; }
        .ss-cell.active { outline: 2px solid ${ACCENT}; outline-offset: -2px; z-index: 1; }
        .ss-cell.locked { background-image: repeating-linear-gradient(45deg, rgba(0,0,0,0.03) 0 4px, transparent 4px 8px); }
        .ss-head { background: #f8fafc; border-right: 1px solid #e2e8f0; border-bottom: 1px solid #cbd5e1; font-size: 12px; font-weight: 600; color: #475569; display:flex; align-items:center; justify-content:center; user-select:none; position: relative; }
        .ss-resize { position:absolute; right:0; top:0; width:5px; height:100%; cursor:col-resize; z-index:6; }
        .ss-btn { border: 1px solid #d1d5db; background: #fff; padding: 5px 9px; border-radius: 6px; font-size: 13px; cursor: pointer; color:#1e293b; }
        .ss-btn:hover { background: #f1f5f9; }
        .ss-btn.on { background: ${ACCENT}; color:#fff; border-color:${ACCENT}; }
        .ss-cell input { border: none; outline: none; width: 100%; height: 100%; font-size: 13px; font-family: ui-monospace, SFMono-Regular, Menlo, monospace; background:#fff; }
        .ss-fillhandle { position:absolute; right:-4px; bottom:-4px; width:7px; height:7px; background:${ACCENT}; border:1px solid #fff; cursor:crosshair; z-index:5; }
        .ss-menu { position:fixed; background:#fff; border:1px solid #d1d5db; border-radius:8px; box-shadow:0 4px 16px rgba(0,0,0,0.12); padding:4px; z-index:1000; min-width:180px; }
        .ss-menu button { display:block; width:100%; text-align:left; border:none; background:none; padding:6px 10px; font-size:13px; border-radius:4px; cursor:pointer; }
        .ss-menu button:hover { background:#f1f5f9; }
        .ss-menu hr { border:none; border-top:1px solid #eee; margin:4px 0; }
        .ss-tab { padding:6px 12px; font-size:13px; border-radius:6px 6px 0 0; cursor:pointer; border:1px solid transparent; }
        .ss-tab.active { background:#fff; border-color:#e2e8f0; border-bottom-color:#fff; font-weight:600; }
        .ss-sep { display:inline-block; width:1px; height:20px; background:#e2e8f0; }
      `}</style>

      {/* Toolbar */}
      <div style={{ display: "flex", gap: 6, alignItems: "center", padding: "6px 10px", borderBottom: "1px solid #e2e8f0", flexWrap: "wrap" }}>
        <button className="ss-btn" onClick={undo} disabled={!history.length} title="Undo (Ctrl+Z)">↶</button>
        <button className="ss-btn" onClick={redo} disabled={!future.length} title="Redo (Ctrl+Y)">↷</button>
        <span className="ss-sep" />
        <button className={`ss-btn${activeStyle.bold ? " on" : ""}`} onClick={() => applyStyle({ bold: !activeStyle.bold })} title="Bold (Ctrl+B)"><b>B</b></button>
        <button className={`ss-btn${activeStyle.italic ? " on" : ""}`} onClick={() => applyStyle({ italic: !activeStyle.italic })} title="Italic (Ctrl+I)"><i>I</i></button>
        <button className={`ss-btn${activeStyle.underline ? " on" : ""}`} onClick={() => applyStyle({ underline: !activeStyle.underline })} title="Underline"><u>U</u></button>
        <button className={`ss-btn${activeStyle.align === "left" ? " on" : ""}`} onClick={() => applyStyle({ align: "left" })}>⯇</button>
        <button className={`ss-btn${activeStyle.align === "center" ? " on" : ""}`} onClick={() => applyStyle({ align: "center" })}>≡</button>
        <button className={`ss-btn${activeStyle.align === "right" ? " on" : ""}`} onClick={() => applyStyle({ align: "right" })}>⯈</button>
        <input type="color" title="Text color" value={activeStyle.color || "#1e293b"} onChange={(e) => applyStyle({ color: e.target.value })} style={{ width: 28, height: 28, padding: 0, border: "1px solid #d1d5db", borderRadius: 6 }} />
        <input type="color" title="Fill color" value={activeStyle.bg || "#ffffff"} onChange={(e) => applyStyle({ bg: e.target.value })} style={{ width: 28, height: 28, padding: 0, border: "1px solid #d1d5db", borderRadius: 6 }} />
        <select className="ss-btn" value={activeStyle.numFmt || "general"} onChange={(e) => applyStyle({ numFmt: e.target.value })}>
          <option value="general">General</option><option value="integer">Integer</option><option value="number">0.00</option>
          <option value="percent">Percent</option><option value="currency">Currency</option><option value="date">Date</option>
        </select>
        <button className={`ss-btn${activeStyle.locked ? " on" : ""}`} onClick={() => applyStyle({ locked: !activeStyle.locked })} title="Lock/unlock selection">🔒</button>
        <span className="ss-sep" />
        <button className="ss-btn" onClick={() => updateSheet((s) => insertRow(s, selection.r2))}>+Row</button>
        <button className="ss-btn" onClick={() => updateSheet((s) => deleteRow(s, selection.r2))}>−Row</button>
        <button className="ss-btn" onClick={() => updateSheet((s) => insertCol(s, selection.c2))}>+Col</button>
        <button className="ss-btn" onClick={() => updateSheet((s) => deleteCol(s, selection.c2))}>−Col</button>
        <button className="ss-btn" onClick={() => { const { rMin, cMin, rMax, cMax } = selBounds(); if (rMin === rMax && cMin === cMax) return; updateSheet((s) => ({ ...s, merges: [...s.merges, { r1: rMin, c1: cMin, r2: rMax, c2: cMax }] }), false); }}>Merge</button>
        <button className="ss-btn" onClick={() => updateSheet((s) => ({ ...s, frozenRows: selection.r2 + 1 }), false)}>Freeze rows</button>
        <button className="ss-btn" onClick={() => updateSheet((s) => ({ ...s, frozenCols: selection.c2 + 1 }), false)}>Freeze cols</button>
        <button className="ss-btn" onClick={() => updateSheet((s) => ({ ...s, frozenRows: 0, frozenCols: 0 }), false)}>Unfreeze</button>
        <span className="ss-sep" />
        <button className="ss-btn" onClick={handleExportCsv}>Export CSV</button>
        <button className="ss-btn" onClick={handleExportXlsx}>Export XLSX</button>
        <button className="ss-btn" onClick={() => fileInputRef.current.click()}>Import</button>
        <input ref={fileInputRef} type="file" accept=".csv,.xlsx" style={{ display: "none" }} onChange={handleImportFile} />
        {status && <span style={{ fontSize: 12, color: "#64748b" }}>{status}</span>}
      </div>

      {/* Formula bar */}
      <div style={{ display: "flex", alignItems: "center", gap: 8, padding: "5px 10px", borderBottom: "1px solid #e2e8f0" }}>
        <input
          value={nameBoxValue} placeholder={colToLetters(selection.c2) + (selection.r2 + 1)}
          onChange={(e) => setNameBoxValue(e.target.value)}
          onKeyDown={(e) => { if (e.key === "Enter") handleNameBoxEnter(); }}
          style={{ width: 90, fontSize: 12, padding: "4px 6px", border: "1px solid #d1d5db", borderRadius: 6 }}
          title="Name box: type a cell/range to jump to it, or a name + Enter to define a named range for the current selection"
        />
        <span style={{ color: "#94a3b8" }}>fx</span>
        <input
          value={editing && editing.r === selection.r2 && editing.c === selection.c2 ? editing.value : (sheet.cells[`${selection.r2}_${selection.c2}`] || "")}
          onChange={(e) => { if (!editing) startEdit(selection.r2, selection.c2, e.target.value); else setEditing((ed) => ({ ...ed, value: e.target.value })); }}
          onFocus={() => { if (!editing) startEdit(selection.r2, selection.c2); }}
          onKeyDown={(e) => { if (e.key === "Enter") commitEdit([1, 0]); if (e.key === "Escape") cancelEdit(); }}
          onBlur={() => { if (editing) commitEdit(null); }}
          style={{ flex: 1, fontSize: 13, padding: "4px 8px", border: "1px solid #d1d5db", borderRadius: 6, fontFamily: "ui-monospace, monospace" }}
        />
      </div>

      {/* Hidden input catches keyboard + clipboard while a cell is selected but not being edited */}
      <textarea ref={hiddenRef} value="" onChange={() => {}} onKeyDown={handleHiddenKeyDown} onCopy={handleCopy} onPaste={handlePaste}
        style={{ position: "absolute", opacity: 0, width: 1, height: 1, pointerEvents: "none" }} />

      {/* Grid */}
      <div ref={containerRef} onScroll={(e) => setScrollTop(e.target.scrollTop)} onMouseUp={commitFillDrag}
        style={{ flex: 1, minHeight: 0, overflow: "auto", position: "relative" }}>
        <div style={{
          display: "grid", gridTemplateColumns, gridTemplateRows: `${HEADER_HEIGHT}px repeat(${frozenRowsCount}, ${ROW_HEIGHT}px)`,
          gridAutoRows: `${ROW_HEIGHT}px`, position: "relative",
        }}>
          {/* corner */}
          <div className="ss-head" style={{ gridColumn: 1, gridRow: 1, height: HEADER_HEIGHT, ...stickyStyleFor(true, true) }} />
          {/* column headers */}
          {cols.map((c) => {
            const isRenamingCol = renaming && renaming.type === "col" && renaming.index === c;
            return (
              <div key={c} className="ss-head" style={{ gridColumn: c + 2, gridRow: 1, height: HEADER_HEIGHT, ...stickyStyleFor(true, false, undefined, c < frozenColsCount ? c : undefined) }}
                onMouseDown={() => setSelection({ r1: 0, c1: c, r2: sheet.numRows - 1, c2: c })}
                onDoubleClick={() => setRenaming({ type: "col", index: c, value: sheet.colLabels[c] || colToLetters(c) })}>
                {isRenamingCol ? (
                  <input autoFocus value={renaming.value} style={{ width: "80%" }}
                    onChange={(e) => setRenaming((rn) => ({ ...rn, value: e.target.value }))}
                    onKeyDown={(e) => { if (e.key === "Enter") { updateSheet((s) => ({ ...s, colLabels: { ...s.colLabels, [c]: renaming.value } }), false); setRenaming(null); } if (e.key === "Escape") setRenaming(null); }}
                    onBlur={() => setRenaming(null)} />
                ) : <>{sheet.colLabels[c] || colToLetters(c)}
                  <span style={{ marginLeft: 4, cursor: "pointer", fontSize: 10 }} onClick={(e) => { e.stopPropagation(); setColMenu({ colIdx: c, x: e.clientX, y: e.clientY }); }}>▾</span>
                </>}
                <div className="ss-resize" onMouseDown={(e) => {
                  e.stopPropagation();
                  const startX = e.clientX, startWidth = colWidth(c);
                  colDragRef.current = { c, startX, startWidth };
                  const onMove = (ev) => { const d = colDragRef.current; if (!d) return; const w = Math.max(40, d.startWidth + (ev.clientX - d.startX)); updateSheet((s) => ({ ...s, colWidths: { ...s.colWidths, [d.c]: w } }), false); };
                  const onUp = () => { colDragRef.current = null; window.removeEventListener("mousemove", onMove); window.removeEventListener("mouseup", onUp); };
                  window.addEventListener("mousemove", onMove); window.addEventListener("mouseup", onUp);
                }} />
              </div>
            );
          })}
          {/* frozen rows */}
          {frozenRowsList.map(({ absRow, gridRow }) => (
            <React.Fragment key={`fr${absRow}`}>
              {renderRowHeader(absRow, gridRow, absRow)}
              {cols.map((c) => renderCell(absRow, c, gridRow))}
            </React.Fragment>
          ))}
          {/* normal (virtualized) rows */}
          {visibleNormalRows.map(({ absRow, gridRow }) => (
            <React.Fragment key={absRow}>
              {renderRowHeader(absRow, gridRow, undefined)}
              {cols.map((c) => renderCell(absRow, c, gridRow))}
            </React.Fragment>
          ))}
        </div>
        {/* fill-drag overlay: track mouse position across the whole scroll area */}
        {fillDrag && (
          <div style={{ position: "absolute", inset: 0, cursor: "crosshair" }}
            onMouseMove={(e) => {
              const rect = containerRef.current.getBoundingClientRect();
              const y = e.clientY - rect.top + scrollTop - HEADER_HEIGHT - frozenRowsCount * ROW_HEIGHT;
              const x = e.clientX - rect.left - ROW_HEADER_WIDTH;
              const row = frozenRowsCount + clamp(Math.floor(y / ROW_HEIGHT), 0, normalRowMapping.length - 1);
              let acc = 0, col = 0;
              for (let i = 0; i < cols.length; i++) { acc += colWidth(i); if (x < acc) { col = i; break; } col = i; }
              updateFillDrag(row, col);
            }}
          />
        )}
      </div>

      {/* Status bar */}
      <StatusBar sheet={sheet} evaluator={evaluator} selection={selection} extraRanges={extraRanges} />

      {/* Sheet tabs */}
      <div style={{ display: "flex", alignItems: "center", gap: 4, padding: "4px 10px", borderTop: "1px solid #e2e8f0", background: "#f8fafc" }}>
        {workbook.sheets.map((s, i) => (
          <div key={s.id} className={`ss-tab${i === workbook.activeSheetIndex ? " active" : ""}`} onClick={() => switchSheet(i)}
            onDoubleClick={() => setRenaming({ type: "sheet", index: i, value: s.name })}>
            {renaming && renaming.type === "sheet" && renaming.index === i ? (
              <input autoFocus value={renaming.value} style={{ width: 80 }}
                onChange={(e) => setRenaming((rn) => ({ ...rn, value: e.target.value }))}
                onKeyDown={(e) => { if (e.key === "Enter") { updateWorkbook((wb) => { const sheets = wb.sheets.slice(); sheets[i] = { ...sheets[i], name: renaming.value }; return { ...wb, sheets }; }, false); setRenaming(null); } if (e.key === "Escape") setRenaming(null); }}
                onBlur={() => setRenaming(null)} onClick={(e) => e.stopPropagation()} />
            ) : (<>{s.name} {workbook.sheets.length > 1 && <span style={{ marginLeft: 6, color: "#94a3b8" }} onClick={(e) => { e.stopPropagation(); deleteSheet(i); }}>×</span>}</>)}
          </div>
        ))}
        <button className="ss-btn" onClick={addSheet} style={{ padding: "3px 8px" }}>+</button>
      </div>

      {/* Context menu */}
      {contextMenu && (
        <div className="ss-menu" style={{ left: contextMenu.x, top: contextMenu.y }} onClick={(e) => e.stopPropagation()}>
          <button onClick={() => { handleCopyMenu(); setContextMenu(null); }}>Copy</button>
          <button onClick={async () => { try { const text = await navigator.clipboard.readText(); pasteText(text); } catch (e) { setStatus("Paste needs clipboard permission — try Ctrl+V instead."); } setContextMenu(null); }}>Paste</button>
          <hr />
          <button onClick={() => { updateSheet((s) => insertRow(s, contextMenu.r)); setContextMenu(null); }}>Insert row above</button>
          <button onClick={() => { updateSheet((s) => insertRow(s, contextMenu.r + 1)); setContextMenu(null); }}>Insert row below</button>
          <button onClick={() => { updateSheet((s) => deleteRow(s, contextMenu.r)); setContextMenu(null); }}>Delete row</button>
          <hr />
          <button onClick={() => { updateSheet((s) => insertCol(s, contextMenu.c)); setContextMenu(null); }}>Insert column left</button>
          <button onClick={() => { updateSheet((s) => insertCol(s, contextMenu.c + 1)); setContextMenu(null); }}>Insert column right</button>
          <button onClick={() => { updateSheet((s) => deleteCol(s, contextMenu.c)); setContextMenu(null); }}>Delete column</button>
          <hr />
          <button onClick={() => { clearSelectionCells(); setContextMenu(null); }}>Clear contents</button>
          <button onClick={() => { clearSelectionFormatting(); setContextMenu(null); }}>Clear formatting</button>
          <button onClick={() => { applyStyle({ locked: !isLocked(contextMenu.r, contextMenu.c) }); setContextMenu(null); }}>{isLocked(contextMenu.r, contextMenu.c) ? "Unlock" : "Lock"} cell(s)</button>
        </div>
      )}

      {/* Column header menu (sort / filter) */}
      {colMenu && <ColumnMenu colMenu={colMenu} sheet={sheet} evaluator={evaluator} onSort={sortByColumn} onFilter={setColumnFilter} onClearFilter={clearColumnFilter} onClose={() => setColMenu(null)} />}
    </div>
  );

  function handleCopyMenu() {
    const { rMin, rMax, cMin, cMax } = selBounds();
    const grid = [];
    for (let r = rMin; r <= rMax; r++) { const row = []; for (let c = cMin; c <= cMax; c++) row.push(sheet.cells[`${r}_${c}`] || ""); grid.push(row); }
    const text = grid.map((row) => row.join("\t")).join("\n");
    lastCopiedRef.current = { originR: rMin, originC: cMin, grid, text };
    navigator.clipboard.writeText(text).catch(() => {});
  }
  function pasteText(text) {
    const destR = selection.r2, destC = selection.c2;
    const rows = text.replace(/\r/g, "").split("\n");
    if (rows[rows.length - 1] === "") rows.pop();
    const grid = rows.map((r) => r.split("\t"));
    updateSheet((s) => {
      const cells = { ...s.cells };
      grid.forEach((row, i) => row.forEach((val, j) => { const k = `${destR + i}_${destC + j}`; if (val === "") delete cells[k]; else cells[k] = val; }));
      return { ...s, cells };
    });
  }
}

function StatusBar({ sheet, evaluator, selection, extraRanges }) {
  const rects = [selection, ...extraRanges].map((sel) => ({ rMin: Math.min(sel.r1, sel.r2), rMax: Math.max(sel.r1, sel.r2), cMin: Math.min(sel.c1, sel.c2), cMax: Math.max(sel.c1, sel.c2) }));
  let sum = 0, count = 0, numCount = 0;
  for (const rect of rects) for (let r = rect.rMin; r <= rect.rMax; r++) for (let c = rect.cMin; c <= rect.cMax; c++) {
    const v = evaluator.getCellDisplay(r, c);
    if (v !== "") { count++; const n = parseFloat(v); if (!isNaN(n)) { sum += n; numCount++; } }
  }
  const avg = numCount ? sum / numCount : 0;
  return (
    <div style={{ display: "flex", gap: 16, padding: "4px 12px", fontSize: 12, color: "#64748b", borderTop: "1px solid #e2e8f0" }}>
      <span>Count: {count}</span>
      {numCount > 0 && <><span>Sum: {Math.round(sum * 1e6) / 1e6}</span><span>Average: {Math.round(avg * 1e6) / 1e6}</span></>}
    </div>
  );
}

function ColumnMenu({ colMenu, sheet, evaluator, onSort, onFilter, onClearFilter, onClose }) {
  const { colIdx, x, y } = colMenu;
  const bounds = getUsedBounds(sheet.cells);
  const values = useMemo(() => { const set = new Set(); for (let r = 0; r < bounds.rows; r++) set.add(String(evaluator.getCellDisplay(r, colIdx))); return Array.from(set); }, [bounds.rows, colIdx, evaluator]);
  const currentFilter = sheet.filters[colIdx];
  const [checked, setChecked] = useState(() => new Set(currentFilter || values));
  return (
    <div className="ss-menu" style={{ left: x, top: y, maxHeight: 320, overflow: "auto" }} onClick={(e) => e.stopPropagation()}>
      <button onClick={() => { onSort(colIdx, true); onClose(); }}>Sort A → Z</button>
      <button onClick={() => { onSort(colIdx, false); onClose(); }}>Sort Z → A</button>
      <hr />
      <div style={{ padding: "4px 10px", fontSize: 12, color: "#64748b" }}>Filter values</div>
      {values.slice(0, 50).map((v) => (
        <label key={v} style={{ display: "flex", gap: 6, alignItems: "center", padding: "3px 10px", fontSize: 13 }}>
          <input type="checkbox" checked={checked.has(v)} onChange={(e) => { const next = new Set(checked); e.target.checked ? next.add(v) : next.delete(v); setChecked(next); }} />
          {v || "(blank)"}
        </label>
      ))}
      <hr />
      <button onClick={() => { onFilter(colIdx, checked); onClose(); }}>Apply filter</button>
      {currentFilter && <button onClick={() => { onClearFilter(colIdx); onClose(); }}>Clear filter</button>}
    </div>
  );
}
