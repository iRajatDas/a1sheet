/* =============================================================================
   XLSX / CSV I/O — zero npm dependencies.
   Implements just enough of ZIP, DEFLATE (RFC1951), and OOXML SpreadsheetML
   to read real .xlsx files (multi-sheet, with basic cell styling) and write
   valid ones back out.

   Known simplifications:
   - Styles carried: bold, italic, background color, font color, number
     format (general/integer/number/percent/currency/date), horizontal align.
     Borders, merged-cell styling nuances, and conditional formatting are not
     round-tripped.
   - Export always writes uncompressed (STORE) zip entries — valid xlsx,
     just not as small as a fully DEFLATE-compressed one.
   - Import decompresses DEFLATE entries with a from-scratch decoder, so it
     can open normally-saved Excel/Sheets/LibreOffice files.
   ============================================================================= */
import { colToLetters, lettersToCol, createEvaluator } from "./formulaEngine.js";

/* ================================ CRC32 ================================== */
const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; }
  return t;
})();
function crc32(bytes) { let c = 0xffffffff; for (let i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xff] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; }

/* ============================ DEFLATE decoder ============================= */
const LENBASE = [3,4,5,6,7,8,9,10,11,13,15,17,19,23,27,31,35,43,51,59,67,83,99,115,131,163,195,227,258];
const LENEXTRA = [0,0,0,0,0,0,0,0,1,1,1,1,2,2,2,2,3,3,3,3,4,4,4,4,5,5,5,5,0];
const DISTBASE = [1,2,3,4,5,7,9,13,17,25,33,49,65,97,129,193,257,385,513,769,1025,1537,2049,3073,4097,6145,8193,12289,16385,24577];
const DISTEXTRA = [0,0,0,0,1,1,2,2,3,3,4,4,5,5,6,6,7,7,8,8,9,9,10,10,11,11,12,12,13,13];
const CLCIDX = [16,17,18,0,8,7,9,6,10,5,11,4,12,3,13,2,14,1,15];

class BitReader {
  constructor(bytes) { this.bytes = bytes; this.pos = 0; this.bitpos = 0; }
  getBit() { if (this.pos >= this.bytes.length) return 0; const bit = (this.bytes[this.pos] >> this.bitpos) & 1; this.bitpos++; if (this.bitpos === 8) { this.bitpos = 0; this.pos++; } return bit; }
  getBits(n) { let v = 0; for (let i = 0; i < n; i++) v |= this.getBit() << i; return v; }
  align() { if (this.bitpos !== 0) { this.bitpos = 0; this.pos++; } }
}
function buildHuffman(lengths) {
  const MAXBITS = 15;
  const counts = new Array(MAXBITS + 1).fill(0);
  for (const len of lengths) if (len) counts[len]++;
  const offs = new Array(MAXBITS + 2).fill(0);
  for (let i = 1; i <= MAXBITS; i++) offs[i + 1] = offs[i] + counts[i];
  const symbols = new Array(lengths.length).fill(0);
  const offsCopy = offs.slice();
  for (let i = 0; i < lengths.length; i++) if (lengths[i]) symbols[offsCopy[lengths[i]]++] = i;
  return { counts, symbols };
}
function decodeSymbol(br, table) {
  let code = 0, first = 0, index = 0;
  for (let len = 1; len <= 15; len++) {
    code |= br.getBit();
    const count = table.counts[len] || 0;
    if (code - first < count) return table.symbols[index + (code - first)];
    index += count; first = (first + count) << 1; code <<= 1;
  }
  throw new Error("bad huffman code");
}
const FIXED_LIT_TABLE = (() => {
  const lens = new Array(288);
  for (let i = 0; i < 144; i++) lens[i] = 8;
  for (let i = 144; i < 256; i++) lens[i] = 9;
  for (let i = 256; i < 280; i++) lens[i] = 7;
  for (let i = 280; i < 288; i++) lens[i] = 8;
  return buildHuffman(lens);
})();
const FIXED_DIST_TABLE = buildHuffman(new Array(30).fill(5));

function inflateRaw(bytes) {
  const br = new BitReader(bytes);
  const out = [];
  let final = 0;
  while (!final) {
    final = br.getBit();
    const type = br.getBits(2);
    if (type === 0) {
      br.align();
      const len = bytes[br.pos] | (bytes[br.pos + 1] << 8);
      br.pos += 4;
      for (let i = 0; i < len; i++) out.push(bytes[br.pos + i]);
      br.pos += len;
    } else {
      let litTable, distTable;
      if (type === 1) { litTable = FIXED_LIT_TABLE; distTable = FIXED_DIST_TABLE; }
      else if (type === 2) {
        const hlit = br.getBits(5) + 257, hdist = br.getBits(5) + 1, hclen = br.getBits(4) + 4;
        const clLengths = new Array(19).fill(0);
        for (let i = 0; i < hclen; i++) clLengths[CLCIDX[i]] = br.getBits(3);
        const clTable = buildHuffman(clLengths);
        const lengths = [];
        while (lengths.length < hlit + hdist) {
          const sym = decodeSymbol(br, clTable);
          if (sym < 16) lengths.push(sym);
          else if (sym === 16) { const rep = br.getBits(2) + 3; const prev = lengths[lengths.length - 1] || 0; for (let i = 0; i < rep; i++) lengths.push(prev); }
          else if (sym === 17) { const rep = br.getBits(3) + 3; for (let i = 0; i < rep; i++) lengths.push(0); }
          else { const rep = br.getBits(7) + 11; for (let i = 0; i < rep; i++) lengths.push(0); }
        }
        litTable = buildHuffman(lengths.slice(0, hlit));
        distTable = buildHuffman(lengths.slice(hlit));
      } else throw new Error("invalid deflate block type");
      while (true) {
        const sym = decodeSymbol(br, litTable);
        if (sym < 256) out.push(sym);
        else if (sym === 256) break;
        else {
          const li = sym - 257;
          const len = LENBASE[li] + br.getBits(LENEXTRA[li]);
          const dsym = decodeSymbol(br, distTable);
          const dist = DISTBASE[dsym] + br.getBits(DISTEXTRA[dsym]);
          const start = out.length - dist;
          for (let i = 0; i < len; i++) out.push(out[start + i]);
        }
      }
    }
  }
  return new Uint8Array(out);
}

/* ================================== ZIP =================================== */
function readU16(b, o) { return b[o] | (b[o + 1] << 8); }
function readU32(b, o) { return (b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24)) >>> 0; }

function unzip(bytes) {
  let eocd = -1;
  for (let i = bytes.length - 22; i >= 0; i--) if (readU32(bytes, i) === 0x06054b50) { eocd = i; break; }
  if (eocd < 0) throw new Error("not a zip file");
  const cdOffset = readU32(bytes, eocd + 16);
  const cdEntries = readU16(bytes, eocd + 10);
  const files = {};
  let p = cdOffset;
  for (let i = 0; i < cdEntries; i++) {
    if (readU32(bytes, p) !== 0x02014b50) break;
    const method = readU16(bytes, p + 10);
    const compSize = readU32(bytes, p + 20);
    const nameLen = readU16(bytes, p + 28);
    const extraLen = readU16(bytes, p + 30);
    const commentLen = readU16(bytes, p + 32);
    const localOffset = readU32(bytes, p + 42);
    const name = new TextDecoder().decode(bytes.subarray(p + 46, p + 46 + nameLen));
    const lfNameLen = readU16(bytes, localOffset + 26);
    const lfExtraLen = readU16(bytes, localOffset + 28);
    const dataStart = localOffset + 30 + lfNameLen + lfExtraLen;
    const compData = bytes.subarray(dataStart, dataStart + compSize);
    files[name] = method === 8 ? inflateRaw(compData) : compData;
    p += 46 + nameLen + extraLen + commentLen;
  }
  return files;
}
function makeZip(fileList) {
  const encoder = new TextEncoder();
  const localParts = [], centralParts = [];
  let offset = 0;
  for (const f of fileList) {
    const nameBytes = encoder.encode(f.name);
    const data = f.data;
    const crc = crc32(data), size = data.length;
    const local = new Uint8Array(30 + nameBytes.length);
    const dv = new DataView(local.buffer);
    dv.setUint32(0, 0x04034b50, true);
    dv.setUint16(4, 20, true);
    dv.setUint32(14, crc, true); dv.setUint32(18, size, true); dv.setUint32(22, size, true);
    dv.setUint16(26, nameBytes.length, true);
    local.set(nameBytes, 30);
    localParts.push(local, data);

    const central = new Uint8Array(46 + nameBytes.length);
    const cdv = new DataView(central.buffer);
    cdv.setUint32(0, 0x02014b50, true);
    cdv.setUint16(4, 20, true); cdv.setUint16(6, 20, true);
    cdv.setUint32(16, crc, true); cdv.setUint32(20, size, true); cdv.setUint32(24, size, true);
    cdv.setUint16(28, nameBytes.length, true);
    cdv.setUint32(42, offset, true);
    central.set(nameBytes, 46);
    centralParts.push(central);
    offset += local.length + data.length;
  }
  const centralStart = offset;
  let centralSize = 0; for (const c of centralParts) centralSize += c.length;
  const eocd = new Uint8Array(22);
  const edv = new DataView(eocd.buffer);
  edv.setUint32(0, 0x06054b50, true);
  edv.setUint16(8, fileList.length, true); edv.setUint16(10, fileList.length, true);
  edv.setUint32(12, centralSize, true); edv.setUint32(16, centralStart, true);
  const all = [...localParts, ...centralParts, eocd];
  let total = 0; for (const a of all) total += a.length;
  const out = new Uint8Array(total);
  let pos = 0; for (const a of all) { out.set(a, pos); pos += a.length; }
  return out;
}

/* ================================== XML ==================================== */
function xmlEscape(s) { return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;"); }
function colorToRgb(hex) { if (!hex) return "FF000000"; const h = hex.replace("#", "").toUpperCase(); return "FF" + (h.length === 6 ? h : "000000"); }
function rgbToColor(rgb) { if (!rgb || rgb.length < 6) return null; return "#" + rgb.slice(-6); }

/* Built-in numFmtIds we use on export; recognized on import too. */
const NUMFMT_TO_ID = { general: 0, integer: 1, number: 2, percent: 10, currency: 164, date: 14 };
const CUSTOM_NUMFMTS = { 164: '"$"#,##0.00' };

function styleKey(style) {
  return JSON.stringify([style.bold, style.italic, style.color || "", style.bg || "", style.numFmt || "general", style.align || ""]);
}

/* ================================ Export ================================== */
export function buildWorkbookXlsx(sheets) {
  // sheets: [{ name, cells, styles, merges }]
  const encoder = new TextEncoder();

  // --- collect distinct styles across all sheets ---
  const styleList = [{ bold: false, italic: false, color: "", bg: "", numFmt: "general", align: "" }];
  const styleIndex = new Map([[styleKey(styleList[0]), 0]]);
  function getXfIndex(style) {
    if (!style) return 0;
    const key = styleKey(style);
    if (styleIndex.has(key)) return styleIndex.get(key);
    styleList.push(style);
    styleIndex.set(key, styleList.length - 1);
    return styleList.length - 1;
  }

  const sheetXmls = sheets.map((sheet) => {
    const evaluator = createEvaluator(sheet.cells, sheet.namedRanges || {});
    let maxR = -1, maxC = -1;
    for (const key in sheet.cells) { if (!sheet.cells[key]) continue; const [r, c] = key.split("_").map(Number); if (r > maxR) maxR = r; if (c > maxC) maxC = c; }
    for (const key in sheet.styles || {}) { const [r, c] = key.split("_").map(Number); if (r > maxR) maxR = r; if (c > maxC) maxC = c; }
    const rows = Math.max(maxR + 1, 1), cols = Math.max(maxC + 1, 1);

    let rowsXml = "";
    for (let r = 0; r < rows; r++) {
      let rowCells = "", hasContent = false;
      for (let c = 0; c < cols; c++) {
        const key = `${r}_${c}`;
        const raw = sheet.cells[key];
        const style = sheet.styles && sheet.styles[key];
        const xfIdx = getXfIndex(style);
        if ((raw === undefined || raw === "") && !style) continue;
        hasContent = true;
        const ref = colToLetters(c) + (r + 1);
        const sAttr = xfIdx ? ` s="${xfIdx}"` : "";
        if (raw && typeof raw === "string" && raw[0] === "=") {
          const val = evaluator.getCellDisplay(r, c);
          const numVal = typeof val === "number" ? val : 0;
          rowCells += `<c r="${ref}"${sAttr}><f>${xmlEscape(raw.slice(1))}</f><v>${numVal}</v></c>`;
        } else if (raw && /^-?\d+(\.\d+)?$/.test(raw)) {
          rowCells += `<c r="${ref}"${sAttr}><v>${raw}</v></c>`;
        } else if (raw) {
          rowCells += `<c r="${ref}"${sAttr} t="inlineStr"><is><t xml:space="preserve">${xmlEscape(raw)}</t></is></c>`;
        } else {
          rowCells += `<c r="${ref}"${sAttr}/>`;
        }
      }
      if (hasContent) rowsXml += `<row r="${r + 1}">${rowCells}</row>`;
    }
    const mergeXml = (sheet.merges || []).length
      ? `<mergeCells count="${sheet.merges.length}">${sheet.merges.map((m) => `<mergeCell ref="${colToLetters(m.c1)}${m.r1 + 1}:${colToLetters(m.c2)}${m.r2 + 1}"/>`).join("")}</mergeCells>`
      : "";
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${rowsXml}</sheetData>${mergeXml}</worksheet>`;
  });

  // --- build styles.xml from collected styleList ---
  const fontList = [], fontIndex = new Map();
  const fillList = [{ bg: null }, { bg: null }], fillIndex = new Map(); // 0=none,1=gray125 reserved
  const xfEntries = styleList.map((s) => {
    const fKey = JSON.stringify([s.bold, s.italic, s.color || ""]);
    let fIdx = fontIndex.get(fKey);
    if (fIdx === undefined) { fIdx = fontList.length; fontList.push(s); fontIndex.set(fKey, fIdx); }
    let fillIdx = 0;
    if (s.bg) { const bKey = s.bg; fillIdx = fillIndex.get(bKey); if (fillIdx === undefined) { fillIdx = fillList.length; fillList.push({ bg: s.bg }); fillIndex.set(bKey, fillIdx); } }
    const numFmtId = NUMFMT_TO_ID[s.numFmt || "general"] || 0;
    return { fontId: fIdx, fillId: fillIdx, numFmtId, align: s.align || "" };
  });
  const numFmtsXml = Object.entries(CUSTOM_NUMFMTS).map(([id, code]) => `<numFmt numFmtId="${id}" formatCode="${xmlEscape(code)}"/>`).join("");
  const fontsXml = fontList.map((s) => `<font>${s.bold ? "<b/>" : ""}${s.italic ? "<i/>" : ""}<sz val="11"/>${s.color ? `<color rgb="${colorToRgb(s.color)}"/>` : `<color theme="1"/>`}<name val="Calibri"/></font>`).join("");
  const fillsXml = fillList.map((f) => f.bg ? `<fill><patternFill patternType="solid"><fgColor rgb="${colorToRgb(f.bg)}"/></patternFill></fill>` : `<fill><patternFill patternType="none"/></fill>`).join("");
  const xfsXml = xfEntries.map((xf) => `<xf numFmtId="${xf.numFmtId}" fontId="${xf.fontId}" fillId="${xf.fillId}" borderId="0" xfId="0" applyFont="1" applyFill="1" applyNumberFormat="1"${xf.align ? ` applyAlignment="1"` : ""}>${xf.align ? `<alignment horizontal="${xf.align}"/>` : ""}</xf>`).join("");
  const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="${Object.keys(CUSTOM_NUMFMTS).length}">${numFmtsXml}</numFmts><fonts count="${fontList.length}">${fontsXml}</fonts><fills count="${fillList.length}">${fillsXml}</fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${xfEntries.length}">${xfsXml}</cellXfs></styleSheet>`;

  const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("")}</Types>`;
  const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`;
  const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${sheets.map((s, i) => `<sheet name="${xmlEscape(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join("")}</sheets></workbook>`;
  const workbookRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join("")}<Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`;

  const files = [
    { name: "[Content_Types].xml", data: encoder.encode(contentTypes) },
    { name: "_rels/.rels", data: encoder.encode(rootRels) },
    { name: "xl/workbook.xml", data: encoder.encode(workbookXml) },
    { name: "xl/_rels/workbook.xml.rels", data: encoder.encode(workbookRels) },
    { name: "xl/styles.xml", data: encoder.encode(stylesXml) },
    ...sheetXmls.map((xml, i) => ({ name: `xl/worksheets/sheet${i + 1}.xml`, data: encoder.encode(xml) })),
  ];
  return makeZip(files);
}
export function downloadXlsx(sheets, filename = "spreadsheet.xlsx") {
  const bytes = buildWorkbookXlsx(sheets);
  const blob = new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename; a.click();
  URL.revokeObjectURL(url);
}

/* ================================ Import =================================== */
function parseStylesXml(xml) {
  if (!xml) return [];
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  const numFmts = {};
  Array.from(doc.getElementsByTagName("numFmt")).forEach((n) => { numFmts[n.getAttribute("numFmtId")] = n.getAttribute("formatCode"); });
  const fonts = Array.from(doc.getElementsByTagName("fonts")[0]?.getElementsByTagName("font") || []).map((f) => ({
    bold: !!f.getElementsByTagName("b").length,
    italic: !!f.getElementsByTagName("i").length,
    color: rgbToColor(f.getElementsByTagName("color")[0]?.getAttribute("rgb")),
  }));
  const fills = Array.from(doc.getElementsByTagName("fills")[0]?.getElementsByTagName("fill") || []).map((f) => {
    const fg = f.getElementsByTagName("fgColor")[0];
    return { bg: fg ? rgbToColor(fg.getAttribute("rgb")) : null };
  });
  function numFmtToKey(id) {
    const n = parseInt(id, 10);
    if (n === 0) return "general";
    if (n === 1) return "integer";
    if (n === 2) return "number";
    if (n === 9 || n === 10) return "percent";
    if (n >= 14 && n <= 22) return "date";
    const code = numFmts[id] || "";
    if (code.includes("$")) return "currency";
    if (code.includes("%")) return "percent";
    if (/[ymd]/i.test(code) && code.includes("/")) return "date";
    return "general";
  }
  const xfs = Array.from(doc.getElementsByTagName("cellXfs")[0]?.getElementsByTagName("xf") || []).map((xf) => {
    const fontId = parseInt(xf.getAttribute("fontId") || "0", 10);
    const fillId = parseInt(xf.getAttribute("fillId") || "0", 10);
    const numFmtId = xf.getAttribute("numFmtId") || "0";
    const align = xf.getElementsByTagName("alignment")[0]?.getAttribute("horizontal") || "";
    const font = fonts[fontId] || {};
    const fill = fills[fillId] || {};
    const style = { bold: !!font.bold, italic: !!font.italic, color: font.color || "", bg: fill.bg || "", numFmt: numFmtToKey(numFmtId), align };
    const isDefault = !style.bold && !style.italic && !style.color && !style.bg && style.numFmt === "general" && !style.align;
    return isDefault ? null : style;
  });
  return xfs;
}

export async function readXlsxFile(file) {
  const buf = await file.arrayBuffer();
  const bytes = new Uint8Array(buf);
  const files = unzip(bytes);
  const decoder = new TextDecoder();

  let sharedStrings = [];
  const ssName = Object.keys(files).find((n) => /sharedStrings\.xml$/.test(n));
  if (ssName) {
    const doc = new DOMParser().parseFromString(decoder.decode(files[ssName]), "application/xml");
    sharedStrings = Array.from(doc.getElementsByTagName("si")).map((si) => si.textContent);
  }
  const stylesName = Object.keys(files).find((n) => /styles\.xml$/.test(n));
  const xfStyles = stylesName ? parseStylesXml(decoder.decode(files[stylesName])) : [];

  // sheet order + names from workbook.xml (falls back to filename order)
  const wbName = Object.keys(files).find((n) => /workbook\.xml$/.test(n));
  let sheetMeta = [];
  if (wbName) {
    const doc = new DOMParser().parseFromString(decoder.decode(files[wbName]), "application/xml");
    sheetMeta = Array.from(doc.getElementsByTagName("sheet")).map((s, i) => ({ name: s.getAttribute("name") || `Sheet${i + 1}` }));
  }
  const sheetFiles = Object.keys(files).filter((n) => /^xl\/worksheets\/sheet\d*\.xml$/.test(n))
    .sort((a, b) => (parseInt(a.match(/(\d+)/)[1]) - parseInt(b.match(/(\d+)/)[1])));

  const sheets = sheetFiles.map((fname, i) => {
    const doc = new DOMParser().parseFromString(decoder.decode(files[fname]), "application/xml");
    const cells = {}, styles = {};
    let maxR = 0, maxC = 0;
    for (const c of Array.from(doc.getElementsByTagName("c"))) {
      const ref = c.getAttribute("r");
      if (!ref) continue;
      const m = ref.match(/^([A-Za-z]+)(\d+)$/);
      if (!m) continue;
      const col = lettersToCol(m[1]), row = parseInt(m[2], 10) - 1;
      maxR = Math.max(maxR, row); maxC = Math.max(maxC, col);
      const type = c.getAttribute("t");
      const fNode = c.getElementsByTagName("f")[0];
      const vNode = c.getElementsByTagName("v")[0];
      let value = "";
      if (fNode) value = "=" + fNode.textContent;
      else if (type === "s") { const idx = parseInt(vNode ? vNode.textContent : "0", 10); value = sharedStrings[idx] || ""; }
      else if (type === "inlineStr") { const isNode = c.getElementsByTagName("is")[0]; value = isNode ? isNode.textContent : ""; }
      else value = vNode ? vNode.textContent : "";
      if (value !== "") cells[`${row}_${col}`] = value;
      const sIdx = c.getAttribute("s");
      if (sIdx && xfStyles[parseInt(sIdx, 10)]) styles[`${row}_${col}`] = xfStyles[parseInt(sIdx, 10)];
    }
    const merges = Array.from(doc.getElementsByTagName("mergeCell")).map((m) => {
      const [from, to] = m.getAttribute("ref").split(":");
      const a = from.match(/^([A-Za-z]+)(\d+)$/), b = to.match(/^([A-Za-z]+)(\d+)$/);
      return { r1: parseInt(a[2]) - 1, c1: lettersToCol(a[1]), r2: parseInt(b[2]) - 1, c2: lettersToCol(b[1]) };
    });
    return { name: sheetMeta[i]?.name || `Sheet${i + 1}`, cells, styles, merges, rows: maxR + 1, cols: maxC + 1 };
  });
  return sheets.length ? sheets : [{ name: "Sheet1", cells: {}, styles: {}, merges: [], rows: 1, cols: 1 }];
}

/* ================================== CSV ==================================== */
export function parseCSV(text) {
  const rows = [];
  let row = [], field = "", inQuotes = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (inQuotes) { if (c === '"') { if (text[i + 1] === '"') { field += '"'; i++; } else inQuotes = false; } else field += c; }
    else {
      if (c === '"') inQuotes = true;
      else if (c === ",") { row.push(field); field = ""; }
      else if (c === "\n") { row.push(field); rows.push(row); row = []; field = ""; }
      else if (c === "\r") { /* skip */ }
      else field += c;
    }
  }
  row.push(field); rows.push(row);
  return rows;
}
export function getUsedBounds(cells) {
  let maxR = -1, maxC = -1;
  for (const key in cells) { if (cells[key] === "" || cells[key] == null) continue; const [r, c] = key.split("_").map(Number); if (r > maxR) maxR = r; if (c > maxC) maxC = c; }
  return { rows: maxR + 1, cols: maxC + 1 };
}
export function cellsToCSV(cells, evaluator) {
  const b = getUsedBounds(cells);
  const lines = [];
  for (let r = 0; r < b.rows; r++) {
    const parts = [];
    for (let c = 0; c < b.cols; c++) {
      let v = evaluator.getCellDisplay(r, c);
      v = v === undefined ? "" : String(v);
      if (/[",\n]/.test(v)) v = '"' + v.replace(/"/g, '""') + '"';
      parts.push(v);
    }
    lines.push(parts.join(","));
  }
  return lines.join("\r\n");
}
export function csvToCells(text) {
  const rows = parseCSV(text.replace(/\r\n/g, "\n"));
  const cells = {};
  let maxCols = 0;
  rows.forEach((row, r) => { maxCols = Math.max(maxCols, row.length); row.forEach((val, c) => { if (val !== "") cells[`${r}_${c}`] = val; }); });
  return { cells, rows: rows.length, cols: maxCols };
}
export function downloadCsv(cells, evaluator, filename = "spreadsheet.csv") {
  const csv = cellsToCSV(cells, evaluator);
  const blob = new Blob([csv], { type: "text/csv" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename; a.click();
  URL.revokeObjectURL(url);
}
