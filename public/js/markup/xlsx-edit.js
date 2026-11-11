// Excel editor (GridCraft-style): a ribbon in the floating glass bar (Home, Insert, Formulas, Data, View), Name Box
// and formula bar with coloured references and function help, a virtualised grid with frozen panes, range selection,
// copy/paste to and from Excel, fill handle, sheet tabs and a status bar. ExcelJS reads and writes the .xlsx;
// HyperFormula (GPL-3.0) calculates formulas live. Charts, images and pivot tables are not kept (ExcelJS drops them).
// On a protected sheet, locked cells stay read-only, and relock() (xlsx-relock.js) copies the original protection back.
import { relock } from './xlsx-relock.js';
import { cssZoom, floatBar, fullScreen, viewBar } from './view-tools.js';
import { fileTab, ribbon, RI, setPressed } from './ribbon.js';
import { pairHex } from './hex-color.js';
import { THEMES, currentTheme } from './themes.js';
import {
  addr, adjustFormula, calcContent, cellCss, colName, FN_GROUPS, FN_HELP, fixText, hexOf, mapRefs, parseRange, parseValue,
  rawText, shiftFormula, shownText, themeColors,
} from './xlsx-core.js';

const EXCELJS = 'https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js';
const HYPERFORMULA = 'https://cdn.jsdelivr.net/npm/hyperformula@3.0.1/dist/hyperformula.full.min.js'; // not on cdnjs
const MAX_CALC_CELLS = 500000; // larger workbooks show the results saved in the file instead of live ones
const HDR = 22; const RH = 46; // column header height, row header width (px)
const REF_COLORS = ['#1a73e8', '#d93025', '#188038', '#9334e6', '#e37400', '#12b5cb'];
const FONTS = ['Calibri', 'Aptos', 'Arial', 'Tahoma', 'TH Sarabun New', 'Cordia New', 'Angsana New', 'Noto Sans Thai', 'LINE Seed Sans TH', 'Times New Roman'];
const SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36];
const NUMFMT = [['General', 'General'], ['#,##0', 'Number 1,234'], ['#,##0.00', 'Number 1,234.00'], ['"฿"#,##0.00', 'Baht ฿1,234.00'],
  ['0%', 'Percent 12%'], ['0.00%', 'Percent 12.34%'], ['d/m/yyyy', 'Date 31/12/2026'], ['[$-th-TH,107]d mmm bbbb', 'Thai date 31 ธ.ค. 2569'], ['@', 'Text']];

const script = (src, global) => window[global] ? Promise.resolve(window[global]) : new Promise((res, rej) => {
  const s = document.createElement('script'); s.src = src; s.crossOrigin = 'anonymous';
  s.onload = () => res(window[global]); s.onerror = () => rej(new Error('Could not load the Excel tools. Try again.'));
  document.head.append(s);
});
const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

// Marks the saved file so Excel recalculates every formula when it opens it. Done last, on the XML.
async function recalcOnOpen(buf) {
  const zip = await window.JSZip.loadAsync(buf); // loaded by relock()
  let xml = await zip.file('xl/workbook.xml').async('string');
  xml = /<calcPr\b/.test(xml) ? xml.replace(/<calcPr\b(?![^>]*fullCalcOnLoad)/, '<calcPr fullCalcOnLoad="1"') : xml.replace(/(<\/definedNames>|<\/sheets>)/, '$1<calcPr fullCalcOnLoad="1"/>');
  zip.file('xl/workbook.xml', xml);
  return zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
}
const bsearch = (pre, v, n) => { let lo = 1, hi = n; while (lo < hi) { const m = (lo + hi + 1) >> 1; if (pre[m] <= v) lo = m; else hi = m - 1; } return lo; };

// Opens the editor in `root`. bytes: the .xlsx (ArrayBuffer). Resolves { isDirty(), save() }.
export async function openXlsxEditor(root, bytes) {
  const ExcelJS = await script(EXCELJS, 'ExcelJS');
  const wb = new ExcelJS.Workbook();
  try { await wb.xlsx.load(bytes.slice(0)); } catch { throw new Error('Could not read this file. Only .xlsx files can be edited.'); }
  if (!wb.worksheets.length) throw new Error('This workbook has no sheets.');
  const theme = themeColors(wb);
  let dirty = false;
  // Shared formulas become ordinary ones, so every formula can be shifted, sorted and adjusted on its own.
  wb.worksheets.forEach((ws) => ws.eachRow({ includeEmpty: false }, (row) => row.eachCell({ includeEmpty: false }, (c) => {
    if (c.value && typeof c.value === 'object' && (c.value.sharedFormula || c.value.shareType)) c.value = { formula: c.formula, result: c.result };
  })));

  // ---------- live calculation ----------
  let HF = null; let hf = null; const sheetId = new Map();
  async function buildHF() {
    hf?.destroy?.(); hf = null; sheetId.clear();
    try {
      const cells = wb.worksheets.reduce((n, ws) => n + ws.rowCount * ws.columnCount, 0);
      if (cells > MAX_CALC_CELLS) return;
      HF = HF || await script(HYPERFORMULA, 'HyperFormula'); fixText(HF);
      const data = {};
      wb.worksheets.forEach((ws) => { data[ws.name] = Array.from({ length: ws.rowCount }, (_, r) => Array.from({ length: ws.columnCount }, (__, c) => calcContent(ws.findCell(r + 1, c + 1)))); });
      hf = HF.buildFromSheets(data, { licenseKey: 'gpl-v3' });
      wb.worksheets.forEach((ws) => sheetId.set(ws, hf.getSheetId(ws.name)));
    } catch (e) { hf = null; console.warn('Live calculation is off:', e); }
  }
  await buildHF();
  function result(ws, cell) {
    if (!hf || !cell.formula) return cell.result;
    const v = hf.getCellValue({ sheet: sheetId.get(ws), row: cell.row - 1, col: cell.col - 1 });
    if (v == null || typeof v !== 'object') return v;
    const saved = cell.result;
    return saved != null && !(typeof saved === 'object' && saved.error) ? saved : { error: v.value || '#ERROR' };
  }
  const valueOf = (ws, cell) => (cell.formula ? result(ws, cell) : cell.value);
  const syncHF = (ws, cells) => {
    if (!hf) return;
    const sid = sheetId.get(ws);
    try { hf.batch(() => cells.forEach((c) => hf.setCellContents({ sheet: sid, row: c.row - 1, col: c.col - 1 }, [[calcContent(c)]]))); } catch (e) { console.warn(e); }
  };

  // ---------- state ----------
  let ws = wb.worksheets[0];
  const view = new Map(); // ws -> { sel, fr, fc, filter }
  const vs = () => {
    if (!view.has(ws)) {
      const v = (ws.views || [])[0] || {};
      const af = typeof ws.autoFilter === 'string' ? parseRange(ws.autoFilter) : ws.autoFilter?.from ? parseRange(`${addr(ws.autoFilter.from.row, ws.autoFilter.from.column)}:${addr(ws.autoFilter.to.row, ws.autoFilter.to.column)}`) : null;
      view.set(ws, { sel: { r: 1, c: 1, r2: 1, c2: 1 }, fr: v.state === 'frozen' ? v.ySplit || 0 : 0, fc: v.state === 'frozen' ? v.xSplit || 0 : 0, filter: af ? { ref: af, crit: new Map(), hidden: new Set() } : null });
    }
    return view.get(ws);
  };
  const hist = { undo: [], redo: [] };
  let clip = null; // the last copied range: { ws, r1, c1, rows: [[{ value, style }]], text }
  let geo = null; // geometry of the active sheet

  // ---------- markup ----------
  const spec = [
    fileTab('xe', { open: 'Open an Excel file (.xlsx)', download: 'Download as…' }),
    { tab: 'Home', groups: [
      { label: 'Clipboard', items: [{ btn: 'xe-cut', icon: RI.cut, title: 'Cut (Ctrl+X)' }, { btn: 'xe-copy', icon: RI.copy, title: 'Copy (Ctrl+C)' }, { btn: 'xe-paste', icon: RI.paste, title: 'Paste (Ctrl+V)' }] },
      { label: 'Font', items: [{ select: 'xe-font', title: 'Font', options: FONTS, width: 132 }, { select: 'xe-size', title: 'Font size', options: SIZES.map(String), width: 58 },
        { btn: 'xe-b', icon: RI.bold, title: 'Bold (Ctrl+B)', toggle: true }, { btn: 'xe-i', icon: RI.italic, title: 'Italic (Ctrl+I)', toggle: true }, { btn: 'xe-u', icon: RI.underline, title: 'Underline (Ctrl+U)', toggle: true },
        { color: 'xe-fc', icon: RI.color, title: 'Font colour' }, { color: 'xe-bg', icon: RI.fill, title: 'Fill colour' }] },
      { label: 'Alignment', items: [{ btn: 'xe-al', icon: RI.left, title: 'Align left', toggle: true }, { btn: 'xe-ac', icon: RI.center, title: 'Center', toggle: true }, { btn: 'xe-ar', icon: RI.right, title: 'Align right', toggle: true },
        { btn: 'xe-vt', icon: RI.top, title: 'Top', toggle: true }, { btn: 'xe-vm', icon: RI.middle, title: 'Middle', toggle: true }, { btn: 'xe-vb', icon: RI.bottom, title: 'Bottom', toggle: true },
        { btn: 'xe-wrap', icon: RI.wrap, title: 'Wrap text', toggle: true }, { btn: 'xe-merge', icon: RI.merge, title: 'Merge & center', toggle: true }] },
      { label: 'Number', items: [{ select: 'xe-numfmt', title: 'Number format', options: NUMFMT, width: 150 }, { btn: 'xe-pct', icon: RI.pct, title: 'Percent style' }, { btn: 'xe-comma', icon: RI.comma, title: 'Comma style' },
        { btn: 'xe-decup', icon: RI.decUp, title: 'More decimals' }, { btn: 'xe-decdn', icon: RI.decDown, title: 'Fewer decimals' }] },
      { label: 'Borders', items: [{ select: 'xe-border', title: 'Borders', options: [['', 'Borders…'], ['all', 'All borders'], ['outside', 'Outside borders'], ['thick', 'Thick outside'], ['bottom', 'Bottom border'], ['top', 'Top border'], ['double', 'Double bottom'], ['none', 'No border']], width: 128 }] },
      { label: 'Styles', items: [{ btn: 'xe-table', icon: RI.table, title: 'Format as table: header row and banded rows in a colour theme', text: 'Format as table' }] },
    ] },
    { tab: 'Insert', groups: [
      { label: 'Rows', items: [{ btn: 'xe-rowup', icon: RI.rowAbove, title: 'Insert rows above' }, { btn: 'xe-rowdn', icon: RI.rowBelow, title: 'Insert rows below' }, { btn: 'xe-rowdel', icon: RI.rowDel, title: 'Delete rows' }] },
      { label: 'Columns', items: [{ btn: 'xe-colleft', icon: RI.colLeft, title: 'Insert columns left' }, { btn: 'xe-colright', icon: RI.colRight, title: 'Insert columns right' }, { btn: 'xe-coldel', icon: RI.colDel, title: 'Delete columns' }] },
    ] },
    { tab: 'Formulas', groups: [
      { label: 'Functions', items: [{ btn: 'xe-fnlist', icon: RI.fx, title: 'Insert function', text: 'Insert function' }, { btn: 'xe-autosum', icon: RI.sum, title: 'AutoSum', text: 'AutoSum' }] },
    ] },
    { tab: 'Data', groups: [
      { label: 'Sort & filter', items: [{ btn: 'xe-sortaz', icon: RI.sortAZ, title: 'Sort A to Z (smallest first)' }, { btn: 'xe-sortza', icon: RI.sortZA, title: 'Sort Z to A (largest first)' },
        { btn: 'xe-filter', icon: RI.filter, title: 'Filter', text: 'Filter', toggle: true }, { btn: 'xe-unfilter', icon: RI.filter, title: 'Clear filters', text: 'Clear' }] },
    ] },
    { tab: 'View', groups: [
      { label: 'Freeze panes', items: [{ select: 'xe-freeze', title: 'Freeze panes', options: [['', 'Freeze panes…'], ['row', 'Freeze top row'], ['col', 'Freeze first column'], ['sel', 'Freeze at selection'], ['off', 'Unfreeze']], width: 160 }] },
    ] },
  ];
  const tabs = wb.worksheets.map((s, i) => `<button type="button" role="tab" class="xe-tab" data-i="${i}">${esc(s.name)}</button>`).join('');
  root.innerHTML = `<div class="oe xe"><div class="oe-bar xe-bar"><div class="xe-top"></div>
      <div class="xe-fxrow"><input id="xe-name" class="xe-name" aria-label="Name Box: the selected cell or range; type an address and press Enter to go there" autocomplete="off" spellcheck="false">
        <span class="xe-fxl" aria-hidden="true">fx</span>
        <div class="xe-fxwrap"><div class="xe-fxback" aria-hidden="true"></div><input id="xe-fxin" class="xe-fxin" aria-label="Formula bar" autocomplete="off" spellcheck="false"></div></div></div>
    <div class="oe-scroll xe-grid" id="xe-grid" tabindex="0" role="grid" aria-label="Sheet cells. Arrow keys move; type to edit; F2 edits the cell."><div class="xe-canvas"></div></div>
    <div class="xe-foot"><div class="xe-tabs" role="tablist" aria-label="Sheets">${tabs}</div><div class="xe-status" id="xe-status" role="status"></div></div>
    <div class="xe-pop" id="xe-pop" hidden></div><div class="xe-hint" id="xe-hint" hidden></div></div>`;
  const $ = (s) => root.querySelector(s);
  const box = $('.oe'); const grid = $('#xe-grid'); const canvas = $('.xe-canvas'); const fxin = $('#xe-fxin'); const nameBox = $('#xe-name');
  const rb = ribbon(spec, 'xe', 1); $('.xe-top').append(rb);
  const tools = document.createElement('span'); tools.className = 'oe-tools'; tools.innerHTML = viewBar('xe', undefined, { help: true }); $('.xe-top').append(tools);
  const zoom = cssZoom(box, grid, () => canvas, 'xe');
  fullScreen(box, 'xe', 'oe-full');
  floatBar(box, $('.oe-bar'));
  const syncFc = pairHex($('#xe-fc')); const syncBg = pairHex($('#xe-bg'));
  const status = (s) => { $('#xe-status').dataset.msg = s; showStatus(); clearTimeout(status.t); status.t = setTimeout(() => { delete $('#xe-status').dataset.msg; showStatus(); }, 4000); };

  // ---------- geometry ----------
  const mctx = document.createElement('canvas').getContext('2d');
  function buildGeo() {
    const v = vs();
    const nR = Math.max(ws.rowCount, v.sel.r2, v.sel.r) + 200; const nC = Math.min(16384, Math.max(ws.columnCount, v.sel.c2, v.sel.c) + 26);
    const dW = Math.round((ws.properties.defaultColWidth || 8.43) * 7 + 5); const dH = Math.round((ws.properties.defaultRowHeight || 15) * 4 / 3);
    const colW = new Float64Array(nC + 2); const rowH = new Float64Array(nR + 2);
    for (let c = 1; c <= nC; c++) { const col = c <= ws.columnCount + 1 ? ws.getColumn(c) : null; colW[c] = col?.hidden ? 0 : col?.width ? Math.round(col.width * 7 + 5) : dW; }
    for (let r = 1; r <= nR; r++) {
      const row = r <= ws.rowCount ? ws.findRow(r) : null;
      if (row?.hidden || v.filter?.hidden.has(r)) { rowH[r] = 0; continue; }
      rowH[r] = row?.height ? Math.round(row.height * 4 / 3) : dH;
      if (row && !row.height) rowH[r] = Math.max(rowH[r], autoHeight(row, colW));
    }
    const colX = new Float64Array(nC + 2); const rowY = new Float64Array(nR + 2);
    for (let c = 1; c <= nC + 1; c++) colX[c] = colX[c - 1] + (colW[c - 1] || 0);
    for (let r = 1; r <= nR + 1; r++) rowY[r] = rowY[r - 1] + (rowH[r - 1] || 0);
    // Merged ranges: master -> span, and covered cells.
    const merges = []; const covered = new Set();
    Object.values(ws._merges || {}).forEach((m) => {
      const mm = m.model || m; const r1 = mm.top, c1 = mm.left, r2 = mm.bottom, c2 = mm.right;
      if (r1 == null) return;
      merges.push({ r1, c1, r2, c2 });
      for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) if (r !== r1 || c !== c1) covered.add(`${r}:${c}`);
    });
    geo = { nR, nC, colW, rowH, colX, rowY, merges, covered, dH };
  }
  // Rows with wrapped text and no height of their own grow to fit their text, as in Excel.
  function autoHeight(row, colW) {
    let h = 0;
    row.eachCell({ includeEmpty: false }, (cell) => {
      const t = shownText(cell, valueOf(ws, cell)); if (!t) return;
      const wrap = cell.alignment?.wrapText || t.includes('\n'); if (!wrap) return;
      const pt = cell.font?.size || 11; const px = pt * 4 / 3; mctx.font = `${px}px ${cell.font?.name ? `"${cell.font.name}",` : ''}sans-serif`;
      const w = Math.max(20, (colW[cell.col] || 64) - 8); let lines = 0;
      t.split('\n').forEach((p) => { lines += Math.max(1, Math.ceil(mctx.measureText(p).width / w)); });
      h = Math.max(h, Math.ceil(lines * px * 1.25 + 6));
    });
    return h;
  }
  const merged = (r, c) => geo.merges.find((m) => r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2);

  // ---------- rendering ----------
  // A 3 × 3 grid of areas: headers, frozen rows/columns and the body; frozen areas stay put with position: sticky.
  canvas.innerHTML = ['a00', 'a01', 'a02', 'a10', 'a11', 'a12', 'a20', 'a21', 'a22'].map((a) => `<div class="xa ${a}"></div>`).join('');
  const A = Object.fromEntries([...canvas.children].map((el) => [el.classList[1], el]));
  let frame = 0; const paint = () => { cancelAnimationFrame(frame); frame = requestAnimationFrame(render); };
  function render() {
    const v = vs(); const { colX, rowY, nR, nC } = geo; const z = zoom.get();
    const fr = Math.min(v.fr, nR), fc = Math.min(v.fc, nC);
    const FH = rowY[fr + 1] - rowY[1], FW = colX[fc + 1] - colX[1];
    const BW = colX[nC + 1] - colX[fc + 1], BH = rowY[nR + 1] - rowY[fr + 1];
    canvas.style.gridTemplateColumns = `${RH}px ${FW}px ${BW}px`; canvas.style.gridTemplateRows = `${HDR}px ${FH}px ${BH}px`;
    canvas.style.setProperty('--hdr', `${HDR}px`); canvas.style.setProperty('--rh', `${RH}px`);
    // Visible body window (canvas px, unzoomed).
    const pt = parseFloat(getComputedStyle(grid).paddingTop) || 0;
    const sx = grid.scrollLeft / z, sy = Math.max(0, grid.scrollTop - pt) / z;
    const vw = grid.clientWidth / z - RH - FW, vh = (grid.clientHeight - pt) / z - HDR - FH;
    const r0 = Math.max(fr + 1, bsearch(rowY, rowY[fr + 1] + sy, nR) - 2), r1 = Math.min(nR, bsearch(rowY, rowY[fr + 1] + sy + vh, nR) + 2);
    const c0 = Math.max(fc + 1, bsearch(colX, colX[fc + 1] + sx, nC) - 1), c1 = Math.min(nC, bsearch(colX, colX[fc + 1] + sx + vw, nC) + 1);
    const range = (a, b) => { const o = []; for (let i = a; i <= b; i++) o.push(i); return o; };
    const fRows = range(1, fr), fCols = range(1, fc), bRows = range(r0, r1), bCols = range(c0, c1);
    const s = norm(v.sel);
    A.a00.innerHTML = '<button type="button" class="xh xcorner" aria-label="Select all" data-all="1"></button>';
    A.a01.innerHTML = colHeads(fCols, 1, s); A.a02.innerHTML = colHeads(bCols, fc + 1, s);
    A.a10.innerHTML = rowHeads(fRows, 1, s); A.a20.innerHTML = rowHeads(bRows, fr + 1, s);
    A.a11.innerHTML = cells(fRows, fCols, 1, 1, s); A.a12.innerHTML = cells(fRows, bCols, 1, fc + 1, s);
    A.a21.innerHTML = cells(bRows, fCols, fr + 1, 1, s); A.a22.innerHTML = cells(bRows, bCols, fr + 1, fc + 1, s);
    placeEditor(); drawRefs();
    // Grow the grid when scrolling near its end.
    if (r1 > nR - 40 && nR < 1048576) { geo = null; buildGeo(); paint(); }
  }
  function colHeads(cols, cBase, s) {
    const { colX, colW } = geo;
    return cols.map((c) => `<div class="xh xch${c >= s.c1 && c <= s.c2 ? ' on' : ''}" data-col="${c}" style="left:${colX[c] - colX[cBase]}px;width:${colW[c]}px">${colW[c] ? colName(c) : ''}<span class="rz" data-rzc="${c}"></span></div>`).join('');
  }
  function rowHeads(rows, rBase, s) {
    const { rowY, rowH } = geo;
    return rows.filter((r) => rowH[r]).map((r) => `<div class="xh xrh${r >= s.r1 && r <= s.r2 ? ' on' : ''}" data-row="${r}" style="top:${rowY[r] - rowY[rBase]}px;height:${rowH[r]}px">${r}<span class="rz" data-rzr="${r}"></span></div>`).join('');
  }
  function cells(rows, cols, rBase, cBase, s) {
    if (!rows.length || !cols.length) return '';
    const { colX, rowY, colW, rowH, covered } = geo; const v = vs(); const out = [];
    const prot = !!ws.sheetProtection?.sheet; const filt = v.filter;
    const one = (r, c, w, h) => {
      const cell = ws.findCell(r, c); const val = cell ? valueOf(ws, cell) : null;
      const text = cell ? shownText(cell, val) : '';
      const cls = ['xc'];
      const inSel = r >= s.r1 && r <= s.r2 && c >= s.c1 && c <= s.c2;
      if (inSel) { cls.push('sel'); if (r === s.r1) cls.push('et'); if (r === s.r2) cls.push('eb'); if (c === s.c1) cls.push('el'); if (c === s.c2) cls.push('er'); }
      if (r === v.sel.r && c === v.sel.c) cls.push('act');
      if (prot && cell && cell.protection?.locked !== false) cls.push('lock');
      const err = val && typeof val === 'object' && val.error;
      const css = `left:${colX[c] - colX[cBase]}px;top:${rowY[r] - rowY[rBase]}px;width:${w}px;height:${h}px;${cell ? cellCss(cell, theme, val) : ''}`;
      const fb = filt && r === filt.ref.r1 && c >= filt.ref.c1 && c <= filt.ref.c2 ? `<button type="button" class="xf${filt.crit.has(c) ? ' on' : ''}" data-fc="${c}" aria-label="Filter column ${colName(c)}">▾</button>` : '';
      const handle = r === s.r2 && c === s.c2 ? '<span class="xfill" data-fill="1" aria-hidden="true"></span>' : '';
      out.push(`<div class="${cls.join(' ')}${err ? ' err' : ''}" data-r="${r}" data-c="${c}" style="${css}"><span class="xt">${esc(text)}</span>${fb}${handle}</div>`);
    };
    for (const r of rows) {
      if (!rowH[r]) continue;
      for (const c of cols) {
        if (!colW[c] || covered.has(`${r}:${c}`)) continue;
        const m = merged(r, c);
        if (m) one(r, c, colX[m.c2 + 1] - colX[c], rowY[m.r2 + 1] - rowY[r]); else one(r, c, colW[c], rowH[r]);
      }
    }
    // Merged ranges whose top-left cell is above or left of the window still show their part in it.
    geo.merges.forEach((m) => {
      if (rows.includes(m.r1) && cols.includes(m.c1)) return;
      if (m.r2 < rows[0] || m.r1 > rows.at(-1) || m.c2 < cols[0] || m.c1 > cols.at(-1)) return;
      if (m.r1 < rBase || m.c1 < cBase) return;
      one(m.r1, m.c1, colX[m.c2 + 1] - colX[m.c1], rowY[m.r2 + 1] - rowY[m.r1]);
    });
    return out.join('');
  }
  const norm = (sel) => ({ r1: Math.min(sel.r, sel.r2), r2: Math.max(sel.r, sel.r2), c1: Math.min(sel.c, sel.c2), c2: Math.max(sel.c, sel.c2) });
  // Which area a cell is drawn in, and its box there.
  function cellBox(r, c) {
    const v = vs(); const { colX, rowY, colW, rowH } = geo;
    const area = A[`a${r <= v.fr ? 1 : 2}${c <= v.fc ? 1 : 2}`];
    const rBase = r <= v.fr ? 1 : v.fr + 1; const cBase = c <= v.fc ? 1 : v.fc + 1;
    const m = merged(r, c);
    return { area, x: colX[c] - colX[cBase], y: rowY[r] - rowY[rBase], w: m ? colX[m.c2 + 1] - colX[c] : colW[c], h: m ? rowY[m.r2 + 1] - rowY[r] : rowH[r] };
  }
  function scrollTo(r, c) {
    const v = vs(); const { colX, rowY, colW, rowH } = geo; const z = zoom.get();
    const pt = parseFloat(getComputedStyle(grid).paddingTop) || 0;
    if (r > v.fr) {
      const y = rowY[r] - rowY[v.fr + 1]; const top = Math.max(0, grid.scrollTop - pt) / z; const vh = (grid.clientHeight - pt) / z - HDR - (rowY[v.fr + 1] - rowY[1]);
      if (y < top) grid.scrollTop = y * z + (y ? pt : 0); else if (y + rowH[r] > top + vh) grid.scrollTop = (y + rowH[r] - vh) * z + pt;
    }
    if (c > v.fc) {
      const x = colX[c] - colX[v.fc + 1]; const left = grid.scrollLeft / z; const vw = grid.clientWidth / z - RH - (colX[v.fc + 1] - colX[1]);
      if (x < left) grid.scrollLeft = x * z; else if (x + colW[c] > left + vw) grid.scrollLeft = (x + colW[c] - vw) * z;
    }
  }

  // ---------- selection, ribbon state, status ----------
  function select(r, c, extend = false, keep = false) {
    const v = vs(); const { nR, nC } = geo;
    r = Math.max(1, Math.min(nR, r)); c = Math.max(1, Math.min(nC, c));
    if (extend) { v.sel.r2 = r; v.sel.c2 = c; } else v.sel = { r, c, r2: r, c2: c };
    // A merged cell selects its whole range.
    const m = !extend && merged(r, c); if (m) v.sel = { r: m.r1, c: m.c1, r2: m.r2, c2: m.c2 };
    if (!keep) scrollTo(extend ? r : v.sel.r, extend ? c : v.sel.c);
    refreshBars(); paint();
  }
  function refreshBars() {
    const v = vs(); const s = norm(v.sel); const cell = ws.getCell(v.sel.r, v.sel.c);
    nameBox.value = s.r1 === s.r2 && s.c1 === s.c2 ? addr(s.r1, s.c1) : `${addr(s.r1, s.c1)}:${addr(s.r2, s.c2)}`;
    if (!editing) { fxin.value = rawText(cell); paintFx(); }
    const f = cell.font || {}; const a = cell.alignment || {};
    $('#xe-font').value = FONTS.includes(f.name) ? f.name : 'Calibri'; $('#xe-size').value = String(f.size || 11);
    setPressed(rb, 'xe-b', f.bold); setPressed(rb, 'xe-i', f.italic); setPressed(rb, 'xe-u', f.underline);
    setPressed(rb, 'xe-al', a.horizontal === 'left'); setPressed(rb, 'xe-ac', a.horizontal === 'center'); setPressed(rb, 'xe-ar', a.horizontal === 'right');
    setPressed(rb, 'xe-vt', a.vertical === 'top'); setPressed(rb, 'xe-vm', a.vertical === 'middle'); setPressed(rb, 'xe-vb', !a.vertical || a.vertical === 'bottom');
    setPressed(rb, 'xe-wrap', a.wrapText); setPressed(rb, 'xe-merge', !!merged(v.sel.r, v.sel.c)); setPressed(rb, 'xe-filter', !!v.filter);
    const fmt = NUMFMT.find(([k]) => k === (cell.numFmt || 'General')); $('#xe-numfmt').value = fmt ? fmt[0] : 'General';
    $('#xe-fc').value = hexOf(f.color, theme) || '#000000'; syncFc();
    $('#xe-bg').value = cell.fill?.type === 'pattern' && cell.fill.pattern !== 'none' ? hexOf(cell.fill.fgColor, theme) || '#ffffff' : '#ffffff'; syncBg();
    $('#xe-undo').disabled = !hist.undo.length; $('#xe-redo').disabled = !hist.redo.length;
    root.querySelectorAll('.xe-tab').forEach((t) => t.setAttribute('aria-selected', String(wb.worksheets[+t.dataset.i] === ws)));
    showStatus();
  }
  function showStatus() {
    const el = $('#xe-status'); if (!geo) return;
    if (el.dataset.msg) { el.textContent = el.dataset.msg; return; }
    const s = norm(vs().sel); const r2 = Math.min(s.r2, ws.rowCount), c2 = Math.min(s.c2, ws.columnCount);
    let n = 0, sum = 0, cnt = 0; const cells = (r2 - s.r1 + 1) * (c2 - s.c1 + 1);
    if (cells > 1 && cells <= 300000) {
      for (let r = s.r1; r <= r2; r++) for (let c = s.c1; c <= c2; c++) {
        const cell = ws.findCell(r, c); if (!cell) continue; const v = valueOf(ws, cell);
        if (v != null && v !== '') cnt++; if (typeof v === 'number') { n++; sum += v; }
      }
    }
    const f = (x) => x.toLocaleString('en-US', { maximumFractionDigits: 2 });
    el.textContent = cells > 1 ? (n ? `Average ${f(sum / n)} · Count ${cnt} · Sum ${f(sum)}` : `Count ${cnt}`) : (hf ? 'Formulas calculate live' : 'Live calculation is off for this file: formulas show the results saved in it');
  }

  // ---------- changes and history ----------
  const isLocked = (cell) => !!ws.sheetProtection?.sheet && cell.protection?.locked !== false;
  const snap = (cell) => ({ value: cell.value == null ? null : structuredClone(cell.value), style: structuredClone(cell.style || {}) });
  const restore = (cell, s) => { cell.value = s.value == null ? null : structuredClone(s.value); cell.style = structuredClone(s.style); };
  // Changes cells of the active sheet: mutate(cell, r, c) for each address; one undo step.
  function change(list, mutate, label = '') {
    const items = []; const cellsDone = [];
    for (const [r, c] of list) {
      const cell = ws.getCell(r, c); if (isLocked(cell)) continue;
      const before = snap(cell); mutate(cell, r, c); items.push({ r, c, before, after: snap(cell) }); cellsDone.push(cell);
    }
    if (!items.length) { if (list.length) status('Locked cells on a protected sheet cannot be changed.'); return; }
    syncHF(ws, cellsDone); dirty = true;
    hist.undo.push({ kind: 'cells', ws, items, label }); hist.redo = [];
    afterChange();
  }
  // Structural changes (rows, columns, merges, sort with formulas, freeze): the whole workbook model is the undo step.
  const models = () => wb.worksheets.map((s) => structuredClone(s.model));
  async function structural(fn, label) {
    const before = models(); fn(); const after = models();
    hist.undo.push({ kind: 'models', before, after, label }); hist.redo = []; dirty = true;
    await buildHF(); afterChange(true);
  }
  function afterChange(full = false) { const keepSel = vs().sel; if (full) view.forEach((v, s) => { if (s !== ws) view.delete(s); }); buildGeo(); vs().sel = keepSel; refreshBars(); paint(); }
  async function step(dir) {
    if (editing) commitEdit();
    const h = hist[dir].pop(); if (!h) return;
    hist[dir === 'undo' ? 'redo' : 'undo'].push(h);
    if (h.kind === 'cells') {
      if (ws !== h.ws) showSheet(wb.worksheets.indexOf(h.ws));
      const cs = h.items.map(({ r, c, before, after }) => { const cell = ws.getCell(r, c); restore(cell, dir === 'undo' ? before : after); return cell; });
      syncHF(ws, cs); afterChange();
      const s = h.items[0]; if (s) select(s.r, s.c, false);
    } else {
      const m = dir === 'undo' ? h.before : h.after;
      wb.worksheets.forEach((s, i) => { s.model = structuredClone(m[i]); });
      await buildHF(); afterChange(true);
    }
    dirty = true; status(`${dir === 'undo' ? 'Undone' : 'Redone'}${h.label ? `: ${h.label}` : ''}`);
  }
  const selCells = (clipToUsed = true) => {
    const s = norm(vs().sel); const r2 = clipToUsed ? Math.min(s.r2, Math.max(ws.rowCount, s.r1)) : s.r2; const c2 = clipToUsed ? Math.min(s.c2, Math.max(ws.columnCount, s.c1)) : s.c2;
    const out = []; for (let r = s.r1; r <= r2; r++) for (let c = s.c1; c <= c2; c++) if (!geo.covered.has(`${r}:${c}`)) out.push([r, c]);
    return out;
  };

  // ---------- editing ----------
  let editing = null; // { r, c, ed (textarea), orig, mode: 'enter' | 'edit' }
  const edEl = document.createElement('textarea'); edEl.className = 'xe-ed'; edEl.spellcheck = false; edEl.setAttribute('aria-label', 'Cell editor');
  function startEdit(mode, initial = null) {
    const v = vs(); const cell = ws.getCell(v.sel.r, v.sel.c);
    if (isLocked(cell)) { status('This cell is locked on a protected sheet.'); return; }
    editing = { r: v.sel.r, c: v.sel.c, orig: rawText(cell), mode };
    edEl.value = initial != null ? initial : editing.orig;
    edEl.style.cssText = cellCss(cell, theme, null).replace(/justify-content[^;]*;|align-items[^;]*;/g, '');
    placeEditor(); edEl.focus(); edEl.setSelectionRange(edEl.value.length, edEl.value.length);
    fxin.value = edEl.value; paintFx(); drawRefs(); fnHelp(edEl);
  }
  function placeEditor() {
    if (!editing) { edEl.remove(); return; }
    const b = cellBox(editing.r, editing.c);
    if (edEl.parentNode !== b.area) b.area.append(edEl);
    Object.assign(edEl.style, { left: `${b.x}px`, top: `${b.y}px`, minWidth: `${b.w}px`, minHeight: `${b.h}px` });
  }
  function commitEdit(move = null) {
    if (!editing) return;
    const { r, c, orig } = editing; const text = edEl.value.replace(/\n$/, '');
    editing = null; edEl.remove(); hidePop(); hideHint();
    if (text !== orig) {
      const ws0 = ws;
      change([[r, c]], (cell) => {
        // A typed number keeps a text cell's number format; a date gets a date format.
        const v = parseValue(text); cell.value = v;
        if (v instanceof Date && !/[dmy]/i.test(cell.numFmt || '')) cell.numFmt = 'd/m/yyyy';
      }, `edit ${addr(r, c)}`);
      if (ws0 !== ws) return;
    }
    if (move) select(r + move[0], c + move[1]); else { refreshBars(); paint(); }
    grid.focus({ preventScroll: true });
  }
  function cancelEdit() { if (!editing) return; editing = null; edEl.remove(); hidePop(); hideHint(); refreshBars(); paint(); grid.focus({ preventScroll: true }); }
  edEl.addEventListener('keydown', (e) => {
    if (fnKey(e, edEl)) return;
    if (e.key === 'Enter' && e.altKey) { e.preventDefault(); const p = edEl.selectionStart; edEl.setRangeText('\n', p, edEl.selectionEnd, 'end'); syncFromEd(); return; }
    if (e.key === 'Enter') { e.preventDefault(); commitEdit(e.shiftKey ? [-1, 0] : [1, 0]); }
    else if (e.key === 'Tab') { e.preventDefault(); commitEdit(e.shiftKey ? [0, -1] : [0, 1]); }
    else if (e.key === 'Escape') { e.preventDefault(); cancelEdit(); }
    else if (editing?.mode === 'enter' && e.key.startsWith('Arrow') && !refPickPos(edEl)) {
      e.preventDefault(); commitEdit({ ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[e.key]);
    }
  });
  const syncFromEd = () => { fxin.value = edEl.value; paintFx(); drawRefs(); fnHelp(edEl); };
  edEl.addEventListener('input', syncFromEd);
  // The formula bar edits the active cell too.
  fxin.addEventListener('focus', () => { if (!editing) { startEdit('edit'); fxin.focus(); } });
  fxin.addEventListener('input', () => { if (!editing) startEdit('edit', fxin.value); edEl.value = fxin.value; paintFx(); drawRefs(); fnHelp(fxin); });
  fxin.addEventListener('keydown', (e) => {
    if (fnKey(e, fxin)) return;
    if (e.key === 'Enter') { e.preventDefault(); edEl.value = fxin.value; commitEdit([1, 0]); }
    else if (e.key === 'Escape') { e.preventDefault(); cancelEdit(); }
  });
  fxin.addEventListener('scroll', () => { $('.xe-fxback').scrollLeft = fxin.scrollLeft; });
  // Typing a formula: is the caret where a cell reference may go (after = ( , operators)?
  const refPickPos = (inp) => inp.value.startsWith('=') && /[=(,+\-*/^&<>:;]\s*$/.test(inp.value.slice(0, inp.selectionStart));
  // Coloured references: in the formula bar text, and as boxes on the grid.
  function formulaRefs(text) {
    const out = []; if (!text.startsWith('=')) return out;
    mapRefs(text, (sh, a, b) => { if (sh == null || sh === ws.name) out.push({ r1: Math.min(a.row, b?.row ?? a.row), c1: Math.min(a.col, b?.col ?? a.col), r2: Math.max(a.row, b?.row ?? a.row), c2: Math.max(a.col, b?.col ?? a.col) }); return null; });
    return out;
  }
  function paintFx() {
    const t = fxin.value; const back = $('.xe-fxback');
    if (!t.startsWith('=')) { back.textContent = t; return; }
    let i = 0;
    back.innerHTML = t.split(/("(?:[^"]|"")*")/).map((part, k) => (k % 2 ? esc(part) : esc(part).replace(/('(?:[^']|'')+'!|[A-Za-z_][\w.]*!)?\$?[A-Z]{1,3}\$?\d{1,7}(?::\$?[A-Z]{1,3}\$?\d{1,7})?(?![\w(])/g, (m) => `<span style="color:${REF_COLORS[i++ % REF_COLORS.length]}">${m}</span>`))).join('');
  }
  function drawRefs() {
    canvas.querySelectorAll('.xe-refbox').forEach((e) => e.remove());
    if (!editing) return;
    formulaRefs(edEl.value).forEach((ref, i) => {
      const a = cellBox(ref.r1, ref.c1); const { colX, rowY } = geo;
      const box1 = document.createElement('div'); box1.className = 'xe-refbox';
      Object.assign(box1.style, { left: `${a.x}px`, top: `${a.y}px`, width: `${colX[ref.c2 + 1] - colX[ref.c1]}px`, height: `${rowY[ref.r2 + 1] - rowY[ref.r1]}px`, borderColor: REF_COLORS[i % REF_COLORS.length], background: `${REF_COLORS[i % REF_COLORS.length]}14` });
      a.area.append(box1);
    });
  }

  // ---------- function help ----------
  const FN_ALL = (() => { try { return (HF?.getRegisteredFunctionNames?.('enGB') || Object.keys(FN_HELP)).slice().sort(); } catch { return Object.keys(FN_HELP).sort(); } })();
  const pop = $('#xe-pop'); const hint = $('#xe-hint'); let popItems = []; let popIdx = 0; let popFor = null;
  const hidePop = () => { pop.hidden = true; popItems = []; popFor = null; };
  const hideHint = () => { hint.hidden = true; };
  const placeUnder = (el, inp) => { const r = inp.getBoundingClientRect(); const b = box.getBoundingClientRect(); el.style.left = `${Math.min(r.left - b.left, b.width - 320)}px`; el.style.top = `${r.bottom - b.top + 4}px`; };
  function fnHelp(inp) {
    const t = inp.value; const pos = inp.selectionStart ?? t.length; const before = t.slice(0, pos);
    if (!t.startsWith('=')) { hidePop(); hideHint(); return; }
    const m = /(?:^=|[(,+\-*/^&<>=:;\s])([A-Za-z][A-Za-z0-9.]*)$/.exec(before);
    if (m && m[1].length >= 1) {
      const q = m[1].toUpperCase(); popItems = FN_ALL.filter((n) => n.startsWith(q)).slice(0, 8);
      if (popItems.length) {
        popIdx = 0; popFor = { inp, start: pos - m[1].length, end: pos };
        pop.innerHTML = popItems.map((n, i) => `<button type="button" class="xe-popi${i ? '' : ' on'}" data-fn="${n}"><b>${n}</b>${FN_HELP[n] ? `<span>${esc(FN_HELP[n].slice(n.length))}</span>` : ''}</button>`).join('');
        pop.hidden = false; placeUnder(pop, inp);
      } else hidePop();
    } else hidePop();
    // Argument hint: the innermost open function before the caret, with the current argument in bold.
    let depth = 0, arg = 0, name = null;
    for (let i = before.length - 1; i >= 0; i--) {
      const ch = before[i];
      if (ch === ')') depth++; else if (ch === '(') { if (depth) depth--; else { name = /([A-Za-z][A-Za-z0-9.]*)$/.exec(before.slice(0, i))?.[1]?.toUpperCase(); break; } } else if (ch === ',' && !depth) arg++;
    }
    const sig = name && FN_HELP[name];
    if (sig) {
      const parts = sig.slice(name.length + 1, -1).split(', ');
      hint.innerHTML = `${name}(${parts.map((p, i) => (i === Math.min(arg, parts.length - 1) ? `<b>${esc(p)}</b>` : esc(p))).join(', ')})`;
      hint.hidden = false; placeUnder(hint, inp); if (!pop.hidden) hint.style.top = `${parseFloat(pop.style.top) + pop.offsetHeight + 4}px`;
    } else hideHint();
  }
  function pickFn(n) {
    const { inp, start, end } = popFor; inp.setRangeText(`${n}(`, start, end, 'end');
    if (inp === edEl) syncFromEd(); else { edEl.value = fxin.value; paintFx(); drawRefs(); fnHelp(fxin); }
    hidePop(); inp.focus();
  }
  function fnKey(e, inp) {
    if (pop.hidden || popFor?.inp !== inp) return false;
    if (e.key === 'ArrowDown' || e.key === 'ArrowUp') { e.preventDefault(); popIdx = (popIdx + (e.key === 'ArrowDown' ? 1 : popItems.length - 1)) % popItems.length; pop.querySelectorAll('.xe-popi').forEach((b, i) => b.classList.toggle('on', i === popIdx)); return true; }
    if (e.key === 'Tab' || (e.key === 'Enter' && popItems.length)) { e.preventDefault(); pickFn(popItems[popIdx]); return true; }
    if (e.key === 'Escape') { e.preventDefault(); hidePop(); return true; }
    return false;
  }
  pop.addEventListener('pointerdown', (e) => e.preventDefault());
  pop.addEventListener('click', (e) => {
    const f = e.target.closest('[data-fn]'); if (f && popFor) { pickFn(f.dataset.fn); return; }
    const ins = e.target.closest('[data-ins]'); if (ins) { hidePop(); startEdit('edit', `=${ins.dataset.ins}(`); }
  });
  // Formulas → Insert function: a searchable list grouped by category.
  function functionList() {
    const groups = { ...FN_GROUPS, 'All functions': FN_ALL };
    pop.innerHTML = `<input class="xe-popq" placeholder="Search functions…" aria-label="Search functions" autocomplete="off"><div class="xe-poplist">${Object.entries(groups).map(([g, list]) => `<div class="xe-popg">${g}</div>${list.map((n) => `<button type="button" class="xe-popi" data-ins="${n}"><b>${n}</b>${FN_HELP[n] ? `<span>${esc(FN_HELP[n].slice(n.length))}</span>` : ''}</button>`).join('')}`).join('')}</div>`;
    pop.hidden = false; popFor = null; placeUnder(pop, $('#xe-fnlist'));
    const q = pop.querySelector('.xe-popq'); q.focus();
    q.oninput = () => { const s = q.value.trim().toUpperCase(); pop.querySelectorAll('.xe-popi').forEach((b) => { b.hidden = s && !b.dataset.ins.includes(s); }); pop.querySelectorAll('.xe-popg').forEach((g) => { g.hidden = !!s; }); };
    q.onkeydown = (e) => { if (e.key === 'Escape') { hidePop(); grid.focus(); } };
  }
  document.addEventListener('pointerdown', (e) => { if (!root.isConnected) return; if (!pop.hidden && !pop.contains(e.target) && !e.target.closest('#xe-fnlist') && popFor == null) hidePop(); });

  // ---------- pointer: select, edit, fill, resize, headers ----------
  let drag = null;
  const capture = (e) => { try { canvas.setPointerCapture(e.pointerId); } catch { /* pointer already gone */ } };
  canvas.addEventListener('pointerdown', (e) => {
    if (e.button !== 0) return;
    const t = e.target;
    if (t === edEl) return;
    if (t.closest('.xf')) { e.preventDefault(); filterMenu(+t.closest('.xf').dataset.fc); return; }
    const rz = t.closest('[data-rzc],[data-rzr]');
    if (rz) { e.preventDefault(); const isC = !!rz.dataset.rzc; const i = +(rz.dataset.rzc || rz.dataset.rzr); drag = { kind: 'resize', isC, i, x: e.clientX, y: e.clientY, size: isC ? geo.colW[i] : geo.rowH[i] }; capture(e); return; }
    if (t.closest('[data-fill]')) { e.preventDefault(); const s = norm(vs().sel); drag = { kind: 'fill', src: s, to: { ...s } }; capture(e); return; }
    if (t.closest('[data-all]')) { vs().sel = { r: 1, c: 1, r2: geo.nR, c2: geo.nC }; refreshBars(); paint(); grid.focus(); return; }
    const ch = t.closest('[data-col]'); const rh = t.closest('[data-row]');
    if (ch || rh) {
      if (editing) commitEdit();
      const v = vs(); const i = +(ch?.dataset.col || rh?.dataset.row);
      if (e.shiftKey) { if (ch) v.sel.c2 = i; else v.sel.r2 = i; }
      else v.sel = ch ? { r: 1, c: i, r2: geo.nR, c2: i } : { r: i, c: 1, r2: i, c2: geo.nC };
      drag = { kind: ch ? 'cols' : 'rows' }; capture(e); refreshBars(); paint(); grid.focus({ preventScroll: true }); return;
    }
    const cel = t.closest('.xc'); if (!cel) return;
    const r = +cel.dataset.r, c = +cel.dataset.c;
    // While typing a formula, a click puts the cell's address into it.
    if (editing && refPickPos(document.activeElement === fxin ? fxin : edEl)) {
      e.preventDefault(); const inp = document.activeElement === fxin ? fxin : edEl;
      inp.setRangeText(addr(r, c), inp.selectionStart, inp.selectionEnd, 'end'); if (inp === edEl) syncFromEd(); else { edEl.value = fxin.value; paintFx(); drawRefs(); }
      drag = { kind: 'ref', inp, start: inp.selectionStart - addr(r, c).length, r, c }; capture(e); return;
    }
    if (editing) commitEdit();
    // Touch: a drag that starts on the active cell selects a range; elsewhere the finger scrolls.
    if (e.pointerType === 'touch' && !cel.classList.contains('act')) { select(r, c, false, true); return; }
    select(r, c, e.shiftKey, true); grid.focus({ preventScroll: true });
    drag = { kind: 'sel' }; capture(e);
  });
  const cellAt = (x, y) => { const el = document.elementFromPoint(x, y)?.closest('.xc,[data-col],[data-row]'); return el ? { r: +(el.dataset.r || el.dataset.row || 0), c: +(el.dataset.c || el.dataset.col || 0) } : null; };
  canvas.addEventListener('pointermove', (e) => {
    if (!drag) return;
    if (drag.kind === 'resize') {
      const z = zoom.get(); const d = (drag.isC ? e.clientX - drag.x : e.clientY - drag.y) / z; const size = Math.max(drag.isC ? 8 : 6, drag.size + d);
      if (drag.isC) geo.colW[drag.i] = size; else geo.rowH[drag.i] = size;
      const pre = drag.isC ? geo.colX : geo.rowY; const arr = drag.isC ? geo.colW : geo.rowH; const n = drag.isC ? geo.nC : geo.nR;
      for (let k = drag.i + 1; k <= n + 1; k++) pre[k] = pre[k - 1] + (arr[k - 1] || 0);
      drag.now = size; paint(); return;
    }
    const at = cellAt(e.clientX, e.clientY); if (!at) return;
    const v = vs();
    if (drag.kind === 'sel') { if (at.r && at.c && (at.r !== v.sel.r2 || at.c !== v.sel.c2)) { v.sel.r2 = at.r; v.sel.c2 = at.c; refreshBars(); paint(); } }
    else if (drag.kind === 'cols' && at.c) { v.sel.c2 = at.c; refreshBars(); paint(); }
    else if (drag.kind === 'rows' && at.r) { v.sel.r2 = at.r; refreshBars(); paint(); }
    else if (drag.kind === 'ref' && at.r && at.c) {
      const a = addr(Math.min(drag.r, at.r), Math.min(drag.c, at.c)); const b = addr(Math.max(drag.r, at.r), Math.max(drag.c, at.c));
      const text = a === b ? a : `${a}:${b}`; const inp = drag.inp; inp.setRangeText(text, drag.start, inp.selectionEnd, 'end');
      if (inp === edEl) syncFromEd(); else { edEl.value = fxin.value; paintFx(); drawRefs(); }
    } else if (drag.kind === 'fill' && at.r && at.c) {
      const s = drag.src; const dr = at.r > s.r2 ? at.r - s.r2 : at.r < s.r1 ? at.r - s.r1 : 0; const dc = at.c > s.c2 ? at.c - s.c2 : at.c < s.c1 ? at.c - s.c1 : 0;
      // Fill goes one way: along the longer drag.
      drag.to = Math.abs(dr) >= Math.abs(dc) ? { ...s, r1: Math.min(s.r1, at.r), r2: Math.max(s.r2, at.r) } : { ...s, c1: Math.min(s.c1, at.c), c2: Math.max(s.c2, at.c) };
      canvas.querySelectorAll('.xc.fillp').forEach((x) => x.classList.remove('fillp'));
      canvas.querySelectorAll('.xc').forEach((x) => { const r = +x.dataset.r, c = +x.dataset.c; if (r >= drag.to.r1 && r <= drag.to.r2 && c >= drag.to.c1 && c <= drag.to.c2) x.classList.add('fillp'); });
    }
  });
  canvas.addEventListener('pointerup', () => {
    const d = drag; drag = null; if (!d) return;
    if (d.kind === 'resize' && d.now != null) {
      const sizes = norm(vs().sel); const all = d.isC ? (sizes.r1 === 1 && sizes.r2 >= geo.nR - 1 && d.i >= sizes.c1 && d.i <= sizes.c2 ? range(sizes.c1, sizes.c2) : [d.i]) : (sizes.c1 === 1 && sizes.c2 >= geo.nC - 1 && d.i >= sizes.r1 && d.i <= sizes.r2 ? range(sizes.r1, sizes.r2) : [d.i]);
      structural(() => all.forEach((i) => { if (d.isC) ws.getColumn(i).width = Math.max(0.5, (d.now - 5) / 7); else ws.getRow(i).height = d.now * 0.75; }), d.isC ? 'column width' : 'row height');
    }
    if (d.kind === 'fill') fill(d.src, d.to);
  });
  const range = (a, b) => { const o = []; for (let i = a; i <= b; i++) o.push(i); return o; };
  canvas.addEventListener('dblclick', (e) => { const cel = e.target.closest('.xc'); if (cel && !editing) { select(+cel.dataset.r, +cel.dataset.c, false, true); startEdit('edit'); } });
  grid.addEventListener('scroll', paint, { passive: true });
  new ResizeObserver(paint).observe(grid);
  new ResizeObserver(paint).observe(canvas);

  // ---------- keyboard ----------
  grid.addEventListener('keydown', (e) => {
    if (e.target !== grid) return;
    const v = vs(); const mod = e.ctrlKey || e.metaKey; const k = e.key;
    const moves = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
    if (moves[k]) {
      e.preventDefault(); const [dr, dc] = moves[k];
      const from = e.shiftKey ? { r: v.sel.r2, c: v.sel.c2 } : { r: v.sel.r, c: v.sel.c };
      const to = mod ? edge(from.r, from.c, dr, dc) : { r: from.r + dr, c: from.c + dc };
      select(to.r, to.c, e.shiftKey); return;
    }
    if (k === 'Enter' || k === 'Tab') { e.preventDefault(); const d = k === 'Enter' ? [e.shiftKey ? -1 : 1, 0] : [0, e.shiftKey ? -1 : 1]; select(v.sel.r + d[0], v.sel.c + d[1]); return; }
    if (k === 'PageDown' || k === 'PageUp') { e.preventDefault(); const n = Math.max(1, Math.round(grid.clientHeight / zoom.get() / geo.dH) - 2); select(v.sel.r + (k === 'PageDown' ? n : -n), v.sel.c, e.shiftKey); return; }
    if (k === 'Home') { e.preventDefault(); select(mod ? 1 : v.sel.r, 1, e.shiftKey); return; }
    if (k === 'F2') { e.preventDefault(); startEdit('edit'); return; }
    if (k === 'Delete' || k === 'Backspace') { e.preventDefault(); clearSel(); return; }
    if (mod && k.toLowerCase() === 'a') { e.preventDefault(); v.sel = { r: 1, c: 1, r2: geo.nR, c2: geo.nC }; refreshBars(); paint(); return; }
    if (mod && (k.toLowerCase() === 'z' || k.toLowerCase() === 'y')) { e.preventDefault(); step(k.toLowerCase() === 'y' || e.shiftKey ? 'redo' : 'undo'); return; }
    if (mod && ['b', 'i', 'u'].includes(k.toLowerCase())) { e.preventDefault(); toggleFont({ b: 'bold', i: 'italic', u: 'underline' }[k.toLowerCase()]); return; }
    if (!mod && !e.altKey && k.length === 1) { e.preventDefault(); startEdit('enter', k); }
  });
  // Ctrl+arrow: to the edge of the data block, as in Excel.
  function edge(r, c, dr, dc) {
    const filled = (rr, cc) => { const x = ws.findCell(rr, cc); return x && x.value != null && x.value !== ''; };
    const maxR = Math.max(ws.rowCount, 1), maxC = Math.max(ws.columnCount, 1);
    const inb = (rr, cc) => rr >= 1 && cc >= 1 && rr <= maxR && cc <= maxC;
    let rr = r + dr, cc = c + dc; if (!inb(rr, cc)) return { r: Math.max(1, Math.min(maxR, rr)), c: Math.max(1, Math.min(maxC, cc)) };
    if (filled(r, c) && filled(rr, cc)) { while (inb(rr + dr, cc + dc) && filled(rr + dr, cc + dc)) { rr += dr; cc += dc; } return { r: rr, c: cc }; }
    while (inb(rr, cc) && !filled(rr, cc)) { rr += dr; cc += dc; }
    return inb(rr, cc) ? { r: rr, c: cc } : { r: Math.max(1, Math.min(maxR, rr - dr)), c: Math.max(1, Math.min(maxC, cc - dc)) };
  }
  function clearSel() { change(selCells(), (cell) => { cell.value = null; }, 'clear'); }

  // ---------- clipboard ----------
  const tsvCell = (s) => (/[\t\n"]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s);
  function copySel(e, cut = false) {
    const s = norm(vs().sel); const r2 = Math.min(s.r2, Math.max(ws.rowCount, s.r1)), c2 = Math.min(s.c2, Math.max(ws.columnCount, s.c1));
    const rows = []; const lines = []; const html = [];
    for (let r = s.r1; r <= r2; r++) {
      const row = []; const line = []; const tr = [];
      for (let c = s.c1; c <= c2; c++) { const cell = ws.getCell(r, c); row.push(snap(cell)); const t = shownText(cell, valueOf(ws, cell)); line.push(tsvCell(t)); tr.push(`<td style="${cellCss(cell, theme, valueOf(ws, cell)).replace(/"/g, "'")}">${esc(t)}</td>`); }
      rows.push(row); lines.push(line.join('\t')); html.push(`<tr>${tr.join('')}</tr>`);
    }
    const text = lines.join('\n');
    clip = { ws, r1: s.r1, c1: s.c1, rows, text };
    if (e?.clipboardData) { e.clipboardData.setData('text/plain', text); e.clipboardData.setData('text/html', `<table>${html.join('')}</table>`); e.preventDefault(); }
    else navigator.clipboard?.writeText(text).catch(() => {});
    if (cut) clearSel();
    status(`${cut ? 'Cut' : 'Copied'} ${addr(s.r1, s.c1)}${r2 > s.r1 || c2 > s.c1 ? `:${addr(r2, c2)}` : ''}`);
  }
  function parseTsv(t) {
    const rows = [[]]; let cur = ''; let q = false;
    for (let i = 0; i < t.length; i++) {
      const ch = t[i];
      if (q) { if (ch === '"' && t[i + 1] === '"') { cur += '"'; i++; } else if (ch === '"') q = false; else cur += ch; }
      else if (ch === '"' && cur === '') q = true; else if (ch === '\t') { rows.at(-1).push(cur); cur = ''; }
      else if (ch === '\n' || ch === '\r') { if (ch === '\r' && t[i + 1] === '\n') i++; rows.at(-1).push(cur); cur = ''; rows.push([]); } else cur += ch;
    }
    rows.at(-1).push(cur); if (rows.length > 1 && rows.at(-1).length === 1 && rows.at(-1)[0] === '') rows.pop();
    return rows;
  }
  function pasteText(text) {
    const v = vs(); const r0 = Math.min(v.sel.r, v.sel.r2), c0 = Math.min(v.sel.c, v.sel.c2);
    // Our own copy: values with formats; formulas move with the cells.
    if (clip && text === clip.text) {
      const list = []; clip.rows.forEach((row, i) => row.forEach((_, j) => list.push([r0 + i, c0 + j])));
      const dr = r0 - clip.r1, dc = c0 - clip.c1;
      change(list, (cell, r, c) => {
        const s = clip.rows[r - r0][c - c0]; const val = structuredClone(s.value);
        if (val && typeof val === 'object' && val.formula) { val.formula = shiftFormula(val.formula, dr, dc); delete val.result; }
        cell.value = val; cell.style = structuredClone(s.style);
      }, 'paste');
    } else {
      const rows = parseTsv(text); const list = []; rows.forEach((row, i) => row.forEach((_, j) => list.push([r0 + i, c0 + j])));
      change(list, (cell, r, c) => { cell.value = parseValue(rows[r - r0][c - c0]); }, 'paste');
    }
    const h = hist.undo.at(-1); if (h?.items?.length) { const last = h.items.at(-1); vs().sel = { r: r0, c: c0, r2: last.r, c2: last.c }; refreshBars(); paint(); }
  }
  const gridHasFocus = () => root.isConnected && (document.activeElement === grid);
  document.addEventListener('copy', (e) => { if (gridHasFocus()) copySel(e); });
  document.addEventListener('cut', (e) => { if (gridHasFocus()) copySel(e, true); });
  document.addEventListener('paste', (e) => { if (!gridHasFocus()) return; e.preventDefault(); pasteText(e.clipboardData.getData('text/plain')); });
  $('#xe-copy').onclick = () => copySel(null); $('#xe-cut').onclick = () => copySel(null, true);
  $('#xe-paste').onclick = async () => { try { pasteText(await navigator.clipboard.readText()); } catch { status('Paste with Ctrl+V (the browser blocked reading the clipboard).'); } };

  // ---------- fill handle ----------
  function fill(src, to) {
    canvas.querySelectorAll('.xc.fillp').forEach((x) => x.classList.remove('fillp'));
    if (to.r1 === src.r1 && to.r2 === src.r2 && to.c1 === src.c1 && to.c2 === src.c2) return;
    const down = to.r1 !== src.r1 || to.r2 !== src.r2; const list = [];
    for (let r = to.r1; r <= to.r2; r++) for (let c = to.c1; c <= to.c2; c++) if (r < src.r1 || r > src.r2 || c < src.c1 || c > src.c2) list.push([r, c]);
    // A run of numbers (or dates) with an even step continues; anything else repeats, formulas moving with the cells.
    const series = (vals) => { if (vals.length < 2 || !vals.every((x) => typeof x === 'number' || x instanceof Date)) return null; const n = vals.map((x) => (x instanceof Date ? x.getTime() : x)); const d = n[1] - n[0]; return n.every((x, i) => !i || Math.abs(x - n[i - 1] - d) < 1e-9) ? { last: n.at(-1), first: n[0], d, date: vals[0] instanceof Date } : null; };
    change(list, (cell, r, c) => {
      const len = down ? src.r2 - src.r1 + 1 : src.c2 - src.c1 + 1;
      const k = down ? r - src.r1 : c - src.c1; const idx = ((k % len) + len) % len;
      const sr = down ? src.r1 + idx : r; const sc = down ? c : src.c1 + idx;
      const s = ws.getCell(sr, sc);
      const line = down ? range(src.r1, src.r2).map((rr) => ws.getCell(rr, c).value) : range(src.c1, src.c2).map((cc) => ws.getCell(r, cc).value);
      const ser = series(line);
      cell.style = structuredClone(s.style || {});
      if (ser) { const v2 = k >= len ? ser.last + ser.d * (k - len + 1) : ser.first + ser.d * k; cell.value = ser.date ? new Date(v2) : v2; return; }
      const val = s.value == null ? null : structuredClone(s.value);
      if (val && typeof val === 'object' && val.formula) { val.formula = shiftFormula(val.formula, r - sr, c - sc); delete val.result; }
      cell.value = val;
    }, 'fill');
    vs().sel = { r: to.r1, c: to.c1, r2: to.r2, c2: to.c2 }; refreshBars(); paint();
  }

  // ---------- Home: formatting ----------
  const fmtSel = (mutate, label) => change(selCells(), mutate, label);
  function toggleFont(key) { const on = !ws.getCell(vs().sel.r, vs().sel.c).font?.[key]; fmtSel((cell) => { cell.font = { ...(cell.font || {}), [key]: on }; }, key); }
  $('#xe-b').onclick = () => toggleFont('bold'); $('#xe-i').onclick = () => toggleFont('italic'); $('#xe-u').onclick = () => toggleFont('underline');
  $('#xe-font').onchange = (e) => fmtSel((cell) => { cell.font = { ...(cell.font || {}), name: e.target.value }; }, 'font');
  $('#xe-size').onchange = (e) => fmtSel((cell) => { cell.font = { ...(cell.font || {}), size: +e.target.value }; }, 'size');
  $('#xe-fc').addEventListener('change', (e) => fmtSel((cell) => { cell.font = { ...(cell.font || {}), color: { argb: `FF${e.target.value.slice(1).toUpperCase()}` } }; }, 'font colour'));
  $('#xe-bg').addEventListener('change', (e) => fmtSel((cell) => { cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${e.target.value.slice(1).toUpperCase()}` } }; }, 'fill'));
  const align = (k, v) => fmtSel((cell) => { cell.alignment = { ...(cell.alignment || {}), [k]: cell.alignment?.[k] === v && k === 'horizontal' ? undefined : v }; }, 'alignment');
  $('#xe-al').onclick = () => align('horizontal', 'left'); $('#xe-ac').onclick = () => align('horizontal', 'center'); $('#xe-ar').onclick = () => align('horizontal', 'right');
  $('#xe-vt').onclick = () => align('vertical', 'top'); $('#xe-vm').onclick = () => align('vertical', 'middle'); $('#xe-vb').onclick = () => align('vertical', 'bottom');
  $('#xe-wrap').onclick = () => { const on = !ws.getCell(vs().sel.r, vs().sel.c).alignment?.wrapText; fmtSel((cell) => { cell.alignment = { ...(cell.alignment || {}), wrapText: on }; }, 'wrap'); };
  $('#xe-merge').onclick = () => {
    const v = vs(); const s = norm(v.sel); const m = merged(v.sel.r, v.sel.c);
    if (m) structural(() => ws.unMergeCells(m.r1, m.c1, m.r2, m.c2), 'unmerge');
    else if (s.r1 !== s.r2 || s.c1 !== s.c2) structural(() => { ws.mergeCells(s.r1, s.c1, s.r2, s.c2); const cl = ws.getCell(s.r1, s.c1); cl.alignment = { ...(cl.alignment || {}), horizontal: 'center', vertical: 'middle' }; }, 'merge');
  };
  $('#xe-numfmt').onchange = (e) => fmtSel((cell) => { cell.numFmt = e.target.value; }, 'number format');
  $('#xe-pct').onclick = () => fmtSel((cell) => { cell.numFmt = '0%'; }, 'percent');
  $('#xe-comma').onclick = () => fmtSel((cell) => { cell.numFmt = '#,##0.00'; }, 'comma');
  const decimals = (d) => fmtSel((cell) => {
    let f = cell.numFmt && cell.numFmt !== 'General' ? cell.numFmt : '0';
    const m = /(\d)(\.(0+))?/.exec(f); if (!m) return;
    const n = Math.max(0, (m[3] || '').length + d); f = f.replace(/0(\.0+)?/, n ? `0.${'0'.repeat(n)}` : '0'); cell.numFmt = f;
  }, 'decimals');
  $('#xe-decup').onclick = () => decimals(1); $('#xe-decdn').onclick = () => decimals(-1);
  $('#xe-border').onchange = (e) => {
    const k = e.target.value; e.target.value = ''; if (!k) return;
    const s = norm(vs().sel); const thin = { style: 'thin', color: { argb: 'FF000000' } }; const thick = { style: 'medium', color: { argb: 'FF000000' } };
    fmtSel((cell, r, c) => {
      const b = { ...(cell.border || {}) };
      if (k === 'none') { cell.border = {}; return; }
      if (k === 'all') { cell.border = { top: thin, left: thin, bottom: thin, right: thin }; return; }
      const o = k === 'thick' ? thick : thin;
      if (k === 'outside' || k === 'thick') { if (r === s.r1) b.top = o; if (r === s.r2) b.bottom = o; if (c === s.c1) b.left = o; if (c === s.c2) b.right = o; }
      if (k === 'bottom' && r === s.r2) b.bottom = thin; if (k === 'top' && r === s.r1) b.top = thin;
      if (k === 'double' && r === s.r2) b.bottom = { style: 'double', color: { argb: 'FF000000' } };
      cell.border = b;
    }, 'borders');
  };

  // ---------- Home: format as table ----------
  // Table styles from the site's colour themes (js/themes.js): a header row in the theme's strongest colour, banded
  // rows in a light tint of it, thin borders and an AutoFilter on the header. Written as cell formatting, so Excel,
  // Google Sheets and Numbers show it the same.
  const rgb = (h) => [1, 3, 5].map((i) => parseInt(h.slice(i, i + 2), 16));
  const lum = (h) => { const [r, g, b] = rgb(h).map((x) => { x /= 255; return x <= 0.03928 ? x / 12.92 : ((x + 0.055) / 1.055) ** 2.4; }); return 0.2126 * r + 0.7152 * g + 0.0722 * b; };
  const mix = (h, k) => `#${rgb(h).map((x) => Math.round(x + (255 - x) * k).toString(16).padStart(2, '0')).join('')}`;
  const tableStyle = (t) => {
    const head = t.chart.find((c) => lum(c) < 0.3) || [...t.chart].sort((a, b) => lum(a) - lum(b))[0];
    return { head, headText: lum(head) > 0.35 ? '#000000' : '#ffffff', band: mix(head, 0.86), line: mix(head, 0.55) };
  };
  const argb = (h) => ({ argb: `FF${h.slice(1).toUpperCase()}` });
  function formatTable(t, { header = true, banded = true, filter = true } = {}) {
    if (ws.sheetProtection?.sheet) { status('This sheet is protected: its cells cannot be formatted here.'); return; }
    const g = region(); const ts = tableStyle(t); const thin = { style: 'thin', color: argb(ts.line) };
    change(range(g.r1, g.r2).flatMap((r) => range(g.c1, g.c2).map((c) => [r, c])), (cell, r) => {
      const isHead = header && r === g.r1; const k = r - g.r1 - (header ? 1 : 0);
      cell.border = { top: thin, left: thin, bottom: thin, right: thin };
      if (isHead) { cell.fill = { type: 'pattern', pattern: 'solid', fgColor: argb(ts.head) }; cell.font = { ...(cell.font || {}), bold: true, color: argb(ts.headText) }; return; }
      cell.fill = banded && k % 2 === 1 ? { type: 'pattern', pattern: 'solid', fgColor: argb(ts.band) } : { type: 'pattern', pattern: 'none' };
      if (cell.font?.color) { const f = { ...cell.font }; delete f.color; cell.font = f; }
    }, 'format as table');
    if (header && filter && g.r2 > g.r1) { structural(() => { ws.autoFilter = `${addr(g.r1, g.c1)}:${addr(g.r1, g.c2)}`; }, 'filter'); view.delete(ws); vs(); }
    vs().sel = { r: g.r1, c: g.c1, r2: g.r2, c2: g.c2 }; buildGeo(); refreshBars(); paint();
    status(`${addr(g.r1, g.c1)}:${addr(g.r2, g.c2)} formatted as a table (${t.name})`);
  }
  const tableOpts = { header: true, banded: true, filter: true };
  $('#xe-table').onclick = () => {
    const cur = currentTheme(); const list = [cur, ...THEMES.filter((t) => t.id !== cur.id)];
    const mini = (t) => { const ts = tableStyle(t); return `<span class="xe-tsw" aria-hidden="true">${[0, 1, 2, 3].map((i) => `<i style="background:${i === 0 ? ts.head : i % 2 === 0 ? ts.band : '#fff'};border-color:${ts.line}"></i>`).join('')}</span>`; };
    pop.innerHTML = `<div class="xe-fm"><div class="xe-popg">Format as table</div>
      <div class="xe-tgal">${list.map((t) => `<button type="button" class="xe-tst" data-t="${t.id}" title="${esc(t.name)}">${mini(t)}<span>${esc(t.name)}</span></button>`).join('')}</div>
      <label class="xe-fmi"><input type="checkbox" data-o="header" ${tableOpts.header ? 'checked' : ''}> Header row</label>
      <label class="xe-fmi"><input type="checkbox" data-o="banded" ${tableOpts.banded ? 'checked' : ''}> Banded rows</label>
      <label class="xe-fmi"><input type="checkbox" data-o="filter" ${tableOpts.filter ? 'checked' : ''}> Filter buttons</label></div>`;
    pop.hidden = false; popFor = null; placeUnder(pop, $('#xe-table'));
    pop.querySelectorAll('[data-o]').forEach((i) => { i.onchange = () => { tableOpts[i.dataset.o] = i.checked; }; });
    pop.querySelectorAll('[data-t]').forEach((b) => { b.onclick = () => { hidePop(); formatTable(THEMES.find((t) => t.id === b.dataset.t), tableOpts); grid.focus(); }; });
  };

  // ---------- Insert: rows and columns ----------
  function spliceSheet(axis, at, n) {
    // Merges and column widths are moved by hand (ExcelJS does not), then every formula in the workbook is adjusted.
    const ms = geo.merges.map((m) => ({ ...m })); ms.forEach((m) => ws.unMergeCells(m.r1, m.c1, m.r2, m.c2));
    const widths = axis === 'col' ? range(1, Math.max(ws.columnCount, at) + Math.abs(n) + 1).map((c) => ws.getColumn(c).width) : null;
    if (axis === 'row') { if (n > 0) ws.spliceRows(at, 0, ...Array.from({ length: n }, () => [])); else ws.spliceRows(at, -n); }
    else if (n > 0) ws.spliceColumns(at, 0, ...Array.from({ length: n }, () => [])); else ws.spliceColumns(at, -n);
    if (widths) {
      const nw = n > 0 ? [...widths.slice(0, at - 1), ...Array(n).fill(undefined), ...widths.slice(at - 1)] : [...widths.slice(0, at - 1), ...widths.slice(at - 1 - n)];
      nw.forEach((w, i) => { ws.getColumn(i + 1).width = w; });
    }
    wb.worksheets.forEach((s) => s.eachRow({ includeEmpty: false }, (row) => row.eachCell({ includeEmpty: false }, (cell) => {
      if (!cell.formula) return; const f = adjustFormula(cell.formula, s.name, ws.name, axis, at, n);
      if (f !== cell.formula) cell.value = { formula: f };
    })));
    const k1 = axis === 'row' ? 'r1' : 'c1', k2 = axis === 'row' ? 'r2' : 'c2';
    ms.forEach((m) => {
      const mv = (x) => (n > 0 ? (x >= at ? x + n : x) : x >= at - n ? x + n : x < at ? x : null);
      const a = mv(m[k1]); const b = mv(m[k2]); if (a == null || b == null) return;
      try { ws.mergeCells(axis === 'row' ? a : m.r1, axis === 'row' ? m.c1 : a, axis === 'row' ? b : m.r2, axis === 'row' ? m.c2 : b); } catch { /* overlapping: dropped */ }
    });
    const v = vs(); if (v.filter) v.filter = null;
  }
  const rowsOf = () => { const s = norm(vs().sel); return { a: s.r1, n: Math.min(s.r2, geo.nR) - s.r1 + 1, ca: s.c1, cn: Math.min(s.c2, geo.nC) - s.c1 + 1 }; };
  $('#xe-rowup').onclick = () => { const { a, n } = rowsOf(); structural(() => spliceSheet('row', a, Math.min(n, 1000)), 'insert rows'); };
  $('#xe-rowdn').onclick = () => { const { a, n } = rowsOf(); structural(() => spliceSheet('row', a + n, Math.min(n, 1000)), 'insert rows'); };
  $('#xe-rowdel').onclick = () => { const { a, n } = rowsOf(); if (n > ws.rowCount) return; structural(() => spliceSheet('row', a, -n), 'delete rows'); };
  $('#xe-colleft').onclick = () => { const { ca, cn } = rowsOf(); structural(() => spliceSheet('col', ca, Math.min(cn, 100)), 'insert columns'); };
  $('#xe-colright').onclick = () => { const { ca, cn } = rowsOf(); structural(() => spliceSheet('col', ca + cn, Math.min(cn, 100)), 'insert columns'); };
  $('#xe-coldel').onclick = () => { const { ca, cn } = rowsOf(); if (cn > ws.columnCount) return; structural(() => spliceSheet('col', ca, -cn), 'delete columns'); };

  // ---------- Formulas ----------
  $('#xe-fnlist').onclick = () => functionList();
  $('#xe-autosum').onclick = () => {
    const v = vs(); let r = v.sel.r - 1; const c = v.sel.c;
    while (r >= 1 && typeof valueOf(ws, ws.getCell(r, c)) !== 'number') r--;
    const end = r; while (r >= 1 && typeof valueOf(ws, ws.getCell(r, c)) === 'number') r--;
    startEdit('edit', end >= 1 ? `=SUM(${addr(r + 1, c)}:${addr(end, c)})` : '=SUM()');
  };

  // ---------- Data: sort and filter ----------
  // The data block around the active cell (or the selection when it is more than one cell).
  function region() {
    const v = vs(); const s = norm(v.sel);
    if (s.r1 !== s.r2 || s.c1 !== s.c2) return { ...s, r2: Math.min(s.r2, ws.rowCount), c2: Math.min(s.c2, ws.columnCount) };
    const has = (r, c) => { const x = ws.findCell(r, c); return x && x.value != null && x.value !== ''; };
    let { r1, r2, c1, c2 } = s; let grew = true;
    while (grew) {
      grew = false;
      if (r1 > 1 && range(c1, c2).some((c) => has(r1 - 1, c))) { r1--; grew = true; }
      if (r2 < ws.rowCount && range(c1, c2).some((c) => has(r2 + 1, c))) { r2++; grew = true; }
      if (c1 > 1 && range(r1, r2).some((r) => has(r, c1 - 1))) { c1--; grew = true; }
      if (c2 < ws.columnCount && range(r1, r2).some((r) => has(r, c2 + 1))) { c2++; grew = true; }
    }
    return { r1, r2, c1, c2 };
  }
  const headerLike = (g) => { const v = vs(); if (v.filter && v.filter.ref.r1 === g.r1) return true; const top = range(g.c1, g.c2).map((c) => valueOf(ws, ws.getCell(g.r1, c))); const next = g.r2 > g.r1 ? range(g.c1, g.c2).map((c) => valueOf(ws, ws.getCell(g.r1 + 1, c))) : []; return top.every((x) => typeof x === 'string' || x == null) && next.some((x) => typeof x === 'number' || x instanceof Date); };
  const cmpVal = (a, b) => {
    const na = a instanceof Date ? a.getTime() : a; const nb = b instanceof Date ? b.getTime() : b;
    if (na == null || na === '') return nb == null || nb === '' ? 0 : 1; if (nb == null || nb === '') return -1;
    if (typeof na === 'number' && typeof nb === 'number') return na - nb;
    if (typeof na === 'number') return -1; if (typeof nb === 'number') return 1;
    return String(na).localeCompare(String(nb), 'th', { numeric: true });
  };
  function sortBy(col, desc, g = region()) {
    const head = headerLike(g) ? 1 : 0; const r1 = g.r1 + head; if (g.r2 <= r1) return;
    const rows = range(r1, g.r2).map((r) => ({ r, key: valueOf(ws, ws.getCell(r, col)), cells: range(g.c1, g.c2).map((c) => snap(ws.getCell(r, c))) }));
    rows.sort((a, b) => (desc ? -cmpVal(a.key, b.key) : cmpVal(a.key, b.key)) || a.r - b.r);
    const list = []; for (let r = r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) list.push([r, c]);
    change(list, (cell, r, c) => {
      const src = rows[r - r1]; const s = src.cells[c - g.c1]; const val = s.value == null ? null : structuredClone(s.value);
      if (val && typeof val === 'object' && val.formula) { val.formula = shiftFormula(val.formula, r - src.r, 0); delete val.result; }
      cell.value = val; cell.style = structuredClone(s.style);
    }, `sort ${desc ? 'Z→A' : 'A→Z'}`);
    status(`Sorted ${addr(r1, g.c1)}:${addr(g.r2, g.c2)} by column ${colName(col)}${head ? ' (header row kept)' : ''}`);
  }
  $('#xe-sortaz').onclick = () => sortBy(vs().sel.c, false); $('#xe-sortza').onclick = () => sortBy(vs().sel.c, true);
  $('#xe-filter').onclick = () => {
    const v = vs();
    if (v.filter) structural(() => { ws.autoFilter = null; v.filter = null; }, 'filter off');
    else { const g = region(); structural(() => { ws.autoFilter = `${addr(g.r1, g.c1)}:${addr(g.r1, g.c2)}`; }, 'filter'); view.delete(ws); vs(); buildGeo(); refreshBars(); paint(); }
  };
  $('#xe-unfilter').onclick = () => { const v = vs(); if (!v.filter) return; v.filter.crit.clear(); applyFilter(); };
  function applyFilter() {
    const v = vs(); const f = v.filter; f.hidden = new Set();
    const last = Math.max(ws.rowCount, f.ref.r1);
    for (let r = f.ref.r1 + 1; r <= last; r++) for (const [c, allowed] of f.crit) { if (!allowed.has(shownText(ws.getCell(r, c), valueOf(ws, ws.getCell(r, c))))) { f.hidden.add(r); break; } }
    buildGeo(); refreshBars(); paint();
    status(f.crit.size ? `${f.hidden.size} rows hidden by the filter` : 'Filter cleared');
  }
  function filterMenu(c) {
    const v = vs(); const f = v.filter; if (!f) return;
    const vals = new Map(); const last = Math.max(ws.rowCount, f.ref.r1);
    for (let r = f.ref.r1 + 1; r <= last; r++) { const t = shownText(ws.getCell(r, c), valueOf(ws, ws.getCell(r, c))); vals.set(t, (vals.get(t) || 0) + 1); if (vals.size > 2000) break; }
    const list = [...vals.keys()].sort((a, b) => cmpVal(a, b)); const allowed = f.crit.get(c);
    pop.innerHTML = `<div class="xe-fm"><div class="xe-fmrow"><button type="button" class="xe-fmb" data-s="az">Sort A→Z</button><button type="button" class="xe-fmb" data-s="za">Sort Z→A</button></div>
      <input class="xe-popq" placeholder="Search…" aria-label="Search values" autocomplete="off">
      <label class="xe-fmi"><input type="checkbox" data-all="1" ${!allowed ? 'checked' : ''}> (Select all)</label>
      <div class="xe-fml">${list.map((t) => `<label class="xe-fmi"><input type="checkbox" data-v="${esc(t)}" ${!allowed || allowed.has(t) ? 'checked' : ''}> ${t === '' ? '(Blanks)' : esc(t)}</label>`).join('')}</div>
      <div class="xe-fmrow"><button type="button" class="xe-fmb xe-fmok">OK</button><button type="button" class="xe-fmb" data-cancel="1">Cancel</button></div></div>`;
    pop.hidden = false; popFor = null;
    const cellEl = canvas.querySelector(`.xc[data-r="${f.ref.r1}"][data-c="${c}"]`) || grid; placeUnder(pop, cellEl);
    const q = pop.querySelector('.xe-popq');
    q.oninput = () => { const s = q.value.toLowerCase(); pop.querySelectorAll('.xe-fml .xe-fmi').forEach((l) => { l.hidden = s && !l.textContent.toLowerCase().includes(s); }); };
    pop.querySelector('[data-all]').onchange = (e) => pop.querySelectorAll('.xe-fml input').forEach((i) => { if (!i.closest('.xe-fmi').hidden) i.checked = e.target.checked; });
    pop.querySelector('.xe-fmok').onclick = () => {
      const on = [...pop.querySelectorAll('.xe-fml input')].filter((i) => i.checked).map((i) => i.dataset.v);
      if (on.length === list.length) f.crit.delete(c); else f.crit.set(c, new Set(on));
      hidePop(); applyFilter();
    };
    pop.querySelector('[data-cancel]').onclick = () => hidePop();
    pop.querySelectorAll('[data-s]').forEach((b) => { b.onclick = () => { hidePop(); sortBy(c, b.dataset.s === 'za', { r1: f.ref.r1, c1: f.ref.c1, c2: f.ref.c2, r2: ws.rowCount }); }; });
  }

  // ---------- View: freeze panes ----------
  $('#xe-freeze').onchange = (e) => {
    const k = e.target.value; e.target.value = ''; if (!k) return;
    const v = vs(); const s = norm(v.sel);
    const [fr, fc] = k === 'row' ? [1, 0] : k === 'col' ? [0, 1] : k === 'sel' ? [s.r1 - 1, s.c1 - 1] : [0, 0];
    structural(() => { ws.views = fr || fc ? [{ state: 'frozen', xSplit: fc, ySplit: fr, topLeftCell: addr(fr + 1, fc + 1) }] : [{}]; }, 'freeze panes');
    v.fr = fr; v.fc = fc; view.set(ws, v); paint();
  };

  // ---------- name box, sheets, tools ----------
  nameBox.addEventListener('keydown', (e) => {
    if (e.key !== 'Enter') return; e.preventDefault();
    const t = nameBox.value.trim(); const bang = t.lastIndexOf('!');
    if (bang > 0) { const name = t.slice(0, bang).replace(/^'|'$/g, ''); const i = wb.worksheets.findIndex((s) => s.name === name); if (i >= 0) showSheet(i); }
    const g = parseRange(bang > 0 ? t.slice(bang + 1) : t); if (!g) { status(`"${t}" is not a cell or range address (for example B12 or A1:C5).`); return; }
    vs().sel = { r: g.r1, c: g.c1, r2: g.r2, c2: g.c2 }; scrollTo(g.r1, g.c1); refreshBars(); paint(); grid.focus();
  });
  function showSheet(i) {
    if (editing) commitEdit();
    ws = wb.worksheets[i]; grid.scrollTop = 0; grid.scrollLeft = 0; buildGeo(); refreshBars(); paint();
  }
  $('.xe-tabs').addEventListener('click', (e) => { const t = e.target.closest('.xe-tab'); if (t) { showSheet(+t.dataset.i); grid.focus(); } });
  $('#xe-undo').onclick = () => step('undo'); $('#xe-redo').onclick = () => step('redo');
  $('#xe-del').title = 'Clear the selected cells'; $('#xe-del').setAttribute('aria-label', 'Clear the selected cells'); $('#xe-del').onclick = clearSel;

  // Other file types: the active sheet as shown (formatted values, like Excel's Save As). CSV is UTF-8 with a BOM so
  // Excel opens Thai text correctly; "Unicode text" is Excel's tab-separated UTF-16 format.
  function sheetText() {
    const rows = []; const nR = ws.rowCount, nC = ws.columnCount;
    for (let r = 1; r <= nR; r++) rows.push(range(1, nC).map((c) => { if (geo.covered.has(`${r}:${c}`)) return ''; /* merged: the value once, in its first cell */ const cell = ws.findCell(r, c); return cell ? shownText(cell, valueOf(ws, cell)) : ''; }));
    while (rows.length && rows.at(-1).every((x) => x === '')) rows.pop();
    return rows;
  }
  const csvCell = (x, sep) => (/["\r\n]/.test(x) || x.includes(sep) || /^\s|\s$/.test(x) ? `"${x.replace(/"/g, '""')}"` : x);
  const joined = (sep) => `${sheetText().map((r) => r.map((x) => csvCell(x, sep)).join(sep)).join('\r\n')}\r\n`;
  const EXPORTS = {
    csv: () => ({ data: `\ufeff${joined(',')}`, ext: 'csv', mime: 'text/csv;charset=utf-8' }),
    tsv: () => ({ data: joined('\t'), ext: 'tsv', mime: 'text/tab-separated-values;charset=utf-8' }),
    utf16: () => { const t = `\ufeff${joined('\t')}`; const b = new Uint8Array(t.length * 2); for (let i = 0; i < t.length; i++) { const c = t.charCodeAt(i); b[2 * i] = c & 255; b[2 * i + 1] = c >> 8; } return { data: b, ext: 'txt', mime: 'text/plain;charset=utf-16le' }; },
    json: () => { const [head = [], ...rest] = sheetText(); const keys = head.map((h, i) => h || colName(i + 1)); return { data: JSON.stringify(rest.map((r) => Object.fromEntries(keys.map((k, i) => [k, r[i] ?? '']))), null, 2), ext: 'json', mime: 'application/json' }; },
    html: () => ({ data: `<!doctype html><meta charset="utf-8"><title>${esc(ws.name)}</title><table border="1" style="border-collapse:collapse;font-family:sans-serif">${sheetText().map((r) => `<tr>${r.map((x) => `<td>${esc(x)}</td>`).join('')}</tr>`).join('')}</table>`, ext: 'html', mime: 'text/html;charset=utf-8' }),
  };

  buildGeo(); refreshBars(); render();
  if (ws.sheetProtection?.sheet) status('Protected sheet: grey cells are locked');

  return {
    isDirty: () => dirty,
    async save() {
      if (editing) commitEdit();
      const edited = await wb.xlsx.writeBuffer();
      const out = await recalcOnOpen(await relock(bytes, edited));
      dirty = false;
      return out;
    },
    // fmt: csv, tsv, utf16, json, html. Resolves { data, ext, mime, sheet }.
    exportAs(fmt) { if (editing) commitEdit(); return { ...EXPORTS[fmt](), sheet: ws.name }; },
  };
}
