// Excel editor (prototype): edit cell values of an .xlsx in the browser with ExcelJS. Styles, formulas, number
// formats, merged cells and column widths are kept; charts, images and pivot tables are not (ExcelJS drops them).
// On a protected sheet, locked cells stay read-only, and relock() (xlsx-relock.js, from TUSUFinanceFeedback) copies
// the original file's protection back onto the saved copy, so its password never has to be known.
// Formulas are calculated live with HyperFormula (GPL-3.0): editing a cell updates every result that depends on it.
import { relock } from './xlsx-relock.js';

const EXCELJS = 'https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js';
const HYPERFORMULA = 'https://cdn.jsdelivr.net/npm/hyperformula@3.0.1/dist/hyperformula.full.min.js'; // not on cdnjs
const MAX_ROWS = 1000; const MAX_COLS = 60; // grid size shown; larger sheets show their top-left part
const MAX_CALC_CELLS = 500000; // larger workbooks show the results saved in the file instead of live ones

const script = (src, global) => window[global] ? Promise.resolve(window[global]) : new Promise((res, rej) => {
  const s = document.createElement('script'); s.src = src; s.crossOrigin = 'anonymous';
  s.onload = () => res(window[global]); s.onerror = () => rej(new Error('Could not load the Excel tools. Try again.'));
  document.head.append(s);
});

const colName = (n) => { let s = ''; for (; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s; };
const colNum = (s) => [...s].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
const parseAddr = (a) => { const m = /^([A-Z]+)(\d+)$/.exec(a); return [+m[2], colNum(m[1])]; };
const pad = (n) => String(n).padStart(2, '0');

// The text a cell shows while it is being edited: "=FORMULA" for formulas, yyyy-mm-dd for dates.
function rawText(cell) {
  if (cell.formula) return `=${cell.formula}`;
  const v = cell.value;
  if (v == null) return '';
  if (v instanceof Date) return `${v.getUTCFullYear()}-${pad(v.getUTCMonth() + 1)}-${pad(v.getUTCDate())}`;
  if (typeof v === 'object') return v.richText ? v.richText.map((r) => r.text).join('') : v.text ?? v.error ?? '';
  return String(v);
}
// The text a cell shows in the grid: a formula's result (`res`: the live one, else the one saved in the file; none:
// blank, as in Excel), else its value, in the cell's number format.
function shownText(cell, res = cell.result) {
  const v = cell.formula ? res : cell.value;
  if (typeof v === 'number') return cell.formula && isDateFmt(cell.numFmt) ? rawText({ value: fromSerial(v) }) : formatNumber(v, cell.numFmt);
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (cell.formula) return v == null ? '' : v instanceof Date ? rawText({ value: v }) : typeof v === 'object' ? v.error ?? '' : String(v);
  return rawText(cell);
}
// Dates are day numbers in formulas (Excel's serial dates, day 0 = 30 Dec 1899).
const EPOCH = Date.UTC(1899, 11, 30);
const toSerial = (d) => (d.getTime() - EPOCH) / 86400000;
const fromSerial = (n) => new Date(EPOCH + Math.round(n * 86400000));
const isDateFmt = (f) => { const x = (f || '').replace(/"[^"]*"|\[[^\]]*\]|\\./g, ''); return /[dy]|m{3,}/i.test(x) && !/[#0]/.test(x); };
// A cell as the calculation engine reads it: "=..." for formulas (without Excel's _xlfn. prefixes), day numbers for
// dates; the hidden cells of a merged range are empty, as in Excel.
function calcContent(cell) {
  if (!cell || (cell.isMerged && cell.master !== cell)) return null;
  if (cell.formula) return `=${cell.formula.replace(/_xlfn\.|_xlws\./g, '')}`;
  const v = cell.value;
  if (v == null) return null;
  if (v instanceof Date) return toSerial(v);
  if (typeof v === 'object') return v.error ?? (v.richText ? v.richText.map((r) => r.text).join('') : v.text ?? null);
  if (typeof v === 'string' && /^[=+-]/.test(v)) return `'${v}`; // text that only looks like a formula
  return v;
}
// Number formats, the common parts: thousands separator, decimals, percent, negative and zero sections, and text
// around the number ("฿", "บาท", brackets). Other formats show the plain number.
function formatNumber(v, fmt) {
  if (!fmt || fmt === 'General') return String(Math.round(v * 1e10) / 1e10);
  const sec = fmt.split(';');
  const f = v < 0 && sec[1] ? sec[1] : v === 0 && sec[2] ? sec[2] : sec[0];
  // Literal parts: [$฿-41E] is a currency symbol, other [...] codes are dropped; "*x" fills (shown as a space) and
  // "_x" spacers; "?" placeholders only pad.
  const lit = (x) => x.replace(/\[\$([^\]-]*)[^\]]*\]/g, '$1').replace(/\[[^\]]*\]/g, '').replace(/"([^"]*)"/g, '$1').replace(/\*./g, ' ').replace(/_./g, '').replace(/\\(.)/g, '$1').replace(/\?/g, '');
  const core = /[#0?][#0?,]*(\.[0#?]+)?%?/.exec(f.replace(/"[^"]*"|\[[^\]]*\]/g, (q) => ' '.repeat(q.length)));
  if (!core || !/[#0]/.test(core[0])) return lit(f).trim() || String(v); // e.g. the zero section "-" of accounting formats
  const dec = (core[1] || '.').length - 1; const pct = core[0].endsWith('%');
  let s = Math.abs(pct ? v * 100 : v).toLocaleString('en-US', { minimumFractionDigits: dec, maximumFractionDigits: dec, useGrouping: core[0].includes(',') });
  if (pct) s += '%';
  const minus = v < 0 && !sec[1] && Number(s.replace(/[^0-9.]/g, '')) !== 0; // as in Excel, no "-0.00"
  return `${minus ? '-' : ''}${lit(f.slice(0, core.index))}${s}${lit(f.slice(core.index + core[0].length))}`.trim();
}

// Colours: argb, or a colour of the workbook theme with its tint.
function themeColors(wb) {
  const xml = wb._themes?.theme1 || '';
  return ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink']
    .map((n) => new RegExp(`<a:${n}>[\\s\\S]*?(?:srgbClr val|lastClr)="([0-9A-Fa-f]{6})"`).exec(xml)?.[1] || null);
}
function cssColor(c, theme) {
  const hex = c?.argb ? c.argb.slice(-6) : c?.theme != null ? theme[c.theme] : null;
  if (!hex) return null;
  const t = c.tint || 0;
  const rgb = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16)).map((x) => Math.round(t < 0 ? x * (1 + t) : x + (255 - x) * t));
  return `rgb(${rgb.join(',')})`;
}
const BORDER = { thin: '1px solid', hair: '1px solid', medium: '2px solid', thick: '3px solid', double: '3px double', dotted: '1px dotted', dashed: '1px dashed', dashDot: '1px dashed', dashDotDot: '1px dotted', mediumDashed: '2px dashed', mediumDashDot: '2px dashed', mediumDashDotDot: '2px dotted', slantDashDot: '2px dashed' };

// The cell's own look in the grid: font, fill, borders, alignment and wrapping, as in Excel.
function styleCell(td, cell, theme) {
  const f = cell.font || {};
  if (f.name) td.style.fontFamily = `"${f.name}", sans-serif`;
  if (f.size) td.style.fontSize = `${f.size}pt`;
  if (f.bold) td.style.fontWeight = 'bold';
  if (f.italic) td.style.fontStyle = 'italic';
  const deco = [f.underline && 'underline', f.strike && 'line-through'].filter(Boolean); if (deco.length) td.style.textDecoration = deco.join(' ');
  const fc = cssColor(f.color, theme); if (fc) td.style.color = fc;
  if (cell.fill?.type === 'pattern' && cell.fill.pattern !== 'none') { const bg = cssColor(cell.fill.fgColor, theme); if (bg) td.style.background = bg; }
  const b = cell.border || {};
  for (const side of ['Top', 'Right', 'Bottom', 'Left']) {
    const e = b[side.toLowerCase()]; if (e?.style) td.style[`border${side}`] = `${BORDER[e.style] || '1px solid'} ${cssColor(e.color, theme) || '#000'}`;
  }
  const a = cell.alignment || {};
  const num = typeof (cell.formula ? cell.result : cell.value) === 'number';
  td.style.textAlign = a.horizontal && a.horizontal !== 'general' ? ({ centerContinuous: 'center', distributed: 'justify', fill: 'left' }[a.horizontal] || a.horizontal) : num ? 'right' : 'left';
  td.style.verticalAlign = a.vertical === 'middle' || a.vertical === 'top' ? a.vertical : 'bottom';
  if (a.wrapText) td.classList.add('xe-wrap');
  if (a.indent) td.style.paddingLeft = `${6 + a.indent * 9}px`;
}
// Typed text -> cell value: "=..." formula, number, date (only where the cell held a date), else text.
function parseValue(text, cell) {
  if (text === '') return null;
  if (text.startsWith('=') && text.length > 1) return { formula: text.slice(1) };
  const num = text.replace(/,/g, '');
  if (/^-?\d+(\.\d+)?$/.test(num)) return Number(num);
  if (cell.value instanceof Date && /^\d{4}-\d{2}-\d{2}$/.test(text)) return new Date(`${text}T00:00:00Z`);
  return text;
}

// Marks the saved file so Excel recalculates every formula when it opens it (edited cells change results, and a
// formula typed here has none yet). Done last, on the XML, because relock() rewrites the workbook.
async function recalcOnOpen(buf) {
  const zip = await window.JSZip.loadAsync(buf); // loaded by relock()
  let xml = await zip.file('xl/workbook.xml').async('string');
  xml = /<calcPr\b/.test(xml) ? xml.replace(/<calcPr\b(?![^>]*fullCalcOnLoad)/, '<calcPr fullCalcOnLoad="1"') : xml.replace(/(<\/definedNames>|<\/sheets>)/, '$1<calcPr fullCalcOnLoad="1"/>');
  zip.file('xl/workbook.xml', xml);
  return zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
}

// Opens the editor in `root`. bytes: the .xlsx (ArrayBuffer); t(en, th): translates labels. Resolves
// { isDirty(), save() }; save() resolves the edited .xlsx as an ArrayBuffer.
export async function openXlsxEditor(root, bytes, { t = (en) => en } = {}) {
  const ExcelJS = await script(EXCELJS, 'ExcelJS');
  const wb = new ExcelJS.Workbook();
  try { await wb.xlsx.load(bytes.slice(0)); } catch { throw new Error(t('Could not read this file. Only .xlsx files can be edited.', 'อ่านไฟล์นี้ไม่ได้ แก้ไขได้เฉพาะไฟล์ .xlsx')); }
  if (!wb.worksheets.length) throw new Error(t('This workbook has no sheets.', 'สมุดงานนี้ไม่มีชีต'));
  const theme = themeColors(wb);
  let dirty = false;

  // Live calculation: every sheet goes into HyperFormula, which keeps track of which cells depend on which. Where it
  // cannot work out a formula that Excel could (a function it does not have, a link to another file), the result
  // saved in the file is shown instead.
  let hf = null; const sheetId = new Map();
  try {
    const cells = wb.worksheets.reduce((n, ws) => n + ws.rowCount * ws.columnCount, 0);
    if (cells <= MAX_CALC_CELLS) {
      const HF = await script(HYPERFORMULA, 'HyperFormula');
      const data = {};
      wb.worksheets.forEach((ws) => {
        data[ws.name] = Array.from({ length: ws.rowCount }, (_, r) => Array.from({ length: ws.columnCount }, (__, c) => calcContent(ws.findCell(r + 1, c + 1))));
      });
      hf = HF.buildFromSheets(data, { licenseKey: 'gpl-v3' });
      wb.worksheets.forEach((ws) => sheetId.set(ws, hf.getSheetId(ws.name)));
    }
  } catch (e) { hf = null; console.warn('Live calculation is off:', e); }
  function result(ws, cell) {
    if (!hf || !cell.formula) return cell.result;
    const v = hf.getCellValue({ sheet: sheetId.get(ws), row: cell.row - 1, col: cell.col - 1 });
    if (v == null || typeof v !== 'object') return v;
    const saved = cell.result;
    return saved != null && !(typeof saved === 'object' && saved.error) ? saved : { error: v.value || '#ERROR' };
  }

  root.innerHTML = `<div class="oe"><div class="oe-bar">
      <label>${t('Sheet', 'ชีต')} <select id="xe-sheet">${wb.worksheets.map((ws, i) => `<option value="${i}">${ws.name.replace(/</g, '&lt;')}</option>`).join('')}</select></label>
      <span class="oe-note" id="xe-note"></span>
      <div class="xe-fx" id="xe-fx"><b></b><span></span></div>
    </div><div class="oe-scroll" id="xe-grid"></div></div>`;
  const grid = root.querySelector('#xe-grid');
  const note = root.querySelector('#xe-note');
  const fx = root.querySelector('#xe-fx');
  // Formula bar: the picked cell's address, and its formula with the result, or its value.
  function showFx(ws, td) {
    const cell = ws.getCell(+td.dataset.r, +td.dataset.c);
    fx.firstChild.textContent = `${colName(+td.dataset.c)}${td.dataset.r}`;
    fx.lastChild.textContent = cell.formula ? `=${cell.formula}   →   ${shownText(cell, result(ws, cell)) || t('(blank)', '(ว่าง)')}` : rawText(cell);
  }

  function render(ws) {
    const rows = Math.min(Math.max(ws.rowCount, 1), MAX_ROWS);
    const cols = Math.min(Math.max(ws.columnCount, 1), MAX_COLS);
    const isProtected = !!ws.sheetProtection?.sheet;
    // Merged ranges: the top-left cell spans the range; the others are not drawn.
    const span = new Map(); const covered = new Set();
    (ws.model.merges || []).forEach((m) => {
      const [a, b] = m.split(':').map(parseAddr);
      span.set(`${a[0]}:${a[1]}`, [b[0] - a[0] + 1, b[1] - a[1] + 1]);
      for (let r = a[0]; r <= b[0]; r++) for (let c = a[1]; c <= b[1]; c++) if (r !== a[0] || c !== a[1]) covered.add(`${r}:${c}`);
    });
    const notes = [];
    if (ws.rowCount > MAX_ROWS || ws.columnCount > MAX_COLS) notes.push(t('Showing the first {rows} rows and {cols} columns', 'แสดง {rows} แถวและ {cols} คอลัมน์แรก').replace('{rows}', rows).replace('{cols}', cols));
    if (isProtected) notes.push(t('Protected sheet: grey cells are locked', 'ชีตที่ป้องกัน: เซลล์สีเทาแก้ไขไม่ได้'));
    notes.push(hf ? t('Formulas calculate live', 'สูตรคำนวณทันทีที่แก้ไข') : t('Live calculation is off for this file: formulas show the results saved in it', 'ไฟล์นี้ไม่คำนวณสูตรทันที: สูตรแสดงผลที่บันทึกไว้ในไฟล์'));
    note.textContent = notes.join(' · ');

    const table = document.createElement('table'); table.className = 'xe-table';
    // Column widths as in Excel (width in characters of the default font, about 7 px each).
    const head = table.insertRow(); head.insertCell().className = 'xe-corner';
    let total = 40;
    for (let c = 1; c <= cols; c++) {
      const th = document.createElement('th'); th.textContent = colName(c);
      const col = ws.getColumn(c); const px = col.hidden ? 0 : Math.round((col.width || 8.43) * 7 + 5);
      th.style.width = `${px}px`; total += px; if (col.hidden) th.className = 'xe-hid';
      head.append(th);
    }
    table.style.width = `${total}px`;
    for (let r = 1; r <= rows; r++) {
      const tr = table.insertRow();
      const row = ws.getRow(r); if (row.hidden) tr.style.display = 'none'; else if (row.height) tr.style.height = `${row.height}pt`;
      const rh = document.createElement('th'); rh.textContent = r; tr.append(rh);
      for (let c = 1; c <= cols; c++) {
        if (covered.has(`${r}:${c}`)) continue;
        const cell = ws.getCell(r, c); const td = tr.insertCell();
        const s = span.get(`${r}:${c}`); if (s) { td.rowSpan = s[0]; td.colSpan = s[1]; }
        td.dataset.r = r; td.dataset.c = c; td.textContent = shownText(cell, result(ws, cell));
        styleCell(td, cell, theme);
        if (ws.getColumn(c).hidden) td.classList.add('xe-hid');
        if (isProtected && cell.protection?.locked !== false) { td.classList.add('xe-locked'); continue; }
        td.contentEditable = 'true'; td.tabIndex = 0;
      }
    }
    grid.replaceChildren(table);

    // Editing: a cell shows its formula while focused; Enter moves down, Escape cancels. Any cell, locked ones too,
    // shows its formula in the formula bar when clicked.
    table.addEventListener('click', (e) => { const td = e.target.closest('td[data-r]'); if (td) showFx(ws, td); });
    table.addEventListener('focusin', (e) => { const td = e.target.closest('td[contenteditable]'); if (td) { td.dataset.before = rawText(ws.getCell(+td.dataset.r, +td.dataset.c)); td.textContent = td.dataset.before; showFx(ws, td); } });
    table.addEventListener('focusout', (e) => {
      const td = e.target.closest('td[contenteditable]'); if (!td) return;
      const cell = ws.getCell(+td.dataset.r, +td.dataset.c);
      const text = td.innerText.replace(/\n$/, '');
      let changes = [];
      if (text !== td.dataset.before) {
        cell.value = parseValue(text, cell);
        td.classList.add('xe-changed'); dirty = true;
        if (hf) try { changes = hf.setCellContents({ sheet: sheetId.get(ws), row: cell.row - 1, col: cell.col - 1 }, [[calcContent(cell)]]); } catch (err) { console.warn(err); }
      }
      // Without live calculation a formula typed here has no result until Excel calculates it: show the formula.
      td.textContent = !hf && td.classList.contains('xe-changed') && cell.formula ? `=${cell.formula}` : shownText(cell, result(ws, cell));
      styleCell(td, cell, theme); // a number typed into a text cell now aligns right, as in Excel
      // Results that changed because of this edit: redraw them and flash them briefly.
      const sid = sheetId.get(ws);
      changes.forEach((ch) => {
        const a = ch.address; if (!a || a.sheet !== sid) return;
        const other = table.querySelector(`td[data-r="${a.row + 1}"][data-c="${a.col + 1}"]`); if (!other || other === td) return;
        other.textContent = shownText(ws.getCell(a.row + 1, a.col + 1), result(ws, ws.getCell(a.row + 1, a.col + 1)));
        other.classList.remove('xe-calc'); void other.offsetWidth; other.classList.add('xe-calc');
      });
      showFx(ws, td);
    });
    table.addEventListener('keydown', (e) => {
      const td = e.target.closest('td[contenteditable]'); if (!td) return;
      if (e.key === 'Escape') { td.textContent = td.dataset.before; td.blur(); }
      if (e.key === 'Enter' && !e.shiftKey) {
        e.preventDefault();
        const next = table.querySelector(`td[data-r="${+td.dataset.r + 1}"][data-c="${td.dataset.c}"][contenteditable]`);
        next ? next.focus() : td.blur();
      }
    });
  }

  const sel = root.querySelector('#xe-sheet');
  sel.onchange = () => render(wb.worksheets[+sel.value]);
  render(wb.worksheets[0]);

  return {
    isDirty: () => dirty,
    async save() {
      document.activeElement?.blur?.();
      const edited = await wb.xlsx.writeBuffer();
      const out = await recalcOnOpen(await relock(bytes, edited));
      dirty = false;
      return out;
    },
  };
}
