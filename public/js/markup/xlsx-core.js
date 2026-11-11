// Spreadsheet helpers for the Excel editor (js/xlsx-edit.js): addresses, cell text, number and date formats, cell
// styles as CSS, formula reference shifting, and HyperFormula set-up (with Excel's TEXT() formats).

export const colName = (n) => { let s = ''; for (; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s; };
export const colNum = (s) => [...s.toUpperCase()].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
export const addr = (r, c) => `${colName(c)}${r}`;
// "B12" or "A1:C5" -> { r1, c1, r2, c2 } (normalised), or null.
export function parseRange(s) {
  const m = /^\s*\$?([A-Za-z]{1,3})\$?(\d{1,7})(?::\$?([A-Za-z]{1,3})\$?(\d{1,7}))?\s*$/.exec(s || '');
  if (!m) return null;
  const a = [+m[2], colNum(m[1])]; const b = m[3] ? [+m[4], colNum(m[3])] : a;
  return { r1: Math.min(a[0], b[0]), c1: Math.min(a[1], b[1]), r2: Math.max(a[0], b[0]), c2: Math.max(a[1], b[1]) };
}
const pad = (n) => String(n).padStart(2, '0');

// ---------- values and text ----------
export const EPOCH = Date.UTC(1899, 11, 30);
export const toSerial = (d) => (d.getTime() - EPOCH) / 86400000;
export const fromSerial = (n) => new Date(EPOCH + Math.round(n * 86400000));
export const isDateFmt = (f) => { const x = (f || '').replace(/"[^"]*"|\[[^\]]*\]|\\./g, ''); return /[dyb]|m{3,}/i.test(x) && !/[#0]/.test(x); };
// The text a cell shows while it is edited: "=FORMULA", yyyy-mm-dd for dates.
export function rawText(cell) {
  if (cell.formula) return `=${cell.formula}`;
  const v = cell.value;
  if (v == null) return '';
  if (v instanceof Date) return `${v.getUTCFullYear()}-${pad(v.getUTCMonth() + 1)}-${pad(v.getUTCDate())}`;
  if (typeof v === 'object') return v.richText ? v.richText.map((r) => r.text).join('') : v.text ?? v.error ?? '';
  return String(v);
}
// The text a cell shows in the grid: a formula's result (`res`), else its value, in the cell's number format.
export function shownText(cell, res = cell.result) {
  const v = cell.formula ? res : cell.value;
  if (typeof v === 'number') return isDateFmt(cell.numFmt) ? formatDate(fromSerial(v), cell.numFmt) : formatNumber(v, cell.numFmt);
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (v instanceof Date) return isDateFmt(cell.numFmt) ? formatDate(v, cell.numFmt) : rawText({ value: v });
  if (cell.formula) return v == null ? '' : typeof v === 'object' ? v.error ?? '' : String(v);
  return rawText(cell);
}
// A cell as the calculation engine reads it: "=..." (without _xlfn. prefixes), day numbers for dates; the hidden cells
// of a merged range are empty.
export function calcContent(cell) {
  if (!cell || (cell.isMerged && cell.master !== cell)) return null;
  if (cell.formula) return `=${cell.formula.replace(/_xlfn\.|_xlws\./g, '')}`;
  const v = cell.value;
  if (v == null) return null;
  if (v instanceof Date) return toSerial(v);
  if (typeof v === 'object') return v.error ?? (v.richText ? v.richText.map((r) => r.text).join('') : v.text ?? null);
  if (typeof v === 'string' && /^[=+-]/.test(v)) return `'${v}`;
  return v;
}
// Typed text -> value: "=..." formula, number (1,234 · 12% · ฿1,234), date (yyyy-mm-dd, d/m/yyyy; B.E. years over
// 2400), TRUE/FALSE, else text.
export function parseValue(text) {
  if (text === '') return null;
  if (text.startsWith('=') && text.length > 1) return { formula: text.slice(1) };
  const t = text.trim();
  const num = t.replace(/^฿|,/g, '').replace(/^\((.*)\)$/, '-$1');
  if (/^-?\d+(\.\d+)?%$/.test(num)) return Number(num.slice(0, -1)) / 100;
  if (/^-?(\d+\.?\d*|\.\d+)(e[+-]?\d+)?$/i.test(num)) return Number(num);
  let m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(t);
  if (m) return new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
  m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(t);
  if (m) { let y = +m[3]; if (y > 2400) y -= 543; return new Date(Date.UTC(y, +m[2] - 1, +m[1])); }
  if (/^(true|false)$/i.test(t)) return /^true$/i.test(t);
  return text;
}

// ---------- number and date formats ----------
// The common parts: thousands separator, decimals, percent, negative and zero sections, and text around the number.
export function formatNumber(v, fmt) {
  if (!fmt || fmt === 'General') return String(Math.round(v * 1e10) / 1e10);
  const sec = fmt.split(';');
  const f = v < 0 && sec[1] ? sec[1] : v === 0 && sec[2] ? sec[2] : sec[0];
  const lit = (x) => x.replace(/\[\$([^\]-]*)[^\]]*\]/g, '$1').replace(/\[[^\]]*\]/g, '').replace(/"([^"]*)"/g, '$1').replace(/\*./g, ' ').replace(/_./g, '').replace(/\\(.)/g, '$1').replace(/\?/g, '');
  const core = /[#0?][#0?,]*(\.[0#?]+)?%?/.exec(f.replace(/"[^"]*"|\[[^\]]*\]/g, (q) => ' '.repeat(q.length)));
  if (!core || !/[#0]/.test(core[0])) return lit(f).trim() || String(v);
  const dec = (core[1] || '.').length - 1; const pct = core[0].endsWith('%');
  let s = Math.abs(pct ? v * 100 : v).toLocaleString('en-US', { minimumFractionDigits: dec, maximumFractionDigits: dec, useGrouping: core[0].includes(',') });
  if (pct) s += '%';
  const minus = v < 0 && !sec[1] && Number(s.replace(/[^0-9.]/g, '')) !== 0;
  return `${minus ? '-' : ''}${lit(f.slice(0, core.index))}${s}${lit(f.slice(core.index + core[0].length))}`.trim();
}
const MONTH_EN = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const MONTH_TH = ['มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน', 'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];
const MON_TH = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
// Dates: d dd ddd dddd m mm mmm mmmm yy yyyy, and Thai Buddhist years bb bbbb; Thai month names with a [$-th…] code.
export function formatDate(d, fmt) {
  const th = /\[\$-(?:th|41E|1041E|[0-9A-F]*041E)/i.test(fmt || '');
  const f = (fmt || 'yyyy-mm-dd').split(';')[0].replace(/\[[^\]]*\]/g, '');
  const Y = d.getUTCFullYear(), M = d.getUTCMonth(), D = d.getUTCDate(), W = d.getUTCDay();
  const DAY_EN = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  return f.replace(/"([^"]*)"|\\(.)|yyyy|yy|bbbb|bb|mmmm|mmm|mm|m|dddd|ddd|dd|d|e/gi, (t, q, esc) => {
    if (q != null) return q; if (esc != null) return esc;
    switch (t.toLowerCase()) {
      case 'yyyy': case 'e': return String(Y); case 'yy': return pad(Y % 100);
      case 'bbbb': return String(Y + 543); case 'bb': return pad((Y + 543) % 100);
      case 'mmmm': return (th ? MONTH_TH : MONTH_EN)[M]; case 'mmm': return th ? MON_TH[M] : MONTH_EN[M].slice(0, 3);
      case 'mm': return pad(M + 1); case 'm': return String(M + 1);
      case 'dddd': return DAY_EN[W]; case 'ddd': return DAY_EN[W].slice(0, 3);
      case 'dd': return pad(D); default: return String(D);
    }
  });
}
export const formatValue = (v, f) => (typeof v === 'number' ? (isDateFmt(f) ? formatDate(fromSerial(v), f) : formatNumber(v, f)) : String(v ?? ''));

// ---------- styles ----------
export function themeColors(wb) {
  const xml = wb._themes?.theme1 || '';
  return ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink']
    .map((n) => new RegExp(`<a:${n}>[\\s\\S]*?(?:srgbClr val|lastClr)="([0-9A-Fa-f]{6})"`).exec(xml)?.[1] || null);
}
export function cssColor(c, theme) {
  const hex = c?.argb ? c.argb.slice(-6) : c?.theme != null ? theme[c.theme] : null;
  if (!hex) return null;
  const t = c.tint || 0;
  const rgb = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16)).map((x) => Math.round(t < 0 ? x * (1 + t) : x + (255 - x) * t));
  return `rgb(${rgb.join(',')})`;
}
export const hexOf = (c, theme) => { const s = cssColor(c, theme); if (!s) return null; return `#${s.match(/\d+/g).map((x) => (+x).toString(16).padStart(2, '0')).join('')}`; };
const BORDER = { thin: '1px solid', hair: '1px solid', medium: '2px solid', thick: '3px solid', double: '3px double', dotted: '1px dotted', dashed: '1px dashed', dashDot: '1px dashed', dashDotDot: '1px dotted', mediumDashed: '2px dashed', mediumDashDot: '2px dashed', mediumDashDotDot: '2px dotted', slantDashDot: '2px dashed' };
// A cell's own look as CSS: font, fill, borders, alignment, wrapping.
export function cellCss(cell, theme, value) {
  const out = []; const f = cell.font || {};
  if (f.name) out.push(`font-family:'${f.name}',sans-serif`);
  if (f.size) out.push(`font-size:${f.size}pt`);
  if (f.bold) out.push('font-weight:bold');
  if (f.italic) out.push('font-style:italic');
  const deco = [f.underline && 'underline', f.strike && 'line-through'].filter(Boolean); if (deco.length) out.push(`text-decoration:${deco.join(' ')}`);
  const fc = cssColor(f.color, theme); if (fc) out.push(`color:${fc}`);
  if (cell.fill?.type === 'pattern' && cell.fill.pattern && cell.fill.pattern !== 'none') { const bg = cssColor(cell.fill.fgColor, theme); if (bg) out.push(`background:${bg}`); }
  const b = cell.border || {};
  for (const side of ['top', 'right', 'bottom', 'left']) { const e = b[side]; if (e?.style) out.push(`border-${side}:${BORDER[e.style] || '1px solid'} ${cssColor(e.color, theme) || '#000'}`); }
  const a = cell.alignment || {};
  const h = a.horizontal && a.horizontal !== 'general' ? ({ centerContinuous: 'center', distributed: 'justify', fill: 'left' }[a.horizontal] || a.horizontal) : typeof value === 'number' ? 'right' : 'left';
  out.push(`text-align:${h}`, `justify-content:${h === 'right' ? 'flex-end' : h === 'center' ? 'center' : 'flex-start'}`);
  out.push(`align-items:${a.vertical === 'top' ? 'flex-start' : a.vertical === 'middle' ? 'center' : 'flex-end'}`);
  if (a.wrapText) out.push('white-space:pre-wrap', 'overflow-wrap:anywhere');
  if (a.indent) out.push(`padding-left:${4 + a.indent * 9}px`);
  return out.join(';');
}

// ---------- formula references ----------
// Calls f(sheetQualifier, ref) for each A1 reference outside string literals; f returns the replacement text.
// A reference: optional sheet ('Sheet 1'! or Sheet1!), then $A$1 or $A$1:$B$2.
const REF = /('(?:[^']|'')+'!|[A-Za-z_][\w.]*!)?(\$?)([A-Z]{1,3})(\$?)(\d{1,7})(?::(\$?)([A-Z]{1,3})(\$?)(\d{1,7}))?(?![\w(])/g;
export function mapRefs(formula, f) {
  return formula.split(/("(?:[^"]|"")*")/).map((part, i) => (i % 2 ? part : part.replace(REF, (m, sh, ac1, c1, ar1, r1, ac2, c2, ar2, r2, off, s) => {
    const prev = s[off - 1]; if (prev && /[\w.]/.test(prev)) return m; // part of a name like LOG10 or a defined name
    const a = { col: colNum(c1), row: +r1, absC: !!ac1, absR: !!ar1 };
    const b = c2 ? { col: colNum(c2), row: +r2, absC: !!ac2, absR: !!ar2 } : null;
    return f(sh ? sh.slice(0, -1).replace(/^'|'$/g, '').replace(/''/g, "'") : null, a, b, m) ?? m;
  }))).join('');
}
const refText = (sh, a, b) => {
  const one = (x) => `${x.absC ? '$' : ''}${colName(x.col)}${x.absR ? '$' : ''}${x.row}`;
  const q = sh == null ? '' : /^[A-Za-z_][\w.]*$/.test(sh) ? `${sh}!` : `'${sh.replace(/'/g, "''")}'!`;
  return `${q}${one(a)}${b ? `:${one(b)}` : ''}`;
};
// Copy/fill: relative parts move by (dr, dc); a reference that would leave the sheet becomes #REF!.
export function shiftFormula(formula, dr, dc) {
  return mapRefs(formula, (sh, a, b) => {
    const mv = (x) => ({ ...x, row: x.absR ? x.row : x.row + dr, col: x.absC ? x.col : x.col + dc });
    const na = mv(a); const nb = b && mv(b);
    if (na.row < 1 || na.col < 1 || (nb && (nb.row < 1 || nb.col < 1))) return '#REF!';
    return refText(sh, na, nb);
  });
}
// Insert (n > 0) or delete (n < 0) rows or columns at `at` in sheet `target`; `own` is the sheet the formula is in.
// References after `at` move; a single reference into deleted rows/columns becomes #REF!, a range shrinks.
export function adjustFormula(formula, own, target, axis, at, n) {
  const key = axis === 'row' ? 'row' : 'col';
  return mapRefs(formula, (sh, a, b) => {
    if ((sh ?? own) !== target) return null;
    const move = (v) => (n > 0 ? (v >= at ? v + n : v) : v >= at - n ? v + n : v < at ? v : null);
    if (!b) { const v = move(a[key]); return v == null ? '#REF!' : refText(sh, { ...a, [key]: v }, null); }
    let v1 = a[key], v2 = b[key];
    if (n > 0) { v1 = move(v1); v2 = move(v2); } else {
      const end = at - n - 1;
      if (v1 >= at && v2 <= end) return '#REF!';
      v1 = v1 < at ? v1 : v1 <= end ? at : v1 + n; v2 = v2 < at ? v2 : v2 <= end ? at - 1 : v2 + n;
    }
    return refText(sh, { ...a, [key]: v1 }, { ...b, [key]: v2 });
  });
}

// ---------- calculation ----------
// HyperFormula's own TEXT() mishandles Excel formats ("#,##0.00" gives "9123457,##0.00"): use the editor's formats.
let textFixed = false;
export function fixText(HF) {
  if (textFixed || !HF.FunctionPlugin) return; textFixed = true;
  class TextPlugin extends HF.FunctionPlugin {
    text(ast, state) { return this.runFunction(ast.args, state, this.metadata('TEXT'), (v, f) => formatValue(v, f)); }
  }
  TextPlugin.implementedFunctions = { TEXT: { method: 'text', parameters: [{ argumentType: HF.FunctionArgumentType.SCALAR }, { argumentType: HF.FunctionArgumentType.STRING }] } };
  try { HF.unregisterFunction('TEXT'); HF.registerFunctionPlugin(TextPlugin, { enGB: { TEXT: 'TEXT' } }); } catch (e) { console.warn('TEXT format fix not applied:', e); }
}

// Argument hints for the functions finance work uses most (others list their name only).
export const FN_HELP = {
  SUM: 'SUM(number1, [number2], …)', SUMIF: 'SUMIF(range, criteria, [sum_range])', SUMIFS: 'SUMIFS(sum_range, criteria_range1, criteria1, …)',
  AVERAGE: 'AVERAGE(number1, [number2], …)', AVERAGEIF: 'AVERAGEIF(range, criteria, [average_range])', AVERAGEIFS: 'AVERAGEIFS(average_range, criteria_range1, criteria1, …)',
  COUNT: 'COUNT(value1, [value2], …)', COUNTA: 'COUNTA(value1, [value2], …)', COUNTIF: 'COUNTIF(range, criteria)', COUNTIFS: 'COUNTIFS(criteria_range1, criteria1, …)',
  MIN: 'MIN(number1, [number2], …)', MAX: 'MAX(number1, [number2], …)', IF: 'IF(logical_test, value_if_true, [value_if_false])', IFS: 'IFS(test1, value1, [test2, value2], …)',
  IFERROR: 'IFERROR(value, value_if_error)', AND: 'AND(logical1, [logical2], …)', OR: 'OR(logical1, [logical2], …)',
  VLOOKUP: 'VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])', HLOOKUP: 'HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])',
  XLOOKUP: 'XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])',
  INDEX: 'INDEX(array, row_num, [column_num])', MATCH: 'MATCH(lookup_value, lookup_array, [match_type])',
  ROUND: 'ROUND(number, num_digits)', ROUNDUP: 'ROUNDUP(number, num_digits)', ROUNDDOWN: 'ROUNDDOWN(number, num_digits)',
  TEXT: 'TEXT(value, format_text)', DATE: 'DATE(year, month, day)', EDATE: 'EDATE(start_date, months)', EOMONTH: 'EOMONTH(start_date, months)',
  TODAY: 'TODAY()', YEAR: 'YEAR(date)', MONTH: 'MONTH(date)', DAY: 'DAY(date)', DATEDIF: 'DATEDIF(start_date, end_date, unit)',
  PMT: 'PMT(rate, nper, pv, [fv], [type])', NPV: 'NPV(rate, value1, [value2], …)', IRR: 'IRR(values, [guess])', FV: 'FV(rate, nper, pmt, [pv], [type])', PV: 'PV(rate, nper, pmt, [fv], [type])',
  ABS: 'ABS(number)', CONCAT: 'CONCAT(text1, [text2], …)', CONCATENATE: 'CONCATENATE(text1, [text2], …)', LEFT: 'LEFT(text, [num_chars])', RIGHT: 'RIGHT(text, [num_chars])', MID: 'MID(text, start_num, num_chars)',
  LEN: 'LEN(text)', TRIM: 'TRIM(text)', SUMPRODUCT: 'SUMPRODUCT(array1, [array2], …)', SUBTOTAL: 'SUBTOTAL(function_num, ref1, …)',
};
export const FN_GROUPS = {
  'Math & totals': ['SUM', 'SUMIF', 'SUMIFS', 'SUMPRODUCT', 'SUBTOTAL', 'ROUND', 'ROUNDUP', 'ROUNDDOWN', 'ABS', 'MIN', 'MAX'],
  Statistics: ['AVERAGE', 'AVERAGEIF', 'AVERAGEIFS', 'COUNT', 'COUNTA', 'COUNTIF', 'COUNTIFS'],
  Logic: ['IF', 'IFS', 'IFERROR', 'AND', 'OR'], Lookup: ['XLOOKUP', 'VLOOKUP', 'HLOOKUP', 'INDEX', 'MATCH'],
  'Date & time': ['TODAY', 'DATE', 'EDATE', 'EOMONTH', 'YEAR', 'MONTH', 'DAY', 'DATEDIF'], Finance: ['PMT', 'NPV', 'IRR', 'FV', 'PV'],
  Text: ['TEXT', 'CONCAT', 'CONCATENATE', 'LEFT', 'RIGHT', 'MID', 'LEN', 'TRIM'],
};
