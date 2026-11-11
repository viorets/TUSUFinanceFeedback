// PDF -> structured text (headings, paragraphs, lists, simple tables) -> Markdown.
// Reads the text a PDF already holds (pdf.js getTextContent), so it works for PDFs made on a computer; scanned
// pages have no text and are listed in `scanned` (OCR is a later step). Thai is cleaned on the way in: legacy
// "shifted mark" characters (U+F700–F71A), ํ + า written for ำ, and marks stored out of order.

// Microsoft/Apple private-use Thai glyph variants (marks shifted left or down, letters without their tail).
const PUA = {
  0xF700: 0x0E10, 0xF701: 0x0E34, 0xF702: 0x0E35, 0xF703: 0x0E36, 0xF704: 0x0E37, 0xF705: 0x0E48, 0xF706: 0x0E49,
  0xF707: 0x0E4A, 0xF708: 0x0E4B, 0xF709: 0x0E4C, 0xF70A: 0x0E48, 0xF70B: 0x0E49, 0xF70C: 0x0E4A, 0xF70D: 0x0E4B,
  0xF70E: 0x0E4C, 0xF70F: 0x0E0D, 0xF710: 0x0E31, 0xF711: 0x0E4D, 0xF712: 0x0E47, 0xF713: 0x0E48, 0xF714: 0x0E49,
  0xF715: 0x0E4A, 0xF716: 0x0E4B, 0xF717: 0x0E4C, 0xF718: 0x0E38, 0xF719: 0x0E39, 0xF71A: 0x0E3A,
};
// Common words in which Word's PDF export turns ำ into า. None of the misspellings is itself a Thai word.
const AM_WORDS = Object.fromEntries(['กำไร', 'ชำระ', 'จำหน่าย', 'สำหรับ', 'จำนวน', 'กำหนด', 'สำคัญ', 'ดำเนิน', 'คำนวณ', 'ตำแหน่ง', 'กำกับ',
  'สำนัก', 'ทำให้', 'ประจำปี', 'อำนาจ', 'สำรอง', 'ลำดับ', 'ทำหน้าที่', 'นำไป', 'นำมา', 'ทำงาน', 'จำกัด', 'คำนึง', 'สำเร็จ', 'คำสั่ง',
  'ทำการ', 'กำลัง', 'ดำรง', 'จำเป็น', 'สำเนา', 'ค้ำประกัน', 'อำนวย', 'ตำบล', 'อำเภอ', 'สำนักงาน', 'ทำรายการ', 'จำแนก'].map((w) => [w.replace(/\u0E33/g, '\u0E32'), w]));
const LOST_AM = new RegExp(Object.keys(AM_WORDS).sort((a, b) => b.length - a.length).join('|'), 'g');
export function cleanThai(s) {
  return s
    .replace(/[-]/g, (c) => String.fromCharCode(PUA[c.charCodeAt(0)]))
    .replace(/ํ([่-๋]?)า/g, '$1ำ') // ํ (tone) า -> (tone) ำ
    .replace(/ำา/g, 'ำ') // ำ copied with its own า tail
    .replace(/([ก-ฮ][่-๋]?) (?=ำ)/g, '$1') // "จ ำกัด": ำ never starts a word (Word's PDF export)
    .replace(/([ก-ฮ][่-๋]?) า/g, '$1ำ') // "ก าหนด": nor does า, so this was ำ
    .replace(/\.{4,}/g, ' … ') // contents-page dot leaders
    .replace(/([่-์])([ัิ-ฺ็])/g, '$2$1') // vowel mark before tone mark
    .replace(/ /g, ' ')
    .replace(/[\u0000-\u0008\u000B-\u001F\u007F]/g, '') // control characters some PDFs map unknown shapes to
    .normalize('NFC');
}

const MARKS = /^[\p{M}ัิ-ฺ็-๎-]+$/u;
const THAI = /[฀-๿]/;
// "ABCDEF+THSarabunNew-Bold" -> "TH Sarabun New"; "Kanit-Regular" -> "Kanit"; "Arial,Bold" -> "Arial".
const KNOWN_FONTS = { thsarabunnew: 'TH Sarabun New', thsarabunpsk: 'TH SarabunPSK', angsananew: 'Angsana New', angsanaupc: 'AngsanaUPC',
  cordianew: 'Cordia New', browallianew: 'Browallia New', leelawadeeui: 'Leelawadee UI', leelawadee: 'Leelawadee', tahoma: 'Tahoma',
  timesnewromanpsmt: 'Times New Roman', timesnewroman: 'Times New Roman', arialmt: 'Arial', calibri: 'Calibri', cambria: 'Cambria' };
function familyOf(name) {
  const base = String(name).replace(/^[A-Z]{6}\+/, '').split(/[-,]/)[0].replace(/(Bold|Italic|Oblique|Regular|Light|Medium|Semi|Black|PS|MT)+$/i, (m) => (/^(PS|MT)$/i.test(m) ? m : ''));
  const key = base.toLowerCase().replace(/[^a-z0-9]/g, '');
  if (!key || /^(sansserif|serif|monospace|g_?d\d)/.test(key)) return '';
  return KNOWN_FONTS[key] || base.replace(/([a-z])([A-Z])/g, '$1 $2');
}
// A figure in a statement: 1,245,300.50, (15,000.00), 12%, ฿500, or "-" for nil.
const FIG = /^(?:[-+(]?\s?[฿$€£¥]?\s?[\d๐-๙][\d๐-๙,]*(?:\.[\d๐-๙]+)?\s?%?\)?|[-–—])$/;
const LEADER = /\s*(?:\.{4,}|…{2,}|(?:\. ){3,}\.?)\s*/; // dot leaders between a label and its figure
const MARKER_ONLY = /^(?:[•●▪■◦○◆❖➢➤►✓✔\-–—*·]|\(?(?:\d{1,3}|[๐-๙]{1,3}|[a-zA-Zก-ฮ])[.)])$/;
const BOLD = /bold|black|heavy|semibold|demi|,b$/i;
const ITALIC = /italic|oblique/i;
const LIST = /^\s*(?:([•●▪■◦○◆❖➢➤►✓✔\-–—*·])|(\d{1,3})[.)]|([๐-๙]{1,3}|[a-zA-Zก-ฮ])[.)]|\(([\d๐-๙]{1,3}|[a-zA-Zก-ฮ])\))\s+/;

// One page -> lines (top to bottom). A line is cells (split at wide gaps); a cell is runs { t, b, i }.
// items: text items from OCR (js/pdf-ocr.js) instead of the page's own text.
async function pageLines(page, items = null, { findGrids = true } = {}) {
  const ops = await page.getOperatorList(); // also loads the fonts, so their names (Bold, Italic) can be read
  const tc = items ? { items, styles: {} } : await page.getTextContent();
  // A font that writes ำ where า belongs (TH Sarabun bold in Word's PDF export: "กิจกำร"): it has many ำ and no า.
  const amaa = {};
  tc.items.forEach((it) => { const c = (amaa[it.fontName] ||= [0, 0]); c[0] += (it.str?.match(/ำ/g) || []).length; c[1] += (it.str?.match(/า/g) || []).length; });
  const swapAm = (name) => amaa[name] && amaa[name][0] >= 4 && amaa[name][1] === 0;
  const fontOf = {};
  for (const name of Object.keys(tc.styles)) {
    let f = null; try { f = page.commonObjs.get(name); } catch { /* not loaded */ }
    const n = f?.name || tc.styles[name].fontFamily || '';
    fontOf[name] = { b: !!f?.bold || BOLD.test(n), i: !!f?.italic || ITALIC.test(n), family: familyOf(n) };
  }
  const vp = page.getViewport({ scale: 1 });
  // Items -> chunks; a chunk that is only combining marks belongs to the chunk before it (stream order), so marks
  // drawn slightly left of their letter do not jump ahead of it when the line is sorted by x.
  const chunks = [];
  // Text drawn again over itself (Canva's outline and shadow effects draw each piece or each line two or three
  // times): a piece that repeats one already on the page at almost the same place is left out, with the marks that follow it.
  const drawn = new Map(); let dup = false;
  for (let it of tc.items) {
    if (!('str' in it) || !it.str.trim()) continue; // spaces and tabs: the gap between chunks says where they were
    const [a, b, c, d, e, f] = it.transform;
    if (Math.abs(b) > 0.01 || Math.abs(c) > 0.01) continue; // rotated text (stamps, side notes)
    const size = Math.abs(d) || Math.abs(a) || it.height; const st = fontOf[it.fontName] || {};
    if (MARKS.test(it.str) ? dup : (dup = !!drawn.get(it.str)?.some((o) => Math.abs(o.x - e) < size * 0.3 && Math.abs(o.y - f) < size * 0.3))) continue;
    if (!drawn.has(it.str)) drawn.set(it.str, []); drawn.get(it.str).push({ x: e, y: f });
    if (swapAm(it.fontName)) it = { ...it, str: it.str.replace(/ำ/g, 'า') };
    const prev = chunks[chunks.length - 1];
    if (prev && MARKS.test(it.str) && Math.abs(prev.y - f) < size * 0.6) { prev.runs.push({ t: it.str, b: st.b, i: st.i }); prev.x2 = Math.max(prev.x2, e + it.width); continue; }
    // "Cash ........ 1,200.00" in one item: label and figure become two chunks (positions shared out by letters).
    // Only amounts (with , . or brackets): a contents line's page number stays on its line.
    const ld = LEADER.exec(it.str); const tail = ld && it.str.slice(ld.index + ld[0].length);
    if (ld && ld.index && FIG.test(tail.trim()) && /[\d๐-๙][,.][\d๐-๙]|\(/.test(tail)) {
      const per = it.width / it.str.length; const head = it.str.slice(0, ld.index);
      chunks.push({ x: e, x2: e + head.length * per, y: f, size, fam: st.family, runs: [{ t: head, b: st.b, i: st.i }] });
      chunks.push({ x: e + it.width - tail.length * per, x2: e + it.width, y: f, size, fam: st.family, runs: [{ t: tail, b: st.b, i: st.i }] });
      continue;
    }
    chunks.push({ x: e, x2: e + it.width, y: f, size, fam: st.family, runs: [{ t: it.str, b: st.b, i: st.i }] });
  }
  // Bordered tables: the page's ruling lines give the cell grid; text inside a grid goes to its cells.
  const tables = (findGrids ? gridTables(rulesOf(ops, vp)) : []).map((t) => ({ ...t, cells: t.ys.slice(1).map(() => t.xs.slice(1).map(() => [])) }));
  const rest = [];
  for (const ch of chunks) {
    const cx = (ch.x + ch.x2) / 2; const cy = ch.y + ch.size * 0.3;
    const t = tables.find((g) => cx > g.xs[0] && cx < g.xs.at(-1) && cy < g.ys[0] && cy > g.ys.at(-1));
    if (!t) { rest.push(ch); continue; }
    const i = t.ys.findIndex((y, k) => k && cy > y) - 1; const j = t.xs.findIndex((x, k) => k && ch.x + 1 < x) - 1;
    t.cells[Math.max(i, 0)][Math.max(j, 0)].push(ch);
  }
  // A grid with text in only one column is a box around text (a frame, a callout): its text goes back to the page.
  const grids = [];
  tables.forEach((t) => {
    const used = t.xs.slice(1).map((_, j) => t.cells.some((r) => r[j].length)).filter(Boolean).length;
    if (used < 2) { t.cells.flat(2).forEach((ch) => rest.push(ch)); return; }
    let rows = t.cells.map((r) => r.map((chs) => cellRuns(linesOf(chs, vp))));
    rows = rows.filter((r) => r.some((c) => c.length));
    const keep = rows[0].map((_, j) => rows.some((r) => r[j].length));
    const widths = t.xs.slice(1).map((x, j) => x - t.xs[j]).filter((_, j) => keep[j]);
    grids.push({ y: t.ys[0], top: vp.height - t.ys[0], rows: rows.map((r) => r.filter((_, j) => keep[j])), widths });
  });
  grids.sort((p, q) => q.y - p.y);
  const images = ops.fnArray.filter((fn) => fn === OPS.paintImageXObject || fn === OPS.paintInlineImageXObject || fn === OPS.paintImageXObjectRepeat).length;
  const fonts = {};
  tc.items.forEach((it) => { const fam = fontOf[it.fontName]?.family; if (fam && it.str) fonts[fam] = (fonts[fam] || 0) + it.str.length; });
  return { lines: linesOf(rest, vp), tables: grids, height: vp.height, width: vp.width, images, fonts };
}

// Chunks -> lines (top to bottom); a line's chunks split into cells at wide gaps.
function linesOf(chunks, vp) {
  const lines = [];
  for (const ch of chunks) {
    const ln = lines.find((l) => Math.abs(l.y - ch.y) < Math.min(l.size, ch.size) * 0.45);
    if (ln) { ln.items.push(ch); ln.size = Math.max(ln.size, ch.size); } else lines.push({ y: ch.y, size: ch.size, items: [ch] });
  }
  lines.sort((p, q) => q.y - p.y);
  lines.forEach((ln) => ln.items.sort((p, q) => p.x - q.x));
  // Column edges: x positions where text starts on 3 or more lines (at the line start or after a gap). A gap that
  // ends on a column edge separates columns; a gap elsewhere is justified spacing inside a sentence.
  const starts = new Map();
  const mark = (x) => { const k = Math.round(x / 2); starts.set(k, (starts.get(k) || 0) + 1); };
  lines.forEach((ln) => ln.items.forEach((ch, k) => { if (!k || ch.x - ln.items[k - 1].x2 > ln.size * 0.9) mark(ch.x); }));
  const edge = (x) => [-1, 0, 1].reduce((n, d) => n + (starts.get(Math.round(x / 2) + d) || 0), 0) >= 3;
  // Figures are right-aligned, so their columns show where they end: x positions where a figure ends on 3 or more lines.
  const textOf = (ch) => ch.runs.map((r) => r.t).join('').trim();
  const ends = new Map();
  lines.forEach((ln) => ln.items.forEach((ch) => { if (FIG.test(textOf(ch))) { const k = Math.round(ch.x2 / 2); ends.set(k, (ends.get(k) || 0) + 1); } }));
  const figEdge = (ch) => FIG.test(textOf(ch)) && [-1, 0, 1].reduce((n, d) => n + (ends.get(Math.round(ch.x2 / 2) + d) || 0), 0) >= 3;
  for (const ln of lines) {
    ln.cells = []; let cell = null; let last = null;
    for (const ch of ln.items) {
      const gap = last ? ch.x - last.x2 : 0;
      if (!cell || gap > ln.size * 3 || (gap > ln.size * 0.9 && (edge(ch.x) || figEdge(ch))) || (gap > ln.size * 0.5 && figEdge(ch) && figEdge(last))) { cell = { x: ch.x, x2: ch.x2, runs: [], fam: ch.fam }; ln.cells.push(cell); }
      else if (gap > ln.size * 0.18 && !/\s$/.test(last.runs.at(-1).t) && !/^\s/.test(ch.runs[0].t)) cell.runs.push({ t: ' ', b: false, i: false });
      cell.runs.push(...ch.runs); cell.x2 = Math.max(cell.x2, ch.x2); last = ch;
    }
    ln.cells.forEach((c) => {
      c.runs = mergeRuns(c.runs.map((r) => ({ ...r, t: cleanThai(r.t) })));
      if (ln.cells.length > 1) { const r = c.runs.at(-1); r.t = r.t.replace(new RegExp(`${LEADER.source}$`), ''); } // leaders before the next column
      c.text = c.runs.map((r) => r.t).join('').trim();
    });
    ln.cells = ln.cells.filter((c) => c.text);
    // "1." or "•" alone in the first cell is a list marker with a hanging indent, not a column
    if (ln.cells.length >= 2 && MARKER_ONLY.test(ln.cells[0].text)) {
      const [m, c] = ln.cells; ln.cells.splice(0, 2, { x: m.x, x2: c.x2, runs: mergeRuns([...m.runs, { t: ' ', b: false, i: false }, ...c.runs]), text: `${m.text} ${c.text}` });
    }
    ln.x = ln.cells[0]?.x ?? 0; ln.x2 = ln.cells.at(-1)?.x2 ?? 0; ln.text = ln.cells.map((c) => c.text).join(' ');
    ln.top = vp.height - ln.y;
  }
  return lines.filter((l) => l.text);
}

// The text of one table cell: its lines joined as a paragraph, except that a line starting with a list marker
// starts a new line in the cell ("\n").
function cellRuns(lines) {
  const wide = Math.max(...lines.map((l) => l.x2), 0);
  return lines.reduce((out, l, k) => addLine(out, l, lines[k - 1], wide), []);
}
// Adds line l to a cell's runs: a new line ("\n") when l starts a list item or the line before it stopped short;
// otherwise the same paragraph.
function addLine(out, l, prev, wide) {
  const runs = runsOf(l);
  if (!out.length) return runs;
  if (LIST.test(l.text) || /^[-•●▪*]/.test(l.text) || (prev && prev.x2 < wide - prev.size * 3)) return mergeRuns([...out, { t: '\n', b: false, i: false }, ...runs]);
  return joinRuns(out, runs);
}

// pdf.js operator numbers used here (pdf.js 4.x OPS).
const OPS = { save: 10, restore: 11, transform: 12, moveTo: 13, lineTo: 14, curveTo: 15, curveTo2: 16, curveTo3: 17, closePath: 18,
  rectangle: 19, stroke: 20, closeStroke: 21, fill: 22, eoFill: 23, fillStroke: 24, eoFillStroke: 25, closeFillStroke: 26,
  closeEOFillStroke: 27, endPath: 28, setFont: 37, paintFormXObjectBegin: 74, paintFormXObjectEnd: 75, paintImageXObject: 85,
  paintInlineImageXObject: 86, paintImageXObjectRepeat: 88, constructPath: 91 };
export const PDF_OPS = OPS; // checked against the loaded pdf.js in testing
const mul = (m, n) => [m[0] * n[0] + m[2] * n[1], m[1] * n[0] + m[3] * n[1], m[0] * n[2] + m[2] * n[3], m[1] * n[2] + m[3] * n[3], m[0] * n[4] + m[2] * n[5] + m[4], m[1] * n[4] + m[3] * n[5] + m[5]];
const at = (m, x, y) => [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]];

// Horizontal and vertical ruling lines on the page (stroked lines, thin filled bars, rectangle edges), in PDF units.
function rulesOf(ops, vp) {
  const H = []; const V = []; let m = [1, 0, 0, 1, 0, 0]; const stack = []; const boxes = [];
  const paints = new Set([OPS.stroke, OPS.closeStroke, OPS.fill, OPS.eoFill, OPS.fillStroke, OPS.eoFillStroke, OPS.closeFillStroke, OPS.closeEOFillStroke]);
  const seg = ([x1, y1], [x2, y2]) => {
    if (Math.abs(y1 - y2) < 0.8 && Math.abs(x1 - x2) > 4) H.push({ y: (y1 + y2) / 2, a: Math.min(x1, x2), b: Math.max(x1, x2) });
    else if (Math.abs(x1 - x2) < 0.8 && Math.abs(y1 - y2) > 4) V.push({ x: (x1 + x2) / 2, a: Math.min(y1, y2), b: Math.max(y1, y2) });
  };
  const { fnArray: fn, argsArray: args } = ops;
  for (let i = 0; i < fn.length; i++) {
    const f = fn[i]; const a = args[i];
    if (f === OPS.save) stack.push(m);
    else if (f === OPS.restore) m = stack.pop() || m;
    else if (f === OPS.transform) m = mul(m, a);
    else if (f === OPS.paintFormXObjectBegin) { stack.push(m); if (a?.[0]) m = mul(m, a[0]); }
    else if (f === OPS.paintFormXObjectEnd) m = stack.pop() || m;
    else if (f === OPS.constructPath) {
      const next = fn.slice(i + 1, i + 4).find((x) => paints.has(x) || x === OPS.endPath || x === OPS.constructPath);
      if (!paints.has(next)) continue; // clipping paths and paths never drawn
      const [sub, co] = a; let k = 0; let cur = null; let start = null;
      for (const op of sub) {
        if (op === OPS.moveTo) { cur = start = at(m, co[k], co[k + 1]); k += 2; }
        else if (op === OPS.lineTo) { const q = at(m, co[k], co[k + 1]); k += 2; if (cur) seg(cur, q); cur = q; }
        else if (op === OPS.rectangle) {
          const [x, y, w, h] = co.slice(k, k + 4); k += 4;
          const c = [at(m, x, y), at(m, x + w, y), at(m, x + w, y + h), at(m, x, y + h)];
          const wide = Math.hypot(c[1][0] - c[0][0], c[1][1] - c[0][1]); const tall = Math.hypot(c[3][0] - c[0][0], c[3][1] - c[0][1]);
          if (wide > vp.width * 0.9 && tall > vp.height * 0.9) continue; // page background
          if (tall < 4.5) seg([c[0][0], (c[0][1] + c[3][1]) / 2], [c[1][0], (c[1][1] + c[2][1]) / 2]); // a thin bar is a line
          else if (wide < 4.5) seg([(c[0][0] + c[1][0]) / 2, c[0][1]], [(c[3][0] + c[2][0]) / 2, c[3][1]]);
          else boxes.push(c); // a shaded box: its edges count unless it sits inside another (see below)
          cur = start = c[0];
        } else if (op === OPS.closePath) { if (cur && start) seg(cur, start); cur = start; }
        else if (op === OPS.curveTo) { cur = at(m, co[k + 4], co[k + 5]); k += 6; }
        else if (op === OPS.curveTo2 || op === OPS.curveTo3) { cur = at(m, co[k + 2], co[k + 3]); k += 4; }
      }
    }
  }
  // Word shades each text line with its own strip on top of the cell's shading: a box inside another box is shading,
  // not a cell, so only the outer boxes give cell edges.
  const span = (c) => { const xs = c.map((q) => q[0]); const ys = c.map((q) => q[1]); return [Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys)]; };
  const spans = boxes.map(span);
  boxes.forEach((c, i) => {
    const [a, b, d, e] = spans[i];
    if (spans.some(([p, q, r, t], j) => j !== i && p <= a + 1 && q <= b + 1 && r >= d - 1 && t >= e - 1 && (r - p) * (t - q) > (d - a) * (e - b) + 1)) return;
    for (let k = 0; k < 4; k++) seg(c[k], c[(k + 1) % 4]);
  });
  const merge = (list, key) => {
    list.sort((p, q) => p[key] - q[key] || p.a - q.a); const out = [];
    for (const l of list) {
      const o = out.find((x) => Math.abs(x[key] - l[key]) < 1.5 && l.a <= x.b + 2 && l.b >= x.a - 2);
      if (o) { o.a = Math.min(o.a, l.a); o.b = Math.max(o.b, l.b); } else out.push({ ...l });
    }
    return out;
  };
  return { H: merge(H, 'y'), V: merge(V, 'x') };
}

// Ruling lines that touch each other form a table; its grid is every distinct line position in the group.
function gridTables({ H, V }) {
  const n = H.length + V.length; const parent = Array.from({ length: n }, (_, i) => i);
  const find = (i) => (parent[i] === i ? i : (parent[i] = find(parent[i])));
  H.forEach((h, i) => V.forEach((v, j) => {
    if (v.x >= h.a - 2 && v.x <= h.b + 2 && h.y >= v.a - 2 && h.y <= v.b + 2) parent[find(i)] = find(H.length + j);
  }));
  const groups = new Map();
  for (let i = 0; i < n; i++) { const g = find(i); if (!groups.has(g)) groups.set(g, { ys: [], xs: [] }); if (i < H.length) groups.get(g).ys.push(H[i].y); else groups.get(g).xs.push(V[i - H.length].x); }
  const uniq = (xs) => xs.sort((a, b) => a - b).filter((x, i, s) => !i || x - s[i - 1] > 3);
  return [...groups.values()].map((g) => ({ ys: uniq(g.ys).reverse(), xs: uniq(g.xs) })).filter((g) => g.ys.length >= 2 && g.xs.length >= 2);
}

function mergeRuns(runs) {
  const out = [];
  // a space or tab takes the style of the text before it, so "*1.* *heading*" stays one italic run
  for (const r of runs) { const p = out.at(-1); if (p && ((p.b === r.b && p.i === r.i) || (!r.t.trim() && r.t !== '\n'))) p.t += r.t; else out.push({ ...r }); }
  return out;
}
const joinRuns = (a, b) => {
  // Thai wraps without a space; anything else gets one back.
  const end = a.at(-1)?.t.slice(-1) || ''; const start = b[0]?.t[0] || '';
  const sep = !end || /\s/.test(end) || /^\s/.test(start) ? '' : THAI.test(end) && THAI.test(start) ? '' : end === '-' && /[a-z]/i.test(start) ? null : ' ';
  const left = a.map((r) => ({ ...r }));
  if (sep === null) left.at(-1).t = left.at(-1).t.slice(0, -1);
  else if (sep) left.push({ t: sep, b: false, i: false });
  return mergeRuns([...left, ...b]);
};
const runsOf = (ln) => mergeRuns(ln.cells.flatMap((c, i) => (i ? [{ t: ' ', b: false, i: false }, ...c.runs] : c.runs)));

// Whole document -> blocks: { type: 'h', level, runs } | { type: 'p', runs } | { type: 'li', marker, ordered, runs }
// | { type: 'table', rows: [[text]] }. Also { pages, scanned: [page numbers without text] }.
// ocr(page) -> text items (optional): scanned pages and pages with damaged Thai are read again by OCR;
// onOcr(k, of, pageNumber) reports progress.
// A table that runs on from the previous page (the last thing on that page, the first on this one, same number of
// columns) joins it: a first row with an empty first cell continues the last row; a repeated header row is dropped.
function joinTables(prev, t) {
  if (prev?.type !== 'table' || prev.page !== t.page - 1 || prev.rows[0].length !== t.rows[0].length) return false;
  const text = (r) => r.map((c) => c.map((x) => x.t).join('').trim());
  let rows = t.rows;
  if (JSON.stringify(text(rows[0])) === JSON.stringify(text(prev.rows[0]))) rows = rows.slice(1);
  if (rows[0] && !text(rows[0])[0]) {
    const last = prev.rows.at(-1);
    rows[0].forEach((c, j) => { if (c.length) last[j] = last[j].length ? joinRuns(last[j], c) : c; });
    rows = rows.slice(1);
  }
  prev.rows.push(...rows); prev.page = t.page; return true;
}

// Damaged Thai: some programs store Thai shapes under the wrong codes, so letters and symbols turn up inside Thai
// words ("เ2ยประ1น;ขภาพ", Safari/Microsoft Forms), or store shifted marks as spaces ("เสี ยง", "เป น", Canva), or drop
// a tone mark that sits on an upper vowel ("เสียง" for เสี่ยง, "ที" for ที่: Canva). Copying from such a PDF gives the
// same; only reading the page image (OCR) recovers it.
const ODD = /[฀-๿][0-9A-Za-z;:@\\^_`{|}~$#<=>+*&%!?]{1,2}(?=[฀-๿])/g; // symbols inside a word
// a space where a mark was: after ั (ปั จจัย), between ํ and า (ตํ า), after เ + one letter (เป น) or เ + letters + ี/ื
// (เสี ยง); none of these ends a Thai word
const GAP = /ั (?=[ก-ฮ])|ํ า|เ[ก-ฮ](?:[ก-ฮ]?[ีื])? (?=[ก-ฮ])/g;
// Words that lost a stacked tone mark or upper vowel (ที for ที่, ชือ for ชื่อ, เปด for เปิด), as Canva's PDFs have them.
// None is a Thai word by itself. In the team's Thai documents (~15,000 words) they turn up at most once per file;
// in a 135-page Canva file, on 22 pages.
const LOST_MARK = new Set(['ที', 'นี', 'ชือ', 'อืน', 'ขึน', 'เพิม', 'เปด', 'ปด', 'ซึง', 'ทัง', 'ตัง', 'นัน', 'เมือ', 'เพือ', 'เชือ', 'ทัว', 'ครัง', 'ผู', 'แลว', 'ดวย', 'อยู']);
const wordSeg = typeof Intl !== 'undefined' && Intl.Segmenter ? new Intl.Segmenter('th', { granularity: 'word' }) : null;
export function damagedThai(t, { canva = false } = {}) {
  const thai = (t.match(/[฀-๿]+/g) || []).length; const odd = (t.match(ODD) || []).length + (t.match(GAP) || []).length;
  if (thai >= 10 && odd >= 3 && odd / thai > 0.03) return true;
  if (!wordSeg) return false;
  let words = 0; let lost = 0;
  for (const { segment } of wordSeg.segment(t)) { if (!/[ก-ฮ]/.test(segment)) continue; words++; if (LOST_MARK.has(segment)) lost++; }
  return canva ? lost >= 1 : lost >= 2 && lost / words >= 0.01;
}
// Made by Canva (Producer or Creator): its PDFs are the ones known to drop stacked tone marks.
export async function madeByCanva(doc) {
  try { const { info } = await doc.getMetadata(); return /canva/i.test(`${info?.Producer || ''} ${info?.Creator || ''}`); } catch { return false; }
}

export async function pdfToBlocks(doc, { onPage = () => {}, ocr = null, onOcr = () => {} } = {}) {
  const pages = [];
  for (let n = 1; n <= doc.numPages; n++) { onPage(n, doc.numPages); pages.push(await pageLines(await doc.getPage(n))); }
  const scanned = pages.map((p, i) => (p.lines.reduce((s, l) => s + l.text.length, 0) < 10 && p.images ? i + 1 : 0)).filter(Boolean);
  const canva = await madeByCanva(doc);
  const damaged = pages.map((p, i) => {
    const t = p.lines.map((l) => l.text).join(' ') + p.tables.flatMap((g) => g.rows.flat(2).map((r) => r.t)).join(' ');
    return damagedThai(t, { canva }) ? i + 1 : 0;
  }).filter(Boolean);
  const ocred = ocr ? [...new Set([...scanned, ...damaged])].sort((a, b) => a - b) : [];
  for (const [k, n] of ocred.entries()) {
    onOcr(k + 1, ocred.length, n);
    const page = await doc.getPage(n); pages[n - 1] = await pageLines(page, await ocr(page));
  }

  // Running headers and footers: lines in the top or bottom 9% of the page that repeat (numbers ignored) on at least
  // half the pages, and lone page numbers.
  const key = (t) => t.replace(/[\d๐-๙]+/g, '#').replace(/\s+/g, ' ').trim();
  const edge = (p, l) => l.top < p.height * 0.09 || l.top > p.height * 0.91;
  const seen = new Map();
  pages.forEach((p) => new Set(p.lines.filter((l) => edge(p, l)).map((l) => key(l.text))).forEach((k) => seen.set(k, (seen.get(k) || 0) + 1)));
  const PAGE_NO = /^[-–(]?\s*(?:page|หน้า(?:ที่)?)?\s*#\s*(?:(?:\/|of|จาก)\s*#)?\s*[-–)]?$/i;
  pages.forEach((p) => { p.lines = p.lines.filter((l) => !(edge(p, l) && (PAGE_NO.test(key(l.text)) || (pages.length >= 3 && seen.get(key(l.text)) >= Math.max(2, pages.length / 2))))); });

  // Body size: the size most characters use. Larger sizes become heading levels 1–3.
  const bySize = new Map();
  pages.forEach((p) => p.lines.forEach((l) => { const s = Math.round(l.size * 2) / 2; bySize.set(s, (bySize.get(s) || 0) + l.text.length); }));
  const body = [...bySize].sort((a, b) => b[1] - a[1])[0]?.[0] || 12;
  const big = [...bySize.keys()].filter((s) => s > body * 1.15).sort((a, b) => b - a);
  const levelOf = (size) => { const i = big.indexOf(Math.round(size * 2) / 2); return i < 0 ? 0 : Math.min(i + 1, 3); };

  const blocks = [];
  const pushTable = (t, atTop) => { if (!(atTop && joinTables(blocks.at(-1), t))) blocks.push(t); };
  let carry = null; // a paragraph that may continue on the next page
  for (const [pi, p] of pages.entries()) {
    const ls = p.lines; const first = blocks.length;
    const gaps = ls.slice(1).map((l, i) => ls[i].y - l.y).filter((g) => g > 0).sort((a, b) => a - b);
    const lead = gaps[Math.floor(gaps.length / 2)] || body * 1.4; // usual line spacing on this page
    const left = Math.min(...ls.filter((l) => l.cells.length === 1).map((l) => l.x), Infinity);
    const right = Math.max(...ls.map((l) => l.x2), 0);
    let cur = null; let table = null;
    const grids = [...p.tables];
    const flushGrids = (y) => { while (grids.length && grids[0].y > y) { close(); closeTable(); carry = null; const g = grids.shift(); pushTable({ type: 'table', rows: g.rows, widths: g.widths, page: pi }, blocks.length === first); } };
    const close = () => { if (cur) blocks.push(cur); cur = null; };
    const closeTable = () => {
      if (!table) return;
      if (table.length >= 2) blocks.push(tableBlock(table));
      else table[0].cells.forEach((c) => splitLines(c.runs).forEach((runs) => blocks.push(asBlock(runs)))); // a label beside its text
      table = null;
    };
    ls.forEach((l, i) => {
      flushGrids(l.y);
      const prev = ls[i - 1]; const gap = prev ? prev.y - l.y : 0;
      const runs = runsOf(l);
      // Rows with columns (wide gaps) gather into a table.
      const row = table?.at(-1);
      // The next line belongs to the row above when it follows at normal line spacing and either the row's last line
      // ran to the margin in words (a cell wrapped; numbers at the margin are a new row of figures), or it is one
      // piece that starts under one of the row's cells.
      const wrapped = row && row.last.cells.some((c) => c.x2 >= right - l.size * 2 && /[\p{L}\p{M}]$/u.test(c.text));
      const under = row && l.cells.length === 1 && (l.x > row.cells[0].x2 || row.cells.some((c) => Math.abs(c.x - l.x) < 3));
      if (row && gap < lead * 1.3 && (wrapped || under)) {
        l.cells.forEach((part) => {
          // the cell whose left edge is nearest, among those starting no further right than about one character
          const c = row.cells.filter((x) => x.x <= part.x + l.size * 1.5).reduce((a, x) => (!a || Math.abs(x.x - part.x) < Math.abs(a.x - part.x) ? x : a), null) || row.cells[0];
          const piece = { text: part.text, cells: [part], x2: part.x2, size: l.size };
          c.wide = Math.max(c.wide || c.x2, c.x2, part.x2);
          c.runs = addLine(c.runs, piece, c.prev || { x2: c.x2, size: l.size }, c.wide); c.prev = piece; c.text = c.runs.map((r) => r.t).join('').trim();
        });
        row.last = l; return;
      }
      // Rows with columns (wide gaps) gather into a table.
      if (l.cells.length >= 2) { close(); if (table && gap > lead * 2.2) closeTable(); l.last = l; (table ||= []).push(l); carry = null; return; }
      // A section label inside a statement ("Current assets", "สินทรัพย์หมุนเวียน"): a short one-piece line with rows
      // of the table close above and below stays in the table as a row of its own.
      const next = ls[i + 1];
      if (table && gap <= lead * 2.2 && next?.cells.length >= 2 && l.y - next.y <= lead * 2.2 && l.x2 < right - l.size * 4 && !LIST.test(l.text)) {
        l.last = l; table.push(l); carry = null; return;
      }
      closeTable();
      const level = levelOf(l.size);
      const m = LIST.exec(l.text);
      const allBold = runs.every((r) => r.b || !r.t.trim());
      if (level) {
        if (cur?.type === 'h' && cur.level === level && gap < Math.max(lead * 1.6, l.size * 1.8)) { cur.runs = joinRuns(cur.runs, runs); return; } // a title over two lines (big type, bigger gap)
        close(); carry = null; cur = { type: 'h', level, runs }; return;
      }
      if (m) {
        close(); carry = null;
        const stripped = stripMarker(runs, m[0].length);
        cur = { type: 'li', ordered: !!m[2], marker: m[2] ? `${m[2]}.` : m[1] ? '-' : m[0].trim(), runs: stripped, x: l.x }; return;
      }
      const sizeChange = prev && Math.abs(prev.size - l.size) > body * 0.1;
      // New paragraph: a size change, a wider gap, a first-line indent, or the line above stopped short of the
      // margin. A list item goes on while its lines are indented past its marker.
      const broke = sizeChange || gap > lead * 1.35 || (prev && /…\s*\S{1,4}$/.test(prev.text)); // a contents line ends at its page number
      const newPara = !cur || cur.type === 'h' || broke
        || (cur.type === 'li' ? l.x <= cur.x + l.size * 0.3 : l.x > left + l.size * 0.9 || (prev && prev.x2 < right - l.size * 4));
      if (!cur && carry && !l.cells[0].runs[0]?.b && l.x <= left + l.size * 0.9 && blocks.at(-1) === carry) { cur = blocks.pop(); cur.runs = joinRuns(cur.runs, runs); return; }
      if (newPara) { close(); cur = { type: 'p', runs, bold: allBold, lines: 1, at: { pg: pi, y: l.y, size: l.size } }; return; }
      cur.runs = joinRuns(cur.runs, runs); cur.lines = (cur.lines || 1) + 1;
    });
    close(); closeTable(); flushGrids(-Infinity);
    const lastB = blocks.at(-1);
    carry = lastB?.type === 'p' && !/[.!?:。]\s*$/.test(lastB.runs.map((r) => r.t).join('')) ? lastB : null;
  }
  // A short, all-bold, one-line paragraph is a sub-heading.
  blocks.forEach((b) => { delete b.page; });
  const sub = Math.min(big.length + 1, 4);
  // Two such lines straight after each other, same size and at line spacing, are one title over two lines.
  let prevSub = null;
  for (let i = 0; i < blocks.length; i++) {
    const b = blocks[i];
    if (!(b.type === 'p' && b.bold && b.lines === 1 && b.runs.map((r) => r.t).join('').trim().length < 80)) { prevSub = null; continue; }
    const p = prevSub?.at; const at = b.at;
    if (p && at && p.pg === at.pg && Math.abs(p.size - at.size) < 0.5 && p.y - at.y > 0 && p.y - at.y < at.size * 1.8) {
      prevSub.runs = joinRuns(prevSub.runs, b.runs); prevSub.at = at; blocks.splice(i--, 1); continue;
    }
    b.type = 'h'; b.level = sub; prevSub = b;
  }
  blocks.forEach((b) => { delete b.at; });
  // Word's PDF export can also drop ำ outright ("กาไร" for กำไร); fixed on whole paragraphs, as a word may span pieces.
  blocks.forEach((b) => (b.rows ? b.rows.flat() : [b.runs]).forEach((rs) => rs.forEach((r) => { r.t = r.t.replace(LOST_AM, (w) => AM_WORDS[w]); })));
  const fam = {}; pages.forEach((p) => Object.entries(p.fonts).forEach(([k, v]) => { fam[k] = (fam[k] || 0) + v; }));
  const font = Object.entries(fam).sort((a, b) => b[1] - a[1])[0]?.[0] || '';
  return { blocks, pages: pages.length, scanned, damaged, ocred, font, size: body, page: { w: pages[0]?.width || 595, h: pages[0]?.height || 842 } };
}

// Runs with "\n" -> one runs list per line.
function splitLines(runs) {
  const parts = [[]];
  runs.forEach((r) => r.t.split('\n').forEach((t, k) => { if (k) parts.push([]); if (t) parts.at(-1).push({ ...r, t }); }));
  return parts.filter((x) => x.some((r) => r.t.trim()));
}
// One line of runs -> a list item (if it starts with a marker) or a paragraph.
function asBlock(runs) {
  const text = runs.map((r) => r.t).join(''); const m = LIST.exec(text);
  return m ? { type: 'li', ordered: !!m[2], marker: m[2] ? `${m[2]}.` : m[1] ? '-' : m[0].trim(), runs: stripMarker(runs, m[0].length) } : { type: 'p', runs };
}

function stripMarker(runs, n) {
  const out = runs.map((r) => ({ ...r })); let k = n;
  while (k > 0 && out.length) { const r = out[0]; if (r.t.length <= k) { k -= r.t.length; out.shift(); } else { r.t = r.t.slice(k); k = 0; } }
  if (out[0]) out[0].t = out[0].t.replace(/^\s+/, '');
  return out;
}

// Rows of cells -> a grid: columns from the row with the most cells; each cell goes to the column it overlaps most.
function tableBlock(rows) {
  const anchor = rows.reduce((a, r) => (r.cells.length > a.cells.length ? r : a)).cells.map((c) => ({ x: c.x, x2: c.x2 }));
  rows.forEach((r) => r.cells.forEach((c) => { // widen the columns to everything that lands in them
    const j = colOf(anchor, c); anchor[j].x = Math.min(anchor[j].x, c.x); anchor[j].x2 = Math.max(anchor[j].x2, c.x2);
  }));
  const right = Math.max(...anchor.map((a) => a.x2));
  return {
    type: 'table', borders: false, // laid out by spacing only, as in the PDF
    widths: anchor.map((a, j) => (anchor[j + 1]?.x ?? right) - a.x), // each column runs to where the next one starts
    rows: rows.map((r) => {
      const out = anchor.map(() => []);
      r.cells.forEach((c) => out[colOf(anchor, c)].push(c.runs));
      return out.map((parts) => parts.reduce((a, b) => (a.length ? joinRuns(a, [{ t: ' ', b: false, i: false }, ...b]) : b), []));
    }),
  };
}
function colOf(cols, c) {
  let best = 0; let score = -Infinity;
  cols.forEach((k, j) => {
    const ov = Math.min(k.x2, c.x2) - Math.max(k.x, c.x);
    const s = ov > 0 ? ov : -Math.abs((k.x + k.x2) / 2 - (c.x + c.x2) / 2) - 1e4;
    if (s > score) { score = s; best = j; }
  });
  return best;
}

// ---------- Markdown ----------
const esc = (t) => t.replace(/([\\`*_[\]<>|])/g, '\\$1');
function inline(runs) {
  if (runs.some((r) => r.t.includes('\n'))) {
    const parts = [[]]; runs.forEach((r) => r.t.split('\n').forEach((t, k) => { if (k) parts.push([]); if (t) parts.at(-1).push({ ...r, t }); }));
    return parts.map(inline).filter(Boolean).join('<br>');
  }
  const rs = mergeRuns(runs).map((r) => { const [, a, body, z] = /^(\s*)(.*?)(\s*)$/s.exec(r.t); return { ...r, a, body, z, styled: !!body && (r.b || r.i) }; });
  return rs.map((r, k) => {
    if (!r.body) return r.t;
    // styled runs that touch with no space between ("**a***b*") break Markdown; use HTML tags there
    const touches = r.styled && ((!r.a && rs[k - 1]?.styled && !rs[k - 1].z) || (!r.z && rs[k + 1]?.styled && !rs[k + 1].a));
    if (touches) {
      const open = (r.b ? '<strong>' : '') + (r.i ? '<em>' : ''); const close = (r.i ? '</em>' : '') + (r.b ? '</strong>' : '');
      return r.a + open + esc(r.body) + close + r.z;
    }
    const mark = r.b && r.i ? '***' : r.b ? '**' : r.i ? '*' : '';
    return r.a + mark + esc(r.body) + mark + r.z;
  }).join('').replace(/\s+/g, ' ').replace(/(?:\s*(?:…|\.{2,})\s*){2,}/g, ' … ').trim();
}
const lineStart = (t) => t.replace(/^(#{1,6}\s|>|[-+]\s|\d+[.)]\s)/, '\\$1');

export function blocksToMd(blocks) {
  const out = [];
  blocks.forEach((b, i) => {
    const prev = blocks[i - 1];
    if (i && !(b.type === 'li' && prev?.type === 'li')) out.push('');
    if (b.type === 'h') out.push(`${'#'.repeat(b.level)} ${inline(b.runs.map((r) => ({ ...r, b: false })))}`);
    else if (b.type === 'li') out.push(`${b.ordered || b.marker === '-' ? b.marker : `- ${esc(b.marker)}`} ${inline(b.runs)}`);
    else if (b.type === 'table') {
      const w = Math.max(...b.rows.map((r) => r.length));
      const row = (r) => `| ${Array.from({ length: w }, (_, j) => inline(r[j] || [])).join(' | ')} |`;
      const rule = Array.from({ length: w }, (_, j) => (figCol(b.rows, j) ? ' ---: |' : ' --- |')).join(''); // columns of figures align right
      out.push(row(b.rows[0]), `|${rule}`, ...b.rows.slice(1).map(row));
    } else out.push(lineStart(inline(b.runs)));
  });
  return `${out.join('\n').trim()}\n`;
}

// ---------- Word (.docx) ----------
// A real Word document from the blocks: the PDF's main font and body size, Heading 1–4, bulleted and numbered lists
// (Word numbering, restarting at each list's first number), bold and italic, and tables with borders. JSZip is passed
// in (the page already loads it for the Office editors).
const x = (t) => t.replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
// Word wraps Thai only where it may break a line; a zero-width space between Thai words (found by the browser's
// Thai word breaker) gives it those places, so lines fill the page as in the PDF.
const SEG = typeof Intl !== 'undefined' && Intl.Segmenter ? new Intl.Segmenter('th', { granularity: 'word' }) : null;
const THAI_END = /[ก-๎]$/; const THAI_START = /^[ก-๎]/;
const wrapThai = (t) => {
  if (!SEG || !THAI.test(t)) return t;
  let out = ''; let prev = '';
  for (const { segment } of SEG.segment(t)) { out += (THAI_END.test(prev) && THAI_START.test(segment) ? '​' : '') + segment; prev = segment; }
  return out;
};
function wRuns(runs) {
  return mergeRuns(runs).map((r) => {
    const pr = `${r.b ? '<w:b/><w:bCs/>' : ''}${r.i ? '<w:i/><w:iCs/>' : ''}`;
    return r.t.split('\n').map((t, k) => `${k ? '<w:r><w:br/></w:r>' : ''}${t ? `<w:r>${pr ? `<w:rPr>${pr}</w:rPr>` : ''}<w:t xml:space="preserve">${x(wrapThai(t))}</w:t></w:r>` : ''}`).join('');
  }).join('');
}
const wP = (runs, ppr = '') => `<w:p>${ppr ? `<w:pPr>${ppr}</w:pPr>` : ''}${wRuns(trimRuns(runs))}</w:p>`;
function trimRuns(runs) {
  const out = mergeRuns(runs).map((r) => ({ ...r }));
  if (out[0]) out[0].t = out[0].t.replace(/^[ \t]+/, '');
  if (out.at(-1)) out.at(-1).t = out.at(-1).t.replace(/[ \t]+$/, '');
  return out.filter((r) => r.t);
}

const runsText = (runs) => runs.map((r) => r.t).join('').trim();
// a column of figures: every filled cell below the header row is a figure
const figCol = (rows, j) => { const cs = rows.slice(1).map((r) => runsText(r[j] || [])).filter(Boolean); return cs.length > 0 && cs.every((t) => FIG.test(t)); };
const NO_BORDERS = `<w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((s) => `<w:${s} w:val="nil"/>`).join('')}</w:tblBorders>`;
export async function blocksToDocx(blocks, { JSZip, font = '', size = 14, page = { w: 595, h: 842 }, title = '' } = {}) {
  const f = font || 'TH Sarabun New'; const sz = Math.round(Math.min(Math.max(size, 8), 28) * 2);
  const twip = (pt) => Math.round(pt * 20);
  const body = []; const nums = []; // one Word list per run of list items: [abstract id, start]
  let list = null;
  blocks.forEach((b) => {
    if (b.type !== 'li') list = null;
    if (b.type === 'h') body.push(wP(b.runs.map((r) => ({ ...r, b: false })), `<w:pStyle w:val="Heading${b.level}"/>`));
    else if (b.type === 'p') body.push(wP(b.runs));
    else if (b.type === 'li') {
      const kind = b.ordered ? 1 : b.marker === '-' ? 0 : 2; // 0 bullet, 1 Word numbering, 2 keep the marker as text
      if (!list || list.kind !== kind) { list = { kind, id: nums.length + 1 }; nums.push([kind, b.ordered ? parseInt(b.marker, 10) || 1 : 1]); }
      const runs = kind === 2 ? [{ t: `${b.marker} `, b: false, i: false }, ...b.runs] : b.runs;
      body.push(wP(runs, `<w:pStyle w:val="ListParagraph"/>${kind === 2 ? '<w:ind w:left="720" w:hanging="360"/>' : `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${list.id}"/></w:numPr>`}`));
    } else if (b.type === 'table') {
      const cols = Math.max(...b.rows.map((r) => r.length));
      const width = twip(page.w - 144); // the page less 1" margins
      const ws = Array.from({ length: cols }, (_, j) => Math.max(b.widths?.[j] || 0, 20)); const sum = ws.reduce((a, v) => a + v, 0);
      const cws = ws.map((v) => Math.floor((v / sum) * width));
      const rows = b.rows.map((r, i) => `<w:tr>${i === 0 ? '<w:trPr><w:tblHeader/></w:trPr>' : ''}${Array.from({ length: cols }, (_, j) => {
        const lines = splitLines(r[j] || []);
        const jc = FIG.test(runsText(r[j] || [])) ? '<w:jc w:val="right"/>' : ''; // figures align right, as in a statement
        return `<w:tc><w:tcPr><w:tcW w:w="${cws[j]}" w:type="dxa"/></w:tcPr>${(lines.length ? lines : [[]]).map((ln) => wP(ln, `<w:spacing w:before="0" w:after="0"/>${jc}`)).join('')}</w:tc>`;
      }).join('')}</w:tr>`).join('');
      body.push(`<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="${width}" w:type="dxa"/>${b.borders === false ? NO_BORDERS : ''}<w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid>${cws.map((v) => `<w:gridCol w:w="${v}"/>`).join('')}</w:tblGrid>${rows}</w:tbl>`, '<w:p/>');
    }
  });
  const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
  const doc = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${body.join('')}<w:sectPr><w:pgSz w:w="${twip(page.w)}" w:h="${twip(page.h)}"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`;
  const fonts = `<w:rFonts w:ascii="${x(f)}" w:hAnsi="${x(f)}" w:eastAsia="${x(f)}" w:cs="${x(f)}"/>`;
  const head = (lvl, k) => `<w:style w:type="paragraph" w:styleId="Heading${lvl}"><w:name w:val="heading ${lvl}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="${lvl - 1}"/></w:pPr><w:rPr><w:b/><w:bCs/><w:sz w:val="${Math.round(sz * k)}"/><w:szCs w:val="${Math.round(sz * k)}"/></w:rPr></w:style>`;
  const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${W}><w:docDefaults><w:rPrDefault><w:rPr>${fonts}<w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/><w:lang w:val="en-US" w:bidi="th-TH"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>`
    + '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>'
    + head(1, 1.6) + head(2, 1.35) + head(3, 1.15) + head(4, 1)
    + '<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:after="60"/><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style>'
    + '<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders><w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>'
    + '</w:styles>';
  const abs = (id, kind) => `<w:abstractNum w:abstractNumId="${id}"><w:multiLevelType w:val="singleLevel"/><w:lvl w:ilvl="0">${kind === 1 ? '<w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>' : '<w:numFmt w:val="bullet"/><w:lvlText w:val="•"/>'}<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>`;
  const numbering = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering ${W}>${abs(0, 0)}${abs(1, 1)}${nums.map(([kind, start], i) => `<w:num w:numId="${i + 1}"><w:abstractNumId w:val="${kind === 1 ? 1 : 0}"/>${kind === 1 ? `<w:lvlOverride w:ilvl="0"><w:startOverride w:val="${start}"/></w:lvlOverride>` : ''}</w:num>`).join('')}</w:numbering>`;
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>');
  zip.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>');
  zip.file('docProps/core.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>${x(title)}</dc:title></cp:coreProperties>`);
  zip.file('word/_rels/document.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>');
  zip.file('word/document.xml', doc); zip.file('word/styles.xml', styles); zip.file('word/numbering.xml', numbering);
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

// ---------- fonts the PDF uses ----------
// Family names of the fonts used on the first `max` pages ("TH Sarabun New", "Angsana New", ...), symbol fonts left out.
export async function pdfFonts(doc, max = 5) {
  const out = new Set();
  for (let n = 1; n <= Math.min(doc.numPages, max); n++) {
    const page = await doc.getPage(n); const ops = await page.getOperatorList();
    ops.fnArray.forEach((fn, i) => {
      if (fn !== OPS.setFont) return;
      let f = null; try { f = page.commonObjs.get(ops.argsArray[i][0]); } catch { /* not loaded */ }
      const fam = familyOf(f?.name || ''); if (fam && !/symbol|dingbat|webding|math/i.test(fam)) out.add(fam);
    });
  }
  return [...out];
}

// ---------- paragraphs to edit (PDF markup "Edit text") ----------
// The page's text as editable units: a paragraph (lines of one size that follow at normal spacing under each other),
// or one table cell line. Each unit: { text, runs, fam, size, lh (line height / size), box: [x, y, w, h] page units
// (top-left origin) covering the letters with their marks, first: first baseline (page units from the top),
// rects: [[x0, y0, x1, y1]] per line in PDF space (what to remove from the page), lines }.
// items: OCR text items for a scanned page (js/pdf-ocr.js), instead of the page's own text.
export async function pageParagraphs(page, items = null) {
  const { lines, height } = await pageLines(page, items, { findGrids: false });
  const units = []; const right = Math.max(0, ...lines.map((l) => l.x2)); // the text's right edge on this page
  const left = Math.min(...lines.map((l) => l.x), right);
  // OCR lines of an uneven scan sit a little off their places: let the line gap be up to 1.35 × the page's usual
  // line pitch as well as 1.9 × the type size, so their paragraphs hold together
  const ocr = !!items?.some((i) => i.fontName === 'ocr');
  const pitches = lines.slice(1).map((l, i) => lines[i].y - l.y).filter((g) => g > 0).sort((a, b) => a - b);
  const pitch = ocr ? pitches[Math.floor(pitches.length / 2)] || 0 : 0;
  for (const ln of lines) {
    for (const c of ln.cells) {
      const seg = { x: c.x, x2: c.x2, y: ln.y, size: ln.size, runs: c.runs, fam: c.fam, text: c.text };
      const u = ln.cells.length === 1 && !LIST.test(c.text) && units.findLast((v) => {
        const last = v.segs.at(-1); const gap = last.y - seg.y;
        return !v.closed && v.single && Math.abs(last.size - seg.size) < seg.size * 0.15 && gap > 0 && gap < Math.max(seg.size * 1.9, pitch * 1.35)
          && seg.x < v.x2 && seg.x2 > v.x
          && (Math.abs(seg.x - v.left) < seg.size * 3 || (v.segs.length === 1 && seg.x < v.left && v.left - seg.x < seg.size * 6)); // or under an indented first line
      });
      if (u) {
        const last = u.segs.at(-1);
        u.segs.push(seg); u.x = Math.min(u.x, seg.x); u.x2 = Math.max(u.x2, seg.x2); u.left = Math.min(u.left, seg.x);
      } else units.push({ segs: [seg], x: seg.x, x2: seg.x2, left: seg.x, single: ln.cells.length === 1 });
      // a line that stops well short of the text's right edge ends its paragraph (a last line, a salutation)
      if (ln.cells.length === 1 && seg.x2 < right - seg.size * 4) (u || units.at(-1)).closed = true;
    }
    // units that did not continue on this line are finished
    units.forEach((u) => { if (u.segs.at(-1).y > ln.y + ln.size * 2) u.closed = true; });
  }
  return units.map((u) => {
    const s0 = u.segs[0]; const sl = u.segs.at(-1); const size = s0.size;
    const gaps = u.segs.slice(1).map((sg, i) => u.segs[i].y - sg.y);
    const lh = gaps.length ? gaps.reduce((a, b) => a + b, 0) / gaps.length / size : 1.35;
    const runs = u.segs.slice(1).reduce((a, sg) => joinRuns(a, sg.runs), u.segs[0].runs);
    const top = height - (s0.y + size * 1.05); const bottom = height - (sl.y - size * 0.42);
    return {
      text: runs.map((r) => r.t).join(''), runs, fam: s0.fam || '', size, lh, first: height - s0.y,
      box: [u.x - 1, top, u.x2 - u.x + 2, bottom - top],
      rects: u.segs.map((sg) => [sg.x - 1, sg.y - sg.size * 0.42, sg.x2 + 1, sg.y + sg.size * 1.05]),
      lines: u.segs.length, col: [left, right],
      // centred: every line sits in the middle of the page's text column; right: every line ends at its right edge
      // and starts well in from the left (a signature block, a date line)
      align: u.segs.every((sg) => sg.x - left > sg.size * 2 && Math.abs((sg.x - left) - (right - sg.x2)) < sg.size * 1.5) ? 'center'
        : u.segs.every((sg) => right - sg.x2 < sg.size * 1.5 && sg.x - left > sg.size * 6) ? 'right' : null,
    };
  });
}
