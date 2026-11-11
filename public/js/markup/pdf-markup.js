// PDF markup editor: text box (rich text per character), highlight text, freehand draw, line, arrow, rectangle;
// adjustable width; Shift keeps lines at 45° steps and rectangles square. Marks are kept as data (page + points in PDF
// units), so a draft is a small JSON list. "Download" writes the marks into a copy of the PDF with pdf-lib; the
// original bytes are never changed. Text boxes are drawn by the browser (correct Thai shaping) and embedded as PNG.
// Libraries load only when the editor opens: pdf.js (pages + text layer) and pdf-lib (writing).
// pdf.js 4.4 needs Promise.withResolvers (Safari 17.4). Older Safari (iPad on iPadOS 15.4–17.3, v2.13) gets the
// legacy build of the same version, which brings its own polyfills; cdnjs has no legacy build, so it comes from jsDelivr.
// Tools sit in one bar on top of the pages (also in full screen).
// Zoom: Ctrl/⌘ + mouse wheel, trackpad or finger pinch, or the − / + buttons. Pages are drawn only while near the
// visible area, so long files and high zoom stay inside iPad Safari's canvas memory limit.
import { normHex, pairHex } from './hex-color.js';
import { addFontFile, fontStack, loadUserFonts, siteFont, textWriter } from './pdf-realtext.js';
import { damagedThai, madeByCanva, pageParagraphs } from './pdf-text.js';
// pdf-redact.js (Edit text) and pdf-ocr.js (scanned pages) load only when used: the Feedback site leaves both out.
import { ICON, clampZoom, floatBar, fullScreen, iconBtn, pinchZoom, viewBar, zoomAround } from './view-tools.js';

const PDFJS_BASE = Promise.withResolvers ? 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.4.168/' : 'https://cdn.jsdelivr.net/npm/pdfjs-dist@4.4.168/legacy/build/';
const PDFJS = `${PDFJS_BASE}pdf.min.mjs`;
const PDFJS_WORKER = `${PDFJS_BASE}pdf.worker.min.mjs`;
const PDFLIB = 'https://cdnjs.cloudflare.com/ajax/libs/pdf-lib/1.17.1/pdf-lib.min.js';
// Size: up to 100 MB; above 12 MB the page warns that iPad Safari may run short of memory and saving is slow
// (decided 11 Nov 2026; was a 12 MB limit). The Feedback site passes its own maxBytes (12 MB, its upload limit).
export const WARN_BYTES = 12 * 1024 * 1024;
export const MAX_BYTES = 100 * 1024 * 1024;
// Text box default (decided 20 Oct): LINE Seed Sans TH, 11 pt, #990000, clear background, no border.
export const TEXT = { font: 'LINE Seed Sans TH', size: 11, color: '#990000', lineHeight: 1.35 };
// Fonts a text box can use (js/pdf-realtext.js FACES: files and the fallback for letters a font lacks). All but
// Georgia are written into the PDF as real text; Georgia boxes stay a picture. EucrosiaUPC and Times New Roman are
// the computer's own fonts on screen.
export const FONTS = ['LINE Seed Sans TH', 'Noto Sans Thai', 'Sarabun', 'TH Sarabun New', 'TH SarabunPSK', 'EucrosiaUPC', 'Angsana New', 'Times New Roman', 'Arial', 'Georgia'];
const fontOf = (m) => m.font || TEXT.font; // text boxes from v2.12 drafts have no font: the old default
const HIGHLIGHT = '#ffe14d';
const HL_OPACITY = 0.45; // highlights from older drafts have no colour or opacity: these
const MAX_CANVAS_PX = 8e6; // pixels per page canvas (iPad Safari: 16.7M per canvas, ~384 MB for all canvases)
// Typed arrows: the text before the caret is replaced when ">" is typed. "—>" is "-->" after iPad smart punctuation.
const RESET_ICON = '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M3 12a9 9 0 1 0 3-6.7"/><path d="M3 4v5h5"/><path d="m9.5 9.5 5 5M14.5 9.5l-5 5"/></svg>';
const ARROWS = { '-->': '→', '==>': '➜', '—>': '→' };
const TOOLS = [
  ['select', 'Select', 'เลือก'], ['edit', 'Edit text', 'แก้ข้อความ'], ['text', 'Text box', 'กล่องข้อความ'], ['highlight', 'Highlight', 'ไฮไลต์'],
  ['pen', 'Draw', 'วาด'], ['line', 'Line', 'เส้น'], ['arrow', 'Arrow', 'ลูกศร'], ['rect', 'Rectangle', 'สี่เหลี่ยม'],
];
// Text formats: [execCommand name, run key, button face, English, Thai]
const FMT = [
  ['bold', 'b', 'B', 'Bold (Ctrl+B)', 'ตัวหนา'], ['italic', 'i', 'I', 'Italic (Ctrl+I)', 'ตัวเอียง'],
  ['underline', 'u', 'U', 'Underline (Ctrl+U)', 'ขีดเส้นใต้'], ['strikeThrough', 's', 'S', 'Strikethrough (Ctrl+Shift+X)', 'ขีดฆ่า'],
  ['hilite', 'hl', ICON.hl, 'Highlight', 'ไฮไลต์'],
];
const KEY = Object.fromEntries(FMT.map(([cmd, k]) => [cmd, k]));

let libs = null;
export function loadLibs() { // pdf.js and pdf-lib, loaded once (also used by js/pdf-merge.js)
  return libs || (libs = Promise.all([
    import(PDFJS).then((m) => { m.GlobalWorkerOptions.workerSrc = PDFJS_WORKER; return m; }),
    new Promise((res, rej) => {
      if (window.PDFLib) { res(window.PDFLib); return; }
      const s = document.createElement('script'); s.src = PDFLIB; s.crossOrigin = 'anonymous';
      s.onload = () => res(window.PDFLib); s.onerror = () => { libs = null; rej(new Error('Could not load the PDF tools. Try again.')); };
      document.head.append(s);
    }),
    document.fonts.load(`${TEXT.size}px "${TEXT.font}"`),
  ]));
}
// Regular, bold and italic faces of each font, so the screen and the exported PNG use the same faces.
const loadFonts = (fonts) => Promise.all([...new Set(fonts)].flatMap((f) => ['', 'bold ', 'italic ', 'italic bold ']
  .map((v) => document.fonts.load(`${v}${TEXT.size}px "${f}"`).catch(() => {}))));

// ---------- colours and rich-text runs ----------
const rgbOf = (h) => { const n = parseInt(h.slice(1), 16); return [(n >> 16) & 255, (n >> 8) & 255, n & 255]; };
const rgba = (h, a) => `rgba(${rgbOf(h).join(', ')}, ${a})`;
function parseColor(s) {
  const m = /rgba?\(\s*([\d.]+)[,\s]+([\d.]+)[,\s]+([\d.]+)(?:\s*[,/]\s*([\d.]+%?))?/i.exec(s);
  if (m) {
    const a = m[4] === undefined ? 1 : m[4].endsWith('%') ? parseFloat(m[4]) / 100 : +m[4];
    return { hex: `#${[m[1], m[2], m[3]].map((v) => Math.round(+v).toString(16).padStart(2, '0')).join('')}`, a };
  }
  const h = normHex(s); return h ? { hex: h, a: 1 } : null;
}
// A text box's content: runs [{ t, b?, i?, u?, s?, hl?, ha? }] (bold, italic, underline, strike, highlight colour and
// opacity). m.text is kept as the plain text. Boxes from older drafts have only text: one plain run.
const runsOf = (m) => m.runs || [{ t: m.text || '' }];
const styleKey = ({ t, ...s }) => JSON.stringify(Object.keys(s).sort().map((k) => [k, s[k]]));
function mergeRuns(rs) {
  const out = [];
  rs.forEach((r) => { const l = out.at(-1); if (l && styleKey(l) === styleKey(r)) l.t += r.t; else if (r.t) out.push({ ...r }); });
  return out.length ? out : [{ t: '' }];
}
// Characters [a, z) of the runs get `patch`.
function styleRange(runs, a, z, patch) {
  const out = []; let i = 0;
  runs.forEach((r) => {
    const e = i + r.t.length;
    [[i, a], [a, z], [z, e]].forEach(([s0, s1]) => {
      const lo = Math.max(s0, i); const hi = Math.min(s1, e);
      if (hi > lo) out.push({ ...r, ...(lo >= a && hi <= z ? patch : {}), t: r.t.slice(lo - i, hi - i) });
    });
    i = e;
  });
  return mergeRuns(out);
}
// One run as DOM: the same elements the browser's own Bold/Italic/... commands make, so they can undo them.
function runNode(r) {
  let n = document.createTextNode(r.t);
  const wrap = (tag) => { const e = document.createElement(tag); e.append(n); n = e; return e; };
  if (r.hl) wrap('span').style.backgroundColor = rgba(r.hl, r.ha ?? HL_OPACITY);
  if (r.s) wrap('strike');
  if (r.u) wrap('u');
  if (r.i) wrap('i');
  if (r.b) wrap('b');
  return n;
}
// The style a text node shows inside a box. Underline, strike and background are not inherited, so the ancestors are read.
function styleOf(el, box) {
  const r = {}; const cs = getComputedStyle(el);
  if (cs.fontWeight === 'bold' || parseInt(cs.fontWeight, 10) >= 600) r.b = 1;
  if (/italic|oblique/.test(cs.fontStyle)) r.i = 1;
  let bg = false;
  for (let e = el; e && e !== box; e = e.parentElement) {
    const c = getComputedStyle(e); const d = `${c.textDecorationLine || c.webkitTextDecorationLine || c.textDecoration || ''}`;
    if (d.includes('underline')) r.u = 1;
    if (d.includes('line-through')) r.s = 1;
    if (!bg && e.style.backgroundColor) { // the nearest background set on an element wins ("transparent" = removed)
      bg = true; const p = parseColor(e.style.backgroundColor);
      if (p && p.a > 0) { r.hl = p.hex; r.ha = Math.round(p.a * 100) / 100; }
    }
  }
  return r;
}
const seg = typeof Intl !== 'undefined' && Intl.Segmenter ? new Intl.Segmenter('th', { granularity: 'grapheme' }) : null;
// The edited box back to runs. A run never splits a Thai vowel or tone mark from its letter (each grapheme takes the
// style of its first character), so the exported PNG shapes Thai the same way as the screen.
function readRuns(box) {
  const chars = [];
  const walk = document.createTreeWalker(box, NodeFilter.SHOW_ELEMENT | NodeFilter.SHOW_TEXT);
  for (let n = walk.nextNode(); n; n = walk.nextNode()) {
    if (n.nodeType === 3) { const s = styleOf(n.parentElement, box); for (const ch of n.data.replace(/ /g, ' ')) chars.push([ch, s]); }
    else if (n.nodeName === 'BR') chars.push(['\n', styleOf(n.parentElement, box)]);
    else if (/^(DIV|P)$/.test(n.nodeName) && chars.length && chars.at(-1)[0] !== '\n') chars.push(['\n', {}]);
  }
  while (chars.length && chars.at(-1)[0] === '\n') chars.pop();
  const text = chars.map((c) => c[0]).join('');
  const out = []; let i = 0;
  for (const g of seg ? Array.from(seg.segment(text), (x) => x.segment) : [...text]) { out.push({ ...chars[i][1], t: g }); i += [...g].length; }
  return mergeRuns(out);
}
function caretAt(x, y) {
  if (document.caretPositionFromPoint) {
    const c = document.caretPositionFromPoint(x, y);
    if (c && c.offsetNode) { const r = document.createRange(); r.setStart(c.offsetNode, c.offset); r.collapse(true); return r; }
  }
  return document.caretRangeFromPoint ? document.caretRangeFromPoint(x, y) : null;
}
const exec = (cmd, v = null) => { try { return document.execCommand(cmd, false, v); } catch { return false; } };

// Opens the editor in `root`. bytes: the PDF (ArrayBuffer); name: file name; draft: saved marks or null.
// t(en, th) translates labels. Resolves an editor handle { marks(), isDirty(), exportPdf() }.
// defaults (admin console, v2.13): { size, textColor, shapeColor, font } of new text boxes and shapes; TEXT where not given.
// editText: the Edit text tool (and its OCR of scanned or damaged pages, when ocr is on). Feedback turns both off.
export async function openPdfEditor(root, { bytes, name, draft = null, t = (en) => en, onDraft = () => {}, onStatus = () => {}, defaults = {}, maxBytes = MAX_BYTES, editText = true, ocr = true, fonts = FONTS }) {
  const textSize = defaults.size || TEXT.size;
  // fonts: the font menu (Feedback shows four); fonts added from this computer before come after them
  const userFonts = (await loadUserFonts()).filter((f) => !fonts.includes(f));
  const fontOptions = (sel) => [...fonts, ...userFonts].map((f) => `<option ${f === sel ? 'selected' : ''}>${f}</option>`).join('');
  const font0 = fonts.includes(defaults.font) || userFonts.includes(defaults.font) ? defaults.font : fonts.includes(TEXT.font) ? TEXT.font : fonts[0];
  const limit = Math.round(maxBytes / 1048576);
  if (bytes.byteLength > maxBytes) throw new Error(t(`This file is over ${limit} MB, so it cannot be marked. Open it in Drive instead.`, `ไฟล์ใหญ่เกิน ${limit} MB จึงทำเครื่องหมายไม่ได้ กรุณาเปิดใน Drive`));
  root.innerHTML = `<div class="mk" data-tool="select"><div class="mk-bar" role="toolbar" aria-label="${t('PDF tools', 'เครื่องมือ PDF')}">
      <div class="mk-tools">${TOOLS.filter(([k]) => editText || k !== 'edit').map(([k, en, th]) => `<button type="button" data-tool="${k}" aria-pressed="false">${t(en, th)}</button>`).join('')}</div>
      <div class="mk-g mk-g-text">
        <div class="mk-fmt" role="group" aria-label="${t('Text format', 'รูปแบบตัวอักษร')}">${FMT.map(([cmd, , face, en, th]) => `<button type="button" data-fmt="${cmd}" aria-pressed="false" title="${t(en, th)}" aria-label="${t(en, th)}">${face}</button>`).join('')}</div>
        <label class="mk-ctl">${t('Font', 'แบบอักษร')} <select id="mk-f">${fontOptions(font0)}</select></label><button type="button" class="mk-addf" id="mk-addf" title="${t('Add a font from this computer (.ttf, .otf, .ttc)', 'เพิ่มแบบอักษรจากเครื่องนี้ (.ttf, .otf, .ttc)')}" aria-label="${t('Add a font from this computer', 'เพิ่มแบบอักษรจากเครื่องนี้')}">＋</button><input type="file" id="mk-ffile" accept=".ttf,.otf,.ttc,font/ttf,font/otf,font/collection" multiple hidden>
        <label class="mk-ctl">${t('Text size', 'ขนาดตัวอักษร')} <span><input type="number" id="mk-s" min="6" max="72" step="1" value="${textSize}"> pt</span></label>
      </div>
      <div class="mk-g mk-g-shape"><label class="mk-ctl">${t('Width', 'ความหนา')} <span><input type="range" id="mk-w" min="1" max="12" step="0.5" value="2"><output id="mk-wv">2</output></span></label></div>
      <div class="mk-g mk-g-color"><label class="mk-ctl">${t('Colour', 'สี')} <span><input type="color" id="mk-c" value="${defaults.shapeColor || TEXT.color}"></span></label></div>
      <div class="mk-g mk-g-hl">
        <button type="button" id="mk-hl" class="mk-hl-btn">${t('Highlight selection', 'ไฮไลต์ข้อความที่เลือก')}</button>
        <label class="mk-ctl">${t('Highlight colour', 'สีไฮไลต์')} <span><input type="color" id="mk-hc" value="${HIGHLIGHT}"></span></label>
        <label class="mk-ctl">${t('Opacity', 'ความทึบ')} <span><input type="range" id="mk-o" min="10" max="100" step="5" value="${HL_OPACITY * 100}"><output id="mk-ov">${HL_OPACITY * 100}%</output></span></label>
      </div>
      <span class="mk-grow"></span>
      ${viewBar('mk', t, { open: true, help: true })}${iconBtn('mk-reset', t('Reset edits: remove every mark (Undo brings them back)', 'ล้างการแก้ไข: ลบเครื่องหมายทั้งหมด (เลิกทำเพื่อนำกลับมา)'), RESET_ICON)}
      <p class="mk-caution" role="note"><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 9v4M12 17h.01"/><path d="M10.3 3.9 1.8 18a2 2 0 0 0 1.7 3h17a2 2 0 0 0 1.7-3L13.7 3.9a2 2 0 0 0-3.4 0z"/></svg><span>${t('Edit text: the text below an edit does not move. A longer edit can overlap it (amber outline); a shorter one leaves a gap.', 'แก้ข้อความ: ข้อความด้านล่างจะไม่เลื่อนตาม ถ้าแก้ให้ยาวขึ้นอาจทับข้อความด้านล่าง (กรอบสีส้ม) ถ้าสั้นลงจะเหลือช่องว่าง')}</span></p>
    </div><div class="mk-pages" id="mk-pages"><p class="mk-wait">${t('Loading…', 'กำลังโหลด…')}</p></div></div>`;
  const [pdfjs, PDFLib] = await loadLibs();
  await loadFonts([font0, ...(draft || []).filter((m) => m.type === 'text').map(fontOf)]);
  const doc = await pdfjs.getDocument({ data: bytes.slice(0) }).promise;
  const canva = madeByCanva(doc); // Canva PDFs can lose tone marks: such pages are read by OCR for Edit text
  const mk = root.querySelector('.mk');
  const pagesEl = root.querySelector('#mk-pages');
  const $ = (s) => root.querySelector(s);
  pagesEl.innerHTML = '';
  const syncC = pairHex($('#mk-c')); const syncHC = pairHex($('#mk-hc'));
  floatBar(mk, $('.mk-bar'));

  // Two current colours (v2.13): one for text boxes, one for shapes; the colour box shows the one in use.
  // Highlights have their own colour and opacity.
  const st = {
    tool: 'select', width: 2, textColor: defaults.textColor || TEXT.color, shapeColor: defaults.shapeColor || TEXT.color,
    hlColor: HIGHLIGHT, hlOpacity: HL_OPACITY, font: font0, size: textSize, zoom: 1,
    marks: draft ? structuredClone(draft) : [], undo: [], redo: [], sel: null, dirty: false,
  };
  const pages = []; // { el, canvas, hls, svg, boxes, text, page, base, fit, scale, w, h, gen, done, near, task }
  const avail = () => Math.max(280, pagesEl.clientWidth - 32);
  const NS = 'http://www.w3.org/2000/svg';
  const svgEl = (tag, attrs) => { const e = document.createElementNS(NS, tag); Object.entries(attrs).forEach(([k, v]) => e.setAttribute(k, v)); return e; };
  const size = (p) => {
    p.el.style.width = `${p.w * p.scale}px`; p.el.style.height = `${p.h * p.scale}px`;
    p.el.style.setProperty('--scale-factor', p.scale);
  };

  for (let n = 1; n <= doc.numPages; n++) {
    const page = await doc.getPage(n);
    const base = page.getViewport({ scale: 1 });
    const fit = Math.min(1.6, avail() / base.width);
    const el = document.createElement('div');
    el.className = 'mk-page'; el.dataset.page = n - 1;
    const canvas = document.createElement('canvas');
    const text = document.createElement('div'); text.className = 'textLayer';
    const mkSvg = (cls) => { const s = svgEl('svg', { viewBox: `0 0 ${base.width} ${base.height}` }); s.classList.add(cls); return s; };
    const hls = mkSvg('mk-hls'); const svg = mkSvg('mk-svg');
    const boxes = document.createElement('div'); boxes.className = 'mk-boxes';
    el.append(canvas, hls, text, svg, boxes);
    pagesEl.append(el);
    const p = { el, canvas, hls, svg, boxes, text, page, base, fit, scale: fit, w: base.width, h: base.height, gen: 0, done: null, near: false, task: null };
    size(p); pages.push(p);
  }

  // ---------- drawing pages (only near the visible area) ----------
  let queue = Promise.resolve();
  async function renderPage(p) {
    if (!p.near || p.done === p.scale) return;
    const gen = ++p.gen; const scale = p.scale;
    const vp = p.page.getViewport({ scale });
    const k = Math.min(window.devicePixelRatio || 1, Math.sqrt(MAX_CANVAS_PX / (vp.width * vp.height)));
    const c = document.createElement('canvas');
    c.width = Math.floor(vp.width * k); c.height = Math.floor(vp.height * k);
    p.task = p.page.render({ canvasContext: c.getContext('2d'), viewport: vp, transform: k !== 1 ? [k, 0, 0, k, 0, 0] : null });
    try { await p.task.promise; } catch { c.width = c.height = 0; return; } // cancelled
    if (gen !== p.gen) { c.width = c.height = 0; return; }
    const old = p.canvas; p.canvas = c; old.replaceWith(c); old.width = old.height = 0;
    p.done = scale;
    try {
      const tl = document.createElement('div'); tl.className = 'textLayer';
      await new pdfjs.TextLayer({ textContentSource: p.page.streamTextContent(), container: tl, viewport: vp }).render();
      if (gen === p.gen) { p.text.replaceWith(tl); p.text = tl; }
    } catch { /* no text layer: highlight is unavailable here */ }
  }
  const want = (p) => { queue = queue.then(() => renderPage(p)).catch(() => {}); };
  const drop = (p) => { p.gen++; p.task?.cancel(); p.canvas.width = p.canvas.height = 0; p.done = null; };
  const io = new IntersectionObserver((es) => es.forEach((en) => {
    const p = pages[+en.target.dataset.page]; p.near = en.isIntersecting;
    if (p.near) want(p); else drop(p);
  }), { root: pagesEl, rootMargin: '150% 0px' });
  pages.forEach((p) => io.observe(p.el));
  let renderTimer = 0;
  const rerender = () => { clearTimeout(renderTimer); renderTimer = setTimeout(() => pages.forEach((p) => p.near && want(p)), 160); };

  // ---------- marks ----------
  // Every mark: { id, page, type, ...geometry in page units (1 = 1/72 in, origin top-left), color, width }.
  // pen: pts [[x,y]…]; line/arrow: a [x,y], b [x,y]; rect: x, y, w, h; highlight: rects [[x,y,w,h]…], color, opacity;
  // text: x, y, text, runs, color, size, font.
  const uid = () => Math.random().toString(36).slice(2, 10);
  const snapshot = () => { st.undo.push(JSON.stringify(st.marks)); if (st.undo.length > 100) st.undo.shift(); st.redo = []; };
  const changed = () => { st.dirty = true; onDraft(st.marks); paint(); };
  const selMark = () => st.marks.find((x) => x.id === st.sel);
  const editingBox = () => root.querySelector('.mk-box.editing');
  // Open arrow head like ←: two strokes from the tip, 30° either side, drawn with the line's own width and round ends.
  const arrowHead = (a, b, w) => {
    const ang = Math.atan2(b[1] - a[1], b[0] - a[0]); const L = 7 + w * 2.6; const s = Math.PI / 6;
    return [b, [b[0] - L * Math.cos(ang - s), b[1] - L * Math.sin(ang - s)], [b[0] - L * Math.cos(ang + s), b[1] - L * Math.sin(ang + s)]];
  };
  function drawMark(m, g) {
    const stroke = { stroke: m.color, 'stroke-width': m.width, fill: 'none', 'stroke-linecap': 'round', 'stroke-linejoin': 'round' };
    if (m.type === 'pen') g.append(svgEl('polyline', { ...stroke, points: m.pts.map((p) => p.join(',')).join(' ') }));
    if (m.type === 'line' || m.type === 'arrow') g.append(svgEl('line', { ...stroke, x1: m.a[0], y1: m.a[1], x2: m.b[0], y2: m.b[1] }));
    if (m.type === 'arrow') { const [t, p1, p2] = arrowHead(m.a, m.b, m.width); g.append(svgEl('polyline', { ...stroke, points: [p1, t, p2].map((p) => p.join(',')).join(' ') })); }
    if (m.type === 'rect') g.append(svgEl('rect', { ...stroke, x: m.x, y: m.y, width: m.w, height: m.h }));
    if (m.type === 'highlight') m.rects.forEach(([x, y, w, h]) => g.append(svgEl('rect', { x, y, width: w, height: h, fill: m.color || HIGHLIGHT, 'fill-opacity': m.opacity ?? HL_OPACITY, class: 'mk-hl' })));
  }
  function paint() {
    pages.forEach((p, i) => {
      p.svg.replaceChildren(); p.hls.replaceChildren(); p.boxes.replaceChildren();
      st.marks.filter((m) => m.page === i).forEach((m) => {
        if (m.type === 'text') {
          const b = document.createElement('div');
          b.className = `mk-box${st.sel === m.id ? ' on' : ''}`; b.dataset.id = m.id;
          runsOf(m).forEach((r) => b.append(runNode(r)));
          b.style.cssText = `left:${m.x * p.scale}px;top:${m.y * p.scale}px;color:${m.color};font-size:${m.size * p.scale}px;font-family:${fontStack(fontOf(m))};line-height:${m.lh || TEXT.lineHeight}`
            + (m.wrap ? `;width:${m.wrap * p.scale}px;white-space:pre-wrap` : '') + (m.align ? `;text-align:${m.align}` : '');
          if (m.cover) { // an edited paragraph: the original is hidden under the page's own background colour
            const c = document.createElement('div'); c.className = 'mk-cover'; const [x, y, w, h] = m.cover.box;
            c.style.cssText = `left:${x * p.scale}px;top:${y * p.scale}px;width:${w * p.scale}px;height:${h * p.scale}px;background:${m.cover.bg}`;
            p.boxes.append(c);
          }
          p.boxes.append(b);
          if (m.cover) { // a PDF cannot move the text below down: say so when the edit needs more room than it had
            const over = b.offsetHeight > (m.cover.box[3] + m.size * 0.6) * p.scale;
            b.classList.toggle('mk-over', over); b.title = over ? 'Longer than the original: it may overlap the text below' : '';
          }
          return;
        }
        const g = svgEl('g', { 'data-id': m.id, class: `mk-mark${st.sel === m.id ? ' on' : ''}` });
        drawMark(m, g);
        // Highlights have their own layer, blended with multiply: the page takes the colour, the ink stays dark.
        (m.type === 'highlight' ? p.hls : p.svg).append(g);
      });
    });
    refreshBar();
  }

  // ---------- the tools rail ----------
  // Which controls show: text (text tool, a picked or open text box), highlight, shape, or none.
  function mode() {
    const m = selMark();
    if (st.tool === 'text' || st.tool === 'edit' || m?.type === 'text' || editingBox()) return 'text';
    if (st.tool === 'highlight' || m?.type === 'highlight') return 'highlight';
    if (m || ['pen', 'line', 'arrow', 'rect'].includes(st.tool)) return 'shape';
    return '';
  }
  // The colour box is about text while the Text box tool is on or a text box is picked, else about shapes.
  function textMode() { return st.tool === 'text' || st.tool === 'edit' || selMark()?.type === 'text'; }
  function refreshBar() {
    const m = selMark();
    mk.dataset.mode = mode();
    $('#mk-del').disabled = !st.sel;
    if (m?.type === 'text') { $('#mk-s').value = m.size; $('#mk-f').value = fontOf(m); } // the picked text box's size and font
    $('#mk-c').value = m && m.type !== 'highlight' && m.color ? m.color : textMode() ? st.textColor : st.shapeColor; syncC();
    $('#mk-hc').value = m?.type === 'highlight' ? m.color || HIGHLIGHT : st.hlColor; syncHC();
    const op = Math.round((m?.type === 'highlight' ? m.opacity ?? HL_OPACITY : st.hlOpacity) * 100);
    $('#mk-o').value = op; $('#mk-ov').textContent = `${op}%`;
    $('#mk-undo').disabled = !st.undo.length;
    $('#mk-redo').disabled = !st.redo.length;
    $('#mk-reset').disabled = !st.marks.length;
    updateFmt();
  }
  // Is style key k on every character of the selection in the open box (the caret's character when nothing is selected)?
  function selectionHas(ed, k) {
    const s = window.getSelection(); if (!s || !s.rangeCount) return false;
    const r = s.getRangeAt(0); if (!ed.contains(r.commonAncestorContainer)) return false;
    if (r.collapsed) { const n = r.startContainer; return !!styleOf(n.nodeType === 3 ? n.parentElement : n, ed)[k]; }
    const w = document.createTreeWalker(ed, NodeFilter.SHOW_TEXT); let any = false;
    for (let n = w.nextNode(); n; n = w.nextNode()) {
      if (!r.intersectsNode(n) || !n.data.length) continue;
      if ((n === r.endContainer && r.endOffset === 0) || (n === r.startContainer && r.startOffset === n.data.length)) continue;
      any = true; if (!styleOf(n.parentElement, ed)[k]) return false;
    }
    return any;
  }
  function updateFmt() {
    const ed = editingBox(); const m = !ed && selMark()?.type === 'text' ? selMark() : null;
    root.querySelectorAll('[data-fmt]').forEach((b) => {
      const k = KEY[b.dataset.fmt];
      b.setAttribute('aria-pressed', String(ed ? selectionHas(ed, k) : m ? runsOf(m).every((r) => r[k]) : false));
    });
  }
  // B / I / U / S / highlight: on the selected characters of the open box, or on the whole picked box.
  function format(cmd) {
    const ed = editingBox();
    if (ed) {
      if (cmd === 'hilite') {
        const on = selectionHas(ed, 'hl');
        exec('styleWithCSS', true); exec('hiliteColor', on ? 'transparent' : rgba(st.hlColor, st.hlOpacity)); exec('styleWithCSS', false);
      } else exec(cmd);
      updateFmt();
      return;
    }
    const m = selMark(); if (m?.type !== 'text') return;
    const k = KEY[cmd]; const runs = runsOf(m).map((r) => ({ ...r })); const all = runs.every((r) => r[k]);
    snapshot();
    runs.forEach((r) => {
      if (all) { delete r[k]; if (k === 'hl') delete r.ha; }
      else if (k === 'hl') { r.hl = st.hlColor; r.ha = st.hlOpacity; }
      else r[k] = 1;
    });
    m.runs = mergeRuns(runs); changed();
  }
  const fmtBar = $('.mk-fmt');
  // Keep the text selection and the focus in the box while a format button is pressed (mouse, touch, Pencil).
  fmtBar.addEventListener('pointerdown', (e) => e.preventDefault());
  fmtBar.addEventListener('mousedown', (e) => e.preventDefault());
  fmtBar.addEventListener('click', (e) => { const b = e.target.closest('[data-fmt]'); if (b) format(b.dataset.fmt); });
  const onSel = () => { if (!root.isConnected) { document.removeEventListener('selectionchange', onSel); return; } if (editingBox()) updateFmt(); };
  document.addEventListener('selectionchange', onSel);

  const setTool = (k) => {
    editingBox()?.blur();
    st.tool = k;
    root.querySelectorAll('button[data-tool]').forEach((b) => b.setAttribute('aria-pressed', String(b.dataset.tool === k)));
    mk.dataset.tool = k;
    if (k !== 'select') st.sel = null;
    paint();
  };
  $('.mk-tools').onclick = (e) => { const b = e.target.closest('button[data-tool]'); if (b) setTool(b.dataset.tool); };
  // A slider or colour drag on a picked mark is one undo step.
  let tweakId = null;
  const tweak = (m, f) => { if (tweakId !== m.id) { snapshot(); tweakId = m.id; } f(m); changed(); };
  ['#mk-w', '#mk-c', '#mk-hc', '#mk-o'].forEach((s) => $(s).addEventListener('change', () => { tweakId = null; }));
  $('#mk-w').oninput = (e) => {
    st.width = +e.target.value; $('#mk-wv').textContent = e.target.value;
    const m = selMark(); if (m && m.width) tweak(m, (x) => { x.width = st.width; });
  };
  // Text size (pt): the picked text box, and every new one after it.
  $('#mk-s').onchange = (e) => {
    const v = Math.min(72, Math.max(6, Math.round(+e.target.value) || textSize));
    e.target.value = v; st.size = v;
    const m = selMark();
    if (m?.type === 'text' && m.size !== v) { snapshot(); m.size = v; changed(); }
  };
  $('#mk-c').oninput = (e) => {
    const v = e.target.value;
    if (textMode()) st.textColor = v; else st.shapeColor = v;
    const m = selMark(); if (m && m.type !== 'highlight') tweak(m, (x) => { x.color = v; });
  };
  // Highlight colour and opacity: the picked highlight, and every new highlight (also inside text boxes).
  $('#mk-hc').oninput = (e) => {
    st.hlColor = e.target.value;
    const m = selMark(); if (m?.type === 'highlight') tweak(m, (x) => { x.color = st.hlColor; });
  };
  $('#mk-o').oninput = (e) => {
    st.hlOpacity = +e.target.value / 100; $('#mk-ov').textContent = `${e.target.value}%`;
    const m = selMark(); if (m?.type === 'highlight') tweak(m, (x) => { x.opacity = st.hlOpacity; });
  };
  // Font: the picked text box, and every new one after it.
  // "Add a font from this computer": .ttf/.otf files, kept in this browser only; a family's bold and italic files too.
  $('#mk-addf').onclick = () => $('#mk-ffile').click();
  $('#mk-ffile').onchange = async (e) => {
    const files = [...e.target.files]; e.target.value = ''; let last = null;
    for (const f of files) {
      try {
        const r = await addFontFile(f); last = r.family;
        if (r.site) onStatus(t(`${r.family} is already one of the site's fonts.`, `${r.family} มีอยู่ในเว็บแล้ว`));
        else { if (!fonts.includes(r.family) && !userFonts.includes(r.family)) userFonts.push(r.family); onStatus(t(`Font added: ${r.family} (kept in this browser only)`, `เพิ่มแบบอักษรแล้ว: ${r.family} (เก็บไว้ในเบราว์เซอร์นี้เท่านั้น)`)); }
      } catch (err) { onStatus(err.message); }
    }
    $('#mk-f').innerHTML = fontOptions(last || st.font);
    if (last) $('#mk-f').dispatchEvent(new Event('change'));
  };
  $('#mk-f').onchange = async (e) => {
    st.font = e.target.value;
    await loadFonts([st.font]);
    const m = selMark();
    if (m?.type === 'text' && fontOf(m) !== st.font) { snapshot(); m.font = st.font; changed(); }
  };
  const undo = () => { editingBox()?.blur(); if (!st.undo.length) return; st.redo.push(JSON.stringify(st.marks)); st.marks = JSON.parse(st.undo.pop()); st.sel = null; changed(); };
  const redo = () => { editingBox()?.blur(); if (!st.redo.length) return; st.undo.push(JSON.stringify(st.marks)); st.marks = JSON.parse(st.redo.pop()); st.sel = null; changed(); };
  const del = () => { editingBox()?.blur(); if (!st.sel) return; snapshot(); st.marks = st.marks.filter((m) => m.id !== st.sel); st.sel = null; changed(); };
  // Reset edits (from TUSUFinanceFeedback v2.16): every mark goes in one step; Undo brings them back.
  const reset = () => { editingBox()?.blur(); if (!st.marks.length) return; snapshot(); st.marks = []; st.sel = null; changed(); };
  $('#mk-undo').onclick = undo; $('#mk-redo').onclick = redo; $('#mk-del').onclick = del; $('#mk-reset').onclick = reset;

  // ---------- zoom ----------
  // z: zoom against "fit to width"; (cx, cy): the screen point that stays put (the cursor or the pinch centre).
  function zoomTo(z, cx, cy) {
    z = clampZoom(z); if (z === st.zoom) return;
    editingBox()?.blur();
    const k = z / st.zoom; st.zoom = z;
    zoomAround(pagesEl, k, cx, cy, () => { pages.forEach((p) => { p.scale = p.fit * z; size(p); }); paint(); });
    $('#mk-zfit').textContent = `${Math.round(z * 100)}%`;
    rerender();
  }
  // Fit to width again after the editor changes size (full screen, iPad rotation, window resize).
  function refit() {
    const a = avail(); const top = pagesEl.scrollTop / Math.max(1, pagesEl.scrollHeight);
    let moved = false;
    pages.forEach((p) => { const f = Math.min(1.6, a / p.base.width); if (Math.abs(f - p.fit) > 0.005) { p.fit = f; p.scale = f * st.zoom; size(p); moved = true; } });
    if (!moved) return;
    paint(); pagesEl.scrollTop = top * pagesEl.scrollHeight; rerender();
  }
  let fitTimer = 0;
  new ResizeObserver(() => { clearTimeout(fitTimer); fitTimer = setTimeout(refit, 120); }).observe(pagesEl);
  $('#mk-zin').onclick = () => zoomTo(st.zoom * 1.25);
  $('#mk-zout').onclick = () => zoomTo(st.zoom / 1.25);
  $('#mk-zfit').onclick = () => zoomTo(1);
  pinchZoom(pagesEl, () => st.zoom, zoomTo);

  // ---------- full screen (tools on top) ----------
  fullScreen(mk, 'mk', 'mk-full', t);

  const onKey = (e) => {
    if (!root.isConnected) { document.removeEventListener('keydown', onKey); return; }
    if (e.target.isContentEditable || /INPUT|TEXTAREA|SELECT/.test(e.target.tagName)) return;
    const mod = e.ctrlKey || e.metaKey;
    if (mod && e.key.toLowerCase() === 'z') { e.preventDefault(); e.shiftKey ? redo() : undo(); }
    else if (mod && e.key.toLowerCase() === 'y') { e.preventDefault(); redo(); }
    else if (mod && (e.key === '=' || e.key === '+')) { e.preventDefault(); zoomTo(st.zoom * 1.25); }
    else if (mod && e.key === '-') { e.preventDefault(); zoomTo(st.zoom / 1.25); }
    else if (mod && e.key === '0') { e.preventDefault(); zoomTo(1); }
    else if (mod && st.sel && selMark()?.type === 'text' && ['b', 'i', 'u'].includes(e.key.toLowerCase())) {
      e.preventDefault(); format({ b: 'bold', i: 'italic', u: 'underline' }[e.key.toLowerCase()]);
    }
    else if (e.key === 'Delete' || e.key === 'Backspace') { if (st.sel) { e.preventDefault(); del(); } }
    else if (e.key === 'Escape') { st.sel = null; paint(); }
  };
  document.addEventListener('keydown', onKey);

  // Pointer position on a page, in page units.
  const at = (p, e) => { const r = p.el.getBoundingClientRect(); return [(e.clientX - r.left) / p.scale, (e.clientY - r.top) / p.scale]; };
  // Shift: lines snap to 45° steps; rectangles become squares.
  const snap = (a, b, shift) => {
    if (!shift) return b;
    const dx = b[0] - a[0], dy = b[1] - a[1]; const ang = Math.round(Math.atan2(dy, dx) / (Math.PI / 4)) * (Math.PI / 4); const len = Math.hypot(dx, dy);
    return [a[0] + len * Math.cos(ang), a[1] + len * Math.sin(ang)];
  };

  let drag = null;
  // A second finger during a finger stroke or move is a pinch: undo what the first finger started.
  const cancelDrag = () => { const d = drag; drag = null; if (!d) return; st.marks = JSON.parse(st.undo.pop()); paint(); };
  // Apple Pencil (from TUSUFinanceFeedback): once a pen has been used, finger touches no longer draw (a resting palm
  // would); switch to Select to scroll by finger. A touch during a pen stroke is the palm: it is ignored.
  let penSeen = false;
  const penOnly = () => penSeen && ['pen', 'line', 'arrow', 'rect'].includes(st.tool);
  pagesEl.addEventListener('pointerdown', (e) => {
    if (e.pointerType === 'pen') penSeen = true;
    if (drag) { if (e.pointerType === 'touch' && drag.ptype !== 'pen') cancelDrag(); return; }
    if (e.pointerType === 'touch' && penOnly()) return;
    if (e.button !== 0) return;
    // Inside the open text box the browser places the caret and selects text. A press anywhere else first saves it.
    const ed = editingBox();
    if (ed) { if (ed.contains(e.target)) return; ed.blur(); }
    tweakId = null;
    const tgt = ed ? document.elementFromPoint(e.clientX, e.clientY) : e.target;
    const pel = tgt?.closest('.mk-page'); if (!pel) return;
    const p = pages[+pel.dataset.page]; const pt = at(p, e);
    if (st.tool === 'highlight') return; // the text layer selects; pointerup turns the selection into a highlight
    if (st.tool === 'edit') { // click a paragraph of the PDF to edit it in place
      const hit = tgt.closest('.mk-box');
      if (hit) { st.sel = hit.dataset.id; paint(); editBox(p, hit.dataset.id, false, [e.clientX, e.clientY]); e.preventDefault(); return; }
      e.preventDefault(); editAt(p, pt).catch((err) => console.error(err));
      return;
    }
    if (st.tool === 'text') {
      const hit = tgt.closest('.mk-box');
      if (hit) { st.sel = hit.dataset.id; paint(); editBox(p, hit.dataset.id, false, [e.clientX, e.clientY]); e.preventDefault(); return; }
      snapshot();
      const m = { id: uid(), page: +pel.dataset.page, type: 'text', x: pt[0], y: pt[1] - st.size * 0.7, text: '', runs: [{ t: '' }], color: st.textColor, size: st.size, font: st.font };
      st.marks.push(m); st.sel = m.id; paint(); editBox(p, m.id, true); e.preventDefault();
      return;
    }
    if (st.tool === 'select') {
      const id = tgt.closest('[data-id]')?.dataset.id || null; st.sel = id; paint();
      if (!id) return;
      const m = selMark();
      if (m.type === 'text' && e.detail === 2) { editBox(p, id, false, [e.clientX, e.clientY]); e.preventDefault(); return; }
      snapshot(); drag = { kind: 'move', p, last: pt, m, moved: false, id: e.pointerId, ptype: e.pointerType };
      pel.setPointerCapture(e.pointerId); e.preventDefault();
      return;
    }
    snapshot();
    const base = { id: uid(), page: +pel.dataset.page, color: st.shapeColor, width: st.width };
    const m = st.tool === 'pen' ? { ...base, type: 'pen', pts: [pt] } : st.tool === 'rect' ? { ...base, type: 'rect', x: pt[0], y: pt[1], w: 0, h: 0 } : { ...base, type: st.tool, a: pt, b: pt };
    st.marks.push(m); drag = { kind: 'draw', p, start: pt, m, id: e.pointerId, ptype: e.pointerType };
    pel.setPointerCapture(e.pointerId); e.preventDefault(); paint();
  });
  // Double-click opens a text box with the Select tool where pointerdown carries no click count.
  pagesEl.addEventListener('dblclick', (e) => {
    if (st.tool !== 'select' || editingBox()) return;
    const hit = e.target.closest('.mk-box'); const pel = e.target.closest('.mk-page'); if (!hit || !pel) return;
    editBox(pages[+pel.dataset.page], hit.dataset.id, false, [e.clientX, e.clientY]);
  });
  pagesEl.addEventListener('pointermove', (e) => {
    if (!drag || e.pointerId !== drag.id) return;
    const pt = at(drag.p, e); const m = drag.m;
    if (drag.kind === 'move') {
      const dx = pt[0] - drag.last[0], dy = pt[1] - drag.last[1]; drag.last = pt; drag.moved = true;
      if (m.pts) m.pts = m.pts.map(([x, y]) => [x + dx, y + dy]);
      if (m.a) { m.a = [m.a[0] + dx, m.a[1] + dy]; m.b = [m.b[0] + dx, m.b[1] + dy]; }
      if (m.rects) m.rects = m.rects.map(([x, y, w, h]) => [x + dx, y + dy, w, h]);
      if ('x' in m) { m.x += dx; m.y += dy; }
    } else if (m.type === 'pen') {
      // Shift while drawing: a straight stroke from where it started.
      if (e.shiftKey) m.pts = [drag.start, snap(drag.start, pt, true)]; else m.pts.push(pt);
    } else if (m.type === 'rect') {
      let w = pt[0] - drag.start[0], h = pt[1] - drag.start[1];
      if (e.shiftKey) { const s = Math.max(Math.abs(w), Math.abs(h)); w = Math.sign(w || 1) * s; h = Math.sign(h || 1) * s; }
      m.x = Math.min(drag.start[0], drag.start[0] + w); m.y = Math.min(drag.start[1], drag.start[1] + h); m.w = Math.abs(w); m.h = Math.abs(h);
    } else m.b = snap(drag.start, pt, e.shiftKey);
    paint();
  });
  const endDrag = (e) => {
    if (!drag || e.pointerId !== drag.id) return;
    const { m, kind, moved } = drag; drag = null;
    if (kind === 'move' && !moved) { st.undo.pop(); return; } // a click that only selected
    const tiny = (m.type === 'rect' && m.w < 2 && m.h < 2) || ((m.type === 'line' || m.type === 'arrow') && Math.hypot(m.b[0] - m.a[0], m.b[1] - m.a[1]) < 2) || (m.type === 'pen' && m.pts.length < 2);
    if (kind === 'draw' && tiny) { st.marks = st.marks.filter((x) => x !== m); st.undo.pop(); paint(); return; }
    changed();
  };
  pagesEl.addEventListener('pointerup', endDrag);
  pagesEl.addEventListener('pointercancel', (e) => { if (drag && e.pointerId === drag.id) cancelDrag(); });

  // Highlight: the text selected with the Highlight tool gets the highlight colour. Over the PDF's own text that is a
  // highlight mark (boxes over those words); inside a text box it is those characters. A mouse highlights on release;
  // touch and Apple Pencil (iPad, v2.13) select with the handles first, then press "Highlight selection"
  // (Safari sends no mouseup after a long-press selection).
  pagesEl.addEventListener('mouseup', () => { if (st.tool === 'highlight') highlightSelection(); });
  $('#mk-hl').onclick = highlightSelection;
  $('#mk-hl').onpointerdown = (e) => e.preventDefault(); // keep the text selection
  $('#mk-hl').onmousedown = (e) => e.preventDefault();
  function highlightSelection() {
    const sel = window.getSelection(); if (!sel || sel.isCollapsed || !sel.rangeCount) return;
    const r0 = sel.getRangeAt(0); const ca = r0.commonAncestorContainer;
    const boxEl = (ca.nodeType === 3 ? ca.parentElement : ca).closest?.('.mk-box');
    if (boxEl) {
      const m = st.marks.find((x) => x.id === boxEl.dataset.id);
      const pre = document.createRange(); pre.selectNodeContents(boxEl); pre.setEnd(r0.startContainer, r0.startOffset);
      const a = pre.toString().length; const z = a + r0.toString().length;
      sel.removeAllRanges();
      if (m && z > a) { snapshot(); m.runs = styleRange(runsOf(m), a, z, { hl: st.hlColor, ha: st.hlOpacity }); changed(); }
      return;
    }
    const byPage = new Map();
    for (let i = 0; i < sel.rangeCount; i++) {
      [...sel.getRangeAt(i).getClientRects()].forEach((r) => {
        if (r.width < 1 || r.height < 1) return;
        const p = pages.find((x) => { const b = x.el.getBoundingClientRect(); return r.top >= b.top - 2 && r.bottom <= b.bottom + 2; });
        if (!p) return;
        const b = p.el.getBoundingClientRect();
        const rect = [(r.left - b.left) / p.scale, (r.top - b.top) / p.scale, r.width / p.scale, r.height / p.scale];
        const k = +p.el.dataset.page; byPage.set(k, [...(byPage.get(k) || []), rect]);
      });
    }
    sel.removeAllRanges();
    if (!byPage.size) return;
    snapshot();
    byPage.forEach((rects, page) => st.marks.push({ id: uid(), page, type: 'highlight', rects: mergeRects(rects), color: st.hlColor, opacity: st.hlOpacity }));
    changed();
  }

  // Replaces "-->" / "==>" before the caret with an arrow; Ctrl+Z gives the typed characters back.
  function arrowize() {
    const s = window.getSelection(); if (!s.rangeCount || !s.isCollapsed) return;
    const { startContainer: n, startOffset: o } = s.getRangeAt(0);
    if (n.nodeType !== 3) return;
    const len = o >= 3 && ARROWS[n.data.slice(o - 3, o)] ? 3 : o >= 2 && ARROWS[n.data.slice(o - 2, o)] ? 2 : 0;
    if (!len) return;
    const r = document.createRange(); r.setStart(n, o - len); r.setEnd(n, o);
    s.removeAllRanges(); s.addRange(r); exec('insertText', ARROWS[n.data.slice(o - len, o)]);
  }

  // Text box editing: rich text per character; Enter = new line; the caret goes where the box was pressed (at the end
  // for a new box); an empty box is removed when it loses focus.
  function editBox(p, id, fresh = false, pt = null) {
    const m = st.marks.find((x) => x.id === id); if (!m) return;
    const b = p.boxes.querySelector(`[data-id="${id}"]`); if (!b || b.classList.contains('editing')) return;
    if (!fresh) snapshot();
    b.contentEditable = 'true'; b.spellcheck = false;
    b.classList.add('editing'); exec('styleWithCSS', false);
    b.focus({ preventScroll: true });
    let r = pt && caretAt(pt[0], pt[1]);
    if (!r || !b.contains(r.startContainer)) { r = document.createRange(); r.selectNodeContents(b); r.collapse(false); }
    const s = window.getSelection(); s.removeAllRanges(); s.addRange(r);
    mk.dataset.mode = 'text'; updateFmt();
    b.onkeydown = (e) => {
      const mod = e.ctrlKey || e.metaKey;
      if (e.key === 'Enter' && !e.isComposing) { e.preventDefault(); if (!exec('insertLineBreak')) exec('insertHTML', '<br>'); }
      else if (mod && e.shiftKey && e.key.toLowerCase() === 'x') { e.preventDefault(); exec('strikeThrough'); updateFmt(); }
      else if (e.key === 'Escape') { e.preventDefault(); b.blur(); }
    };
    b.onpaste = (e) => { e.preventDefault(); exec('insertText', e.clipboardData?.getData('text/plain') || ''); };
    b.oninput = (e) => { if (e.inputType === 'insertText' && e.data === '>') arrowize(); updateFmt(); };
    b.onblur = () => {
      b.onblur = null; b.classList.remove('editing'); b.contentEditable = 'false';
      m.runs = readRuns(b); m.text = m.runs.map((x) => x.t).join('');
      if (!m.text.trim() && !m.cover) { st.marks = st.marks.filter((x) => x.id !== id); if (st.sel === id) st.sel = null; } // an emptied edit deletes the paragraph
      changed();
    };
  }

  // ---------- Edit text ----------
  // The paragraph under point pt (page units) becomes a text box in its place: its text, the nearest site font, its
  // size, line spacing and colour; the original is covered on screen and removed from the downloaded PDF.
  async function editAt(p, pt) {
    const paras = await (p.paras ||= pageText(p));
    const pad = 2; const u = paras.findLast(({ box: [x, y, w, h] }) => pt[0] >= x - pad && pt[0] <= x + w + pad && pt[1] >= y - pad && pt[1] <= y + h + pad);
    if (!u) return;
    const n = pages.indexOf(p); const key = `${n}:${u.box.map((v) => v.toFixed(1)).join(',')}`;
    const had = st.marks.find((m) => m.cover?.key === key);
    if (had) { st.sel = had.id; paint(); editBox(p, had.id); return; }
    const font = u.fam ? siteFont(u.fam) : 'Sarabun'; await loadFonts([font]); // OCR text has no font name
    const { ink, bg } = sampleColours(p, u.box);
    const c = document.createElement('canvas').getContext('2d'); c.font = `${u.size}px ${fontStack(font)}`;
    const mt = c.measureText('Hg'); const asc = mt.fontBoundingBoxAscent ?? u.size * 0.8; const desc = mt.fontBoundingBoxDescent ?? u.size * 0.25;
    const lh = u.lines > 1 ? u.lh : TEXT.lineHeight;
    snapshot();
    const m = {
      id: uid(), page: n, type: 'text', x: u.align ? u.col[0] : u.box[0] + 1, y: u.first - ((lh * u.size - asc - desc) / 2 + asc), align: u.align,
      runs: u.runs.map((r) => ({ t: r.t, ...(r.b ? { b: 1 } : {}), ...(r.i ? { i: 1 } : {}) })), text: u.text, color: ink, size: u.size, font, lh,
      // wrap width: a paragraph keeps its own width; one line may run to the page's text margin; centred text
      // stays centred in the page's text column; right-aligned text keeps its own right edge
      wrap: u.align === 'right' ? u.box[0] + u.box[2] - 1 - u.col[0] : u.align ? u.col[1] - u.col[0] : u.lines > 1 ? u.box[2] + u.size * 0.3 : Math.max(u.box[2], u.col[1] - u.box[0]) + u.size * 0.3,
      cover: { key, box: u.box, bg, scan: !!p.scanned, rects: u.rects.map(([x0, y0, x1, y1]) => [x0 - u.size * 0.8, y0, x1, y1]) },
    };
    st.marks.push(m); st.sel = m.id; changed(); editBox(p, m.id, true);
  }
  // The page's paragraphs; a page without text (a scan), or whose Thai text the PDF holds damaged, is read by OCR
  // first, once.
  async function pageText(p) {
    const own = await pageParagraphs(p.page);
    const text = own.map((u) => u.text).join(' ');
    const damaged = text.trim().length >= 10 && damagedThai(text, { canva: await canva });
    if ((text.trim().length >= 10 && !damaged) || !ocr) return own;
    try {
      onStatus(damaged ? t('The Thai text in this page is damaged in the PDF: reading it with OCR…', 'ข้อความภาษาไทยในหน้านี้เสียหาย กำลังอ่านด้วย OCR…')
        : t('Reading this scanned page with OCR…', 'กำลังอ่านหน้าสแกนด้วย OCR…'));
      const { ocrPage } = await import('./pdf-ocr.js');
      const items = await ocrPage(p.page, onStatus); p.scanned = !damaged; // a damaged page still has text to remove
      const read = await pageParagraphs(p.page, items);
      const chk = (items.check || []).slice(0, 6).map(([w, s]) => (s ? `${w} (${s}?)` : w)).join(', ');
      const more = chk ? t(` Words to check: ${chk}${items.check.length > 6 ? '…' : ''}.`, ` คำที่ควรตรวจ: ${chk}${items.check.length > 6 ? '…' : ''}`) : '';
      onStatus((damaged ? t(`Page read by OCR: ${read.length} paragraphs to edit. Check names and figures.`, `อ่านหน้านี้ด้วย OCR แล้ว ${read.length} ย่อหน้า กรุณาตรวจชื่อและตัวเลข`)
        : t(`Scanned page read by OCR: ${read.length} paragraphs to edit. Check names and figures.`, `อ่านหน้าสแกนแล้ว ${read.length} ย่อหน้า กรุณาตรวจชื่อและตัวเลข`)) + more);
      return read;
    } catch (err) { p.paras = null; onStatus(err.message); return own; }
  }
  // The text and background colours under a box, from the drawn page: the darkest and the most common light pixels.
  function sampleColours(p, [x, y, w, h]) {
    const cv = p.canvas; const k = cv.width / p.w; const fall = { ink: '#000000', bg: '#ffffff' };
    if (!cv.width) return fall;
    try {
      const d = cv.getContext('2d').getImageData(Math.max(0, x * k), Math.max(0, y * k), Math.max(1, w * k), Math.max(1, h * k)).data;
      const px = []; for (let i = 0; i < d.length; i += 16) px.push([d[i], d[i + 1], d[i + 2], 0.3 * d[i] + 0.59 * d[i + 1] + 0.11 * d[i + 2]]);
      px.sort((a, b) => a[3] - b[3]);
      const avg = (list) => `#${[0, 1, 2].map((j) => Math.round(list.reduce((s, q) => s + q[j], 0) / list.length).toString(16).padStart(2, '0')).join('')}`;
      return { ink: avg(px.slice(0, Math.max(1, px.length * 0.04))), bg: avg(px.slice(Math.floor(px.length * 0.6))) };
    } catch { return fall; }
  }
  // An edited paragraph's runs with its line breaks where the screen wraps them (the PDF has no wrapping).
  function wrapRuns(m) {
    const runs = runsOf(m); const pv = pages[m.page]; const box = pv?.boxes.querySelector(`[data-id="${m.id}"]`);
    if (!box) return runs;
    const breaks = new Set(); let off = 0; let top = null; const r = document.createRange();
    const walk = document.createTreeWalker(box, NodeFilter.SHOW_TEXT);
    for (let t = walk.nextNode(); t; t = walk.nextNode()) {
      for (const { segment, index } of (seg ? [...seg.segment(t.data)] : [...t.data].map((s, i) => ({ segment: s, index: i })))) {
        if (segment === '\n') { top = null; off += 1; continue; }
        r.setStart(t, index); r.setEnd(t, index + segment.length); const rect = r.getClientRects()[0];
        if (rect) { if (top != null && rect.top > top + rect.height / 2) breaks.add(off); top = rect.top; }
        off += segment.length;
      }
    }
    const out = []; off = 0;
    for (const run of runs) {
      let t = '';
      for (const ch of run.t) { if (breaks.has(off)) t = `${t.replace(/ $/, '')}\n`; t += ch; off++; }
      out.push({ ...run, t });
    }
    return out;
  }

  setTool('select');
  paint();

  // ---------- writing the marked PDF ----------
  async function exportPdf() {
    document.activeElement?.blur?.();
    await loadFonts(st.marks.filter((m) => m.type === 'text').map(fontOf));
    const out = await PDFLib.PDFDocument.load(bytes.slice(0));
    const { rgb, LineCapStyle } = PDFLib;
    const hex = (c) => { const [r, g, b] = rgbOf(c); return rgb(r / 255, g / 255, b / 255); };
    const outPages = out.getPages();
    const P0 = (n, [x, y]) => { const pg = outPages[n]; const { x: ox, y: oy } = pg.getMediaBox ? pg.getMediaBox() : { x: 0, y: 0 }; return { x: ox + x, y: oy + pages[n].h - y }; };
    // Text boxes as real text (searchable, copyable); a box whose font cannot be embedded is drawn as a picture.
    const writer = st.marks.some((m) => m.type === 'text') ? await textWriter(out, PDFLib) : null;
    // Edited paragraphs: their old text is removed from the page (search and copy no longer find it); where it cannot
    // be removed (text inside a form XObject), it is covered with the page's background colour.
    const edits = st.marks.filter((m) => m.type === 'text' && m.cover); let covered = 0;
    const { removeText } = edits.length ? await import('./pdf-redact.js') : {};
    for (const n of [...new Set(edits.map((m) => m.page))]) {
      const mine = edits.filter((m) => m.page === n); const areas = mine.flatMap((m) => m.cover.rects);
      const hits = removeText(out, PDFLib, n, areas); let k = 0;
      for (const m of mine) {
        const got = m.cover.rects.map(() => hits[k++]);
        // a scan has no text to remove: its picture is covered (no note); otherwise covering is the fallback
        if (got.some((h) => !h)) { if (!m.cover.scan) covered++; const [x, y, w, h] = m.cover.box; const a = P0(n, [x, y + h]); outPages[n].drawRectangle({ x: a.x, y: a.y, width: w, height: h, color: hex(m.cover.bg) }); }
      }
    }
    // Highlights first, so every other mark (and the page's ink, through multiply) sits on top of them.
    const order = [...st.marks.filter((m) => m.type === 'highlight'), ...st.marks.filter((m) => m.type !== 'highlight')];
    for (const m of order) {
      const pg = outPages[m.page]; const pv = pages[m.page]; if (!pg || !pv) continue;
      const { x: ox, y: oy } = pg.getMediaBox ? pg.getMediaBox() : { x: 0, y: 0 };
      // Page units (top-left origin) -> PDF user space (bottom-left origin). Pages with /Rotate are not handled yet.
      const P = ([x, y]) => ({ x: ox + x, y: oy + pv.h - y });
      const line = (a, b, w, c) => pg.drawLine({ start: P(a), end: P(b), thickness: w, color: hex(c), lineCap: LineCapStyle.Round });
      if (m.type === 'pen') for (let i = 1; i < m.pts.length; i++) line(m.pts[i - 1], m.pts[i], m.width, m.color);
      if (m.type === 'line' || m.type === 'arrow') line(m.a, m.b, m.width, m.color);
      if (m.type === 'arrow') { const [t, p1, p2] = arrowHead(m.a, m.b, m.width); line(p1, t, m.width, m.color); line(p2, t, m.width, m.color); }
      if (m.type === 'rect') {
        const c = [[m.x, m.y], [m.x + m.w, m.y], [m.x + m.w, m.y + m.h], [m.x, m.y + m.h]];
        for (let i = 0; i < 4; i++) line(c[i], c[(i + 1) % 4], m.width, m.color);
      }
      if (m.type === 'highlight') {
        m.rects.forEach(([x, y, w, h]) => {
          const a = P([x, y + h]);
          pg.drawRectangle({ x: a.x, y: a.y, width: w, height: h, color: hex(m.color || HIGHLIGHT), opacity: m.opacity ?? HL_OPACITY, blendMode: PDFLib.BlendMode?.Multiply });
        });
      }
      if (m.type === 'text' && m.text.trim() && await writer.draw(pg, { ...m, runs: m.wrap ? wrapRuns(m) : runsOf(m) }, (x, y) => P([x, y]), { lineHeight: m.lh || TEXT.lineHeight, fontFamily: fontOf })) continue;
      if (m.type === 'text' && m.text.trim()) {
        const png = await textPng({ ...m, runs: m.wrap ? wrapRuns(m) : runsOf(m) });
        const img = await out.embedPng(png.bytes);
        const a = P([m.x, m.y - png.oy + png.h]);
        pg.drawImage(img, { x: a.x, y: a.y, width: png.w, height: png.h });
      }
    }
    exportNotes = writer ? [...writer.notes] : [];
    if (covered) exportNotes.push(`${covered} edited paragraph${covered > 1 ? 's' : ''} could not be removed from the page (the PDF draws ${covered > 1 ? 'them' : 'it'} in a way this tool cannot change), so ${covered > 1 ? 'they were' : 'it was'} covered instead`);
    return out.save();
  }
  let exportNotes = [];

  return { marks: () => structuredClone(st.marks), isDirty: () => st.dirty, clean: () => { st.dirty = false; }, exportPdf, notes: () => exportNotes, name, doc };
}

// A text box drawn by the browser at 4x (so Thai vowels and tone marks are placed correctly), as PNG bytes, its size
// in page units and oy, how far above the box's top the image starts. Lines are laid out as CSS does: each line is
// line-height tall, with the font's ascent + descent centred in it. Highlights are filled first, under all letters.
async function textPng(m) {
  const k = 4; const px = m.size * k; const lh = m.size * TEXT.lineHeight * k; const fam = fontStack(fontOf(m));
  const fontFor = (r) => `${r.i ? 'italic ' : ''}${r.b ? 'bold ' : ''}${px}px ${fam}`;
  const lines = [[]];
  runsOf(m).forEach((r) => r.t.split('\n').forEach((piece, i) => { if (i) lines.push([]); if (piece) lines.at(-1).push({ ...r, t: piece }); }));
  const c = document.createElement('canvas'); const ctx = c.getContext('2d');
  ctx.font = fontFor({}); const mt = ctx.measureText('Hg');
  const asc = mt.fontBoundingBoxAscent ?? px * 0.8; const desc = mt.fontBoundingBoxDescent ?? px * 0.25;
  const lay = lines.map((l) => { let x = 0; return l.map((r) => { ctx.font = fontFor(r); const w = ctx.measureText(r.t).width; const o = { r, x, w }; x += w; return o; }); });
  const pad = Math.max(0, (asc + desc - lh) / 2); // tall Thai fonts reach above and below the line box
  let W = Math.max(1, ...lay.map((l) => (l.length ? l.at(-1).x + l.at(-1).w : 0))) + px * 0.15 + 2 * k;
  if ((m.align === 'center' || m.align === 'right') && m.wrap) { // place each line in the box's width, as on screen
    W = Math.max(W, m.wrap * k);
    lay.forEach((l) => { const lw = l.length ? l.at(-1).x + l.at(-1).w : 0; const dx = (W - lw) / (m.align === 'center' ? 2 : 1); l.forEach((o) => { o.x += dx; }); });
  }
  c.width = Math.ceil(W); c.height = Math.ceil(lh * lines.length + 2 * pad);
  const top = (i) => pad + i * lh + (lh - asc - desc) / 2; // top of the font box on line i
  lay.forEach((l, i) => l.forEach(({ r, x, w }) => { if (r.hl) { ctx.fillStyle = rgba(r.hl, r.ha ?? HL_OPACITY); ctx.fillRect(x, top(i), w, asc + desc); } }));
  ctx.textBaseline = 'alphabetic'; ctx.fillStyle = m.color; const th = Math.max(k, px / 14);
  lay.forEach((l, i) => l.forEach(({ r, x, w }) => {
    const base = top(i) + asc; ctx.font = fontFor(r); ctx.fillText(r.t, x, base);
    if (r.u) ctx.fillRect(x, base + px * 0.1, w, th);
    if (r.s) ctx.fillRect(x, base - px * 0.3, w, th);
  }));
  const blob = await new Promise((res) => c.toBlob(res, 'image/png'));
  return { bytes: new Uint8Array(await blob.arrayBuffer()), w: c.width / k, h: c.height / k, oy: pad / k };
}

// Selection rectangles of one line come in pieces: join neighbours on the same line.
function mergeRects(rs) {
  const out = [];
  rs.sort((a, b) => a[1] - b[1] || a[0] - b[0]).forEach((r) => {
    const last = out.at(-1);
    if (last && Math.abs(last[1] - r[1]) < 2 && r[0] <= last[0] + last[2] + 3) {
      const x2 = Math.max(last[0] + last[2], r[0] + r[2]); last[2] = x2 - last[0]; last[3] = Math.max(last[3], r[3]);
    } else out.push([...r]);
  });
  return out;
}
