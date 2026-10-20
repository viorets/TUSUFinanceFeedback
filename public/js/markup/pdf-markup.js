// PDF markup editor (v2.12 test run): text box, highlight text, freehand draw, line, arrow, rectangle; adjustable
// width; Shift keeps lines at 45° steps and rectangles square. Marks are kept as data (page + points in PDF units),
// so a draft is a small JSON list. "Download" writes the marks into a copy of the PDF with pdf-lib; the original
// bytes are never changed. Text boxes are drawn by the browser (correct Thai shaping) and embedded as PNG.
// Libraries load only when the editor opens: pdf.js (pages + text layer) and pdf-lib (writing).
const PDFJS = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.4.168/pdf.min.mjs';
const PDFJS_WORKER = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.4.168/pdf.worker.min.mjs';
const PDFLIB = 'https://cdnjs.cloudflare.com/ajax/libs/pdf-lib/1.17.1/pdf-lib.min.js';
export const MAX_BYTES = 12 * 1024 * 1024; // marked files: 12 MB (decided 20 Oct)
// Text box default (decided 20 Oct): LINE Seed Sans TH, 11 pt, #990000, clear background, no border.
export const TEXT = { font: 'LINE Seed Sans TH', size: 11, color: '#990000', lineHeight: 1.35 };
const HIGHLIGHT = '#ffe14d';
const TOOLS = [
  ['select', 'Select', 'เลือก'], ['text', 'Text box', 'กล่องข้อความ'], ['highlight', 'Highlight', 'ไฮไลต์'],
  ['pen', 'Draw', 'วาด'], ['line', 'Line', 'เส้น'], ['arrow', 'Arrow', 'ลูกศร'], ['rect', 'Rectangle', 'สี่เหลี่ยม'],
];

let libs = null;
function loadLibs() {
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

// Opens the editor in `root`. bytes: the PDF (ArrayBuffer); name: file name; draft: saved marks or null.
// t(en, th) translates labels. Resolves an editor handle { marks(), isDirty(), exportPdf() }.
export async function openPdfEditor(root, { bytes, name, draft = null, t = (en) => en, onDraft = () => {} }) {
  if (bytes.byteLength > MAX_BYTES) throw new Error(t('This file is over 12 MB, so it cannot be marked. Open it in Drive instead.', 'ไฟล์ใหญ่เกิน 12 MB จึงทำเครื่องหมายไม่ได้ กรุณาเปิดใน Drive'));
  root.innerHTML = `<div class="mk"><div class="mk-bar" role="toolbar">
      <div class="mk-tools">${TOOLS.map(([k, en, th]) => `<button type="button" data-tool="${k}" aria-pressed="false">${t(en, th)}</button>`).join('')}</div>
      <label class="mk-ctl">${t('Width', 'ความหนา')} <input type="range" id="mk-w" min="1" max="12" step="0.5" value="2"><output id="mk-wv">2</output></label>
      <label class="mk-ctl">${t('Colour', 'สี')} <input type="color" id="mk-c" value="${TEXT.color}"></label>
      <span class="mk-grow"></span>
      <button type="button" id="mk-undo" title="Ctrl+Z">${t('Undo', 'เลิกทำ')}</button><button type="button" id="mk-redo" title="Ctrl+Y">${t('Redo', 'ทำซ้ำ')}</button>
      <button type="button" id="mk-del" disabled>${t('Delete', 'ลบ')}</button>
    </div><div class="mk-pages" id="mk-pages"><p class="mk-wait">${t('Loading…', 'กำลังโหลด…')}</p></div></div>`;
  const [pdfjs, PDFLib] = await loadLibs();
  const doc = await pdfjs.getDocument({ data: bytes.slice(0) }).promise;
  const pagesEl = root.querySelector('#mk-pages');
  pagesEl.innerHTML = '';

  const st = { tool: 'select', width: 2, color: TEXT.color, marks: draft ? structuredClone(draft) : [], undo: [], redo: [], sel: null, dirty: false };
  const pages = []; // { el, svg, boxes, vp, scale, w, h }
  const avail = Math.max(320, pagesEl.clientWidth - 32);

  for (let n = 1; n <= doc.numPages; n++) {
    const page = await doc.getPage(n);
    const base = page.getViewport({ scale: 1 });
    const scale = Math.min(1.6, avail / base.width);
    const vp = page.getViewport({ scale });
    const el = document.createElement('div');
    el.className = 'mk-page'; el.dataset.page = n - 1;
    el.style.width = `${vp.width}px`; el.style.height = `${vp.height}px`;
    el.style.setProperty('--scale-factor', scale);
    const canvas = document.createElement('canvas');
    const dpr = window.devicePixelRatio || 1;
    canvas.width = Math.floor(vp.width * dpr); canvas.height = Math.floor(vp.height * dpr);
    const text = document.createElement('div'); text.className = 'textLayer';
    const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg'); svg.classList.add('mk-svg');
    svg.setAttribute('viewBox', `0 0 ${base.width} ${base.height}`);
    const boxes = document.createElement('div'); boxes.className = 'mk-boxes';
    el.append(canvas, text, svg, boxes);
    pagesEl.append(el);
    pages.push({ el, svg, boxes, text, base, scale, w: base.width, h: base.height });
    await page.render({ canvasContext: canvas.getContext('2d'), viewport: vp, transform: dpr !== 1 ? [dpr, 0, 0, dpr, 0, 0] : null }).promise;
    try { await new pdfjs.TextLayer({ textContentSource: page.streamTextContent(), container: text, viewport: vp }).render(); } catch { /* no text layer: highlight is unavailable here */ }
  }

  // ---------- marks ----------
  // Every mark: { id, page, type, ...geometry in page units (1 = 1/72 in, origin top-left), color, width }.
  // pen: pts [[x,y]…]; line/arrow: a [x,y], b [x,y]; rect: x, y, w, h; highlight: rects [[x,y,w,h]…];
  // text: x, y, text, color, size.
  const uid = () => Math.random().toString(36).slice(2, 10);
  const snapshot = () => { st.undo.push(JSON.stringify(st.marks)); if (st.undo.length > 100) st.undo.shift(); st.redo = []; };
  const changed = () => { st.dirty = true; onDraft(st.marks); paint(); };
  const NS = 'http://www.w3.org/2000/svg';
  const svgEl = (tag, attrs) => { const e = document.createElementNS(NS, tag); Object.entries(attrs).forEach(([k, v]) => e.setAttribute(k, v)); return e; };
  const arrowHead = (a, b, w) => {
    const ang = Math.atan2(b[1] - a[1], b[0] - a[0]); const L = 6 + w * 3; const s = Math.PI / 7;
    return [b, [b[0] - L * Math.cos(ang - s), b[1] - L * Math.sin(ang - s)], [b[0] - L * Math.cos(ang + s), b[1] - L * Math.sin(ang + s)]];
  };
  function drawMark(m, g) {
    const stroke = { stroke: m.color, 'stroke-width': m.width, fill: 'none', 'stroke-linecap': 'round', 'stroke-linejoin': 'round' };
    if (m.type === 'pen') g.append(svgEl('polyline', { ...stroke, points: m.pts.map((p) => p.join(',')).join(' ') }));
    if (m.type === 'line' || m.type === 'arrow') g.append(svgEl('line', { ...stroke, x1: m.a[0], y1: m.a[1], x2: m.b[0], y2: m.b[1] }));
    if (m.type === 'arrow') g.append(svgEl('polygon', { fill: m.color, points: arrowHead(m.a, m.b, m.width).map((p) => p.join(',')).join(' ') }));
    if (m.type === 'rect') g.append(svgEl('rect', { ...stroke, x: m.x, y: m.y, width: m.w, height: m.h }));
    if (m.type === 'highlight') m.rects.forEach(([x, y, w, h]) => g.append(svgEl('rect', { x, y, width: w, height: h, fill: HIGHLIGHT, 'fill-opacity': 0.45, class: 'mk-hl' })));
  }
  function paint() {
    pages.forEach((p, i) => {
      p.svg.replaceChildren(); p.boxes.replaceChildren();
      st.marks.filter((m) => m.page === i).forEach((m) => {
        if (m.type === 'text') {
          const b = document.createElement('div');
          b.className = `mk-box${st.sel === m.id ? ' on' : ''}`; b.dataset.id = m.id; b.textContent = m.text;
          b.style.cssText = `left:${m.x * p.scale}px;top:${m.y * p.scale}px;color:${m.color};font-size:${m.size * p.scale}px;font-family:"${TEXT.font}";line-height:${TEXT.lineHeight}`;
          p.boxes.append(b);
          return;
        }
        const g = svgEl('g', { 'data-id': m.id, class: `mk-mark${st.sel === m.id ? ' on' : ''}` });
        drawMark(m, g);
        p.svg.append(g);
      });
    });
    root.querySelector('#mk-del').disabled = !st.sel;
    root.querySelector('#mk-undo').disabled = !st.undo.length;
    root.querySelector('#mk-redo').disabled = !st.redo.length;
  }

  // ---------- tools ----------
  const setTool = (k) => {
    st.tool = k;
    root.querySelectorAll('[data-tool]').forEach((b) => b.setAttribute('aria-pressed', String(b.dataset.tool === k)));
    root.querySelector('.mk').dataset.tool = k;
    if (k !== 'select') { st.sel = null; paint(); }
  };
  root.querySelector('.mk-tools').onclick = (e) => { const b = e.target.closest('[data-tool]'); if (b) setTool(b.dataset.tool); };
  const wIn = root.querySelector('#mk-w');
  wIn.oninput = () => { st.width = +wIn.value; root.querySelector('#mk-wv').textContent = wIn.value; const m = st.marks.find((x) => x.id === st.sel); if (m && m.width) { m.width = st.width; changed(); } };
  root.querySelector('#mk-c').oninput = (e) => { st.color = e.target.value; const m = st.marks.find((x) => x.id === st.sel); if (m && m.type !== 'highlight') { m.color = st.color; changed(); } };
  const undo = () => { if (!st.undo.length) return; st.redo.push(JSON.stringify(st.marks)); st.marks = JSON.parse(st.undo.pop()); st.sel = null; changed(); };
  const redo = () => { if (!st.redo.length) return; st.undo.push(JSON.stringify(st.marks)); st.marks = JSON.parse(st.redo.pop()); st.sel = null; changed(); };
  const del = () => { if (!st.sel) return; snapshot(); st.marks = st.marks.filter((m) => m.id !== st.sel); st.sel = null; changed(); };
  root.querySelector('#mk-undo').onclick = undo; root.querySelector('#mk-redo').onclick = redo; root.querySelector('#mk-del').onclick = del;
  const onKey = (e) => {
    if (!root.isConnected) { document.removeEventListener('keydown', onKey); return; }
    if (e.target.isContentEditable || /INPUT|TEXTAREA/.test(e.target.tagName)) return;
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z') { e.preventDefault(); e.shiftKey ? redo() : undo(); }
    else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'y') { e.preventDefault(); redo(); }
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
  const markAt = (e) => e.target.closest('[data-id]')?.dataset.id || null;

  let drag = null;
  pagesEl.addEventListener('pointerdown', (e) => {
    const pel = e.target.closest('.mk-page'); if (!pel || e.button !== 0) return;
    const p = pages[+pel.dataset.page]; const pt = at(p, e);
    if (st.tool === 'highlight') return; // the text layer selects; pointerup turns the selection into a highlight
    if (st.tool === 'text') {
      const hit = e.target.closest('.mk-box');
      if (hit) { editBox(p, hit.dataset.id); e.preventDefault(); return; }
      snapshot();
      const m = { id: uid(), page: p.el.dataset.page * 1, type: 'text', x: pt[0], y: pt[1] - TEXT.size * 0.7, text: '', color: st.color, size: TEXT.size };
      st.marks.push(m); st.sel = m.id; paint(); editBox(p, m.id, true); e.preventDefault();
      return;
    }
    if (st.tool === 'select') {
      const id = markAt(e); st.sel = id; paint();
      if (!id) return;
      const m = st.marks.find((x) => x.id === id);
      if (m.type === 'text' && e.detail === 2) { editBox(p, id); return; }
      snapshot(); drag = { kind: 'move', p, last: pt, m, moved: false };
      pel.setPointerCapture(e.pointerId); e.preventDefault();
      return;
    }
    snapshot();
    const base = { id: uid(), page: +pel.dataset.page, color: st.color, width: st.width };
    const m = st.tool === 'pen' ? { ...base, type: 'pen', pts: [pt] } : st.tool === 'rect' ? { ...base, type: 'rect', x: pt[0], y: pt[1], w: 0, h: 0 } : { ...base, type: st.tool, a: pt, b: pt };
    st.marks.push(m); drag = { kind: 'draw', p, start: pt, m };
    pel.setPointerCapture(e.pointerId); e.preventDefault(); paint();
  });
  pagesEl.addEventListener('pointermove', (e) => {
    if (!drag) return;
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
  const endDrag = () => {
    if (!drag) return;
    const { m, kind, moved } = drag; drag = null;
    if (kind === 'move' && !moved) { st.undo.pop(); return; } // a click that only selected
    const tiny = (m.type === 'rect' && m.w < 2 && m.h < 2) || ((m.type === 'line' || m.type === 'arrow') && Math.hypot(m.b[0] - m.a[0], m.b[1] - m.a[1]) < 2) || (m.type === 'pen' && m.pts.length < 2);
    if (kind === 'draw' && tiny) { st.marks = st.marks.filter((x) => x !== m); st.undo.pop(); paint(); return; }
    changed();
  };
  pagesEl.addEventListener('pointerup', endDrag);
  pagesEl.addEventListener('pointercancel', endDrag);

  // Highlight: the text selected with the Highlight tool becomes yellow boxes over those words.
  pagesEl.addEventListener('mouseup', () => {
    if (st.tool !== 'highlight') return;
    const sel = window.getSelection(); if (!sel || sel.isCollapsed) return;
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
    byPage.forEach((rects, page) => st.marks.push({ id: uid(), page, type: 'highlight', rects: mergeRects(rects) }));
    changed();
  });

  // Text box editing: plain text, Enter = new line; an empty box is removed when it loses focus.
  function editBox(p, id, fresh = false) {
    const m = st.marks.find((x) => x.id === id); if (!m) return;
    const b = p.boxes.querySelector(`[data-id="${id}"]`); if (!b) return;
    if (!fresh) snapshot();
    b.contentEditable = 'plaintext-only'; if (b.contentEditable !== 'plaintext-only') b.contentEditable = 'true';
    b.classList.add('editing'); b.focus();
    const range = document.createRange(); range.selectNodeContents(b); range.collapse(false); const s = window.getSelection(); s.removeAllRanges(); s.addRange(range);
    b.onblur = () => {
      m.text = b.innerText.replace(/ /g, ' ').replace(/\n$/, '');
      if (!m.text.trim()) { st.marks = st.marks.filter((x) => x.id !== id); if (st.sel === id) st.sel = null; }
      changed();
    };
  }

  setTool('select');
  paint();

  // ---------- writing the marked PDF ----------
  async function exportPdf() {
    document.activeElement?.blur?.();
    const out = await PDFLib.PDFDocument.load(bytes.slice(0));
    const { rgb, LineCapStyle } = PDFLib;
    const hex = (c) => { const n = parseInt(c.slice(1), 16); return rgb(((n >> 16) & 255) / 255, ((n >> 8) & 255) / 255, (n & 255) / 255); };
    const outPages = out.getPages();
    for (const m of st.marks) {
      const pg = outPages[m.page]; const pv = pages[m.page]; if (!pg || !pv) continue;
      const { x: ox, y: oy } = pg.getMediaBox ? pg.getMediaBox() : { x: 0, y: 0 };
      // Page units (top-left origin) -> PDF user space (bottom-left origin). Pages with /Rotate are not handled in this test run.
      const P = ([x, y]) => ({ x: ox + x, y: oy + pv.h - y });
      const line = (a, b, w, c) => pg.drawLine({ start: P(a), end: P(b), thickness: w, color: hex(c), lineCap: LineCapStyle.Round });
      if (m.type === 'pen') for (let i = 1; i < m.pts.length; i++) line(m.pts[i - 1], m.pts[i], m.width, m.color);
      if (m.type === 'line' || m.type === 'arrow') line(m.a, m.b, m.width, m.color);
      if (m.type === 'arrow') {
        const [t0, t1, t2] = arrowHead(m.a, m.b, m.width).map(P);
        pg.drawSvgPath(`M ${t0.x} ${-t0.y} L ${t1.x} ${-t1.y} L ${t2.x} ${-t2.y} Z`, { x: 0, y: 0, color: hex(m.color), borderWidth: 0 });
      }
      if (m.type === 'rect') {
        const c = [[m.x, m.y], [m.x + m.w, m.y], [m.x + m.w, m.y + m.h], [m.x, m.y + m.h]];
        for (let i = 0; i < 4; i++) line(c[i], c[(i + 1) % 4], m.width, m.color);
      }
      if (m.type === 'highlight') {
        m.rects.forEach(([x, y, w, h]) => {
          const a = P([x, y + h]);
          pg.drawRectangle({ x: a.x, y: a.y, width: w, height: h, color: hex(HIGHLIGHT), opacity: 0.45, blendMode: PDFLib.BlendMode?.Multiply });
        });
      }
      if (m.type === 'text' && m.text.trim()) {
        const png = await textPng(m);
        const img = await out.embedPng(png.bytes);
        const a = P([m.x, m.y + png.h]);
        pg.drawImage(img, { x: a.x, y: a.y, width: png.w, height: png.h });
      }
    }
    return out.save();
  }

  return { marks: () => structuredClone(st.marks), isDirty: () => st.dirty, clean: () => { st.dirty = false; }, exportPdf, name };
}

// A text box drawn by the browser at 4x (so Thai vowels and tone marks are placed correctly), as PNG bytes and its
// size in page units.
async function textPng(m) {
  const k = 4; const lines = m.text.split('\n'); const lh = m.size * TEXT.lineHeight;
  const c = document.createElement('canvas'); const ctx = c.getContext('2d');
  ctx.font = `${m.size * k}px "${TEXT.font}"`;
  const w = Math.max(...lines.map((l) => ctx.measureText(l).width)) / k + 2; const h = lh * lines.length;
  c.width = Math.ceil(w * k); c.height = Math.ceil(h * k);
  ctx.font = `${m.size * k}px "${TEXT.font}"`; ctx.fillStyle = m.color; ctx.textBaseline = 'middle';
  lines.forEach((l, i) => ctx.fillText(l, 0, (i + 0.5) * lh * k));
  const blob = await new Promise((r) => c.toBlob(r, 'image/png'));
  return { bytes: new Uint8Array(await blob.arrayBuffer()), w: c.width / k, h: c.height / k };
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
