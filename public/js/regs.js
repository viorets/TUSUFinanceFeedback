// Regulations sidebar (v2.17): the menu button at the left of the top bar opens a panel from the left with two modes,
// สรุป (rates compared side by side, by topic) and ระเบียบ (each regulation's text, laid out like its PDF), from
// data/regs.json (tools/build-regs.py). Search covers both modes as you type (Thai and Arabic digits match), with a
// count per mode, next / previous and a jump to the hit. Zoom scales only the panel's text: Ctrl + wheel or a trackpad
// pinch inside the panel (the browser's own zoom is held off there), a two-finger pinch on a touch screen, or the
// slider and − / +; it is remembered on this device. Labels are Thai whatever the site's language (decided 26 Oct).
// Signed-in users only: the button is drawn with the top bar (app.js shell), which needs a sign-in.
import { esc, icon } from './ui.js';

const ZOOM = { min: 0.7, max: 2, key: 'regs:zoom' };
const WIDTH = { min: 360, def: 520, key: 'regs:width' }; // the panel's width, dragged by its right edge (not on phones)
let el = null; // the panel
let data = null;
let mode = 'sum'; // 'sum' | 'reg'
let hits = { sum: [], reg: [] }; let at = -1;
let zoom = 1;

const keep = { get: (k) => { try { return localStorage.getItem(k); } catch { return null; } }, set: (k, v) => { try { localStorage.setItem(k, v); } catch { /* private mode */ } } };

export async function toggle() {
  if (el?.classList.contains('open')) return close();
  if (!el) build();
  el.classList.add('open'); scrim().classList.add('open');
  document.documentElement.classList.add('rg-lock');
  $('[data-q]').focus({ preventScroll: true });
  if (!data) await load();
}
function close() {
  el.classList.remove('open'); scrim().classList.remove('open');
  document.documentElement.classList.remove('rg-lock');
  document.getElementById('regs')?.focus({ preventScroll: true });
}
const $ = (s) => el.querySelector(s);
const scrim = () => document.querySelector('.rg-scrim');

function build() {
  const s = document.createElement('div'); s.className = 'rg-scrim'; s.onclick = close; document.body.append(s);
  el = document.createElement('aside');
  el.className = 'rg-panel'; el.setAttribute('role', 'dialog'); el.setAttribute('aria-label', 'ระเบียบและอัตรา'); el.lang = 'th';
  el.innerHTML = `<header class="rg-head">
      <div class="rg-row"><div class="seg rg-seg" role="tablist">
        <button role="tab" data-mode="sum" aria-selected="true">สรุป</button><button role="tab" data-mode="reg" aria-selected="false">ระเบียบ</button></div>
        <span class="rg-grow"></span>
        <button class="rg-icon" data-close aria-label="ปิด" title="ปิด (Esc)">${icon('x', 20)}</button></div>
      <div class="rg-row rg-find"><label class="rg-search">${icon('search', 16)}<input type="search" data-q placeholder="ค้นหา เช่น 8708, เงินยืม, ที่พัก" aria-label="ค้นหา" enterkeyhint="search"></label>
        <span class="rg-count" data-count aria-live="polite"></span>
        <button class="rg-icon rg-up" data-prev aria-label="ก่อนหน้า" title="ก่อนหน้า (Shift+Enter)" disabled>${icon('down', 18)}</button>
        <button class="rg-icon" data-next aria-label="ถัดไป" title="ถัดไป (Enter)" disabled>${icon('down', 18)}</button></div>
      <div class="rg-row rg-zoom"><span class="rg-lbl">ซูม</span>
        <button class="rg-icon" data-zoom="-1" aria-label="ซูมออก">−</button>
        <input type="range" data-zoomr min="${ZOOM.min * 100}" max="${ZOOM.max * 100}" step="5" aria-label="ขนาดตัวอักษร">
        <button class="rg-icon" data-zoom="1" aria-label="ซูมเข้า">+</button><span class="rg-pct" data-pct></span></div>
      <nav class="rg-nav" data-nav aria-label="หัวข้อ"></nav>
    </header>
    <div class="rg-body" data-body tabindex="-1"><div data-pane="sum"></div><div data-pane="reg" hidden></div>
      <p class="rg-wait" data-wait>กำลังโหลด…</p></div>
    <div class="rg-resize" role="separator" aria-orientation="vertical" aria-label="ปรับความกว้าง" title="ลากเพื่อปรับความกว้าง" tabindex="0"></div>`;
  document.body.append(el);
  el.addEventListener('click', onClick);
  $('[data-q]').addEventListener('input', () => search($('[data-q]').value));
  $('[data-q]').addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); step(e.shiftKey ? -1 : 1); } });
  $('[data-zoomr]').addEventListener('input', (e) => setZoom(e.target.value / 100));
  document.addEventListener('keydown', (e) => {
    if (e.key !== 'Escape') return;
    if (document.querySelector('.rg-pic')) closePic(); else if (el.classList.contains('open')) close();
  });
  zoomGestures();
  setZoom(+keep.get(ZOOM.key) || 1);
  resizer();
}

async function load() {
  try {
    data = await (await fetch('data/regs.json', { cache: 'no-cache' })).json();
  } catch {
    $('[data-wait]').textContent = 'โหลดไม่สำเร็จ กรุณาลองอีกครั้ง'; return;
  }
  $('[data-wait]').remove();
  $('[data-pane="sum"]').innerHTML = `<h1 class="rg-title">${esc(data.summary.title)}</h1>${data.summary.html}`;
  $('[data-pane="reg"]').innerHTML = data.regs.map((r) => `<article class="rg-doc${r.kind === 'guide' ? ' rg-guide' : ''}" id="rg-doc-${r.id}" style="--rg-font:'${r.font}'">
      <h1 class="rg-doctitle">${esc(r.title)}</h1>${r.kind === 'guide' ? '<p class="rg-tag">แนวปฏิบัติ ไม่ใช่ระเบียบ</p>' : ''}${r.html}</article>`).join('');
  linkClauses();
  nav();
  if ($('[data-q]').value) search($('[data-q]').value);
}

// ---------- modes and the topic list ----------
function setMode(m) {
  mode = m;
  el.querySelectorAll('[data-mode]').forEach((b) => b.setAttribute('aria-selected', String(b.dataset.mode === m)));
  $('[data-pane="sum"]').hidden = m !== 'sum'; $('[data-pane="reg"]').hidden = m !== 'reg';
  $('[data-body]').scrollTop = 0;
  el.querySelector('mark.rg-cur')?.classList.remove('rg-cur'); at = -1; // the hit count starts again in this mode
  nav(); if (hits[m].length) step(1); else count();
}
function nav() {
  if (!data) return;
  $('[data-nav]').innerHTML = mode === 'sum'
    ? data.summary.toc.map((t) => `<a href="#" data-to="${t.id}">${esc(t.title)}</a>`).join('')
    : data.regs.map((r) => `<a href="#" data-to="rg-doc-${r.id}" title="${esc(r.title)}">${esc(r.short)}</a>`).join('');
}
function go(id, { flash = false } = {}) {
  const t = document.getElementById(id); if (!t) return;
  const body = $('[data-body]');
  body.scrollTop += t.getBoundingClientRect().top - body.getBoundingClientRect().top - 8;
  if (flash) { t.classList.remove('rg-flash'); void t.offsetWidth; t.classList.add('rg-flash'); }
}

function onClick(e) {
  if (e.target.closest('[data-close]')) return close();
  const pic = e.target.closest('.rg-card a[href]');
  if (pic) { e.preventDefault(); return showPic(pic.getAttribute('href'), pic.querySelector('img')?.alt || ''); }
  const m = e.target.closest('[data-mode]'); if (m) return setMode(m.dataset.mode);
  const to = e.target.closest('[data-to]'); if (to) { e.preventDefault(); return go(to.dataset.to); }
  const z = e.target.closest('[data-zoom]'); if (z) return setZoom(zoom + z.dataset.zoom * 0.1);
  if (e.target.closest('[data-next]')) return step(1);
  if (e.target.closest('[data-prev]')) return step(-1);
  const r = e.target.closest('a[data-reg]');
  if (r) { // a link to a regulation (or a ข้อ in it): Regulation mode at that place
    e.preventDefault();
    if (mode !== 'reg') setMode('reg');
    go(r.dataset.id || (r.dataset.at ? `r${r.dataset.reg}-${r.dataset.at}` : `rg-doc-${r.dataset.reg}`), { flash: true });
  }
}

// Summary cells name a ข้อ of the regulation their column is about ("05 ...", "04 ..."): each becomes a link to the
// most specific heading of that regulation whose ข้อ range holds it.
function linkClauses() {
  const ranges = {};
  data.regs.forEach((r) => {
    ranges[r.id] = [...document.querySelectorAll(`#rg-doc-${r.id} h2[id], #rg-doc-${r.id} h3[id]`)].flatMap((h) => {
      const m = h.textContent.match(/ข้อ\s?(\d+)(?:\s?[–-]\s?(\d+))?/); if (!m) return [];
      return [{ id: h.id, a: +m[1], b: +(m[2] || m[1]) }];
    });
  });
  const find = (doc, n) => (ranges[doc] || []).filter((x) => x.a <= n && n <= x.b).sort((x, y) => (x.b - x.a) - (y.b - y.a))[0];
  $('[data-pane="sum"]').querySelectorAll('table').forEach((t) => {
    const docs = [...t.querySelectorAll('thead th')].map((th) => th.textContent.match(/^(0\d)\b/)?.[1] || '');
    t.querySelectorAll('tbody tr').forEach((tr) => [...tr.children].forEach((td, k) => {
      const doc = docs[k]; if (!doc) return;
      td.innerHTML = td.innerHTML.replace(/\(ข้อ\s?(\d+)([^)]*)\)/g, (all, n) => {
        const h = find(doc, +n); return h ? `(<a href="#" data-reg="${doc}" data-id="${h.id}" title="เปิดในโหมดระเบียบ">ข้อ ${n}</a>${all.slice(all.indexOf(n) + n.length, -1)})` : all;
      });
    }));
  });
}

// ---------- a sample picture, full size on a dark screen in this page (v2.17): a tap on the dark part or Esc closes it ----------
function showPic(src, alt) {
  closePic();
  const box = document.createElement('div');
  box.className = 'rg-pic'; box.setAttribute('role', 'dialog'); box.setAttribute('aria-label', alt);
  box.innerHTML = `<img src="${esc(src)}" alt="${esc(alt)}"><button class="rg-pic-x" aria-label="ปิด">${icon('x', 22)}</button>`;
  box.onclick = (e) => { if (e.target !== box.querySelector('img')) closePic(); };
  document.body.append(box);
  box.querySelector('.rg-pic-x').focus({ preventScroll: true });
}
const closePic = () => document.querySelector('.rg-pic')?.remove();

// ---------- search ----------
const TH = '๐๑๒๓๔๕๖๗๘๙';
const norm = (s) => s.replace(/[๐-๙]/g, (d) => String(TH.indexOf(d))).toLowerCase(); // same length: offsets match
function clearMarks(root) {
  root.querySelectorAll('mark.rg-hit').forEach((m) => m.replaceWith(...m.childNodes));
  root.normalize();
}
function mark(root, q) {
  const out = [];
  const w = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, { acceptNode: (n) => (n.parentElement.closest('script, style') ? NodeFilter.FILTER_REJECT : NodeFilter.FILTER_ACCEPT) });
  const nodes = []; for (let n = w.nextNode(); n; n = w.nextNode()) nodes.push(n);
  for (const n of nodes) {
    const t = norm(n.data); let i = t.indexOf(q); let node = n;
    let base = 0;
    while (i >= 0) {
      const hit = node.splitText(i - base); const rest = hit.splitText(q.length);
      const m = document.createElement('mark'); m.className = 'rg-hit'; hit.replaceWith(m); m.append(hit);
      out.push(m); node = rest; base = i + q.length; i = t.indexOf(q, base);
    }
  }
  return out;
}
let searchT = 0;
function search(raw) {
  clearTimeout(searchT);
  searchT = setTimeout(() => {
    if (!data) return;
    const q = norm(raw.trim());
    ['sum', 'reg'].forEach((k) => clearMarks($(`[data-pane="${k}"]`)));
    hits = { sum: [], reg: [] }; at = -1;
    if (q.length >= 1) ['sum', 'reg'].forEach((k) => { hits[k] = mark($(`[data-pane="${k}"]`), q); });
    count();
    if (hits[mode].length) step(1);
  }, 150);
}
function count() {
  const q = $('[data-q]').value.trim(); const c = $('[data-count]');
  const other = mode === 'sum' ? 'reg' : 'sum';
  if (!q || !data) { c.innerHTML = ''; } else {
    const n = hits[mode].length;
    c.innerHTML = `${n ? `${at + 1 > 0 ? at + 1 : 0}/${n}` : 'ไม่พบ'}${hits[other].length ? ` · <button class="rg-link" data-mode="${other}">${other === 'sum' ? 'สรุป' : 'ระเบียบ'} ${hits[other].length}</button>` : ''}`;
  }
  $('[data-prev]').disabled = $('[data-next]').disabled = !hits[mode].length;
}
function step(d) {
  const list = hits[mode]; if (!list.length) return;
  list[at]?.classList.remove('rg-cur');
  at = (at + d + list.length) % list.length;
  const m = list[at]; m.classList.add('rg-cur');
  const body = $('[data-body]'); const r = m.getBoundingClientRect(); const b = body.getBoundingClientRect();
  body.scrollTop += r.top - b.top - b.height / 3;
  const scroller = m.closest('.rg-table'); if (scroller) scroller.scrollLeft += r.left - scroller.getBoundingClientRect().left - 40;
  count();
}

// ---------- width: drag the right edge (mouse, pen or finger), or ← / → on the focused edge ----------
function setWidth(w, save = true) {
  const max = Math.max(WIDTH.min, innerWidth - 40);
  w = Math.round(Math.min(max, Math.max(WIDTH.min, w)));
  el.style.setProperty('--rg-w', `${w}px`);
  el.querySelector('.rg-resize').setAttribute('aria-valuenow', w);
  if (save) keep.set(WIDTH.key, String(w));
}
function resizer() {
  const h = el.querySelector('.rg-resize');
  setWidth(+keep.get(WIDTH.key) || WIDTH.def, false);
  h.addEventListener('pointerdown', (e) => {
    e.preventDefault(); h.setPointerCapture(e.pointerId); el.classList.add('rg-dragging');
    const left = el.getBoundingClientRect().left;
    const move = (ev) => setWidth(ev.clientX - left, false);
    const up = (ev) => { setWidth(ev.clientX - left); el.classList.remove('rg-dragging'); h.removeEventListener('pointermove', move); h.removeEventListener('pointerup', up); h.removeEventListener('pointercancel', up); };
    h.addEventListener('pointermove', move); h.addEventListener('pointerup', up); h.addEventListener('pointercancel', up);
  });
  h.addEventListener('keydown', (e) => { if (e.key === 'ArrowLeft' || e.key === 'ArrowRight') { e.preventDefault(); setWidth(el.offsetWidth + (e.key === 'ArrowRight' ? 40 : -40)); } });
  h.addEventListener('dblclick', () => setWidth(WIDTH.def)); // back to the standard width
}

// ---------- zoom ----------
function setZoom(z) {
  zoom = Math.min(ZOOM.max, Math.max(ZOOM.min, Math.round(z * 100) / 100));
  const body = $('[data-body]'); const keepAt = body.scrollHeight ? body.scrollTop / body.scrollHeight : 0; // the same place stays in view
  el.style.setProperty('--rg-zoom', zoom);
  body.scrollTop = keepAt * body.scrollHeight;
  $('[data-zoomr]').value = Math.round(zoom * 100); $('[data-pct]').textContent = `${Math.round(zoom * 100)}%`;
  keep.set(ZOOM.key, String(zoom));
}
function zoomGestures() {
  // Ctrl + wheel, and a trackpad pinch (Chrome / Edge send it as Ctrl + wheel): the panel's text, not the page.
  el.addEventListener('wheel', (e) => { if (!e.ctrlKey) return; e.preventDefault(); setZoom(zoom * Math.exp(-Math.max(-100, Math.min(100, e.deltaY)) / 600)); /* a wheel notch ≈ 18% */ }, { passive: false });
  // Safari's trackpad pinch.
  let g0 = 1;
  el.addEventListener('gesturestart', (e) => { e.preventDefault(); g0 = zoom; });
  el.addEventListener('gesturechange', (e) => { e.preventDefault(); setZoom(g0 * e.scale); });
  // Two fingers on a touch screen (the panel's touch-action leaves pinching to this).
  let d0 = 0, z0 = 1;
  const dist = (t) => Math.hypot(t[0].clientX - t[1].clientX, t[0].clientY - t[1].clientY);
  el.addEventListener('touchstart', (e) => { if (e.touches.length === 2) { d0 = dist(e.touches); z0 = zoom; } }, { passive: true });
  el.addEventListener('touchmove', (e) => { if (e.touches.length === 2 && d0) { e.preventDefault(); setZoom(z0 * dist(e.touches) / d0); } }, { passive: false });
  el.addEventListener('touchend', (e) => { if (e.touches.length < 2) d0 = 0; });
}
