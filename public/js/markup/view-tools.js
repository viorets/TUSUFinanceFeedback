// View tools shared by the PDF, Excel and Word editors: icon buttons, zoom by Ctrl/⌘ + wheel or pinch (trackpad, finger),
// a zoom control, and full screen with the tools on top.
const svg = (d) => `<svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${d}</svg>`;
export const ICON = {
  undo: svg('<path d="M9 14 4 9l5-5"/><path d="M4 9h10.5a5.5 5.5 0 0 1 0 11H11"/>'),
  redo: svg('<path d="m15 14 5-5-5-5"/><path d="M20 9H9.5a5.5 5.5 0 0 0 0 11H13"/>'),
  del: svg('<path d="M3 6h18"/><path d="M8 6V4h8v2"/><path d="m19 6-1 14H6L5 6"/><path d="M10 11v6M14 11v6"/>'),
  full: svg('<path d="M8 3H3v5M16 3h5v5M3 16v5h5M21 16v5h-5"/>'),
  unfull: svg('<path d="M3 8h5V3M21 8h-5V3M3 16h5v5M21 16h-5v5"/>'),
  zin: svg('<path d="M12 5v14M5 12h14"/>'),
  zout: svg('<path d="M5 12h14"/>'),
  sort: svg('<path d="M7 4v16M3 8l4-4 4 4M17 20V4M13 16l4 4 4-4"/>'),
  sample: svg('<path d="M12 3l1.9 4.6 4.6 1.9-4.6 1.9L12 16l-1.9-4.6-4.6-1.9 4.6-1.9z"/><path d="M19 15l.9 2.1L22 18l-2.1.9L19 21l-.9-2.1L16 18l2.1-.9z"/>'),
  save: svg('<path d="M5 3h11l3 3v13a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2z"/><path d="M8 3v5h7V3M8 21v-7h8v7"/>'),
  open: svg('<path d="M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v1"/><path d="M3 7v10a2 2 0 0 0 2 2h13l3-9H8l-3 9"/>'),
  help: svg('<circle cx="12" cy="12" r="9"/><path d="M9.5 9.5a2.5 2.5 0 1 1 3.5 2.3c-.6.3-1 .9-1 1.6v.6"/><circle cx="12" cy="17" r=".6" fill="currentColor"/>'),
  dl: svg('<path d="M12 3v12M7 10l5 5 5-5"/><path d="M4 17v2a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-2"/>'),
  chev: svg('<path d="m6 9 6 6 6-6"/>'),
  hl: svg('<path d="m9 11-6 6v3h9l3-3"/><path d="m22 12-4.6 4.6a2 2 0 0 1-2.8 0l-5.2-5.2a2 2 0 0 1 0-2.8L14 4"/>'),
};
export const ZOOM_MIN = 0.5; export const ZOOM_MAX = 4;
export const clampZoom = (z) => Math.min(ZOOM_MAX, Math.max(ZOOM_MIN, Math.round(z * 100) / 100));
export const iconBtn = (id, label, face) => `<button type="button" id="${id}" class="vt-ic" title="${label}" aria-label="${label}">${face}</button>`;
// Zoom − / % / + and Save draft, Download, Undo, Redo, Delete, Full screen, as one markup string. `t(en, th)` translates.
// Save draft and Download sit here so they work in full screen too; the page wires them (index.html, fileActions).
// opts.open: an Open button first (the PDF editor, which has no File tab); opts.help: a "?" button last (tips).
export const viewBar = (p, t = (en) => en, { open = false, help = false } = {}) => `<div class="vt-zoom" role="group" aria-label="${t('Zoom', 'ซูม')}">${iconBtn(`${p}-zout`, t('Zoom out', 'ซูมออก'), ICON.zout)}<button type="button" id="${p}-zfit" class="vt-zpct" title="${t('Back to 100%', 'กลับเป็น 100%')}">100%</button>${iconBtn(`${p}-zin`, t('Zoom in', 'ซูมเข้า'), ICON.zin)}</div>
  <div class="vt-acts">${open ? iconBtn(`${p}-open`, t('Open another file', 'เปิดไฟล์อื่น'), ICON.open) : ''}${iconBtn(`${p}-draft`, t('Save draft on this device', 'บันทึกฉบับร่างในเครื่องนี้'), ICON.save)}${iconBtn(`${p}-dl`, t('Download', 'ดาวน์โหลด'), ICON.dl)}${iconBtn(`${p}-undo`, t('Undo (Ctrl+Z)', 'เลิกทำ (Ctrl+Z)'), ICON.undo)}${iconBtn(`${p}-redo`, t('Redo (Ctrl+Y)', 'ทำซ้ำ (Ctrl+Y)'), ICON.redo)}${iconBtn(`${p}-del`, t('Delete', 'ลบ'), ICON.del)}${iconBtn(`${p}-full`, t('Full screen', 'เต็มจอ'), ICON.full)}${help ? iconBtn(`${p}-help`, t('Tips and shortcuts', 'เคล็ดลับและคีย์ลัด'), ICON.help) : ''}</div>`;

// Zoom gestures on `el`: Ctrl/⌘ + wheel (Chrome and Edge also send a trackpad pinch this way; a plain wheel still
// scrolls), and finger pinch: Safari (iPad, Mac trackpad) sends gesture events with the scale, other browsers two
// touches. zoomTo(z, clientX, clientY) gets the new zoom and the screen point that should stay put.
export function pinchZoom(el, getZoom, zoomTo) {
  el.addEventListener('wheel', (e) => {
    if (!(e.ctrlKey || e.metaKey)) return;
    e.preventDefault(); zoomTo(getZoom() * Math.exp(-e.deltaY * 0.0025), e.clientX, e.clientY);
  }, { passive: false });
  if ('ongesturestart' in window) {
    let z0 = 1;
    el.addEventListener('gesturestart', (e) => { e.preventDefault(); z0 = getZoom(); });
    el.addEventListener('gesturechange', (e) => { e.preventDefault(); zoomTo(z0 * e.scale, e.clientX, e.clientY); });
    el.addEventListener('gestureend', (e) => e.preventDefault());
    return;
  }
  let pinch = null;
  const dist = (ts) => Math.hypot(ts[0].clientX - ts[1].clientX, ts[0].clientY - ts[1].clientY);
  el.addEventListener('touchstart', (e) => { if (e.touches.length === 2) pinch = { d: dist(e.touches), z: getZoom() }; }, { passive: true });
  el.addEventListener('touchmove', (e) => {
    if (!pinch || e.touches.length !== 2) return;
    e.preventDefault();
    zoomTo(pinch.z * dist(e.touches) / pinch.d, (e.touches[0].clientX + e.touches[1].clientX) / 2, (e.touches[0].clientY + e.touches[1].clientY) / 2);
  }, { passive: false });
  el.addEventListener('touchend', (e) => { if (e.touches.length < 2) pinch = null; });
}

// Keeps the content point under (cx, cy) in place while `scroller`'s content grows or shrinks by k = new / old zoom.
// apply() resizes the content. No point: the centre of the scroller.
export function zoomAround(scroller, k, cx, cy, apply) {
  const r = scroller.getBoundingClientRect();
  if (cx == null) { cx = r.left + r.width / 2; cy = r.top + r.height / 2; }
  const ox = scroller.scrollLeft + cx - r.left; const oy = scroller.scrollTop + cy - r.top;
  apply();
  scroller.scrollLeft = ox * k - (cx - r.left); scroller.scrollTop = oy * k - (cy - r.top);
}

// Zoom by CSS `zoom` on the element content() returns (Excel grid, Word page); gestures, keys (Ctrl/⌘ + = − 0 while
// the focus is inside `box`) and the −/%/+ buttons with id prefix p.
export function cssZoom(box, scroller, content, p) {
  let z = 1;
  const pct = box.querySelector(`#${p}-zfit`);
  const zoomTo = (nz, cx, cy) => {
    nz = clampZoom(nz); if (nz === z) return;
    const k = nz / z; z = nz;
    zoomAround(scroller, k, cx, cy, () => { const c = content(); if (c) c.style.zoom = z; });
    pct.textContent = `${Math.round(z * 100)}%`;
  };
  pinchZoom(scroller, () => z, zoomTo);
  box.querySelector(`#${p}-zin`).onclick = () => zoomTo(z * 1.25);
  box.querySelector(`#${p}-zout`).onclick = () => zoomTo(z / 1.25);
  pct.onclick = () => zoomTo(1);
  box.addEventListener('keydown', (e) => {
    if (!(e.ctrlKey || e.metaKey)) return;
    if (e.key === '=' || e.key === '+') { e.preventDefault(); zoomTo(z * 1.25); }
    else if (e.key === '-') { e.preventDefault(); zoomTo(z / 1.25); }
    else if (e.key === '0') { e.preventDefault(); zoomTo(1); }
  });
  return { get: () => z, zoomTo };
}

// A toolbar floating over the content that scrolls under it: keeps --bar-h on `box` equal to the bar's height (it
// wraps onto more rows on narrow screens), so the content and sticky headers start below it.
export function floatBar(box, bar) {
  const set = () => box.style.setProperty('--bar-h', `${bar.offsetHeight}px`);
  new ResizeObserver(set).observe(bar); set();
}

// Full screen for `box`: class `cls` makes the box fill the browser window (Escape leaves). The button with id `${p}-full` toggles it and swaps its icon.
export function fullScreen(box, p, cls, t = (en) => en, onChange = () => {}) {
  const fsEl = () => document.fullscreenElement || document.webkitFullscreenElement;
  const btn = box.querySelector(`#${p}-full`);
  let api = false;
  const isOn = () => box.classList.contains(cls);
  const set = (on) => {
    if (on === isOn()) return;
    box.classList.toggle(cls, on);
    const label = on ? t('Exit full screen', 'ออกจากเต็มจอ') : t('Full screen', 'เต็มจอ');
    btn.innerHTML = on ? ICON.unfull : ICON.full; btn.title = label; btn.setAttribute('aria-label', label);
    onChange(on);
  };
  const names = ['fullscreenchange', 'webkitfullscreenchange'];
  const onFs = () => {
    if (!box.isConnected) { names.forEach((n) => document.removeEventListener(n, onFs)); return; }
    if (fsEl() === box) api = true; else if (api) { api = false; set(false); }
  };
  names.forEach((n) => document.addEventListener(n, onFs));
  const onKey = (e) => {
    if (!box.isConnected) { document.removeEventListener('keydown', onKey); return; }
    if (e.key !== 'Escape' || e.defaultPrevented || !isOn() || fsEl() || /INPUT|SELECT/.test(e.target.tagName)) return; // typing in the text or page still leaves
    set(false);
  };
  document.addEventListener('keydown', onKey);
  btn.onclick = async () => {
    document.activeElement?.blur?.();
    if (isOn()) { if (fsEl()) (document.exitFullscreen || document.webkitExitFullscreen).call(document); set(false); return; }
    // The editor fills the browser window. The browser's own full screen is not used: Microsoft Edge's scareware
    // blocker flags pages that take over the screen (seen on 10 Nov 2026). F11 still makes the window full screen.
    set(true);
  };
  return { isOn };
}

// Tips and shortcuts under the "?" button: a small panel in the editor box (so it shows in full screen too); the button
// opens and closes it, as do Escape and a click outside.
export function helpPop(btn, html) {
  btn.setAttribute('aria-expanded', 'false');
  btn.onclick = () => {
    const box = btn.closest('.oe, .mk') || document.body;
    let m = box.querySelector(':scope > .vt-help');
    if (m && !m.hidden) { m.hidden = true; btn.setAttribute('aria-expanded', 'false'); return; }
    if (!m) { m = document.createElement('div'); m.className = 'vt-help'; m.setAttribute('role', 'dialog'); m.setAttribute('aria-label', 'Tips and shortcuts'); m.tabIndex = -1; box.append(m); }
    m.innerHTML = `<h3>Tips and shortcuts</h3>${html}`;
    const r = btn.getBoundingClientRect();
    m.hidden = false; m.style.top = `${r.bottom + 6}px`; m.style.left = `${Math.max(8, Math.min(r.right - 340, innerWidth - 348))}px`;
    btn.setAttribute('aria-expanded', 'true'); m.focus({ preventScroll: true });
    const close = (e) => { if (e && (m.contains(e.target) || btn.contains(e.target))) return; m.hidden = true; btn.setAttribute('aria-expanded', 'false'); removeEventListener('pointerdown', close, true); removeEventListener('keydown', key, true); };
    const key = (e) => { if (e.key === 'Escape') { e.stopPropagation(); close(); btn.focus(); } };
    addEventListener('pointerdown', close, true); addEventListener('keydown', key, true);
  };
}

// A small menu under `btn` (the Download button's file formats): items [[value, label, note]]; onPick(value). Placed
// in the button's editor box, so it shows in full screen too; Escape or a click outside closes it.
export function actionMenu(btn, items, onPick) {
  const box = btn.closest('.oe') || document.body;
  let m = box.querySelector(':scope > .vt-menu');
  if (!m) { m = document.createElement('div'); m.className = 'vt-menu'; m.setAttribute('role', 'menu'); box.append(m); }
  m.innerHTML = items.map(([v, label, note]) => `<button type="button" role="menuitem" data-v="${v}">${label}<small>${note}</small></button>`).join('');
  btn.setAttribute('aria-haspopup', 'menu');
  const r = btn.getBoundingClientRect();
  m.hidden = false; m.style.top = `${r.bottom + 6}px`; m.style.left = `${Math.max(8, Math.min(r.right - 240, innerWidth - 248))}px`;
  m.querySelector('button').focus();
  const close = (e) => { if (e && m.contains(e.target)) return; m.hidden = true; removeEventListener('pointerdown', close, true); removeEventListener('keydown', key, true); };
  const key = (e) => {
    if (e.key === 'Escape') { e.stopPropagation(); e.preventDefault(); close(); btn.focus(); return; }
    const all = [...m.querySelectorAll('button')]; const i = all.indexOf(document.activeElement);
    const to = { ArrowDown: i + 1, ArrowUp: i - 1, Home: 0, End: all.length - 1 }[e.key]; if (to == null) return;
    e.preventDefault(); all[(to + all.length) % all.length].focus();
  };
  addEventListener('pointerdown', close, true); addEventListener('keydown', key, true);
  m.onclick = (e) => { const b = e.target.closest('[data-v]'); if (!b) return; close(); onPick(b.dataset.v); };
}

// Alert window (from TUSUFinanceFeedback public/js/ui.js dialog): a title, a message, buttons; the page behind is
// dimmed. Resolves the value of the button pressed (Esc or a click outside gives `none`). html is already escaped.
// buttons: [{ v, label, primary }]; the last one gets the focus.
export function noticeBox(title, html, buttons = [{ v: true, label: 'OK', primary: true }], { none = null, tone = '' } = {}) {
  return new Promise((resolve) => {
    const back = document.createElement('div'); back.className = 'vt-modal-back';
    const esc = (t) => String(t).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
    const id = `vt-m-${Date.now()}`;
    back.innerHTML = `<div class="vt-modal ${tone}" role="alertdialog" aria-modal="true" aria-labelledby="${id}">
      ${tone === 'warn' ? `<svg class="vt-modal-ic" width="26" height="26" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 9v4M12 17h.01"/><path d="M10.3 3.9 1.8 18a2 2 0 0 0 1.7 3h17a2 2 0 0 0 1.7-3L13.7 3.9a2 2 0 0 0-3.4 0z"/></svg>` : ''}
      <h2 id="${id}">${esc(title)}</h2>${html}
      <div class="vt-modal-acts">${buttons.map((b, i) => `<button type="button" class="vt-modal-btn${b.primary ? ' primary' : ''}" data-i="${i}">${esc(b.label)}</button>`).join('')}</div></div>`;
    const before = document.activeElement;
    const done = (v) => { back.remove(); removeEventListener('keydown', key, true); before?.focus?.({ preventScroll: true }); resolve(v); };
    const key = (e) => {
      if (e.key === 'Escape') { e.stopPropagation(); e.preventDefault(); done(none); }
      if (e.key === 'Tab') { // keep the focus on the window's buttons
        const bs = [...back.querySelectorAll('button')]; const i = bs.indexOf(document.activeElement);
        e.preventDefault(); bs[(i + (e.shiftKey ? -1 : 1) + bs.length) % bs.length].focus();
      }
    };
    back.addEventListener('click', (e) => { if (e.target === back) done(none); const b = e.target.closest('[data-i]'); if (b) done(buttons[Number(b.dataset.i)].v); });
    addEventListener('keydown', key, true);
    document.body.append(back);
    back.querySelector('.vt-modal-acts button:last-child').focus();
  });
}
