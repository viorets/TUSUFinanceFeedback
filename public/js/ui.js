import { tr, locale } from './i18n.js';

export const $ = (s, root = document) => root.querySelector(s);
export const $$ = (s, root = document) => [...root.querySelectorAll(s)];

export const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

// Money is always shown with thousands separators and 2 decimals: 1,000.00
export const num2 = (n) => (+n || 0).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
export const moneyOr = (n) => (n == null || n === '' ? '—' : num2(n));
export const moneyBox = (n) => (n == null || n === '' ? '' : num2(n));
// Amount typed into a money box → number (or null when empty). Commas are ignored.
export const parseMoney = (v) => { const t = String(v ?? '').replace(/[^\d.]/g, ''); return t === '' || t === '.' ? null : Math.round(parseFloat(t) * 100) / 100; };
// While typing: keep digits and one dot (max 2 decimals), add commas, keep the caret in place.
export function liveMoney(input) {
  const raw = input.value; const caret = input.selectionStart ?? raw.length;
  const digitsBefore = raw.slice(0, caret).replace(/[^\d.]/g, '').length;
  let [int = '', dec] = raw.replace(/[^\d.]/g, '').split(/\.(.*)/s);
  int = int.replace(/^0+(?=\d)/, '');
  const hasDot = raw.includes('.');
  dec = hasDot ? String(dec ?? '').replace(/\./g, '').slice(0, 2) : undefined;
  const out = int.replace(/\B(?=(\d{3})+(?!\d))/g, ',') + (hasDot ? `.${dec}` : '');
  input.value = out;
  let pos = 0; let seen = 0;
  while (pos < out.length && seen < digitsBefore) { if (/[\d.]/.test(out[pos])) seen++; pos++; }
  try { input.setSelectionRange(pos, pos); } catch {}
}

// Phone: digits only, shown as XXX-XXX-XXXX.
export const phoneFmt = (v) => { const d = String(v || '').replace(/\D/g, ''); return [d.slice(0, 3), d.slice(3, 6), d.slice(6)].filter(Boolean).join('-'); };
export const dateTimeFmt = (v) => {
  const d = v ? new Date(v) : null;
  return d && !Number.isNaN(+d) ? d.toLocaleString(locale(), { day: 'numeric', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' }) : '—';
};

const P = {
  x: '<path d="M18 6 6 18M6 6l12 12"/>',
  file: '<path d="M14 3H6a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V9z"/><path d="M14 3v6h6"/>',
  upload: '<path d="M12 16V4M7 9l5-5 5 5M4 20h16"/>',
  download: '<path d="M12 4v12M7 11l5 5 5-5M4 20h16"/>',
  chev: '<path d="m9 6 6 6-6 6"/>',
  down: '<path d="m6 9 6 6 6-6"/>',
  back: '<path d="m15 6-6 6 6 6"/>',
  search: '<circle cx="11" cy="11" r="7"/><path d="m20 20-3.5-3.5"/>',
  user: '<circle cx="12" cy="8" r="4"/><path d="M4 21a8 8 0 0 1 16 0"/>',
  plus: '<path d="M12 5v14M5 12h14"/>',
  chart: '<path d="M4 20V10M10 20V4M16 20v-7M22 20H2"/>',
  eye: '<path d="M2 12s3.6-7 10-7 10 7 10 7-3.6 7-10 7S2 12 2 12z"/><circle cx="12" cy="12" r="3"/>',
  grid: '<rect x="3" y="3" width="7" height="7" rx="1.5"/><rect x="14" y="3" width="7" height="7" rx="1.5"/><rect x="3" y="14" width="7" height="7" rx="1.5"/><rect x="14" y="14" width="7" height="7" rx="1.5"/>',
  shield: '<path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/><path d="m9 12 2 2 4-4"/>',
  out: '<path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4M16 17l5-5-5-5M21 12H9"/>',
  globe: '<circle cx="12" cy="12" r="9"/><path d="M3 12h18M12 3a14 14 0 0 1 0 18M12 3a14 14 0 0 0 0 18"/>',
  alert: '<path d="M12 9v4M12 17h.01"/><path d="M10.3 3.9 1.8 18a2 2 0 0 0 1.7 3h17a2 2 0 0 0 1.7-3L13.7 3.9a2 2 0 0 0-3.4 0z"/>',
  check: '<circle cx="12" cy="12" r="9"/><path d="m8 12 3 3 5-6"/>',
  clock: '<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/>',
  moon: '<path d="M21 13A9 9 0 1 1 11 3a7 7 0 0 0 10 10z"/>',
  sun: '<circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4.9 4.9l1.4 1.4M17.7 17.7l1.4 1.4M2 12h2M20 12h2M4.9 19.1l1.4-1.4M17.7 6.3l1.4-1.4"/>',
};
export const icon = (name, size = 20) =>
  `<svg class="ic" width="${size}" height="${size}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${P[name] || ''}</svg>`;

let toastTimer;
export function toast(msg, tone = '') {
  let el = $('#toast');
  if (!el) { el = document.createElement('div'); el.id = 'toast'; el.setAttribute('role', 'status'); document.body.append(el); }
  el.className = `toast show ${tone}`;
  el.textContent = msg;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { el.className = 'toast'; }, 3400);
}

// Promise<boolean>. Apple-style alert.
export function confirmBox(title, body = '', { ok = tr('OK', 'ตกลง'), danger = false } = {}) {
  return new Promise((resolve) => {
    const back = document.createElement('div');
    back.className = 'modal-back';
    back.innerHTML = `<div class="modal" role="alertdialog" aria-modal="true"><h2>${esc(title)}</h2><p>${esc(body)}</p>
      <div class="actions"><button class="btn" data-v="0">${tr('Cancel', 'ยกเลิก')}</button><button class="btn ${danger ? 'danger' : 'primary'}" data-v="1">${esc(ok)}</button></div></div>`;
    const done = (v) => { back.remove(); document.removeEventListener('keydown', key); resolve(v); };
    const key = (e) => { if (e.key === 'Escape') done(false); };
    back.addEventListener('click', (e) => { if (e.target === back) done(false); const b = e.target.closest('[data-v]'); if (b) done(b.dataset.v === '1'); });
    document.addEventListener('keydown', key);
    document.body.append(back);
    back.querySelector('[data-v="1"]').focus();
  });
}

const sk = (w, h, extra = '') => `<div class="skel" style="width:${w};height:${h}px;${extra}"></div>`;
export function skeleton(rows = 5) {
  return `<div aria-busy="true" aria-label="Loading">${sk('45%', 30, 'margin:8px 0 10px')}${sk('30%', 14, 'margin-bottom:22px')}
    <div class="group">${Array.from({ length: rows }, (_, i) => `<div class="row"><span class="grow" style="display:grid;gap:7px">${sk(`${62 - (i % 3) * 12}%`, 13)}${sk(`${34 + (i % 2) * 10}%`, 10)}</span>${sk('70px', 14, 'flex:none')}</div>`).join('')}</div></div>`;
}
