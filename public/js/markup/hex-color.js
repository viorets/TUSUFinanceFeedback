// HEX entry for a colour box: a text field beside an <input type=color> that takes #abc, abc, #aabbcc or aabbcc.
// A valid value sets the colour box and fires its 'input' and 'change' events, so the page's own listeners run.
// The colour box writes back into the field on 'input'. Code that sets the colour box's value directly calls the
// returned sync(), because setting .value fires no event.
export const normHex = (v) => {
  const m = /^#?([0-9a-f]{3}|[0-9a-f]{6})$/i.exec(String(v).trim());
  if (!m) return null;
  const h = m[1].length === 3 ? [...m[1]].map((c) => c + c).join('') : m[1];
  return `#${h.toLowerCase()}`;
};

// HEX <-> HSL for scrubbing.
function toHsl(hex) {
  const n = parseInt(hex.slice(1), 16); const r = ((n >> 16) & 255) / 255, g = ((n >> 8) & 255) / 255, b = (n & 255) / 255;
  const mx = Math.max(r, g, b), mn = Math.min(r, g, b), l = (mx + mn) / 2, d = mx - mn;
  if (!d) return [0, 0, l];
  const s = d / (1 - Math.abs(2 * l - 1));
  const h = mx === r ? ((g - b) / d) % 6 : mx === g ? (b - r) / d + 2 : (r - g) / d + 4;
  return [(h * 60 + 360) % 360, s, l];
}
function toHex([h, s, l]) {
  const c = (1 - Math.abs(2 * l - 1)) * s, x = c * (1 - Math.abs(((h / 60) % 2) - 1)), m = l - c / 2;
  const [r, g, b] = h < 60 ? [c, x, 0] : h < 120 ? [x, c, 0] : h < 180 ? [0, c, x] : h < 240 ? [0, x, c] : h < 300 ? [x, 0, c] : [c, 0, x];
  return `#${[r, g, b].map((v) => Math.round((v + m) * 255).toString(16).padStart(2, '0')).join('')}`;
}
// Scrub: press on a colour box and drag. Left/right shifts the hue, up/down makes it lighter/darker; 'input' fires as
// it changes and 'change' on release. A press without a drag still opens the picker.
function scrub(input) {
  input.title = 'Click to pick · drag ← → for hue, ↑ ↓ for lighter or darker';
  let start = null; let dragged = false;
  input.addEventListener('pointerdown', (e) => {
    if (e.button !== 0) return;
    start = { x: e.clientX, y: e.clientY, hsl: toHsl(input.value), id: e.pointerId }; dragged = false;
  });
  input.addEventListener('pointermove', (e) => {
    if (!start || e.pointerId !== start.id) return;
    const dx = e.clientX - start.x, dy = e.clientY - start.y;
    if (!dragged && Math.hypot(dx, dy) < 5) return;
    if (!dragged) { dragged = true; input.setPointerCapture(e.pointerId); }
    const [h, s, l] = start.hsl;
    const v = toHex([(h + dx * 0.6 + 3600) % 360, s, Math.min(0.97, Math.max(0.03, l - dy * 0.004))]);
    if (v !== input.value) { input.value = v; input.dispatchEvent(new Event('input', { bubbles: true })); }
  });
  const end = (e) => {
    if (!start || (e && e.pointerId !== start.id)) return;
    if (dragged) input.dispatchEvent(new Event('change', { bubbles: true }));
    start = null;
  };
  input.addEventListener('pointerup', end); input.addEventListener('pointercancel', end);
  // A drag must not open the picker on release.
  input.addEventListener('click', (e) => { if (dragged) { e.preventDefault(); dragged = false; } });
}

export function pairHex(input) {
  scrub(input);
  const f = document.createElement('input');
  f.type = 'text'; f.className = 'hex-in'; f.maxLength = 7; f.spellcheck = false;
  f.autocomplete = 'off'; f.setAttribute('autocapitalize', 'off'); f.setAttribute('aria-label', `${input.getAttribute('aria-label') || input.title || input.closest('label')?.title || 'Colour'}: HEX`); // each one says which colour
  f.placeholder = '#rrggbb';
  input.after(f);
  const sync = () => { f.value = input.value; f.removeAttribute('aria-invalid'); };
  const apply = () => {
    const v = normHex(f.value);
    if (!v) { f.setAttribute('aria-invalid', 'true'); return false; }
    f.removeAttribute('aria-invalid'); f.value = v;
    if (input.value !== v) {
      input.value = v;
      input.dispatchEvent(new Event('input', { bubbles: true }));
      input.dispatchEvent(new Event('change', { bubbles: true }));
    }
    return true;
  };
  f.addEventListener('input', () => { if (normHex(f.value) && f.value.replace('#', '').length === 6) apply(); });
  f.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); if (!apply()) sync(); } });
  f.addEventListener('blur', () => { if (!apply()) sync(); });
  input.addEventListener('input', sync);
  input.addEventListener('change', sync);
  sync();
  return sync;
}
