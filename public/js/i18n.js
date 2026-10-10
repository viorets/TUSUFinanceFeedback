// UI language. English by default; the choice is remembered per device.
// Strings are written inline as tr('English', 'ไทย') so both versions stay side by side. A line with values in it
// marks them {name} and passes them last: tr('{n} pending', 'รอ {n}', { n }).
// The Owner can replace any line in the Control tab (Texts tab in our Sheet); the built-in English line is the key.

let lang = 'en';
try { if (localStorage.getItem('lang') === 'th') lang = 'th'; } catch {}
document.documentElement.lang = lang;

let custom = {}; // key (built-in English) → { en, th }; an empty value keeps the built-in line

export const getLang = () => lang;
export const locale = () => (lang === 'th' ? 'th-TH' : 'en-GB');
export function tr(en, th, vars) {
  const c = custom[en];
  const line = lang === 'th' ? (c?.th || th) : (c?.en || en);
  return vars ? line.replace(/\{(\w+)\}/g, (m, k) => (k in vars ? vars[k] : m)) : line;
}
export const setTexts = (t) => { custom = t && typeof t === 'object' ? t : {}; };

export function setLang(l) {
  lang = l === 'th' ? 'th' : 'en';
  document.documentElement.lang = lang;
  try { localStorage.setItem('lang', lang); } catch {}
}
