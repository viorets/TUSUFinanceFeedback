// UI language. English by default; the choice is remembered per device.
// Strings are written inline as tr('English', 'ไทย') so both versions stay side by side.

let lang = 'en';
try { if (localStorage.getItem('lang') === 'th') lang = 'th'; } catch {}
document.documentElement.lang = lang;

export const getLang = () => lang;
export const locale = () => (lang === 'th' ? 'th-TH' : 'en-GB');
export const tr = (en, th) => (lang === 'th' ? th : en);

export function setLang(l) {
  lang = l === 'th' ? 'th' : 'en';
  document.documentElement.lang = lang;
  try { localStorage.setItem('lang', lang); } catch {}
}
