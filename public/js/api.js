// Every read and write goes to the Apps Script web app (apps-script/Code.gs), which runs as the sheet owner.
import { SCRIPT_URL } from './config.js';
import { idToken } from './firebase.js';
import { tr } from './i18n.js';
import { dateTimeFmt } from './ui.js';

const MAX_MB = 25; // keep in step with MAX_BYTES in Code.gs
const TTL = 60e3;
const cache = new Map(); // reads only: key → { t, p }

// Server busy (too many requests at once, or a save waited too long for the lock): nothing was saved, so the same
// call is tried again, up to twice, after 1.5 s and 3 s. Errors in English or Thai (the script owner's language).
const BUSY = /lock timeout|หมดเวลาการล็อก|simultaneous|พร้อมกันมากเกินไป|invoked too many|^BUSY/i;

export async function call(action, body = {}) {
  for (let tries = 0; ; tries++) {
    try {
      return await once(action, body);
    } catch (e) {
      if (tries >= 2 || !BUSY.test(e.message)) throw e;
      await new Promise((r) => { setTimeout(r, 1500 * (tries + 1)); });
    }
  }
}

async function once(action, body) {
  // text/plain keeps this a "simple" CORS request (Apps Script can't answer preflights).
  const res = await fetch(SCRIPT_URL, {
    method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify({ ...body, action, idToken: await idToken() }),
  });
  const text = await res.text();
  let out;
  try { out = JSON.parse(text); } catch { throw new Error(tr('BUSY: the server is busy — please try again', 'BUSY: เซิร์ฟเวอร์ไม่ว่าง — กรุณาลองใหม่')); } // Google's "too many" page
  if (!out.ok) throw Object.assign(new Error(message(out.error)), { code: out.error });
  return out;
}

// Save conflicts (v2.8) come back as codes: someone else saved after this page loaded, so nothing was written.
// Every other error is shown as the server wrote it.
function message(err) {
  if (!err) return 'Request failed';
  if (err === 'CONFLICT_USERS') return tr('Someone else changed the user list after this page loaded, so nothing was saved. Reload the page and make your changes again.', 'มีคนอื่นแก้ไขรายชื่อผู้ใช้หลังจากเปิดหน้านี้ จึงยังไม่ได้บันทึก กรุณาโหลดหน้าใหม่แล้วแก้ไขอีกครั้ง');
  if (err.startsWith('CONFLICT:')) {
    let c = {};
    try { c = JSON.parse(err.slice(9)); } catch { /* shown without names */ }
    return tr('{who} saved this at {when}, after this page loaded, so your changes were not saved. Copy what you typed, reload the page, then save again.', '{who} บันทึกรายการนี้เมื่อ {when} หลังจากเปิดหน้านี้ จึงยังไม่ได้บันทึกการแก้ไขของคุณ กรุณาคัดลอกสิ่งที่พิมพ์ไว้ โหลดหน้าใหม่ แล้วบันทึกอีกครั้ง', { who: c.by || tr('Someone', 'มีคน'), when: c.at ? dateTimeFmt(c.at) : '—' });
  }
  return err;
}

// Same read again within a minute → the same answer, no network. Views mutate the answer after a save.
// Every answer is also kept on this device (per signed-in email) so the next visit can show it at once: peek().
export function cached(action, body = {}) {
  const k = action + JSON.stringify(body);
  const c = cache.get(k);
  if (c && Date.now() - c.t < TTL) return c.p;
  const p = call(action, body).then((v) => { keep(k, v); return v; }).catch((e) => { cache.delete(k); throw e; });
  cache.set(k, { t: Date.now(), p });
  return p;
}
export const forget = (action) => { for (const k of cache.keys()) if (k.startsWith(action)) cache.delete(k); };

/* ---------- copies kept on this device ---------- */
let who = '';
export const setWho = (email) => { who = email || ''; };
const sk = (k) => `swr:${who}:${k}`;
function keep(k, v) { try { localStorage.setItem(sk(k), JSON.stringify(v)); } catch {} }
// The last answer this device saw, or null. Views show it first, then the fresh answer from cached().
export function peek(action, body = {}) { try { return JSON.parse(localStorage.getItem(sk(action + JSON.stringify(body))) || 'null'); } catch { return null; } }
// After a view changed an answer in place (a save), keep the changed copy.
export const remember = (action, body, v) => keep(action + JSON.stringify(body), v);
export function clearKept() { try { Object.keys(localStorage).filter((k) => k.startsWith('swr:')).forEach((k) => localStorage.removeItem(k)); } catch {} }

const MIME = {
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  pdf: 'application/pdf',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  xls: 'application/vnd.ms-excel',
};
const EXCEL = ['xlsx', 'xls'];
const toBase64 = (blob) => new Promise((res, rej) => {
  const r = new FileReader();
  r.onload = () => res(String(r.result).split(',')[1]);
  r.onerror = rej;
  r.readAsDataURL(blob);
});

// Upload one file straight to Drive. only = 'docx' | 'pdf' | 'excel' (.xlsx or .xls) | 'any' (any of the four) |
// '' (.docx or PDF). ref = the project; org = the organization for purpose 'form' (budget-form feedback).
// Resolves { id, name }.
export async function upload(file, { ref, org, purpose, only = '' }) {
  const kind = (String(file.name).match(/\.([a-z0-9]+)$/i) || [])[1]?.toLowerCase();
  const fits = only === 'any' ? !!MIME[kind] : only === 'excel' ? EXCEL.includes(kind) : MIME[kind] && !EXCEL.includes(kind) && (!only || kind === only);
  if (!fits) {
    throw new Error(only === 'docx' ? tr('Only a Word file (.docx) can go here', 'ช่องนี้รับเฉพาะไฟล์ Word (.docx)')
      : only === 'pdf' ? tr('Only a PDF file can go here', 'ช่องนี้รับเฉพาะไฟล์ PDF')
        : only === 'excel' ? tr('Only an Excel file (.xlsx or .xls) can go here', 'ช่องนี้รับเฉพาะไฟล์ Excel (.xlsx หรือ .xls)')
          : only === 'any' ? tr('Only .xlsx, .xls, .docx or PDF files', 'รับเฉพาะไฟล์ .xlsx, .xls, .docx หรือ PDF')
            : tr('Only .docx or PDF files', 'รับเฉพาะไฟล์ .docx หรือ PDF'));
  }
  if (file.size > MAX_MB * 1048576) throw new Error(tr('File is larger than {mb} MB', 'ไฟล์ใหญ่เกิน {mb} MB', { mb: MAX_MB }));
  const out = await call('upload', { ref, org, purpose, only: only === 'any' ? '' : only, name: file.name, mime: MIME[kind], data: await toBase64(file) });
  return out.file;
}

// Drive links: view page and direct download.
const idOf = (u) => (String(u || '').match(/\/d\/([\w-]+)|[?&]id=([\w-]+)/) || []).slice(1).find(Boolean) || '';
const fileId = (u) => (typeof u === 'string' ? idOf(u) : u?.id || '');
export const viewUrl = (u) => { const id = fileId(u); return id ? `https://drive.google.com/file/d/${encodeURIComponent(id)}/view` : ''; };
// Drive's own viewer (PDF and .docx), shown inside the site by openPreview() in project.js.
export const previewUrl = (u) => { const id = fileId(u); return id ? `https://drive.google.com/file/d/${encodeURIComponent(id)}/preview` : ''; };
export const downloadUrl = (u) => { const id = fileId(u); return id ? `https://drive.google.com/uc?export=download&id=${encodeURIComponent(id)}` : ''; };
