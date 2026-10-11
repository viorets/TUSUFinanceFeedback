// Control (the script owner and Data team staff; Code.gs checks too):
// 1. Required / Not required for each feedback question (Settings tab, key "Required"). Required questions get a *
//    on the form and decide when a project counts as finished.
// 2. Every line of site text in English and Thai (Texts tab). The list is read from the site's own code: each
//    tr('English', 'ไทย') is one line, keyed by its built-in English. An edit left equal to the built-in line is not
//    stored, so the line follows the code again. Marks like {n} stand for values and must stay. Lines are split into
//    tabs by type (the page they appear on); a line used on several pages shows on each of those tabs.
import { tr, setTexts } from '../i18n.js';
import { $, esc, icon, toast, errorBox } from '../ui.js';
import { call, forget, remember } from '../api.js';
import { QUESTIONS, isReq } from '../project.js';

// Modules whose text is listed, with the page name shown in the table.
const SOURCES = [
  ['js/app.js', () => tr('Top bar and sign-in', 'แถบด้านบนและการลงชื่อเข้าใช้')],
  ['js/views/list.js', () => tr('Organizations', 'องค์กร')],
  ['js/views/org.js', () => tr('Organization page', 'หน้าองค์กร')],
  ['js/views/prepitch.js', () => tr('Pre-pitch', 'ก่อนการนำเสนอ')],
  ['js/views/pitching.js', () => tr('Pitching', 'การนำเสนอ')],
  ['js/views/resubmit.js', () => tr('Re-submission', 'การส่งแก้ไข')],
  ['js/project.js', () => tr('Project cards', 'การ์ดโครงการ')],
  ['js/views/dashboard.js', () => tr('Dashboard', 'แดชบอร์ด')],
  ['js/views/admin.js', () => tr('Admin console', 'คอนโซลผู้ดูแล')],
  ['js/views/control.js', () => tr('Control', 'ควบคุม')],
  ['js/ui.js', () => tr('Windows and messages', 'หน้าต่างและข้อความ')],
  ['js/api.js', () => tr('Uploads', 'การอัปโหลด')],
];
const TR = /\btr\(\s*'((?:[^'\\]|\\.)*)'\s*,\s*'((?:[^'\\]|\\.)*)'/g;
const unq = (x) => x.replace(/\\(.)/g, '$1');
const marks = (x) => (String(x).match(/\{\w+\}/g) || []).sort().join(' ');

let st = null; // { lines: [{ key, th, pages }], edit: { key: { en, th } }, saved: JSON, req, reqSaved, save: { req, texts } }
let tab = ''; // site-text tab: a page name, '' = all (kept while moving around the site)

const reqDirty = () => !!st && JSON.stringify(st.req) !== st.reqSaved;
const textsDirty = () => !!st && JSON.stringify(st.edit) !== st.saved;
export const isDirty = () => reqDirty() || textsDirty();
// Unsaved-changes window: Save and leave saves both parts (throws if a line lost its {marks}).
export const canSaveAll = () => isDirty();
export async function saveAll() {
  if (reqDirty()) await st.save.req();
  if (textsDirty()) await st.save.texts();
}

export async function render(page, ctx) {
  page.innerHTML = `<h1 class="grad">${tr('Control', 'ควบคุม')}</h1>
    <p class="muted">${tr('Owner and Data team staff. Changes show for everyone on their next page load.', 'สำหรับเจ้าของและเจ้าหน้าที่ทีมข้อมูล การเปลี่ยนแปลงจะแสดงให้ทุกคนเห็นเมื่อโหลดหน้าใหม่')}</p>
    <h2 class="section-title">${tr('Required questions', 'คำถามที่จำเป็น')}</h2>
    <div class="group scroll-x"><table class="grid-table static"><thead><tr><th>${tr('Tab', 'แท็บ')}</th><th>${tr('Question', 'คำถาม')}</th><th>${tr('Required', 'จำเป็น')}</th></tr></thead>
      <tbody id="req"></tbody></table>
      <div class="table-actions"><span class="grow muted small">${tr('A project is finished when every required question is answered.', 'โครงการจะครบเมื่อตอบคำถามที่จำเป็นทุกข้อ')}</span><button class="btn primary" id="saveReq">${tr('Save', 'บันทึก')}</button></div></div>
    <h2 class="section-title">${tr('Site text', 'ข้อความบนเว็บไซต์')} <span class="muted" id="count"></span></h2>
    <nav class="seg-tabs text-tabs" id="textTabs" aria-label="${esc(tr('Text type', 'ประเภทข้อความ'))}"></nav>
    <div class="toolbar"><input type="search" id="q" placeholder="${esc(tr('Search English or Thai', 'ค้นหาภาษาอังกฤษหรือภาษาไทย'))}">
      <label class="check"><input type="checkbox" id="changedOnly">${tr('Changed only', 'เฉพาะที่แก้ไข')}</label></div>
    <div class="group scroll-x"><table class="grid-table edit texts"><thead><tr><th>${tr('Page', 'หน้า')}</th><th>English</th><th>ไทย</th><th></th></tr></thead>
      <tbody id="texts"><tr><td colspan="4" class="muted">${tr('Loading…', 'กำลังโหลด…')}</td></tr></tbody></table>
      <div class="table-actions"><span class="grow muted small">${tr('Keep marks like {n} — they stand for numbers and names.', 'คงเครื่องหมายเช่น {n} ไว้ — แทนตัวเลขและชื่อ')}</span><button class="btn primary" id="saveTexts">${tr('Save text', 'บันทึกข้อความ')}</button></div></div>`;

  const texts = ctx.boot?.texts || {};
  const required = { ...(ctx.boot?.required || {}) };
  st = { lines: [], edit: {}, saved: '{}', req: Object.fromEntries(QUESTIONS.map((q) => [q.id, isReq({ required }, q.id)])), reqSaved: '', save: {} };
  st.reqSaved = JSON.stringify(st.req);

  // Required questions
  const tabName = { prepitch: () => tr('Pre-pitch', 'ก่อนการนำเสนอ'), pitching: () => tr('Pitching', 'การนำเสนอ'), resubmit: () => tr('Re-submission', 'การส่งแก้ไข') };
  $('#req', page).innerHTML = QUESTIONS.map((q) => `<tr><td>${esc(tabName[q.tab]())}</td><td>${esc(q.label())}</td>
    <td><label class="check"><input type="checkbox" data-q="${q.id}" ${st.req[q.id] ? 'checked' : ''}>${tr('Required', 'จำเป็น')}</label></td></tr>`).join('');
  $('#req', page).onchange = (e) => { const id = e.target.dataset.q; if (id) st.req[id] = e.target.checked; };
  st.save.req = async () => {
    const out = await call('saveRequired', { required: st.req });
    st.reqSaved = JSON.stringify(st.req);
    if (ctx.boot) { ctx.boot.required = out.required; remember('boot', {}, ctx.boot); }
    forget('boot'); forget('org'); forget('dashboard');
  };
  $('#saveReq', page).onclick = async (e) => {
    e.target.disabled = true;
    try {
      await st.save.req();
      toast(tr('Required questions saved', 'บันทึกคำถามที่จำเป็นแล้ว'));
    } catch (err) {
      toast(err.message, 'bad');
    }
    e.target.disabled = false;
  };

  // Site text: every tr('…', '…') in the site's code.
  const found = new Map();
  await Promise.all(SOURCES.map(async ([path, pageName]) => {
    try {
      const src = await (await fetch(path, { cache: 'no-cache' })).text();
      for (const m of src.matchAll(TR)) {
        const key = unq(m[1]);
        const line = found.get(key) || { key, th: unq(m[2]), pages: [] };
        if (!line.pages.includes(pageName())) line.pages.push(pageName());
        found.set(key, line);
      }
    } catch { /* that page's lines are left out */ }
  }));
  if (!page.isConnected) return;
  st.lines = [...found.values()].sort((a, b) => a.pages[0].localeCompare(b.pages[0]) || a.key.localeCompare(b.key));
  // Lines the Owner changed earlier but no longer in the code stay stored; they are simply not listed.
  st.edit = Object.fromEntries(st.lines.filter((l) => texts[l.key]).map((l) => [l.key, { en: texts[l.key].en || '', th: texts[l.key].th || '' }]));
  st.saved = JSON.stringify(st.edit);

  const value = (l, k) => st.edit[l.key]?.[k] || (k === 'en' ? l.key : l.th);
  const changed = (l) => !!(st.edit[l.key]?.en || st.edit[l.key]?.th);
  // One tab per page that has lines, in SOURCES order, after "All".
  const pageNames = SOURCES.map(([, n]) => n()).filter((n) => st.lines.some((l) => l.pages.includes(n)));
  if (tab && !pageNames.includes(tab)) tab = '';
  const paintTabs = () => {
    const count = (n) => st.lines.filter((l) => !n || l.pages.includes(n)).length;
    $('#textTabs', page).innerHTML = ['', ...pageNames].map((n) => `<a href="#" data-tab="${esc(n)}" class="${n === tab ? 'on' : ''}">${esc(n || tr('All', 'ทั้งหมด'))} <span class="small">${count(n)}</span></a>`).join('');
  };
  const paint = () => {
    const q = $('#q', page).value.trim().toLowerCase();
    const only = $('#changedOnly', page).checked;
    const rows = st.lines.filter((l) => (!tab || l.pages.includes(tab)) && (!only || changed(l)) && (!q || `${l.key} ${l.th} ${value(l, 'en')} ${value(l, 'th')}`.toLowerCase().includes(q)));
    $('#count', page).textContent = tr('{n} lines · {c} changed', '{n} บรรทัด · แก้ไข {c}', { n: st.lines.length, c: st.lines.filter(changed).length });
    $('#texts', page).innerHTML = rows.map((l) => `<tr data-key="${esc(l.key)}" class="${changed(l) ? 'changed' : ''}">
      <td class="small muted">${esc(l.pages.join(', '))}</td>
      <td><textarea rows="1" data-k="en">${esc(value(l, 'en'))}</textarea></td>
      <td><textarea rows="1" data-k="th">${esc(value(l, 'th'))}</textarea></td>
      <td>${changed(l) ? `<button class="icon-btn" data-reset title="${esc(tr('Back to the built-in text', 'กลับเป็นข้อความเดิม'))}">${icon('x', 16)}</button>` : ''}</td></tr>`).join('')
      || `<tr><td colspan="4" class="muted">${tr('No lines match', 'ไม่มีข้อความที่ตรงกัน')}</td></tr>`;
  };
  paintTabs();
  paint();
  $('#textTabs', page).onclick = (e) => {
    const a = e.target.closest('[data-tab]');
    if (!a) return;
    e.preventDefault();
    tab = a.dataset.tab;
    paintTabs();
    paint();
  };
  $('#q', page).oninput = paint;
  $('#changedOnly', page).onchange = paint;
  const body = $('#texts', page);
  body.oninput = (e) => {
    const k = e.target.dataset.k; const key = e.target.closest('[data-key]')?.dataset.key;
    if (!k || key == null) return;
    const line = st.lines.find((l) => l.key === key);
    const builtIn = k === 'en' ? line.key : line.th;
    const v = e.target.value === builtIn ? '' : e.target.value;
    const cur = { en: '', th: '', ...st.edit[key], [k]: v };
    if (cur.en || cur.th) st.edit[key] = cur; else delete st.edit[key];
    e.target.closest('tr').classList.toggle('changed', !!st.edit[key]);
  };
  body.onclick = (e) => {
    const key = e.target.closest('[data-reset]') && e.target.closest('[data-key]')?.dataset.key;
    if (key == null || !key) return;
    delete st.edit[key];
    paint();
  };
  // Every {mark} of the built-in line must still be there.
  const badLines = () => st.lines.filter((l) => st.edit[l.key] && ['en', 'th'].some((k) => st.edit[l.key][k] && marks(st.edit[l.key][k]) !== marks(l.key)));
  const marksMsg = (bad) => tr('Keep the {marks} of these lines: {lines}', 'กรุณาคงเครื่องหมาย {marks} ในข้อความเหล่านี้: {lines}', { marks: '{…}', lines: bad.map((l) => l.key).join(' · ') });
  st.save.texts = async () => {
    const bad = badLines();
    if (bad.length) throw new Error(marksMsg(bad));
    const out = await call('saveTexts', { texts: Object.entries(st.edit).map(([key, v]) => ({ key, en: v.en, th: v.th })) });
    st.saved = JSON.stringify(st.edit);
    setTexts(out.texts);
    if (ctx.boot) { ctx.boot.texts = out.texts; remember('boot', {}, ctx.boot); }
    forget('boot');
  };
  $('#saveTexts', page).onclick = async (e) => {
    const bad = badLines();
    if (bad.length) { await errorBox(marksMsg(bad), tr('Marks missing', 'เครื่องหมายหายไป')); return; }
    e.target.disabled = true;
    try {
      await st.save.texts();
      toast(tr('Site text saved', 'บันทึกข้อความแล้ว'));
    } catch (err) {
      toast(err.message, 'bad');
    }
    e.target.disabled = false;
  };
}
