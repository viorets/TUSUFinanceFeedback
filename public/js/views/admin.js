// Admin console: users and roles (with the organizations each person is in charge of), organizations removed on
// the pitching site, pitching phases (On / Off, applied by Save), data export (Excel / CSV), and (Admins only) the
// Pre-pitch edit log. Admins and Chiefs; Data team staff see the edit log only. The server checks too.
import { tr } from '../i18n.js';
import { $, esc, icon, toast, dateTimeFmt, moneyBox, parseMoney, liveMoney, errorBox, ROLE_LABEL } from '../ui.js';
import { call, cached, forget, peek, remember } from '../api.js';
import { ORGS } from '../masterdata.js';
import { PENDING } from '../project.js';

const ROLES = ['Admin', 'Chief', 'Staff', 'Data team staff'];
const PHASE_LIST = [['prepitch', () => tr('Pre-pitch', 'ก่อนการนำเสนอ')], ['pitching', () => tr('Pitching', 'การนำเสนอ')], ['resubmit', () => tr('Re-submission', 'การส่งแก้ไข')]];
const roleLabel = ROLE_LABEL;
// SheetJS builds the Excel file in the browser; loaded only when an Excel export is saved.
const XLSX_URL = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
let st = null; // { users (working copy), saved (JSON), phases (draft), phasesSaved (JSON), save: { users, phases } }

const usersDirty = () => !!st?.users && JSON.stringify(st.users) !== st.saved;
const phasesDirty = () => !!st?.phases && JSON.stringify(st.phases) !== st.phasesSaved;
export const isDirty = () => usersDirty() || phasesDirty();
// Unsaved-changes window: Save and leave saves users and phases.
export const canSaveAll = () => isDirty();
export async function saveAll() {
  if (usersDirty()) await st.save.users();
  if (phasesDirty()) await st.save.phases();
}

// This device's last copy shows at once; the fresh answer repaints unless someone is mid-edit.
export async function render(page, ctx) {
  const kept = peek('admin');
  if (kept) draw(page, ctx, kept);
  try {
    const fresh = await cached('admin');
    const busy = isDirty() || (page.contains(document.activeElement) && document.activeElement.matches('input, select'));
    if (!kept || (!busy && JSON.stringify(fresh) !== JSON.stringify(kept))) draw(page, ctx, fresh);
  } catch (e) {
    if (!kept) throw e;
    toast(e.message, 'bad');
  }
}

function draw(page, ctx, d) {
  if (d.logOnly) { drawLog(page, ctx, d); return; }
  // The owner is always Admin; listed first so a name and organizations can be set for them too.
  const list = d.users.map((u) => ({ ...u }));
  if (!list.some((u) => u.email === d.owner)) list.unshift({ email: d.owner, role: 'Admin', orgs: [], name: '' });
  const chief = ctx.me.role !== 'Admin';
  const phases = { prepitch: true, pitching: true, resubmit: true, ...(d.phases || {}) };
  st = { users: list, saved: JSON.stringify(list), savedAdmins: new Set(list.filter((u) => u.role === 'Admin').map((u) => u.email)),
    phases, phasesSaved: JSON.stringify(phases), save: {} };
  const name = (ref) => ORGS.find((o) => o.code === ref)?.en || '';

  page.innerHTML = `<h1 class="grad">${tr('Admin console', 'คอนโซลผู้ดูแล')}</h1>
    <h2 class="section-title">${tr('Users and roles', 'ผู้ใช้และสิทธิ์')}</h2>
    <div class="group scroll-x"><table class="grid-table edit"><thead><tr><th>${tr('Name shown on the site', 'ชื่อที่แสดง')}</th><th>${tr('Gmail address', 'อีเมล Gmail')}</th><th>${tr('Role', 'สิทธิ์')}</th>
      <th>${tr('In charge of (RefCodes, comma separated)', 'รับผิดชอบองค์กร (รหัส คั่นด้วยจุลภาค)')}</th><th></th></tr></thead>
      <tbody id="users"></tbody></table>
      <div class="table-actions"><button class="btn" id="add">${icon('plus', 16)}${tr('Add user', 'เพิ่มผู้ใช้')}</button><span class="grow"></span>
        <button class="btn primary" id="save">${tr('Save users', 'บันทึกผู้ใช้')}</button></div></div>
    <p class="foot">${tr('Admin: everything, including the edit log. Chief: everything except the edit log; cannot add, change or remove Admins. Staff: edit organization data, no admin console. Data team staff: as Staff, plus Control and the edit log. Emails not listed, or without a role, cannot open the site.', 'ผู้ดูแล: ทำได้ทุกอย่าง รวมถึงดูบันทึกการแก้ไข หัวหน้า: ทำได้ทุกอย่างยกเว้นดูบันทึกการแก้ไข และเพิ่ม เปลี่ยน หรือลบผู้ดูแลไม่ได้ เจ้าหน้าที่: แก้ไขข้อมูลองค์กรได้ แต่เข้าคอนโซลผู้ดูแลไม่ได้ เจ้าหน้าที่ทีมข้อมูล: เหมือนเจ้าหน้าที่ และจัดการหน้าควบคุมและดูบันทึกการแก้ไขได้ อีเมลที่ไม่อยู่ในรายชื่อหรือยังไม่มีสิทธิ์จะเข้าเว็บไซต์ไม่ได้')}</p>

    <h2 class="section-title">${tr('Pitching phases', 'ช่วงการเสนองบ')}</h2>
    <div class="group" id="phases">${PHASE_LIST.map(([k, label]) => `<label class="row phase-row"><span class="grow">${esc(label())}</span>
      <span class="muted small" data-state="${k}">${phases[k] ? tr('On', 'เปิด') : esc(PENDING())}</span>
      <input type="checkbox" data-phase="${k}" ${phases[k] ? 'checked' : ''} aria-label="${esc(label())}"></label>`).join('')}
      <div class="table-actions"><span class="grow muted small" id="phasesNote"></span><button class="btn primary" id="savePhases" disabled>${tr('Save', 'บันทึก')}</button></div></div>
    <p class="foot">${tr('A phase that is off shows "Pending Phase" on every organization page, and nothing in it can be changed. Changes apply only when you press Save.', 'ช่วงที่ปิดอยู่จะแสดง "รอเปิดช่วง" ในทุกหน้าองค์กร และแก้ไขข้อมูลในช่วงนั้นไม่ได้ การเปลี่ยนแปลงจะมีผลเมื่อกดบันทึกเท่านั้น')}</p>

    <h2 class="section-title">${tr('Export data', 'ส่งออกข้อมูล')}</h2>
    <div class="group" id="export">
      <div class="row"><span class="grow">${tr('Organizations and projects', 'องค์กรและโครงการ')}</span><span class="muted small">${tr('Always included', 'รวมเสมอ')}</span></div>
      ${PHASE_LIST.map(([k, label]) => `<label class="row phase-row"><span class="grow">${esc(label())}</span><input type="checkbox" data-ex="${k}" checked aria-label="${esc(label())}"></label>`).join('')}
      <div class="row"><span class="grow">${tr('Format', 'รูปแบบไฟล์')}</span>
        <div class="seg" role="radiogroup"><label><input type="radio" name="exfmt" value="xlsx" checked><span>Excel</span></label><label><input type="radio" name="exfmt" value="csv"><span>CSV</span></label></div></div>
      <div class="table-actions"><span class="grow muted small">${tr('Nothing downloads until you press Save.', 'จะดาวน์โหลดเมื่อกดบันทึกเท่านั้น')}</span><button class="btn primary" id="saveExport">${icon('download', 16)}${tr('Save', 'บันทึก')}</button></div></div>
    <p class="foot">${tr('Excel: one sheet for organizations, one for projects and one per included phase. CSV: one row per project with every included column. Organizations removed on the pitching site are left out.', 'Excel: ชีตองค์กร ชีตโครงการ และหนึ่งชีตต่อช่วงที่เลือก CSV: หนึ่งแถวต่อโครงการ พร้อมทุกคอลัมน์ที่เลือก ไม่รวมองค์กรที่ถูกนำออกในเว็บไซต์เสนองบ')}</p>

    <h2 class="section-title">${tr('Allocated budget', 'กรอบงบประมาณ')}</h2>
    <div class="group alloc"><span class="muted">฿</span><input type="text" inputmode="decimal" id="alloc" value="${esc(moneyBox(d.allocated))}" placeholder="0.00" aria-label="${esc(tr('Allocated budget', 'กรอบงบประมาณ'))}">
      <button class="btn primary" id="saveAlloc">${tr('Save', 'บันทึก')}</button></div>
    <p class="foot">${tr('The Dashboard compares the total proposed budget and the final budget with this amount. Admins and Chiefs can change it.', 'แดชบอร์ดเปรียบเทียบงบประมาณที่เสนอทั้งหมดและงบประมาณอนุมัติขั้นสุดท้ายกับจำนวนนี้ ผู้ดูแลและหัวหน้าแก้ไขได้')}</p>

    <h2 class="section-title">${tr('Organizations removed on the pitching site', 'องค์กรที่ถูกนำออกในเว็บไซต์เสนองบ')}</h2>
    <div class="group">${d.removed.length ? d.removed.map((r) => `<a class="row" href="#/org/${encodeURIComponent(r)}"><span class="mono">${esc(r)}</span><span class="grow">${esc(name(r))}</span>${icon('chev', 16)}</a>`).join('')
    : `<div class="row muted">${tr('None', 'ไม่มี')}</div>`}</div>
    <p class="foot">${tr('Hidden from the public list.', 'ไม่แสดงในรายชื่อองค์กร')}</p>

    ${d.log ? logSection() : ''}`;

  const paintUsers = () => {
    $('#users', page).innerHTML = st.users.map((u, i) => {
      const owner = u.email === d.owner;
      const locked = owner || (chief && st.savedAdmins.has(u.email));
      return `<tr data-i="${i}">
      <td><input type="text" data-f="name" value="${esc(u.name || '')}" maxlength="80" placeholder="${esc(tr('e.g. Mint (Finance)', 'เช่น มิ้นท์ (ฝ่ายการเงิน)'))}"></td>
      <td><input type="email" data-f="email" value="${esc(u.email)}" placeholder="name@gmail.com" ${locked ? 'readonly' : ''}></td>
      <td>${locked ? `<span class="muted small">${owner ? tr('Admin · script owner', 'ผู้ดูแล · เจ้าของสคริปต์') : roleLabel('Admin')}</span>`
    : `<select data-f="role">${u.role ? '' : `<option value="" selected>${tr('Choose a role', 'เลือกสิทธิ์')}</option>`}${ROLES.filter((r) => !chief || r !== 'Admin').map((r) => `<option value="${r}" ${u.role === r ? 'selected' : ''}>${esc(roleLabel(r))}</option>`).join('')}</select>`}</td>
      <td><input type="text" data-f="orgs" value="${esc(u.orgs.join(', '))}" placeholder="RS-ART-03, RS-OTH-01" list="refs"></td>
      <td>${locked ? '' : `<button class="icon-btn" data-rm title="${esc(tr('Remove', 'ลบ'))}">${icon('x', 16)}</button>`}</td></tr>`;
    }).join('')
      + `<datalist id="refs">${ORGS.map((o) => `<option value="${esc(o.code)}">${esc(o.en)}</option>`).join('')}</datalist>`;
  };
  paintUsers();
  if (d.log) paintLogRows(page, ctx, d);

  // Phase switches: a draft until Save (v2.4). Save stays off until something changes.
  const phaseState = () => {
    PHASE_LIST.forEach(([k]) => { $(`[data-state="${k}"]`, page).textContent = st.phases[k] ? tr('On', 'เปิด') : PENDING(); });
    $('#savePhases', page).disabled = !phasesDirty();
    $('#phasesNote', page).textContent = phasesDirty() ? tr('Not saved yet', 'ยังไม่ได้บันทึก') : '';
  };
  $('#phases', page).onchange = (e) => {
    const k = e.target.dataset.phase;
    if (!k) return;
    st.phases[k] = e.target.checked;
    phaseState();
  };
  st.save.phases = async () => {
    const out = await call('savePhases', { phases: st.phases });
    st.phases = { ...out.phases };
    st.phasesSaved = JSON.stringify(st.phases);
    d.phases = out.phases;
    remember('admin', {}, d); forget('admin');
    if (ctx.boot) { ctx.boot.phases = out.phases; remember('boot', {}, ctx.boot); }
    forget('boot'); forget('org');
  };
  $('#savePhases', page).onclick = async (e) => {
    e.target.disabled = true;
    try {
      await st.save.phases();
      toast(tr('Phases saved', 'บันทึกช่วงแล้ว'));
    } catch (err) {
      await errorBox(err.message); // the draft stays on screen
    }
    phaseState();
  };

  // Export: download on Save only.
  $('#saveExport', page).onclick = async (e) => {
    const want = Object.fromEntries(PHASE_LIST.map(([k]) => [k, $(`[data-ex="${k}"]`, page).checked]));
    const fmt = $('input[name=exfmt]:checked', page).value;
    e.target.disabled = true;
    try {
      const data = await call('exportData', { phases: want });
      await (fmt === 'csv' ? saveCsv(data) : saveXlsx(data));
      toast(tr('Export saved', 'บันทึกไฟล์ส่งออกแล้ว'));
    } catch (err) {
      await errorBox(err.message, tr('Couldn’t export', 'ส่งออกไม่สำเร็จ'));
    }
    e.target.disabled = false;
  };

  $('#alloc', page).oninput = (e) => liveMoney(e.target);
  $('#saveAlloc', page).onclick = async (e) => {
    e.target.disabled = true;
    try {
      const out = await call('saveAllocated', { allocated: parseMoney($('#alloc', page).value) });
      $('#alloc', page).value = moneyBox(out.allocated);
      d.allocated = out.allocated;
      remember('admin', {}, d); forget('admin');
      forget('dashboard');
      toast(tr('Allocated budget saved', 'บันทึกกรอบงบประมาณแล้ว'));
    } catch (err) {
      toast(err.message, 'bad');
    }
    e.target.disabled = false;
  };

  const users = $('#users', page);
  users.oninput = users.onchange = (e) => {
    const i = Number(e.target.closest('[data-i]')?.dataset.i); const f = e.target.dataset.f;
    if (!f) return;
    st.users[i][f] = f === 'orgs' ? e.target.value.split(/[,\s]+/).map((x) => x.trim().toUpperCase()).filter(Boolean) : e.target.value.trim();
  };
  users.onclick = (e) => {
    if (!e.target.closest('[data-rm]')) return;
    st.users.splice(Number(e.target.closest('[data-i]').dataset.i), 1);
    paintUsers();
  };
  $('#add', page).onclick = () => { st.users.push({ email: '', role: 'Staff', orgs: [], name: '' }); paintUsers(); $('#users tr:last-child input', page)?.focus(); };
  st.save.users = async () => {
    const out = await call('saveUsers', { users: st.users.filter((u) => u.email) });
    st.users = out.users.map((u) => ({ ...u }));
    st.saved = JSON.stringify(st.users);
    d.users = out.users;
    remember('admin', {}, d); forget('admin');
    if (ctx.boot) { ctx.boot.names = Object.fromEntries(out.users.filter((u) => u.name).map((u) => [u.email, u.name])); remember('boot', {}, ctx.boot); }
    if (page.isConnected) paintUsers();
    forget('boot'); // persons in charge on the list
  };
  $('#save', page).onclick = async (e) => {
    e.target.disabled = true;
    try {
      await st.save.users();
      toast(tr('Users saved', 'บันทึกผู้ใช้แล้ว'));
    } catch (err) {
      toast(err.message, 'bad');
    }
    e.target.disabled = false;
  };
}

// Data team staff: the edit log only.
function drawLog(page, ctx, d) {
  st = null;
  page.innerHTML = `<h1 class="grad">${tr('Admin console', 'คอนโซลผู้ดูแล')}</h1>${logSection()}`;
  paintLogRows(page, ctx, d);
}

// Edit logs (v2.6): Pre-pitch (every changed field) and Re-submission (decisions, case close / open).
let logKind = 'pre';
const logSection = () => `<h2 class="section-title">${tr('Edit log', 'บันทึกการแก้ไข')} <span class="muted">${tr('latest 500', '500 รายการล่าสุด')}</span></h2>
    <div class="toolbar"><div class="seg" role="radiogroup" id="logKind" aria-label="${esc(tr('Log', 'บันทึก'))}">
      <label><input type="radio" name="logKind" value="pre" ${logKind === 'pre' ? 'checked' : ''}><span>${tr('Pre-pitch', 'ก่อนการนำเสนอ')}</span></label>
      <label><input type="radio" name="logKind" value="resub" ${logKind === 'resub' ? 'checked' : ''}><span>${tr('Re-submission', 'การส่งแก้ไข')}</span></label></div>
      <input type="search" id="lq" placeholder="${esc(tr('Filter by RefCode, person or field', 'กรองตามรหัส ผู้แก้ไข หรือช่อง'))}"></div>
    <div class="group scroll-x"><table class="grid-table static"><thead><tr><th>${tr('When', 'เวลา')}</th><th>${tr('By', 'โดย')}</th><th>${tr('Project', 'โครงการ')}</th>
      <th>${tr('Field', 'ช่อง')}</th><th>${tr('Old', 'เดิม')}</th><th>${tr('New', 'ใหม่')}</th></tr></thead><tbody id="log"></tbody></table></div>`;

function paintLogRows(page, ctx, d) {
  const cell = (s) => { const f = /^\{"id"/.test(s) ? (() => { try { return JSON.parse(s).name; } catch { return s; } })() : s; return esc(f); };
  const paint = () => {
    const q = $('#lq', page).value.trim().toLowerCase();
    const list = (logKind === 'resub' ? d.resubLog : d.log) || [];
    const rows = list.filter((r) => !q || `${r.ref} ${r.org} ${r.by} ${ctx.nameOf(r.by)} ${r.field}`.toLowerCase().includes(q));
    // A case row has no project: it shows the organization's RefCode.
    $('#log', page).innerHTML = rows.map((r) => `<tr><td class="small">${dateTimeFmt(r.at)}</td><td class="small">${esc(ctx.nameOf(r.by))}</td><td class="mono">${esc(r.ref || r.org)}</td>
      <td>${esc(r.field)}</td><td class="small">${cell(r.old)}</td><td class="small">${cell(r.now)}</td></tr>`).join('')
      || `<tr><td colspan="6" class="muted">${tr('No edits yet', 'ยังไม่มีการแก้ไข')}</td></tr>`;
  };
  paint();
  $('#lq', page).oninput = paint;
  $('#logKind', page).onchange = (e) => { logKind = e.target.value; paint(); };
}

/* ---------- export files ---------- */

const PHASE_SHEETS = [['prepitch', 'Pre-pitch'], ['pitching', 'Pitching'], ['resubmit', 'Re-submission']];
const stamp = () => new Date().toISOString().slice(0, 10);
function download(blob, name) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = name;
  document.body.append(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 1000);
}

// One row per project: organization columns, project columns, then each included phase's columns (prefixed).
function flatRows(d) {
  const orgs = Object.fromEntries(d.orgs.map((o) => [o['Org Ref'], o]));
  const phase = PHASE_SHEETS.filter(([k]) => d[k]).map(([k, label]) => [label, Object.fromEntries(d[k].map((r) => [r['Project Ref'], r]))]);
  return d.projects.map((p) => {
    const row = {};
    Object.entries(orgs[p['Org Ref']] || { 'Org Ref': p['Org Ref'] }).forEach(([k, v]) => { row[k] = v; });
    Object.entries(p).forEach(([k, v]) => { if (k !== 'Org Ref') row[k] = v; });
    phase.forEach(([label, by]) => Object.entries(by[p['Project Ref']] || {}).forEach(([k, v]) => { if (k !== 'Project Ref') row[`${label}: ${k}`] = v; }));
    return row;
  });
}

function saveCsv(d) {
  const rows = flatRows(d);
  const head = [...new Set(rows.flatMap((r) => Object.keys(r)))];
  const q = (v) => { const t = v == null ? '' : String(v); return /[",\n\r]/.test(t) ? `"${t.replace(/"/g, '""')}"` : t; };
  const text = [head.map(q).join(','), ...rows.map((r) => head.map((h) => q(r[h])).join(','))].join('\r\n');
  download(new Blob([`﻿${text}`], { type: 'text/csv;charset=utf-8' }), `pitch-feedback-${stamp()}.csv`); // BOM: Excel reads Thai as UTF-8
}

let xlsxLoad = null;
const loadXlsx = () => xlsxLoad || (xlsxLoad = new Promise((res, rej) => {
  if (window.XLSX) { res(window.XLSX); return; }
  const s = document.createElement('script');
  s.src = XLSX_URL; s.crossOrigin = 'anonymous';
  s.onload = () => res(window.XLSX);
  s.onerror = () => { xlsxLoad = null; rej(new Error(tr('Couldn’t load the Excel tool. Try CSV, or try again.', 'โหลดเครื่องมือ Excel ไม่ได้ ลองใช้ CSV หรือลองใหม่'))); };
  document.head.append(s);
}));

async function saveXlsx(d) {
  const X = await loadXlsx();
  const wb = X.utils.book_new();
  X.utils.book_append_sheet(wb, X.utils.json_to_sheet(d.orgs), 'Organizations');
  X.utils.book_append_sheet(wb, X.utils.json_to_sheet(d.projects), 'Projects');
  PHASE_SHEETS.forEach(([k, label]) => { if (d[k]) X.utils.book_append_sheet(wb, X.utils.json_to_sheet(d[k]), label); });
  const out = X.write(wb, { bookType: 'xlsx', type: 'array' });
  download(new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), `pitch-feedback-${stamp()}.xlsx`);
}
