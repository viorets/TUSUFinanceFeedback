// Admin console: users and roles (with the organizations each person is in charge of), organizations removed on
// the pitching site, pitching phases (On / Off), and (Admins only) the Pre-pitch edit log. Admins and Chiefs;
// the server checks too.
import { tr } from '../i18n.js';
import { $, esc, icon, toast, dateTimeFmt, moneyBox, parseMoney, liveMoney } from '../ui.js';
import { call, cached, forget, peek, remember } from '../api.js';
import { ORGS } from '../masterdata.js';
import { PENDING } from '../project.js';

const ROLES = ['Admin', 'Chief', 'Staff'];
const PHASE_LIST = [['prepitch', () => tr('Pre-pitch', 'ก่อนการนำเสนอ')], ['pitching', () => tr('Pitching', 'การนำเสนอ')], ['resubmit', () => tr('Re-submission', 'การส่งแก้ไข')]];
const roleLabel = (r) => ({ Admin: tr('Admin', 'ผู้ดูแล'), Chief: tr('Chief', 'หัวหน้า'), Staff: tr('Staff', 'เจ้าหน้าที่') }[r] || r);
let st = null; // { users (working copy), saved (JSON) }

export const isDirty = () => !!st && JSON.stringify(st.users) !== st.saved;

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
  // The owner is always Admin; listed first so a name and organizations can be set for them too.
  const list = d.users.map((u) => ({ ...u }));
  if (!list.some((u) => u.email === d.owner)) list.unshift({ email: d.owner, role: 'Admin', orgs: [], name: '' });
  const chief = ctx.me.role !== 'Admin';
  st = { users: list, saved: JSON.stringify(list), savedAdmins: new Set(list.filter((u) => u.role === 'Admin').map((u) => u.email)) };
  const name = (ref) => ORGS.find((o) => o.code === ref)?.en || '';
  const phases = { prepitch: true, pitching: true, resubmit: true, ...(d.phases || {}) };

  page.innerHTML = `<h1 class="grad">${tr('Admin console', 'คอนโซลผู้ดูแล')}</h1>
    <h2 class="section-title">${tr('Users and roles', 'ผู้ใช้และสิทธิ์')}</h2>
    <div class="group scroll-x"><table class="grid-table edit"><thead><tr><th>${tr('Name shown on the site', 'ชื่อที่แสดง')}</th><th>${tr('Gmail address', 'อีเมล Gmail')}</th><th>${tr('Role', 'สิทธิ์')}</th>
      <th>${tr('In charge of (RefCodes, comma separated)', 'รับผิดชอบองค์กร (รหัส คั่นด้วยจุลภาค)')}</th><th></th></tr></thead>
      <tbody id="users"></tbody></table>
      <div class="table-actions"><button class="btn" id="add">${icon('plus', 16)}${tr('Add user', 'เพิ่มผู้ใช้')}</button><span class="grow"></span>
        <button class="btn primary" id="save">${tr('Save users', 'บันทึกผู้ใช้')}</button></div></div>
    <p class="foot">${tr('Admin: everything, including the edit log. Chief: everything except the edit log; cannot add, change or remove Admins. Staff: edit organization data, no admin console. Emails not listed, or without a role, cannot open the site.', 'ผู้ดูแล: ทำได้ทุกอย่าง รวมถึงดูบันทึกการแก้ไข หัวหน้า: ทำได้ทุกอย่างยกเว้นดูบันทึกการแก้ไข และเพิ่ม เปลี่ยน หรือลบผู้ดูแลไม่ได้ เจ้าหน้าที่: แก้ไขข้อมูลองค์กรได้ แต่เข้าคอนโซลผู้ดูแลไม่ได้ อีเมลที่ไม่อยู่ในรายชื่อหรือยังไม่มีสิทธิ์จะเข้าเว็บไซต์ไม่ได้')}</p>

    <h2 class="section-title">${tr('Pitching phases', 'ช่วงการเสนองบ')}</h2>
    <div class="group" id="phases">${PHASE_LIST.map(([k, label]) => `<label class="row switch-row"><span class="grow">${esc(label())}</span>
      <span class="muted small" data-state="${k}">${phases[k] ? tr('On', 'เปิด') : esc(PENDING())}</span>
      <input type="checkbox" class="switch" data-phase="${k}" ${phases[k] ? 'checked' : ''} aria-label="${esc(label())}"></label>`).join('')}</div>
    <p class="foot">${tr('A phase that is off shows "Pending Phase" on every organization page, and nothing in it can be changed.', 'ช่วงที่ปิดอยู่จะแสดง "รอเปิดช่วง" ในทุกหน้าองค์กร และแก้ไขข้อมูลในช่วงนั้นไม่ได้')}</p>

    <h2 class="section-title">${tr('Allocated budget', 'กรอบงบประมาณ')}</h2>
    <div class="group alloc"><span class="muted">฿</span><input type="text" inputmode="decimal" id="alloc" value="${esc(moneyBox(d.allocated))}" placeholder="0.00" aria-label="${esc(tr('Allocated budget', 'กรอบงบประมาณ'))}">
      <button class="btn primary" id="saveAlloc">${tr('Save', 'บันทึก')}</button></div>
    <p class="foot">${tr('The Dashboard compares the total proposed budget and the final budget with this amount. Admins and Chiefs can change it.', 'แดชบอร์ดเปรียบเทียบงบประมาณที่เสนอทั้งหมดและงบประมาณอนุมัติขั้นสุดท้ายกับจำนวนนี้ ผู้ดูแลและหัวหน้าแก้ไขได้')}</p>

    <h2 class="section-title">${tr('Organizations removed on the pitching site', 'องค์กรที่ถูกนำออกในเว็บไซต์เสนองบ')}</h2>
    <div class="group">${d.removed.length ? d.removed.map((r) => `<a class="row" href="#/org/${encodeURIComponent(r)}"><span class="mono">${esc(r)}</span><span class="grow">${esc(name(r))}</span>${icon('chev', 16)}</a>`).join('')
    : `<div class="row muted">${tr('None', 'ไม่มี')}</div>`}</div>
    <p class="foot">${tr('Hidden from the public list.', 'ไม่แสดงในรายชื่อองค์กร')}</p>

    ${d.log ? `<h2 class="section-title">${tr('Pre-pitch edit log', 'บันทึกการแก้ไขก่อนการนำเสนอ')} <span class="muted">${tr('latest 500', '500 รายการล่าสุด')}</span></h2>
    <div class="toolbar"><input type="search" id="lq" placeholder="${esc(tr('Filter by RefCode, person or field', 'กรองตามรหัส ผู้แก้ไข หรือช่อง'))}"></div>
    <div class="group scroll-x"><table class="grid-table static"><thead><tr><th>${tr('When', 'เวลา')}</th><th>${tr('By', 'โดย')}</th><th>${tr('Project', 'โครงการ')}</th>
      <th>${tr('Field', 'ช่อง')}</th><th>${tr('Old', 'เดิม')}</th><th>${tr('New', 'ใหม่')}</th></tr></thead><tbody id="log"></tbody></table></div>` : ''}`;

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
  const cell = (s) => { const f = /^\{"id"/.test(s) ? (() => { try { return JSON.parse(s).name; } catch { return s; } })() : s; return esc(f); };
  const paintLog = () => {
    const q = $('#lq', page).value.trim().toLowerCase();
    const rows = d.log.filter((r) => !q || `${r.ref} ${r.by} ${ctx.nameOf(r.by)} ${r.field}`.toLowerCase().includes(q));
    $('#log', page).innerHTML = rows.map((r) => `<tr><td class="small">${dateTimeFmt(r.at)}</td><td class="small">${esc(ctx.nameOf(r.by))}</td><td class="mono">${esc(r.ref)}</td>
      <td>${esc(r.field)}</td><td class="small">${cell(r.old)}</td><td class="small">${cell(r.now)}</td></tr>`).join('')
      || `<tr><td colspan="6" class="muted">${tr('No edits yet', 'ยังไม่มีการแก้ไข')}</td></tr>`;
  };
  paintUsers();
  if (d.log) { paintLog(); $('#lq', page).oninput = paintLog; }

  // Phase switch: saves at once; on failure the switch flips back.
  $('#phases', page).onchange = async (e) => {
    const k = e.target.dataset.phase;
    if (!k) return;
    const on = e.target.checked;
    e.target.disabled = true;
    try {
      const out = await call('savePhases', { phases: { ...phases, [k]: on } });
      Object.assign(phases, out.phases);
      d.phases = out.phases;
      remember('admin', {}, d); forget('admin');
      if (ctx.boot) { ctx.boot.phases = out.phases; remember('boot', {}, ctx.boot); }
      forget('boot'); forget('org');
      $(`[data-state="${k}"]`, page).textContent = on ? tr('On', 'เปิด') : PENDING();
      toast(on ? tr('Phase on', 'เปิดช่วงแล้ว') : tr('Phase off — shows Pending Phase', 'ปิดช่วงแล้ว — แสดง รอเปิดช่วง'));
    } catch (err) {
      e.target.checked = !on;
      toast(err.message, 'bad');
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
  $('#save', page).onclick = async (e) => {
    e.target.disabled = true;
    try {
      const out = await call('saveUsers', { users: st.users.filter((u) => u.email) });
      st.users = out.users.map((u) => ({ ...u }));
      st.saved = JSON.stringify(st.users);
      d.users = out.users;
      remember('admin', {}, d); forget('admin');
      if (ctx.boot) { ctx.boot.names = Object.fromEntries(out.users.filter((u) => u.name).map((u) => [u.email, u.name])); remember('boot', {}, ctx.boot); }
      paintUsers();
      forget('boot'); // persons in charge on the list
      toast(tr('Users saved', 'บันทึกผู้ใช้แล้ว'));
    } catch (err) {
      toast(err.message, 'bad');
    }
    e.target.disabled = false;
  };
}
