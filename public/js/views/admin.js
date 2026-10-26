// Admin console (v2.7 order): Open Google Sheet, data export (Excel / CSV; Admins only, confirmed and logged — it
// carries phones and emails), users and roles (with the organizations and projects each person is in charge of),
// pitching phases (On / Off, applied by Save), allocated budget, organizations removed on the pitching site (Admins
// and Chiefs can add one back to our list), and (Admins only) the edit logs: Pre-pitch, Pitching, Re-submission.
// Admins and Chiefs; Data team staff see the edit logs only. The server checks too.
// v2.9: every save runs in the progress window, exports in a "Preparing export…" window; "Export data for mailing"
// (Admins only, logged) builds the two sheets of the budget allocation email; "In charge of" is folded into a chip.
// v2.13: round codes (First-Submission or Re-Submission per round of the pitching site); PDF editor defaults (text
// size, text colour, shape colour, font).
import { tr } from '../i18n.js';
import { $, esc, icon, toast, dateTimeFmt, moneyBox, parseMoney, liveMoney, errorBox, confirmBox, ROLE_LABEL, busy, saving, dialog } from '../ui.js';
import { call, cached, forget, forgetAll, peek, remember } from '../api.js';
import { ORGS } from '../masterdata.js';
import { PENDING } from '../project.js';
import { STEPS } from '../timeline.js';
import { FONTS } from '../markup/pdf-markup.js';

const ROLES = ['Admin', 'Chief', 'Staff', 'Data team staff'];
const PHASE_LIST = [['prepitch', () => tr('Pre-pitch', 'ก่อนการนำเสนอ')], ['pitching', () => tr('Pitching', 'การนำเสนอ')], ['resubmit', () => tr('Re-submission', 'การส่งแก้ไข')]];
const roleLabel = ROLE_LABEL;
const title = (ic, text, extra = '') => `<h2 class="section-title">${icon(ic, 20)}${text}${extra}</h2>`;
// SheetJS builds the Excel file in the browser; loaded only when an Excel export is saved.
const XLSX_URL = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
let st = null; // { users (working copy), saved (JSON), phases (draft), phasesSaved (JSON), save: { users, phases } }

const usersDirty = () => !!st?.users && JSON.stringify(st.users) !== st.saved;
const phasesDirty = () => !!st?.phases && JSON.stringify(st.phases) !== st.phasesSaved;
const datesDirty = () => !!st?.dates && JSON.stringify(st.dates) !== st.datesSaved;
const roundsDirty = () => !!st?.rounds && JSON.stringify(st.rounds) !== st.roundsSaved;
export const isDirty = () => usersDirty() || phasesDirty() || datesDirty() || roundsDirty();
// Unsaved-changes window: Save and leave saves users and phases.
export const canSaveAll = () => isDirty();
export async function saveAll() {
  if (usersDirty()) await st.save.users();
  if (phasesDirty()) await st.save.phases();
  if (datesDirty()) await st.save.dates();
  if (roundsDirty()) await st.save.rounds();
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
  if (!list.some((u) => u.email === d.owner)) list.unshift({ email: d.owner, role: 'Admin', orgs: [], name: '', email2: '' });
  const chief = ctx.me.role !== 'Admin';
  const phases = { prepitch: true, pitching: true, resubmit: true, ...(d.phases || {}) };
  st = { users: list, saved: JSON.stringify(list), savedAdmins: new Set(list.filter((u) => u.role === 'Admin').map((u) => u.email)),
    phases, phasesSaved: JSON.stringify(phases), save: {} };
  // Timeline dates (v2.12): a draft until Save, as the phase switches.
  st.dates = d.dates ? structuredClone(d.dates) : null;
  st.datesSaved = JSON.stringify(st.dates);
  // Round codes (v2.13): { code: kind }, a draft until Save.
  st.rounds = d.rounds ? Object.fromEntries(d.rounds.map((r) => [r.code, r.kind])) : null;
  st.roundsSaved = JSON.stringify(st.rounds);
  const name = (ref) => ORGS.find((o) => o.code === ref)?.en || '';

  const admin = ctx.me.role === 'Admin';
  const restored = new Set(d.restored || []);
  const usersPart = `${title('user', tr('Users and roles', 'ผู้ใช้และสิทธิ์'))}
    <div class="group scroll-x"><table class="grid-table edit"><thead><tr><th>${tr('Name shown on the site', 'ชื่อที่แสดง')}</th><th>${tr('Gmail address', 'อีเมล Gmail')}</th><th>${tr('Second Gmail (optional)', 'อีเมล Gmail ที่สอง (ถ้ามี)')}</th><th>${tr('Role', 'สิทธิ์')}</th>
      <th>${tr('In charge of (RefCodes, comma separated)', 'รับผิดชอบองค์กร (รหัส คั่นด้วยจุลภาค)')}</th><th class="r">${tr('Projects in charge', 'โครงการที่รับผิดชอบ')}</th><th></th></tr></thead>
      <tbody id="users"></tbody></table>
      <div class="table-actions"><button class="btn" id="add">${icon('plus', 16)}${tr('Add user', 'เพิ่มผู้ใช้')}</button>
        <button class="btn" id="importCsv">${icon('upload', 16)}${tr('Import CSV', 'นำเข้า CSV')}</button><input type="file" id="csvFile" accept=".csv,text/csv" hidden><span class="grow"></span>
        <button class="btn primary" id="save">${tr('Save users', 'บันทึกผู้ใช้')}</button></div></div>
    <p class="foot">${tr('Admin: everything, including the edit log. Chief: everything except the edit log; cannot add, change or remove Admins. Staff: edit organization data, no admin console. Data team staff: as Staff, plus Control, the edit log and closing or opening a case. Emails not listed, or without a role, cannot open the site. Second Gmail: one person can sign in with up to 2 Gmail addresses; both get the same role, name and organizations. Import CSV: columns Username, Gmail, Gmail#2 (comma or tab separated; Gmail#2 may be empty). Invalid emails are ignored, and an invalid Gmail falls back to Gmail#2. An email already listed updates the name and second Gmail of that user; new emails are added as Staff; rows with no valid email are listed in a window. Press Save users to keep them.', 'ผู้ดูแล: ทำได้ทุกอย่าง รวมถึงดูบันทึกการแก้ไข หัวหน้า: ทำได้ทุกอย่างยกเว้นดูบันทึกการแก้ไข และเพิ่ม เปลี่ยน หรือลบผู้ดูแลไม่ได้ เจ้าหน้าที่: แก้ไขข้อมูลองค์กรได้ แต่เข้าคอนโซลผู้ดูแลไม่ได้ เจ้าหน้าที่ทีมข้อมูล: เหมือนเจ้าหน้าที่ และจัดการหน้าควบคุม ดูบันทึกการแก้ไข และปิดหรือเปิดเคสได้ อีเมลที่ไม่อยู่ในรายชื่อหรือยังไม่มีสิทธิ์จะเข้าเว็บไซต์ไม่ได้ อีเมล Gmail ที่สอง: หนึ่งคนเข้าสู่ระบบได้ด้วย Gmail สูงสุด 2 บัญชี ทั้งสองบัญชีได้สิทธิ์ ชื่อ และองค์กรที่รับผิดชอบเดียวกัน นำเข้า CSV: คอลัมน์ Username, Gmail, Gmail#2 (คั่นด้วยจุลภาคหรือแท็บ Gmail#2 เว้นว่างได้) อีเมลที่ไม่ถูกต้องจะถูกข้าม ถ้า Gmail ไม่ถูกต้องจะใช้ Gmail#2 แทน อีเมลที่มีอยู่แล้วจะแก้ไขชื่อและ Gmail ที่สองของผู้ใช้นั้น อีเมลใหม่จะได้สิทธิ์เจ้าหน้าที่ แถวที่ไม่มีอีเมลที่ถูกต้องจะแสดงในหน้าต่าง กดบันทึกผู้ใช้เพื่อเก็บไว้')}</p>`;

  const phasesPart = `${title('clock', tr('Pitching phases', 'ช่วงการเสนองบ'))}
    <div class="group" id="phases">${PHASE_LIST.map(([k, label]) => `<label class="row phase-row"><span class="grow">${esc(label())}</span>
      <span class="muted small" data-state="${k}">${phases[k] ? tr('On', 'เปิด') : esc(PENDING())}</span>
      <input type="checkbox" data-phase="${k}" ${phases[k] ? 'checked' : ''} aria-label="${esc(label())}"></label>`).join('')}
      <div class="table-actions"><span class="grow muted small" id="phasesNote"></span><button class="btn primary" id="savePhases" disabled>${tr('Save', 'บันทึก')}</button></div></div>
    <p class="foot">${tr('A phase that is off shows "Pending Phase" on every organization page, and nothing in it can be changed. Changes apply only when you press Save.', 'ช่วงที่ปิดอยู่จะแสดง "รอเปิดช่วง" ในทุกหน้าองค์กร และแก้ไขข้อมูลในช่วงนั้นไม่ได้ การเปลี่ยนแปลงจะมีผลเมื่อกดบันทึกเท่านั้น')}</p>`;

  const datesPart = st.dates ? `${title('clock', tr('Timeline dates', 'วันที่ในไทม์ไลน์'))}
    <div class="group" id="dates">${STEPS.filter(([k]) => st.dates[k]).map(([k, label]) => `<div class="row date-row"><span class="grow">${esc(label())}</span>
      <input type="date" data-date="${k}" data-end="from" value="${esc(st.dates[k].from)}" aria-label="${esc(tr('{step}: start', '{step}: เริ่ม', { step: label() }))}">
      <span class="muted">–</span><input type="date" data-date="${k}" data-end="to" value="${esc(st.dates[k].to)}" aria-label="${esc(tr('{step}: end', '{step}: สิ้นสุด', { step: label() }))}"></div>`).join('')}
      <div class="table-actions"><span class="grow muted small" id="datesNote"></span><button class="btn primary" id="saveDates" disabled>${tr('Save', 'บันทึก')}</button></div></div>
    <p class="foot">${tr('The timeline on the Organizations tab is drawn from these dates (first filled from the "TUSU:" calendar events). They do not switch phases on or off. Admins and Chiefs can change them.', 'ไทม์ไลน์ในหน้าองค์กรวาดจากวันที่เหล่านี้ (ตั้งต้นจากกิจกรรม "TUSU:" ในปฏิทิน) วันที่ไม่ได้เปิดหรือปิดช่วง ผู้ดูแลและหัวหน้าแก้ไขได้')}</p>` : '';

  const exportPart = `${title('download', tr('Export data', 'ส่งออกข้อมูล'))}
    <div class="group" id="export">
      <div class="row"><span class="grow">${tr('Organizations and projects', 'องค์กรและโครงการ')}</span><span class="muted small">${tr('Always included', 'รวมเสมอ')}</span></div>
      ${[...PHASE_LIST, ['emails', () => tr('Notification emails (one row per email)', 'อีเมลรับการแจ้งเตือน (หนึ่งแถวต่ออีเมล)')]].map(([k, label]) => `<label class="row phase-row"><span class="grow">${esc(label())}</span><input type="checkbox" data-ex="${k}" checked aria-label="${esc(label())}"></label>`).join('')}
      <div class="row"><span class="grow">${tr('Format', 'รูปแบบไฟล์')}</span>
        <div class="seg" role="radiogroup"><label><input type="radio" name="exfmt" value="xlsx" checked><span>Excel</span></label><label><input type="radio" name="exfmt" value="csv"><span>CSV</span></label></div></div>
      <div class="table-actions"><span class="grow muted small">${tr('Nothing downloads until you press Save.', 'จะดาวน์โหลดเมื่อกดบันทึกเท่านั้น')}</span><button class="btn primary" id="saveExport">${icon('download', 16)}${tr('Save', 'บันทึก')}</button></div>
      <div class="row"><span class="grow">${tr('Data for mailing', 'ข้อมูลสำหรับส่งอีเมล')}<small class="muted block">${tr('Excel: Organizations (emails, budgets, category and campus benchmarks, shares) and Projects (contacts, dates, budgets, Committee feedback, links)', 'Excel: ชีตองค์กร (อีเมล งบประมาณ ค่าเฉลี่ยหมวดหมู่และศูนย์ สัดส่วน) และชีตโครงการ (ผู้ติดต่อ วันที่ งบประมาณ ความเห็นคณะกรรมการ ลิงก์)')}</small></span>
        <button class="btn" id="saveMail">${icon('download', 16)}${tr('Export data for mailing', 'ส่งออกข้อมูลสำหรับส่งอีเมล')}</button></div></div>
    <p class="foot">${tr('Admins only. Excel: one sheet for organizations, one for projects, one per included phase and one for emails. CSV: one row per project with every included column, plus a second file of emails. Organizations hidden from the list are left out. The file holds phone numbers and emails: every export is logged.', 'เฉพาะผู้ดูแล Excel: ชีตองค์กร ชีตโครงการ หนึ่งชีตต่อช่วงที่เลือก และชีตอีเมล CSV: หนึ่งแถวต่อโครงการ พร้อมทุกคอลัมน์ที่เลือก และไฟล์อีเมลอีกหนึ่งไฟล์ ไม่รวมองค์กรที่ซ่อนจากรายชื่อ ไฟล์มีเบอร์โทรและอีเมล ทุกการส่งออกจะถูกบันทึก')}</p>`;

  // Round codes (v2.13): which rounds of the pitching site feed Pre-pitch and Pitching, and which the Re-submission tab.
  const KINDS = [['first', () => tr('First-Submission', 'ส่งครั้งแรก')], ['resub', () => tr('Re-Submission', 'ส่งแก้ไข')]];
  const roundsPart = st.rounds ? `${title('clock', tr('Round codes', 'รหัสรอบ'))}
    <div class="group" id="rounds">${Object.keys(st.rounds).length ? Object.entries(st.rounds).map(([code, kind]) => `<label class="row phase-row"><span class="grow mono">${esc(code)}</span>
      <select data-round="${esc(code)}" aria-label="${esc(tr('Round {code}', 'รอบ {code}', { code }))}">${KINDS.map(([k, l]) => `<option value="${k}" ${kind === k ? 'selected' : ''}>${esc(l())}</option>`).join('')}</select></label>`).join('')
    : `<div class="row muted">${tr('No submissions yet', 'ยังไม่มีการส่งข้อเสนอ')}</div>`}
      <div class="table-actions"><span class="grow muted small" id="roundsNote"></span><button class="btn primary" id="saveRounds" disabled>${tr('Save', 'บันทึก')}</button></div></div>
    <p class="foot">${tr('Round codes come from the pitching site (Settings → Round code). First-Submission rounds feed Pre-pitch and Pitching; Re-Submission rounds feed the Re-submission tab. A round not set here: the first code is First-Submission, every later one Re-Submission (2569A, then 2569B onwards). Admins and Chiefs can change it.', 'รหัสรอบมาจากเว็บไซต์เสนองบ (ตั้งค่า → รหัสรอบ) รอบส่งครั้งแรกใช้ในช่วงก่อนการนำเสนอและการนำเสนอ รอบส่งแก้ไขใช้ในแท็บการส่งแก้ไข รอบที่ยังไม่ได้ตั้ง: รหัสแรกเป็นส่งครั้งแรก รหัสถัดไปทั้งหมดเป็นส่งแก้ไข (2569A แล้ว 2569B เป็นต้นไป) ผู้ดูแลและหัวหน้าแก้ไขได้')}</p>` : '';

  // PDF editor (v2.12 size, v2.13 colours and font): defaults of new text boxes and shapes; each person can still
  // change them in the editor.
  const mk = { size: 11, textColor: '#990000', shapeColor: '#990000', font: FONTS[0], ...(d.markup || {}) };
  const sizePart = `${title('pen', tr('PDF editor', 'เครื่องมือแก้ไข PDF'))}
    <div class="group" id="markup">
      <div class="row"><span class="grow">${tr('Default text box size', 'ขนาดกล่องข้อความเริ่มต้น')}</span>
        <input type="number" id="mkSize" min="6" max="72" step="1" value="${esc(mk.size)}" aria-label="${esc(tr('Default text box size (pt)', 'ขนาดกล่องข้อความเริ่มต้น (pt)'))}"><span class="muted">pt</span></div>
      <div class="row"><span class="grow">${tr('Default text colour', 'สีข้อความเริ่มต้น')}</span><input type="color" id="mkText" value="${esc(mk.textColor)}" aria-label="${esc(tr('Default text colour', 'สีข้อความเริ่มต้น'))}"></div>
      <div class="row"><span class="grow">${tr('Default shape colour (draw, line, arrow, rectangle)', 'สีรูปร่างเริ่มต้น (วาด เส้น ลูกศร สี่เหลี่ยม)')}</span><input type="color" id="mkShape" value="${esc(mk.shapeColor)}" aria-label="${esc(tr('Default shape colour', 'สีรูปร่างเริ่มต้น'))}"></div>
      <div class="row"><span class="grow">${tr('Default font', 'แบบอักษรเริ่มต้น')}</span><select id="mkFont" aria-label="${esc(tr('Default font', 'แบบอักษรเริ่มต้น'))}">${FONTS.map((f) => `<option ${f === mk.font ? 'selected' : ''}>${esc(f)}</option>`).join('')}</select></div>
      <div class="table-actions"><span class="grow"></span><button class="btn primary" id="saveMarkup">${tr('Save', 'บันทึก')}</button></div></div>
    <p class="foot">${tr('New text boxes in the PDF editor start at this size (6–72 pt), colour and font; new shapes start at the shape colour. Staff can still change any text box or shape. Admins and Chiefs can change these.', 'กล่องข้อความใหม่ในเครื่องมือแก้ไข PDF จะเริ่มที่ขนาด (6–72 pt) สี และแบบอักษรนี้ รูปร่างใหม่จะเริ่มที่สีรูปร่าง เจ้าหน้าที่ยังปรับแต่ละกล่องหรือรูปร่างได้ ผู้ดูแลและหัวหน้าแก้ไขได้')}</p>`;

  const allocPart = `${title('chart', tr('Allocated budget', 'กรอบงบประมาณ'))}
    <div class="group alloc"><span class="muted">฿</span><input type="text" inputmode="decimal" id="alloc" value="${esc(moneyBox(d.allocated))}" placeholder="0.00" aria-label="${esc(tr('Allocated budget', 'กรอบงบประมาณ'))}">
      <button class="btn primary" id="saveAlloc">${tr('Save', 'บันทึก')}</button></div>
    <p class="foot">${tr('The Dashboard compares the total proposed budget and the final budget with this amount. Admins and Chiefs can change it.', 'แดชบอร์ดเปรียบเทียบงบประมาณที่เสนอทั้งหมดและงบประมาณอนุมัติขั้นสุดท้ายกับจำนวนนี้ ผู้ดูแลและหัวหน้าแก้ไขได้')}</p>`;

  // Removed on the pitching site: hidden from our list unless added back (v2.7).
  const removedPart = `${title('eye', tr('Organizations removed on the pitching site', 'องค์กรที่ถูกนำออกในเว็บไซต์เสนองบ'))}
    <div class="group" id="removed">${d.removed.length ? d.removed.map((r) => `<div class="row removed-row"><a class="mono" href="#/org/${encodeURIComponent(r)}">${esc(r)}</a><span class="grow">${esc(name(r))}</span>
      ${restored.has(r) ? `<span class="pill green">${tr('On the list', 'อยู่ในรายชื่อ')}</span><button class="btn plain" data-restore="${esc(r)}" data-on="0">${tr('Remove again', 'นำออกอีกครั้ง')}</button>`
    : `<button class="btn" data-restore="${esc(r)}" data-on="1">${icon('plus', 16)}${tr('Add to list', 'เพิ่มในรายชื่อ')}</button>`}</div>`).join('')
    : `<div class="row muted">${tr('None', 'ไม่มี')}</div>`}</div>
    <p class="foot">${tr('Hidden from the list, Dashboard, export and print unless added back. Admins and Chiefs can add one back or remove it again; each change is logged.', 'ไม่แสดงในรายชื่อ แดชบอร์ด การส่งออก และการพิมพ์ เว้นแต่เพิ่มกลับ ผู้ดูแลและหัวหน้าเพิ่มกลับหรือนำออกอีกครั้งได้ ทุกการเปลี่ยนแปลงจะถูกบันทึก')}</p>`;

  page.innerHTML = `<h1 class="grad">${tr('Admin console', 'คอนโซลผู้ดูแล')}</h1>
    <div class="admin-top">${d.sheetUrl ? `<a class="btn" href="${esc(d.sheetUrl)}" target="_blank" rel="noopener">${icon('ext', 16)}${tr('Open Google Sheet', 'เปิด Google Sheet')}</a>` : ''}</div>
    ${admin ? exportPart : ''}${usersPart}${phasesPart}${datesPart}${roundsPart}${sizePart}${allocPart}${removedPart}
    ${d.log ? logSection() : ''}`;

  // Projects in charge: the project counts of the user's organizations (boot data), e.g. "9 (3 orgs)".
  const inCharge = (orgs) => {
    const n = orgs.reduce((t, r) => t + (ctx.boot?.orgs?.[r]?.count || 0), 0);
    return orgs.length ? `${n} <span class="muted">${tr('({o} orgs)', '({o} องค์กร)', { o: orgs.length })}</span>` : '<span class="muted">—</span>';
  };
  // "In charge of" folds into a chip ("3 orgs ▾"); a click opens the RefCode field, a second click closes it.
  const chipText = (orgs) => `${orgs.length ? esc(tr('{n} orgs', '{n} องค์กร', { n: orgs.length })) : esc(tr('None', 'ไม่มี'))} ▾`;
  st.open = new Set();
  const paintUsers = () => {
    $('#users', page).innerHTML = st.users.map((u, i) => {
      const owner = u.email === d.owner;
      const locked = owner || (chief && st.savedAdmins.has(u.email));
      return `<tr data-i="${i}">
      <td><input type="text" data-f="name" value="${esc(u.name || '')}" maxlength="80" placeholder="${esc(tr('e.g. Mint (Finance)', 'เช่น มิ้นท์ (ฝ่ายการเงิน)'))}"></td>
      <td><input type="email" data-f="email" value="${esc(u.email)}" placeholder="name@gmail.com" ${locked ? 'readonly' : ''}></td>
      <td><input type="email" data-f="email2" value="${esc(u.email2 || '')}" placeholder="${esc(tr('another@gmail.com', 'อีกบัญชี@gmail.com'))}" ${chief && locked ? 'readonly' : ''}></td>
      <td>${locked ? `<span class="muted small">${owner ? tr('Admin · script owner', 'ผู้ดูแล · เจ้าของสคริปต์') : roleLabel('Admin')}</span>`
    : `<select data-f="role">${u.role ? '' : `<option value="" selected>${tr('Choose a role', 'เลือกสิทธิ์')}</option>`}${ROLES.filter((r) => !chief || r !== 'Admin').map((r) => `<option value="${r}" ${u.role === r ? 'selected' : ''}>${esc(roleLabel(r))}</option>`).join('')}</select>`}</td>
      <td><button type="button" class="chip-btn" data-fold="${i}" aria-expanded="${st.open.has(i)}">${chipText(u.orgs)}</button>
        <input type="text" data-f="orgs" value="${esc(u.orgs.join(', '))}" placeholder="RS-ART-03, RS-OTH-01" list="refs" ${st.open.has(i) ? '' : 'hidden'}></td>
      <td class="r small" data-count>${inCharge(u.orgs)}</td>
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
    const out = await saving(() => call('savePhases', { phases: st.phases }));
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

  $('#saveMarkup', page).onclick = async (e) => {
    e.target.disabled = true;
    try {
      const v = { size: $('#mkSize', page).value, textColor: $('#mkText', page).value, shapeColor: $('#mkShape', page).value, font: $('#mkFont', page).value };
      const out = await saving(() => call('saveMarkupDefaults', { defaults: v }));
      $('#mkSize', page).value = out.markup.size;
      d.markup = out.markup;
      remember('admin', {}, d); forget('admin');
      toast(tr('PDF editor defaults saved', 'บันทึกค่าเริ่มต้นของเครื่องมือแก้ไข PDF แล้ว'));
    } catch (err) {
      await errorBox(err.message);
    }
    e.target.disabled = false;
  };

  // Round codes: Save stays off until something changes. Saving changes what every tab reads, so all reads go.
  if (st.rounds) {
    const roundsState = () => {
      $('#saveRounds', page).disabled = !roundsDirty();
      $('#roundsNote', page).textContent = roundsDirty() ? tr('Not saved yet', 'ยังไม่ได้บันทึก') : '';
    };
    $('#rounds', page).onchange = (e) => {
      const k = e.target.dataset.round;
      if (!k) return;
      st.rounds[k] = e.target.value;
      roundsState();
    };
    st.save.rounds = async () => {
      const out = await saving(() => call('saveRoundKinds', { kinds: st.rounds }));
      st.rounds = Object.fromEntries(out.rounds.map((r) => [r.code, r.kind]));
      st.roundsSaved = JSON.stringify(st.rounds);
      d.rounds = out.rounds;
      remember('admin', {}, d);
      forgetAll();
    };
    $('#saveRounds', page).onclick = async (e) => {
      e.target.disabled = true;
      try {
        await st.save.rounds();
        toast(tr('Round codes saved', 'บันทึกรหัสรอบแล้ว'));
      } catch (err) {
        await errorBox(err.message);
      }
      roundsState();
    };
  }

  // Timeline dates: Save stays off until something changes.
  if (st.dates) {
    const datesState = () => {
      $('#saveDates', page).disabled = !datesDirty();
      $('#datesNote', page).textContent = datesDirty() ? tr('Not saved yet', 'ยังไม่ได้บันทึก') : '';
    };
    $('#dates', page).onchange = (e) => {
      const k = e.target.dataset.date;
      if (!k) return;
      st.dates[k][e.target.dataset.end] = e.target.value;
      datesState();
    };
    st.save.dates = async () => {
      const out = await saving(() => call('savePhaseDates', { dates: st.dates }));
      st.dates = structuredClone(out.dates);
      st.datesSaved = JSON.stringify(st.dates);
      if (ctx.boot) { ctx.boot.dates = out.dates; remember('boot', {}, ctx.boot); }
      d.dates = out.dates;
      remember('admin', {}, d); forget('admin');
      forget('boot');
    };
    $('#saveDates', page).onclick = async (e) => {
      e.target.disabled = true;
      try {
        await st.save.dates();
        toast(tr('Dates saved', 'บันทึกวันที่แล้ว'));
      } catch (err) {
        await errorBox(err.message);
      }
      datesState();
    };
  }

  // Export (Admins only): download on Save only, after a reminder that the file holds personal data.
  if (admin) $('#saveExport', page).onclick = async (e) => {
    const want = Object.fromEntries(['prepitch', 'pitching', 'resubmit', 'emails'].map((k) => [k, $(`[data-ex="${k}"]`, page).checked]));
    const fmt = $('input[name=exfmt]:checked', page).value;
    const ok = await confirmBox(tr('Export personal data?', 'ส่งออกข้อมูลส่วนบุคคล?'),
      tr('The file holds names, phone numbers and emails. Keep it private and delete it when you are done. Your export is logged.', 'ไฟล์มีชื่อ เบอร์โทร และอีเมล โปรดเก็บเป็นความลับและลบเมื่อใช้เสร็จ การส่งออกนี้จะถูกบันทึก'),
      { ok: tr('Export', 'ส่งออก') });
    if (!ok) return;
    e.target.disabled = true;
    try {
      await busy(PREPARING(), async () => {
        const data = await call('exportData', { phases: want, format: fmt });
        await (fmt === 'csv' ? saveCsv(data) : saveXlsx(data));
      });
      toast(tr('Export saved', 'บันทึกไฟล์ส่งออกแล้ว'));
    } catch (err) {
      await errorBox(err.message, tr('Couldn’t export', 'ส่งออกไม่สำเร็จ'));
    }
    e.target.disabled = false;
  };

  // Export data for mailing (v2.9, Admins only, logged): the fields of the budget allocation email.
  if (admin) $('#saveMail', page).onclick = async (e) => {
    const ok = await confirmBox(tr('Export personal data?', 'ส่งออกข้อมูลส่วนบุคคล?'),
      tr('The file holds names, phone numbers and emails. Keep it private and delete it when you are done. Your export is logged.', 'ไฟล์มีชื่อ เบอร์โทร และอีเมล โปรดเก็บเป็นความลับและลบเมื่อใช้เสร็จ การส่งออกนี้จะถูกบันทึก'),
      { ok: tr('Export', 'ส่งออก') });
    if (!ok) return;
    const btn = e.target.closest('button');
    btn.disabled = true;
    try {
      await busy(PREPARING(), async () => saveMailXlsx(await call('mailData')));
      toast(tr('Export saved', 'บันทึกไฟล์ส่งออกแล้ว'));
    } catch (err) {
      await errorBox(err.message, tr('Couldn’t export', 'ส่งออกไม่สำเร็จ'));
    }
    btn.disabled = false;
  };

  $('#alloc', page).oninput = (e) => liveMoney(e.target);
  $('#saveAlloc', page).onclick = async (e) => {
    e.target.disabled = true;
    try {
      const out = await saving(() => call('saveAllocated', { allocated: parseMoney($('#alloc', page).value) }));
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

  $('#removed', page).onclick = async (e) => {
    const b = e.target.closest('[data-restore]');
    if (!b) return;
    b.disabled = true;
    try {
      const out = await saving(() => call('setRestored', { ref: b.dataset.restore, on: b.dataset.on === '1' }));
      d.restored = out.restored;
      remember('admin', {}, d); forget('admin');
      forget('boot'); forget('dashboard');
      draw(page, ctx, d);
      toast(b.dataset.on === '1' ? tr('{ref} is back on the list', '{ref} กลับเข้ารายชื่อแล้ว', { ref: b.dataset.restore }) : tr('{ref} is hidden again', 'ซ่อน {ref} อีกครั้งแล้ว', { ref: b.dataset.restore }));
    } catch (err) {
      toast(err.message, 'bad');
      b.disabled = false;
    }
  };

  const users = $('#users', page);
  users.oninput = users.onchange = (e) => {
    const i = Number(e.target.closest('[data-i]')?.dataset.i); const f = e.target.dataset.f;
    if (!f) return;
    st.users[i][f] = f === 'orgs' ? e.target.value.split(/[,\s]+/).map((x) => x.trim().toUpperCase()).filter(Boolean) : e.target.value.trim();
    if (f === 'orgs') {
      e.target.closest('tr').querySelector('[data-count]').innerHTML = inCharge(st.users[i].orgs);
      e.target.closest('tr').querySelector('[data-fold]').innerHTML = chipText(st.users[i].orgs);
    }
  };
  users.onclick = (e) => {
    const fold = e.target.closest('[data-fold]');
    if (fold) {
      const i = Number(fold.dataset.fold);
      if (st.open.has(i)) st.open.delete(i); else st.open.add(i);
      const input = fold.parentElement.querySelector('[data-f="orgs"]');
      input.hidden = !st.open.has(i);
      fold.setAttribute('aria-expanded', String(st.open.has(i)));
      if (!input.hidden) input.focus();
      return;
    }
    if (!e.target.closest('[data-rm]')) return;
    st.open.clear();
    st.users.splice(Number(e.target.closest('[data-i]').dataset.i), 1);
    paintUsers();
  };
  $('#add', page).onclick = () => { st.users.push({ email: '', role: 'Staff', orgs: [], name: '', email2: '' }); st.open.add(st.users.length - 1); paintUsers(); $('#users tr:last-child input', page)?.focus(); };
  // Import CSV (v2.11): Username, Gmail, Gmail#2 per row, comma or tab separated (columns found by the header row
  // when there is one). Invalid emails are ignored; an invalid Gmail falls back to Gmail#2. A row whose email is
  // already listed (on the site or earlier in the file) edits that user: Username and Gmail#2. Others join as Staff.
  // Rows with no valid email are listed in a window. Save users keeps the changes.
  $('#importCsv', page).onclick = () => $('#csvFile', page).click();
  $('#csvFile', page).onchange = async (e) => {
    const file = e.target.files[0];
    e.target.value = '';
    if (!file) return;
    let added = 0; let edited = 0; const failed = [];
    const text = (await file.text()).replace(/^﻿/, '');
    const rows = parseCsv(text, /\t/.test(text.split(/\r?\n/)[0]) ? '\t' : ',');
    // Columns: by the header row's names, else Username, Gmail, Gmail#2 in that order.
    const key = (h) => String(h || '').trim().toLowerCase().replace(/[\s_-]/g, '');
    const head = (rows[0] || []).map(key);
    const col = { name: 0, mail: 1, mail2: 2 };
    let first = 1; // file line of the first data row, for the window
    if (head.some((h) => ['username', 'gmail', 'gmail#2', 'email'].includes(h))) {
      rows.shift();
      first = 2;
      col.name = head.indexOf('username');
      col.mail = head.findIndex((h) => h === 'gmail' || h === 'email');
      col.mail2 = head.findIndex((h) => h === 'gmail#2' || h === 'gmail2' || h === 'email2');
    }
    const valid = (x) => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(x);
    const at = (r, i) => (i >= 0 ? String(r[i] ?? '').trim() : '');
    const owner = (m) => st.users.find((u) => u.email.toLowerCase() === m || (u.email2 || '').toLowerCase() === m);
    rows.forEach((r, i) => {
      const name = at(r, col.name).slice(0, 80);
      const m1 = at(r, col.mail).toLowerCase();
      const m2 = at(r, col.mail2).toLowerCase();
      if (!name && !m1 && !m2) return; // blank line
      // Invalid emails are dropped; Gmail#2 becomes the main email when Gmail is not valid.
      const [email, extra = ''] = [m1, m2].filter(valid);
      if (!email) { failed.push({ name, row: i + first, mails: [m1, m2].filter(Boolean) }); return; }
      const u = owner(email) || (extra && owner(extra));
      if (u) {
        // Already listed: update its name and second Gmail (a second Gmail someone else uses is left out).
        if (name) u.name = name;
        const mine = [email, extra].filter(Boolean);
        const second = mine.find((m) => m !== u.email.toLowerCase() && (!owner(m) || owner(m) === u));
        if (second && !mine.includes((u.email2 || '').toLowerCase())) u.email2 = second;
        edited++;
        return;
      }
      st.users.push({ email, role: 'Staff', orgs: [], name, email2: extra && !owner(extra) ? extra : '' });
      added++;
    });
    st.open.clear();
    paintUsers();
    const summary = tr('Added {a}, updated {e}. Press Save users to keep them.', 'เพิ่ม {a} คน แก้ไข {e} คน กดบันทึกผู้ใช้เพื่อเก็บไว้', { a: added, e: edited });
    if (!failed.length) { toast(summary, added || edited ? '' : 'bad'); return; }
    await dialog(tr('{n} users could not be added', 'เพิ่มผู้ใช้ไม่ได้ {n} คน', { n: failed.length }),
      `<p>${esc(summary)}</p><p>${esc(tr('These rows have no valid Gmail:', 'แถวเหล่านี้ไม่มีอีเมล Gmail ที่ถูกต้อง:'))}</p>
      <ul class="unsaved">${failed.map((f) => `<li><span>${esc(f.name || tr('Row {n} (no username)', 'แถว {n} (ไม่มีชื่อ)', { n: f.row }))}</span><span class="muted">${esc(f.mails.join(', ') || tr('no email', 'ไม่มีอีเมล'))}</span></li>`).join('')}</ul>`,
      [{ v: true, label: tr('OK', 'ตกลง'), cls: 'primary' }], { none: true });
  };
  st.save.users = async () => {
    const out = await saving(() => call('saveUsers', { users: st.users.filter((u) => u.email), version: d.usersVer }));
    st.users = out.users.map((u) => ({ ...u }));
    st.saved = JSON.stringify(st.users);
    d.users = out.users;
    d.usersVer = out.usersVer;
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

// CSV text to rows of cells: commas (or `sep`, e.g. a tab), quoted cells ("" inside quotes is a quote), CRLF or LF line ends.
function parseCsv(text, sep = ',') {
  const rows = []; let row = []; let cell = ''; let q = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (q) {
      if (c === '"' && text[i + 1] === '"') { cell += '"'; i++; } else if (c === '"') q = false; else cell += c;
    } else if (c === '"') q = true;
    else if (c === sep) { row.push(cell); cell = ''; }
    else if (c === '\n' || c === '\r') {
      if (c === '\r' && text[i + 1] === '\n') i++;
      row.push(cell); rows.push(row); row = []; cell = '';
    } else cell += c;
  }
  if (cell || row.length) { row.push(cell); rows.push(row); }
  return rows;
}

// Data team staff: the edit log only.
function drawLog(page, ctx, d) {
  st = null;
  page.innerHTML = `<h1 class="grad">${tr('Admin console', 'คอนโซลผู้ดูแล')}</h1>${logSection()}`;
  paintLogRows(page, ctx, d);
}

// Edit logs: Pre-pitch (every changed field, budget-form feedback too), Pitching (v2.7, every changed field) and
// Re-submission (decisions, case close / open, organizations added back to the list).
let logKind = 'pre';
const LOGS = [['pre', 'log', () => tr('Pre-pitch', 'ก่อนการนำเสนอ')], ['pitch', 'pitchLog', () => tr('Pitching', 'การนำเสนอ')], ['resub', 'resubLog', () => tr('Re-submission', 'การส่งแก้ไข')]];
const logSection = () => `${title('file', tr('Edit log', 'บันทึกการแก้ไข'), ` <span class="muted">${tr('latest 500', '500 รายการล่าสุด')}</span>`)}
    <div class="toolbar"><div class="seg" role="radiogroup" id="logKind" aria-label="${esc(tr('Log', 'บันทึก'))}">
      ${LOGS.map(([k, , label]) => `<label><input type="radio" name="logKind" value="${k}" ${logKind === k ? 'checked' : ''}><span>${esc(label())}</span></label>`).join('')}</div>
      <input type="search" id="lq" placeholder="${esc(tr('Filter by RefCode, person or field', 'กรองตามรหัส ผู้แก้ไข หรือช่อง'))}"></div>
    <div class="group scroll-x"><table class="grid-table static"><thead><tr><th>${tr('When', 'เวลา')}</th><th>${tr('By', 'โดย')}</th><th>${tr('Project', 'โครงการ')}</th>
      <th>${tr('Field', 'ช่อง')}</th><th>${tr('Old', 'เดิม')}</th><th>${tr('New', 'ใหม่')}</th></tr></thead><tbody id="log"></tbody></table></div>`;

function paintLogRows(page, ctx, d) {
  const cell = (s) => { const f = /^\{"id"/.test(s) ? (() => { try { return JSON.parse(s).name; } catch { return s; } })() : s; return esc(f); };
  const paint = () => {
    const q = $('#lq', page).value.trim().toLowerCase();
    const list = d[(LOGS.find(([k]) => k === logKind) || LOGS[0])[1]] || [];
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

const PREPARING = () => tr('Preparing export…', 'กำลังเตรียมไฟล์ส่งออก…');
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
  // Emails: their own file, one row per email.
  if (d.emails) {
    const eh = ['Org Ref', 'Organization', 'Campus', 'Email'];
    const et = [eh.join(','), ...d.emails.map((r) => eh.map((h) => q(r[h])).join(','))].join('\r\n');
    setTimeout(() => download(new Blob([`﻿${et}`], { type: 'text/csv;charset=utf-8' }), `pitch-feedback-emails-${stamp()}.csv`), 400);
  }
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

// v2.10: the Pre-pitch and Pitching sheets end with the same summary row: project count and the budget total.
function withTotal(rows) {
  if (!rows.length) return rows;
  const head = Object.keys(rows[0]);
  const total = { [head[0]]: tr('Total ({n} projects)', 'รวม ({n} โครงการ)', { n: rows.length }) };
  head.forEach((h) => { if (/Budget/.test(h)) total[h] = rows.reduce((t, r) => t + (+r[h] || 0), 0); });
  return [...rows, total];
}

// v2.15: a sheet whose cells holding one web address are hyperlinks (one click opens it); the text stays the address,
// for mail-merge templates. A cell with several addresses (one per line) stays text: Excel links one per cell.
function linked(X, rows) {
  const ws = X.utils.json_to_sheet(rows);
  Object.keys(ws).forEach((k) => {
    const c = ws[k];
    if (k[0] !== '!' && c.t === 's' && /^https?:\/\/\S+$/.test(c.v)) c.l = { Target: c.v, Tooltip: c.v };
  });
  return ws;
}

async function saveXlsx(d) {
  const X = await loadXlsx();
  const wb = X.utils.book_new();
  X.utils.book_append_sheet(wb, linked(X, d.orgs), 'Organizations');
  X.utils.book_append_sheet(wb, linked(X, d.projects), 'Projects');
  PHASE_SHEETS.forEach(([k, label]) => { if (d[k]) X.utils.book_append_sheet(wb, linked(X, k === 'resubmit' ? d[k] : withTotal(d[k])), label); });
  if (d.emails) X.utils.book_append_sheet(wb, linked(X, d.emails.length ? d.emails : [{ 'Org Ref': '', Organization: '', Campus: '', Email: '' }]), 'Emails');
  const out = X.write(wb, { bookType: 'xlsx', type: 'array' });
  download(new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), `pitch-feedback-${stamp()}.xlsx`);
}

// Mailing file (v2.9): Organizations and Projects sheets, money and percentages as numbers for a mail-merge template.
async function saveMailXlsx(d) {
  const X = await loadXlsx();
  const wb = X.utils.book_new();
  X.utils.book_append_sheet(wb, linked(X, d.orgs.length ? d.orgs : [{ 'Org Ref': '' }]), 'Organizations');
  X.utils.book_append_sheet(wb, linked(X, d.projects.length ? d.projects : [{ 'Project Ref': '' }]), 'Projects');
  const out = X.write(wb, { bookType: 'xlsx', type: 'array' });
  download(new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), `pitch-mailing-${stamp()}.xlsx`);
}
