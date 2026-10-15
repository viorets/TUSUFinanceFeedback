// Pitching: at the top, the organization's "Correct Excel file" (one .xlsx / .xls per organization since v2.9, with
// its own Save, logged in Pitching_Log). Then per project: the pitching slot (from the scheduling site, read only),
// the summary, the final approved budget and "Noted documents" (any number of .docx / PDF). Unfinished = a
// required question unanswered (project.js). Every save and upload runs in a progress window (busy); a save that
// can't reach the server is kept on this device and shows "Pending" until it is sent (api.save).
import { tr } from '../i18n.js';
import { $, esc, icon, toast, moneyBox, parseMoney, liveMoney, dateTimeFmt, busy, errorBox } from '../ui.js';
import { save as send, upload, forget } from '../api.js';
import { cardTop, fileChip, progress, slotLine, star, capNote } from '../project.js';

let st = null; // { el, d, can, form: { ref: working copy }, xl: Correct Excel working copy (file or null) }

const pick = (x = {}) => ({ summary: x.summary || '', finalBudget: x.finalBudget ?? null, noted: [...(x.noted || [])] });
const dirty = (ref) => JSON.stringify(st.form[ref]) !== JSON.stringify(pick(st.d.pitch[ref]));
const xlDirty = () => !!st && (st.xl?.id || '') !== (st.d.orgPitch?.file?.id || '');
export const isDirty = () => !!st && (xlDirty() || st.d.projects.some((p) => dirty(p.ref)));
// Unsaved edits, for the leave window: [{ ref, name }] (the Correct Excel box first, ref 'excel').
export const unsaved = () => (st ? [...(xlDirty() ? [{ ref: 'excel', name: tr('Correct Excel file', 'ไฟล์ Excel ฉบับแก้ไข') }] : []),
  ...st.d.projects.filter((p) => dirty(p.ref)).map((p) => ({ ref: p.ref, name: p.name }))] : []);
// "Save and leave": saves each unsaved part; throws at the first that fails.
export async function saveAll() { for (const p of unsaved()) await (p.ref === 'excel' ? saveXl() : save(p.ref)); }

export function render(el, d, ctx) {
  st = { el, d, ctx, can: ctx.canEdit, form: Object.fromEntries(d.projects.map((p) => [p.ref, pick(d.pitch[p.ref])])), xl: d.orgPitch?.file || null };
  el.innerHTML = `<div data-progress></div>${xlBox()}${d.projects.map(card).join('')}`;
  progress(el, d, 'pitching');
  el.oninput = onInput;
  el.onchange = onChange;
  el.onclick = onClick;
}

const orgRef = () => (st.d.projects[0]?.ref || '').replace(/-P\d+$/, '');

// One per organization (v2.9): Upload / Replace / Remove, then Save.
function xlBox() {
  return `<section class="group xl-box"><div class="fb"><h4>${tr('Correct Excel file', 'ไฟล์ Excel ฉบับแก้ไข')}</h4>
    <p class="muted small">${tr('One per organization, for all its projects.', 'หนึ่งไฟล์ต่อองค์กร ใช้กับทุกโครงการ')}</p>
    <div data-xl-slot>${xlSlot()}</div>
    <div class="actions"><span class="muted small" data-xl-saved>${savedLine(st.d.orgPitch)}</span>${st.can ? `<button class="btn primary" data-save-xl>${tr('Save', 'บันทึก')}</button>` : ''}</div></div></section>`;
}
function xlSlot() {
  const f = st.xl;
  return `<div class="upl">${fileChip(f)}${!f && !st.can ? '<span class="muted">—</span>' : ''}
    ${st.can ? `<label class="btn">${icon('upload', 16)}${f ? tr('Replace', 'เปลี่ยนไฟล์') : tr('Upload Excel', 'อัปโหลด Excel')}<input type="file" hidden accept=".xlsx,.xls" data-xl></label>` : ''}
    ${f && st.can ? `<button type="button" class="btn plain" data-xl-clear>${tr('Remove', 'นำออก')}</button>` : ''}</div>${st.can ? capNote('.xlsx / .xls') : ''}`;
}

function card(p) {
  const f = st.form[p.ref];
  const dis = st.can ? '' : 'disabled';
  return `<section class="group pcard" data-ref="${esc(p.ref)}">
    ${cardTop(p, [[tr('Total project budget', 'งบประมาณโครงการ'), p.budget]], '', slotLine(st.d.slots?.[p.ref]))}
    <div class="fb">
      <div class="field"><label>${tr('Summary', 'สรุป')}${star(st.d, 'pitch.summary')}</label><textarea data-k="summary" maxlength="20000" ${dis}>${esc(f.summary)}</textarea></div>
      <div class="field"><label>${tr('Final approved budget', 'งบประมาณอนุมัติขั้นสุดท้าย')}${star(st.d, 'pitch.finalBudget')}</label><input type="text" inputmode="decimal" data-k="finalBudget" value="${esc(moneyBox(f.finalBudget))}" placeholder="0.00" ${dis}></div>
      <div class="field"><label>${tr('Noted documents', 'เอกสารบันทึก')}${star(st.d, 'pitch.noted')}</label><div data-noted>${noted(p.ref)}</div></div>
      <div class="actions"><span class="muted small" data-saved>${savedLine(st.d.pitch[p.ref])}</span>${st.can ? `<button class="btn primary" data-save>${tr('Save', 'บันทึก')}</button>` : ''}</div>
    </div></section>`;
}

const savedLine = (s) => (s?.pending ? `<span class="pill orange">${tr('Pending', 'รอส่ง')}</span> ${tr('Kept on this device; it is sent when the server can be reached', 'เก็บไว้ในเครื่องนี้ จะส่งเมื่อติดต่อเซิร์ฟเวอร์ได้')}`
  : s?.at ? tr('Saved {when} by {who}', 'บันทึก {when} โดย {who}', { when: dateTimeFmt(s.at), who: esc(st.ctx.nameOf(s.by)) }) : tr('Not saved yet', 'ยังไม่ได้บันทึก'));

function noted(ref) {
  const list = st.form[ref].noted;
  return `<div class="upl">${list.map((f, i) => `${fileChip(f)}${st.can ? `<button class="icon-btn" data-rm="${i}" title="${esc(tr('Remove from list', 'นำออกจากรายการ'))}">${icon('x', 16)}</button>` : ''}`).join('')}
    ${!list.length && !st.can ? '<span class="muted">—</span>' : ''}
    ${st.can ? `<label class="btn">${icon('upload', 16)}${tr('Upload .docx or PDF', 'อัปโหลด .docx หรือ PDF')}<input type="file" hidden multiple accept=".docx,.pdf" data-up></label>` : ''}</div>${st.can ? capNote('.docx / PDF') : ''}`;
}

const refOf = (e) => e.target.closest('[data-ref]')?.dataset.ref;
const uploading = (name) => tr('Uploading {file}…', 'กำลังอัปโหลด {file}…', { file: name });
const UPLOADED = () => tr('Uploaded — press Save to keep it', 'อัปโหลดแล้ว — กดบันทึกเพื่อเก็บ');

function onInput(e) {
  const ref = refOf(e); const k = e.target.dataset.k;
  if (!ref || !k) return;
  if (k === 'finalBudget') { liveMoney(e.target); st.form[ref].finalBudget = parseMoney(e.target.value); }
  if (k === 'summary') st.form[ref].summary = e.target.value;
}

async function onChange(e) {
  if (e.target.hasAttribute('data-xl') && e.target.files?.[0]) {
    const file = e.target.files[0];
    try {
      st.xl = await busy(uploading(file.name), (set) => upload(file, { org: orgRef(), purpose: 'excel', only: 'excel' }, set));
      toast(UPLOADED());
    } catch (err) {
      await errorBox(err.message, tr('Couldn’t upload', 'อัปโหลดไม่สำเร็จ'));
    }
    $('[data-xl-slot]', st.el).innerHTML = xlSlot();
    return;
  }
  const ref = refOf(e);
  if (!ref || !e.target.hasAttribute('data-up')) return;
  const box = e.target.closest('[data-noted]');
  let got = 0;
  for (const file of [...e.target.files]) {
    try {
      st.form[ref].noted.push(await busy(uploading(file.name), (set) => upload(file, { ref, purpose: 'noted' }, set)));
      got++;
    } catch (err) {
      await errorBox(`${file.name}: ${err.message}`, tr('Couldn’t upload', 'อัปโหลดไม่สำเร็จ'));
    }
  }
  box.innerHTML = noted(ref);
  if (got) toast(UPLOADED());
}

async function onClick(e) {
  if (e.target.closest('[data-xl-clear]')) { st.xl = null; $('[data-xl-slot]', st.el).innerHTML = xlSlot(); return; }
  if (e.target.closest('[data-save-xl]')) { await run(e.target.closest('button'), saveXl); return; }
  const ref = refOf(e);
  if (!ref) return;
  const rm = e.target.closest('[data-rm]');
  if (rm) {
    st.form[ref].noted.splice(Number(rm.dataset.rm), 1);
    rm.closest('[data-noted]').innerHTML = noted(ref);
    return;
  }
  const btn = e.target.closest('[data-save]');
  if (btn) await run(btn, () => save(ref));
}

// A Save button: the save, then "Saved" / "Nothing changed" / "Pending", or the error window (the edits stay).
async function run(btn, fn) {
  btn.disabled = true;
  try {
    const out = await fn();
    if (out.queued) toast(tr('Kept on this device — it will be sent automatically', 'เก็บไว้ในเครื่องนี้ — จะส่งให้อัตโนมัติ'));
    else toast(out.changed === 0 ? tr('Nothing changed', 'ไม่มีการเปลี่ยนแปลง') : tr('Saved', 'บันทึกแล้ว'));
  } catch (err) {
    await errorBox(err.message);
  }
  btn.disabled = false;
}

// Saves one project's form and repaints its card. Resolves the server's answer, or { queued: true } when the save
// is kept on this device (the card then shows Pending); throws when the save is refused.
async function save(ref) {
  const p = st.d.projects.find((x) => x.ref === ref);
  const prev = st.d.pitch[ref];
  let out;
  try {
    out = await busy(tr('Saving {name}…', 'กำลังบันทึก {name}…', { name: p?.name || ref }),
      () => send('savePitching', { ref, data: st.form[ref], base: prev?.at || '' }, `pitch:${ref}`));
  } catch (e) {
    if (e.code !== 'QUEUED') throw e;
    out = { pitch: { ...st.form[ref], at: prev?.at || '', by: prev?.by || '', pending: true }, queued: true };
  }
  st.d.pitch[ref] = out.pitch;
  progress(st.el, st.d, 'pitching');
  st.ctx.persist?.();
  forget('dashboard');
  st.form[ref] = pick(out.pitch);
  const sec = st.el.querySelector(`.pcard[data-ref="${CSS.escape(ref)}"]`);
  if (sec) $('[data-saved]', sec).innerHTML = savedLine(out.pitch);
  return out;
}

// Saves the organization's Correct Excel file (v2.9).
async function saveXl() {
  const prev = st.d.orgPitch;
  let out;
  try {
    out = await busy(tr('Saving {name}…', 'กำลังบันทึก {name}…', { name: tr('Correct Excel file', 'ไฟล์ Excel ฉบับแก้ไข') }),
      () => send('saveOrgPitching', { org: orgRef(), data: { excel: st.xl }, base: prev?.at || '' }, `xl:${orgRef()}`));
  } catch (e) {
    if (e.code !== 'QUEUED') throw e;
    out = { orgPitch: { file: st.xl, at: prev?.at || '', by: prev?.by || '', pending: true }, queued: true };
  }
  st.d.orgPitch = out.orgPitch;
  st.xl = out.orgPitch?.file || null;
  st.ctx.persist?.();
  $('[data-xl-saved]', st.el).innerHTML = savedLine(out.orgPitch);
  $('[data-xl-slot]', st.el).innerHTML = xlSlot();
  return out;
}
