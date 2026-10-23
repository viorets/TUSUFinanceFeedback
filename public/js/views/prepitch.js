// Pre-pitch: each project as submitted, then a Note form for self-noting (every role can edit; nothing is
// required and nothing is tracked since v2.2). A No on .docx / PDF offers a feedback file, uploaded straight to
// Drive when picked; Save writes the row and the server logs every changed field in PrePitch_Log.
// Top of the tab: the organization's Budget proposal form (.xlsx) from the submission (v2.4), preview + download,
// and under it a feedback box on that form (v2.7): a note + one file, saved with its own Save, logged like the cards.
import { tr } from '../i18n.js';
import { $, esc, icon, toast, moneyBox, parseMoney, liveMoney, dateTimeFmt, num2, dialog, errorBox, busy } from '../ui.js';
import { save as send, upload, forget } from '../api.js';
import { cardTop, dlBtn, fileChip, capNote, newFiles, markCard, kept, earlierNotice } from '../project.js';
import { ORGS } from '../masterdata.js';

let st = null; // { el, d, can, form: { ref: working copy }, box: budget-form feedback working copy }

const blank = { approvedBudget: null, docxOk: '', docxFile: null, pdfOk: '', pdfFile: null, note: '', docs: [] };
const pick = (x) => ({ approvedBudget: x.approvedBudget ?? null, docxOk: x.docxOk || '', docxFile: x.docxFile || null, pdfOk: x.pdfOk || '', pdfFile: x.pdfFile || null, note: x.note || '', docs: [...(x.docs || [])] });
const dirty = (ref) => JSON.stringify(st.form[ref]) !== JSON.stringify(pick(st.d.pre[ref] || blank));
const pickBox = (x) => ({ note: x?.note || '', file: x?.file || null });
const boxDirty = () => !!st && JSON.stringify(st.box) !== JSON.stringify(pickBox(st.d.orgPre));
export const isDirty = () => !!st && (boxDirty() || st.d.projects.some((p) => dirty(p.ref)));
// Unsaved edits, for the leave window: [{ ref, code, name, files }] (the budget-form box first, ref 'form', shown
// under the organization's code); files = the names of files not kept with Save yet.
const BOX = () => tr('Feedback on the budget form', 'ข้อเสนอแนะต่อแบบฟอร์มงบประมาณ');
const filesOf = (x) => [x?.docxFile, x?.pdfFile, ...(x?.docs || [])].filter(Boolean);
const projNew = (ref, work = st.form[ref]) => newFiles(filesOf(work), filesOf(st.d.pre[ref])).map((f) => f.name);
const boxNew = (work = st.box) => newFiles(work?.file, st.d.orgPre?.file).map((f) => f.name);
export const unsaved = () => (st ? [...(boxDirty() ? [{ ref: 'form', code: orgRef(), name: BOX(), files: boxNew() }] : []),
  ...st.d.projects.filter((p) => dirty(p.ref)).map((p) => ({ ref: p.ref, name: p.name, files: projNew(p.ref) }))] : []);
// Leave without saving: what was kept on this device for this tab goes too.
export const discard = () => { if (st) kept.write(keptKey(), null); };
// "Save and leave": saves each unsaved part; throws at the first that fails (the rest stay unsaved).
export async function saveAll() { for (const p of unsaved()) await (p.ref === 'form' ? saveBox() : save(p.ref)); }

// restored: unsaved work kept on this device and put back ({ parts: { ref: working copy }, box? }).
export function render(el, d, ctx, restored = null) {
  st = { el, d, ctx, can: ctx.canEdit, form: Object.fromEntries(d.projects.map((p) => [p.ref, restored?.parts?.[p.ref] || pick(d.pre[p.ref] || blank)])),
    box: restored?.box || pickBox(d.orgPre) };
  el.innerHTML = `<div class="notice blue">${icon('alert', 18)}<span>${tr('This section is for your own notes. You don’t have to fill it in.', 'ส่วนนี้มีไว้สำหรับจดบันทึกของคุณเอง ไม่จำเป็นต้องกรอก')}</span></div>${budgetForm(d)}${d.projects.map(card).join('')}`;
  el.oninput = (e) => { onInput(e); marks(); };
  el.onchange = async (e) => { await onChange(e); marks(); };
  el.onclick = async (e) => { await onClick(e); marks(); };
  marks();
  if (!restored) offerKept();
}

// ---------- unsaved work (v2.13) ----------
// "Not saved" on each card / box with unsaved changes, and the unsaved parts kept on this device.
function marks() {
  if (!st?.el.isConnected) return;
  st.el.querySelectorAll('.pcard').forEach((sec) => markCard(sec, dirty(sec.dataset.ref)));
  markCard(st.el.querySelector('.bform'), boxDirty());
  keepSoon();
}
const keptKey = () => `unsaved:${st.ctx.me?.email || ''}:${orgRef()}:prepitch`;
let keepT = 0;
const keepSoon = () => { clearTimeout(keepT); keepT = setTimeout(keepNow, 400); };
function keepNow() {
  if (!st?.can || !st.el.isConnected) return;
  const parts = {};
  st.d.projects.forEach((p) => { if (dirty(p.ref)) parts[p.ref] = { work: st.form[p.ref], base: st.d.pre[p.ref]?.at || '' }; });
  let box = boxDirty() ? { work: st.box, base: st.d.orgPre?.at || '' } : null;
  if (st.offered) {
    const old = kept.read(keptKey()) || {};
    Object.entries(old.parts || {}).forEach(([r, v]) => { if (!parts[r]) parts[r] = v; });
    box = box || old.box || null;
  }
  kept.write(keptKey(), Object.keys(parts).length || box ? { at: new Date().toISOString(), parts, box } : null);
}
document.addEventListener('visibilitychange', () => { if (document.hidden) keepNow(); });
addEventListener('pagehide', () => keepNow());

function offerKept() {
  const k = kept.read(keptKey());
  if (!st.can || !k || (!Object.keys(k.parts || {}).length && !k.box)) return;
  st.offered = true;
  const nameOf = (r) => st.d.projects.find((p) => p.ref === r)?.name || '';
  const line = (code, name, files) => `${code} · ${name}${files.length ? ` — ${files.join(', ')}` : ''}`;
  const names = [...(k.box ? [line(orgRef(), BOX(), boxNew(k.box.work))] : []),
    ...Object.entries(k.parts || {}).map(([r, v]) => line(r, nameOf(r), projNew(r, v.work)))];
  earlierNotice(st.el, { at: k.at, names }, () => restoreKept(k), () => { st.offered = false; kept.write(keptKey(), null); keepNow(); });
}
function restoreKept(k) {
  const parts = {}; const skipped = [];
  Object.entries(k.parts || {}).forEach(([r, v]) => {
    if (!st.d.projects.some((p) => p.ref === r)) return;
    if ((st.d.pre[r]?.at || '') === v.base) parts[r] = v.work; else skipped.push(r);
  });
  const restored = { parts };
  if (k.box) { if ((st.d.orgPre?.at || '') === k.box.base) restored.box = k.box.work; else skipped.push(BOX()); }
  st.offered = false;
  render(st.el, st.d, st.ctx, restored);
  if (skipped.length) {
    errorBox(tr('Someone saved these on the site after your changes were kept, so they were not restored: {list}', 'มีผู้บันทึกรายการเหล่านี้บนเว็บไซต์หลังจากเก็บการแก้ไขของคุณไว้ จึงไม่ได้กู้คืน: {list}', { list: skipped.join(', ') }),
      tr('Some changes were not restored', 'การแก้ไขบางส่วนไม่ได้กู้คืน'));
  } else toast(tr('Restored — press Save to keep them', 'กู้คืนแล้ว — กดบันทึกเพื่อเก็บ'));
}

// One per organization: the Budget proposal form (.xlsx) uploaded with the submission (a Drive link).
function budgetForm(d) {
  const src = d.org?.budgetForm || '';
  return `<section class="group budget-form"><div class="row"><span class="grow"><b>${tr('Budget proposal form (.xlsx)', 'แบบฟอร์มรายการงบประมาณ (.xlsx)')}</b>
    <small class="muted">${src ? tr('One per organization, from the submission', 'หนึ่งไฟล์ต่อองค์กร จากการส่งข้อมูล') : tr('No file submitted', 'ยังไม่มีไฟล์ที่ส่งเข้ามา')}</small></span>
    ${dlBtn(src, tr('Budget proposal form', 'แบบฟอร์มงบประมาณ'))}</div>
    <div class="fb bform"><h4>${tr('Feedback on the budget form', 'ข้อเสนอแนะต่อแบบฟอร์มงบประมาณ')}</h4>
      <div class="field"><label>${tr('Note', 'บันทึก')}</label><textarea data-box="note" maxlength="5000" ${st.can ? '' : 'disabled'}>${esc(st.box.note)}</textarea></div>
      <div class="field"><label>${tr('Feedback file', 'ไฟล์ข้อเสนอแนะ')}</label><div data-slot="box">${boxSlot()}</div></div>
      <div class="actions"><span class="muted small" data-box-saved>${savedLine(d.orgPre)}</span>${st.can ? `<button class="btn primary" data-save-box>${tr('Save', 'บันทึก')}</button>` : ''}</div>
    </div></section>`;
}

// The budget-form feedback file: .xlsx, .xls, .docx or PDF.
function boxSlot() {
  const f = st.box.file;
  return `<div class="upl">${fileChip(f, { unsaved: !!f && f.id !== st.d.orgPre?.file?.id })}${st.can ? `<label class="btn">${icon('upload', 16)}${f ? tr('Replace', 'เปลี่ยนไฟล์') : tr('Upload feedback file', 'อัปโหลดไฟล์ข้อเสนอแนะ')}
    <input type="file" hidden accept=".xlsx,.xls,.docx,.pdf" data-up="box"></label>` : ''}${f && st.can ? `<button type="button" class="btn plain" data-box-clear>${tr('Remove', 'นำออก')}</button>` : ''}</div>${st.can ? capNote('.xlsx / .xls / .docx / PDF') : ''}`;
}

async function saveBox() {
  const prev = st.d.orgPre;
  let out;
  try {
    out = await busy(tr('Saving {name}…', 'กำลังบันทึก {name}…', { name: tr('Feedback on the budget form', 'ข้อเสนอแนะต่อแบบฟอร์มงบประมาณ') }),
      () => send('saveOrgPrePitch', { org: orgRef(), data: st.box, base: prev?.at || '' }, `form:${orgRef()}`));
  } catch (e) {
    if (e.code !== 'QUEUED') throw e;
    out = { orgPre: { ...st.box, at: prev?.at || '', by: prev?.by || '', pending: true }, queued: true };
  }
  st.d.orgPre = out.orgPre;
  st.box = pickBox(out.orgPre);
  const line = $('[data-box-saved]', st.el);
  if (line) line.innerHTML = savedLine(out.orgPre);
  const bs = $('[data-slot="box"]', st.el); if (bs) bs.innerHTML = boxSlot();
  marks();
  st.ctx.persist?.();
  return out;
}
const orgRef = () => (st.d.projects[0]?.ref || '').replace(/-P\d+$/, '');

function card(p) {
  const f = st.form[p.ref];
  const saved = st.d.pre[p.ref];
  const dis = st.can ? '' : 'disabled';
  const yn = (k) => `<div class="seg" role="radiogroup">${['Yes', 'No'].map((v) => `<label><input type="radio" name="${k}-${esc(p.ref)}" value="${v}" data-k="${k}Ok" ${f[`${k}Ok`] === v ? 'checked' : ''} ${dis}><span>${v === 'Yes' ? tr('Yes', 'ใช่') : tr('No', 'ไม่')}</span></label>`).join('')}</div>`;
  return `<section class="group pcard" data-ref="${esc(p.ref)}">
    ${cardTop(p, [[tr('Total proposed project budget', 'งบประมาณโครงการที่เสนอ'), p.budget]], `${dlBtn(p.pdf, tr('Project PDF', 'โครงการ PDF'))}${dlBtn(p.docx, tr('Project .docx', 'โครงการ .docx'))}`)}
    <div class="fb"><h4>${tr('Note', 'บันทึก')}</h4>
      <div class="field"><label>${tr('Estimated Approved budget', 'งบประมาณที่คาดว่าจะได้รับอนุมัติ')}</label><input type="text" inputmode="decimal" data-k="approvedBudget" value="${esc(moneyBox(f.approvedBudget))}" placeholder="0.00" ${dis}></div>
      <div class="field"><label>${tr('Approved project .docx?', 'อนุมัติไฟล์โครงการ .docx?')}</label><div>${yn('docx')}<div data-slot="docx">${slot(p.ref, 'docx')}</div></div></div>
      <div class="field"><label>${tr('Approved project PDF?', 'อนุมัติไฟล์โครงการ PDF?')}</label><div>${yn('pdf')}<div data-slot="pdf">${slot(p.ref, 'pdf')}</div></div></div>
      <div class="field"><label>${tr('Note', 'บันทึก')}</label><textarea data-k="note" maxlength="5000" ${dis}>${esc(f.note)}</textarea></div>
      <div class="field"><label>${tr('Documents', 'เอกสาร')}</label><div data-docs>${docs(p.ref)}</div></div>
      ${st.can ? `<div class="actions"><span class="muted small" data-saved>${savedLine(saved)}</span><button class="btn primary" data-save>${tr('Save', 'บันทึก')}</button></div>`
    : `<div class="actions"><span class="muted small">${savedLine(saved)}</span></div>`}
    </div></section>`;
}

const QUEUED_TOAST = () => tr('Kept on this device — it will be sent automatically', 'เก็บไว้ในเครื่องนี้ — จะส่งให้อัตโนมัติ');
const savedLine = (s) => (s?.pending ? `<span class="pill orange">${tr('Pending', 'รอส่ง')}</span> ${tr('Kept on this device; it is sent when the server can be reached', 'เก็บไว้ในเครื่องนี้ จะส่งเมื่อติดต่อเซิร์ฟเวอร์ได้')}` : s?.at ? tr('Saved {when} by {who}', 'บันทึก {when} โดย {who}', { when: dateTimeFmt(s.at), who: esc(st.ctx.nameOf(s.by)) }) : tr('Not saved yet', 'ยังไม่ได้บันทึก'));

// Upload area under a Yes/No: shown only for No.
function slot(ref, k) {
  const f = st.form[ref];
  if (f[`${k}Ok`] !== 'No') return '';
  const file = f[`${k}File`];
  const isNew = !!file && file.id !== st.d.pre[ref]?.[`${k}File`]?.id;
  const label = k === 'docx' ? tr('Upload feedback .docx', 'อัปโหลดไฟล์ข้อเสนอแนะ .docx') : tr('Upload feedback PDF', 'อัปโหลดไฟล์ข้อเสนอแนะ PDF');
  return `<div class="upl">${fileChip(file, { unsaved: isNew })}${st.can ? `<label class="btn">${icon('upload', 16)}${file ? tr('Replace', 'เปลี่ยนไฟล์') : esc(label)}
    <input type="file" hidden accept="${k === 'docx' ? '.docx' : '.pdf'}" data-up="${k}"></label>` : ''}
    </div>${st.can ? capNote(k === 'docx' ? '.docx' : 'PDF') : ''}`;
}

// Documents (v2.12): any number of .docx / PDF per project, uploaded when picked; Save keeps the list.
function docs(ref) {
  const list = st.form[ref].docs;
  const fresh = new Set(newFiles(list, st.d.pre[ref]?.docs).map((f) => f.id));
  return `<div class="upl">${list.map((f, i) => `${fileChip(f, { unsaved: fresh.has(f.id) })}${st.can ? `<button type="button" class="icon-btn" data-rm-doc="${i}" title="${esc(tr('Remove from list', 'นำออกจากรายการ'))}">${icon('x', 16)}</button>` : ''}`).join('')}
    ${!list.length && !st.can ? '<span class="muted">—</span>' : ''}
    ${st.can ? `<label class="btn">${icon('upload', 16)}${tr('Upload .docx or PDF', 'อัปโหลด .docx หรือ PDF')}<input type="file" hidden multiple accept=".docx,.pdf" data-up-docs></label>` : ''}</div>${st.can ? capNote('.docx / PDF') : ''}`;
}

const refOf = (e) => e.target.closest('[data-ref]')?.dataset.ref;

function onInput(e) {
  if (e.target.dataset.box === 'note') { st.box.note = e.target.value; return; }
  const ref = refOf(e); const k = e.target.dataset.k;
  if (!ref || !k) return;
  if (k === 'approvedBudget') { liveMoney(e.target); st.form[ref].approvedBudget = parseMoney(e.target.value); }
  if (k === 'note') st.form[ref].note = e.target.value;
}

async function onChange(e) {
  if (e.target.dataset.up === 'box') { await boxUpload(e); return; }
  const ref = refOf(e);
  if (!ref) return;
  if (e.target.hasAttribute('data-up-docs')) {
    const box = e.target.closest('[data-docs]');
    let got = 0;
    for (const file of [...e.target.files]) {
      try {
        st.form[ref].docs.push(await busy(tr('Uploading {file}…', 'กำลังอัปโหลด {file}…', { file: file.name }), (set) => upload(file, { ref, purpose: 'docs' }, set)));
        got++;
      } catch (err) {
        await errorBox(`${file.name}: ${err.message}`, tr('Couldn’t upload', 'อัปโหลดไม่สำเร็จ'));
      }
    }
    box.innerHTML = docs(ref);
    if (got) toast(tr('Uploaded — press Save to keep it', 'อัปโหลดแล้ว — กดบันทึกเพื่อเก็บ'));
    return;
  }
  const k = e.target.dataset.k;
  if (k === 'docxOk' || k === 'pdfOk') {
    st.form[ref][k] = e.target.value;
    const kind = k.slice(0, -2);
    $(`[data-slot="${kind}"]`, e.target.closest('.pcard')).innerHTML = slot(ref, kind);
    return;
  }
  const kind = e.target.dataset.up;
  const file = e.target.files?.[0];
  if (!kind || !file) return;
  const box = e.target.closest('[data-slot]');
  try {
    st.form[ref][`${kind}File`] = await busy(tr('Uploading {file}…', 'กำลังอัปโหลด {file}…', { file: file.name }), (set) => upload(file, { ref, purpose: 'feedback', only: kind }, set));
    toast(tr('Uploaded — press Save to keep it', 'อัปโหลดแล้ว — กดบันทึกเพื่อเก็บ'));
  } catch (err) {
    await errorBox(err.message, tr('Couldn’t upload', 'อัปโหลดไม่สำเร็จ'));
  }
  box.innerHTML = slot(ref, kind);
}

async function boxUpload(e) {
  const file = e.target.files?.[0];
  if (!file) return;
  const box = e.target.closest('[data-slot]');
  try {
    st.box.file = await busy(tr('Uploading {file}…', 'กำลังอัปโหลด {file}…', { file: file.name }), (set) => upload(file, { org: orgRef(), purpose: 'form', only: 'any' }, set));
    toast(tr('Uploaded — press Save to keep it', 'อัปโหลดแล้ว — กดบันทึกเพื่อเก็บ'));
  } catch (err) {
    await errorBox(err.message, tr('Couldn’t upload', 'อัปโหลดไม่สำเร็จ'));
  }
  box.innerHTML = boxSlot();
}

// Saves one project's form and repaints its card. Returns the server's answer; throws when the save fails.
async function save(ref) {
  const prev = st.d.pre[ref];
  const name = st.d.projects.find((x) => x.ref === ref)?.name || ref;
  let out;
  try {
    out = await busy(tr('Saving {name}…', 'กำลังบันทึก {name}…', { name }),
      () => send('savePrePitch', { ref, data: st.form[ref], base: prev?.at || '' }, `pre:${ref}`));
  } catch (e) {
    if (e.code !== 'QUEUED') throw e;
    out = { pre: { ...st.form[ref], at: prev?.at || '', by: prev?.by || '', pending: true }, queued: true };
  }
  const sec = st.el.querySelector(`.pcard[data-ref="${CSS.escape(ref)}"]`);
  st.d.pre[ref] = out.pre;
  st.form[ref] = pick(out.pre);
  if (sec) {
    $('[data-saved]', sec).innerHTML = savedLine(out.pre);
    ['docx', 'pdf'].forEach((k) => { const x = $(`[data-slot="${k}"]`, sec); if (x) x.innerHTML = slot(ref, k); });
    const dx = $('[data-docs]', sec); if (dx) dx.innerHTML = docs(ref);
  }
  marks();
  forget('boot'); // list counts
  forget('dashboard');
  st.ctx.persist?.();
  return out;
}

// Shown after a save that changed something; stays until OK or Cancel (both just close it).
function savedWindow(ref) {
  const p = st.d.projects.find((x) => x.ref === ref);
  const code = ref.replace(/-P\d+$/, '');
  const org = ORGS.find((o) => o.code === code);
  const v = st.d.pre[ref]?.approvedBudget;
  const kv = [
    [tr('Organization', 'องค์กร'), org?.en || code],
    [tr('Project number', 'รหัสโครงการ'), ref],
    [tr('Project name', 'ชื่อโครงการ'), p?.name || '—'],
    [tr('Estimated Approved budget', 'งบประมาณที่คาดว่าจะได้รับอนุมัติ'), v == null ? '—' : `฿${num2(v)}`],
  ];
  return dialog(tr('Saved', 'บันทึกแล้ว'), `<dl class="kv saved-kv">${kv.map(([k, x]) => `<dt>${esc(k)}</dt><dd>${esc(x)}</dd>`).join('')}</dl>`,
    [{ v: false, label: tr('Cancel', 'ยกเลิก') }, { v: true, label: tr('OK', 'ตกลง'), cls: 'primary' }], { sticky: true });
}

async function onClick(e) {
  if (e.target.closest('[data-box-clear]')) { st.box.file = null; $('[data-slot="box"]', st.el).innerHTML = boxSlot(); return; }
  const bb = e.target.closest('[data-save-box]');
  if (bb) {
    bb.disabled = true;
    try {
      const out = await saveBox();
      toast(out.queued ? QUEUED_TOAST() : out.changed ? tr('Saved', 'บันทึกแล้ว') : tr('Nothing changed', 'ไม่มีการเปลี่ยนแปลง'));
    } catch (err) {
      await errorBox(err.message);
    }
    bb.disabled = false;
    return;
  }
  const rmDoc = e.target.closest('[data-rm-doc]');
  if (rmDoc) {
    const r = refOf(e);
    st.form[r].docs.splice(Number(rmDoc.dataset.rmDoc), 1);
    rmDoc.closest('[data-docs]').innerHTML = docs(r);
    return;
  }
  const btn = e.target.closest('[data-save]');
  if (!btn) return;
  const ref = refOf(e);
  btn.disabled = true;
  try {
    const out = await save(ref);
    if (out.queued) toast(QUEUED_TOAST());
    else if (out.changed) await savedWindow(ref);
    else toast(tr('Nothing changed', 'ไม่มีการเปลี่ยนแปลง'));
  } catch (err) {
    await errorBox(err.message); // the edits stay on screen
  }
  btn.disabled = false;
}
