// Pitching: at the top, the organization's "Correct Excel file" (one .xlsx / .xls per organization since v2.9, with
// its own Save, logged in Pitching_Log). Then per project: the pitching slot (from the scheduling site, read only),
// the summary, the final approved budget and "Noted documents" (v2.15: one Word file, .docx or .doc, and one PDF,
// each may be empty; Upload / Replace / Remove as on Pre-pitch). Unfinished = a
// required question unanswered (project.js). Every save and upload runs in a progress window (busy); a save that
// can't reach the server is kept on this device and shows "Pending" until it is sent (api.save).
import { tr } from '../i18n.js';
import { $, esc, icon, toast, moneyBox, parseMoney, liveMoney, dateTimeFmt, busy, errorBox, moneyOr } from '../ui.js';
import { save as send, upload, forget, call } from '../api.js';
import { cardTop, dlBtn, fileChip, progress, slotLine, star, capNote, editBtn, onMarked, newFiles, markCard, kept, earlierNotice, UNFINISHED } from '../project.js';

// "Show Pre-pitch data" (v2.12): each card also shows its Pre-pitch notes, read only. Kept on this device.
// "Show Pre-pitch data" (v2.13): one toggle per project card (and one for the Correct Excel box, key 'xl'); open ones
// stay open while this page is open.
const openPre = new Set();

let st = null; // { el, d, can, form: { ref: working copy }, xl: Correct Excel working copy (file or null) }

// Noted documents (v2.15): 'word' (.docx / .doc) or 'pdf' by the file name; the list is [Word, PDF], either may be
// missing (the server keeps the same pair: Code.gs notedPair_).
const NK = (f) => (/\.pdf$/i.test(f?.name || '') ? 'pdf' : /\.docx?$/i.test(f?.name || '') ? 'word' : '');
const notedOf = (list, k) => (list || []).filter((f) => NK(f) === k).pop() || null;
const pair = (list) => [notedOf(list, 'word'), notedOf(list, 'pdf')].filter(Boolean);
const setNoted = (ref, k, file) => { st.form[ref].noted = pair([...st.form[ref].noted.filter((f) => NK(f) !== k), ...(file ? [file] : [])]); };
const pick = (x = {}) => ({ summary: x.summary || '', finalBudget: x.finalBudget ?? null, noted: pair(x.noted) });
const dirty = (ref) => JSON.stringify(st.form[ref]) !== JSON.stringify(pick(st.d.pitch[ref]));
const xlDirty = () => !!st && (st.xl?.id || '') !== (st.d.orgPitch?.file?.id || '');
export const isDirty = () => !!st && (xlDirty() || st.d.projects.some((p) => dirty(p.ref)));
// Unsaved edits, for the leave window: [{ ref, code, name, files }] (the Correct Excel box first, ref 'excel', shown
// under the organization's code); files = the names of files not kept with Save yet.
const XL = () => tr('Correct Excel file', 'ไฟล์ Excel ฉบับแก้ไข');
const notedNew = (ref, work = st.form[ref]) => newFiles(work.noted, st.d.pitch[ref]?.noted).map((f) => f.name);
export const unsaved = () => (st ? [...(xlDirty() ? [{ ref: 'excel', code: orgRef(), name: XL(), files: st.xl ? [st.xl.name] : [] }] : []),
  ...st.d.projects.filter((p) => dirty(p.ref)).map((p) => ({ ref: p.ref, name: p.name, files: notedNew(p.ref) }))] : []);
// Leave without saving: what was kept on this device for this tab goes too.
export const discard = () => { if (st) kept.write(keptKey(), null); };
// "Save and leave": saves each unsaved part; throws at the first that fails.
export async function saveAll() { for (const p of unsaved()) await (p.ref === 'excel' ? saveXl() : save(p.ref)); }

// restored: unsaved work kept on this device and put back ({ parts: { ref: working copy }, xl? }).
export function render(el, d, ctx, restored = null) {
  st = { el, d, ctx, can: ctx.canEdit, form: Object.fromEntries(d.projects.map((p) => [p.ref, restored?.parts?.[p.ref] ? { ...restored.parts[p.ref], noted: pair(restored.parts[p.ref].noted) } : pick(d.pitch[p.ref])])),
    xl: restored && 'xl' in restored ? restored.xl : d.orgPitch?.file || null };
  el.innerHTML = `<div data-progress></div>${xlBox()}${d.projects.map(card).join('')}`;
  progress(el, d, 'pitching');
  el.oninput = (e) => { onInput(e); marks(); };
  el.onchange = async (e) => { await onChange(e); marks(); };
  el.onclick = async (e) => { await onClick(e); marks(); };
  onMarked(marked);
  marks();
  if (!restored) offerKept();
}

// ---------- unsaved work (v2.13) ----------
// "Not saved" on each card / box with unsaved changes, and the unsaved parts kept on this device.
function marks() {
  if (!st?.el.isConnected) return;
  st.el.querySelectorAll('.pcard').forEach((sec) => markCard(sec, dirty(sec.dataset.ref)));
  markCard(st.el.querySelector('.xl-box'), xlDirty());
  keepSoon();
}
const keptKey = () => `unsaved:${st.ctx.me?.email || ''}:${orgRef()}:pitching`;
let keepT = 0;
const keepSoon = () => { clearTimeout(keepT); keepT = setTimeout(keepNow, 400); };
// The dirty parts with the saved time they started from (base). While the earlier copy hasn't been answered, its
// parts are kept too.
function keepNow() {
  if (!st?.can || !st.el.isConnected) return;
  const parts = {};
  st.d.projects.forEach((p) => { if (dirty(p.ref)) parts[p.ref] = { work: st.form[p.ref], base: st.d.pitch[p.ref]?.at || '' }; });
  let xl = xlDirty() ? { work: st.xl, base: st.d.orgPitch?.at || '' } : null;
  if (st.offered) {
    const old = kept.read(keptKey()) || {};
    Object.entries(old.parts || {}).forEach(([r, v]) => { if (!parts[r]) parts[r] = v; });
    xl = xl || old.xl || null;
  }
  kept.write(keptKey(), Object.keys(parts).length || xl ? { at: new Date().toISOString(), parts, xl } : null);
}
document.addEventListener('visibilitychange', () => { if (document.hidden) keepNow(); });
addEventListener('pagehide', () => keepNow());

// Work kept from an earlier visit: Restore (only parts nobody saved since) or Discard.
function offerKept() {
  const k = kept.read(keptKey());
  if (!st.can || !k || (!Object.keys(k.parts || {}).length && !k.xl)) return;
  st.offered = true;
  const nameOf = (r) => st.d.projects.find((p) => p.ref === r)?.name || '';
  const names = [...(k.xl ? [`${orgRef()} · ${XL()}${k.xl.work ? ` — ${k.xl.work.name}` : ''}`] : []),
    ...Object.entries(k.parts || {}).map(([r, v]) => { const f = notedNew(r, v.work); return `${r} · ${nameOf(r)}${f.length ? ` — ${f.join(', ')}` : ''}`; })];
  earlierNotice(st.el, { at: k.at, names }, () => restoreKept(k), () => { st.offered = false; kept.write(keptKey(), null); keepNow(); });
}
function restoreKept(k) {
  const parts = {}; const skipped = [];
  Object.entries(k.parts || {}).forEach(([r, v]) => {
    if (!st.d.projects.some((p) => p.ref === r)) return;
    if ((st.d.pitch[r]?.at || '') === v.base) parts[r] = v.work; else skipped.push(r);
  });
  const restored = { parts };
  if (k.xl) { if ((st.d.orgPitch?.at || '') === k.xl.base) restored.xl = k.xl.work; else skipped.push(XL()); }
  st.offered = false;
  render(st.el, st.d, st.ctx, restored);
  if (skipped.length) {
    errorBox(tr('Someone saved these on the site after your changes were kept, so they were not restored: {list}', 'มีผู้บันทึกรายการเหล่านี้บนเว็บไซต์หลังจากเก็บการแก้ไขของคุณไว้ จึงไม่ได้กู้คืน: {list}', { list: skipped.join(', ') }),
      tr('Some changes were not restored', 'การแก้ไขบางส่วนไม่ได้กู้คืน'));
  } else toast(tr('Restored — press Save to keep them', 'กู้คืนแล้ว — กดบันทึกเพื่อเก็บ'));
}

// A file saved by the editor tab (v2.12): it replaces the Correct Excel file, or the Noted Word file / PDF of its slot
// (v2.15; a slot from before v2.15 goes by the file's type). Since v2.16 the editor has saved it into the card
// already (pitch / orgPitch: the card as saved): only the card's other unsaved edits stay unsaved here.
function marked({ org, slot, file, pitch, orgPitch }) {
  if (!st?.el.isConnected || !st.can || org !== orgRef()) return false;
  if (slot === 'xl') {
    st.xl = file;
    if (orgPitch) { st.d.orgPitch = orgPitch; st.ctx.persist?.(); $('[data-xl-saved]', st.el).innerHTML = savedLine(orgPitch); }
    $('[data-xl-slot]', st.el).innerHTML = xlSlot();
  } else {
    const m = String(slot).match(/^noted:([^:]+)(?::(pdf|docx))?$/);
    const ref = m?.[1];
    const k = m?.[2] ? (m[2] === 'pdf' ? 'pdf' : 'word') : NK(file);
    const sec = ref && k && st.form[ref] && st.el.querySelector(`.pcard[data-ref="${CSS.escape(ref)}"]`);
    if (!sec) return false;
    setNoted(ref, k, file);
    if (pitch) {
      st.d.pitch[ref] = pitch; st.ctx.persist?.(); forget('dashboard'); progress(st.el, st.d, 'pitching');
      $('[data-saved]', sec).innerHTML = savedLine(pitch); $('[data-finish]', sec).innerHTML = finishRow(ref);
    }
    $('[data-noted]', sec).innerHTML = noted(ref);
  }
  toast(pitch || orgPitch ? tr('Edited file saved', 'บันทึกไฟล์ที่แก้ไขแล้ว') : tr('Edited file attached — press Save to keep it', 'แนบไฟล์ที่แก้ไขแล้ว — กดบันทึกเพื่อเก็บ'));
  marks();
  return true;
}

const orgRef = () => (st.d.projects[0]?.ref || '').replace(/-P\d+$/, '');

// One per organization (v2.9): Upload / Replace / Remove, then Save.
function xlBox() {
  return `<section class="group xl-box"><div class="fb"><h4>${tr('Correct Excel file', 'ไฟล์ Excel ฉบับแก้ไข')}</h4>
    <p class="muted small">${tr('One per organization, for all its projects.', 'หนึ่งไฟล์ต่อองค์กร ใช้กับทุกโครงการ')}</p>
    <div class="dl-row src-row"><span class="muted small">${tr('Submitted form', 'แบบฟอร์มที่ส่งมา')}</span>${dlBtn(st.d.org?.budgetForm || '', tr('Budget proposal form', 'แบบฟอร์มงบประมาณ'))}
      ${st.can ? editBtn(st.d.org?.budgetForm || '', { org: orgRef(), slot: 'xl', kind: 'xlsx', master: st.d.org?.budgetForm }) : ''}</div>
    <div data-pre-form>${preForm()}</div>
    <div data-xl-slot>${xlSlot()}</div>
    <div class="actions"><span class="muted small" data-xl-saved>${savedLine(st.d.orgPitch)}</span>${st.can ? `<button class="btn primary" data-save-xl>${tr('Save', 'บันทึก')}</button>` : ''}</div></div></section>`;
}
function xlSlot() {
  const f = st.xl;
  return `<div class="upl">${fileChip(f, { unsaved: xlDirty() })}${st.can ? editBtn(f, { org: orgRef(), slot: 'xl', plain: true, master: st.d.org?.budgetForm }) : ''}${!f && !st.can ? '<span class="muted">—</span>' : ''}
    ${st.can ? `<label class="btn">${icon('upload', 16)}${f ? tr('Replace', 'เปลี่ยนไฟล์') : tr('Upload Excel', 'อัปโหลด Excel')}<input type="file" hidden accept=".xlsx,.xls" data-xl></label>` : ''}
    ${f && st.can ? `<button type="button" class="btn plain" data-xl-clear>${tr('Remove', 'นำออก')}</button>` : ''}</div>${st.can ? capNote('.xlsx / .xls') : ''}`;
}

function card(p) {
  const f = st.form[p.ref];
  const dis = st.can ? '' : 'disabled';
  const files = `${dlBtn(p.pdf, tr('Project PDF', 'โครงการ PDF'))}${st.can ? editBtn(p.pdf, { org: orgRef(), slot: `noted:${p.ref}:pdf`, kind: 'pdf', master: p.pdf }) : ''}${dlBtn(p.docx, tr('Project .docx', 'โครงการ .docx'))}${st.can ? editBtn(p.docx, { org: orgRef(), slot: `noted:${p.ref}:docx`, kind: 'docx', master: p.docx }) : ''}`;
  return `<section class="group pcard" data-ref="${esc(p.ref)}">
    ${cardTop(p, [[tr('Total project budget', 'งบประมาณโครงการ'), p.budget]], files, slotLine(st.d.slots?.[p.ref]))}
    <div data-pre>${prePeek(p.ref)}</div>
    <div class="fb">
      <div class="field"><label>${tr('Summary', 'สรุป')}${star(st.d, 'pitch.summary')}</label><textarea data-k="summary" maxlength="20000" ${dis}>${esc(f.summary)}</textarea></div>
      <div class="field"><label>${tr('Final approved budget', 'งบประมาณอนุมัติขั้นสุดท้าย')}${star(st.d, 'pitch.finalBudget')}</label><input type="text" inputmode="decimal" data-k="finalBudget" value="${esc(moneyBox(f.finalBudget))}" placeholder="0.00" ${dis}></div>
      <div class="field"><label>${tr('Noted documents', 'เอกสารบันทึก')}${star(st.d, 'pitch.noted')}</label><div data-noted>${noted(p.ref)}</div></div>
      <div class="actions"><span class="muted small" data-saved>${savedLine(st.d.pitch[p.ref])}</span>${st.can ? `<button class="btn primary" data-save>${tr('Save', 'บันทึก')}</button>` : ''}</div>
      <div class="actions finish-row" data-finish>${finishRow(p.ref)}</div>
    </div></section>`;
}

// "Mark as finished" (v2.16): the staff confirm the project is finished (after its required answers are saved).
function finishRow(ref) {
  const f = st.d.pitch[ref]?.finished;
  if (f) {
    return `<span class="pill green">${icon('check', 13)}${tr('Finished', 'เสร็จสิ้น')}</span><span class="muted small grow">${tr('{when} by {who}', '{when} โดย {who}', { when: dateTimeFmt(f.at), who: esc(st.ctx.nameOf(f.by)) })}</span>
      ${st.can ? `<button type="button" class="btn plain" data-finish-set="0">${tr('Mark as not finished', 'ยกเลิกเสร็จสิ้น')}</button>` : ''}`;
  }
  return st.can ? `<button type="button" class="btn" data-finish-set="1">${icon('check', 16)}${tr('Mark as finished', 'ทำเครื่องหมายว่าเสร็จสิ้น')}</button>` : '';
}
async function setFinished(ref, finished, btn) {
  const p = st.d.projects.find((x) => x.ref === ref);
  if (finished && dirty(ref)) return errorBox(tr('Save this project first.', 'กรุณาบันทึกโครงการนี้ก่อน'), tr('Not saved yet', 'ยังไม่ได้บันทึก'));
  if (finished && UNFINISHED.pitching(st.d, p)) return errorBox(tr('Answer the required questions (*) and save first.', 'กรุณาตอบคำถามที่จำเป็น (*) และบันทึกก่อน'), tr('Not finished yet', 'ยังไม่ครบ'));
  btn.disabled = true;
  try {
    const out = await busy(tr('Saving {name}…', 'กำลังบันทึก {name}…', { name: p?.name || ref }), () => call('setFinished', { ref, finished }));
    st.d.pitch[ref] = { ...st.d.pitch[ref], ...out.pitch };
    st.ctx.persist?.();
    forget('dashboard');
    toast(finished ? tr('Marked as finished', 'ทำเครื่องหมายว่าเสร็จสิ้นแล้ว') : tr('Marked as not finished', 'ยกเลิกเสร็จสิ้นแล้ว'));
  } catch (err) {
    await errorBox(err.message);
  }
  const sec = st.el.querySelector(`.pcard[data-ref="${CSS.escape(ref)}"]`);
  if (sec) $('[data-finish]', sec).innerHTML = finishRow(ref);
}

// Pre-pitch notes, read only (shown with "Show Pre-pitch data").
const yesNo = (v) => (v === 'Yes' ? tr('Yes', 'ใช่') : v === 'No' ? tr('No', 'ไม่') : '—');
const preToggle = (key) => `<button type="button" class="btn plain pre-toggle" data-pre-toggle="${esc(key)}" aria-expanded="${openPre.has(key)}">${icon('down', 14)}${openPre.has(key)
  ? tr('Hide Pre-pitch data', 'ซ่อนข้อมูลก่อนการนำเสนอ') : tr('Show Pre-pitch data', 'แสดงข้อมูลก่อนการนำเสนอ')}</button>`;
function prePeek(ref) {
  if (!openPre.has(ref)) return preToggle(ref);
  const x = st.d.pre?.[ref] || {};
  const row = (k, v) => `<dt>${esc(k)}</dt><dd>${v}</dd>`;
  return `${preToggle(ref)}<div class="fb pre-peek"><h4>${tr('Pre-pitch data', 'ข้อมูลก่อนการนำเสนอ')}</h4><dl class="kv">
    ${row(tr('Estimated Approved budget', 'งบประมาณที่คาดว่าจะได้รับอนุมัติ'), esc(moneyOr(x.approvedBudget)))}
    ${row(tr('Approved project .docx?', 'อนุมัติไฟล์โครงการ .docx?'), `${esc(yesNo(x.docxOk))} ${fileChip(x.docxFile)}`)}
    ${row(tr('Approved project PDF?', 'อนุมัติไฟล์โครงการ PDF?'), `${esc(yesNo(x.pdfOk))} ${fileChip(x.pdfFile)}`)}
    ${row(tr('Note', 'บันทึก'), x.note ? esc(x.note).replace(/\n/g, '<br>') : '—')}
    ${row(tr('Documents', 'เอกสาร'), (x.docs || []).map(fileChip).join(' ') || '—')}</dl></div>`;
}
function preForm() {
  if (!openPre.has('xl')) return preToggle('xl');
  const x = st.d.orgPre || {};
  return `${preToggle('xl')}<dl class="kv pre-peek"><dt>${tr('Feedback on the budget form', 'ข้อเสนอแนะต่อแบบฟอร์มงบประมาณ')}</dt><dd>${x.note ? esc(x.note).replace(/\n/g, '<br>') : '—'} ${fileChip(x.file)}</dd></dl>`;
}

const savedLine = (s) => (s?.pending ? `<span class="pill orange">${tr('Pending', 'รอส่ง')}</span> ${tr('Kept on this device; it is sent when the server can be reached', 'เก็บไว้ในเครื่องนี้ จะส่งเมื่อติดต่อเซิร์ฟเวอร์ได้')}`
  : s?.at ? tr('Saved {when} by {who}', 'บันทึก {when} โดย {who}', { when: dateTimeFmt(s.at), who: esc(st.ctx.nameOf(s.by)) }) : tr('Not saved yet', 'ยังไม่ได้บันทึก'));

// Noted documents (v2.15): a Word row and a PDF row, one file each. A .doc gets no Edit button (the editor opens .docx).
function noted(ref) {
  const list = st.form[ref].noted;
  const proj = st.d.projects.find((p) => p.ref === ref) || {}; // its submitted PDF / .docx: the editor's Reset All
  const fresh = new Set(newFiles(list, st.d.pitch[ref]?.noted).map((f) => f.id));
  const row = (k) => {
    const f = notedOf(list, k);
    const name = k === 'pdf' ? tr('Noted PDF', 'เอกสารบันทึก PDF') : tr('Noted .docx', 'เอกสารบันทึก .docx');
    return `<div class="noted-row"><span class="muted small noted-k">${name}</span><div class="upl">${fileChip(f, { unsaved: !!f && fresh.has(f.id) })}
      ${f && st.can ? editBtn(f, { org: orgRef(), slot: `noted:${ref}:${k === 'pdf' ? 'pdf' : 'docx'}`, plain: true, master: k === 'pdf' ? proj.pdf : proj.docx }) : ''}${!f && !st.can ? '<span class="muted">—</span>' : ''}
      ${st.can ? `<label class="btn">${icon('upload', 16)}${f ? tr('Replace', 'เปลี่ยนไฟล์') : k === 'pdf' ? tr('Upload PDF', 'อัปโหลด PDF') : tr('Upload .docx', 'อัปโหลด .docx')}<input type="file" hidden accept="${k === 'pdf' ? '.pdf' : '.docx,.doc'}" data-up="${k}"></label>` : ''}
      ${f && st.can ? `<button type="button" class="btn plain" data-rm="${k}">${tr('Remove', 'นำออก')}</button>` : ''}</div>${st.can ? capNote(k === 'pdf' ? 'PDF' : '.docx / .doc') : ''}</div>`;
  };
  return `${row('word')}${row('pdf')}`;
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
  const k = e.target.dataset.up;
  const file = e.target.files?.[0];
  if (!ref || !k || !file) return;
  const box = e.target.closest('[data-noted]');
  try {
    setNoted(ref, k, await busy(uploading(file.name), (set) => upload(file, { ref, purpose: 'noted', only: k }, set)));
    toast(UPLOADED());
  } catch (err) {
    await errorBox(`${file.name}: ${err.message}`, tr('Couldn’t upload', 'อัปโหลดไม่สำเร็จ'));
  }
  box.innerHTML = noted(ref);
}

async function onClick(e) {
  const pt = e.target.closest('[data-pre-toggle]');
  if (pt) {
    const key = pt.dataset.preToggle;
    if (openPre.has(key)) openPre.delete(key); else openPre.add(key);
    if (key === 'xl') $('[data-pre-form]', st.el).innerHTML = preForm();
    else $('[data-pre]', pt.closest('.pcard')).innerHTML = prePeek(key);
    return;
  }
  if (e.target.closest('[data-xl-clear]')) { st.xl = null; $('[data-xl-slot]', st.el).innerHTML = xlSlot(); return; }
  if (e.target.closest('[data-save-xl]')) { await run(e.target.closest('button'), saveXl); return; }
  const ref = refOf(e);
  if (!ref) return;
  const rm = e.target.closest('[data-rm]');
  if (rm) {
    setNoted(ref, rm.dataset.rm, null);
    rm.closest('[data-noted]').innerHTML = noted(ref);
    return;
  }
  const fin = e.target.closest('[data-finish-set]');
  if (fin) { await setFinished(ref, fin.dataset.finishSet === '1', fin); return; }
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
    out = { pitch: { ...st.form[ref], at: prev?.at || '', by: prev?.by || '', finished: prev?.finished || null, pending: true }, queued: true };
  }
  st.d.pitch[ref] = out.pitch;
  progress(st.el, st.d, 'pitching');
  st.ctx.persist?.();
  forget('dashboard');
  st.form[ref] = pick(out.pitch);
  const sec = st.el.querySelector(`.pcard[data-ref="${CSS.escape(ref)}"]`);
  if (sec) { $('[data-saved]', sec).innerHTML = savedLine(out.pitch); $('[data-noted]', sec).innerHTML = noted(ref); $('[data-finish]', sec).innerHTML = finishRow(ref); }
  marks();
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
  marks();
  return out;
}
