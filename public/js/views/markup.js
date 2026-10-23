// Edit a file in the site (v2.12), opened in its own tab from an organization page (editBtn in project.js):
// #/markup?org=<Org Ref>&id=<Drive file id>&kind=pdf|docx|xlsx&slot=<where the result goes>.
// PDF: the markup editor (markup/pdf-markup.js). Its marks are a draft saved on the server (Save draft, and with every
// Save to site) and kept on this device after each change, so a closed tab loses nothing.
// .docx / .xlsx (v2.13 prototype): edited in this page (markup/docx-edit.js, markup/xlsx-edit.js), no Google account
// needed; only the edited parts of the file change, and an Excel file keeps its locks. A draft (the edited file) is
// kept on this device (IndexedDB): with Save draft, and by itself when the tab is hidden (an iPad switching apps).
// "Open in Google Docs / Sheets instead" keeps the v2.12 way for bigger changes: a Google copy in the organization's
// folder, shared with the signed-in Gmail only; an Excel file gets its original's locks back (markup/xlsx-relock.js).
// Save to site uploads the result as that slot's feedback file and hands it to the organization tab (sendMarked),
// where Save keeps it, as with any upload. Download gives the same file.
import { tr } from '../i18n.js';
import { $, esc, icon, toast, busy, errorBox, confirmBox, dateTimeFmt } from '../ui.js';
import { call, upload } from '../api.js';
import { sendMarked } from '../project.js';

const MIME = {
  pdf: 'application/pdf',
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
};

let st = null; // { org, id, kind, slot, name, ed (PDF editor), office (Word/Excel editor), draft: { at, by } | null, copy (Google copy id), ctx }

export const isDirty = () => !!st?.ed?.isDirty() || !!st?.office?.isDirty();
export const unsaved = () => [];

// Where the slot's result is uploaded (as in the organization tabs).
function target(slot, org) {
  if (slot === 'form') return { org, purpose: 'form', only: 'any' };
  if (slot === 'xl') return { org, purpose: 'excel', only: 'excel' };
  let m = String(slot).match(/^pre:(.+):(pdf|docx)$/);
  if (m) return { ref: m[1], purpose: 'feedback', only: m[2] };
  m = String(slot).match(/^noted:(.+)$/);
  if (m) return { ref: m[1], purpose: 'noted' };
  throw new Error('Bad slot');
}

const bytesOf = async (b64) => (await fetch(`data:application/octet-stream;base64,${b64}`)).arrayBuffer();
const localKey = () => `markup:${st.ctx.me.email}:${st.id}`;
const readLocal = () => { try { return JSON.parse(localStorage.getItem(localKey()) || 'null'); } catch { return null; } };
const writeLocal = (v) => { try { if (v) localStorage.setItem(localKey(), JSON.stringify(v)); else localStorage.removeItem(localKey()); } catch { /* private mode */ } };
const outName = () => {
  const base = st.name.replace(/\.(pdf|docx|xlsx)$/i, '');
  return `${base} (${st.kind === 'pdf' ? 'marked' : 'edited'}).${st.kind}`;
};

export async function render(el, ctx, route) {
  const q = route.q;
  st = { org: q.get('org') || '', id: q.get('id') || '', kind: q.get('kind') || '', slot: q.get('slot') || '', name: '', ed: null, draft: null, copy: '', ctx };
  if (!ctx.canEdit) throw new Error(tr('Only staff can edit files.', 'เฉพาะเจ้าหน้าที่เท่านั้นที่แก้ไขไฟล์ได้'));
  target(st.slot, st.org); // a bad link stops here
  el.classList.add('markup-page');
  el.innerHTML = `<div class="mk-head group"><span class="grow"><b data-name>${esc(tr('Opening…', 'กำลังเปิด…'))}</b>
      <small class="muted" data-status></small></span>
    <span class="actions" data-actions></span></div><div class="mk-root" data-root><p class="muted">${esc(tr('Loading the file…', 'กำลังโหลดไฟล์…'))}</p></div>`;
  if (st.kind === 'pdf') await openPdf(el);
  else await openOffice(el);
}

const status = (el, html) => { $('[data-status]', el).innerHTML = html; };
const draftLine = () => (st.draft?.at ? tr('Draft saved {when} by {who}', 'บันทึกฉบับร่าง {when} โดย {who}', { when: dateTimeFmt(st.draft.at), who: esc(st.ctx.nameOf(st.draft.by)) }) : tr('No draft yet', 'ยังไม่มีฉบับร่าง'));

async function openPdf(el) {
  const [{ openPdfEditor }, out] = await Promise.all([import('../markup/pdf-markup.js'), call('markupOpen', { org: st.org, id: st.id, bytes: true })]);
  st.name = out.name;
  st.draft = out.draft ? { at: out.draft.at, by: out.draft.by } : null;
  $('[data-name]', el).textContent = st.name;
  // Marks on this device win when they were made on top of the draft the server still has.
  const local = readLocal();
  const useLocal = local && (local.base || '') === (st.draft?.at || '');
  st.ed = await openPdfEditor($('[data-root]', el), {
    bytes: await bytesOf(out.data), name: st.name, draft: useLocal ? local.marks : out.draft?.marks || null, t: tr,
    onDraft: (marks) => writeLocal({ marks, base: st.draft?.at || '' }), textSize: out.textSize,
  });
  if (local && !useLocal) writeLocal(null); // made on an older draft: the newer draft on the server is shown
  status(el, useLocal ? tr('Unsaved marks from this device were restored', 'กู้คืนเครื่องหมายที่ยังไม่บันทึกจากเครื่องนี้แล้ว') : draftLine());
  $('[data-actions]', el).innerHTML = `<button class="btn" data-a="draft">${tr('Save draft', 'บันทึกฉบับร่าง')}</button>
    <button class="btn" data-a="dl">${icon('download', 16)}${tr('Download', 'ดาวน์โหลด')}</button>
    <button class="btn primary" data-a="site">${icon('upload', 16)}${tr('Save to site', 'บันทึกเข้าเว็บไซต์')}</button>`;
  $('[data-actions]', el).onclick = (e) => act(el, e.target.closest('[data-a]')?.dataset.a);
}

async function saveDraft() {
  const marks = st.ed.marks();
  const out = await call('saveMarkupDraft', { org: st.org, id: st.id, marks, base: st.draft?.at || '' });
  st.draft = out.draft;
  writeLocal(null);
  st.ed.clean();
}

// ---------- Word / Excel in this page (v2.13 prototype) ----------
// Device drafts: the edited file's bytes in IndexedDB (a .docx or .xlsx is too big for localStorage), one per user
// and file. base = the original's size when the draft was made, so a draft made on an older original is flagged.
const DB = 'tusu-markup';
function drafts(mode, fn) {
  return new Promise((res, rej) => {
    const open = indexedDB.open(DB, 1);
    open.onupgradeneeded = () => open.result.createObjectStore('office');
    open.onerror = () => rej(open.error);
    open.onsuccess = () => {
      const tx = open.result.transaction('office', mode); const q = fn(tx.objectStore('office'));
      tx.oncomplete = () => { open.result.close(); res(q.result); };
      tx.onerror = () => { open.result.close(); rej(tx.error); };
    };
  });
}
const readDraft = () => drafts('readonly', (o) => o.get(localKey())).catch(() => null);
const writeDraft = (v) => drafts('readwrite', (o) => (v ? o.put(v, localKey()) : o.delete(localKey())));

async function openOffice(el, { fresh = false } = {}) {
  const [mod, out, draft] = await Promise.all([
    st.kind === 'xlsx' ? import('../markup/xlsx-edit.js') : import('../markup/docx-edit.js'),
    call('markupOpen', { org: st.org, id: st.id, bytes: true }),
    fresh ? null : readDraft(),
  ]);
  st.name = out.name; st.size = out.size; st.copy = '';
  $('[data-name]', el).textContent = st.name;
  st.local = draft?.bytes ? draft : null;
  const root = $('[data-root]', el);
  root.classList.add('mk-office'); root.innerHTML = '';
  const open = st.kind === 'xlsx' ? mod.openXlsxEditor : mod.openDocxEditor;
  try {
    st.office = await open(root, st.local ? st.local.bytes.slice(0) : await bytesOf(out.data), { t: tr });
  } catch (err) {
    if (!st.local) throw err;
    st.local = null; // a draft this editor can't read: start again from the file
    st.office = await open(root, await bytesOf(out.data), { t: tr });
  }
  status(el, officeLine());
  const app = st.kind === 'xlsx' ? 'Google Sheets' : 'Google Docs';
  $('[data-actions]', el).innerHTML = `<button class="btn plain" data-a="google" title="${esc(tr('For bigger changes (inserting rows, pictures, ...): edit a Google copy instead', 'สำหรับการแก้ไขใหญ่ (แทรกแถว รูปภาพ ฯลฯ): แก้ไขในสำเนา Google แทน'))}">${icon('ext', 16)}${tr('Open in {app} instead', 'เปิดใน {app} แทน', { app })}</button>
    ${st.local ? `<button class="btn plain" data-a="fresh">${tr('Discard draft', 'ทิ้งฉบับร่าง')}</button>` : ''}
    <button class="btn" data-a="draft">${tr('Save draft', 'บันทึกฉบับร่าง')}</button>
    <button class="btn" data-a="dl">${icon('download', 16)}${tr('Download', 'ดาวน์โหลด')}</button>
    <button class="btn primary" data-a="site">${icon('upload', 16)}${tr('Save to site', 'บันทึกเข้าเว็บไซต์')}</button>`;
  $('[data-actions]', el).onclick = (e) => act(el, e.target.closest('[data-a]')?.dataset.a);
}

// "Draft from this device, saved <when>" (flagged when the original changed since), or how edits are kept.
const officeLine = () => {
  if (!st.local) return tr('Editing in this page · edits are kept on this device when you press Save draft', 'แก้ไขในหน้านี้ · กดบันทึกฉบับร่างเพื่อเก็บการแก้ไขไว้ในเครื่องนี้');
  const line = tr('Draft from this device, saved {when}', 'ฉบับร่างจากเครื่องนี้ บันทึกเมื่อ {when}', { when: dateTimeFmt(st.local.at) });
  return st.local.base !== st.size ? `${line} · <b>${tr('the original file has changed since', 'ไฟล์ต้นฉบับมีการเปลี่ยนแปลงหลังจากนั้น')}</b>` : line;
};

async function saveOfficeDraft() {
  const bytes = await st.office.save();
  st.local = { bytes, at: new Date().toISOString(), base: st.local?.base ?? st.size };
  await writeDraft(st.local);
}

// The tab hidden (another app on an iPad, which may close the tab): keep the edits as the device draft.
document.addEventListener('visibilitychange', () => {
  if (document.hidden && st?.office?.isDirty()) saveOfficeDraft().catch(() => { /* private mode: nothing kept */ });
});

async function toGoogle(el) {
  if (st.office?.isDirty() && !(await confirmBox(tr('Open a Google copy instead?', 'เปิดสำเนา Google แทนหรือไม่'),
    tr('The changes made in this page are not in the Google copy. Save a draft first to keep them on this device.', 'การแก้ไขในหน้านี้จะไม่อยู่ในสำเนา Google กรุณาบันทึกฉบับร่างก่อนเพื่อเก็บไว้ในเครื่องนี้'),
    { ok: tr('Open Google copy', 'เปิดสำเนา Google') }))) return;
  st.office = null;
  $('[data-root]', el).classList.remove('mk-office');
  await openGoogle(el);
}

async function openGoogle(el) {
  const out = await call('markupCopy', { org: st.org, id: st.id });
  st.name = out.name; st.copy = out.copy;
  const app = st.kind === 'xlsx' ? 'Google Sheets' : 'Google Docs';
  $('[data-name]', el).textContent = st.name;
  st.edited = out.edited;
  status(el, editLine());
  // Back from Google: the last edit again (someone may have edited since).
  document.addEventListener('visibilitychange', async () => {
    if (document.hidden || !el.isConnected || st.office) return; // back in this page's editor: nothing to check
    try { st.edited = (await call('markupCopy', { org: st.org, id: st.id, check: true })).edited; status(el, editLine()); } catch { /* keep the last one */ }
  });
  $('[data-root]', el).innerHTML = `<section class="group mk-steps"><ol>
      <li>${tr('Open the copy in {app} and make your changes. Google saves them as you type.', 'เปิดสำเนาใน {app} แล้วแก้ไข Google จะบันทึกให้อัตโนมัติ', { app })}
        <p><a class="btn primary" href="${esc(out.url)}" target="_blank" rel="noopener">${icon('ext', 16)}${tr('Open in {app}', 'เปิดใน {app}', { app })}</a></p></li>
      <li>${tr('Come back to this tab and press Save to site. The file goes to the organization page, where you press Save.', 'กลับมาที่แท็บนี้แล้วกดบันทึกเข้าเว็บไซต์ ไฟล์จะไปอยู่ในหน้าองค์กร แล้วกดบันทึกที่นั่น')}</li>
    </ol>
    ${st.kind === 'xlsx' ? `<p class="muted small mk-lock">${icon('lock', 14)}${tr('The form’s locks (protected sheets, locked cells, password) are put back from the original when you save.', 'การล็อกของแบบฟอร์ม (ชีตที่ป้องกัน เซลล์ที่ล็อก รหัสผ่าน) จะถูกใส่กลับจากไฟล์ต้นฉบับเมื่อบันทึก')}</p>` : ''}
    <p class="muted small">${tr('Everyone who opens this file here edits the same copy.', 'ทุกคนที่เปิดไฟล์นี้จากที่นี่จะแก้ไขสำเนาเดียวกัน')}</p>
    <p><button class="btn plain" data-back>${tr('Edit in this page instead', 'แก้ไขในหน้านี้แทน')}</button></p></section>`;
  $('[data-back]', el).onclick = () => openOffice(el).catch((err) => errorBox(err.message, tr('Couldn’t open', 'เปิดไม่สำเร็จ')));
  $('[data-actions]', el).innerHTML = `<button class="btn" data-a="dl">${icon('download', 16)}${tr('Download', 'ดาวน์โหลด')}</button>
    <button class="btn primary" data-a="site">${icon('upload', 16)}${tr('Save to site', 'บันทึกเข้าเว็บไซต์')}</button>`;
  $('[data-actions]', el).onclick = (e) => act(el, e.target.closest('[data-a]')?.dataset.a);
}

// Google copy: "Last edited <when> by <who>" (Google's own record), or "Not edited yet" while it equals the original.
const editLine = () => (st.edited?.at ? tr('Google copy last edited {when} by {who} · shared only with you', 'สำเนา Google แก้ไขล่าสุด {when} โดย {who} · แชร์ให้คุณคนเดียว',
  { when: dateTimeFmt(st.edited.at), who: esc(st.ctx.nameOf(st.edited.by)) }) : tr('A Google copy for editing, shared only with you', 'สำเนา Google สำหรับแก้ไข แชร์ให้คุณคนเดียว'));

// The finished file as a File: the marked PDF, or the Google copy exported (an Excel file re-locked).
async function produce(set) {
  if (st.kind === 'pdf') { set(0.3); return new Blob([await st.ed.exportPdf()], { type: MIME.pdf }); }
  if (st.office) { set(0.3); return new Blob([await st.office.save()], { type: MIME[st.kind] }); }
  const out = await call('markupExport', { org: st.org, copy: st.copy });
  set(0.4);
  let bytes = await bytesOf(out.data);
  if (st.kind === 'xlsx') {
    const [{ relock }, orig] = await Promise.all([import('../markup/xlsx-relock.js'), call('markupOpen', { org: st.org, id: out.src, bytes: true })]);
    set(0.7);
    bytes = await relock(await bytesOf(orig.data), bytes);
  }
  return new Blob([bytes], { type: MIME[st.kind] });
}

async function act(el, a) {
  if (!a) return;
  const btns = [...el.querySelectorAll('[data-a]')];
  btns.forEach((b) => { b.disabled = true; });
  try {
    if (a === 'google') await toGoogle(el);
    if (a === 'fresh' && await confirmBox(tr('Discard the draft?', 'ทิ้งฉบับร่างหรือไม่'), tr('The file opens again as it is on the site. The draft on this device is deleted.', 'ไฟล์จะเปิดใหม่ตามที่อยู่ในเว็บไซต์ และฉบับร่างในเครื่องนี้จะถูกลบ'), { ok: tr('Discard', 'ทิ้ง'), danger: true })) {
      await writeDraft(null).catch(() => {});
      await openOffice(el, { fresh: true });
    }
    if (a === 'draft' && st.office) {
      await busy(tr('Saving the draft…', 'กำลังบันทึกฉบับร่าง…'), saveOfficeDraft);
      status(el, officeLine());
      toast(tr('Draft saved on this device', 'บันทึกฉบับร่างในเครื่องนี้แล้ว'));
    } else if (a === 'draft') {
      await busy(tr('Saving the draft…', 'กำลังบันทึกฉบับร่าง…'), saveDraft);
      status(el, draftLine());
      toast(tr('Draft saved', 'บันทึกฉบับร่างแล้ว'));
    }
    if (a === 'dl') {
      const blob = await busy(tr('Preparing {file}…', 'กำลังเตรียม {file}…', { file: outName() }), produce);
      const link = document.createElement('a');
      link.href = URL.createObjectURL(blob); link.download = outName();
      document.body.append(link); link.click(); link.remove();
      setTimeout(() => URL.revokeObjectURL(link.href), 5000);
    }
    if (a === 'site') {
      const file = await busy(tr('Saving {file} to the site…', 'กำลังบันทึก {file} เข้าเว็บไซต์…', { file: outName() }), async (set) => {
        const f = new File([await produce((x) => set(x * 0.5))], outName(), { type: MIME[st.kind] });
        const saved = await upload(f, target(st.slot, st.org), (x) => set(0.5 + x * 0.5));
        if (st.ed) await saveDraft(); // the marks of the saved file stay editable
        if (st.office) { await writeDraft(null).catch(() => {}); st.local = null; } // saved to the site: no draft needed
        return saved;
      });
      if (st.ed) status(el, draftLine());
      if (st.office) status(el, tr('Saved to the site · you can keep editing', 'บันทึกเข้าเว็บไซต์แล้ว · แก้ไขต่อได้'));
      await hand(el, file);
    }
  } catch (err) {
    await errorBox(err.message, a === 'draft' ? undefined : tr('Couldn’t finish', 'ทำไม่สำเร็จ'));
  }
  btns.forEach((b) => { b.disabled = false; });
}

// The saved file to the organization tab; if no tab took it, say so and offer to send it again.
async function hand(el, file) {
  const msg = { org: st.org, slot: st.slot, file };
  if (await sendMarked(msg)) {
    toast(tr('Sent to the organization page — press Save there to keep it', 'ส่งไปที่หน้าองค์กรแล้ว — กดบันทึกที่หน้านั้นเพื่อเก็บ'));
    return;
  }
  const box = document.createElement('div');
  box.className = 'notice orange mk-orphan';
  box.innerHTML = `${icon('alert', 18)}<span class="grow">${tr('{file} is saved in Drive, but the organization page that opened this tab is closed or on another tab. Open its {tab} tab, then press Send again.', '{file} บันทึกใน Drive แล้ว แต่หน้าองค์กรที่เปิดแท็บนี้ปิดอยู่หรืออยู่ที่แท็บอื่น กรุณาเปิดแท็บ {tab} ขององค์กร แล้วกดส่งอีกครั้ง',
    { file: esc(file.name), tab: /^(xl|noted:)/.test(st.slot) ? tr('Pitching', 'การนำเสนอ') : tr('Pre-pitch', 'ก่อนการนำเสนอ') })}</span>
    <a class="btn" href="#/org/${encodeURIComponent(st.org)}/${/^(xl|noted:)/.test(st.slot) ? 'pitching' : 'prepitch'}" target="_blank" rel="opener">${tr('Open organization', 'เปิดหน้าองค์กร')}</a>
    <button class="btn primary">${tr('Send again', 'ส่งอีกครั้ง')}</button>`;
  el.querySelector('.mk-orphan')?.remove();
  $('.mk-head', el).after(box);
  box.querySelector('button').onclick = async () => {
    if (await sendMarked(msg)) { box.remove(); toast(tr('Sent to the organization page — press Save there to keep it', 'ส่งไปที่หน้าองค์กรแล้ว — กดบันทึกที่หน้านั้นเพื่อเก็บ')); }
    else toast(tr('The organization page still isn’t open', 'ยังไม่ได้เปิดหน้าองค์กร'), 'bad');
  };
}
