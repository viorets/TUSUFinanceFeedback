// Edit a file in the site (v2.12), opened from an organization page (editBtn in project.js; same tab since v2.16.3):
// #/markup?org=<Org Ref>&id=<Drive file id>&kind=pdf|docx|xlsx&slot=<where the result goes>.
// PDF: the markup editor (markup/pdf-markup.js). Its marks are a draft saved on the server with every Save, and kept on
// this device after each change, so a closed tab loses nothing.
// .docx / .xlsx (v2.13 prototype): edited in this page (markup/docx-edit.js, markup/xlsx-edit.js), no Google account
// needed; only the edited parts of the file change, and an Excel file keeps its locks. A draft (the edited file) is
// kept on this device (IndexedDB) by itself when the tab is hidden (an iPad switching apps, a closed tab).
// "Open in Google Docs / Sheets instead" keeps the v2.12 way for bigger changes: a Google copy in the organization's
// folder, shared with the signed-in Gmail only; an Excel file gets its original's locks back (markup/xlsx-relock.js).
// Save (v2.16) uploads the result as that slot's feedback file and saves it into its card on the server (saveMarked in
// Code.gs; no second Save on the organization page), then tells an open organization tab (sendMarked). Save and quit
// does the same, then goes back to the organization page. Download gives the same file. Before it replaces the slot's
// file, Save checks that nobody saved a newer one since this page opened.
// Reset All (v2.16.2): the editor opens the slot's master file (the submitted PDF / .docx / budget form, from the
// link's master) again with no marks or edits; Save then replaces the slot's file with it.
// Who is editing (v2.13): one person at a time edits a .docx / .xlsx here (presence in Code.gs). Anyone else who
// opens it waits while the editor is asked to share: Yes saves the editor's work to the site and moves everyone to
// one Google copy made from that saved file; No, or no answer within a minute, leaves them read-only.
import { tr } from '../i18n.js';
import { $, esc, icon, toast, busy, errorBox, confirmBox, dialog, dateTimeFmt } from '../ui.js';
import { call, upload, previewUrl, viewUrl, downloadUrl, forget, peek, remember } from '../api.js';
import { sendMarked, openPreview } from '../project.js';
import { ORGS } from '../masterdata.js';
import { idToken } from '../firebase.js';
import { SCRIPT_URL } from '../config.js';

const MIME = {
  pdf: 'application/pdf',
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
};

let st = null; // { org, id, kind, slot, name, ed (PDF editor), office (Word/Excel editor), draft: { at, by } | null, copy (Google copy id), ctx }

// Unsaved = not on the site yet. A draft (on this device, or the PDF draft on the server) still counts: only Save to
// site clears st.unsent. Refreshing or closing the tab then asks first (app.js); the site's leave window names the file.
export const isDirty = () => !!st && (!!st.unsent || !!st.ed?.isDirty() || !!st.office?.isDirty());
export const unsaved = () => (isDirty() ? [{ ref: slotCode(), name: st.name, files: [] }] : []);
const slotCode = () => String(st.slot).match(/^(?:pre|noted):([^:]+)/)?.[1] || st.org;
// Card slots (Pitching tab) are saved into their card by Save; Pre-pitch slots (links from before v2.13) are handed over.
const inCard = () => /^(xl|noted:)/.test(st.slot);
const orgHref = () => `#/org/${encodeURIComponent(st.org)}/${inCard() ? 'pitching' : 'prepitch'}`;

// Where the slot's result is uploaded (as in the organization tabs).
function target(slot, org) {
  if (slot === 'form') return { org, purpose: 'form', only: 'any' };
  if (slot === 'xl') return { org, purpose: 'excel', only: 'excel' };
  let m = String(slot).match(/^pre:(.+):(pdf|docx)$/);
  if (m) return { ref: m[1], purpose: 'feedback', only: m[2] };
  // Noted documents (v2.15): the Word slot or the PDF slot; 'noted:<ref>' = a link from before v2.15 (adds a file).
  m = String(slot).match(/^noted:([^:]+)(?::(pdf|docx))?$/);
  if (m) return { ref: m[1], purpose: 'noted', ...(m[2] ? { only: m[2] === 'pdf' ? 'pdf' : 'word' } : {}) };
  throw new Error('Bad slot');
}

const bytesOf = async (b64) => (await fetch(`data:application/octet-stream;base64,${b64}`)).arrayBuffer();
const localKey = () => `markup:${st.ctx.me.email}:${st.id}`;
const readLocal = () => { try { return JSON.parse(localStorage.getItem(localKey()) || 'null'); } catch { return null; } };
const writeLocal = (v) => { try { if (v) localStorage.setItem(localKey(), JSON.stringify(v)); else localStorage.removeItem(localKey()); } catch { /* private mode */ } };
// The saved file's name: "<name> (Feedback)" in the Pitching tab (Correct Excel file, Noted documents), else
// "(marked)" / "(edited)". A file saved here before is not tagged twice.
const outName = () => {
  const base = st.name.replace(/\.(pdf|docx|xlsx)$/i, '').replace(/ \((marked|edited|Feedback)\)$/i, '');
  const tag = /^(xl|noted:)/.test(st.slot) ? 'Feedback' : st.kind === 'pdf' ? 'marked' : 'edited';
  return `${base} (${tag}).${st.kind}`;
};

export async function render(el, ctx, route) {
  const q = route.q;
  st = { org: q.get('org') || '', id: q.get('id') || '', kind: q.get('kind') || '', slot: q.get('slot') || '', master: q.get('master') || '', name: '', ed: null, draft: null, copy: '', ctx };
  if (!ctx.canEdit) throw new Error(tr('Only staff can edit files.', 'เฉพาะเจ้าหน้าที่เท่านั้นที่แก้ไขไฟล์ได้'));
  target(st.slot, st.org); // a bad link stops here
  el.classList.add('markup-page');
  const orgName = ORGS.find((o) => o.code === st.org)?.en || st.org;
  el.innerHTML = `<a class="btn plain back" href="${esc(orgHref())}">${icon('back', 18)}${esc(tr('Back to {org}', 'กลับไปที่ {org}', { org: orgName }))}</a>
    <div class="mk-head group"><span class="grow"><b data-name>${esc(tr('Opening…', 'กำลังเปิด…'))}</b>
      <small class="muted" data-status></small></span>
    <span class="actions" data-actions></span></div><div class="mk-root" data-root><p class="muted">${esc(tr('Loading the file…', 'กำลังโหลดไฟล์…'))}</p></div>
    <p class="mk-foot muted small" data-foot></p>`;
  if (st.kind === 'pdf') await openPdf(el);
  else await openOffice(el);
}

const status = (el, html) => { $('[data-status]', el).innerHTML = html; foot(el); };
// Under the editor (v2.13): the last save of this file on the site (the slot's file: when, by whom), and the file
// this page saved there, if any.
function foot(el) {
  const f = $('[data-foot]', el); if (!f) return;
  const parts = [];
  if (st.saved) parts.push(inCard() ? tr('You saved {file} at {when}.', 'คุณบันทึก {file} เมื่อ {when}', { file: esc(st.saved.name), when: dateTimeFmt(st.saved.at) })
    : tr('You saved {file} to the site at {when}. Press Save on the organization page to keep it.', 'คุณบันทึก {file} เข้าเว็บไซต์เมื่อ {when} กดบันทึกที่หน้าองค์กรเพื่อเก็บไฟล์', { file: esc(st.saved.name), when: dateTimeFmt(st.saved.at) }));
  if (st.site?.at) parts.push(tr('Last saved on the site {when} by {who}', 'บันทึกล่าสุดบนเว็บไซต์ {when} โดย {who}', { when: dateTimeFmt(st.site.at), who: esc(st.ctx.nameOf(st.site.by) || '—') }));
  else if (st.site) parts.push(tr('Not saved on the site yet', 'ยังไม่เคยบันทึกบนเว็บไซต์'));
  f.innerHTML = parts.join(' · ');
}
const draftLine = () => {
  // The marks are kept on the server with every Save (and a hand-over), so they open again for more editing.
  const line = st.draft?.at ? tr('Marks saved {when} by {who}', 'บันทึกเครื่องหมาย {when} โดย {who}', { when: dateTimeFmt(st.draft.at), who: esc(st.ctx.nameOf(st.draft.by)) }) : tr('No marks saved yet', 'ยังไม่มีเครื่องหมายที่บันทึก');
  return st.unsent ? `${line} · <b>${tr('not saved yet', 'ยังไม่ได้บันทึก')}</b>` : line;
};

// "You have unsaved changes from earlier": a notice under the header with Restore and Discard (the iPad fallback:
// Safari there doesn't ask before a tab is closed, so the work kept on the device is offered when the file opens).
function earlier(el, at, extra, onRestore, onDiscard) {
  el.querySelector('.mk-earlier')?.remove();
  const box = document.createElement('div');
  box.className = 'notice orange mk-earlier';
  box.innerHTML = `${icon('alert', 18)}<span class="grow">${esc(at ? tr('You have unsaved changes from earlier (kept on this device {when}).', 'คุณมีการแก้ไขที่ยังไม่บันทึกจากก่อนหน้านี้ (เก็บไว้ในเครื่องนี้เมื่อ {when})', { when: dateTimeFmt(at) })
    : tr('You have unsaved changes from earlier, kept on this device.', 'คุณมีการแก้ไขที่ยังไม่บันทึกจากก่อนหน้านี้ เก็บไว้ในเครื่องนี้'))}${extra ? ` <b>${esc(extra)}</b>` : ''}</span>
    <button class="btn" data-discard>${esc(tr('Discard', 'ทิ้ง'))}</button><button class="btn primary" data-restore>${esc(tr('Restore', 'กู้คืน'))}</button>`;
  $('.mk-head', el).after(box);
  box.querySelector('[data-restore]').onclick = async () => {
    box.remove();
    try { await onRestore(); } catch (err) { errorBox(err.message, tr('Couldn’t restore', 'กู้คืนไม่สำเร็จ')); }
  };
  box.querySelector('[data-discard]').onclick = async () => {
    if (!(await confirmBox(tr('Discard the changes from earlier?', 'ทิ้งการแก้ไขจากก่อนหน้านี้หรือไม่'), tr('They are deleted from this device. The file stays as it is on the site.', 'การแก้ไขจะถูกลบจากเครื่องนี้ ไฟล์ยังคงเหมือนที่อยู่ในเว็บไซต์'), { ok: tr('Discard', 'ทิ้ง'), danger: true }))) return;
    box.remove();
    await onDiscard();
  };
}

// PDF (v2.13): who is marking it up is checked as for Word / Excel (claim runs alongside the download); someone
// else marking it up = the waiting screen, and Yes hands the PDF over with the marks saved as the draft.
async function openPdf(el, { restore = false, blank = false } = {}) {
  stopPresence();
  const [{ openPdfEditor }, out, claim] = await Promise.all([import('../markup/pdf-markup.js'), st.pdfOut || call('markupOpen', { org: st.org, id: st.id, bytes: true }, { wait: 90e3 }),
    presence('claim').catch(() => ({ state: 'none' }))]);
  st.pdfOut = out;
  st.name = out.name; if (claim.site) st.site = claim.site;
  st.draft = out.draft ? { at: out.draft.at, by: out.draft.by } : null;
  $('[data-name]', el).textContent = st.name;
  if (claim.state === 'busy') return showWaiting(el, claim.by);
  el.querySelector('.mk-earlier')?.remove();
  // Marks kept on this device (made on top of the draft the server still has) are offered: Restore or Discard.
  const local = readLocal();
  const usable = local && (local.base || '') === (st.draft?.at || '');
  if (local && !usable) writeLocal(null); // made on an older draft: the newer draft on the server is shown
  st.ed = await openPdfEditor($('[data-root]', el), {
    bytes: await bytesOf(out.data), name: st.name, draft: blank ? null : restore && usable ? local.marks : out.draft?.marks || null, t: tr,
    onDraft: (marks) => { st.unsent = true; writeLocal({ marks, base: st.draft?.at || '', at: new Date().toISOString() }); }, defaults: out.defaults,
  });
  if (restore && usable) st.unsent = true;
  if (claim.state === 'mine') startBeat(el);
  status(el, draftLine());
  if (usable && !restore && !blank) {
    earlier(el, local.at, '', () => openPdf(el, { restore: true }), () => writeLocal(null));
  }
  $('[data-actions]', el).innerHTML = `<button class="btn" data-a="dl">${icon('download', 16)}${tr('Download', 'ดาวน์โหลด')}</button>
    ${allBtn()}<button class="btn" data-a="site">${icon('upload', 16)}${tr('Save', 'บันทึก')}</button>
    <button class="btn primary" data-a="quit">${tr('Save and quit', 'บันทึกและออก')}</button>`;
  $('[data-actions]', el).onclick = (e) => act(el, e.target.closest('[data-a]')?.dataset.a);
}

// The organization page this tab goes back to: its kept copy gets the saved card at once (no flash of the old one),
// and the next read asks the server.
function keepCard({ pitch, orgPitch }) {
  forget('org'); forget('dashboard');
  const body = { ref: st.org }; const d = peek('org', body); if (!d) return;
  const ref = String(st.slot).match(/^noted:([^:]+)/)?.[1];
  if (pitch && ref && d.pitch) d.pitch[ref] = pitch;
  if (orgPitch) d.orgPitch = orgPitch;
  remember('org', body, d);
}
const allBtn = () => (st.master ? `<button class="btn plain" data-a="all" title="${esc(tr('Open the submitted file again, with no marks or edits', 'เปิดไฟล์ที่ส่งมาใหม่อีกครั้ง โดยไม่มีเครื่องหมายหรือการแก้ไข'))}">${tr('Reset All', 'เริ่มใหม่ทั้งหมด')}</button>` : '');
async function resetAll(el) {
  if (!(await confirmBox(tr('Start again from the submitted file?', 'เริ่มใหม่จากไฟล์ที่ส่งมาหรือไม่'),
    tr('The editor opens the submitted file with no marks or edits. The file on the site stays as it is until you press Save.', 'ตัวแก้ไขจะเปิดไฟล์ที่ส่งมาโดยไม่มีเครื่องหมายหรือการแก้ไข ไฟล์บนเว็บไซต์ยังคงเดิมจนกว่าคุณจะกดบันทึก'),
    { ok: tr('Reset All', 'เริ่มใหม่ทั้งหมด'), danger: true }))) return;
  if (st.hold) await presence('release').catch(() => {}); // let go of the file being edited
  stopPresence();
  st.id = st.master; st.pdfOut = null; st.office = null; st.ed = null; st.copy = ''; st.unsent = false;
  const q = new URLSearchParams(location.hash.split('?')[1] || ''); q.set('id', st.id);
  history.replaceState(null, '', `#/markup?${q}`); // a refresh opens the master too
  await busy(tr('Opening the submitted file…', 'กำลังเปิดไฟล์ที่ส่งมา…'), () => (isPdf() ? openPdf(el, { blank: true }) : openOffice(el, { fresh: true, page: true })));
  toast(tr('Started again from the submitted file — press Save to replace the file on the site', 'เริ่มใหม่จากไฟล์ที่ส่งมาแล้ว — กดบันทึกเพื่อแทนที่ไฟล์บนเว็บไซต์'));
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

// claim runs alongside the file download, so asking who is editing adds no time. A server without presence (or a
// network error) leaves the page editing as before, without the check.
async function openOffice(el, { fresh = false, page = false, restore = false } = {}) {
  stopPresence();
  const [mod, out, draft, claim] = await Promise.all([
    st.kind === 'xlsx' ? import('../markup/xlsx-edit.js') : import('../markup/docx-edit.js'),
    call('markupOpen', { org: st.org, id: st.id, bytes: true }, { wait: 90e3 }),
    fresh ? null : readDraft(),
    presence('claim', page ? { page: true } : {}).catch(() => ({ state: 'none' })),
  ]);
  st.name = out.name; st.size = out.size; st.copy = '';
  $('[data-name]', el).textContent = st.name;
  if (claim.site) st.site = claim.site;
  if (claim.state === 'busy') return showWaiting(el, claim.by);
  if (claim.state === 'google') return joinGoogle(el, claim.src, claim.by);
  el.querySelector('.mk-earlier')?.remove();
  const kept = draft?.bytes ? draft : null;
  st.local = restore ? kept : null;
  if (restore && kept) st.unsent = true;
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
  if (claim.state === 'mine') startBeat(el);
  status(el, officeLine());
  if (kept && !restore) {
    earlier(el, kept.at, kept.base !== st.size ? tr('The original file has changed since.', 'ไฟล์ต้นฉบับมีการเปลี่ยนแปลงหลังจากนั้น') : '',
      () => openOffice(el, { restore: true }), () => writeDraft(null).catch(() => {}));
  }
  const app = st.kind === 'xlsx' ? 'Google Sheets' : 'Google Docs';
  $('[data-actions]', el).innerHTML = `<button class="btn plain" data-a="google" title="${esc(tr('For bigger changes (inserting rows, pictures, ...): edit a Google copy instead', 'สำหรับการแก้ไขใหญ่ (แทรกแถว รูปภาพ ฯลฯ): แก้ไขในสำเนา Google แทน'))}">${icon('ext', 16)}${tr('Open in {app} instead', 'เปิดใน {app} แทน', { app })}</button>
    ${st.local ? `<button class="btn plain" data-a="fresh">${tr('Discard draft', 'ทิ้งฉบับร่าง')}</button>` : ''}
    <button class="btn" data-a="dl">${icon('download', 16)}${tr('Download', 'ดาวน์โหลด')}</button>
    ${allBtn()}<button class="btn" data-a="site">${icon('upload', 16)}${tr('Save', 'บันทึก')}</button>
    <button class="btn primary" data-a="quit">${tr('Save and quit', 'บันทึกและออก')}</button>`;
  $('[data-actions]', el).onclick = (e) => act(el, e.target.closest('[data-a]')?.dataset.a);
}

// "Draft from this device, saved <when>" (flagged when the original changed since), or how edits are kept.
const officeLine = () => {
  if (!st.local) {
    return st.hold ? tr('You are editing this file · anyone else who opens it asks you first', 'คุณกำลังแก้ไขไฟล์นี้ · ผู้อื่นที่เปิดไฟล์จะต้องขอคุณก่อน')
      : tr('Editing in this page · unsaved edits are kept on this device if the tab closes', 'แก้ไขในหน้านี้ · การแก้ไขที่ยังไม่บันทึกจะเก็บไว้ในเครื่องนี้หากปิดแท็บ');
  }
  const line = tr('Draft saved on this device {when}, not yet on the site', 'บันทึกฉบับร่างในเครื่องนี้เมื่อ {when} ยังไม่ได้บันทึกเข้าเว็บไซต์', { when: dateTimeFmt(st.local.at) });
  return st.local.base !== st.size ? `${line} · <b>${tr('the original file has changed since', 'ไฟล์ต้นฉบับมีการเปลี่ยนแปลงหลังจากนั้น')}</b>` : line;
};

async function saveOfficeDraft() {
  const bytes = await st.office.save(); // the editor counts itself clean now; st.unsent keeps the page unsaved
  st.unsent = true;
  st.local = { bytes, at: new Date().toISOString(), base: st.local?.base ?? st.size };
  await writeDraft(st.local);
}

// The tab hidden (another app on an iPad, which may close the tab): keep the edits as the device draft.
const keepOnDevice = () => { if (st?.office?.isDirty()) saveOfficeDraft().catch(() => { /* private mode: nothing kept */ }); };
document.addEventListener('visibilitychange', () => { if (document.hidden) keepOnDevice(); });
addEventListener('pagehide', keepOnDevice);

async function toGoogle(el) {
  if (st.office?.isDirty() && !(await confirmBox(tr('Open a Google copy instead?', 'เปิดสำเนา Google แทนหรือไม่'),
    tr('The changes made in this page are not in the Google copy. Press Save first to keep them.', 'การแก้ไขในหน้านี้จะไม่อยู่ในสำเนา Google กรุณากดบันทึกก่อนเพื่อเก็บไว้'),
    { ok: tr('Open Google copy', 'เปิดสำเนา Google') }))) return;
  const out = await call('markupCopy', { org: st.org, id: st.id });
  if (st.hold) await presence('handoff', { src: st.id, copy: out.copy }).catch(() => {}); // newcomers join this copy
  stopPresence();
  st.office = null;
  $('[data-root]', el).classList.remove('mk-office');
  await openGoogle(el, { out });
}

// The Google copy page. id = the file the copy is made from (the original, or the file just saved to the site in a
// hand-off); out = markupCopy's answer when already asked; note = a line shown above the steps.
async function openGoogle(el, { id = st.id, out = null, note = '' } = {}) {
  stopPresence();
  out = out || await call('markupCopy', { org: st.org, id });
  st.name = out.name; st.copy = out.copy; st.copyOf = id;
  const app = appName();
  $('[data-name]', el).textContent = st.name;
  st.edited = out.edited;
  status(el, editLine());
  // Back from Google: the last edit again (someone may have edited since).
  if (!st.googleWatch) {
    st.googleWatch = true;
    document.addEventListener('visibilitychange', async () => {
      if (document.hidden || !el.isConnected || st.office || !st.copy) return; // back in this page's editor: nothing to check
      try { st.edited = (await call('markupCopy', { org: st.org, id: st.copyOf, check: true })).edited; status(el, editLine()); } catch { /* keep the last one */ }
    });
  }
  $('[data-root]', el).classList.remove('mk-office');
  $('[data-root]', el).innerHTML = `<section class="group mk-steps">${note ? `<p class="notice blue mk-note">${icon('check', 18)}<span>${esc(note)}</span></p>` : ''}<ol>
      <li>${tr('Open the copy in {app} and make your changes. Google saves them as you type.', 'เปิดสำเนาใน {app} แล้วแก้ไข Google จะบันทึกให้อัตโนมัติ', { app })}
        <p><a class="btn primary" href="${esc(out.url)}" target="_blank" rel="noopener">${icon('ext', 16)}${tr('Open in {app}', 'เปิดใน {app}', { app })}</a></p></li>
      <li>${tr('Come back to this tab and press Save. The file is saved into its card on the organization page.', 'กลับมาที่แท็บนี้แล้วกดบันทึก ไฟล์จะถูกบันทึกลงในการ์ดของหน้าองค์กร')}</li>
    </ol>
    ${st.kind === 'xlsx' ? `<p class="muted small mk-lock">${icon('lock', 14)}${tr('The form’s locks (protected sheets, locked cells, password) are put back from the original when you save.', 'การล็อกของแบบฟอร์ม (ชีตที่ป้องกัน เซลล์ที่ล็อก รหัสผ่าน) จะถูกใส่กลับจากไฟล์ต้นฉบับเมื่อบันทึก')}</p>` : ''}
    <p class="muted small">${tr('Everyone who opens this file here edits the same copy.', 'ทุกคนที่เปิดไฟล์นี้จากที่นี่จะแก้ไขสำเนาเดียวกัน')}</p>
    <p><button class="btn plain" data-back>${tr('Edit in this page instead', 'แก้ไขในหน้านี้แทน')}</button></p></section>`;
  $('[data-back]', el).onclick = () => { st.copy = ''; openOffice(el, { page: true }).catch((err) => errorBox(err.message, tr('Couldn’t open', 'เปิดไม่สำเร็จ'))); };
  $('[data-actions]', el).innerHTML = `<button class="btn" data-a="dl">${icon('download', 16)}${tr('Download', 'ดาวน์โหลด')}</button>
    ${allBtn()}<button class="btn" data-a="site">${icon('upload', 16)}${tr('Save', 'บันทึก')}</button>
    <button class="btn primary" data-a="quit">${tr('Save and quit', 'บันทึกและออก')}</button>`;
  $('[data-actions]', el).onclick = (e) => act(el, e.target.closest('[data-a]')?.dataset.a);
}

// Google copy: "Last edited <when> by <who>" (Google's own record), or "Not edited yet" while it equals the original.
const editLine = () => (st.edited?.at ? tr('Google copy last edited {when} by {who}', 'สำเนา Google แก้ไขล่าสุด {when} โดย {who}',
  { when: dateTimeFmt(st.edited.at), who: esc(st.ctx.nameOf(st.edited.by)) }) : tr('A Google copy for editing, shared with the staff who open it here', 'สำเนา Google สำหรับแก้ไข แชร์ให้เจ้าหน้าที่ที่เปิดจากที่นี่'));

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
    if (a === 'all') await resetAll(el);
    if (a === 'fresh' && await confirmBox(tr('Discard the draft?', 'ทิ้งฉบับร่างหรือไม่'), tr('The file opens again as it is on the site. The draft on this device is deleted.', 'ไฟล์จะเปิดใหม่ตามที่อยู่ในเว็บไซต์ และฉบับร่างในเครื่องนี้จะถูกลบ'), { ok: tr('Discard', 'ทิ้ง'), danger: true })) {
      await writeDraft(null).catch(() => {});
      await openOffice(el, { fresh: true });
    }
    if (a === 'dl') {
      const blob = await busy(tr('Preparing {file}…', 'กำลังเตรียม {file}…', { file: outName() }), produce);
      const link = document.createElement('a');
      link.href = URL.createObjectURL(blob); link.download = outName();
      document.body.append(link); link.click(); link.remove();
      setTimeout(() => URL.revokeObjectURL(link.href), 5000);
    }
    if (a === 'site' || a === 'quit') {
      const file = await saveToSite(el);
      if (file) await hand(el, file);
      if (file && a === 'quit') { location.hash = orgHref(); return; }
    }
  } catch (err) {
    await errorBox(err.message, tr('Couldn’t finish', 'ทำไม่สำเร็จ'));
  }
  btns.forEach((b) => { b.disabled = false; });
}

// Save: the slot's file is checked first (someone may have saved a newer one since this page opened), then
// the file is uploaded. Resolves the saved file { id, name }, or null when the person chose not to replace.
async function saveToSite(el, { ask = true } = {}) {
  if (ask && !(await siteStillSame())) return null;
  const file = await busy(tr('Saving {file} to the site…', 'กำลังบันทึก {file} เข้าเว็บไซต์…', { file: outName() }), async (set) => {
    const f = new File([await produce((x) => set(x * 0.5))], outName(), { type: MIME[st.kind] });
    const saved = await upload(f, target(st.slot, st.org), (x) => set(0.5 + x * 0.4));
    // Into its card at once (v2.16): the card as saved comes back for the organization tab.
    st.card = inCard() ? await call('saveMarked', { org: st.org, slot: st.slot, file: saved }) : null;
    if (st.card) keepCard(st.card);
    if (st.ed) await saveDraft(); // the marks of the saved file stay editable
    if (st.office) { await writeDraft(null).catch(() => {}); st.local = null; } // saved to the site: no draft needed
    return saved;
  });
  st.unsent = false;
  st.saved = { name: file.name, at: new Date().toISOString() };
  (st.own = st.own || new Set()).add(file.id); // my own saves never count as "changed by someone else"
  if (st.ed) status(el, draftLine());
  if (st.office) status(el, tr('Saved · you can keep editing', 'บันทึกแล้ว · แก้ไขต่อได้'));
  return file;
}

// true = go ahead: the slot's file is the one it was when this page opened (or one this page saved), or the person
// chose to replace / add anyway.
async function siteStillSame() {
  if (!st.site) return true;
  let now;
  try { now = (await presence('slot')).site; } catch { return true; } // can't tell: as before
  if (!now?.id || now.id === st.site.id || st.own?.has(now.id)) return true;
  const who = st.ctx.nameOf(now.by) || tr('Someone', 'มีคน'); const when = now.at ? dateTimeFmt(now.at) : '—';
  const ok = await confirmBox(tr('The file on the site has changed', 'ไฟล์บนเว็บไซต์มีการเปลี่ยนแปลง'),
    st.site.adds ? tr('{who} added a file here at {when}, after you opened this page. Add yours as well?', '{who} เพิ่มไฟล์ที่นี่เมื่อ {when} หลังจากคุณเปิดหน้านี้ ต้องการเพิ่มไฟล์ของคุณด้วยหรือไม่', { who, when })
      : tr('{who} saved a newer file on the site at {when}, after you opened this page. Replace it with yours?', '{who} บันทึกไฟล์ที่ใหม่กว่าบนเว็บไซต์เมื่อ {when} หลังจากคุณเปิดหน้านี้ ต้องการแทนที่ด้วยไฟล์ของคุณหรือไม่', { who, when }),
    { ok: st.site.adds ? tr('Add mine', 'เพิ่มของฉัน') : tr('Replace', 'แทนที่'), danger: !st.site.adds });
  if (ok) st.site = now;
  return ok;
}

// ---------- who is editing (v2.13): presence in Code.gs ----------
const BEAT_MS = 20e3; // the editor's heartbeat (Code.gs frees a hold after a minute without one)
const ASK_MS = 4e3; // how often a waiting person asks for the answer
const presence = (op, extra = {}) => call('presence', { op, org: st.org, id: st.id, kind: st.kind, slot: st.slot, ...extra });
const appName = () => (st.kind === 'xlsx' ? 'Google Sheets' : 'Google Docs');
const isPdf = () => st.kind === 'pdf';
const reopen = (el) => { if (isPdf()) { st.pdfOut = null; return openPdf(el); } return openOffice(el); }; // fresh: the newest draft
const nameOf = (email) => st.ctx.nameOf(email);
const names = (list) => (list.length > 1 ? `${list.slice(0, -1).join(', ')}${tr(' and ', ' และ ')}${list[list.length - 1]}` : list[0] || '');

function stopPresence() {
  clearInterval(st?.beatT); clearInterval(st?.askT);
  if (st) { st.beatT = 0; st.askT = 0; st.hold = false; st.waiting = false; }
  closeNotice();
}

// The editor: a heartbeat every 20 s; its answer brings anyone waiting to edit.
function startBeat(el) {
  st.hold = true;
  keepToken();
  st.beatT = setInterval(() => beat(el), BEAT_MS);
}
const keepToken = () => Promise.resolve(idToken?.()).then((t) => { if (t) st.token = t; }).catch(() => {});
async function beat(el) {
  if (document.hidden || !el.isConnected || !st.hold || st.handing) return;
  keepToken();
  let o;
  try { o = await presence('beat'); } catch { return; } // a missed beat: the next one counts
  if (!st.hold) return;
  if (o.lost) return lostHold(el, o.lost);
  if (o.reqs?.length) notice(el, o.reqs); else closeNotice();
}

// The editor's notice: "<name> is trying to edit ... Save this file and switch to Google?" It closes by itself when
// the people asking stop waiting (cancelled, or a minute without an answer).
function notice(el, reqs) {
  const key = reqs.map((r) => r.by).join(',');
  if (st.notice?.key === key) return;
  closeNotice();
  let closed; const gone = new Promise((r) => { closed = r; });
  const who = names(reqs.map((r) => nameOf(r.by)));
  const box = dialog(tr('Someone wants to edit this file', 'มีคนต้องการแก้ไขไฟล์นี้'),
    `<p data-presence>${esc(isPdf() ? (reqs.length > 1
      ? tr('{who} are trying to mark up this PDF with you. Save your marks as a draft and hand the PDF over?', '{who} กำลังพยายามทำเครื่องหมายบน PDF นี้พร้อมกับคุณ บันทึกเครื่องหมายของคุณเป็นฉบับร่างและส่งต่อ PDF หรือไม่', { who })
      : tr('{who} is trying to mark up this PDF with you. Save your marks as a draft and hand the PDF over?', '{who} กำลังพยายามทำเครื่องหมายบน PDF นี้พร้อมกับคุณ บันทึกเครื่องหมายของคุณเป็นฉบับร่างและส่งต่อ PDF หรือไม่', { who }))
      : reqs.length > 1
      ? tr('{who} are trying to edit this document with you. Save this file and switch to {app}?', '{who} กำลังพยายามแก้ไขเอกสารนี้พร้อมกับคุณ บันทึกไฟล์นี้และเปลี่ยนไปใช้ {app} หรือไม่', { who, app: appName() })
      : tr('{who} is trying to edit this document with you. Save this file and switch to {app}?', '{who} กำลังพยายามแก้ไขเอกสารนี้พร้อมกับคุณ บันทึกไฟล์นี้และเปลี่ยนไปใช้ {app} หรือไม่', { who, app: appName() }))}</p>`,
    [{ v: false, label: tr('No', 'ไม่') }, { v: true, label: tr('Yes', 'ใช่'), cls: 'primary' }], { sticky: true });
  st.notice = { key, close: () => { document.querySelector('[data-presence]')?.closest('.modal-back')?.remove(); closed('closed'); } };
  Promise.race([box, gone]).then((v) => {
    if (v === 'closed') return;
    st.notice = null;
    answer(el, v, who);
  });
}
function closeNotice() { const n = st?.notice; if (n) { st.notice = null; n.close(); } }

// Yes: answer, save to the site, make the Google copy from the saved file, then hand off; everyone waiting joins it.
// No: the people asking are told; the editor carries on.
async function answer(el, yes, who) {
  if (!yes) {
    await presence('answer', { yes: false }).catch(() => {});
    toast(tr('You carry on editing; {who} was told.', 'คุณแก้ไขต่อได้ และแจ้ง {who} แล้ว', { who }));
    return;
  }
  st.handing = true;
  if (isPdf()) {
    try {
      await presence('answer', { yes: true });
      await busy(tr('Saving your marks as a draft…', 'กำลังบันทึกเครื่องหมายเป็นฉบับร่าง…'), saveDraft);
      await presence('pass');
      st.handing = false;
      stopPresence();
      st.unsent = false; writeLocal(null); st.ed = null; // the marks are the draft now; the next person goes on from them
      showPassed(el, who);
    } catch (err) {
      st.handing = false;
      await presence('answer', { fail: true }).catch(() => {});
      await errorBox(err.message, tr('Couldn’t hand the PDF over', 'ส่งต่อ PDF ไม่สำเร็จ'));
    }
    return;
  }
  try {
    await presence('answer', { yes: true });
    const file = await saveToSite(el, { ask: true });
    if (!file) throw new Error(tr('Nothing was saved, so the file was not shared.', 'ยังไม่ได้บันทึก จึงยังไม่ได้แชร์ไฟล์'));
    const out = await busy(tr('Making the {app} copy…', 'กำลังสร้างสำเนา {app}…', { app: appName() }), () => call('markupCopy', { org: st.org, id: file.id }));
    await presence('handoff', { src: file.id, copy: out.copy });
    st.handing = false;
    stopPresence();
    st.office = null;
    await openGoogle(el, { id: file.id, out, note: tr('Your edits are saved to the site. You and {who} are editing in {app} now.', 'บันทึกการแก้ไขของคุณเข้าเว็บไซต์แล้ว ตอนนี้คุณและ {who} แก้ไขใน {app}', { who, app: appName() }) });
    await hand(el, file);
  } catch (err) {
    st.handing = false;
    await presence('answer', { fail: true }).catch(() => {});
    await errorBox(err.message, tr('Couldn’t share the file', 'แชร์ไฟล์ไม่สำเร็จ'));
  }
}

// The hold lapsed while this page was away (an iPad asleep) and someone else has the file now: keep the edits as the
// device draft, then wait like anyone else.
async function lostHold(el, by) {
  stopPresence();
  if (st.office?.isDirty()) await saveOfficeDraft().catch(() => {}); // a PDF's marks are kept on this device already
  st.office = null;
  showWaiting(el, by, tr('This page was away for over a minute, so {who} took over. Your edits are kept as a draft on this device.', 'หน้านี้ไม่ได้ใช้งานนานกว่าหนึ่งนาที {who} จึงแก้ไขต่อ การแก้ไขของคุณถูกเก็บเป็นฉบับร่างในเครื่องนี้', { who: nameOf(by) }), { ask: false });
}

// Waiting for the editor's answer: a spinner, a one-minute bar and Cancel; asks every 4 s.
function showWaiting(el, by, note = '', { ask = true } = {}) {
  stopPresence();
  st.office = null; st.ed = null; st.waiting = true;
  const root = $('[data-root]', el);
  root.classList.remove('mk-office');
  root.innerHTML = `<section class="group mk-wait-box" role="status">
      <div class="mk-spin" aria-hidden="true"></div>
      <p class="mk-wait-msg" data-msg><b>${esc(ask ? (isPdf() ? tr('{who} is marking up this PDF. Asking them to hand it over…', '{who} กำลังทำเครื่องหมายบน PDF นี้ กำลังขอให้ส่งต่อ…', { who: nameOf(by) })
    : tr('{who} is editing this file. Asking them to share…', '{who} กำลังแก้ไขไฟล์นี้ กำลังขอให้แชร์…', { who: nameOf(by) })) : tr('{who} is editing this file now.', '{who} กำลังแก้ไขไฟล์นี้อยู่', { who: nameOf(by) }))}</b></p>
      ${note ? `<p class="muted small">${esc(note)}</p>` : ''}
      ${ask ? `<div class="mk-count" aria-hidden="true"><span></span></div>
      <p class="muted small">${esc(isPdf() ? tr('They have about a minute to answer. If they say yes, their marks are saved as a draft and the PDF passes to you.', 'ผู้แก้ไขมีเวลาตอบประมาณหนึ่งนาที หากตอบตกลง เครื่องหมายของเขาจะถูกบันทึกเป็นฉบับร่างและ PDF จะส่งต่อมาให้คุณ')
    : tr('They have about a minute to answer. If they say yes, their edits are saved and you both move to {app}.', 'ผู้แก้ไขมีเวลาตอบประมาณหนึ่งนาที หากตอบตกลง การแก้ไขของเขาจะถูกบันทึกและทั้งคู่จะย้ายไปแก้ไขใน {app}', { app: appName() }))}</p>
      <p><button class="btn" data-cancel>${esc(tr('Cancel', 'ยกเลิก'))}</button></p>`
    : `<p><button class="btn primary" data-retry>${esc(tr('Ask them to share', 'ขอให้แชร์'))}</button></p>`}</section>`;
  readOnlyActions(el);
  status(el, tr('Read only while someone else is editing', 'ดูได้อย่างเดียวขณะที่ผู้อื่นกำลังแก้ไข'));
  $('[data-retry]', el)?.addEventListener('click', () => reopen(el).catch((err) => errorBox(err.message)));
  if (!ask) return;
  $('[data-cancel]', el).onclick = async () => {
    stopPresence();
    await presence('cancel').catch(() => {});
    showRefused(el, by, 'cancel');
  };
  st.askT = setInterval(() => poll(el), ASK_MS);
}
async function poll(el) {
  if (document.hidden || !el.isConnected || !st.waiting) return;
  let o;
  try { o = await presence('ask'); } catch { return; }
  if (!st.waiting) return;
  if (o.state === 'saving') {
    const m = $('[data-msg]', el);
    if (m) m.innerHTML = `<b>${esc(isPdf() ? tr('{who} said yes and is saving their marks…', '{who} ตอบตกลงและกำลังบันทึกเครื่องหมาย…', { who: nameOf(o.by) })
      : tr('{who} said yes and is saving their edits…', '{who} ตอบตกลงและกำลังบันทึกการแก้ไข…', { who: nameOf(o.by) }))}</b>`;
    el.querySelector('.mk-count')?.classList.add('done');
  } else if (o.state === 'yes') joinGoogle(el, o.src, o.by).catch((err) => errorBox(err.message, tr('Couldn’t open', 'เปิดไม่สำเร็จ')));
  else if (o.state === 'no' || o.state === 'timeout') { stopPresence(); showRefused(el, o.by, o.state); }
  else if (o.state === 'pass') {
    stopPresence();
    reopen(el).then(() => toast(tr('{who} saved their marks and handed the PDF to you.', '{who} บันทึกเครื่องหมายและส่งต่อ PDF ให้คุณแล้ว', { who: nameOf(o.by) })))
      .catch((err) => errorBox(err.message, tr('Couldn’t open', 'เปิดไม่สำเร็จ')));
  } else if (o.state === 'free' || o.state === 'gone') { stopPresence(); reopen(el).catch((err) => errorBox(err.message)); }
}

// Moved to the Google copy made from the file the editor just saved.
async function joinGoogle(el, src, by) {
  stopPresence();
  await openGoogle(el, { id: src || st.id, note: tr('{who} saved their edits. You’re both editing in {app} now.', '{who} บันทึกการแก้ไขแล้ว ตอนนี้คุณทั้งคู่แก้ไขใน {app}', { who: nameOf(by), app: appName() }) });
}

// No, no answer, or cancelled: read only (view in Drive's viewer, or download); Try again asks once more.
function showRefused(el, by, why) {
  const who = nameOf(by);
  const line = why === 'no' ? tr('{who} is still editing this file. Try again later, or ask them directly.', '{who} ยังคงแก้ไขไฟล์นี้อยู่ กรุณาลองใหม่ภายหลัง หรือติดต่อผู้แก้ไขโดยตรง', { who })
    : why === 'timeout' ? tr('{who} didn’t answer. Try again later, or ask them directly.', '{who} ไม่ได้ตอบกลับ กรุณาลองใหม่ภายหลัง หรือติดต่อผู้แก้ไขโดยตรง', { who })
      : tr('You stopped asking. {who} is still editing this file.', 'คุณยกเลิกคำขอแล้ว {who} ยังคงแก้ไขไฟล์นี้อยู่', { who });
  const root = $('[data-root]', el);
  root.classList.remove('mk-office');
  root.innerHTML = `<section class="group mk-wait-box"><p class="mk-wait-msg"><b>${esc(line)}</b></p>
      <p class="muted small">${esc(tr('You can view or download the file, but not edit it.', 'คุณดูหรือดาวน์โหลดไฟล์ได้ แต่แก้ไขไม่ได้'))}</p>
      <p><button class="btn primary" data-retry>${esc(tr('Try again', 'ลองอีกครั้ง'))}</button></p></section>`;
  readOnlyActions(el);
  status(el, tr('Read only', 'ดูได้อย่างเดียว'));
  $('[data-retry]', el).onclick = () => reopen(el).catch((err) => errorBox(err.message));
}
function showPassed(el, who) {
  const root = $('[data-root]', el);
  root.innerHTML = `<section class="group mk-wait-box"><p class="mk-wait-msg"><b>${esc(tr('Your marks are saved as a draft. {who} is marking up this PDF now.', 'บันทึกเครื่องหมายของคุณเป็นฉบับร่างแล้ว ตอนนี้ {who} กำลังทำเครื่องหมายบน PDF นี้', { who }))}</b></p>
      <p class="muted small">${esc(tr('You can view or download the PDF. To mark it up again, ask them to hand it back.', 'คุณดูหรือดาวน์โหลด PDF ได้ หากต้องการทำเครื่องหมายอีกครั้ง ให้ขอให้ส่งคืน'))}</p>
      <p><button class="btn primary" data-retry>${esc(tr('Ask to edit again', 'ขอแก้ไขอีกครั้ง'))}</button></p></section>`;
  readOnlyActions(el);
  status(el, tr('Handed over · read only', 'ส่งต่อแล้ว · ดูได้อย่างเดียว'));
  $('[data-retry]', el).onclick = () => reopen(el).catch((err) => errorBox(err.message));
}
function readOnlyActions(el) {
  $('[data-actions]', el).innerHTML = `<button class="btn" data-view>${icon('eye', 16)}${esc(tr('View', 'ดู'))}</button>
    <a class="btn" href="${esc(downloadUrl(st.id))}" download rel="noopener">${icon('download', 16)}${esc(tr('Download', 'ดาวน์โหลด'))}</a>`;
  $('[data-actions]', el).onclick = (e) => {
    if (e.target.closest('[data-view]')) openPreview({ preview: previewUrl(st.id), view: viewUrl(st.id), dl: downloadUrl(st.id), title: st.name });
  };
}

// Back from the background (iPad): the editor beats at once (and finds out if someone took over); a waiting page
// asks at once. Closing the tab lets go of the file straight away.
document.addEventListener('visibilitychange', () => {
  if (document.hidden || !st) return;
  const el = document.querySelector('.markup-page'); if (!el) return;
  if (st.hold) beat(el); else if (st.waiting) poll(el);
});
// Leaving for another page of the site (Back, Save and quit, v2.16): let go of the file at once and stop asking.
export function leave() {
  if (!st) return;
  if (st.hold) presence('release').catch(() => {});
  stopPresence();
  st = null;
}
addEventListener('pagehide', () => {
  if (!st?.hold || !st.token || !navigator.sendBeacon) return;
  navigator.sendBeacon(SCRIPT_URL, new Blob([JSON.stringify({ action: 'presence', op: 'release', org: st.org, id: st.id, idToken: st.token })], { type: 'text/plain;charset=utf-8' }));
});

// The saved file to the organization tab; if no tab took it, say so and offer to send it again.
async function hand(el, file) {
  const msg = { org: st.org, slot: st.slot, file, pitch: st.card?.pitch, orgPitch: st.card?.orgPitch };
  if (st.card) { // saved into its card already: an organization page open in another tab only shows it (no waiting)
    sendMarked(msg);
    toast(tr('Saved into the card', 'บันทึกลงในการ์ดแล้ว'));
    return;
  }
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
