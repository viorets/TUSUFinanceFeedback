// Edit a file in the site (v2.12), opened in its own tab from an organization page (editBtn in project.js):
// #/markup?org=<Org Ref>&id=<Drive file id>&kind=pdf|docx|xlsx&slot=<where the result goes>.
// PDF: the markup editor (markup/pdf-markup.js). Its marks are a draft saved on the server (Save draft, and with every
// Save to site) and kept on this device after each change, so a closed tab loses nothing.
// The editors come from TUSU Finance Team Tools (copied into markup/; that repository is the source of truth).
// PDF: Edit text and OCR are off here; text boxes are written as real Thai text. Files up to 12 MB.
// .docx / .xlsx: edited in the page (markup/docx-edit.js: pages, headers and footers, page numbers; markup/xlsx-edit.js:
// formulas, formats; an Excel file keeps its original's locks, markup/xlsx-relock.js). "Edit in Google instead" keeps
// the v2.12 route for files the editors here cannot keep (charts, pictures, pivot tables): a Google Docs / Sheets copy
// in the organization's folder, shared with the signed-in Gmail only.
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

let st = null; // { org, id, kind, slot, name, ed (PDF editor), office (Excel/Word editor), draft: { at, by } | null, copy (Google copy id), ctx }

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
  const [{ openPdfEditor }, { setFontDir }, out] = await Promise.all([import('../markup/pdf-markup.js'), import('../markup/pdf-realtext.js'), call('markupOpen', { org: st.org, id: st.id, bytes: true }, { wait: 90e3 })]);
  setFontDir('fonts/'); // the fonts are in public/fonts, beside js/
  st.name = out.name;
  st.draft = out.draft ? { at: out.draft.at, by: out.draft.by } : null;
  $('[data-name]', el).textContent = st.name;
  // Marks on this device win when they were made on top of the draft the server still has.
  const local = readLocal();
  const useLocal = local && (local.base || '') === (st.draft?.at || '');
  st.ed = await openPdfEditor($('[data-root]', el), {
    bytes: await bytesOf(out.data), name: st.name, draft: useLocal ? local.marks : out.draft?.marks || null, t: tr,
    onDraft: (marks) => writeLocal({ marks, base: st.draft?.at || '' }), defaults: out.defaults,
    editText: false, ocr: false, maxBytes: 12 * 1024 * 1024, onStatus: (m) => status(el, esc(m)),
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

// Excel and Word in the page. The edits live in this tab until Save to site (or Download); leaving asks first.
async function openOffice(el) {
  const kind = st.kind;
  const [mod, rb, out] = await Promise.all([import(kind === 'xlsx' ? '../markup/xlsx-edit.js' : '../markup/docx-edit.js'), import('../markup/ribbon.js'),
    call('markupOpen', { org: st.org, id: st.id, bytes: true }, { wait: 90e3 })]);
  rb.setRibbonLang(tr);
  st.name = out.name; $('[data-name]', el).textContent = st.name;
  const bytes = await bytesOf(out.data);
  const root = $('[data-root]', el); root.innerHTML = '';
  try { st.office = await (kind === 'xlsx' ? mod.openXlsxEditor(root, bytes) : mod.openDocxEditor(root, bytes)); }
  catch (err) { root.innerHTML = `<p class="muted">${esc(err.message)}</p>`; }
  // What the editor here would not keep: charts, pictures and pivot tables of a workbook (then Google is better).
  let lost = '';
  if (kind === 'xlsx') {
    try {
      if (!window.JSZip) await new Promise((res, rej) => { const sc = document.createElement('script'); sc.src = 'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js'; sc.crossOrigin = 'anonymous'; sc.onload = res; sc.onerror = rej; document.head.append(sc); });
      const names = Object.keys((await window.JSZip.loadAsync(bytes)).files);
      const n = (re) => names.filter((x) => re.test(x)).length;
      lost = [[n(/^xl\/charts\/chart\d+\.xml$/), tr('charts', 'แผนภูมิ')], [n(/^xl\/media\//), tr('pictures', 'รูปภาพ')], [n(/^xl\/pivotTables\//), tr('pivot tables', 'ตาราง Pivot')]].filter(([k]) => k).map(([k, w]) => `${k} ${w}`).join(', ');
    } catch { /* keep going */ }
  }
  status(el, lost ? `<span class="warn">${tr('Not kept if edited here: {what}. Use Edit in Google instead.', 'ถ้าแก้ไขที่นี่จะไม่เก็บ: {what} กรุณาใช้แก้ไขใน Google แทน', { what: esc(lost) })}</span>`
    : tr('Edited here, in this tab: press Save to site when done', 'แก้ไขที่นี่ในแท็บนี้ กดบันทึกเข้าเว็บไซต์เมื่อเสร็จ'));
  $('[data-actions]', el).innerHTML = `<button class="btn" data-a="google">${icon('ext', 16)}${tr('Edit in Google instead', 'แก้ไขใน Google แทน')}</button>
    <button class="btn" data-a="dl">${icon('download', 16)}${tr('Download', 'ดาวน์โหลด')}</button>
    <button class="btn primary" data-a="site">${icon('upload', 16)}${tr('Save to site', 'บันทึกเข้าเว็บไซต์')}</button>`;
  $('[data-actions]', el).onclick = (e) => act(el, e.target.closest('[data-a]')?.dataset.a);
}

async function openGoogle(el) {
  st.office = null;
  const out = await call('markupCopy', { org: st.org, id: st.id });
  st.name = out.name; st.copy = out.copy;
  const app = st.kind === 'xlsx' ? 'Google Sheets' : 'Google Docs';
  $('[data-name]', el).textContent = st.name;
  st.edited = out.edited;
  status(el, editLine());
  // Back from Google: the last edit again (someone may have edited since).
  document.addEventListener('visibilitychange', async () => {
    if (document.hidden || !el.isConnected) return;
    try { st.edited = (await call('markupCopy', { org: st.org, id: st.id, check: true })).edited; status(el, editLine()); } catch { /* keep the last one */ }
  });
  $('[data-root]', el).innerHTML = `<section class="group mk-steps"><ol>
      <li>${tr('Open the copy in {app} and make your changes. Google saves them as you type.', 'เปิดสำเนาใน {app} แล้วแก้ไข Google จะบันทึกให้อัตโนมัติ', { app })}
        <p><a class="btn primary" href="${esc(out.url)}" target="_blank" rel="noopener">${icon('ext', 16)}${tr('Open in {app}', 'เปิดใน {app}', { app })}</a></p></li>
      <li>${tr('Come back to this tab and press Save to site. The file goes to the organization page, where you press Save.', 'กลับมาที่แท็บนี้แล้วกดบันทึกเข้าเว็บไซต์ ไฟล์จะไปอยู่ในหน้าองค์กร แล้วกดบันทึกที่นั่น')}</li>
    </ol>
    ${st.kind === 'xlsx' ? `<p class="muted small mk-lock">${icon('lock', 14)}${tr('The form’s locks (protected sheets, locked cells, password) are put back from the original when you save.', 'การล็อกของแบบฟอร์ม (ชีตที่ป้องกัน เซลล์ที่ล็อก รหัสผ่าน) จะถูกใส่กลับจากไฟล์ต้นฉบับเมื่อบันทึก')}</p>` : ''}
    <p class="muted small">${tr('Everyone who opens this file here edits the same copy.', 'ทุกคนที่เปิดไฟล์นี้จากที่นี่จะแก้ไขสำเนาเดียวกัน')}</p></section>`;
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
  if (st.office) { set(0.3); return new Blob([await st.office.save()], { type: MIME[st.kind] }); } // xlsx: locks put back in save()
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
    if (a === 'google') {
      if (st.office?.isDirty() && !(await confirmBox(tr('Leave the edits made here and edit a Google copy instead?', 'ทิ้งการแก้ไขที่ทำไว้ที่นี่ แล้วแก้ไขสำเนา Google แทน?')))) { btns.forEach((b) => { b.disabled = false; }); return; }
      await busy(tr('Making a Google copy…', 'กำลังสร้างสำเนา Google…'), () => openGoogle(el));
    }
    if (a === 'draft') {
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
        return saved;
      });
      if (st.ed) status(el, draftLine());
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
