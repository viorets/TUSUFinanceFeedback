// Pre-pitch: each project as submitted, then the feedback form (every role can edit).
// A No on .docx / PDF asks for a feedback file, uploaded straight to Drive when picked; Save writes the row and
// the server logs every changed field in PrePitch_Log. "Finished" is decided by the server (Code.gs finished_).
import { tr } from '../i18n.js';
import { $, esc, icon, toast, moneyBox, parseMoney, liveMoney, dateTimeFmt } from '../ui.js';
import { call, upload, forget } from '../api.js';
import { cardTop, dlBtn, fileChip, progress } from '../project.js';

let st = null; // { el, d, can, form: { ref: working copy } }

const blank = { approvedBudget: null, docxOk: '', docxFile: null, pdfOk: '', pdfFile: null, note: '' };
const pick = (x) => ({ approvedBudget: x.approvedBudget ?? null, docxOk: x.docxOk || '', docxFile: x.docxFile || null, pdfOk: x.pdfOk || '', pdfFile: x.pdfFile || null, note: x.note || '' });
const dirty = (ref) => JSON.stringify(st.form[ref]) !== JSON.stringify(pick(st.d.pre[ref] || blank));
export const isDirty = () => !!st && st.d.projects.some((p) => dirty(p.ref));

export function render(el, d, ctx) {
  st = { el, d, ctx, can: ctx.canEdit, form: Object.fromEntries(d.projects.map((p) => [p.ref, pick(d.pre[p.ref] || blank)])) };
  el.innerHTML = `<div data-progress></div>${d.projects.map(card).join('')}`;
  progress(el, d, 'prepitch');
  el.oninput = onInput;
  el.onchange = onChange;
  el.onclick = onClick;
}

function card(p) {
  const f = st.form[p.ref];
  const saved = st.d.pre[p.ref];
  const dis = st.can ? '' : 'disabled';
  const yn = (k) => `<div class="seg" role="radiogroup">${['Yes', 'No'].map((v) => `<label><input type="radio" name="${k}-${esc(p.ref)}" value="${v}" data-k="${k}Ok" ${f[`${k}Ok`] === v ? 'checked' : ''} ${dis}><span>${v === 'Yes' ? tr('Yes', 'ใช่') : tr('No', 'ไม่')}</span></label>`).join('')}</div>`;
  return `<section class="group pcard" data-ref="${esc(p.ref)}">
    ${cardTop(p, [[tr('Total proposed project budget', 'งบประมาณโครงการที่เสนอ'), p.budget]], `${dlBtn(p.pdf, tr('Project PDF', 'โครงการ PDF'))}${dlBtn(p.docx, tr('Project .docx', 'โครงการ .docx'))}`)}
    <div class="fb"><h4>${tr('Feedback', 'ข้อเสนอแนะ')} <span data-status>${status(saved)}</span></h4>
      <div class="field"><label>${tr('Approved budget', 'งบประมาณที่อนุมัติ')}</label><input type="text" inputmode="decimal" data-k="approvedBudget" value="${esc(moneyBox(f.approvedBudget))}" placeholder="0.00" ${dis}></div>
      <div class="field"><label>${tr('Approved project .docx?', 'อนุมัติไฟล์โครงการ .docx?')}</label><div>${yn('docx')}<div data-slot="docx">${slot(p.ref, 'docx')}</div></div></div>
      <div class="field"><label>${tr('Approved project PDF?', 'อนุมัติไฟล์โครงการ PDF?')}</label><div>${yn('pdf')}<div data-slot="pdf">${slot(p.ref, 'pdf')}</div></div></div>
      <div class="field"><label>${tr('Feedback note', 'บันทึกข้อเสนอแนะ')}</label><textarea data-k="note" maxlength="5000" ${dis}>${esc(f.note)}</textarea></div>
      ${st.can ? `<div class="actions"><span class="muted small" data-saved>${savedLine(saved)}</span><button class="btn primary" data-save>${tr('Save', 'บันทึก')}</button></div>`
    : `<div class="actions"><span class="muted small">${savedLine(saved)}</span></div>`}
    </div></section>`;
}

const status = (s) => (s?.finished ? `<span class="pill green">${tr('Finished', 'ครบแล้ว')}</span>` : `<span class="pill orange">${tr('Unfinished', 'ยังไม่ครบ')}</span>`);
const savedLine = (s) => (s?.at ? tr(`Saved ${dateTimeFmt(s.at)} by ${esc(st.ctx.nameOf(s.by))}`, `บันทึก ${dateTimeFmt(s.at)} โดย ${esc(st.ctx.nameOf(s.by))}`) : tr('Not saved yet', 'ยังไม่ได้บันทึก'));

// Upload area under a Yes/No: shown only for No.
function slot(ref, k) {
  const f = st.form[ref];
  if (f[`${k}Ok`] !== 'No') return '';
  const file = f[`${k}File`];
  const label = k === 'docx' ? tr('Upload feedback .docx', 'อัปโหลดไฟล์ข้อเสนอแนะ .docx') : tr('Upload feedback PDF', 'อัปโหลดไฟล์ข้อเสนอแนะ PDF');
  return `<div class="upl">${fileChip(file)}${st.can ? `<label class="btn">${icon('upload', 16)}${file ? tr('Replace', 'เปลี่ยนไฟล์') : esc(label)}
    <input type="file" hidden accept="${k === 'docx' ? '.docx' : '.pdf'}" data-up="${k}"></label>` : ''}
    ${!file ? `<span class="small warn">${tr('Required to finish', 'ต้องมีไฟล์จึงจะครบ')}</span>` : ''}</div>`;
}

const refOf = (e) => e.target.closest('[data-ref]')?.dataset.ref;

function onInput(e) {
  const ref = refOf(e); const k = e.target.dataset.k;
  if (!ref || !k) return;
  if (k === 'approvedBudget') { liveMoney(e.target); st.form[ref].approvedBudget = parseMoney(e.target.value); }
  if (k === 'note') st.form[ref].note = e.target.value;
}

async function onChange(e) {
  const ref = refOf(e);
  if (!ref) return;
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
  box.querySelector('.upl')?.classList.add('busy');
  try {
    st.form[ref][`${kind}File`] = await upload(file, { ref, purpose: 'feedback', only: kind });
    toast(tr('Uploaded to Drive — press Save to keep it', 'อัปโหลดไปยัง Drive แล้ว — กดบันทึกเพื่อเก็บ'));
  } catch (err) {
    toast(err.message, 'bad');
  }
  box.innerHTML = slot(ref, kind);
}

async function onClick(e) {
  const btn = e.target.closest('[data-save]');
  if (!btn) return;
  const ref = refOf(e);
  const sec = btn.closest('.pcard');
  btn.disabled = true;
  try {
    const out = await call('savePrePitch', { ref, data: st.form[ref] });
    st.d.pre[ref] = out.pre;
    st.form[ref] = pick(out.pre);
    $('[data-status]', sec).innerHTML = status(out.pre);
    $('[data-saved]', sec).textContent = savedLine(out.pre);
    progress(st.el, st.d, 'prepitch');
    forget('boot'); // list counts
    forget('dashboard');
    st.ctx.persist?.();
    toast(out.changed ? tr('Saved', 'บันทึกแล้ว') : tr('Nothing changed', 'ไม่มีการเปลี่ยนแปลง'));
  } catch (err) {
    toast(err.message, 'bad');
  }
  btn.disabled = false;
}
