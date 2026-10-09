// Pitching: per project, the pitching slot (from the scheduling site, read only), the summary, the final approved
// budget and "Noted documents" (any number of .docx / PDF). Unfinished = no final approved budget yet.
import { tr } from '../i18n.js';
import { $, esc, icon, toast, moneyBox, parseMoney, liveMoney, dateTimeFmt } from '../ui.js';
import { call, upload, forget } from '../api.js';
import { cardTop, fileChip, progress, slotLine } from '../project.js';

let st = null; // { el, d, can, form: { ref: working copy } }

const pick = (x = {}) => ({ summary: x.summary || '', finalBudget: x.finalBudget ?? null, noted: [...(x.noted || [])] });
const dirty = (ref) => JSON.stringify(st.form[ref]) !== JSON.stringify(pick(st.d.pitch[ref]));
export const isDirty = () => !!st && st.d.projects.some((p) => dirty(p.ref));

export function render(el, d, ctx) {
  st = { el, d, ctx, can: ctx.canEdit, form: Object.fromEntries(d.projects.map((p) => [p.ref, pick(d.pitch[p.ref])])) };
  el.innerHTML = `<div data-progress></div>${d.projects.map(card).join('')}`;
  progress(el, d, 'pitching');
  el.oninput = onInput;
  el.onchange = onChange;
  el.onclick = onClick;
}

function card(p) {
  const f = st.form[p.ref];
  const dis = st.can ? '' : 'disabled';
  return `<section class="group pcard" data-ref="${esc(p.ref)}">
    ${cardTop(p, [[tr('Total project budget', 'งบประมาณโครงการ'), p.budget]], '', slotLine(st.d.slots?.[p.ref]))}
    <div class="fb">
      <div class="field"><label>${tr('Summary', 'สรุป')}</label><textarea data-k="summary" maxlength="20000" ${dis}>${esc(f.summary)}</textarea></div>
      <div class="field"><label>${tr('Final approved budget', 'งบประมาณอนุมัติขั้นสุดท้าย')}</label><input type="text" inputmode="decimal" data-k="finalBudget" value="${esc(moneyBox(f.finalBudget))}" placeholder="0.00" ${dis}></div>
      <div class="field"><label>${tr('Noted documents', 'เอกสารบันทึก')}</label><div data-noted>${noted(p.ref)}</div></div>
      <div class="actions"><span class="muted small" data-saved>${savedLine(st.d.pitch[p.ref])}</span>${st.can ? `<button class="btn primary" data-save>${tr('Save', 'บันทึก')}</button>` : ''}</div>
    </div></section>`;
}

const savedLine = (s) => (s?.at ? tr(`Saved ${dateTimeFmt(s.at)} by ${esc(st.ctx.nameOf(s.by))}`, `บันทึก ${dateTimeFmt(s.at)} โดย ${esc(st.ctx.nameOf(s.by))}`) : tr('Not saved yet', 'ยังไม่ได้บันทึก'));

function noted(ref) {
  const list = st.form[ref].noted;
  return `<div class="upl">${list.map((f, i) => `${fileChip(f)}${st.can ? `<button class="icon-btn" data-rm="${i}" title="${esc(tr('Remove from list', 'นำออกจากรายการ'))}">${icon('x', 16)}</button>` : ''}`).join('')}
    ${!list.length && !st.can ? '<span class="muted">—</span>' : ''}
    ${st.can ? `<label class="btn">${icon('upload', 16)}${tr('Upload .docx or PDF', 'อัปโหลด .docx หรือ PDF')}<input type="file" hidden multiple accept=".docx,.pdf" data-up></label>` : ''}</div>`;
}

const refOf = (e) => e.target.closest('[data-ref]')?.dataset.ref;

function onInput(e) {
  const ref = refOf(e); const k = e.target.dataset.k;
  if (!ref || !k) return;
  if (k === 'finalBudget') { liveMoney(e.target); st.form[ref].finalBudget = parseMoney(e.target.value); }
  if (k === 'summary') st.form[ref].summary = e.target.value;
}

async function onChange(e) {
  const ref = refOf(e);
  if (!ref || !e.target.hasAttribute('data-up')) return;
  const box = e.target.closest('[data-noted]');
  box.querySelector('.upl')?.classList.add('busy');
  for (const file of [...e.target.files]) {
    try {
      st.form[ref].noted.push(await upload(file, { ref, purpose: 'noted' }));
    } catch (err) {
      toast(`${file.name}: ${err.message}`, 'bad');
    }
  }
  box.innerHTML = noted(ref);
  if (st.form[ref].noted.length) toast(tr('Uploaded to Drive — press Save to keep it', 'อัปโหลดไปยัง Drive แล้ว — กดบันทึกเพื่อเก็บ'));
}

async function onClick(e) {
  const ref = refOf(e);
  if (!ref) return;
  const rm = e.target.closest('[data-rm]');
  if (rm) {
    st.form[ref].noted.splice(Number(rm.dataset.rm), 1);
    rm.closest('[data-noted]').innerHTML = noted(ref);
    return;
  }
  const btn = e.target.closest('[data-save]');
  if (!btn) return;
  btn.disabled = true;
  try {
    const out = await call('savePitching', { ref, data: st.form[ref] });
    st.d.pitch[ref] = out.pitch;
    progress(st.el, st.d, 'pitching');
    st.ctx.persist?.();
    forget('dashboard');
    st.form[ref] = pick(out.pitch);
    $('[data-saved]', btn.closest('.pcard')).textContent = savedLine(out.pitch);
    toast(tr('Saved', 'บันทึกแล้ว'));
  } catch (err) {
    toast(err.message, 'bad');
  }
  btn.disabled = false;
}
