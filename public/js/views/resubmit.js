// Re-submission: PROTOTYPE cards. The resubmission website and its sheet aren't built yet, so every field that will
// come from them (project name, dates, budgets, files) shows "No data submitted yet". The Approved / Not Approved
// decision is real: it saves at once to our Resubmission tab.
import { tr } from '../i18n.js';
import { $, esc, icon, toast, dateTimeFmt } from '../ui.js';
import { call, forget } from '../api.js';
import { progress, projNo, star } from '../project.js';

const CHOICES = [['Approved', () => tr('Approved', 'อนุมัติ')], ['Not Approved', () => tr('Not Approved', 'ไม่อนุมัติ')]];
export const NO_DATA = () => tr('No data submitted yet', 'ยังไม่มีข้อมูลที่ส่งเข้ามา');
// Fields the resubmission sheet will fill, in card order.
export const RESUB_FIELDS = () => [
  tr('Project name', 'ชื่อโครงการ'), tr('Project dates', 'วันที่จัดโครงการ'),
  tr('Total project budget proposed', 'งบประมาณที่เสนอ'), tr('Total budget approved', 'งบประมาณที่อนุมัติรวม'),
  tr('Resubmitted PDF', 'PDF ที่ส่งแก้ไข'), tr('Resubmitted .docx', '.docx ที่ส่งแก้ไข'),
];

export function render(el, d, ctx) {
  d.resub = d.resub || {};
  const saved = (s) => (s?.at ? tr('Saved {when} by {who}', 'บันทึก {when} โดย {who}', { when: dateTimeFmt(s.at), who: esc(ctx.nameOf(s.by)) }) : tr('No decision yet', 'ยังไม่ได้ตัดสิน'));
  el.innerHTML = `<div data-progress></div><div class="notice blue">${icon('alert', 18)}<span>${tr('Prototype: the resubmission website is not built yet, so no resubmitted data is shown.', 'ต้นแบบ: ยังไม่มีเว็บไซต์สำหรับส่งแก้ไข จึงยังไม่มีข้อมูลที่ส่งแก้ไข')}</span></div>
    ${d.projects.map((p) => {
    const s = d.resub[p.ref];
    return `<section class="group pcard" data-ref="${esc(p.ref)}">
      <header class="pcard-head"><span class="eyebrow">${esc(tr('Project {n}', 'โครงการที่ {n}', { n: projNo(p.ref) }))}</span><span class="code">${esc(p.ref)}</span></header>
      <div class="pcard-body"><dl class="kv nodata">${RESUB_FIELDS().map((k) => `<dt>${esc(k)}</dt><dd class="muted">${esc(NO_DATA())}</dd>`).join('')}</dl></div>
      <div class="fb"><div class="field"><label>${tr('Decision', 'ผลการพิจารณา')}${star(d, 'resub.decision')}</label>
        <div><div class="seg" role="radiogroup">${CHOICES.map(([v, label]) => `<label><input type="radio" name="rs-${esc(p.ref)}" value="${v}" ${s?.decision === v ? 'checked' : ''} ${ctx.canEdit ? '' : 'disabled'}><span>${esc(label())}</span></label>`).join('')}</div>
        <p class="muted small" data-saved>${saved(s)}</p></div></div></div>
    </section>`;
  }).join('')}`;
  progress(el, d, 'resubmit');
  // A click saves at once; on failure the previous choice comes back.
  el.onchange = async (e) => {
    const box = e.target.closest('[data-ref]');
    if (!box || e.target.type !== 'radio') return;
    const ref = box.dataset.ref;
    const radios = [...box.querySelectorAll('input[type=radio]')];
    radios.forEach((r) => { r.disabled = true; });
    try {
      const out = await call('saveResubmit', { ref, decision: e.target.value });
      d.resub[ref] = out.resub;
      $('[data-saved]', box).innerHTML = saved(out.resub);
      progress(el, d, 'resubmit');
      ctx.persist?.();
      forget('dashboard');
      toast(tr('Saved', 'บันทึกแล้ว'));
    } catch (err) {
      toast(err.message, 'bad');
      radios.forEach((r) => { r.checked = r.value === d.resub[ref]?.decision; });
    }
    radios.forEach((r) => { r.disabled = !ctx.canEdit; });
  };
}
