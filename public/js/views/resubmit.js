// Re-submission: PROTOTYPE cards. The resubmission website and its sheet aren't built yet, so the cards show the
// current submission as sample data, "Total budget approved" stays empty until that sheet is connected, and the
// download buttons stay off. The Approved / Not Approved decision is real: it saves at once to our Resubmission tab.
import { tr } from '../i18n.js';
import { $, esc, icon, toast, dateTimeFmt } from '../ui.js';
import { call, forget } from '../api.js';
import { cardTop, dlBtn, progress } from '../project.js';

const CHOICES = [['Approved', () => tr('Approved', 'อนุมัติ')], ['Not Approved', () => tr('Not Approved', 'ไม่อนุมัติ')]];

export function render(el, d, ctx) {
  d.resub = d.resub || {};
  const saved = (s) => (s?.at ? tr(`Saved ${dateTimeFmt(s.at)} by ${esc(ctx.nameOf(s.by))}`, `บันทึก ${dateTimeFmt(s.at)} โดย ${esc(ctx.nameOf(s.by))}`) : tr('No decision yet', 'ยังไม่ได้ตัดสิน'));
  el.innerHTML = `<div data-progress></div><div class="notice blue">${icon('alert', 18)}<span>${tr('Prototype: the resubmission website is not built yet. Cards show sample data from the current submission; Total budget approved will come from the resubmission sheet.', 'ต้นแบบ: ยังไม่มีเว็บไซต์สำหรับส่งแก้ไข การ์ดแสดงข้อมูลตัวอย่างจากการส่งครั้งปัจจุบัน งบประมาณที่อนุมัติรวมจะดึงจากชีตการส่งแก้ไข')}</span></div>
    ${d.projects.map((p) => {
    const s = d.resub[p.ref];
    return `<section class="group pcard" data-ref="${esc(p.ref)}">
      ${cardTop(p, [[tr('Total project budget proposed', 'งบประมาณที่เสนอ'), p.budget], [tr('Total budget approved', 'งบประมาณที่อนุมัติรวม'), null]],
    `${dlBtn('', tr('Resubmitted PDF', 'PDF ที่ส่งแก้ไข'))}${dlBtn('', tr('Resubmitted .docx', '.docx ที่ส่งแก้ไข'))}`)}
      <div class="fb"><div class="field"><label>${tr('Decision', 'ผลการพิจารณา')}</label>
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
