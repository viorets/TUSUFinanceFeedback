// Re-submission: PROTOTYPE cards. The resubmission website and its sheet aren't built yet, so every field that will
// come from them (project name, dates, budgets, files) shows "No data submitted yet". The Approved / Not Approved
// decision is real: it saves at once to our Resubmission tab.
// Top of the tab: the case status, Open or Closed. Admins and Chiefs close / reopen it here (v2.2; it was on the
// organization list before). Closed = nothing in the organization can be changed (Code.gs refuses the writes).
// A case can close only when every project is Approved (v2.4); Code.gs checks too. Reopening is always allowed.
import { tr } from '../i18n.js';
import { $, esc, icon, toast, dateTimeFmt, confirmBox, dialog } from '../ui.js';
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
  const closed = d.closed;
  // Projects not Approved yet (Not Approved or no decision): each one blocks closing.
  const notApproved = () => d.projects.filter((p) => d.resub[p.ref]?.decision !== 'Approved').length;
  const caseBar = () => {
    const n = closed ? 0 : notApproved();
    return `<span class="grow"><b>${tr('Case', 'เคส')}</b>
      ${closed ? `<span class="pill closed">${icon('lock', 13)}${tr('Closed', 'ปิดแล้ว')}</span>` : `<span class="pill green">${tr('Open', 'เปิด')}</span>`}
      ${n && ctx.console ? `<small class="muted case-why">${tr('{n} projects not Approved — every project must be Approved to close the case', 'ยังไม่อนุมัติ {n} โครงการ — ต้องอนุมัติทุกโครงการก่อนปิดเคส', { n })}</small>` : ''}</span>
      ${ctx.console ? `<button class="btn ${closed ? '' : 'danger'}" data-case ${n ? 'disabled' : ''}>${icon('lock', 16)}${closed ? tr('Open case', 'เปิดเคส') : tr('Close case', 'ปิดเคส')}</button>` : ''}`;
  };
  el.innerHTML = `<div class="group casebar" data-casebar>${caseBar()}</div><div data-progress></div><div class="notice blue">${icon('alert', 18)}<span>${tr('Prototype: the resubmission website is not built yet, so no resubmitted data is shown.', 'ต้นแบบ: ยังไม่มีเว็บไซต์สำหรับส่งแก้ไข จึงยังไม่มีข้อมูลที่ส่งแก้ไข')}</span></div>
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
  // Close / open the case: confirm, save, then the whole organization page repaints locked or unlocked.
  el.onclick = async (e) => {
    const b = e.target.closest('[data-case]');
    if (!b) return;
    const ref = d.projects[0].ref.replace(/-P\d+$/, ''); // the tab only shows when there are projects
    const closing = !d.closed;
    const reason = closing ? '' : await reopenReason();
    const ok = closing ? await confirmBox(tr('Close this case?', 'ปิดเคสนี้?'),
      tr('No one can change this organization’s data until the case is opened again.', 'จะไม่มีใครแก้ไขข้อมูลขององค์กรนี้ได้จนกว่าจะเปิดเคสอีกครั้ง'),
      { ok: tr('Close case', 'ปิดเคส'), danger: true }) : !!reason;
    if (!ok) return;
    b.disabled = true;
    try {
      const out = await call('setOrgStatus', { ref, closed: closing, reason });
      d.closed = out.closed;
      forget('boot');
      ctx.persist?.();
      toast(closing ? tr('Case closed', 'ปิดเคสแล้ว') : tr('Case opened', 'เปิดเคสแล้ว'));
      ctx.redraw?.();
    } catch (err) {
      toast(err.message, 'bad');
      b.disabled = false;
    }
  };
  // A click saves at once; on failure the previous choice comes back.
  el.onchange = async (e) => {
    const box = e.target.closest('[data-ref]');
    if (!box || e.target.type !== 'radio') return;
    const ref = box.dataset.ref;
    const radios = [...box.querySelectorAll('input[type=radio]')];
    radios.forEach((r) => { r.disabled = true; });
    try {
      const out = await call('saveResubmit', { ref, decision: e.target.value, base: d.resub[ref]?.at || '' });
      d.resub[ref] = out.resub;
      $('[data-saved]', box).innerHTML = saved(out.resub);
      progress(el, d, 'resubmit');
      $('[data-casebar]', el).innerHTML = caseBar();
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

// Opening a closed case again (v2.8): the reason is required and goes into the Re-submission log. Resolves the
// reason, or '' when cancelled.
async function reopenReason() {
  let why = '';
  for (;;) {
    const p = dialog(tr('Open this case again?', 'เปิดเคสนี้อีกครั้ง?'),
      `<p>${esc(tr('Its data can be changed again. Say why; the reason is kept in the Re-submission log.', 'จะแก้ไขข้อมูลได้อีกครั้ง กรุณาระบุเหตุผล ซึ่งจะบันทึกไว้ในประวัติการแก้ไขการยื่นใหม่'))}</p>
      <div class="field"><label for="reopen-why">${esc(tr('Reason', 'เหตุผล'))}</label><textarea id="reopen-why" maxlength="500">${esc(why)}</textarea></div>`,
      [{ v: false, label: tr('Cancel', 'ยกเลิก') }, { v: true, label: tr('Open case', 'เปิดเคส'), cls: 'primary' }], { none: false });
    const box = document.getElementById('reopen-why');
    box?.addEventListener('input', () => { why = box.value; });
    box?.focus();
    if (!await p) return '';
    if (why.trim()) return why.trim();
    toast(tr('Give a reason for opening the case again', 'กรุณาระบุเหตุผลในการเปิดเคสอีกครั้ง'), 'bad');
  }
}
