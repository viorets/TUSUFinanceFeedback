// Re-submission: one card per project. v2.13: the re-submitted data is the organization's latest Re-Submission round
// on the pitching site (round codes in the admin console): project name, dates, proposed budget and the two files.
// A project with nothing re-submitted shows "No data submitted yet". The Approved / Not Approved decision saves at
// once to our Resubmission tab.
// Top of the tab: the case status, Open or Closed. Admins, Chiefs and Data team staff (v2.11) close / reopen it here (v2.2; it was on the
// organization list before). Closed = nothing in the organization can be changed (Code.gs refuses the writes).
// A case can close only when every project is Approved (v2.4); Code.gs checks too. Reopening is always allowed.
import { tr } from '../i18n.js';
import { $, esc, icon, toast, dateTimeFmt, confirmBox, dialog, saving, errorBox } from '../ui.js';
import { call, save as send, forget } from '../api.js';
import { progress, projNo, star, cardTop, dlBtn } from '../project.js';

const CHOICES = [['Approved', () => tr('Approved', 'อนุมัติ')], ['Not Approved', () => tr('Not Approved', 'ไม่อนุมัติ')]];
export const NO_DATA = () => tr('No data submitted yet', 'ยังไม่มีข้อมูลที่ส่งเข้ามา');
// Fields the resubmission sheet will fill, in card order.
// Where the re-submitted data comes from, in one line (this tab and the print).
export const resubHead = (d) => (d.resubmitted
  ? tr('Re-submitted in round {round}, version {v}, on {when}', 'ส่งแก้ไขในรอบ {round} ฉบับที่ {v} เมื่อ {when}', { round: d.resubmitted.round, v: d.resubmitted.version ?? '—', when: d.resubmitted.at ? dateTimeFmt(d.resubmitted.at) : '—' })
  : NO_DATA());
export const RESUB_FIELDS = () => [
  tr('Project name', 'ชื่อโครงการ'), tr('Project dates', 'วันที่จัดโครงการ'),
  tr('Total project budget proposed', 'งบประมาณที่เสนอ'), tr('Total budget approved', 'งบประมาณที่อนุมัติรวม'),
  tr('Resubmitted PDF', 'PDF ที่ส่งแก้ไข'), tr('Resubmitted .docx', '.docx ที่ส่งแก้ไข'),
];

export function render(el, d, ctx) {
  d.resub = d.resub || {};
  const fbLine = (s) => (s?.decision === 'Not Approved' && s.feedback ? `<span class="fb-note"><b>${tr('Feedback', 'ข้อเสนอแนะ')}:</b> ${esc(s.feedback).replace(/\n/g, '<br>')}</span><br>` : '');
  const saved = (s) => fbLine(s) + (s?.pending ? `<span class="pill orange">${tr('Pending', 'รอส่ง')}</span> ${tr('Kept on this device; it is sent when the server can be reached', 'เก็บไว้ในเครื่องนี้ จะส่งเมื่อติดต่อเซิร์ฟเวอร์ได้')}` : s?.at ? tr('Saved {when} by {who}', 'บันทึก {when} โดย {who}', { when: dateTimeFmt(s.at), who: esc(ctx.nameOf(s.by)) }) : tr('No decision yet', 'ยังไม่ได้ตัดสิน'));
  const closed = d.closed;
  // Projects not Approved yet (Not Approved or no decision): each one blocks closing.
  const notApproved = () => d.projects.filter((p) => d.resub[p.ref]?.decision !== 'Approved').length;
  const caseBar = () => {
    const n = closed ? 0 : notApproved();
    return `<span class="grow"><b>${tr('Case', 'เคส')}</b>
      ${closed ? `<span class="pill closed">${icon('lock', 13)}${tr('Closed', 'ปิดแล้ว')}</span>` : `<span class="pill green">${tr('Open', 'เปิด')}</span>`}
      ${n && ctx.canCase ? `<small class="muted case-why">${tr('{n} projects not Approved — every project must be Approved to close the case', 'ยังไม่อนุมัติ {n} โครงการ — ต้องอนุมัติทุกโครงการก่อนปิดเคส', { n })}</small>` : ''}</span>
      ${ctx.canCase ? `<button class="btn ${closed ? '' : 'danger'}" data-case ${n ? 'disabled' : ''}>${icon('lock', 16)}${closed ? tr('Open case', 'เปิดเคส') : tr('Close case', 'ปิดเคส')}</button>` : ''}`;
  };
  el.innerHTML = `<div class="group casebar" data-casebar>${caseBar()}</div><div data-progress></div><div class="notice blue">${icon(d.resubmitted ? 'check' : 'alert', 18)}<span>${esc(resubHead(d))}</span></div>
    ${d.projects.map((p) => {
    const s = d.resub[p.ref];
    const r = d.resubmitted?.projects?.[p.ref];
    const [, , proposed, approved, pdf, docx] = RESUB_FIELDS();
    const top = r ? cardTop({ ...r, ref: p.ref }, [[proposed, r.budget], [approved, d.pitch?.[p.ref]?.finalBudget]], dlBtn(r.pdf, pdf) + dlBtn(r.docx, docx))
      : `<header class="pcard-head"><span class="eyebrow">${esc(tr('Project {n}', 'โครงการที่ {n}', { n: projNo(p.ref) }))}</span><h3>${esc(p.name || '—')}</h3><span class="code">${esc(p.ref)}</span></header>
      <div class="pcard-body"><dl class="kv nodata">${RESUB_FIELDS().map((k) => `<dt>${esc(k)}</dt><dd class="muted">${esc(NO_DATA())}</dd>`).join('')}</dl></div>`;
    return `<section class="group pcard" data-ref="${esc(p.ref)}">
      ${top}
      <div class="fb"><div class="field"><label>${tr('Decision', 'ผลการพิจารณา')}${star(d, 'resub.decision')}</label>
        <div><div class="seg" role="radiogroup">${CHOICES.map(([v, label]) => `<label><input type="radio" name="rs-${esc(p.ref)}" value="${v}" ${s?.decision === v ? 'checked' : ''} ${ctx.canEdit ? '' : 'disabled'}><span>${esc(label())}</span></label>`).join('')}</div>
        <p class="muted small" data-saved>${saved(s)}</p>
        ${s?.decision === 'Not Approved' && ctx.canEdit ? `<button class="btn plain" data-fb>${tr('Edit feedback', 'แก้ไขข้อเสนอแนะ')}</button>` : ''}</div></div></div>
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
      const out = await saving(() => call('setOrgStatus', { ref, closed: closing, reason }));
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
  // Not Approved asks for feedback first (v2.10): required, saved with the decision.
  const decide = async (box, decision, feedback) => {
    const ref = box.dataset.ref;
    const radios = [...box.querySelectorAll('input[type=radio]')];
    radios.forEach((r) => { r.disabled = true; });
    const prev = d.resub[ref];
    try {
      let out;
      try {
        out = await saving(() => send('saveResubmit', { ref, decision, feedback, base: prev?.at || '' }, `resub:${ref}`));
      } catch (err) {
        if (err.code !== 'QUEUED') throw err;
        out = { resub: { decision, feedback, at: prev?.at || '', by: prev?.by || '', pending: true }, queued: true };
      }
      d.resub[ref] = out.resub;
      $('[data-saved]', box).innerHTML = saved(out.resub);
      $('[data-fb]', box)?.remove();
      if (out.resub.decision === 'Not Approved' && ctx.canEdit) $('[data-saved]', box).insertAdjacentHTML('afterend', `<button class="btn plain" data-fb>${tr('Edit feedback', 'แก้ไขข้อเสนอแนะ')}</button>`);
      progress(el, d, 'resubmit');
      $('[data-casebar]', el).innerHTML = caseBar();
      ctx.persist?.();
      forget('dashboard');
      forget('boot');
      toast(out.queued ? tr('Kept on this device — it will be sent automatically', 'เก็บไว้ในเครื่องนี้ — จะส่งให้อัตโนมัติ') : tr('Saved', 'บันทึกแล้ว'));
    } catch (err) {
      await errorBox(err.message);
      radios.forEach((r) => { r.checked = r.value === d.resub[ref]?.decision; });
    }
    radios.forEach((r) => { r.disabled = !ctx.canEdit; });
  };
  el.onchange = async (e) => {
    const box = e.target.closest('[data-ref]');
    if (!box || e.target.type !== 'radio') return;
    if (e.target.value !== 'Not Approved') { await decide(box, e.target.value, ''); return; }
    const fb = await askFeedback(d.resub[box.dataset.ref]?.feedback || '');
    if (fb) await decide(box, 'Not Approved', fb);
    else [...box.querySelectorAll('input[type=radio]')].forEach((r) => { r.checked = r.value === d.resub[box.dataset.ref]?.decision; });
  };
  el.addEventListener('click', async (e) => {
    const box = e.target.closest('[data-fb]')?.closest('[data-ref]');
    if (!box) return;
    const fb = await askFeedback(d.resub[box.dataset.ref]?.feedback || '');
    if (fb) await decide(box, 'Not Approved', fb);
  });
}

// Feedback for a Not Approved project (v2.10): required. Resolves the text, or '' when cancelled.
async function askFeedback(prev) {
  let text = prev;
  for (;;) {
    const p = dialog(tr('Why is this project Not Approved?', 'เหตุผลที่ไม่อนุมัติโครงการนี้'),
      `<div class="field"><label for="nap-fb">${esc(tr('Feedback', 'ข้อเสนอแนะ'))}</label><textarea id="nap-fb" maxlength="2000">${esc(text)}</textarea></div>`,
      [{ v: false, label: tr('Cancel', 'ยกเลิก') }, { v: true, label: tr('Save', 'บันทึก'), cls: 'primary' }], { none: false });
    const box = document.getElementById('nap-fb');
    box?.addEventListener('input', () => { text = box.value; });
    box?.focus();
    if (!await p) return '';
    if (text.trim()) return text.trim();
    toast(tr('Write feedback for a Not Approved project', 'กรุณาเขียนข้อเสนอแนะสำหรับโครงการที่ไม่อนุมัติ'), 'bad');
  }
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
