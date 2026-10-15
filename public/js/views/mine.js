// My projects (v2.7): every project of the organizations I'm in charge of (the same rule as "Mine only" on the list),
// one row per project, grouped by organization. Shows where each project stands — Pitching progress (required
// answers, as on the Pitching tab), Re-submission decision — and when and where it pitches. No person in charge
// column: they are all mine. One `myProjects` call (the org data of each organization, cached per organization).
import { tr } from '../i18n.js';
import { $, esc, icon, num2, toast } from '../ui.js';
import { cached, forget, peek } from '../api.js';
import { ORGS } from '../masterdata.js';
import { projNo, slotText, meetLink, timeLeft, UNFINISHED, phaseOn, PENDING } from '../project.js';
import { statusPill } from './org.js';

export async function render(page, ctx) {
  page.innerHTML = `<h1 class="grad">${tr('My projects', 'โครงการของฉัน')}</h1><p class="muted" id="sum">&nbsp;</p><div id="rows" class="org-list"></div>`;
  const kept = peek('myProjects');
  if (kept) paint(page, ctx, kept);
  forget('myProjects'); // always ask again: saves on organization pages change it
  try {
    const fresh = await cached('myProjects');
    if (!kept || JSON.stringify(fresh) !== JSON.stringify(kept)) paint(page, ctx, fresh);
  } catch (e) {
    if (!kept) throw e;
    toast(e.message, 'bad');
  }
}

const COLS = '<colgroup><col style="width:5%"><col style="width:24%"><col style="width:12%"><col style="width:12%"><col style="width:13%"><col style="width:17%"><col style="width:8%"><col style="width:9%"></colgroup>';
let tick = 0; // the Time left column refreshes every minute while the page is open

const left = (slot) => timeLeft(slot) || '<span class="muted">—</span>';

function paint(page, ctx, out) {
  const orgs = out.orgs || [];
  const projects = orgs.reduce((n, o) => n + o.data.projects.length, 0);
  $('#sum', page).textContent = orgs.length
    ? tr('{o} organizations · {p} projects', '{o} องค์กร · {p} โครงการ', { o: orgs.length, p: projects })
    : '';
  if (!orgs.length) {
    $('#rows', page).innerHTML = `<div class="empty">${icon('user', 28)}<p>${tr('You are not in charge of any organization with a submission yet.', 'คุณยังไม่ได้รับผิดชอบองค์กรที่ส่งข้อมูลแล้ว')}</p></div>`;
    return;
  }
  const head = `<tr><th>#</th><th>${tr('Project', 'โครงการ')}</th><th class="r">${tr('Proposed', 'เสนอ')}</th>
    <th>${tr('Pitching progress', 'ความคืบหน้าการนำเสนอ')}</th><th>${tr('Re-submission progress', 'ความคืบหน้าการส่งแก้ไข')}</th>
    <th>${tr('Pitching time', 'เวลานำเสนอ')}</th><th>${tr('Time left', 'เหลือเวลา')}</th><th>${tr('Meeting link', 'ลิงก์ประชุม')}</th></tr>`;
  $('#rows', page).innerHTML = orgs.map(({ ref, data: d }) => {
    const o = ORGS.find((x) => x.code === ref);
    const top = `<tr class="grp"><th colspan="8"><a href="#/org/${encodeURIComponent(ref)}"><span class="mono">${esc(ref)}</span> · ${esc(o?.en || ref)}</a> ${statusPill(d.closed)}</th></tr>`;
    const rows = d.projects.map((p) => {
      const pitch = !phaseOn(d, 'pitching') ? `<span class="pill">${esc(PENDING())}</span>`
        : UNFINISHED.pitching(d, p) ? `<span class="pill orange">${tr('Unfinished', 'ยังไม่ครบ')}</span>` : `<span class="pill green">${tr('Done', 'ครบ')}</span>`;
      const dec = d.resub?.[p.ref]?.decision;
      const resub = !phaseOn(d, 'resubmit') ? `<span class="pill">${esc(PENDING())}</span>`
        : dec === 'Approved' ? `<span class="pill green">${tr('Approved', 'อนุมัติ')}</span>`
          : dec === 'Not Approved' ? `<span class="pill red">${tr('Not Approved', 'ไม่อนุมัติ')}</span>` : `<span class="muted">${tr('No decision yet', 'ยังไม่ได้ตัดสิน')}</span>`;
      const slot = d.slots?.[p.ref];
      return `<tr data-ref="${esc(ref)}" data-tab="pitching" tabindex="0"><td>${projNo(p.ref)}</td><td>${esc(p.name || '—')}<small class="muted mono block">${esc(p.ref)}</small></td>
        <td class="r num">${p.budget == null ? '—' : num2(p.budget)}</td><td>${pitch}</td><td>${resub}</td>
        <td class="small">${slot?.date ? esc(slotText(slot)) : '<span class="muted">—</span>'}</td><td class="small" data-left="${esc(JSON.stringify(slot || {}))}">${left(slot)}</td>
        <td class="small">${meetLink(slot, '<span class="muted">—</span>')}</td></tr>`;
    }).join('') || `<tr class="none"><td colspan="8" class="muted">${tr('No submitted projects', 'ไม่มีโครงการที่ส่งเข้ามา')}</td></tr>`;
    return `<div class="group scroll-x campus-card"><table class="grid-table quotes">${COLS}<thead>${top}${head}</thead><tbody>${rows}</tbody></table></div>`;
  }).join('');
  clearInterval(tick);
  tick = setInterval(() => {
    if (!page.isConnected) { clearInterval(tick); return; }
    page.querySelectorAll('[data-left]').forEach((td) => { td.innerHTML = left(JSON.parse(td.dataset.left)); });
  }, 60e3);
  // A row opens its organization's Pitching tab (links inside the row work as links).
  $('#rows', page).onclick = (e) => {
    if (e.target.closest('a')) return;
    const r = e.target.closest('tr[data-ref]');
    if (r) ctx.go(`#/org/${encodeURIComponent(r.dataset.ref)}/pitching`);
  };
  $('#rows', page).onkeydown = (e) => { if (e.key === 'Enter' && e.target.matches('tr[data-ref]')) ctx.go(`#/org/${encodeURIComponent(e.target.dataset.ref)}/pitching`); };
}
