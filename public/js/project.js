// Pieces shared by the three detail tabs: the project card top, the read-only calendar, file links, the pitching
// slot and the progress banner (unfinished projects + the "unfinished only" filter).
import { tr, locale } from './i18n.js';
import { esc, icon, num2 } from './ui.js';
import { downloadUrl, viewUrl, previewUrl } from './api.js';

const pad = (n) => String(n).padStart(2, '0');
const near = (iso, n) => { const [y, m, d] = iso.split('-').map(Number); const t = new Date(y, m - 1, d + n); return `${t.getFullYear()}-${pad(t.getMonth() + 1)}-${pad(t.getDate())}`; };

// The pitching site's calendar, read-only: one month grid per month from the first to the last project day.
// A run of consecutive days is a tinted band with a solid circle at each end (classes as in its form.css):
// l / r = the day before / after is a project day too; lc / rc = the band breaks at the row or month edge.
export function calendar(dates) {
  const on = new Set((dates || []).filter((d) => /^\d{4}-\d{2}-\d{2}$/.test(d)));
  const ds = [...on].sort();
  if (!ds.length) return `<p class="muted small">${tr('No project dates', 'ไม่มีวันที่จัดโครงการ')}</p>`;
  const wk = Array.from({ length: 7 }, (_, d) => new Date(2024, 0, 1 + d).toLocaleDateString(locale(), { weekday: 'narrow' }));
  const [y0, m0] = ds[0].split('-').map(Number);
  const [y1, m1] = ds.at(-1).split('-').map(Number);
  const months = [];
  for (let y = y0, m = m0; y < y1 || (y === y1 && m <= m1); m === 12 ? (y++, m = 1) : m++) {
    const lead = (new Date(y, m - 1, 1).getDay() + 6) % 7; // Monday first
    const last = new Date(y, m, 0).getDate();
    const cells = Array.from({ length: lead }, () => '<span></span>');
    for (let d = 1; d <= last; d++) {
      const iso = `${y}-${pad(m)}-${pad(d)}`;
      const wd = new Date(y, m - 1, d).getDay();
      const sel = on.has(iso);
      const L = sel && on.has(near(iso, -1));
      const R = sel && on.has(near(iso, 1));
      const cls = ['day', sel && 'on', L && 'l', R && 'r', L && (wd === 1 || d === 1) && 'lc', R && (wd === 0 || d === last) && 'rc'].filter(Boolean).join(' ');
      cells.push(`<span class="${cls}">${d}</span>`);
    }
    months.push(`<div class="month"><div class="mname">${esc(new Date(y, m - 1, 1).toLocaleDateString(locale(), { month: 'long', year: 'numeric' }))}</div>
      <div class="days">${wk.map((w) => `<span class="wd">${esc(w)}</span>`).join('')}${cells.join('')}</div></div>`);
  }
  return `<div class="cal">${months.join('')}</div><p class="cal-sum"><b>${(ds.length > 1 ? tr('{n} days', '{n} วัน', { n: ds.length }) : tr('{n} day', '{n} วัน', { n: ds.length }))}</b></p>`;
}

// File button for a Drive link or a { id } file: the label opens a preview, the arrow downloads at once.
// Greyed out when there is no file.
export const dlBtn = (src, label) => {
  const u = downloadUrl(src);
  if (!u) return `<span class="btn split" aria-disabled="true">${icon('eye', 16)}${esc(label)}</span>`;
  return `<span class="split"><button type="button" class="btn" data-preview="${esc(previewUrl(src))}" data-view="${esc(viewUrl(src))}" data-dl="${esc(u)}" data-title="${esc(label)}">${icon('eye', 16)}${esc(label)}</button><a class="btn" href="${esc(u)}" download rel="noopener" title="${esc(tr('Download', 'ดาวน์โหลด'))}" aria-label="${esc(tr('Download {file}', 'ดาวน์โหลด {file}', { file: label }))}">${icon('download', 16)}</a></span>`;
};

// Preview window: Drive's viewer in an iframe, with Open in Drive, Download and Close. Esc or a click outside closes.
export function openPreview({ preview, view, dl, title }) {
  const back = document.createElement('div');
  back.className = 'modal-back';
  back.innerHTML = `<div class="viewer" role="dialog" aria-modal="true" aria-label="${esc(title)}">
    <header><b class="grow">${esc(title)}</b>
      <a class="btn plain" href="${esc(view)}" target="_blank" rel="noopener">${tr('Open in Drive', 'เปิดใน Drive')}</a>
      <a class="btn" href="${esc(dl)}" download rel="noopener">${icon('download', 16)}${tr('Download', 'ดาวน์โหลด')}</a>
      <button type="button" class="icon-btn" data-close aria-label="${esc(tr('Close', 'ปิด'))}">${icon('x', 20)}</button></header>
    <iframe src="${esc(preview)}" title="${esc(title)}" allow="autoplay"></iframe>
    <p class="muted small">${tr('Nothing showing? You may need access to this file in Google Drive — use Open in Drive.', 'ไม่แสดงไฟล์? คุณอาจต้องมีสิทธิ์เข้าถึงไฟล์นี้ใน Google Drive — กดเปิดใน Drive')}</p></div>`;
  const close = () => { back.remove(); document.removeEventListener('keydown', key); };
  const key = (e) => { if (e.key === 'Escape') close(); };
  back.addEventListener('click', (e) => { if (e.target === back || e.target.closest('[data-close]')) close(); });
  document.addEventListener('keydown', key);
  document.body.append(back);
  back.querySelector('[data-close]').focus();
}
// One listener for every preview button and file name on the page.
document.addEventListener('click', (e) => {
  const b = e.target.closest('[data-preview]');
  if (!b) return;
  e.preventDefault();
  openPreview({ preview: b.dataset.preview, view: b.dataset.view, dl: b.dataset.dl, title: b.dataset.title });
});

// A saved file: name (opens the preview) and a download link.
export const fileChip = (f) => (f?.id ? `<span class="fchip">${icon('file', 16)}<a href="${esc(viewUrl(f))}" data-preview="${esc(previewUrl(f))}" data-view="${esc(viewUrl(f))}" data-dl="${esc(downloadUrl(f))}" data-title="${esc(f.name || f.id)}" title="${esc(tr('Preview', 'ดูตัวอย่าง'))}">${esc(f.name || f.id)}</a>
  <a class="icon-btn" href="${esc(downloadUrl(f))}" download title="${esc(tr('Download', 'ดาวน์โหลด'))}">${icon('download', 16)}</a></span>` : '');

// Number = the project's position in the submission (the -P<nn> end of its code).
export const projNo = (ref) => Number(String(ref).match(/-P(\d+)$/)?.[1]) || '';

// Feedback questions the Owner can mark Required / Not required (Control tab). d.required (from the org call)
// holds the Owner's choices; a question not in it keeps its default. Keep in step with QUESTIONS in Code.gs.
export const QUESTIONS = [
  { id: 'pre.approvedBudget', tab: 'prepitch', req: true, label: () => tr('Approved budget', 'งบประมาณที่อนุมัติ') },
  { id: 'pre.docx', tab: 'prepitch', req: true, label: () => tr('Approved project .docx? (feedback file when No)', 'อนุมัติไฟล์โครงการ .docx? (ไฟล์ข้อเสนอแนะเมื่อไม่อนุมัติ)') },
  { id: 'pre.pdf', tab: 'prepitch', req: true, label: () => tr('Approved project PDF? (feedback file when No)', 'อนุมัติไฟล์โครงการ PDF? (ไฟล์ข้อเสนอแนะเมื่อไม่อนุมัติ)') },
  { id: 'pre.note', tab: 'prepitch', req: false, label: () => tr('Feedback note', 'บันทึกข้อเสนอแนะ') },
  { id: 'pitch.summary', tab: 'pitching', req: false, label: () => tr('Summary', 'สรุป') },
  { id: 'pitch.finalBudget', tab: 'pitching', req: true, label: () => tr('Final approved budget', 'งบประมาณอนุมัติขั้นสุดท้าย') },
  { id: 'pitch.noted', tab: 'pitching', req: false, label: () => tr('Noted documents', 'เอกสารบันทึก') },
  { id: 'resub.decision', tab: 'resubmit', req: true, label: () => tr('Decision', 'ผลการพิจารณา') },
];
export const isReq = (d, id) => (d?.required && id in d.required ? !!d.required[id] : !!QUESTIONS.find((q) => q.id === id)?.req);
// Pitching phases (admin console): a phase is on unless switched off. Off = its tab shows "Pending Phase".
export const phaseOn = (d, k) => d?.phases?.[k] !== false;
export const PENDING = () => tr('Pending Phase', 'รอเปิดช่วง');
export const star = (d, id) => (isReq(d, id) ? `<span class="req" title="${esc(tr('Required', 'จำเป็น'))}">*</span>` : '');

// Card top: name + ref, lead (e.g. the pitching slot), calendar, budget line(s), extra (download buttons).
export function cardTop(p, budgets, extra = '', lead = '') {
  const n = projNo(p.ref);
  return `<header class="pcard-head"><span class="eyebrow">${esc(tr('Project {n}', 'โครงการที่ {n}', { n }))}</span><h3>${esc(p.name || '—')}</h3><span class="code">${esc(p.ref)}</span></header>
    <div class="pcard-body">${lead}${calendar(p.dates)}
      <dl class="budgets">${budgets.map(([k, v]) => `<div><dt>${esc(k)}</dt><dd class="num">${v == null || v === '' ? '—' : num2(v)}</dd></div>`).join('')}</dl>
      ${extra ? `<div class="dl-row">${extra}</div>` : ''}</div>`;
}

// Pitching slot from the scheduling site: { date: 'yyyy-mm-dd', start, end, room } (room may be a meeting link).
export function slotLine(s) {
  if (!s?.date) return `<p class="slot none">${icon('clock', 18)}${tr('Pitching time not scheduled yet', 'ยังไม่มีกำหนดเวลานำเสนอ')}</p>`;
  const d = new Date(`${s.date}T00:00`);
  const day = Number.isNaN(+d) ? s.date : d.toLocaleDateString(locale(), { weekday: 'short', day: 'numeric', month: 'short', year: 'numeric' });
  const room = /^https?:\/\//.test(s.room) ? `<a href="${esc(s.room)}" target="_blank" rel="noopener">${tr('Join meeting', 'เข้าร่วมประชุม')}</a>` : esc(s.room);
  return `<p class="slot">${icon('clock', 18)}<span class="muted">${tr('Pitching', 'นำเสนอ')}</span><b>${esc(day)}</b><b class="num">${esc(s.start)}${s.end ? `–${esc(s.end)}` : ''}</b>${room ? `<span>${room}</span>` : ''}</p>`;
}

// The same slot as plain text (print).
export function slotText(s) {
  if (!s?.date) return tr('Pitching time not scheduled yet', 'ยังไม่มีกำหนดเวลานำเสนอ');
  const d = new Date(`${s.date}T00:00`);
  const day = Number.isNaN(+d) ? s.date : d.toLocaleDateString(locale(), { weekday: 'short', day: 'numeric', month: 'short', year: 'numeric' });
  return [day, `${s.start || ''}${s.end ? `–${s.end}` : ''}`, s.room].filter(Boolean).join(' · ');
}

// What "unfinished" means on each tab: a required question unanswered. Pre-pitch: decided by the server (Code.gs finished_).
export const UNFINISHED = {
  prepitch: (d, p) => !d.pre[p.ref]?.finished,
  pitching: (d, p) => {
    const x = d.pitch[p.ref] || {};
    return (isReq(d, 'pitch.summary') && !String(x.summary || '').trim()) || (isReq(d, 'pitch.finalBudget') && x.finalBudget == null)
      || (isReq(d, 'pitch.noted') && !(x.noted || []).length);
  },
  resubmit: (d, p) => isReq(d, 'resub.decision') && !d.resub?.[p.ref]?.decision,
};
const DONE = {
  prepitch: () => tr('All approvals finished', 'อนุมัติครบทุกโครงการ'),
  pitching: () => tr('Every project has its required answers', 'ทุกโครงการตอบคำถามที่จำเป็นครบแล้ว'),
  resubmit: () => tr('Every project has a decision', 'ทุกโครงการตัดสินแล้ว'),
};
let onlyOpen = false; // "Unfinished only": kept while moving between tabs and organizations

// Banner at the top of a tab (an element with data-progress inside el): unfinished count, meter, a link to each
// unfinished project, and the filter that hides finished cards. Call again after every save.
export function progress(el, d, tab) {
  const host = el.querySelector('[data-progress]');
  const open = d.projects.filter((p) => UNFINISHED[tab](d, p));
  const n = open.length; const all = d.projects.length;
  host.className = `progress ${n ? 'open' : 'done'}`;
  host.refresh = () => progress(el, d, tab);
  host.innerHTML = `<div class="pg-main">${icon(n ? 'alert' : 'check', 28)}
      <div class="grow"><b>${n ? tr('{n} of {all} projects unfinished', 'ยังไม่ครบ {n} จาก {all} โครงการ', { n, all }) : esc(DONE[tab]())}</b>
        <div class="meter" role="meter" aria-valuemin="0" aria-valuemax="${all}" aria-valuenow="${all - n}" aria-label="${esc(tr('Finished projects', 'โครงการที่ครบแล้ว'))}"><i style="width:${all ? ((all - n) / all) * 100 : 100}%"></i></div></div>
      ${n ? `<label class="check"><input type="checkbox" data-only ${onlyOpen ? 'checked' : ''}>${tr('Unfinished only', 'เฉพาะที่ยังไม่ครบ')}</label>` : ''}</div>
    ${n ? `<div class="pg-jump">${open.map((p) => `<a href="#" data-jump="${esc(p.ref)}" title="${esc(p.name || '')}">${esc(tr('Project {n}', 'โครงการที่ {n}', { n: projNo(p.ref) }))}</a>`).join('')}</div>` : ''}`;
  const keep = new Set(open.map((p) => p.ref));
  el.querySelectorAll('.pcard[data-ref]').forEach((c) => { c.hidden = onlyOpen && n > 0 && !keep.has(c.dataset.ref); });
}
document.addEventListener('change', (e) => {
  if (!e.target.matches('[data-only]')) return;
  onlyOpen = e.target.checked;
  e.target.closest('[data-progress]').refresh();
});
document.addEventListener('click', (e) => {
  const j = e.target.closest('[data-jump]');
  if (!j) return;
  e.preventDefault();
  document.querySelector(`.pcard[data-ref="${CSS.escape(j.dataset.jump)}"]`)?.scrollIntoView({ behavior: 'smooth', block: 'start' });
});
