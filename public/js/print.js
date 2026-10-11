// Print (v2.6): one organization (Print on its page) or every organization of a campus (Print campus on the list).
// Per organization: a cover (master data, totals, project index), then one page per phase — Pre-pitch, Pitching,
// Re-submission — each listing every project, instead of one page per project. Styled like the site (teal bands,
// cards, pills) in light colours whatever the theme. Every document is a link to its Drive file, so Save as PDF
// keeps them clickable. Uses only data the page already has (the org call); no extra request for one organization.
import { tr, getLang } from './i18n.js';
import { esc, num2, dateTimeFmt, phoneFmt, dialog } from './ui.js';
import { viewUrl } from './api.js';
import { CAMPUSES, CATEGORIES, ORGS } from './masterdata.js';
import { projNo, slotText, phaseOn, PENDING } from './project.js';
import { NO_DATA, RESUB_FIELDS } from './views/resubmit.js';

export const PHASES = [
  ['prepitch', () => tr('Pre-pitch', 'ก่อนการนำเสนอ')],
  ['pitching', () => tr('Pitching', 'การนำเสนอ')],
  ['resubmit', () => tr('Re-submission', 'การส่งแก้ไข')],
];

// Print options window: the cover always prints; each phase can be left out. Resolves { prepitch, pitching,
// resubmit } or null (Cancel).
export async function askPhases(title) {
  const pick = { prepitch: true, pitching: true, resubmit: true };
  const onChange = (e) => { const k = e.target.dataset?.pp; if (k) pick[k] = e.target.checked; };
  document.addEventListener('change', onChange);
  const html = `<div class="print-opts"><label class="check"><input type="checkbox" checked disabled>${tr('Cover and master data', 'หน้าปกและข้อมูลหลัก')}</label>
    ${PHASES.map(([k, label]) => `<label class="check"><input type="checkbox" data-pp="${k}" checked>${esc(label())}</label>`).join('')}</div>`;
  const ok = await dialog(title, html, [{ v: false, label: tr('Cancel', 'ยกเลิก') }, { v: true, label: tr('Print', 'พิมพ์'), cls: 'primary' }], { none: false });
  document.removeEventListener('change', onChange);
  return ok ? pick : null;
}

const name = (list, code) => { const x = list.find((i) => i.code === code); return x ? (getLang() === 'th' ? x.th : x.en) : code || ''; };
const money = (v) => (v == null || v === '' ? '—' : `฿${num2(v)}`);
const sum = (vals) => { const n = vals.filter((v) => v != null && v !== '' && Number.isFinite(+v)); return n.length ? n.reduce((t, v) => t + +v, 0) : null; };
const text = (v) => (String(v || '').trim() ? esc(v).replace(/\n/g, '<br>') : '<span class="p-muted">—</span>');
const pill = (cls, t) => `<span class="p-pill ${cls}">${esc(t)}</span>`;
const yn = (v) => (v === 'Yes' ? pill('ok', tr('Yes', 'ใช่')) : v === 'No' ? pill('bad', tr('No', 'ไม่')) : pill('none', '—'));
// A document: its name, linked to the file in Drive.
const link = (src, label) => { const u = viewUrl(src); return u ? `<a class="p-link" href="${esc(u)}">${esc(label)}</a>` : ''; };
const files = (list) => (list || []).filter((f) => f?.id).map((f) => link(f, f.name || f.id)).join('<br>') || '<span class="p-muted">—</span>';
const kv = (rows) => `<dl class="p-kv">${rows.map(([k, v]) => `<dt>${esc(k)}</dt><dd>${v}</dd>`).join('')}</dl>`;
const table = (head, rows, num = []) => `<table class="p-table"><thead><tr>${head.map((h, i) => `<th class="${num.includes(i) ? 'r' : ''}">${esc(h)}</th>`).join('')}</tr></thead>
  <tbody>${rows.map((r) => `<tr>${r.map((c, i) => `<td class="${num.includes(i) ? 'r num' : ''}">${c}</td>`).join('')}</tr>`).join('')}</tbody></table>`;
const proj = (p) => `<b>${esc(p.name || '—')}</b><small>${esc(p.ref)}</small>`;

// One organization: cover + the chosen phases. d = the org call's answer; nameOf = display name for an email.
export function printSheet(ref, d, show = { prepitch: true, pitching: true, resubmit: true }, nameOf = (e) => e) {
  const o = ORGS.find((x) => x.code === ref);
  const s = d.org;
  const title = o?.en || ref;
  const contact = (n, p, ph) => [n, p, phoneFmt(ph)].filter(Boolean).map(esc).join(' · ') || '—';
  const day = (iso) => new Date(`${iso}T00:00`).toLocaleDateString(getLang() === 'th' ? 'th-TH' : 'en-GB', { day: 'numeric', month: 'short', year: 'numeric' });
  const dates = (p) => (p.dates || []).length ? `${day(p.dates[0])}${p.dates.length > 1 ? ` – ${day(p.dates.at(-1))} · ${tr('{n} days', '{n} วัน', { n: p.dates.length })}` : ''}` : '—';
  const ps = d.projects;
  const band = (label, n) => `<header class="p-band"><span class="p-step">${n}</span><h2>${esc(label)}</h2><span class="p-grow"></span><span class="p-org-tag">${esc(ref)} · ${esc(title)}</span></header>`;

  const cover = `<section class="p-cover">
    <header class="p-top"><div><b>Pitch feedback#2</b><span>TUSU Budget Pitching</span></div><span>${tr('Printed {when}', 'พิมพ์เมื่อ {when}', { when: dateTimeFmt(new Date().toISOString()) })}</span></header>
    <div class="p-title"><h1>${esc(title)}</h1>
      <div class="p-tags">${pill('ref', ref)}${d.closed ? pill('closed', tr('Closed', 'ปิดแล้ว')) : pill('ok', tr('Open', 'เปิด'))}</div>
      <p class="p-sub">${esc(name(CAMPUSES, o?.campus))} · ${esc(name(CATEGORIES, o?.category))}</p></div>
    <div class="p-kpis">
      <div><span>${tr('Total proposed budget', 'งบประมาณที่เสนอทั้งหมด')}</span><b class="num">${money(sum(ps.map((p) => p.budget)))}</b></div>
      <div><span>${tr('Estimated Approved budget', 'งบประมาณที่คาดว่าจะได้รับอนุมัติ')}</span><b class="num">${money(sum(ps.map((p) => d.pre[p.ref]?.approvedBudget)))}</b></div>
      <div><span>${tr('Final approved budget', 'งบประมาณอนุมัติขั้นสุดท้าย')}</span><b class="num">${money(sum(ps.map((p) => d.pitch[p.ref]?.finalBudget)))}</b></div>
    </div>
    <div class="p-card"><h3>${tr('Master data', 'ข้อมูลหลัก')}</h3>${kv([
    [tr('Submitter', 'ผู้ส่ง'), s ? contact(s.name, s.position, s.phone) : '—'],
    [tr('Emergency contact', 'ผู้ติดต่อกรณีฉุกเฉิน'), s ? contact(s.emName, s.emPosition, s.emPhone) : '—'],
    [tr('Latest submission', 'ส่งล่าสุด'), s ? `${dateTimeFmt(s.at)} · ${tr('version {n}', 'ครั้งที่ {n}', { n: s.version })}` : tr('Not submitted yet', 'ยังไม่ได้ส่ง')],
    [tr('Budget proposal form (.xlsx)', 'แบบฟอร์มรายการงบประมาณ (.xlsx)'), link(s?.budgetForm, tr('Budget proposal form', 'แบบฟอร์มงบประมาณ')) || '—'],
  ])}</div>
    <div class="p-card"><h3>${tr('Projects', 'โครงการ')}</h3>${ps.length ? table(
    ['#', tr('Project', 'โครงการ'), tr('Project dates', 'วันที่จัดโครงการ'), tr('Proposed', 'เสนอ')],
    ps.map((p) => [projNo(p.ref), proj(p), dates(p), money(p.budget)]), [3],
  ) : `<p class="p-muted">${tr('This organization has no submitted projects yet.', 'องค์กรนี้ยังไม่มีโครงการที่ส่งเข้ามา')}</p>`}</div>
  </section>`;

  const off = (k) => !phaseOn(d, k);
  const pending = `<p class="p-pending">${pill('none', PENDING())} ${tr('This phase has not started yet.', 'ช่วงนี้ยังไม่เริ่ม')}</p>`;

  const pre = () => {
    if (off('prepitch')) return pending;
    const notes = ps.filter((p) => { const x = d.pre[p.ref] || {}; return String(x.note || '').trim() || x.docxFile || x.pdfFile; });
    return `${table(
      ['#', tr('Project', 'โครงการ'), tr('Proposed', 'เสนอ'), tr('Estimated Approved budget', 'งบประมาณที่คาดว่าจะได้รับอนุมัติ'), '.docx', 'PDF', tr('Proposal files', 'ไฟล์โครงการ')],
      ps.map((p) => { const x = d.pre[p.ref] || {}; return [projNo(p.ref), proj(p), money(p.budget), money(x.approvedBudget), yn(x.docxOk), yn(x.pdfOk), [link(p.pdf, 'PDF'), link(p.docx, '.docx')].filter(Boolean).join(' · ') || '—']; }),
      [2, 3],
    )}${notes.map((p) => { const x = d.pre[p.ref] || {}; return `<div class="p-card"><h3>${esc(tr('Project {n}', 'โครงการที่ {n}', { n: projNo(p.ref) }))} · ${esc(p.name || '—')}</h3>${kv([
      [tr('Note', 'บันทึก'), text(x.note)],
      [tr('Feedback .docx', 'ไฟล์ข้อเสนอแนะ .docx'), files([x.docxFile])],
      [tr('Feedback PDF', 'ไฟล์ข้อเสนอแนะ PDF'), files([x.pdfFile])],
    ])}</div>`; }).join('')}`;
  };

  const pitch = () => {
    if (off('pitching')) return pending;
    const more = ps.filter((p) => { const x = d.pitch[p.ref] || {}; return String(x.summary || '').trim() || (x.noted || []).length || x.excel; });
    return `${table(
      ['#', tr('Project', 'โครงการ'), tr('Pitching time', 'เวลานำเสนอ'), tr('Final approved budget', 'งบประมาณอนุมัติขั้นสุดท้าย')],
      ps.map((p) => [projNo(p.ref), proj(p), esc(slotText(d.slots?.[p.ref])), money(d.pitch[p.ref]?.finalBudget)]),
      [3],
    )}${more.map((p) => { const x = d.pitch[p.ref] || {}; return `<div class="p-card"><h3>${esc(tr('Project {n}', 'โครงการที่ {n}', { n: projNo(p.ref) }))} · ${esc(p.name || '—')}</h3>${kv([
      [tr('Summary', 'สรุป'), text(x.summary)],
      [tr('Noted documents', 'เอกสารบันทึก'), files(x.noted)],
      [tr('Correct Excel file', 'ไฟล์ Excel ฉบับแก้ไข'), files([x.excel])],
    ])}</div>`; }).join('')}`;
  };

  const resub = () => {
    if (off('resubmit')) return pending;
    const dec = (v) => (v === 'Approved' ? pill('ok', tr('Approved', 'อนุมัติ')) : v === 'Not Approved' ? pill('bad', tr('Not Approved', 'ไม่อนุมัติ')) : pill('none', tr('No decision yet', 'ยังไม่ได้ตัดสิน')));
    const by = (x) => (x?.at ? `${dateTimeFmt(x.at)}<small>${esc(nameOf(x.by || ''))}</small>` : '—');
    return `<p class="p-case">${tr('Case', 'เคส')}: ${d.closed ? `${pill('closed', tr('Closed', 'ปิดแล้ว'))} ${esc(tr('Closed by {who} on {when}', 'ปิดโดย {who} เมื่อ {when}', { who: nameOf(d.closed.by), when: dateTimeFmt(d.closed.at) }))}` : pill('ok', tr('Open', 'เปิด'))}</p>
      <p class="p-muted">${esc(RESUB_FIELDS().join(', '))}: ${esc(NO_DATA())}</p>
      ${table(['#', tr('Project', 'โครงการ'), tr('Decision', 'ผลการพิจารณา'), tr('Decided', 'ตัดสินเมื่อ')], ps.map((p) => [projNo(p.ref), proj(p), dec(d.resub?.[p.ref]?.decision), by(d.resub?.[p.ref])]))}`;
  };

  const parts = [['prepitch', pre], ['pitching', pitch], ['resubmit', resub]].filter(([k]) => show[k] && ps.length);
  return `<article class="p-orgdoc">${cover}${parts.map(([k, fn], i) => `<section class="p-phase">${band(PHASES.find(([x]) => x === k)[1](), i + 1)}${fn()}</section>`).join('')}</article>`;
}

// Mounts the sheet, opens the browser's print window (which also offers Save as PDF), then removes it.
export function printNow(html) {
  const sheet = document.createElement('div');
  sheet.className = 'print-sheet';
  sheet.innerHTML = html;
  document.body.append(sheet);
  document.body.classList.add('printing');
  const done = () => { sheet.remove(); document.body.classList.remove('printing'); window.removeEventListener('afterprint', done); };
  window.addEventListener('afterprint', done);
  window.print();
}
