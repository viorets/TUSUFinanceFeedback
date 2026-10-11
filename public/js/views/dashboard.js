// Dashboard: budget at each stage against the allocated budget (กรอบงบประมาณ) — proposed, approved (Pre-pitch),
// pitched (Pitching final approved budget), final (Re-submission: pitched budget of projects Approved there) — project
// count, the top 5 projects and organizations by proposed budget, proposed budget by campus and by category (pies +
// legend tables), and a table of every project (search, group by, sort by).
// Data: one `dashboard` call (every project of the organizations on the list). Amounts shown as ฿1,234,567.
import { tr, getLang } from '../i18n.js';
import { $, esc, icon, toast } from '../ui.js';
import { cached, peek } from '../api.js';
import { CAMPUSES, CATEGORIES, ORGS } from '../masterdata.js';

// Categorical slots: CSS tokens --c0..--c5, set for the dashboard by .dash-ci in app.css from the Mosaic Blue
// palette (#00758F, the dashboard CI). Colour follows the entity: campus / category N always gets slot N.
const SLOTS = 6;

const baht = (n) => `฿${Math.round(n || 0).toLocaleString('en-US')}`;
const pct = (a, b) => (b ? Math.round((a / b) * 1000) / 10 : 0);
const label = (x) => (x ? (getLang() === 'th' ? x.th : x.en) : '—');
const sumOf = (list, k) => list.reduce((a, p) => a + (p[k] || 0), 0);
// Final budget: the pitched budget, once the re-submission is Approved.
const finalOf = (p) => (p.decision === 'Approved' ? p.pitched : null);

// Project table state (kept while moving around the site).
const view = { q: '', group: 'none', sort: 'budget' };

export async function render(page) {
  page.innerHTML = `<div class="dash-ci"><h1 class="grad">${tr('Dashboard', 'แดชบอร์ด')}</h1><div id="dash"></div></div>`;
  const kept = peek('dashboard');
  if (kept) draw($('#dash', page), kept);
  try {
    const fresh = await cached('dashboard');
    if (!kept || JSON.stringify(fresh) !== JSON.stringify(kept)) draw($('#dash', page), fresh);
  } catch (e) {
    if (!kept) throw e;
    toast(e.message, 'bad');
  }
}

function draw(el, d) {
  const ps = d.projects.map((p) => ({ ...p, final: finalOf(p) }));
  const total = sumOf(ps, 'budget');
  const has = (k) => ps.filter((p) => p[k] != null);
  const of = (k) => tr('{n} of {all} projects', '{n} จาก {all} โครงการ', { n: has(k).length, all: ps.length });
  const orgsWith = new Set(ps.map((p) => p.org)).size;

  const byOrg = {};
  ps.forEach((p) => { byOrg[p.org] = (byOrg[p.org] || 0) + p.budget; });
  const topOrgs = Object.entries(byOrg).sort((a, b) => b[1] - a[1]).slice(0, 5);
  const topProjects = [...ps].sort((a, b) => b.budget - a.budget).slice(0, 5);
  const orgName = (ref) => ORGS.find((o) => o.code === ref)?.en || ref;

  // Org code = <campus>-<category>-<nn>, so both come from the code (works for every project).
  const sum = (list, part) => list.map((x, i) => ({ x, color: `var(--c${i % SLOTS})`, v: ps.filter((p) => p.org.split('-')[part] === x.code).reduce((a, p) => a + p.budget, 0) }));

  el.innerHTML = `<div class="kpis">
      ${tile(tr('Total proposed budget', 'งบประมาณที่เสนอทั้งหมด'), total, d.allocated, tr('{n} projects', '{n} โครงการ', { n: ps.length }))}
      ${tile(tr('Estimated Approved budget', 'งบประมาณที่คาดว่าจะได้รับอนุมัติ'), sumOf(ps, 'approved'), d.allocated, `${tr('Pre-pitch', 'ก่อนการนำเสนอ')} · ${of('approved')}`)}
      ${tile(tr('Pitched budget', 'งบประมาณหลังการนำเสนอ'), sumOf(ps, 'pitched'), d.allocated, `${tr('Final approved budget at pitching', 'งบประมาณอนุมัติขั้นสุดท้ายจากการนำเสนอ')} · ${of('pitched')}`)}
      ${tile(tr('Final budget', 'งบประมาณสุดท้าย'), sumOf(ps, 'final'), d.allocated, `${tr('Re-submission approved', 'อนุมัติหลังส่งแก้ไข')} · ${of('final')}`)}
      <div class="kpi"><span>${tr('Total projects', 'จำนวนโครงการทั้งหมด')}</span><b class="hero">${ps.length.toLocaleString('en-US')}</b>
        <p class="muted small">${tr('from {n} organizations', 'จาก {n} องค์กร', { n: orgsWith })}</p></div>
    </div>
    ${d.allocated == null ? `<p class="foot">${icon('alert', 14)} ${tr('Allocated budget is not set yet — an Admin or Chief can set it in the Admin console.', 'ยังไม่ได้กำหนดกรอบงบประมาณ — ผู้ดูแลหรือหัวหน้ากำหนดได้ในคอนโซลผู้ดูแล')}</p>` : ''}

    <div class="dash-grid">
      <section><h2 class="section-title">${tr('Top 5 projects by budget', '5 โครงการที่ใช้งบประมาณสูงสุด')}</h2>
        ${rank(topProjects.map((p) => ({ name: p.name, sub: `${p.ref} · ${orgName(p.org)}`, v: p.budget, href: `#/org/${encodeURIComponent(p.org)}` })))}</section>
      <section><h2 class="section-title">${tr('Top 5 organizations by budget', '5 องค์กรที่ใช้งบประมาณสูงสุด')}</h2>
        ${rank(topOrgs.map(([ref, v]) => ({ name: orgName(ref), sub: ref, v, href: `#/org/${encodeURIComponent(ref)}` })))}</section>
      <section><h2 class="section-title">${tr('Project budget by campus', 'งบประมาณโครงการตามวิทยาเขต')}</h2>${pie(sum(CAMPUSES, 0), total)}</section>
      <section><h2 class="section-title">${tr('Project budget by category', 'งบประมาณโครงการตามหมวด')}</h2>${pie(sum(CATEGORIES, 1), total)}</section>
    </div>

    <h2 class="section-title">${tr('All project budgets', 'งบประมาณทุกโครงการ')}</h2>
    <div class="toolbar">
      <input type="search" id="pq" value="${esc(view.q)}" placeholder="${esc(tr('Filter by project, organization or RefCode', 'กรองตามโครงการ องค์กร หรือรหัส'))}">
      <label class="ctl">${tr('Group by', 'จัดกลุ่มตาม')}<select id="pgroup">${opts(GROUPS, view.group)}</select></label>
      <label class="ctl">${tr('Sort by', 'เรียงตาม')}<select id="psort">${opts(SORTS, view.sort)}</select></label>
    </div>
    <div class="group scroll-x"><table class="grid-table budgets-table"><thead><tr>
      <th>${tr('Project', 'โครงการ')}</th><th>${tr('Organization', 'องค์กร')}</th>
      ${COLS.map(([, l]) => `<th class="r">${esc(l())}</th>`).join('')}</tr></thead><tbody id="ptable"></tbody></table></div>`;
  hover(el);
  const paint = () => { $('#ptable', el).innerHTML = table(ps); };
  paint();
  $('#pq', el).oninput = (e) => { view.q = e.target.value; paint(); };
  $('#pgroup', el).onchange = (e) => { view.group = e.target.value; paint(); };
  $('#psort', el).onchange = (e) => { view.sort = e.target.value; paint(); };
  const open = (e) => { const r = e.target.closest('tr[data-org]'); if (r) location.hash = `#/org/${encodeURIComponent(r.dataset.org)}`; };
  $('#ptable', el).onclick = open;
  $('#ptable', el).onkeydown = (e) => { if (e.key === 'Enter') open(e); };
}

// ---------- every-project table ----------
const orgName = (ref) => ORGS.find((o) => o.code === ref)?.en || ref;
const COLS = [
  ['budget', () => tr('Proposed', 'เสนอ')], ['approved', () => tr('Estimated approved', 'คาดว่าจะอนุมัติ')],
  ['pitched', () => tr('Pitched', 'หลังนำเสนอ')], ['final', () => tr('Final', 'สุดท้าย')],
];
// Group key + heading per choice. Codes are <campus>-<category>-<nn>.
const GROUPS = {
  none: { label: () => tr('Nothing', 'ไม่จัดกลุ่ม'), key: () => '' },
  campus: { label: () => tr('Campus', 'วิทยาเขต'), key: (p) => p.org.split('-')[0], head: (k) => label(CAMPUSES.find((c) => c.code === k)) },
  category: { label: () => tr('Category', 'หมวด'), key: (p) => p.org.split('-')[1], head: (k) => label(CATEGORIES.find((c) => c.code === k)) },
  org: { label: () => tr('Organization', 'องค์กร'), key: (p) => p.org, head: (k) => `${orgName(k)} · ${k}` },
};
// Amounts high to low (not set last); names and codes A to Z.
const byNum = (k) => (a, b) => (b[k] ?? -1) - (a[k] ?? -1);
const SORTS = {
  budget: { label: () => tr('Proposed budget', 'งบที่เสนอ'), fn: byNum('budget') },
  approved: { label: () => tr('Estimated Approved budget', 'งบประมาณที่คาดว่าจะได้รับอนุมัติ'), fn: byNum('approved') },
  pitched: { label: () => tr('Pitched budget', 'งบหลังนำเสนอ'), fn: byNum('pitched') },
  final: { label: () => tr('Final budget', 'งบสุดท้าย'), fn: byNum('final') },
  name: { label: () => tr('Project name', 'ชื่อโครงการ'), fn: (a, b) => String(a.name).localeCompare(String(b.name)) },
  ref: { label: () => 'RefCode', fn: (a, b) => a.ref.localeCompare(b.ref) },
};
const opts = (o, on) => Object.entries(o).map(([k, x]) => `<option value="${k}" ${k === on ? 'selected' : ''}>${esc(x.label())}</option>`).join('');
const money = (v) => (v == null ? '<span class="muted">—</span>' : baht(v));
const totals = (list, name, cls) => `<tr class="${cls}"><th colspan="2">${esc(name)} <span class="muted">· ${tr('{n} projects', '{n} โครงการ', { n: list.length })}</span></th>
  ${COLS.map(([k]) => `<th class="r num">${baht(sumOf(list, k))}</th>`).join('')}</tr>`;

function table(ps) {
  const q = view.q.trim().toLowerCase();
  const list = ps.filter((p) => !q || `${p.name} ${p.ref} ${p.org} ${orgName(p.org)}`.toLowerCase().includes(q)).sort(SORTS[view.sort].fn);
  if (!list.length) return `<tr><td colspan="6" class="muted">${tr('No projects match', 'ไม่มีโครงการที่ตรงกัน')}</td></tr>`;
  // Total first, so it reads before the rows and groups.
  const row = (p) => `<tr data-org="${esc(p.org)}" tabindex="0"><td><b>${esc(p.name || '—')}</b><small class="muted mono">${esc(p.ref)}</small></td>
    <td class="small">${esc(orgName(p.org))}</td>${COLS.map(([k]) => `<td class="r num">${money(p[k])}</td>`).join('')}</tr>`;
  const g = GROUPS[view.group];
  let body;
  if (view.group === 'none') body = list.map(row).join('');
  else {
    // Groups in order of their total proposed budget, high to low; rows inside keep the chosen sort.
    const groups = new Map();
    list.forEach((p) => { const k = g.key(p); if (!groups.has(k)) groups.set(k, []); groups.get(k).push(p); });
    body = [...groups].sort((a, b) => sumOf(b[1], 'budget') - sumOf(a[1], 'budget'))
      .map(([k, rows]) => totals(rows, g.head(k) || k, 'grp') + rows.map(row).join('')).join('');
  }
  return totals(list, tr('Total', 'รวม'), 'total') + body;
}

// Big number + a meter against the allocated budget (track and fill in one hue; over the allocation = warning).
function tile(name, v, alloc, note) {
  const has = alloc != null && alloc > 0;
  const p = has ? pct(v, alloc) : 0;
  const over = has && v > alloc;
  return `<div class="kpi"><span>${esc(name)}</span><b class="hero num">${baht(v)}</b>
    ${has ? `<div class="meter${over ? ' over' : ''}" role="meter" aria-valuemin="0" aria-valuemax="${alloc}" aria-valuenow="${v}" aria-label="${esc(name)}"><i style="width:${Math.min(100, p)}%"></i></div>
      <p class="small ${over ? 'warn' : 'muted'}">${over ? `${icon('alert', 13)} ` : ''}${tr('{p}% of {alloc} allocated', '{p}% ของกรอบงบประมาณ {alloc}', { p, alloc: baht(alloc) })}${over ? tr(' · {over} over', ' · เกิน {over}', { over: baht(v - alloc) }) : ''}</p>` : ''}
    <p class="muted small">${esc(note)}</p></div>`;
}

// Ranked list with a thin magnitude bar (one hue).
function rank(rows) {
  if (!rows.length) return `<div class="group empty">${tr('No projects yet', 'ยังไม่มีโครงการ')}</div>`;
  const max = rows[0].v || 1;
  return `<ol class="group rank">${rows.map((r, i) => `<li><a href="${esc(r.href)}"><span class="no">${i + 1}</span>
    <span class="grow"><b>${esc(r.name)}</b><small class="muted">${esc(r.sub)}</small><i class="bar" style="width:${Math.max(2, (r.v / max) * 100)}%"></i></span>
    <span class="num">${baht(r.v)}</span></a></li>`).join('')}</ol>`;
}

// Pie of the parts with budget, 2px surface gaps between slices, hover tooltip; the legend table lists every part.
function pie(parts, total) {
  const shown = parts.filter((p) => p.v > 0);
  if (!total || !shown.length) return `<div class="group empty">${tr('No budget yet', 'ยังไม่มีงบประมาณ')}</div>`;
  const R = 100; let a = -Math.PI / 2;
  const pt = (t) => `${(R + R * Math.cos(t)).toFixed(2)} ${(R + R * Math.sin(t)).toFixed(2)}`;
  const tip = (p) => `${label(p.x)}: ${baht(p.v)} (${pct(p.v, total)}%)`;
  const slices = shown.map((p) => {
    const da = (p.v / total) * Math.PI * 2;
    const path = shown.length === 1 ? `<circle cx="${R}" cy="${R}" r="${R}"` : `<path d="M${R} ${R} L${pt(a)} A${R} ${R} 0 ${da > Math.PI ? 1 : 0} 1 ${pt(a + da)} Z"`;
    a += da;
    return `${path} style="fill:${p.color}" data-tip="${esc(tip(p))}"/>`;
  }).join('');
  return `<div class="group pie"><svg viewBox="-2 -2 204 204" role="img" aria-label="${esc(shown.map(tip).join('; '))}">${slices}</svg>
    <table class="legend"><tbody>${parts.map((p) => `<tr class="${p.v ? '' : 'muted'}"><td><span class="sw" style="background:${p.color}"></span>${esc(label(p.x))}</td>
      <td class="num">${baht(p.v)}</td><td class="num">${pct(p.v, total)}%</td></tr>`).join('')}</tbody></table></div>`;
}

// One tooltip for every slice on the page.
function hover(el) {
  let t = document.getElementById('viz-tip');
  if (!t) { t = document.createElement('div'); t.id = 'viz-tip'; t.setAttribute('role', 'tooltip'); t.hidden = true; document.body.append(t); }
  el.onpointermove = (e) => {
    const s = e.target.closest('[data-tip]');
    if (!s) { t.hidden = true; return; }
    t.hidden = false;
    t.textContent = s.dataset.tip;
    t.style.left = `${Math.min(e.clientX + 14, innerWidth - t.offsetWidth - 8)}px`;
    t.style.top = `${e.clientY + 14}px`;
  };
  el.onpointerleave = () => { t.hidden = true; };
}

