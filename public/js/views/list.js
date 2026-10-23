// Organization list: every organization in the masterfile, grouped by campus, sorted by RefCode.
// Organizations an admin removed on the pitching site are left out here (the admin console lists them).
// Persons in charge (v2.10): anyone with a role picks who is in charge of an organization; several people can share one.
// Budget totals per organization: proposed (its projects' budgets) and approved (Pitching final approved budgets).
// Tabs (v2.10): All / Pre-pitch / Pitching / Re-submission. Each has its own budget column and a "% of total" column
// (the organization's budget over the total of the organizations shown, so it follows the filters).
// Filters (v2.10): category, campus and a budget range, kept on this device; they combine with search and "Mine only".
// A mini dashboard above the list totals what is shown, for the open tab.
// Status: Open / Closed, shown here; Admins and Chiefs change it at the top of the Re-submission tab.
// A closed case counts as Done whatever its projects' progress (v2.3).
// Each campus is its own card; fixed column widths keep the cards lined up.
// "Mine only" (v2.6) keeps the organizations I'm in charge of; it combines with the filters and is kept on this device.
// "Print campus" (v2.7) prints a summary of that campus: each organization with its projects as sub-rows
// (print.js campusSheet, one printData call).
import { tr, getLang } from '../i18n.js';
import { $, esc, icon, toast, num2, dateTimeFmt, dialog } from '../ui.js';
import { cached, call, peek, remember, forget } from '../api.js';
import { CAMPUSES, CATEGORIES, ORGS } from '../masterdata.js';
import { slotText, timeLeft } from '../project.js';
import { timeline, bindTimeline } from '../timeline.js';

let query = '';
let tab = 'all'; // all | pre | pitch | resub
let sub = 'all'; // Pitching and Re-submission: all | pending | done
let mine = false;
// Filters: cats / camps = codes picked (none picked = no filter); on / min / max = budget range; type = budget the
// range uses on the All tab; rounds = submission round codes picked (Re-submission tab only, v2.13).
let F = { cats: [], camps: [], rounds: [], on: false, min: null, max: null, type: 'proposed' };
try {
  mine = localStorage.getItem('mineOnly') === '1';
  F = { ...F, ...JSON.parse(localStorage.getItem('listFilters') || '{}') };
} catch { /* private mode */ }
const saveF = () => { try { localStorage.setItem('listFilters', JSON.stringify(F)); } catch { /* this visit only */ } };

// The budget each tab is about, and the labels of the choices on All.
const VAL = { proposed: (s) => s.proposed || 0, est: (s) => s.est || 0, approved: (s) => s.approved || 0, fin: (s) => s.fin || 0 };
const TAB_VAL = { pre: 'est', pitch: 'approved', resub: 'fin' };
const TYPE_LABEL = {
  proposed: () => tr('Proposed budget', 'งบประมาณที่เสนอ'), est: () => tr('Estimated budget', 'งบประมาณที่คาดว่าจะอนุมัติ'),
  approved: () => tr('Approved budget', 'งบประมาณที่อนุมัติ'), fin: () => tr('Final budget', 'งบประมาณสุดท้าย'),
};

export async function render(page, ctx) {
  const optList = (list, picked, name) => list.map((x) => `<label class="msel-opt"><input type="checkbox" data-f="${name}" value="${esc(x.code)}" ${picked.includes(x.code) ? 'checked' : ''}>${esc(getLang() === 'th' ? x.th : x.en)}</label>`).join('');
  page.innerHTML = `<h1 class="grad">${tr('Organizations', 'องค์กร')}</h1><p class="muted" id="sum">&nbsp;</p><div id="timeline"></div>
    <div class="phase-bar"><div class="seg filter" role="radiogroup" aria-label="${esc(tr('Phase', 'ช่วง'))}" id="show"></div>
      <div class="track" id="track" aria-live="polite"></div></div>
    <div id="mini" class="kpis mini"></div>
    <div class="toolbar"><input type="search" id="q" placeholder="${esc(tr('Search name, RefCode or person in charge', 'ค้นหาชื่อ รหัส หรือผู้รับผิดชอบ'))}" value="${esc(query)}">
      <details class="msel" id="fcat"><summary>${tr('Category', 'หมวด')} <small id="fcatN"></small></summary><div class="msel-body">${optList(CATEGORIES, F.cats, 'cats')}</div></details>
      <details class="msel" id="fcamp"><summary>${tr('Campus', 'วิทยาเขต')} <small id="fcampN"></small></summary><div class="msel-body">${optList(CAMPUSES, F.camps, 'camps')}</div></details>
      <details class="msel" id="fround" hidden><summary>${tr('Round code', 'รหัสรอบ')} <small id="froundN"></small></summary><div class="msel-body" id="froundBody"></div></details>
      <label class="check"><input type="checkbox" id="mine" ${mine ? 'checked' : ''}>${tr('Mine only', 'เฉพาะของฉัน')} <small class="muted" id="mineN"></small></label></div>
    <div class="toolbar range" id="range">
      <label class="check sw-on"><input type="checkbox" id="ron" ${F.on ? 'checked' : ''}>${tr('Budget range', 'ช่วงงบประมาณ')}</label>
      <select id="rtype" aria-label="${esc(tr('Budget', 'งบประมาณ'))}">${Object.entries(TYPE_LABEL).map(([k, l]) => `<option value="${k}" ${F.type === k ? 'selected' : ''}>${esc(l())}</option>`).join('')}</select>
      <input type="number" id="rmin" min="0" step="any" placeholder="${esc(tr('Min', 'ต่ำสุด'))}" aria-label="${esc(tr('Minimum', 'ต่ำสุด'))}">
      <div class="dual"><input type="range" id="rlo" min="0" max="100" step="1" aria-label="${esc(tr('Minimum', 'ต่ำสุด'))}"><input type="range" id="rhi" min="0" max="100" step="1" aria-label="${esc(tr('Maximum', 'สูงสุด'))}"></div>
      <input type="number" id="rmax" min="0" step="any" placeholder="${esc(tr('Max', 'สูงสุด'))}" aria-label="${esc(tr('Maximum', 'สูงสุด'))}"></div>
    <div id="subbar" class="toolbar"></div>
    <div id="rows" class="org-list"></div>`;

  let boot = null;
  const closedOf = (ref) => boot.closed?.[ref] || null;
  const label = (list, code) => { const x = list.find((i) => i.code === code); return x ? (getLang() === 'th' ? x.th : x.en) : code; };
  const stat = (ref) => boot.orgs[ref] || (boot.orgs[ref] = { count: 0, pending: 0, people: [] });
  // Projects still unfinished (a closed case has none): Pitching + Re-submission together, or each alone.
  const left = (ref) => (closedOf(ref) ? 0 : stat(ref).pending);
  const leftP = (ref) => (closedOf(ref) ? 0 : stat(ref).pp || 0);
  const leftR = (ref) => (closedOf(ref) ? 0 : stat(ref).rp || 0);
  const unfinished = (o) => (tab === 'pitch' ? leftP(o.code) : tab === 'resub' ? leftR(o.code) : left(o.code));
  const valKey = () => (tab === 'all' ? F.type : TAB_VAL[tab] || 'proposed');
  const valOf = (o) => VAL[valKey()](stat(o.code));

  const money = (v) => (v == null ? '—' : num2(v));
  const status = (o) => (closedOf(o.code) ? `<span class="pill closed">${icon('lock', 13)}${tr('Closed', 'ปิดแล้ว')}</span>` : `<span class="pill green">${tr('Open', 'เปิด')}</span>`);

  // Several people can be in charge: a picker of checkboxes, or the names where picking is off. Admins and Chiefs pick
  // anyone with a role; Staff and Data team staff only themselves (v2.11, the server checks too).
  const picText = (people) => people.map((p) => esc(ctx.nameOf(p))).join(', ') || '—';
  const pic = (o) => {
    const people = stat(o.code).people;
    if (!ctx.canEdit) return people.map((p) => esc(ctx.nameOf(p))).join('<br>') || '<span class="muted">—</span>';
    const users = ctx.console ? [...(boot.users || [])].sort((a, b) => ctx.nameOf(a).localeCompare(ctx.nameOf(b))) : [ctx.me.email];
    return `<details class="pic-pick" data-pic-for="${esc(o.code)}"><summary title="${esc(tr('Person in charge', 'ผู้รับผิดชอบ'))}">${picText(people)}</summary>
      <div class="msel-body">${users.map((u) => `<label class="msel-opt"><input type="checkbox" data-pic="${esc(o.code)}" value="${esc(u)}" ${people.includes(u) ? 'checked' : ''}>${esc(ctx.nameOf(u))}</label>`).join('') || `<span class="muted">${tr('No users', 'ไม่มีผู้ใช้')}</span>`}</div></details>`;
  };

  // The next pitching slot of an organization (one still to come, else none).
  const nextSlot = (s) => (s.slots || []).map(([date, start, end]) => ({ date, start, end }))
    .sort((a, b) => `${a.date} ${a.start}`.localeCompare(`${b.date} ${b.start}`))
    .find((x) => { const t = timeLeft(x); return t && t !== tr('Pitched', 'นำเสนอแล้ว'); }) || null;

  // Re-submission: Submitted when a re-submission was sent; the dropdown lists every version.
  const approvedNow = (s) => s.count > 0 && s.ok >= s.count;
  const decisionPill = (s) => (approvedNow(s) ? `<span class="pill green">${tr('Approved', 'อนุมัติ')}</span>` : `<span class="pill orange">${tr('Not approved', 'ยังไม่อนุมัติ')}</span>`);
  const resubCell = (s) => {
    if (!s.rs) return `<span class="pill">${tr('Not submitted', 'ยังไม่ส่ง')}</span>`;
    const items = (s.vers || []).map((v) => `<li><b>v${v.v ?? '?'}</b> <span class="muted">${v.at ? esc(dateTimeFmt(v.at)) : '—'}</span>
      <span class="num">${v.total == null ? '' : `฿${num2(v.total)}`}</span> <span class="muted">${v.n == null ? '' : tr('{n} projects', '{n} โครงการ', { n: v.n })}</span> ${decisionPill(s)}</li>`).join('');
    return `<details class="vers"><summary><span class="pill blue">${tr('Submitted', 'ส่งแล้ว')}</span> ${decisionPill(s)}</summary><ul class="vers-list">${items}</ul></details>`;
  };
  const progCell = (o, s) => {
    if (!s.count && !closedOf(o.code)) return '—';
    const lp = leftP(o.code);
    return lp ? `<span class="pill orange">${s.count - lp} / ${s.count}</span>` : `<span class="pill green">${tr('Done', 'ครบ')}</span>`;
  };
  const unfCell = (o, s, n) => (s.count || closedOf(o.code) ? (n ? `<span class="pill orange">${n} / ${s.count}</span>` : `<span class="pill green">${tr('Done', 'ครบ')}</span>`) : '—');

  // Columns: heading, width weight, cell. `share` is filled in paint (needs the total shown).
  const COL = {
    cat: { h: () => tr('Category', 'หมวด'), w: 6, c: (o) => `<td class="mono" title="${esc(label(CATEGORIES, o.category))}">${esc(o.category)}</td>` },
    name: { h: () => tr('Name', 'ชื่อ'), w: 20, c: (o) => `<td>${esc(o.en)}</td>` },
    ref: { h: () => tr('RefCode', 'รหัส'), w: 9, c: (o) => `<td class="mono">${esc(o.code)}</td>` },
    n: { h: () => tr('Projects', 'โครงการ'), w: 6, r: 1, c: (o, s) => `<td class="r">${s.count || '—'}</td>` },
    prop: { h: () => tr('Total proposed project budget', 'งบประมาณโครงการที่เสนอรวม'), w: 11, r: 1, c: (o, s) => `<td class="r">${s.count ? money(s.proposed) : '—'}</td>` },
    appr: { h: () => tr('Total approved budget', 'งบประมาณที่อนุมัติรวม'), w: 11, r: 1, c: (o, s) => `<td class="r">${money(s.approved)}</td>` },
    est: { h: () => tr('Estimated budget', 'งบประมาณที่คาดว่าจะอนุมัติ'), w: 11, r: 1, c: (o, s) => `<td class="r">${s.count ? money(s.est) : '—'}</td>` },
    fin: { h: () => tr('Final budget', 'งบประมาณสุดท้าย'), w: 11, r: 1, c: (o, s) => `<td class="r">${money(s.fin)}</td>` },
    share: { h: () => tr('% of total', '% ของยอดรวม'), w: 6, r: 1, c: (o, s, x) => `<td class="r">${x.total ? `${(Math.round((VAL[valKey()](s) / x.total) * 1000) / 10).toLocaleString('en-US')}%` : '—'}</td>` },
    next: { h: () => tr('Next pitching', 'นำเสนอครั้งถัดไป'), w: 13, c: (o, s) => { const n = nextSlot(s); return `<td class="small">${n ? esc(slotText(n)).replace(/ · Microsoft Teams$/, '') : '<span class="muted">—</span>'}</td>`; } },
    prog: { h: () => tr('Pitching progress', 'ความคืบหน้าการนำเสนอ'), w: 11, r: 1, c: (o, s) => { const n = nextSlot(s); const t = n ? timeLeft(n) : ''; return `<td class="r">${progCell(o, s)}${t ? `<small class="muted block">${esc(t)}</small>` : ''}</td>`; } },
    rs: { h: () => tr('Re-submission', 'การส่งแก้ไข'), w: 14, c: (o, s) => `<td class="small">${resubCell(s)}</td>` },
    pic: { h: () => tr('Person in charge', 'ผู้รับผิดชอบ'), w: 15, c: (o) => `<td class="small">${pic(o)}</td>` },
    unf: { h: () => tr('Unfinished', 'ยังไม่ครบ'), w: 8, r: 1, c: (o, s) => `<td class="r">${unfCell(o, s, left(o.code))}</td>` },
    unfR: { h: () => tr('Unfinished', 'ยังไม่ครบ'), w: 8, r: 1, c: (o, s) => `<td class="r">${unfCell(o, s, leftR(o.code))}</td>` },
    case: { h: () => tr('Status', 'สถานะ'), w: 8, c: (o) => `<td>${status(o)}</td>` },
  };
  const COLSET = {
    all: ['cat', 'name', 'ref', 'n', 'prop', 'appr', 'share', 'pic', 'unf', 'case'],
    pre: ['cat', 'name', 'ref', 'n', 'prop', 'est', 'share', 'pic', 'case'],
    pitch: ['cat', 'name', 'ref', 'n', 'prop', 'appr', 'share', 'next', 'prog', 'pic', 'case'],
    resub: ['cat', 'name', 'ref', 'n', 'prop', 'fin', 'share', 'rs', 'pic', 'unfR', 'case'],
  };
  const TABS = { all: () => tr('All', 'ทั้งหมด'), pre: () => tr('Pre-pitch', 'ก่อนการนำเสนอ'), pitch: () => tr('Pitching', 'การนำเสนอ'), resub: () => tr('Re-submission', 'การส่งแก้ไข') };
  const SUBS = { all: () => tr('All', 'ทั้งหมด'), pending: () => tr('Unfinished', 'ยังไม่ครบ'), done: () => tr('Finished', 'ครบแล้ว') };

  // Mini dashboard tiles for the open tab, from the organizations shown.
  // Tracking line in the phase bar (v2.11, Pitching and Re-submission): projects finished of all, as a meter, and
  // the organizations still open, for the organizations shown. A closed case counts as finished.
  const track = (list) => {
    if (tab !== 'pitch' && tab !== 'resub') return '';
    const left = tab === 'pitch' ? leftP : leftR;
    const all = list.reduce((t, o) => t + (stat(o.code).count || 0), 0);
    const open = list.reduce((t, o) => t + left(o.code), 0);
    const orgsLeft = list.filter((o) => left(o.code) > 0).length;
    const p = all ? Math.round(((all - open) / all) * 100) : 0;
    const what = tab === 'pitch' ? tr('finished', 'ที่ครบแล้ว') : tr('decided', 'ที่ตัดสินแล้ว');
    return `<div class="meter" role="meter" aria-valuemin="0" aria-valuemax="${all}" aria-valuenow="${all - open}" aria-label="${esc(what)}"><i style="width:${p}%"></i></div>
      <span><b>${(all - open).toLocaleString('en-US')} / ${all.toLocaleString('en-US')}</b> ${esc(tr('projects {what} ({p}%)', 'โครงการ{what} ({p}%)', { what, p }))}</span>
      <span class="muted">${orgsLeft ? esc(tr('{n} organizations unfinished', 'ยังไม่ครบ {n} องค์กร', { n: orgsLeft })) : esc(tr('all organizations finished', 'ครบทุกองค์กร'))}</span>`;
  };
  const mini = (list) => {
    const sum = (f) => list.reduce((t, o) => t + (f(stat(o.code), o) || 0), 0);
    const t = (name, v, note = '') => `<div class="kpi"><span>${esc(name)}</span><b class="hero num">${v}</b>${note ? `<p class="muted small">${note}</p>` : ''}</div>`;
    const baht = (v) => `฿${num2(v)}`;
    const nProj = sum((s) => s.count);
    if (tab === 'pre') return t(tr('Total proposed budget', 'งบประมาณที่เสนอทั้งหมด'), baht(sum((s) => s.proposed)), tr('{n} projects', '{n} โครงการ', { n: nProj }))
      + t(tr('Estimated budget', 'งบประมาณที่คาดว่าจะอนุมัติ'), baht(sum((s) => s.est)));
    if (tab === 'pitch') {
      const notYet = sum((s, o) => (closedOf(o.code) ? 0 : s.pp));
      const next = list.map((o) => nextSlot(stat(o.code))).filter(Boolean).sort((a, b) => `${a.date} ${a.start}`.localeCompare(`${b.date} ${b.start}`))[0];
      return t(tr('Total approved budget', 'งบประมาณที่อนุมัติรวม'), baht(sum((s) => s.approved)))
        + t(tr('Pitched', 'นำเสนอแล้ว'), nProj - notYet, tr('of {n} projects', 'จาก {n} โครงการ', { n: nProj }))
        + t(tr('Not yet pitched', 'ยังไม่ได้นำเสนอ'), notYet)
        + t(tr('Next pitching', 'นำเสนอครั้งถัดไป'), next ? esc(next.date) : '—', next ? esc(`${next.start}${next.end ? `–${next.end}` : ''} · ${timeLeft(next)}`) : '');
    }
    if (tab === 'resub') {
      const sent = list.filter((o) => stat(o.code).rs).length;
      const ok = list.filter((o) => approvedNow(stat(o.code))).length;
      return t(tr('Submitted', 'ส่งแล้ว'), sent, tr('{n} not submitted', 'ยังไม่ส่ง {n}', { n: list.length - sent }))
        + t(tr('Approved', 'อนุมัติ'), ok, tr('{n} not approved', 'ยังไม่อนุมัติ {n}', { n: list.length - ok }))
        + t(tr('Final budget', 'งบประมาณสุดท้าย'), baht(sum((s) => s.fin)));
    }
    return t(tr('Total proposed budget', 'งบประมาณที่เสนอทั้งหมด'), baht(sum((s) => s.proposed)))
      + t(tr('Total approved budget', 'งบประมาณที่อนุมัติรวม'), baht(sum((s) => s.approved)))
      + t(tr('Organizations', 'องค์กร'), list.length.toLocaleString('en-US'))
      + t(tr('Projects', 'โครงการ'), nProj.toLocaleString('en-US'));
  };

  // The range controls follow the open tab: the maximum is the largest budget of the kind the tab uses.
  const syncRange = (rows) => {
    const max = Math.max(0, ...rows.map(valOf));
    const top = Math.max(1, Math.ceil(max / 1000) * 1000);
    const lo = F.min == null ? 0 : Math.min(F.min, top); const hi = F.max == null ? top : Math.min(F.max, top);
    ['#rlo', '#rhi'].forEach((id) => { $(id, page).max = top; $(id, page).step = Math.max(1, Math.round(top / 200)); $(id, page).disabled = !F.on; });
    $('#rlo', page).value = lo; $('#rhi', page).value = hi;
    $('#rmin', page).value = F.min ?? ''; $('#rmax', page).value = F.max ?? '';
    ['#rmin', '#rmax'].forEach((id) => { $(id, page).disabled = !F.on; });
    $('#rtype', page).hidden = tab !== 'all'; $('#rtype', page).disabled = !F.on;
  };

  const paint = () => {
    const removed = new Set(boot.removed);
    const rows = ORGS.filter((o) => !removed.has(o.code));
    $('#show', page).innerHTML = Object.entries(TABS).map(([k, l]) => `<label><input type="radio" name="show" value="${k}" ${tab === k ? 'checked' : ''}><span>${esc(l())}</span></label>`).join('');
    syncRange(rows);
    $('#fcatN', page).textContent = F.cats.length || ''; $('#fcampN', page).textContent = F.camps.length || '';
    // Round code (Re-submission tab): the rounds found in the submissions, picked ones first kept even if gone.
    const roundOf = (o) => stat(o.code).round || '';
    const byRound = tab === 'resub' && F.rounds.length > 0;
    $('#fround', page).hidden = tab !== 'resub';
    if (tab === 'resub') {
      const rounds = [...new Set([...rows.map(roundOf).filter(Boolean), ...F.rounds])].sort();
      $('#froundN', page).textContent = F.rounds.length || '';
      $('#froundBody', page).innerHTML = rounds.length ? rounds.map((r) => `<label class="msel-opt"><input type="checkbox" data-f="rounds" value="${esc(r)}" ${F.rounds.includes(r) ? 'checked' : ''}>${esc(r)}
        <small class="muted">${rows.filter((o) => roundOf(o) === r).length}</small></label>`).join('')
        : `<p class="muted small msel-opt">${esc(tr('No round codes in the submissions yet', 'ยังไม่มีรหัสรอบในข้อมูลที่ส่งเข้ามา'))}</p>`;
    }
    const q = query.trim().toLowerCase();
    const isMine = (o) => stat(o.code).people.includes(ctx.me.email);
    $('#mineN', page).textContent = rows.filter(isMine).length;
    // Finished / Unfinished filter: Pitching and Re-submission (Pre-pitch is not tracked).
    const tracked = tab === 'pitch' || tab === 'resub';
    const isDone = (o) => !!closedOf(o.code) || (stat(o.code).count > 0 && !unfinished(o));
    const subKeep = { all: () => true, pending: (o) => unfinished(o) > 0, done: isDone };
    $('#subbar', page).hidden = !tracked;
    $('#subbar', page).innerHTML = tracked ? `<div class="seg filter" role="radiogroup" aria-label="${esc(tr('Show', 'แสดง'))}" id="sub">${Object.entries(SUBS).map(([k, l]) => `<label><input type="radio" name="sub" value="${k}" ${sub === k ? 'checked' : ''}><span>${esc(l())}
      <small>${rows.filter((o) => subKeep[k](o)).length}</small></span></label>`).join('')}</div>` : '';
    const inRange = (o) => { if (!F.on) return true; const v = valOf(o); return v >= (F.min ?? 0) && v <= (F.max ?? Infinity); };
    const hit = (o) => (!F.cats.length || F.cats.includes(o.category)) && (!F.camps.length || F.camps.includes(o.campus)) && inRange(o)
      && (!byRound || F.rounds.includes(roundOf(o)))
      && (!tracked || subKeep[sub](o)) && (!mine || isMine(o)) && (!q || `${o.en} ${o.code} ${stat(o.code).people.map((p) => `${p} ${ctx.nameOf(p)}`).join(' ')}`.toLowerCase().includes(q));
    const shown = rows.filter(hit);
    const pending = shown.filter((o) => left(o.code) > 0).length;
    $('#sum', page).innerHTML = tr('{n} organizations · ', '{n} องค์กร · ', { n: shown.length })
      + (pending ? `<b class="warn">${tr('{n} with unfinished projects', '{n} องค์กรมีโครงการที่ยังไม่ครบ', { n: pending })}</b>` : tr('all projects finished', 'ทุกโครงการครบแล้ว'));
    $('#mini', page).innerHTML = mini(shown);
    $('#track', page).innerHTML = track(shown);
    const keys = COLSET[tab];
    const weight = keys.reduce((t, k) => t + COL[k].w, 0);
    const COLS = `<colgroup>${keys.map((k) => `<col style="width:${(COL[k].w / weight * 100).toFixed(2)}%">`).join('')}</colgroup>`;
    const HEAD = `<tr>${keys.map((k) => `<th${COL[k].r ? ' class="r"' : ''}>${esc(COL[k].h())}</th>`).join('')}</tr>`;
    const card = (top, body) => `<div class="group scroll-x campus-card"><table class="grid-table quotes">${COLS}<thead>${top}${HEAD}</thead><tbody>${body}</tbody></table></div>`;
    const x = { total: shown.reduce((t, o) => t + valOf(o), 0) };
    const narrowed = q || mine || F.cats.length || F.camps.length || byRound || F.on || (tracked && sub !== 'all');
    $('#rows', page).innerHTML = CAMPUSES.map((c) => {
      const list = shown.filter((o) => o.campus === c.code).sort((a, b) => a.code.localeCompare(b.code));
      const pend = list.filter((o) => unfinished(o) > 0).length;
      const head = `<tr class="grp"><th colspan="${keys.length}">${esc(c.en.toUpperCase())}<span class="muted"> · ${esc(c.th)} · ${list.length}</span>
        ${pend ? `<span class="pill orange">${tr('{n} pending', 'รอ {n}', { n: pend })}</span>` : ''}
        ${rows.some((o) => o.campus === c.code && stat(o.code).count) ? `<button class="btn plain campus-print" data-print-campus="${esc(c.code)}">${icon('print', 15)}${tr('Print campus', 'พิมพ์ทั้งวิทยาเขต')}</button>` : ''}</th></tr>`;
      // Searching or filtering: campuses with no match are left out.
      if (!list.length) return narrowed ? '' : card(head, `<tr class="none"><td colspan="${keys.length}" class="muted">${tr('No organizations', 'ไม่มีองค์กร')}</td></tr>`);
      return card(head, list.map((o) => {
        const s = stat(o.code);
        return `<tr data-ref="${esc(o.code)}" tabindex="0">${keys.map((k) => COL[k].c(o, s, x)).join('')}</tr>`;
      }).join(''));
    }).join('') || card('', `<tr class="none"><td colspan="${keys.length}" class="muted">${tr('No organizations match', 'ไม่มีองค์กรที่ตรงกัน')}</td></tr>`);
  };

  // This device's last copy first (instant), then the fresh one if it differs.
  const kept = peek('boot');
  // Timeline (v2.12): every organization on the list, whatever the filters; a hover shows Pitching / Re-submission progress.
  const paintTimeline = () => {
    const removed = new Set(boot.removed);
    const rows = ORGS.filter((o) => !removed.has(o.code));
    const all = rows.reduce((t, o) => t + (stat(o.code).count || 0), 0);
    const done = (left) => ({ done: all - rows.reduce((t, o) => t + left(o.code), 0), all });
    const prog = all ? { pitching: done(leftP), resubCampus: done(leftR), resubCentre: done(leftR) } : {};
    $('#timeline', page).innerHTML = timeline(boot.dates, prog);
    bindTimeline(page);
  };
  if (kept) { boot = ctx.boot = kept; paint(); paintTimeline(); }
  try {
    const fresh = await cached('boot');
    if (!kept || JSON.stringify(fresh) !== JSON.stringify(kept)) { boot = ctx.boot = fresh; paint(); paintTimeline(); }
  } catch (e) {
    if (!kept) throw e;
    toast(e.message, 'bad');
  }

  $('#q', page).oninput = (e) => { query = e.target.value; paint(); };
  page.addEventListener('change', (e) => {
    const t = e.target;
    if (t.name === 'show') { tab = t.value; paint(); }
    else if (t.name === 'sub') { sub = t.value; paint(); }
    else if (t.dataset.f) { // category / campus pickers
      F[t.dataset.f] = [...page.querySelectorAll(`input[data-f="${t.dataset.f}"]:checked`)].map((i) => i.value);
      saveF(); paint();
    }
  });
  $('#mine', page).onchange = (e) => {
    mine = e.target.checked;
    try { localStorage.setItem('mineOnly', mine ? '1' : '0'); } catch { /* this visit only */ }
    paint();
  };
  // Budget range: the switch, the type (All tab), the two sliders and the two boxes all write F.min / F.max.
  const num = (v) => (v === '' || v == null || Number.isNaN(+v) ? null : Math.max(0, +v));
  const setRange = (min, max) => { F.min = min; F.max = max; if (min != null && max != null && min > max) [F.min, F.max] = [max, min]; saveF(); paint(); };
  $('#ron', page).onchange = (e) => { F.on = e.target.checked; saveF(); paint(); };
  $('#rtype', page).onchange = (e) => { F.type = e.target.value; F.min = null; F.max = null; saveF(); paint(); };
  $('#rmin', page).onchange = (e) => setRange(num(e.target.value), F.max);
  $('#rmax', page).onchange = (e) => setRange(F.min, num(e.target.value));
  $('#rlo', page).oninput = (e) => { const top = +$('#rhi', page).max; setRange(+e.target.value <= 0 ? null : +e.target.value, F.max == null || F.max >= top ? null : F.max); };
  $('#rhi', page).oninput = (e) => { const top = +e.target.max; setRange(F.min, +e.target.value >= top ? null : +e.target.value); };

  const rowsEl = $('#rows', page);
  const open = (e) => {
    if (e.target.closest('details, label, select, a')) return;
    const r = e.target.closest('tr[data-ref]');
    if (r) ctx.go(`#/org/${encodeURIComponent(r.dataset.ref)}`);
  };
  rowsEl.onclick = (e) => {
    const b = e.target.closest('[data-print-campus]');
    if (b) { printCampus(b); return; }
    open(e);
  };
  // Print campus: load every organization of the campus in one call, print the summary table.
  const printCampus = async (b) => {
    const c = CAMPUSES.find((x) => x.code === b.dataset.printCampus);
    const { campusSheet, printNow } = await import('../print.js');
    b.disabled = true;
    toast(tr('Preparing the print…', 'กำลังเตรียมการพิมพ์…'));
    try {
      const out = await call('printData', { campus: c.code });
      const orgs = out.orgs.filter((x) => x.data.org);
      if (orgs.length) printNow(campusSheet(c.code, orgs), 'A4'); else toast(tr('No submitted organizations in this campus', 'ไม่มีองค์กรที่ส่งข้อมูลในวิทยาเขตนี้'), 'bad');
    } catch (err) {
      toast(err.message, 'bad');
    }
    b.disabled = false;
  };
  rowsEl.onkeydown = (e) => { if (e.key === 'Enter') open(e); };

  // Hovering a row for a moment starts loading that organization, so the click opens it faster.
  let hover;
  rowsEl.onpointerover = (e) => {
    const r = e.target.closest('tr[data-ref]');
    clearTimeout(hover);
    if (r && stat(r.dataset.ref).count) hover = setTimeout(() => cached('org', { ref: r.dataset.ref }).catch(() => {}), 250);
  };

  // Person in charge: ticking a name adds that person, unticking removes them (the others stay). Staff ticking
  // themselves where someone else is in charge get a caution window first: add themselves too, or replace them.
  rowsEl.onchange = async (e) => {
    const box = e.target.closest('input[data-pic]');
    if (!box) return;
    const ref = box.dataset.pic;
    const on = box.checked;
    const others = stat(ref).people.filter((p) => p !== box.value);
    let replace = false;
    if (on && !ctx.console && others.length) {
      const pick = await dialog(tr('Someone is already in charge', 'มีผู้รับผิดชอบอยู่แล้ว'),
        `<p>${esc(tr('{who} is in charge of {ref}.', '{who} รับผิดชอบ {ref} อยู่', { who: others.map((p) => ctx.nameOf(p)).join(', '), ref }))}</p>
        <p>${esc(tr('Replace removes them from this organization.', 'แทนที่ จะนำผู้รับผิดชอบเดิมออกจากองค์กรนี้'))}</p>`,
        [{ v: '', label: tr('Cancel', 'ยกเลิก') }, { v: 'replace', label: tr('Replace', 'แทนที่'), cls: 'danger' }, { v: 'add', label: tr('Add me too', 'เพิ่มฉันด้วย'), cls: 'primary' }], { none: '' });
      if (!pick) { box.checked = false; return; }
      replace = pick === 'replace';
    }
    const group = [...rowsEl.querySelectorAll(`input[data-pic="${ref}"]`)];
    group.forEach((i) => { i.disabled = true; });
    try {
      // Replace = the take-over call (no `on`): the person becomes the only one in charge.
      const out = await call('setInCharge', replace ? { ref, email: box.value } : { ref, email: box.value, on });
      stat(ref).people = out.people;
      remember('boot', {}, boot);
      forget('org');
      forget('admin'); // the console's Users version changed
      const sum = rowsEl.querySelector(`details[data-pic-for="${ref}"] summary`);
      if (sum) sum.innerHTML = picText(out.people);
      $('#mineN', page).textContent = ORGS.filter((o) => stat(o.code).people.includes(ctx.me.email)).length;
      toast(on ? tr('{who} is in charge of {ref}', '{who} รับผิดชอบ {ref}', { who: box.value, ref }) : tr('{who} is no longer in charge of {ref}', '{who} ไม่ได้รับผิดชอบ {ref} แล้ว', { who: box.value, ref }));
    } catch (err) {
      toast(err.message, 'bad');
      box.checked = !on;
    }
    group.forEach((i) => { i.disabled = false; });
  };
}
