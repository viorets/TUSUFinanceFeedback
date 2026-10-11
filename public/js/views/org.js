// Organization detail: compact master-data header (with Closed / Unclosed and Print), then three tabs
// (Pre-pitch · Pitching · Re-submission). One `org` call loads everything; each tab's code is imported only when that
// tab is opened. A closed organization's tabs are read-only for everyone (the server refuses its writes too).
// The header also shows the budget totals, as on the list: proposed (projects' budgets) and approved (Pitching final
// approved budgets); they repaint after each save.
import { tr, getLang } from '../i18n.js';
import { $, esc, icon, phoneFmt, dateTimeFmt, toast, num2 } from '../ui.js';
import { cached, peek, remember } from '../api.js';
import { CAMPUSES, CATEGORIES, ORGS } from '../masterdata.js';
import { UNFINISHED, phaseOn, PENDING } from '../project.js';
import { askPhases, printSheet, printNow } from '../print.js';

const TABS = {
  prepitch: { load: () => import('./prepitch.js'), label: () => tr('Pre-pitch', 'ก่อนการนำเสนอ') },
  pitching: { load: () => import('./pitching.js'), label: () => tr('Pitching', 'การนำเสนอ') },
  resubmit: { load: () => import('./resubmit.js'), label: () => tr('Re-submission', 'การส่งแก้ไข') },
};

let state = null; // { page, ctx, ref, data, tab, mod }

export const isDirty = () => !!state?.mod?.isDirty?.();
export const unsaved = () => state?.mod?.unsaved?.() || [];
export const saveAll = () => state?.mod?.saveAll?.();
export const canSaveAll = () => !!state?.mod?.saveAll;

export async function render(page, ctx, route) {
  const body = { ref: route.ref };
  // Tabs call ctx.persist() after a save so this device keeps the changed copy.
  ctx.persist = () => { remember('org', body, state.data); paintTabs(); };
  // After the case is closed or opened (Re-submission tab): repaint the header and lock / unlock every tab.
  ctx.redraw = () => draw(page, ctx, { ...route, tab: state.tab }, state.data);
  // This device's last copy first (instant), then the fresh one — unless edits are in progress.
  const kept = peek('org', body);
  if (kept) await draw(page, ctx, route, kept);
  try {
    const fresh = await cached('org', body);
    if (!kept) return draw(page, ctx, route, fresh);
    if (state.ref !== route.ref || state.page !== page || !page.isConnected) return; // left this page meanwhile
    if (JSON.stringify(fresh) !== JSON.stringify(kept) && !isDirty()) await draw(page, ctx, { ...route, tab: state.tab }, fresh);
  } catch (e) {
    if (!kept) throw e;
    toast(e.message, 'bad');
  }
}

async function draw(page, ctx, route, data) {
  const o = ORGS.find((x) => x.code === route.ref);
  state = { page, ctx, ref: route.ref, data, tab: null, mod: null };
  const label = (list, code) => { const x = list.find((i) => i.code === code); return x ? (getLang() === 'th' ? x.th : x.en) : code; };
  const s = data.org;
  const contact = (n, p, ph) => [n, p, phoneFmt(ph)].filter(Boolean).map(esc).join(' · ') || '—';
  const closed = data.closed;

  page.innerHTML = `<a class="btn plain back" href="#/">${icon('back', 18)}${tr('Organizations', 'องค์กร')}</a>
    <section class="group orghead">
      <div class="oh-top">${statusPill(closed)}<span class="grow"></span><button class="btn" id="print">${icon('print', 16)}${tr('Print', 'พิมพ์')}</button></div>
      <div class="oh-line"><span class="code ref">${esc(route.ref)}</span><h1>${esc(o?.en || route.ref)}</h1></div>
      <div class="oh-line muted">${esc(label(CAMPUSES, o?.campus))} · ${esc(label(CATEGORIES, o?.category))}${ORGS.includes(o) ? '' : ` · ${tr('not in the masterfile', 'ไม่อยู่ในรายชื่อหลัก')}`}</div>
      <dl class="oh-totals" id="totals"></dl>
      <details><summary>${tr('More information', 'ข้อมูลเพิ่มเติม')}</summary>
        <dl class="kv">
          <dt>${tr('Submitter', 'ผู้ส่ง')}</dt><dd>${s ? contact(s.name, s.position, s.phone) : '—'}${s?.phone ? ` <span class="pill blue">Telegram</span>` : ''}</dd>
          <dt>${tr('Emergency contact', 'ผู้ติดต่อกรณีฉุกเฉิน')}</dt><dd>${s ? contact(s.emName, s.emPosition, s.emPhone) : '—'}</dd>
          <dt>${tr('Person in charge', 'ผู้รับผิดชอบ')}</dt><dd>${data.people.map((p) => esc(ctx.nameOf(p))).join(', ') || '—'}</dd>
          <dt>${tr('Latest submission', 'ส่งล่าสุด')}</dt><dd>${s ? `${dateTimeFmt(s.at)} · ${tr('version {n}', 'ครั้งที่ {n}', { n: s.version })}` : tr('Not submitted yet', 'ยังไม่ได้ส่ง')}</dd>
        </dl></details>
    </section>
    ${closed ? `<div class="notice closed">${icon('lock', 18)}<span>${tr('Closed by {who} on {when}. Its data cannot be changed.', 'ปิดโดย {who} เมื่อ {when} ไม่สามารถแก้ไขข้อมูลได้', { who: esc(ctx.nameOf(closed.by)), when: dateTimeFmt(closed.at) })}</span></div>` : ''}
    <nav class="seg-tabs" id="tabs"></nav>
    <div id="tab"></div>`;
  // Print: pick the phases, then the cover + one page per phase (print.js).
  $('#print', page).onclick = async () => { const show = await askPhases(tr('Print', 'พิมพ์')); if (show) printNow(printSheet(route.ref, data, show, ctx.nameOf)); };
  await switchTab(route.tab);
}

export const statusPill = (closed) => (closed ? `<span class="pill closed">${icon('lock', 13)}${tr('Closed', 'ปิดแล้ว')}</span>`
  : `<span class="pill green">${tr('Open', 'เปิด')}</span>`);

// Tab strip; each tab shows how many projects are still unfinished there (none once the case is closed).
function paintTabs() {
  const { page, ref, data, tab } = state;
  const fins = data.projects.map((p) => data.pitch?.[p.ref]?.finalBudget).filter((v) => v != null && v !== '' && Number.isFinite(+v));
  const money = (v) => (v == null ? '—' : `฿${num2(v)}`);
  $('#totals', page).innerHTML = `<div><dt>${tr('Total proposed project budget', 'งบประมาณโครงการที่เสนอรวม')}</dt><dd>${data.projects.length ? money(data.projects.reduce((t, p) => t + (+p.budget || 0), 0)) : '—'}</dd></div>
    <div><dt>${tr('Total approved budget', 'งบประมาณที่อนุมัติรวม')}</dt><dd>${money(fins.length ? fins.reduce((t, v) => t + +v, 0) : null)}</dd></div>`;
  $('#tabs', page).innerHTML = Object.entries(TABS).map(([k, t]) => {
    const n = phaseOn(data, k) && !data.closed ? data.projects.filter((p) => UNFINISHED[k](data, p)).length : 0;
    return `<a href="#/org/${encodeURIComponent(ref)}/${k}" class="${k === tab ? 'on' : ''}">${esc(t.label())}${n ? `<span class="badge" title="${esc(tr('{n} unfinished', 'ยังไม่ครบ {n}', { n }))}">${n}</span>` : ''}</a>`;
  }).join('');
}

export async function switchTab(tab) {
  if (!TABS[tab]) tab = 'prepitch';
  const { page, ref } = state;
  state.tab = tab;
  paintTabs();
  // A fresh element per tab: each tab module binds its own handlers on it.
  const el = document.createElement('div');
  $('#tab', page).replaceChildren(el);
  // Phase switched off in the admin console: no cards, nothing to change.
  if (!phaseOn(state.data, tab)) {
    el.innerHTML = `<div class="empty pending-phase">${icon('clock', 32)}<h2>${esc(PENDING())}</h2><p>${tr('This phase has not started yet.', 'ช่วงนี้ยังไม่เริ่ม')}</p></div>`;
    state.mod = null;
    return;
  }
  if (!state.data.projects.length) {
    el.innerHTML = `<div class="empty">${tr('This organization has no submitted projects yet.', 'องค์กรนี้ยังไม่มีโครงการที่ส่งเข้ามา')}</div>`;
    state.mod = null;
    return;
  }
  const mod = await TABS[tab].load();
  if (state.tab !== tab) return; // another tab was picked while this one loaded
  state.mod = mod;
  // Closed: the tab sees canEdit = false, so every field, upload and Save is off.
  mod.render(el, state.data, state.data.closed ? Object.create(state.ctx, { canEdit: { value: false } }) : state.ctx);
}
