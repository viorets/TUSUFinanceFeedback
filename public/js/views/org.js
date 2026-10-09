// Organization detail: compact master-data header, then three tabs (Pre-pitch · Pitching · Re-submission).
// One `org` call loads everything; each tab's code is imported only when that tab is opened.
import { tr, getLang } from '../i18n.js';
import { $, esc, icon, phoneFmt, dateTimeFmt, toast } from '../ui.js';
import { cached, peek, remember } from '../api.js';
import { CAMPUSES, CATEGORIES, ORGS } from '../masterdata.js';
import { UNFINISHED } from '../project.js';

const TABS = {
  prepitch: { load: () => import('./prepitch.js'), label: () => tr('Pre-pitch', 'ก่อนการนำเสนอ') },
  pitching: { load: () => import('./pitching.js'), label: () => tr('Pitching', 'การนำเสนอ') },
  resubmit: { load: () => import('./resubmit.js'), label: () => tr('Re-submission', 'การส่งแก้ไข') },
};

let state = null; // { page, ctx, ref, data, tab, mod }

export const isDirty = () => !!state?.mod?.isDirty?.();

export async function render(page, ctx, route) {
  const body = { ref: route.ref };
  // Tabs call ctx.persist() after a save so this device keeps the changed copy.
  ctx.persist = () => { remember('org', body, state.data); paintTabs(); };
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

  page.innerHTML = `<a class="btn plain back" href="#/">${icon('back', 18)}${tr('Organizations', 'องค์กร')}</a>
    <section class="group orghead">
      <div class="oh-line"><span class="code ref">${esc(route.ref)}</span><h1>${esc(o?.en || route.ref)}</h1></div>
      <div class="oh-line muted">${esc(label(CAMPUSES, o?.campus))} · ${esc(label(CATEGORIES, o?.category))}${ORGS.includes(o) ? '' : ` · ${tr('not in the masterfile', 'ไม่อยู่ในรายชื่อหลัก')}`}</div>
      <details><summary>${tr('More information', 'ข้อมูลเพิ่มเติม')}</summary>
        <dl class="kv">
          <dt>${tr('Submitter', 'ผู้ส่ง')}</dt><dd>${s ? contact(s.name, s.position, s.phone) : '—'}${s?.phone ? ` <span class="pill blue">Telegram</span>` : ''}</dd>
          <dt>${tr('Emergency contact', 'ผู้ติดต่อกรณีฉุกเฉิน')}</dt><dd>${s ? contact(s.emName, s.emPosition, s.emPhone) : '—'}</dd>
          <dt>${tr('Person in charge', 'ผู้รับผิดชอบ')}</dt><dd>${data.people.map((p) => esc(ctx.nameOf(p))).join(', ') || '—'}</dd>
          <dt>${tr('Latest submission', 'ส่งล่าสุด')}</dt><dd>${s ? `${dateTimeFmt(s.at)} · ${tr(`version ${s.version}`, `ครั้งที่ ${s.version}`)}` : tr('Not submitted yet', 'ยังไม่ได้ส่ง')}</dd>
        </dl></details>
    </section>
    <nav class="seg-tabs" id="tabs"></nav>
    <div id="tab"></div>`;
  await switchTab(route.tab);
}

// Tab strip; each tab shows how many projects are still unfinished there.
function paintTabs() {
  const { page, ref, data, tab } = state;
  $('#tabs', page).innerHTML = Object.entries(TABS).map(([k, t]) => {
    const n = data.projects.filter((p) => UNFINISHED[k](data, p)).length;
    return `<a href="#/org/${encodeURIComponent(ref)}/${k}" class="${k === tab ? 'on' : ''}">${esc(t.label())}${n ? `<span class="badge" title="${esc(tr(`${n} unfinished`, `ยังไม่ครบ ${n}`))}">${n}</span>` : ''}</a>`;
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
  if (!state.data.projects.length) {
    el.innerHTML = `<div class="empty">${tr('This organization has no submitted projects yet.', 'องค์กรนี้ยังไม่มีโครงการที่ส่งเข้ามา')}</div>`;
    state.mod = null;
    return;
  }
  const mod = await TABS[tab].load();
  if (state.tab !== tab) return; // another tab was picked while this one loaded
  state.mod = mod;
  mod.render(el, state.data, state.ctx);
}
