// Organization list: every organization in the masterfile, grouped by campus, sorted by RefCode.
// Organizations an admin removed on the pitching site are left out here (the admin console lists them).
// Admins pick the person in charge of each organization from a dropdown of users; others see the name.
// Filter: all / pending (some Pre-pitch approval unfinished) / done (every project finished).
// Status: Unclosed / Closed. Admins and Chiefs switch it with the button in the row; a closed organization's data
// cannot be changed (Code.gs refuses the writes).
import { tr, getLang } from '../i18n.js';
import { $, esc, icon, toast, confirmBox } from '../ui.js';
import { cached, call, peek, remember, forget } from '../api.js';
import { CAMPUSES, CATEGORIES, ORGS } from '../masterdata.js';

let query = '';
let show = 'all';

export async function render(page, ctx) {
  page.innerHTML = `<h1 class="grad">${tr('Organizations', 'องค์กร')}</h1><p class="muted" id="sum">&nbsp;</p>
    <div class="toolbar"><input type="search" id="q" placeholder="${esc(tr('Search name, RefCode or person in charge', 'ค้นหาชื่อ รหัส หรือผู้รับผิดชอบ'))}" value="${esc(query)}">
      <div class="seg filter" role="radiogroup" aria-label="${esc(tr('Show', 'แสดง'))}" id="show"></div></div>
    <div class="group scroll-x"><table class="grid-table quotes">
      <thead><tr><th>${tr('Category', 'หมวด')}</th><th>${tr('Name', 'ชื่อ')}</th><th>${tr('RefCode', 'รหัส')}</th>
        <th class="r">${tr('Projects', 'โครงการ')}</th><th>${tr('Person in charge', 'ผู้รับผิดชอบ')}</th><th class="r">${tr('Pending approvals', 'รออนุมัติ')}</th><th>${tr('Status', 'สถานะ')}</th></tr></thead>
      <tbody id="rows"></tbody></table></div>`;

  let boot = null;
  const label = (list, code) => { const x = list.find((i) => i.code === code); return x ? (getLang() === 'th' ? x.th : x.en) : code; };
  const stat = (ref) => boot.orgs[ref] || (boot.orgs[ref] = { count: 0, pending: 0, people: [] });

  const pic = (o) => {
    const people = stat(o.code).people;
    if (!ctx.console) return people.map((p) => esc(ctx.nameOf(p))).join('<br>') || '<span class="muted">—</span>';
    const users = boot.users || [];
    return `<select class="pic" data-pic="${esc(o.code)}" aria-label="${esc(tr('Person in charge', 'ผู้รับผิดชอบ'))}">
      <option value="">—</option>${[...users].sort((a, b) => ctx.nameOf(a).localeCompare(ctx.nameOf(b))).map((u) => `<option value="${esc(u)}" ${people[0] === u ? 'selected' : ''}>${esc(ctx.nameOf(u))}</option>`).join('')}</select>`;
  };

  const closedOf = (ref) => boot.closed?.[ref] || null;
  const status = (o) => {
    const c = closedOf(o.code);
    const label = c ? `${icon('lock', 13)}${tr('Closed', 'ปิดแล้ว')}` : tr('Unclosed', 'ยังไม่ปิด');
    if (!ctx.console) return `<span class="pill ${c ? 'closed' : 'green'}">${label}</span>`;
    return `<button class="pill btn-pill ${c ? 'closed' : 'green'}" data-close="${esc(o.code)}" title="${esc(c ? tr('Click to reopen', 'คลิกเพื่อเปิดอีกครั้ง') : tr('Click to close', 'คลิกเพื่อปิด'))}">${label}</button>`;
  };

  const SHOW = {
    all: { label: () => tr('All', 'ทั้งหมด'), keep: () => true },
    pending: { label: () => tr('Pending approval', 'รออนุมัติ'), keep: (s) => s.pending > 0 },
    done: { label: () => tr('Done', 'ครบแล้ว'), keep: (s) => s.count > 0 && !s.pending },
  };
  const paint = () => {
    const removed = new Set(boot.removed);
    const rows = ORGS.filter((o) => !removed.has(o.code));
    const pending = rows.filter((o) => stat(o.code).pending > 0).length;
    $('#sum', page).innerHTML = tr('{n} organizations · ', '{n} องค์กร · ', { n: rows.length })
      + (pending ? `<b class="warn">${tr('{n} with unfinished approvals', '{n} องค์กรยังอนุมัติไม่ครบ', { n: pending })}</b>` : tr('all approvals finished', 'อนุมัติครบทุกองค์กร'));
    $('#show', page).innerHTML = Object.entries(SHOW).map(([k, f]) => `<label><input type="radio" name="show" value="${k}" ${show === k ? 'checked' : ''}><span>${esc(f.label())}
      <small>${rows.filter((o) => f.keep(stat(o.code))).length}</small></span></label>`).join('');
    const q = query.trim().toLowerCase();
    const hit = (o) => SHOW[show].keep(stat(o.code)) && (!q || `${o.en} ${o.code} ${stat(o.code).people.map((p) => `${p} ${ctx.nameOf(p)}`).join(' ')}`.toLowerCase().includes(q));
    const narrowed = q || show !== 'all';
    $('#rows', page).innerHTML = CAMPUSES.map((c) => {
      const list = rows.filter((o) => o.campus === c.code && hit(o)).sort((a, b) => a.code.localeCompare(b.code));
      const pend = list.filter((o) => stat(o.code).pending > 0).length;
      const head = `<tr class="grp"><th colspan="7">${esc(c.en.toUpperCase())}<span class="muted"> · ${esc(c.th)} · ${list.length}</span>
        ${pend ? `<span class="pill orange">${tr('{n} pending', 'รอ {n}', { n: pend })}</span>` : ''}</th></tr>`;
      // Searching or filtering: campuses with no match are left out.
      if (!list.length) return narrowed ? '' : `${head}<tr class="none"><td colspan="7" class="muted">${tr('No organizations', 'ไม่มีองค์กร')}</td></tr>`;
      return head + list.map((o) => {
        const s = stat(o.code);
        return `<tr data-ref="${esc(o.code)}" tabindex="0">
        <td class="mono" title="${esc(label(CATEGORIES, o.category))}">${esc(o.category)}</td>
        <td>${esc(o.en)}</td><td class="mono">${esc(o.code)}</td>
        <td class="r">${s.count || '—'}</td>
        <td class="small">${pic(o)}</td>
        <td class="r">${s.count ? (s.pending ? `<span class="pill orange">${s.pending} / ${s.count}</span>` : `<span class="pill green">${tr('Done', 'ครบ')}</span>`) : '—'}</td>
        <td>${status(o)}</td></tr>`;
      }).join('');
    }).join('') || `<tr class="none"><td colspan="7" class="muted">${tr('No organizations match', 'ไม่มีองค์กรที่ตรงกัน')}</td></tr>`;
  };

  // This device's last copy first (instant), then the fresh one if it differs.
  const kept = peek('boot');
  if (kept) { boot = ctx.boot = kept; paint(); }
  try {
    const fresh = await cached('boot');
    if (!kept || JSON.stringify(fresh) !== JSON.stringify(kept)) { boot = ctx.boot = fresh; paint(); }
  } catch (e) {
    if (!kept) throw e;
    toast(e.message, 'bad');
  }

  $('#q', page).oninput = (e) => { query = e.target.value; paint(); };
  $('#show', page).onchange = (e) => { show = e.target.value; paint(); };
  const rowsEl = $('#rows', page);
  const open = (e) => {
    if (e.target.closest('select, button')) return;
    const r = e.target.closest('tr[data-ref]');
    if (r) ctx.go(`#/org/${encodeURIComponent(r.dataset.ref)}`);
  };
  rowsEl.onclick = async (e) => {
    const b = e.target.closest('[data-close]');
    if (!b) return open(e);
    const ref = b.dataset.close;
    const closing = !closedOf(ref);
    const ok = await confirmBox(closing ? tr('Close {ref}?', 'ปิด {ref}?', { ref }) : tr('Reopen {ref}?', 'เปิด {ref} อีกครั้ง?', { ref }),
      closing ? tr('No one can change its data until it is reopened.', 'จะไม่มีใครแก้ไขข้อมูลได้จนกว่าจะเปิดอีกครั้ง') : tr('Its data can be changed again.', 'จะแก้ไขข้อมูลได้อีกครั้ง'),
      { ok: closing ? tr('Close', 'ปิด') : tr('Reopen', 'เปิดอีกครั้ง'), danger: closing });
    if (!ok) return;
    b.disabled = true;
    try {
      const out = await call('setOrgStatus', { ref, closed: closing });
      boot.closed = { ...(boot.closed || {}) };
      if (out.closed) boot.closed[ref] = out.closed; else delete boot.closed[ref];
      remember('boot', {}, boot);
      forget('org');
      paint();
      toast(closing ? tr('{ref} is closed', '{ref} ปิดแล้ว', { ref }) : tr('{ref} is reopened', '{ref} เปิดอีกครั้งแล้ว', { ref }));
    } catch (err) {
      toast(err.message, 'bad');
      b.disabled = false;
    }
  };
  rowsEl.onkeydown = (e) => { if (e.key === 'Enter') open(e); };

  // Hovering a row for a moment starts loading that organization, so the click opens it faster.
  let hover;
  rowsEl.onpointerover = (e) => {
    const r = e.target.closest('tr[data-ref]');
    clearTimeout(hover);
    if (r && stat(r.dataset.ref).count) hover = setTimeout(() => cached('org', { ref: r.dataset.ref }).catch(() => {}), 250);
  };

  rowsEl.onchange = async (e) => {
    const sel = e.target.closest('select[data-pic]');
    if (!sel) return;
    const ref = sel.dataset.pic;
    sel.disabled = true;
    try {
      const out = await call('setInCharge', { ref, email: sel.value });
      stat(ref).people = out.people;
      remember('boot', {}, boot);
      forget('org');
      toast(sel.value ? tr('{who} is in charge of {ref}', '{who} รับผิดชอบ {ref}', { who: sel.value, ref }) : tr('No one in charge of {ref}', '{ref} ไม่มีผู้รับผิดชอบ', { ref }));
    } catch (err) {
      toast(err.message, 'bad');
      sel.value = stat(ref).people[0] || '';
    }
    sel.disabled = false;
  };
}
