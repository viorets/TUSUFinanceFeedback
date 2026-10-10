import { firebaseConfig, SCRIPT_URL } from './config.js';
import { tr, getLang, setLang, setTexts } from './i18n.js';
import { $, esc, icon, toast, dialog, errorBox, skeleton } from './ui.js';

const root = $('#app');

// Night mode: follows the device until the moon / sun button picks one; the choice is kept on this device.
// index.html applies the kept choice before the first paint.
const dark = () => document.documentElement.dataset.theme === 'dark'
  || (!document.documentElement.dataset.theme && matchMedia('(prefers-color-scheme: dark)').matches);
const themeColor = () => document.querySelector('meta[name=theme-color]')?.setAttribute('content', dark() ? '#111416' : '#dfe9ec');
function toggleTheme() {
  const next = dark() ? 'light' : 'dark';
  document.documentElement.dataset.theme = next;
  try { localStorage.setItem('theme', next); } catch { /* private mode: this visit only */ }
  themeColor();
  $('#theme').innerHTML = icon(dark() ? 'sun' : 'moon', 18);
}
themeColor();
matchMedia('(prefers-color-scheme: dark)').addEventListener('change', () => { themeColor(); const b = $('#theme'); if (b) b.innerHTML = icon(dark() ? 'sun' : 'moon', 18); });
const TITLE = 'Pitch feedback#2';

if (!firebaseConfig.apiKey || !SCRIPT_URL) {
  root.innerHTML = `<div class="signin"><div class="card"><img src="icon.svg" alt=""><h2 style="margin-top:14px">Setup needed</h2>
    <p class="muted">Fill in <code>firebaseConfig</code> and <code>SCRIPT_URL</code> in <code>public/js/config.js</code>. See README.md.</p></div></div>`;
  throw new Error('config missing');
}

const { onUser, signIn, logOut } = await import('./firebase.js');
const api = await import('./api.js');

const VIEWS = {
  list: () => import('./views/list.js'),
  org: () => import('./views/org.js'),
  admin: () => import('./views/admin.js'),
  dashboard: () => import('./views/dashboard.js'),
  control: () => import('./views/control.js'),
};

const ctx = {
  user: null, me: null, boot: null,
  // Admin console: Admins and Chiefs (only Admins see the edit log; the server enforces both).
  get console() { return ctx.me?.role === 'Admin' || ctx.me?.role === 'Chief'; },
  get canEdit() { return ['Admin', 'Chief', 'Staff'].includes(ctx.me?.role); },
  // Owner = the script owner: the Control tab (site text, required questions). The server checks too.
  get owner() { return !!ctx.me?.owner; },
  // Display name set in the admin console, else the email.
  nameOf: (email) => ctx.boot?.names?.[email] || email,
  go: (h) => { location.hash = h; },
};
let current = null; // the mounted view module (may export isDirty)
let lastHash = location.hash;

function parse(hash) {
  const [, a = '', b = '', c = ''] = hash.replace(/^#/, '').split('/');
  if (a === 'org' && b) return { view: 'org', ref: decodeURIComponent(b), tab: c || 'prepitch' };
  if (a === 'admin') return { view: 'admin' };
  if (a === 'dashboard') return { view: 'dashboard' };
  if (a === 'control') return { view: 'control' };
  return { view: 'list' };
}

function shell(route) {
  const tab = (href, ic, label, on) => `<a href="${href}" class="${on ? 'on' : ''}">${icon(ic, 18)}<span>${label}</span></a>`;
  const roleLabel = { Admin: tr('Admin', 'ผู้ดูแล'), Chief: tr('Chief', 'หัวหน้า'), Staff: tr('Staff', 'เจ้าหน้าที่') }[ctx.me.role] || ctx.me.role;
  return `<div class="chrome"><header class="topbar">
    <a class="brand" href="#/"><img class="brand-logo" src="icon.svg" alt=""><span class="grow"><b>${TITLE}</b><span>TUSU Budget Pitching</span></span></a>
    <nav class="topnav">
      <button id="theme" title="${tr('Night mode', 'โหมดกลางคืน')}" aria-label="${tr('Night mode', 'โหมดกลางคืน')}">${icon(dark() ? 'sun' : 'moon', 18)}</button>
      <button id="lang" title="${tr('Language', 'ภาษา')}">${icon('globe', 18)}<span class="lbl">${getLang() === 'th' ? 'EN' : 'ไทย'}</span></button>
      <div class="acct"><button id="acct" aria-haspopup="menu" aria-expanded="false" title="${esc(ctx.me.email)}">${icon('user', 18)}<span class="lbl acct-mail">${esc(ctx.nameOf(ctx.me.email))}</span>${icon('down', 14)}</button></div>
    </nav></header>
    <nav class="tabs" aria-label="${tr('Sections', 'เมนู')}"><div class="tabs-in">
      ${tab('#/dashboard', 'chart', tr('Dashboard', 'แดชบอร์ด'), route.view === 'dashboard')}
      ${tab('#/', 'grid', tr('Organizations', 'องค์กร'), route.view === 'list' || route.view === 'org')}
      ${ctx.console ? tab('#/admin', 'shield', tr('Admin console', 'คอนโซลผู้ดูแล'), route.view === 'admin') : ''}
      ${ctx.owner ? tab('#/control', 'sliders', tr('Control', 'ควบคุม'), route.view === 'control') : ''}
      <span class="pill blue role">${esc(roleLabel)}</span>
    </div></nav></div>
    <main class="page wide" id="page">${skeleton()}</main>`;
}

function accountMenu(btn) {
  const box = btn.parentElement;
  if (box.querySelector('.menu')) { box.querySelector('.menu').remove(); btn.setAttribute('aria-expanded', 'false'); return; }
  const m = document.createElement('div');
  m.className = 'menu'; m.setAttribute('role', 'menu');
  m.innerHTML = `<div class="menu-who"><b>${esc(ctx.boot?.names?.[ctx.me.email] || ctx.user.displayName || '')}</b><span>${esc(ctx.me.email)}</span></div>
    <button role="menuitem" data-a="switch">${icon('user', 18)}${tr('Switch account', 'เปลี่ยนบัญชี')}</button>
    <button role="menuitem" data-a="out" class="danger">${icon('out', 18)}${tr('Sign out', 'ออกจากระบบ')}</button>`;
  box.append(m);
  btn.setAttribute('aria-expanded', 'true');
  m.onclick = async (e) => {
    const a = e.target.closest('[data-a]')?.dataset.a;
    if (!a) return;
    m.remove();
    if (!(await confirmLeave())) return;
    api.clearKept();
    await logOut();
    if (a === 'switch') signIn().catch(signInError);
  };
}
document.addEventListener('click', (e) => { if (!e.target.closest('.acct')) document.querySelector('.acct .menu')?.remove(); });
// The floating top bar gets a little smaller once the page scrolls.
addEventListener('scroll', () => document.querySelector('.chrome')?.classList.toggle('scrolled', scrollY > 24), { passive: true });

async function render() {
  const route = parse(location.hash);
  if ((route.view === 'admin' && !ctx.console) || (route.view === 'control' && !ctx.owner)) { ctx.go('#/'); return; }
  root.innerHTML = shell(route);
  $('#lang').onclick = async () => { if (!(await confirmLeave())) return; setLang(getLang() === 'th' ? 'en' : 'th'); current = null; render(); };
  $('#acct').onclick = (e) => accountMenu(e.currentTarget);
  $('#theme').onclick = toggleTheme;
  // Start loading the admin console when the pointer reaches its tab.
  document.querySelector('a[href="#/admin"]')?.addEventListener('pointerenter', () => api.cached('admin').catch(() => {}), { once: true });
  const page = $('#page');
  try {
    current = await VIEWS[route.view]();
    await current.render(page, ctx, route);
  } catch (e) {
    console.error(e);
    page.innerHTML = `<div class="empty">${icon('alert', 28)}<p>${esc(e.message || e)}</p><a class="btn" href="#/">${tr('Back to organizations', 'กลับไปที่รายชื่อองค์กร')}</a></div>`;
  }
}

// Leaving a page (or tab, language, account) with unsaved edits: a window naming each unsaved project, with
// Save and leave / Leave without saving / Stay. Closing the tab, refreshing or typing a new URL is not guarded
// (option 2 in the plan): those edits are lost without a warning.
async function confirmLeave() {
  if (!current?.isDirty?.()) return true;
  const list = current.unsaved?.() || [];
  const canSave = list.length > 0 && !!current.canSaveAll?.();
  const html = list.length
    ? `<p>${tr('These projects have unsaved changes:', 'โครงการเหล่านี้มีการแก้ไขที่ยังไม่ได้บันทึก:')}</p>
      <ul class="unsaved">${list.map((p) => `<li><span class="code">${esc(p.ref)}</span><span>${esc(p.name || '—')}</span></li>`).join('')}</ul>`
    : `<p>${tr('Your unsaved changes will be lost.', 'การแก้ไขที่ยังไม่บันทึกจะหายไป')}</p>`;
  const buttons = [
    { v: 'leave', label: tr('Leave without saving', 'ออกโดยไม่บันทึก'), cls: 'danger' },
    { v: 'stay', label: tr('Stay', 'อยู่ต่อ'), cls: canSave ? '' : 'primary' },
  ];
  if (canSave) buttons.push({ v: 'save', label: tr('Save and leave', 'บันทึกแล้วออก'), cls: 'primary' });
  const v = await dialog(tr('Unsaved changes', 'มีการแก้ไขที่ยังไม่บันทึก'), html, buttons, { none: 'stay' });
  if (v === 'leave') return true;
  if (v !== 'save') return false;
  try { await current.saveAll(); return true; } catch (e) { await errorBox(e.message); return false; }
}

window.addEventListener('hashchange', async () => {
  if (location.hash === lastHash) return;
  // Switching tabs inside one organization keeps the view; it asks its own tab about unsaved changes.
  const a = parse(lastHash); const b = parse(location.hash);
  if (a.view === 'org' && b.view === 'org' && a.ref === b.ref && current?.switchTab) {
    if (!(await confirmLeave())) { history.replaceState(null, '', lastHash || '#/'); return; }
    lastHash = location.hash;
    current.switchTab(b.tab);
    return;
  }
  if (!(await confirmLeave())) { history.replaceState(null, '', lastHash || '#/'); return; }
  current = null;
  lastHash = location.hash;
  render();
});

const signInError = (e) => { if (e.code !== 'auth/popup-closed-by-user') toast(e.message, 'bad'); };

function card(inner) { root.innerHTML = `<div class="signin"><div class="card"><img src="icon.svg" alt="">${inner}</div></div>`; }

function signInScreen() {
  card(`<h1 class="grad" style="margin-top:18px">${TITLE}</h1>
    <p class="muted">${tr('Thammasat University Student Union. Sign in with your Google account to review budget pitches.', 'องค์การนักศึกษามหาวิทยาลัยธรรมศาสตร์ ลงชื่อเข้าใช้ด้วยบัญชี Google เพื่อตรวจการเสนองบประมาณ')}</p>
    <button class="btn primary lg" id="go" style="width:100%;margin-top:14px">${tr('Sign in with Google', 'ลงชื่อเข้าใช้ด้วย Google')}</button>
    <button class="btn plain" id="lang2" style="margin-top:12px">${getLang() === 'th' ? 'English' : 'ภาษาไทย'}</button>`);
  $('#go').onclick = () => signIn().catch(signInError);
  $('#lang2').onclick = () => { setLang(getLang() === 'th' ? 'en' : 'th'); signInScreen(); };
}

onUser(async (user) => {
  ctx.user = user;
  current = null;
  if (!user) { signInScreen(); return; }
  api.setWho(user.email);
  // Show this device's last copy at once; the fresh answer follows (views repaint from it).
  const kept = api.peek('boot');
  if (kept) { ctx.boot = kept; ctx.me = kept.me; setTexts(kept.texts); lastHash = location.hash; render(); }
  else root.innerHTML = `<div class="page">${skeleton()}</div>`;
  try {
    ctx.boot = await api.cached('boot');
    const textsChanged = JSON.stringify(ctx.boot.texts || {}) !== JSON.stringify(kept?.texts || {});
    setTexts(ctx.boot.texts);
    const roleChanged = ctx.boot.me.role !== kept?.me.role || !!ctx.boot.me.owner !== !!kept?.me.owner;
    ctx.me = ctx.boot.me;
    // Same role: repaint only for new site text, and never over unsaved edits.
    if (!roleChanged && (!textsChanged || current?.isDirty?.())) return;
  } catch (e) {
    api.clearKept();
    const none = e.code === 'NO_ACCESS';
    card(`<h2 style="margin-top:14px">${none ? tr('No access yet', 'ยังไม่มีสิทธิ์เข้าใช้') : tr('Couldn’t connect', 'เชื่อมต่อไม่ได้')}</h2>
      <p class="muted">${none ? tr('{email} is not on the user list. Ask an Admin to add you.', '{email} ยังไม่อยู่ในรายชื่อผู้ใช้ กรุณาขอให้ผู้ดูแลเพิ่มสิทธิ์', { email: esc(user.email) }) : esc(e.message)}</p>
      <div style="display:flex;gap:10px;justify-content:center;margin-top:14px"><button class="btn" id="retry">${tr('Try again', 'ลองใหม่')}</button><button class="btn danger" id="out">${tr('Sign out', 'ออกจากระบบ')}</button></div>`);
    $('#retry').onclick = () => location.reload();
    $('#out').onclick = () => logOut();
    return;
  }
  lastHash = location.hash;
  current = null;
  render();
});
