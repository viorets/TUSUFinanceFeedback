// Timeline (v2.12): the round as a Gantt chart at the top of the Organizations tab. Left: the steps grouped by phase,
// with start, end and days. Right: a month / week scale, week lines, one coloured bar per step (each its own colour,
// apart from the CI teal), a group bar per phase, a "today" line and, for Pitching and Re-submission, the projects
// finished as a darker fill. Dates: boot.dates (Settings "Phase Dates", edited in the admin console; keys and
// defaults in TIMELINE in Code.gs). Folded or not is kept on this device.
import { tr, locale } from './i18n.js';
import { esc } from './ui.js';

// Keep in step with TIMELINE in Code.gs. [key, label, colour]; GROUPS put the steps under their phase.
export const STEPS = [
  ['systemTest', () => tr('System test', 'ทดสอบระบบ'), '#8b5cf6'],
  ['submission', () => tr('Project submission #2', 'ส่งโครงการ ครั้งที่ 2'), '#3b82f6'],
  ['pitching', () => tr('Pitching #2 (Pre-pitch and Pitching)', 'นำเสนอ ครั้งที่ 2 (ก่อนนำเสนอและนำเสนอ)'), '#f59e0b'],
  ['mailing', () => tr('Mailing #2', 'ส่งอีเมล ครั้งที่ 2'), '#ef4444'],
  ['resubCampus', () => tr('Campus check', 'ศูนย์ตรวจ'), '#ec4899'],
  ['resubCentre', () => tr('Centre team check', 'ทีมส่วนกลางตรวจ'), '#22a35a'],
  ['council', () => tr('Documentation and council submission', 'จัดทำเอกสารและส่งสภา'), '#6366f1'],
];
const GROUPS = [
  [() => tr('Preparation', 'เตรียมระบบ'), ['systemTest']],
  [() => tr('Submission', 'การส่งโครงการ'), ['submission']],
  [() => tr('Pitching', 'การนำเสนอ'), ['pitching', 'mailing']],
  [() => tr('Re-submission', 'การส่งแก้ไข'), ['resubCampus', 'resubCentre']],
  [() => tr('Council', 'สภา'), ['council']],
];

const DAY = 864e5;
const at = (iso) => Date.UTC(...iso.split('-').map((n, i) => (i === 1 ? n - 1 : +n)));
const todayIso = () => new Date().toLocaleDateString('en-CA', { timeZone: 'Asia/Bangkok' });
const fmt = (t, o) => new Date(t).toLocaleDateString(locale(), { timeZone: 'UTC', ...o });
const short = (iso) => fmt(at(iso), { day: 'numeric', month: 'short' });
export const range = (d) => (d.from === d.to ? fmt(at(d.from), { day: 'numeric', month: 'short', year: 'numeric' })
  : `${short(d.from)} – ${fmt(at(d.to), { day: 'numeric', month: 'short', year: 'numeric' })}`);
const days = (d) => Math.round((at(d.to) - at(d.from)) / DAY) + 1;

let open = true;
try { open = localStorage.getItem('timelineOpen') !== '0'; } catch { /* private mode */ }

// dates: { key: { from, to } }; progress: { key: { done, all } } (Pitching and Re-submission).
export function timeline(dates, progress = {}) {
  const steps = STEPS.filter(([k]) => dates?.[k]);
  if (!steps.length) return '';
  const today = todayIso();
  const now = at(today);
  // From the Monday on or before the first day (or today) to the Monday after the last.
  const first = Math.min(now, ...steps.map(([k]) => at(dates[k].from)));
  const last = Math.max(now, ...steps.map(([k]) => at(dates[k].to)));
  const t0 = first - ((new Date(first).getUTCDay() + 6) % 7) * DAY;
  const t1 = last + (7 - ((new Date(last).getUTCDay() + 6) % 7)) * DAY;
  const pct = (t) => `${(((t - t0) / (t1 - t0)) * 100).toFixed(3)}%`;
  const len = (a, b) => `${(((b - a) / (t1 - t0)) * 100).toFixed(3)}%`;

  const weeks = [];
  for (let t = t0; t < t1; t += 7 * DAY) weeks.push(t);
  const months = [];
  for (let t = t0; t < t1;) {
    const d = new Date(t);
    const next = Math.min(t1, Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + 1, 1));
    months.push(`<span class="g-month" style="left:${pct(t)};width:${len(t, next)}">${esc(fmt(t, { month: 'long', year: 'numeric' }))}</span>`);
    t = next;
  }
  const scale = `<div class="g-scale">${months.join('')}${weeks.map((t) => `<span class="g-week" style="left:${pct(t)};width:${len(t, t + 7 * DAY)}">${esc(fmt(t, { day: 'numeric' }))}</span>`).join('')}</div>`;
  const grid = `<div class="g-grid">${weeks.map((t) => `<i style="left:${pct(t)}"></i>`).join('')}<b class="g-today" style="left:${pct(now + DAY / 2)}" title="${esc(tr('Today', 'วันนี้'))}"></b></div>`;

  const byKey = Object.fromEntries(steps.map((s) => [s[0], s]));
  const rows = GROUPS.map(([glabel, keys]) => {
    const ks = keys.filter((k) => byKey[k]);
    if (!ks.length) return '';
    const g = { from: ks.map((k) => dates[k].from).sort()[0], to: ks.map((k) => dates[k].to).sort().at(-1) };
    const head = `<div class="g-row g-group"><span class="g-name">${esc(glabel())}</span><span>${esc(short(g.from))}</span><span>${esc(short(g.to))}</span><span class="num">${days(g)}</span>
      <span class="g-track"><span class="g-sum" style="left:${pct(at(g.from))};width:${len(at(g.from), at(g.to) + DAY)}"></span></span></div>`;
    return head + ks.map((k) => {
      const [, label, colour] = byKey[k];
      const d = dates[k];
      const state = today > d.to ? 'past' : today >= d.from ? 'now' : 'next';
      const p = progress[k];
      const share = p?.all ? p.done / p.all : null;
      const tip = [label(), range(d), share != null ? tr('{n} / {all} projects finished', 'ครบแล้ว {n} / {all} โครงการ', { n: p.done, all: p.all }) : ''].filter(Boolean).join(' · ');
      const end = (at(d.to) + DAY - t0) / (t1 - t0);
      const wide = (at(d.to) + DAY - at(d.from)) / (t1 - t0) > 0.12; // room for the label inside the bar
      const pctText = share != null ? ` ${Math.round(share * 100)}%` : '';
      // A short bar keeps its label beside it: on the right, or on the left near the chart's end.
      const out = wide ? '' : end > 0.72
        ? `<em class="g-out l" style="right:calc(${(100 - ((at(d.from) - t0) / (t1 - t0)) * 100).toFixed(3)}% + 6px)">${esc(label())}${pctText}</em>`
        : `<em class="g-out" style="left:calc(${pct(at(d.to) + DAY)} + 6px)">${esc(label())}${pctText}</em>`;
      return `<div class="g-row g-step ${state}"><span class="g-name">${esc(label())}</span><span>${esc(short(d.from))}</span><span>${esc(short(d.to))}</span><span class="num">${days(d)}</span>
        <span class="g-track"><span class="g-bar" style="--c:${colour};left:${pct(at(d.from))};width:${len(at(d.from), at(d.to) + DAY)}" title="${esc(tip)}" tabindex="0" aria-label="${esc(tip)}">${share != null ? `<i style="width:${Math.round(share * 100)}%"></i>` : ''}${wide ? `<em>${esc(label())}${pctText}</em>` : ''}</span>${out}</span></div>`;
    }).join('');
  }).join('');

  return `<details class="group timeline" ${open ? 'open' : ''}><summary>${tr('Timeline', 'ไทม์ไลน์')} <small class="muted">${esc(tr('Today: {d}', 'วันนี้: {d}', { d: range({ from: today, to: today }) }))}</small></summary>
    <div class="gantt"><div class="g-row g-headrow"><span class="g-name">${tr('Phase', 'ช่วง')}</span><span>${tr('Start', 'เริ่ม')}</span><span>${tr('End', 'สิ้นสุด')}</span><span class="num">${tr('Days', 'วัน')}</span>${scale}</div>
      <div class="g-body">${grid}${rows}</div></div></details>`;
}

// Remember folded / open after the chart is on the page.
export function bindTimeline(root) {
  const d = root.querySelector('details.timeline');
  if (d) d.ontoggle = () => { open = d.open; try { localStorage.setItem('timelineOpen', open ? '1' : '0'); } catch { /* this visit only */ } };
}
