// Ribbon for the Word and Excel editors (WordCraft / GridCraft style), inside the floating glass toolbar: a row of tabs
// and, under it, the open tab's groups of buttons, selects and colour boxes. Built from a plain spec; the editor wires
// the behaviour by id. Buttons never take the focus away from the document (pointerdown/mousedown preventDefault), so a
// text selection or the active cell stays put while a button is pressed.
const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');
const svg = (d) => `<svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${d}</svg>`;
export const RI = {
  cut: svg('<circle cx="6" cy="6" r="3"/><circle cx="6" cy="18" r="3"/><path d="M20 4 8.1 15.9M14.5 14.5 20 20M8.1 8.1 12 12"/>'),
  copy: svg('<rect x="9" y="9" width="12" height="12" rx="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/>'),
  paste: svg('<rect x="8" y="2" width="8" height="4" rx="1"/><path d="M16 4h2a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h2"/>'),
  bold: '<b style="font:700 15px Georgia,serif">B</b>', italic: '<i style="font:italic 15px Georgia,serif">I</i>',
  underline: '<u style="font:15px Georgia,serif">U</u>', strike: '<s style="font:15px Georgia,serif">S</s>',
  color: svg('<path d="M5 19h14"/><path d="m7 15 5-11 5 11M9 11h6"/>'),
  fill: svg('<path d="m19 11-8-8-8.5 8.5a2 2 0 0 0 0 2.8l5.7 5.7a2 2 0 0 0 2.8 0Z"/><path d="M5 2l5 5M2 13h15M21 15s1.5 2 1.5 3a1.5 1.5 0 0 1-3 0c0-1 1.5-3 1.5-3Z"/>'),
  hl: svg('<path d="m9 11-6 6v3h9l3-3"/><path d="m22 12-4.6 4.6a2 2 0 0 1-2.8 0l-5.2-5.2a2 2 0 0 1 0-2.8L14 4"/>'),
  left: svg('<path d="M3 6h18M3 12h12M3 18h16"/>'), center: svg('<path d="M3 6h18M6 12h12M4 18h16"/>'),
  right: svg('<path d="M3 6h18M9 12h12M5 18h16"/>'), justify: svg('<path d="M3 6h18M3 12h18M3 18h18"/>'),
  top: svg('<path d="M4 3h16M12 21V7M8 11l4-4 4 4"/>'), middle: svg('<path d="M4 12h16M12 3v6M12 15v6M9 6l3 3 3-3M9 18l3-3 3 3"/>'),
  bottom: svg('<path d="M4 21h16M12 3v14M8 13l4 4 4-4"/>'),
  wrap: svg('<path d="M3 6h18M3 12h15a3 3 0 0 1 0 6h-4M3 18h7"/><path d="m16 16-2 2 2 2"/>'),
  merge: svg('<rect x="3" y="5" width="18" height="14" rx="1"/><path d="M8 12h8M10 10l-2 2 2 2M14 10l2 2-2 2"/>'),
  table: svg('<rect x="3" y="4" width="18" height="16" rx="1.5"/><path d="M3 9.5h18M3 15h18M9.5 9.5V20"/><path d="M4 5h16v4H4z" fill="currentColor" stroke="none" opacity=".35"/>'),
  borders: svg('<rect x="3" y="3" width="18" height="18" rx="1"/><path d="M3 12h18M12 3v18"/>'),
  pct: svg('<path d="M19 5 5 19"/><circle cx="6.5" cy="6.5" r="2.5"/><circle cx="17.5" cy="17.5" r="2.5"/>'),
  comma: '<span style="font:700 14px system-ui">,000</span>',
  decUp: '<span style="font:600 11px system-ui">.0→.00</span>', decDown: '<span style="font:600 11px system-ui">.00→.0</span>',
  rowAbove: svg('<rect x="3" y="12" width="18" height="8" rx="1"/><path d="M12 3v6M9 6h6"/>'), rowBelow: svg('<rect x="3" y="4" width="18" height="8" rx="1"/><path d="M12 15v6M9 18h6"/>'),
  colLeft: svg('<rect x="12" y="3" width="8" height="18" rx="1"/><path d="M3 12h6M6 9v6"/>'), colRight: svg('<rect x="4" y="3" width="8" height="18" rx="1"/><path d="M15 12h6M18 9v6"/>'),
  rowDel: svg('<rect x="3" y="8" width="18" height="8" rx="1"/><path d="m9 10 6 4M15 10l-6 4"/>'), colDel: svg('<rect x="8" y="3" width="8" height="18" rx="1"/><path d="m10 9 4 6M14 9l-4 6"/>'),
  sortAZ: '<span style="font:700 11px system-ui;line-height:1">A→Z</span>', sortZA: '<span style="font:700 11px system-ui;line-height:1">Z→A</span>',
  filter: svg('<path d="M3 4h18l-7 8.5V19l-4 2v-8.5Z"/>'),
  freeze: svg('<rect x="3" y="3" width="18" height="18" rx="1"/><path d="M3 9h18M9 3v18"/>'),
  fx: '<span style="font:italic 700 14px Georgia,serif">fx</span>', sum: '<span style="font:700 16px Georgia,serif">Σ</span>',
  find: svg('<circle cx="11" cy="11" r="7"/><path d="m21 21-4.3-4.3"/>'),
  bullets: svg('<path d="M9 6h12M9 12h12M9 18h12"/><circle cx="4" cy="6" r="1"/><circle cx="4" cy="12" r="1"/><circle cx="4" cy="18" r="1"/>'),
  numbers: svg('<path d="M10 6h11M10 12h11M10 18h11M4 6h1v4M4 10h2M6 18H4c0-1 2-2 2-3s-1-1.5-2-1"/>'),
  indent: svg('<path d="M3 6h18M11 12h10M11 18h10M3 10l4 3-4 3"/>'), outdent: svg('<path d="M3 6h18M11 12h10M11 18h10M7 10l-4 3 4 3"/>'),
  spacing: svg('<path d="M10 6h11M10 12h11M10 18h11M4 4v16M2 6l2-2 2 2M2 18l2 2 2-2"/>'),
  textbox: svg('<rect x="3" y="5" width="18" height="14" rx="1"/><path d="M8 9h8M12 9v6"/>'),
  pen: svg('<path d="M12 20h9"/><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4Z"/>'),
  line: svg('<path d="M5 19 19 5"/>'), arrow: svg('<path d="M5 19 19 5M10 5h9v9"/>'), rect: svg('<rect x="4" y="6" width="16" height="12" rx="1"/>'),
  select: svg('<path d="m4 4 7 17 2.5-7.5L21 11Z"/>'),
  open: svg('<path d="M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v1"/><path d="M3 7v10a2 2 0 0 0 2 2h13l3-9H8l-3 9"/>'),
  newDoc: svg('<path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/><path d="M14 3v5h5M12 11v6M9 14h6"/>'),
  save: svg('<path d="M5 3h11l3 3v13a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2z"/><path d="M8 3v5h7V3M8 21v-7h8v7"/>'),
  dl: svg('<path d="M12 3v12M7 10l5 5 5-5"/><path d="M4 17v2a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-2"/>'),
  pdf: '<span style="font:700 10px system-ui">PDF</span>', txt: '<span style="font:700 10px system-ui">TXT</span>',
  md: '<span style="font:700 10px system-ui">MD</span>', html: '<span style="font:700 9px system-ui">HTML</span>',
};

// spec: [{ tab: 'Home', groups: [{ label: 'Font', items: [item, …] }] }]. item:
//   { btn: id, icon, title, toggle? }  { select: id, title, options: [[value, label]] | [label], width? }
//   { color: id, title, icon }  (a colour box with a small icon above it; the editor pairs it with HEX + scrub)
//   { input: id, title, width?, placeholder? }   { html: '…' } (anything else)
// `start`: the tab open at first (the File tab comes first, as in Office, but the editor opens on Home).
// Thai labels for the ribbons, for a page shown in Thai (the Feedback site): setRibbonLang(t) with its t(en, th).
// Labels not in the list stay in English.
let say = (en) => en;
export const setRibbonLang = (t) => { say = (en) => (TH_LABELS[en] ? t(en, TH_LABELS[en]) : en); };
const TH_LABELS = {
  "Align left": "จัดชิดซ้าย",
  "Align right": "จัดชิดขวา",
  "Alignment": "การจัดแนว",
  "Arrow": "ลูกศร",
  "AutoSum": "ผลรวมอัตโนมัติ",
  "Bold (Ctrl+B)": "ตัวหนา (Ctrl+B)",
  "Borders": "เส้นขอบ",
  "Bottom": "ชิดล่าง",
  "Bullets": "สัญลักษณ์หัวข้อ",
  "Center": "กึ่งกลาง",
  "Clear filters": "ล้างตัวกรอง",
  "Clear formatting": "ล้างการจัดรูปแบบ",
  "Clear": "ล้าง",
  "Clipboard": "คลิปบอร์ด",
  "Columns": "คอลัมน์",
  "Comma style": "ลักษณะจุลภาค",
  "Copy (Ctrl+C)": "คัดลอก (Ctrl+C)",
  "Cut (Ctrl+X)": "ตัด (Ctrl+X)",
  "Data": "ข้อมูล",
  "Decrease indent (Shift+Tab in a list)": "ลดการเยื้อง (Shift+Tab ในรายการ)",
  "Delete columns": "ลบคอลัมน์",
  "Delete rows": "ลบแถว",
  "Draw on the page": "วาดบนหน้า",
  "Draw": "วาด",
  "Edit the header and footer (or double-click one)": "แก้หัวกระดาษและท้ายกระดาษ (หรือดับเบิลคลิก)",
  "Export Markdown (.md)": "ส่งออก Markdown (.md)",
  "Export PDF (print to PDF)": "ส่งออก PDF (พิมพ์เป็น PDF)",
  "Export plain text (.txt)": "ส่งออกข้อความ (.txt)",
  "Export web page (.html)": "ส่งออกหน้าเว็บ (.html)",
  "Export": "ส่งออก",
  "Fewer decimals": "ลดทศนิยม",
  "File name": "ชื่อไฟล์",
  "File": "ไฟล์",
  "Fill colour": "สีพื้น",
  "Filter": "ตัวกรอง",
  "Find & export": "ค้นหาและส่งออก",
  "Find & replace": "ค้นหาและแทนที่",
  "Find and replace (Ctrl+F / Ctrl+H)": "ค้นหาและแทนที่ (Ctrl+F / Ctrl+H)",
  "Find": "ค้นหา",
  "Font colour": "สีตัวอักษร",
  "Font size": "ขนาดตัวอักษร",
  "Font": "แบบอักษร",
  "Format as table": "จัดรูปแบบเป็นตาราง",
  "Format as table: header row and banded rows in a colour theme": "จัดรูปแบบเป็นตาราง: แถวหัวและแถวสลับสีตามธีม",
  "Formulas": "สูตร",
  "Freeze panes": "ตรึงแนว",
  "Functions": "ฟังก์ชัน",
  "Header & footer": "หัวและท้ายกระดาษ",
  "Highlight colour": "สีไฮไลต์",
  "Home": "หน้าแรก",
  "Increase indent (Tab in a list)": "เพิ่มการเยื้อง (Tab ในรายการ)",
  "Insert columns left": "แทรกคอลัมน์ทางซ้าย",
  "Insert columns right": "แทรกคอลัมน์ทางขวา",
  "Insert function": "แทรกฟังก์ชัน",
  "Insert rows above": "แทรกแถวด้านบน",
  "Insert rows below": "แทรกแถวด้านล่าง",
  "Insert": "แทรก",
  "Italic (Ctrl+I)": "ตัวเอียง (Ctrl+I)",
  "Justify": "จัดเต็มบรรทัด",
  "Keep a draft on this device": "เก็บฉบับร่างไว้ในเครื่องนี้",
  "Line spacing": "ระยะห่างบรรทัด",
  "Line width": "ความหนาเส้น",
  "Line": "เส้น",
  "Merge & center": "ผสานและจัดกึ่งกลาง",
  "Middle": "กึ่งกลางแนวตั้ง",
  "More decimals": "เพิ่มทศนิยม",
  "New": "ใหม่",
  "Number format": "รูปแบบตัวเลข",
  "Number": "ตัวเลข",
  "Numbering": "ลำดับเลข",
  "Open…": "เปิด…",
  "Page numbers": "เลขหน้า",
  "Page numbers: format, first number, start and stop page": "เลขหน้า: รูปแบบ เลขแรก หน้าเริ่มและหน้าสุดท้าย",
  "Paragraph": "ย่อหน้า",
  "Paste (Ctrl+V)": "วาง (Ctrl+V)",
  "Percent style": "ลักษณะเปอร์เซ็นต์",
  "Rectangle": "สี่เหลี่ยม",
  "Rows": "แถว",
  "Save draft": "บันทึกฉบับร่าง",
  "Save": "บันทึก",
  "Shape and text box colour": "สีรูปร่างและกล่องข้อความ",
  "Shape style": "ลักษณะรูปร่าง",
  "Sort & filter": "เรียงและกรอง",
  "Sort A to Z (smallest first)": "เรียง ก ถึง ฮ (น้อยไปมาก)",
  "Sort Z to A (largest first)": "เรียง ฮ ถึง ก (มากไปน้อย)",
  "Start a new document": "เริ่มเอกสารใหม่",
  "Strikethrough": "ขีดฆ่า",
  "Styles": "ลักษณะ",
  "Text box": "กล่องข้อความ",
  "Top": "ชิดบน",
  "Type in the document": "พิมพ์ในเอกสาร",
  "Underline (Ctrl+U)": "ขีดเส้นใต้ (Ctrl+U)",
  "View": "มุมมอง",
  "Wrap text": "ตัดข้อความ",
  "Open a Word file (.docx)": "เปิดไฟล์ Word (.docx)",
  "Download .docx": "ดาวน์โหลด .docx",
  "Download": "ดาวน์โหลด",
  "Open a file": "เปิดไฟล์",
};
export function ribbon(spec, prefix, start = 0) {
  const el = document.createElement('div'); el.className = 'rb';
  const tabs = spec.map((t, i) => `<button type="button" role="tab" class="rb-tab${t.file ? ' rb-file' : ''}" id="${prefix}-tab-${i}" aria-selected="${i === start}" aria-controls="${prefix}-pan-${i}">${esc(say(t.tab))}</button>`).join('');
  const item = (it) => {
    if (it.btn) return `<button type="button" class="rb-btn" id="${it.btn}" title="${esc(say(it.title))}" aria-label="${esc(say(it.title))}"${it.toggle ? ' aria-pressed="false"' : ''}>${it.icon}${it.text ? `<span class="rb-txt">${esc(say(it.text))}</span>` : ''}</button>`;
    if (it.select) return `<select class="rb-sel" id="${it.select}" title="${esc(say(it.title))}" aria-label="${esc(say(it.title))}"${it.width ? ` style="width:${it.width}px"` : ''}>${it.options.map((o) => (Array.isArray(o) ? `<option value="${esc(o[0])}">${esc(o[1])}</option>` : `<option>${esc(o)}</option>`)).join('')}</select>`;
    if (it.color) return `<label class="rb-color" title="${esc(say(it.title))}">${it.icon}<input type="color" id="${it.color}" aria-label="${esc(say(it.title))}"></label>`;
    if (it.input) return `<input class="rb-in" id="${it.input}" title="${esc(it.title)}" aria-label="${esc(it.title)}" placeholder="${esc(it.placeholder || '')}" autocomplete="off" spellcheck="false"${it.width ? ` style="width:${it.width}px"` : ''}>`;
    return it.html || '';
  };
  const panels = spec.map((t, i) => `<div class="rb-pan" role="tabpanel" id="${prefix}-pan-${i}" aria-labelledby="${prefix}-tab-${i}"${i !== start ? ' hidden' : ''}>${t.groups.map((g) => `<div class="rb-grp" role="group" aria-label="${esc(say(g.label))}"><div class="rb-items">${g.items.map(item).join('')}</div><span class="rb-glabel">${esc(say(g.label))}</span></div>`).join('')}</div>`).join('');
  el.innerHTML = `<div class="rb-tabs" role="tablist">${tabs}</div><div class="rb-pans">${panels}</div>`;
  el.querySelector('.rb-tabs').addEventListener('click', (e) => {
    const b = e.target.closest('.rb-tab'); if (!b) return;
    // The open tab again: fold the groups away or back (Office's collapsed ribbon); any other tab unfolds them.
    const fold = b.getAttribute('aria-selected') === 'true' && !el.classList.contains('rb-min');
    el.classList.toggle('rb-min', fold);
    el.querySelectorAll('.rb-tab').forEach((x) => x.setAttribute('aria-expanded', String(!fold)));
    b.title = fold ? 'Show the ribbon' : 'Click again to hide the ribbon';
    el.querySelectorAll('.rb-tab').forEach((x) => x.setAttribute('aria-selected', String(x === b)));
    el.querySelectorAll('.rb-pan').forEach((p) => { p.hidden = p.id !== b.getAttribute('aria-controls'); });
  });
  // Arrow keys move between tabs.
  el.querySelector('.rb-tabs').addEventListener('keydown', (e) => {
    if (e.key !== 'ArrowRight' && e.key !== 'ArrowLeft') return;
    const all = [...el.querySelectorAll('.rb-tab')]; const i = all.indexOf(document.activeElement); if (i < 0) return;
    const n = all[(i + (e.key === 'ArrowRight' ? 1 : all.length - 1)) % all.length]; n.focus(); n.click();
  });
  // Buttons keep the document's selection and focus.
  el.addEventListener('pointerdown', (e) => { if (e.target.closest('.rb-btn')) e.preventDefault(); });
  el.addEventListener('mousedown', (e) => { if (e.target.closest('.rb-btn')) e.preventDefault(); });
  return el;
}
// The File tab (Excel, Word, Markdown): Open (and New), Save draft with the time it was saved, Download. Ids:
// `${p}-f-open`, `${p}-f-new`, `${p}-f-draft`, `${p}-f-saved`, `${p}-f-dl`; the page or the editor wires them
// (index.html barActions, js/md-edit.js). `name`: a file name field (its id) in the Save group.
export function fileTab(p, { open = 'Open a file', newDoc = false, name = null, download = 'Download' } = {}) {
  return { tab: 'File', file: true, groups: [
    { label: 'File', items: [{ btn: `${p}-f-open`, icon: RI.open, text: 'Open…', title: open }, ...(newDoc ? [{ btn: `${p}-f-new`, icon: RI.newDoc, text: 'New', title: 'Start a new document' }] : [])] },
    { label: 'Save', items: [...(name ? [{ input: name, title: 'File name', placeholder: 'File name', width: 180 }] : []),
      { btn: `${p}-f-draft`, icon: RI.save, text: 'Save draft', title: 'Keep a draft on this device' },
      { html: `<span class="rb-note" id="${p}-f-saved" aria-live="polite"></span>` },
      { btn: `${p}-f-dl`, icon: RI.dl, text: download, title: download }] },
  ] };
}
/** "Saved 14:02" beside Save draft in the File tab. */
export const markSaved = (root, p, at = Date.now()) => { const n = root.querySelector(`#${p}-f-saved`); if (n) n.textContent = `Saved ${new Date(at).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}`; };
export const setPressed = (root, id, on) => { const b = root.querySelector(`#${id}`); if (b) b.setAttribute('aria-pressed', String(!!on)); };
