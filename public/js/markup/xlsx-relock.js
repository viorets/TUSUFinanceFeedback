// Put an Excel file's locks back after it was edited in Google Sheets (v2.12 test run).
// Google Sheets drops sheet protection, the workbook structure lock and the locked / unlocked flag of every cell.
// relock(original, edited) copies all three from the organization's ORIGINAL file onto the edited one:
// - each cell's locked flag (by sheet name and address), with its own style copy (ExcelJS shares styles between cells);
// - each sheet's <sheetProtection> element and the <workbookProtection> element, copied as XML, so the password hash
//   stays the same and the site never needs the password.
// Both arguments are ArrayBuffers of .xlsx files; resolves the re-locked .xlsx as an ArrayBuffer.
const EXCELJS = 'https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js';
const JSZIP = 'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js';

const script = (src, global) => new Promise((res, rej) => {
  if (window[global]) { res(window[global]); return; }
  const s = document.createElement('script'); s.src = src; s.crossOrigin = 'anonymous';
  s.onload = () => res(window[global]); s.onerror = () => rej(new Error('Could not load the Excel tools. Try again.'));
  document.head.append(s);
});

// Sheet name -> path of its XML part inside the .xlsx (from workbook.xml and its relationships).
async function sheetParts(zip) {
  const wb = await zip.file('xl/workbook.xml').async('string');
  const rels = await zip.file('xl/_rels/workbook.xml.rels').async('string');
  const target = {};
  for (const m of rels.matchAll(/<Relationship\b[^>]*>/g)) {
    const id = /Id="([^"]+)"/.exec(m[0])?.[1]; const t = /Target="([^"]+)"/.exec(m[0])?.[1];
    if (id && t) target[id] = t.startsWith('/') ? t.slice(1) : `xl/${t.replace(/^\.\//, '')}`;
  }
  const out = {};
  for (const m of wb.matchAll(/<sheet\b[^>]*>/g)) {
    const name = /name="([^"]+)"/.exec(m[0])?.[1]; const rid = /r:id="([^"]+)"/.exec(m[0])?.[1];
    if (name && rid && target[rid]) out[unxml(name)] = target[rid];
  }
  return out;
}
const unxml = (s) => s.replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');

export async function relock(original, edited) {
  const [ExcelJS, JSZip] = await Promise.all([script(EXCELJS, 'ExcelJS'), script(JSZIP, 'JSZip')]);

  // 1. Cell flags: which cells the original leaves open, per sheet.
  const src = new ExcelJS.Workbook(); await src.xlsx.load(original.slice(0));
  const open = {};
  src.worksheets.forEach((ws) => {
    const set = new Set();
    ws.eachRow({ includeEmpty: true }, (row) => row.eachCell({ includeEmpty: true }, (cell) => { if (cell.protection && cell.protection.locked === false) set.add(cell.address); }));
    open[ws.name] = set;
  });
  const out = new ExcelJS.Workbook(); await out.xlsx.load(edited.slice(0));
  out.worksheets.forEach((ws) => {
    const set = open[ws.name] || new Set();
    ws.eachRow({ includeEmpty: true }, (row) => row.eachCell({ includeEmpty: true }, (cell) => {
      cell.style = { ...cell.style, protection: { locked: !set.has(cell.address) } };
    }));
    set.forEach((addr) => { const cell = ws.getCell(addr); cell.style = { ...cell.style, protection: { locked: false } }; }); // open cells that are empty
    // Layout Google changes: column widths and merged cells come back from the original (the form's own layout).
    const orig = src.getWorksheet(ws.name);
    if (!orig) return;
    for (let c = 1; c <= Math.max(orig.columnCount, 26); c++) { const w = orig.getColumn(c).width; if (w) ws.getColumn(c).width = w; }
    const want = orig.model.merges || []; const have = ws.model.merges || [];
    want.filter((m) => !have.includes(m)).forEach((m) => {
      const tl = m.split(':')[0];
      have.filter((h) => h.split(':')[0] === tl).forEach((h) => { try { ws.unMergeCells(h); } catch { /* already gone */ } });
      try { ws.mergeCells(m); } catch { /* overlaps an edit: leave Google's merge */ }
    });
  });
  const flagged = await out.xlsx.writeBuffer();

  // 2. Protection elements, copied as XML from the original.
  const a = await JSZip.loadAsync(original.slice(0));
  const b = await JSZip.loadAsync(flagged);
  const pa = await sheetParts(a); const pb = await sheetParts(b);
  for (const [name, partB] of Object.entries(pb)) {
    if (!pa[name]) continue;
    const xa = await a.file(pa[name]).async('string');
    const prot = /<sheetProtection\b[^>]*\/>/.exec(xa)?.[0];
    let xb = await b.file(partB).async('string');
    xb = xb.replace(/<sheetProtection\b[^>]*\/>/g, '');
    if (prot) xb = xb.replace(/(<\/sheetData>|<sheetData\s*\/>)/, `$1${prot}`); // OOXML order: right after sheetData
    b.file(partB, xb);
  }
  const wa = await a.file('xl/workbook.xml').async('string');
  const wprot = /<workbookProtection\b[^>]*\/>/.exec(wa)?.[0];
  if (wprot) {
    let wb = await b.file('xl/workbook.xml').async('string');
    wb = wb.replace(/<workbookProtection\b[^>]*\/>/g, '');
    wb = /<bookViews\b/.test(wb) ? wb.replace(/<bookViews\b/, `${wprot}<bookViews`) : wb.replace(/<sheets\b/, `${wprot}<sheets`);
    b.file('xl/workbook.xml', wb);
  }
  return b.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
}
