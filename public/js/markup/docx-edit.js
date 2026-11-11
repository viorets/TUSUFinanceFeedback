// Word editor (prototype): the document is shown as a page in its own fonts, sizes, colours, alignment and spacing,
// with its tables and images, and the text of any paragraph or table cell can be changed in place. The file is opened
// as a zip (JSZip) and only word/document.xml changes, so styles, headers, footers, images and page layout stay as
// they were. Only the changed characters are rewritten, so the rest of a paragraph keeps its bold, colours and fonts.
// New text takes the font and size of the text before it; typed into the document (not a text box) it is in the
// feedback colour (TYPED, v2.16.4), and it is written into the file as it is typed, so it shows in that colour at once.
// The toolbar sets the font, size and colour of selected text, and adds text boxes, lines and arrows (Word shapes);
// shapes and text boxes already in the file are shown, can be moved or deleted, and their text edited.
// The text is shown in pages with each page's header and footer (typing there keeps the file's format), and page
// numbers can be added (Word fields). New text: Markup (the feedback colour, TYPED) or Edit (the format of the text
// before it). Pages, headers and footers and page numbers come from TUSU Finance Team Tools' Word editor.
const JSZIP = 'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js';
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const WPS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const XML_NS = 'http://www.w3.org/XML/1998/namespace';
const SVG = 'http://www.w3.org/2000/svg';
const THAI = /[฀-๿]/;
const HIGHLIGHT = { darkYellow: '#808000', lightGray: '#d3d3d3', darkGray: '#a9a9a9' }; // the rest are CSS colour names
const EMU = 9525; // drawing units per screen pixel (96 dpi)
// Fonts offered (after the ones the file uses). The page and Word show a font only where it is installed; LINE Seed
// Sans TH, Sarabun and Noto Sans Thai are loaded by the page.
const FONTS = ['LINE Seed Sans TH', 'TH Sarabun New', 'TH SarabunPSK', 'Sarabun', 'Noto Sans Thai', 'Leelawadee UI', 'Tahoma', 'Angsana New', 'Cordia New', 'Calibri', 'Arial', 'Times New Roman'];
const SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 32, 36, 48, 72];
const WIDTHS = [0.75, 1, 1.5, 2.25, 3, 4.5, 6]; // line widths (pt), as Word offers them
// New text boxes and shapes take the PDF tab's defaults: LINE Seed Sans TH, 11 pt, #990000, clear, no border.
const BOX = { font: 'LINE Seed Sans TH', size: 11, color: '#990000' };
const SHAPE = { color: '#990000', width: 1.5 };
const TYPED = BOX.color; // text typed into the document (v2.16.4): the feedback colour, as in text boxes
const LINES = /^(line|straightConnector1|bentConnector\d|curvedConnector\d)$/;
// Run properties must be written in this order, or Word reports the file as damaged.
const RPR_ORDER = ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike', 'outline', 'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing', 'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'effect', 'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath', 'rPrChange'];

const loadJSZip = () => window.JSZip ? Promise.resolve(window.JSZip) : new Promise((res, rej) => {
  const s = document.createElement('script'); s.src = JSZIP; s.crossOrigin = 'anonymous';
  s.onload = () => res(window.JSZip); s.onerror = () => rej(new Error('Could not load the Word tools. Try again.'));
  document.head.append(s);
});

const parse = (s) => new DOMParser().parseFromString(s, 'application/xml');
const kids = (el, name) => el ? [...el.children].filter((c) => c.namespaceURI === W && c.localName === name) : [];
const kid = (el, name) => kids(el, name)[0] || null;
const attr = (el, name) => el?.getAttributeNS(W, name) || null;
const val = (el) => attr(el, 'val');
const on = (el) => !['0', 'false', 'off'].includes(val(el));
const nsKid = (el, ns, name) => el ? [...el.children].find((c) => c.namespaceURI === ns && c.localName === name) || null : null;
const ancestor = (el, name, ns = W) => { for (let n = el.parentNode; n && n.nodeType === 1; n = n.parentNode) if (n.namespaceURI === ns && n.localName === name) return n; return null; };
// A shape's text is stored twice when the file has a fallback copy for old Word versions; only the first is used.
const inFallback = (el) => !!ancestor(el, 'Fallback', MC);
// The drawing of an mc:AlternateContent run child: the one in mc:Choice.
const choiceDrawing = (ac) => nsKid(ac, MC, 'Choice')?.getElementsByTagNameNS(W, 'drawing')[0] || null;
// Text pieces of a run: w:t, and w:tab / w:br directly in a run (a w:tab under w:tabs is a tab stop; page breaks are kept).
const isPiece = (el) => el.parentNode?.localName === 'r' && (el.localName === 't' || el.localName === 'tab' || el.localName === 'cr' || (el.localName === 'br' && !el.getAttributeNS(W, 'type')));
const pieceText = (el) => el.localName === 't' ? el.textContent : el.localName === 'tab' ? '\t' : '\n';
const pt = (twips) => `${twips / 20}pt`;

// Formatting resolved the way Word does it: document defaults, then the paragraph style (and the styles it is based
// on), then the character style, then the run's own properties. Theme fonts come from the theme part.
function styleReader(stylesXml, themeXml, numberingXml) {
  const tf = {}; const scheme = {};
  if (themeXml) {
    const t = parse(themeXml);
    // Theme colours, for shapes coloured by the theme (bg1 = lt1, tx1 = dk1, ...).
    for (const c of t.getElementsByTagNameNS(A, 'clrScheme')[0]?.children || []) {
      const v = c.firstElementChild; const hex = v?.getAttribute('val') && v.localName === 'srgbClr' ? v.getAttribute('val') : v?.getAttribute('lastClr');
      if (hex) scheme[c.localName] = `#${hex}`;
    }
    Object.assign(scheme, { bg1: scheme.lt1, tx1: scheme.dk1, bg2: scheme.lt2, tx2: scheme.dk2 });
    for (const kind of ['major', 'minor']) {
      const f = t.getElementsByTagNameNS(A, `${kind}Font`)[0]; if (!f) continue;
      const latin = f.getElementsByTagNameNS(A, 'latin')[0]?.getAttribute('typeface');
      const cs = f.getElementsByTagNameNS(A, 'cs')[0]?.getAttribute('typeface');
      const thai = [...f.getElementsByTagNameNS(A, 'font')].find((x) => x.getAttribute('script') === 'Thai')?.getAttribute('typeface');
      tf[`${kind}HAnsi`] = tf[`${kind}Ascii`] = latin; tf[`${kind}Bidi`] = cs || thai;
    }
  }
  function rProps(rPr) {
    const o = {}; if (!rPr) return o;
    const f = kid(rPr, 'rFonts');
    if (f) { const a = attr(f, 'ascii') || tf[attr(f, 'asciiTheme')]; const c = attr(f, 'cs') || tf[attr(f, 'cstheme')]; if (a) o.ascii = a; if (c) o.cs = c; }
    for (const k of ['b', 'bCs', 'i', 'iCs', 'strike', 'caps']) { const e = kid(rPr, k); if (e) o[k] = on(e); }
    const u = kid(rPr, 'u'); if (u) o.u = val(u) !== 'none';
    const sz = kid(rPr, 'sz'); if (sz) o.sz = +val(sz) / 2;
    const szCs = kid(rPr, 'szCs'); if (szCs) o.szCs = +val(szCs) / 2;
    const c = kid(rPr, 'color'); if (val(c) && val(c) !== 'auto') o.color = `#${val(c)}`;
    const h = kid(rPr, 'highlight'); if (val(h) && val(h) !== 'none') o.highlight = HIGHLIGHT[val(h)] || val(h);
    const fill = attr(kid(rPr, 'shd'), 'fill'); if (fill && fill !== 'auto') o.shade = `#${fill}`;
    const va = kid(rPr, 'vertAlign'); if (va) o.vertAlign = val(va);
    return o;
  }
  function pProps(pPr) {
    const o = {}; if (!pPr) return o;
    const jc = kid(pPr, 'jc'); if (jc) o.jc = val(jc);
    const sp = kid(pPr, 'spacing');
    if (sp) { for (const k of ['before', 'after', 'line']) { const v = attr(sp, k); if (v) o[k] = +v; } if (attr(sp, 'lineRule')) o.lineRule = attr(sp, 'lineRule'); }
    const ind = kid(pPr, 'ind');
    if (ind) {
      for (const [k, alt] of [['left', 'start'], ['right', 'end']]) { const v = attr(ind, k) || attr(ind, alt); if (v) o[k] = +v; }
      if (attr(ind, 'hanging')) { o.hanging = +attr(ind, 'hanging'); o.firstLine = 0; } else if (attr(ind, 'firstLine')) { o.firstLine = +attr(ind, 'firstLine'); o.hanging = 0; }
    }
    const num = kid(pPr, 'numPr');
    if (num) { const id = val(kid(num, 'numId')); if (id) o.numId = id; const lv = val(kid(num, 'ilvl')); if (lv) o.ilvl = +lv; }
    const fill = attr(kid(pPr, 'shd'), 'fill'); if (fill && fill !== 'auto') o.shade = `#${fill}`;
    const bdr = kid(pPr, 'pBdr');
    if (bdr) {
      o.border = {};
      for (const side of ['top', 'bottom', 'left', 'right']) {
        const e = kid(bdr, side); if (!e || ['nil', 'none'].includes(val(e))) continue;
        o.border[side] = { width: Math.max(0.5, (+attr(e, 'sz') || 4) / 8), color: attr(e, 'color') && attr(e, 'color') !== 'auto' ? `#${attr(e, 'color')}` : '#000', space: +attr(e, 'space') || 0 };
      }
    }
    return o;
  }
  const doc = stylesXml ? parse(stylesXml) : null;
  const styles = new Map(); let defaultPara = null;
  for (const s of doc?.getElementsByTagNameNS(W, 'style') || []) {
    styles.set(attr(s, 'styleId'), s);
    if (attr(s, 'type') === 'paragraph' && ['1', 'true', 'on'].includes(attr(s, 'default'))) defaultPara = attr(s, 'styleId');
  }
  const dd = doc?.getElementsByTagNameNS(W, 'docDefaults')[0];
  const base = { r: rProps(kid(kid(dd, 'rPrDefault'), 'rPr')), p: pProps(kid(kid(dd, 'pPrDefault'), 'pPr')) };
  const memo = new Map();
  function style(id) {
    if (!id || !styles.has(id)) return { r: {}, p: {}, borders: false };
    if (memo.has(id)) return memo.get(id);
    memo.set(id, { r: {}, p: {}, borders: false }); // a style based on itself stops here
    const s = styles.get(id); const parent = style(val(kid(s, 'basedOn')));
    const tb = kid(kid(s, 'tblPr'), 'tblBorders');
    const out = { r: { ...parent.r, ...rProps(kid(s, 'rPr')) }, p: { ...parent.p, ...pProps(kid(s, 'pPr')) }, borders: tb ? hasBorder(tb) : parent.borders };
    memo.set(id, out); return out;
  }
  const hasBorder = (tb) => [...tb.children].some((b) => !['nil', 'none'].includes(val(b)));

  // Lists (numbering.xml): a paragraph's numId points to a list definition; each level has a number format, the
  // marker text ("%1.") and its own indent. Each numId counts on its own (checked against Word).
  const nums = new Map(); const abstracts = new Map();
  if (numberingXml) {
    const nd = parse(numberingXml);
    for (const a of nd.getElementsByTagNameNS(W, 'abstractNum')) abstracts.set(attr(a, 'abstractNumId'), a);
    for (const n of nd.getElementsByTagNameNS(W, 'num')) nums.set(attr(n, 'numId'), n);
  }
  function level(numId, ilvl) {
    const n = nums.get(numId); if (!n) return null;
    const absId = val(kid(n, 'abstractNumId')); const a = abstracts.get(absId); if (!a) return null;
    const lvl = (x) => kids(x, 'lvl').find((l) => +attr(l, 'ilvl') === ilvl);
    const override = kids(n, 'lvlOverride').find((o) => +attr(o, 'ilvl') === ilvl);
    const l = lvl(override) || lvl(a); if (!l) return null;
    const startOverride = val(kid(override, 'startOverride'));
    return { list: numId, fmt: val(kid(l, 'numFmt')) || 'decimal', text: val(kid(l, 'lvlText')) ?? '', start: +(startOverride ?? val(kid(l, 'start')) ?? 1), p: pProps(kid(l, 'pPr')), r: rProps(kid(l, 'rPr')) };
  }
  return {
    scheme,
    para(p) {
      const pPr = kid(p, 'pPr'); const st = style(val(kid(pPr, 'pStyle')) || defaultPara);
      const own = pProps(pPr); const merged = { ...base.p, ...st.p, ...own };
      const lv = merged.numId && merged.numId !== '0' ? level(merged.numId, merged.ilvl || 0) : null;
      // The list level's indent applies unless the paragraph or its style sets one.
      return { p: lv ? { ...base.p, ...st.p, ...lv.p, ...own } : merged, r: { ...base.r, ...st.r }, level: lv, ilvl: merged.ilvl || 0 };
    },
    run(r, paraR) { const rPr = kid(r, 'rPr'); return { ...paraR, ...style(val(kid(rPr, 'rStyle'))).r, ...rProps(rPr) }; },
    tableBorders(tbl) { const tblPr = kid(tbl, 'tblPr'); const tb = kid(tblPr, 'tblBorders'); return tb ? hasBorder(tb) : style(val(kid(tblPr, 'tblStyle'))).borders; },
  };
}

// List markers: numbers in the level's format; bullets from symbol fonts shown as the matching Unicode character.
const ROMAN = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']];
const roman = (n) => ROMAN.reduce((s, [v, r]) => { while (n >= v) { s += r; n -= v; } return s; }, '');
const letters = (n, abc) => abc[(n - 1) % abc.length].repeat(Math.floor((n - 1) / abc.length) + 1);
const THAI_LETTERS = 'กขคงจฉชซฌญฎฏฐฑฒณดตถทธนบปผฝพฟภมยรลวศษสหฬอฮ';
function formatNumber(n, fmt) {
  switch (fmt) {
    case 'lowerLetter': return letters(n, 'abcdefghijklmnopqrstuvwxyz');
    case 'upperLetter': return letters(n, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ');
    case 'lowerRoman': return roman(n);
    case 'upperRoman': return roman(n).toUpperCase();
    case 'decimalZero': return String(n).padStart(2, '0');
    case 'thaiNumbers': case 'thaiCounting': return String(n).replace(/\d/g, (d) => String.fromCharCode(0x0E50 + +d));
    case 'thaiLetters': return letters(n, THAI_LETTERS);
    case 'none': return '';
    default: return String(n);
  }
}
const BULLETS = { '': '•', '': '▪', '': '➢', '': '✓', '': '❖', '': '◦', o: '◦' };

function paraCss(p) {
  const list = p.numId && p.numId !== '0';
  const s = [`margin:${pt(p.before || 0)} ${pt(p.right || 0)} ${pt(p.after || 0)} ${pt(p.left || (list ? 720 : 0))}`];
  if (p.jc) s.push(`text-align:${{ both: 'justify', distribute: 'justify', start: 'left', end: 'right' }[p.jc] || p.jc}`);
  // A list paragraph's marker hangs in the indent (::before); its text starts at the left indent.
  if (list) s.push(`--hang:${pt(p.hanging || 360)}`);
  else if (p.firstLine) s.push(`text-indent:${pt(p.firstLine)}`); else if (p.hanging) s.push(`text-indent:-${pt(p.hanging)}`);
  for (const [side, b] of Object.entries(p.border || {})) s.push(`border-${side}:${b.width}pt solid ${b.color}`, `padding-${side}:${b.space}pt`);
  if (p.line) s.push(p.lineRule === 'exact' || p.lineRule === 'atLeast' ? `line-height:${pt(p.line)}` : `line-height:${(p.line / 240 * 1.2).toFixed(2)}`);
  if (p.shade) s.push(`background:${p.shade}`);
  return s.join(';');
}
// Thai text uses the complex-script font, size, bold and italic, as Word does.
function runCss(r, thai) {
  const fonts = (thai ? [r.cs, r.ascii] : [r.ascii, r.cs]).filter(Boolean).map((f) => `"${f}"`);
  const size = (thai ? r.szCs ?? r.sz : r.sz) || 11;
  const s = [`font-family:${[...fonts, 'sans-serif'].join(',')}`, `font-size:${size}pt`];
  if (thai ? r.bCs ?? r.b : r.b) s.push('font-weight:bold');
  if (thai ? r.iCs ?? r.i : r.i) s.push('font-style:italic');
  const deco = [r.u && 'underline', r.strike && 'line-through'].filter(Boolean); if (deco.length) s.push(`text-decoration:${deco.join(' ')}`);
  if (r.color) s.push(`color:${r.color}`);
  if (r.highlight || r.shade) s.push(`background:${r.highlight || r.shade}`);
  if (r.caps) s.push('text-transform:uppercase');
  if (r.vertAlign === 'superscript' || r.vertAlign === 'subscript') s.push(`vertical-align:${r.vertAlign === 'superscript' ? 'super' : 'sub'};font-size:${size * 0.65}pt`);
  return s.join(';');
}

// Opens the editor in `root`. bytes: the .docx (ArrayBuffer); t(en, th): translates labels. Resolves
// { isDirty(), save() }; save() resolves the edited .docx as an ArrayBuffer.
export async function openDocxEditor(root, bytes, { t = (en) => en } = {}) {
  const JSZip = await loadJSZip();
  let zip;
  try { zip = await JSZip.loadAsync(bytes.slice(0)); } catch { zip = null; }
  const part = zip?.file('word/document.xml');
  if (!part) throw new Error(t('Could not read this file. Only .docx files can be edited (not old .doc files).', 'อ่านไฟล์นี้ไม่ได้ แก้ไขได้เฉพาะไฟล์ .docx (ไม่ใช่ไฟล์ .doc แบบเก่า)'));
  const xml = parse(await part.async('string'));
  let body = xml.getElementsByTagNameNS(W, 'body')[0]; // replaced by Undo
  const text = (path) => zip.file(path)?.async('string');
  const st = styleReader(await text('word/styles.xml'), await text('word/theme/theme1.xml'), await text('word/numbering.xml'));
  const rels = {};
  for (const m of (await zip.file('word/_rels/document.xml.rels')?.async('string') || '').matchAll(/<Relationship\b[^>]*>/g)) {
    const id = /Id="([^"]+)"/.exec(m[0])?.[1]; const t = /Target="([^"]+)"/.exec(m[0])?.[1];
    if (id && t && !/TargetMode="External"/.test(m[0])) rels[id] = t.startsWith('/') ? t.slice(1) : `word/${t}`.replace(/[^/]+\/\.\.\//g, '');
  }
  // New shapes need a drawing id no other drawing in the file has (headers and footers included).
  let nextId = 1;
  for (const f of [part, ...zip.file(/^word\/(header|footer)\d*\.xml$/)]) {
    for (const m of (await f.async('string')).matchAll(/<wp:docPr\b[^>]*?\bid="(\d+)"/g)) nextId = Math.max(nextId, +m[1] + 1);
  }

  // Each paragraph's own text pieces (pieces of a text box inside it belong to the text box's paragraphs).
  const collect = (p) => [...p.getElementsByTagNameNS(W, '*')].filter((el) => isPiece(el) && ancestor(el, 'p') === p && !inFallback(el));
  const pieces = new Map();
  function scanPieces() {
    pieces.clear();
    for (const el of xml.getElementsByTagNameNS(W, '*')) {
      if (!isPiece(el)) continue;
      const p = ancestor(el, 'p'); if (!p || inFallback(el)) continue;
      if (!pieces.has(p)) pieces.set(p, []);
      pieces.get(p).push(el);
    }
  }
  scanPieces();

  const items = []; // { p, div, text }: one per paragraph shown
  const divOf = new Map(); // w:p -> its div
  function image(drawing, relMap = rels) {
    const span = document.createElement('span'); span.className = 'de-img'; span.contentEditable = 'false';
    const ext = drawing.getElementsByTagNameNS(WP, 'extent')[0];
    const w = +ext?.getAttribute('cx') / EMU || 0; const h = +ext?.getAttribute('cy') / EMU || 0;
    const rid = drawing.getElementsByTagNameNS(A, 'blip')[0]?.getAttributeNS(R, 'embed');
    const file = rid && relMap[rid] && zip.file(relMap[rid]);
    if (!file) return span;
    const img = document.createElement('img'); img.alt = ''; if (w) { img.width = w; img.height = h; }
    file.async('blob').then((b) => { img.src = URL.createObjectURL(b); });
    span.append(img);
    return span;
  }
  // List numbers count on in document order, per list; a level resets the levels below it.
  const counters = new Map();
  function marker(lv, ilvl) {
    if (lv.fmt === 'bullet') return BULLETS[lv.text] || lv.text || '•';
    const c = counters.get(lv.list) || []; counters.set(lv.list, c);
    c[ilvl] = c[ilvl] == null ? lv.start : c[ilvl] + 1; c.length = ilvl + 1;
    return lv.text.replace(/%(\d)/g, (_, k) => formatNumber(c[k - 1] ?? 1, k - 1 === ilvl ? lv.fmt : 'decimal'));
  }
  function renderPara(p, i = items.length) {
    const { p: pp, r: pr, level: lv, ilvl } = st.para(p);
    const div = document.createElement('div');
    div.className = 'de-p'; div.dataset.i = i;
    // The paragraph's own font: used for empty lines and new text (an empty paragraph: its paragraph mark's font).
    const own = pieces.get(p)?.length ? pr : st.run(kid(p, 'pPr'), pr);
    div.style.cssText = `${runCss(own, false)};${paraCss(pp)}`;
    const mark = items[i]?.marker ?? (lv ? marker(lv, ilvl) : null); // a redrawn paragraph keeps its number
    if (mark) { div.classList.add('de-li'); div.dataset.marker = mark; if (lv?.r.b) div.style.setProperty('--mark-weight', 'bold'); }
    items[i] = { p, div, text: null, marker: mark }; // taken before the runs: a text box inside gets the next number
    for (const r of p.getElementsByTagNameNS(W, 'r')) {
      if (ancestor(r, 'p') !== p || ancestor(r, 'del') || inFallback(r)) continue;
      const rp = st.run(r, pr);
      const text = [...r.children].filter(isPiece).map(pieceText).join('');
      let span = null;
      for (const c of r.children) {
        if (isPiece(c)) {
          if (!span) { span = document.createElement('span'); span.style.cssText = runCss(rp, THAI.test(text)); div.append(span); }
          span.textContent += pieceText(c);
        } else if (c.localName === 'drawing' || (c.namespaceURI === MC && c.localName === 'AlternateContent')) {
          span = null;
          const d = c.localName === 'drawing' ? c : choiceDrawing(c); if (!d) continue;
          if (isShape(d)) { if (!shapeOf.has(d)) addShape(d, p); } // floating shapes are drawn on the page layer
          else div.append(image(d));
        }
      }
    }
    div.contentEditable = 'plaintext-only'; if (div.contentEditable !== 'plaintext-only') div.contentEditable = 'true';
    divOf.set(p, div);
    return div;
  }
  function renderTable(tbl) {
    const table = document.createElement('table');
    table.className = `de-tbl${st.tableBorders(tbl) ? ' de-borders' : ''}`;
    if (val(kid(kid(tbl, 'tblPr'), 'jc')) === 'center') table.style.margin = '0 auto';
    for (const tr of kids(tbl, 'tr')) {
      const row = table.insertRow();
      for (const tc of kids(tr, 'tc')) {
        const tcPr = kid(tc, 'tcPr'); const td = row.insertCell();
        const span = +val(kid(tcPr, 'gridSpan')) || 1; if (span > 1) td.colSpan = span;
        const vm = kid(tcPr, 'vMerge'); if (vm && val(vm) !== 'restart') td.classList.add('de-vcont'); // rest of a vertically merged cell
        const w = kid(tcPr, 'tcW'); if (attr(w, 'type') === 'dxa') td.style.width = `${+attr(w, 'w') / 15}px`;
        const fill = attr(kid(tcPr, 'shd'), 'fill'); if (fill && fill !== 'auto') td.style.background = `#${fill}`;
        const va = val(kid(tcPr, 'vAlign')); if (va) td.style.verticalAlign = va === 'center' ? 'middle' : va;
        renderBlock(tc, td);
      }
    }
    return table;
  }
  function renderBlock(el, into) {
    for (const c of el.children) {
      if (c.namespaceURI !== W) continue;
      if (c.localName === 'p') into.append(renderPara(c));
      else if (c.localName === 'tbl') into.append(renderTable(c));
      else if (c.localName === 'sdt' || c.localName === 'customXml') renderBlock(kid(c, 'sdtContent') || c, into);
    }
  }

  // ---------- shapes: text boxes, lines, arrows ----------
  // A shape is a floating drawing (wp:anchor holding a wps:wsp) in a run of its anchor paragraph. This page has no
  // page breaks, so shapes are placed from the column's left edge and the anchor paragraph's top, the one position
  // that does not depend on where Word breaks pages. A shape that is moved is anchored to the body paragraph it now
  // sits by. Shapes placed from the page or margin are shown at their paragraph until moved.
  const layer = document.createElement('div'); layer.className = 'de-layer';
  const shapes = []; const shapeOf = new Map(); // w:drawing -> shape
  let sel = null; // the picked shape
  let dirty = false; // formatting or shapes changed
  const sect = kid(body, 'sectPr'); const pgW = +attr(kid(sect, 'pgSz'), 'w') || 11906; const mar = kid(sect, 'pgMar');
  const margin = (side) => (+attr(mar, side) || 1440) / 15;
  const colLeft = margin('left'); const colW = pgW / 15 - colLeft - margin('right');
  const page = document.createElement('div'); page.className = 'de-page';
  page.style.cssText = `width:${pgW / 15}px;padding:${margin('top')}px ${margin('right')}px ${margin('bottom')}px ${colLeft}px`;
  const svgEl = (tag, attrs) => { const e = document.createElementNS(SVG, tag); for (const [k, v] of Object.entries(attrs)) e.setAttribute(k, v); return e; };
  const E = (px) => String(Math.round(px * EMU));
  const anchorOf = (d) => nsKid(d, WP, 'anchor');
  const wspOf = (anchor) => nsKid(anchor?.getElementsByTagNameNS(A, 'graphicData')[0], WPS, 'wsp');
  const isShape = (d) => !!wspOf(anchorOf(d));
  function dmlColor(el) { // a:solidFill, a:lnRef, ... -> CSS colour
    const c = el?.firstElementChild; if (!c) return null;
    if (c.localName === 'srgbClr') return `#${c.getAttribute('val')}`;
    if (c.localName === 'schemeClr') return st.scheme[c.getAttribute('val')] || null;
    if (c.localName === 'sysClr') return `#${c.getAttribute('lastClr') || '000000'}`;
    if (c.localName === 'prstClr') return c.getAttribute('val');
    return null;
  }
  const styleRef = (s, name) => nsKid(nsKid(s.wsp, WPS, 'style'), A, name);
  function lineOf(s) {
    const ln = nsKid(s.spPr, A, 'ln');
    const end = (n) => { const t = nsKid(ln, A, n)?.getAttribute('type'); return !!t && t !== 'none'; };
    const color = dmlColor(nsKid(ln, A, 'solidFill')) || dmlColor(styleRef(s, 'lnRef')) || '#000000';
    const none = !!nsKid(ln, A, 'noFill') || (!ln?.firstElementChild && !styleRef(s, 'lnRef') && s.kind !== 'line');
    return { color, none, width: ln?.getAttribute('w') ? +ln.getAttribute('w') / EMU : 1, head: end('headEnd'), tail: end('tailEnd') };
  }
  function fillOf(s) {
    if (nsKid(s.spPr, A, 'noFill')) return 'none';
    const f = nsKid(s.spPr, A, 'solidFill'); if (f) return dmlColor(f) || 'none';
    const ref = styleRef(s, 'fillRef'); return ref && ref.getAttribute('idx') !== '0' ? dmlColor(ref) || 'none' : 'none';
  }
  function addShape(d, host) {
    const anchor = anchorOf(d); const wsp = wspOf(anchor); const spPr = nsKid(wsp, WPS, 'spPr');
    const prst = nsKid(spPr, A, 'prstGeom')?.getAttribute('prst') || 'rect';
    const txbx = nsKid(nsKid(wsp, WPS, 'txbx'), W, 'txbxContent');
    const s = { d, anchor, wsp, spPr, host, txbx, kind: txbx ? 'box' : LINES.test(prst) ? 'line' : prst === 'ellipse' ? 'ellipse' : 'shape' };
    const el = document.createElement('div'); el.className = `de-shape de-k-${s.kind}`; s.el = el;
    el.append(svgEl('svg', { class: 'de-svg' }));
    if (txbx) {
      const bp = nsKid(wsp, WPS, 'bodyPr'); const inset = (k, d0) => +(bp?.getAttribute(k) ?? d0) / EMU;
      s.auto = !!nsKid(bp, A, 'spAutoFit');
      const content = document.createElement('div'); content.className = 'de-txbx';
      content.style.padding = `${inset('tIns', 45720)}px ${inset('rIns', 91440)}px ${inset('bIns', 45720)}px ${inset('lIns', 91440)}px`;
      renderBlock(txbx, content);
      const grip = document.createElement('span'); grip.className = 'de-grip'; grip.title = t('Drag to move', 'ลากเพื่อย้าย');
      const size = document.createElement('span'); size.className = 'de-size'; size.title = t('Drag to resize', 'ลากเพื่อปรับขนาด');
      el.append(content, grip, size);
    }
    shapes.push(s); shapeOf.set(d, s); layer.append(el);
    return s;
  }
  const off = (pos) => +(nsKid(pos, WP, 'posOffset')?.textContent || 0) / EMU;
  let zoom = 1; // the page is shrunk to fit a narrow screen (an iPad held upright); page measures are unzoomed
  const hostTop = (p) => { const div = divOf.get(p); return div?.isConnected ? (div.getBoundingClientRect().top - page.getBoundingClientRect().top) / zoom : 0; };
  function geom(s) { // the shape's box on the page (px)
    const ext = nsKid(s.anchor, WP, 'extent'); const w = +ext?.getAttribute('cx') / EMU || 0; const h = +ext?.getAttribute('cy') / EMU || 0;
    const H = nsKid(s.anchor, WP, 'positionH'); const V = nsKid(s.anchor, WP, 'positionV');
    const align = nsKid(H, WP, 'align')?.textContent;
    const x = align ? colLeft + (align === 'center' ? (colW - w) / 2 : align === 'right' ? colW - w : 0)
      : (['page', 'leftMargin'].includes(H?.getAttribute('relativeFrom')) ? 0 : colLeft) + off(H);
    const y = hostTop(s.host) + (['paragraph', 'line'].includes(V?.getAttribute('relativeFrom')) ? off(V) : 0);
    const xfrm = nsKid(s.spPr, A, 'xfrm');
    return { x, y, w, h, flipH: xfrm?.getAttribute('flipH') === '1', flipV: xfrm?.getAttribute('flipV') === '1' };
  }
  const ends = (g) => [[g.x + (g.flipH ? g.w : 0), g.y + (g.flipV ? g.h : 0)], [g.x + (g.flipH ? 0 : g.w), g.y + (g.flipV ? 0 : g.h)]];
  const boxOf = (a, b) => ({ x: Math.min(a[0], b[0]), y: Math.min(a[1], b[1]), w: Math.abs(b[0] - a[0]), h: Math.abs(b[1] - a[1]), flipH: b[0] < a[0], flipV: b[1] < a[1] });
  const arrowHead = (a, b, w) => {
    const ang = Math.atan2(b[1] - a[1], b[0] - a[0]); const L = Math.max(7, w * 3.5); const sp = Math.PI / 7;
    return [b, [b[0] - L * Math.cos(ang - sp), b[1] - L * Math.sin(ang - sp)], [b[0] - L * Math.cos(ang + sp), b[1] - L * Math.sin(ang + sp)]];
  };
  function draw(s) {
    const g = s.g = geom(s); const ln = lineOf(s); const el = s.el; const svg = el.firstChild;
    el.style.cssText = `left:${g.x}px;top:${g.y}px;width:${g.w}px;${s.auto ? 'min-height' : 'height'}:${g.h}px`;
    el.classList.toggle('de-sel', s === sel);
    svg.replaceChildren(); el.querySelectorAll('.de-h').forEach((h) => h.remove());
    if (s.kind === 'line') {
      const [a, b] = ends(g).map(([x, y]) => [x - g.x, y - g.y]);
      const line = (attrs) => svgEl('line', { x1: a[0], y1: a[1], x2: b[0], y2: b[1], ...attrs });
      svg.append(line({ stroke: ln.color, 'stroke-width': ln.width, 'stroke-linecap': ln.head || ln.tail ? 'butt' : 'round' }));
      if (ln.tail) svg.append(svgEl('polygon', { fill: ln.color, points: arrowHead(a, b, ln.width).join(' ') }));
      if (ln.head) svg.append(svgEl('polygon', { fill: ln.color, points: arrowHead(b, a, ln.width).join(' ') }));
      svg.append(line({ class: 'de-hit', stroke: 'transparent', 'stroke-width': Math.max(12, ln.width + 8) }));
      if (s === sel) [a, b].forEach(([x, y], k) => { const h = document.createElement('span'); h.className = 'de-h'; h.dataset.h = k; h.style.cssText = `left:${x}px;top:${y}px`; el.append(h); });
    } else {
      const outline = { stroke: ln.none ? 'none' : ln.color, 'stroke-width': ln.width, fill: fillOf(s) };
      svg.append(s.kind === 'ellipse' ? svgEl('ellipse', { cx: '50%', cy: '50%', rx: '50%', ry: '50%', ...outline }) : svgEl('rect', { width: '100%', height: '100%', ...outline }));
    }
  }
  let queued = false;
  const layout = () => { if (queued) return; queued = true; requestAnimationFrame(() => { queued = false; shapes.forEach(draw); }); };

  // Writing a shape's position, size and direction into its drawing. A run holding a shape and text is split first,
  // so the shape can move on its own.
  function ownRun(s) {
    const run = ancestor(s.d, 'r');
    if (![...run.children].some(isPiece)) return run;
    let top = s.d; while (top.parentNode !== run) top = top.parentNode;
    const nr = xml.createElementNS(W, 'w:r'); nr.append(top); run.after(nr);
    return nr;
  }
  const bodyParas = (el = body) => [...el.children].flatMap((c) => c.namespaceURI !== W ? [] : c.localName === 'p' ? [c]
    : c.localName === 'sdt' || c.localName === 'customXml' ? bodyParas(kid(c, 'sdtContent') || c) : []);
  function hostAt(y) { // the last body paragraph that starts above y (inside a table: the paragraph before it)
    let best = null;
    for (const p of bodyParas()) { if (!divOf.get(p)?.isConnected) continue; if (!best || hostTop(p) <= y + 1) best = p; else break; }
    return best;
  }
  function wpEl(anchor, name, before) { // wp:positionH / wp:positionV / wp:extent, added in schema order if missing
    let el = nsKid(anchor, WP, name);
    if (!el) { el = xml.createElementNS(WP, `wp:${name}`); anchor.insertBefore(el, before.map((n) => nsKid(anchor, WP, n)).find(Boolean) || null); }
    return el;
  }
  function place(s, g) {
    const host = hostAt(g.y);
    if (host && host !== s.host) {
      const run = ownRun(s); const pPr = kid(host, 'pPr');
      host.insertBefore(run, pPr ? pPr.nextSibling : host.firstChild);
      s.host = host;
    }
    const a = s.anchor; a.setAttribute('simplePos', '0');
    const pos = (name, from, px) => {
      const el = wpEl(a, name, name === 'positionH' ? ['positionV', 'extent'] : ['extent']);
      el.setAttribute('relativeFrom', from);
      const o = xml.createElementNS(WP, 'wp:posOffset'); o.textContent = E(px); el.replaceChildren(o);
    };
    pos('positionH', 'column', g.x - colLeft); pos('positionV', 'paragraph', g.y - hostTop(s.host));
    const ext = wpEl(a, 'extent', ['effectExtent', 'wrapNone', 'wrapSquare', 'wrapTight', 'wrapThrough', 'wrapTopAndBottom', 'docPr']);
    ext.setAttribute('cx', E(g.w)); ext.setAttribute('cy', E(g.h));
    let xfrm = nsKid(s.spPr, A, 'xfrm');
    if (!xfrm) { xfrm = xml.createElementNS(A, 'a:xfrm'); s.spPr.prepend(xfrm); }
    if (!nsKid(xfrm, A, 'off')) { const o = xml.createElementNS(A, 'a:off'); o.setAttribute('x', '0'); o.setAttribute('y', '0'); xfrm.prepend(o); }
    let xe = nsKid(xfrm, A, 'ext'); if (!xe) { xe = xml.createElementNS(A, 'a:ext'); xfrm.append(xe); }
    xe.setAttribute('cx', E(g.w)); xe.setAttribute('cy', E(g.h));
    for (const f of ['flipH', 'flipV']) if (g[f]) xfrm.setAttribute(f, '1'); else xfrm.removeAttribute(f);
    dirty = true; draw(s);
  }
  // Line colour and width (0: no line, for text boxes and shapes).
  function setLine(s, color, width) {
    checkpoint();
    let ln = nsKid(s.spPr, A, 'ln');
    if (!ln) { ln = xml.createElementNS(A, 'a:ln'); s.spPr.insertBefore(ln, ['effectLst', 'effectDag', 'scene3d', 'sp3d', 'extLst'].map((n) => nsKid(s.spPr, A, n)).find(Boolean) || null); }
    for (const c of [...ln.children]) if (['noFill', 'solidFill', 'gradFill', 'pattFill'].includes(c.localName)) c.remove();
    if (!width && s.kind !== 'line') { ln.prepend(xml.createElementNS(A, 'a:noFill')); }
    else {
      const fill = xml.createElementNS(A, 'a:solidFill'); const c = xml.createElementNS(A, 'a:srgbClr');
      c.setAttribute('val', color.slice(1).toUpperCase()); fill.append(c); ln.prepend(fill);
      ln.setAttribute('w', String(Math.round((width || SHAPE.width) * 12700)));
    }
    dirty = true; draw(s);
  }
  function newShape(kind, g) {
    checkpoint();
    const id = nextId++; const name = { box: 'Text Box', line: 'Straight Connector', arrow: 'Straight Arrow Connector' }[kind];
    const f = BOX.font; const sz = BOX.size * 2; const color = BOX.color.slice(1).toUpperCase();
    const rPr = `<w:rPr><w:rFonts w:ascii="${f}" w:hAnsi="${f}" w:eastAsia="${f}" w:cs="${f}"/><w:color w:val="${color}"/><w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/></w:rPr>`;
    const ln = kind === 'box' ? '<a:noFill/><a:ln><a:noFill/></a:ln>'
      : `<a:ln w="${Math.round(lineWidth() * 12700)}"><a:solidFill><a:srgbClr val="${lineColor().slice(1).toUpperCase()}"/></a:solidFill>${kind === 'arrow' ? '<a:tailEnd type="triangle"/>' : ''}</a:ln>`;
    const wsp = kind === 'box'
      ? `<wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${ln}</wps:spPr>`
        + `<wps:txbx><w:txbxContent><w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>${rPr}</w:pPr></w:p></w:txbxContent></wps:txbx>`
        + '<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t" anchorCtr="0"><a:spAutoFit/></wps:bodyPr>'
      : `<wps:cNvCnPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm><a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>${ln}</wps:spPr><wps:bodyPr/>`;
    const run = xml.importNode(parse(`<w:r xmlns:w="${W}" xmlns:wp="${WP}" xmlns:a="${A}" xmlns:wps="${WPS}" xmlns:mc="${MC}"><w:rPr><w:noProof/></w:rPr>`
      + `<mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="${251659264 + id}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">`
      + '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>'
      + `<wp:extent cx="0" cy="0"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="${id}" name="${name} ${id}"/><wp:cNvGraphicFramePr/>`
      + `<a:graphic><a:graphicData uri="${WPS}"><wps:wsp>${wsp}</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice></mc:AlternateContent></w:r>`).documentElement, true);
    const host = hostAt(g.y); const pPr = kid(host, 'pPr');
    host.insertBefore(run, pPr ? pPr.nextSibling : host.firstChild);
    const first = items.length;
    const s = addShape(run.getElementsByTagNameNS(W, 'drawing')[0], host);
    for (let i = first; i < items.length; i++) items[i].text = textOf(items[i].div);
    place(s, g);
    return s;
  }
  function removeShape(s) {
    checkpoint();
    ownRun(s).remove(); s.el.remove(); shapes.splice(shapes.indexOf(s), 1);
    items.forEach((it, i) => { if (s.el.contains(it.div)) changed.delete(i); }); // its paragraphs go with it
    dirty = true; select(null);
  }

  renderBlock(body, page);
  if (!items.length) throw new Error(t('No text was found in this document.', 'ไม่พบข้อความในเอกสารนี้'));
  const opened = body.cloneNode(true); // the file as it opened, for Reset edits

  // ---------- toolbar ----------
  const fontsUsed = new Set();
  for (const f of xml.getElementsByTagNameNS(W, 'rFonts')) for (const k of ['ascii', 'cs']) { const v = attr(f, k); if (v) fontsUsed.add(v); }
  const fontList = [...fontsUsed].filter((f) => !FONTS.includes(f)).sort();
  const opt = (v, label = v) => `<option value="${String(v).replace(/"/g, '&quot;')}">${label}</option>`;
  // Preset colours next to each colour box: quicker on a touch screen, and there where a browser has no colour picker.
  const PRESETS = ['#000000', '#C00000', '#990000', '#0070C0', '#00B050', '#ED7D31'];
  const presets = (id) => `<span class="de-presets" data-for="${id}">${PRESETS.map((c) => `<button type="button" data-c="${c}" style="background:${c}" title="${c}" aria-label="${c}"></button>`).join('')}</span>`;
  const EDIT_NOTE = t('Select text, then pick a font, size or colour · edited paragraphs are marked on the left', 'เลือกข้อความ แล้วเลือกแบบอักษร ขนาด หรือสี · ย่อหน้าที่แก้ไขจะมีแถบสีส้มด้านซ้าย');
  root.innerHTML = `<div class="oe"><div class="oe-bar de-bar">
      <div class="de-group" role="group" aria-label="${t('Text', 'ข้อความ')}">
        <select id="de-font" title="${t('Font of the selected text', 'แบบอักษรของข้อความที่เลือก')}">${fontList.length ? `<optgroup label="${t('In this file', 'ในไฟล์นี้')}">${fontList.map((f) => opt(f)).join('')}</optgroup><optgroup label="${t('More fonts', 'แบบอักษรอื่น')}">` : ''}${FONTS.map((f) => opt(f)).join('')}${fontList.length ? '</optgroup>' : ''}</select>
        <select id="de-size" title="${t('Size of the selected text (pt)', 'ขนาดของข้อความที่เลือก (pt)')}">${SIZES.map((n) => opt(n)).join('')}</select>
        <label class="de-swatch" title="${t('Colour of the selected text', 'สีของข้อความที่เลือก')}">A <input type="color" id="de-color" value="${BOX.color}"></label>${presets('de-color')}
      </div>
      <div class="de-group" role="group" aria-label="${t('Insert', 'แทรก')}">
        <button type="button" data-tool="edit" aria-pressed="true" title="${t('Click text to edit it; click a shape to pick it', 'แตะข้อความเพื่อแก้ไข แตะรูปร่างเพื่อเลือก')}">${t('Edit', 'แก้ไข')}</button>
        <button type="button" data-tool="box" aria-pressed="false" title="${t('Drag on the page to add a text box', 'ลากบนหน้าเพื่อเพิ่มกล่องข้อความ')}">${t('Text box', 'กล่องข้อความ')}</button>
        <button type="button" data-tool="line" aria-pressed="false" title="${t('Drag to draw a line (Shift: 45° steps)', 'ลากเพื่อวาดเส้น (Shift: ทีละ 45°)')}">${t('Line', 'เส้น')}</button>
        <button type="button" data-tool="arrow" aria-pressed="false" title="${t('Drag to draw an arrow (Shift: 45° steps)', 'ลากเพื่อวาดลูกศร (Shift: ทีละ 45°)')}">${t('Arrow', 'ลูกศร')}</button>
        <label class="de-swatch" title="${t('Line colour (also the border of a picked text box)', 'สีเส้น (และขอบของกล่องข้อความที่เลือก)')}"><svg width="14" height="14" aria-hidden="true"><path d="M2 12 12 2" stroke="currentColor" stroke-width="2"/></svg> <input type="color" id="de-lcolor" value="${SHAPE.color}"></label>${presets('de-lcolor')}
        <select id="de-lwidth" title="${t('Line width', 'ความหนาเส้น')}">${opt(0, t('No line', 'ไม่มีเส้น'))}${WIDTHS.map((w) => opt(w, `${w} pt`)).join('')}</select>
        <button type="button" id="de-del" disabled title="${t('Delete the picked shape (Delete key)', 'ลบรูปร่างที่เลือก (ปุ่ม Delete)')}">${t('Delete', 'ลบ')}</button>
      </div>
      <div class="de-group" role="group" aria-label="${t('Undo', 'เลิกทำ')}">
        <button type="button" id="de-undo" disabled title="${t('Undo the last formatting or shape change (Ctrl+Z)', 'เลิกทำการจัดรูปแบบหรือรูปร่างล่าสุด (Ctrl+Z)')}">${t('Undo', 'เลิกทำ')}</button>
        <button type="button" id="de-redo" disabled title="${t('Redo (Ctrl+Y)', 'ทำซ้ำ (Ctrl+Y)')}">${t('Redo', 'ทำซ้ำ')}</button>
        <button type="button" id="de-reset" title="${t('Back to the file as it opened: every edit, text box, line and arrow goes (Undo brings them back)', 'กลับไปเป็นไฟล์ตอนเปิด: ลบการแก้ไข กล่องข้อความ เส้น และลูกศรทั้งหมด (เลิกทำเพื่อนำกลับมา)')}">${t('Reset edits', 'ล้างการแก้ไข')}</button>
      </div>
      <div class="de-group de-emode" role="group" aria-label="${t('New text', 'ข้อความใหม่')}"><span class="de-cap" aria-hidden="true">${t('New text:', 'ข้อความใหม่:')}</span>
        <button type="button" data-emode="markup" aria-pressed="true" title="${t('Typed text in the feedback colour', 'ข้อความที่พิมพ์เป็นสีความเห็น')}">${t('Markup', 'มาร์กอัป')}</button><button type="button" data-emode="edit" aria-pressed="false" title="${t('Typed text takes the format of the text before it', 'ข้อความที่พิมพ์ใช้รูปแบบของข้อความก่อนหน้า')}">${t('Edit', 'แก้ไข')}</button>
      </div>
      <div class="de-group" role="group" aria-label="${t('Page', 'หน้า')}">
        <button type="button" id="de-hf" aria-pressed="false" title="${t('Type in the headers and footers', 'พิมพ์ในหัวกระดาษและท้ายกระดาษ')}">${t('Header & footer', 'หัว/ท้ายกระดาษ')}</button><button type="button" id="de-pn" title="${t('Page numbers: format, first number, start and stop page', 'เลขหน้า: รูปแบบ เลขแรก หน้าเริ่มและหน้าสุดท้าย')}">${t('Page numbers', 'เลขหน้า')}</button>
      </div>
      <span class="oe-note" id="de-note">${EDIT_NOTE}</span>
    </div><div class="oe-scroll de-doc"></div></div>`;
  const $ = (id) => root.querySelector(`#${id}`);
  $('de-lwidth').value = String(SHAPE.width);
  // The colour and width new lines and arrows take (v2.16.2): the person's last choice. Picking a shape only shows that
  // shape's own line (a text box: black, No line) while it is picked; it never becomes the drawing default.
  const pen = { color: SHAPE.color, width: SHAPE.width };
  const lineColor = () => pen.color;
  const lineWidth = () => pen.width;
  const note = (text) => { $('de-note').textContent = text; };
  let markup = (() => { try { return localStorage.getItem('fb-de-markup') !== '0'; } catch { return true; } })(); // typed text in TYPED (Markup) or not (Edit)
  const capture = document.createElement('div'); capture.className = 'de-capture';
  page.append(layer, capture);
  // Fit to width: on a screen narrower than the page, the page is scaled down (pointer and layout measures are
  // divided by the zoom); a wrapper takes the scaled size so the scroll area matches.
  const scroller = root.querySelector('.de-doc'); const fit = document.createElement('div'); fit.className = 'de-fit';
  fit.append(page); scroller.append(fit);
  function fitWidth() {
    const room = scroller.clientWidth - 2 * (parseFloat(getComputedStyle(scroller).paddingLeft) || 0);
    zoom = room > 0 ? Math.min(1, room / page.offsetWidth) : 1;
    page.style.transform = zoom < 1 ? `scale(${zoom})` : '';
    fit.style.width = `${page.offsetWidth * zoom}px`; fit.style.height = `${page.offsetHeight * zoom}px`;
  }
  fitWidth();
  new ResizeObserver(fitWidth).observe(scroller);
  const textOf = (div) => div.innerText.replace(/\n$/, '');
  items.forEach((it) => { it.text = textOf(it.div); }); // read once the page is laid out
  shapes.forEach(draw);
  new ResizeObserver(() => { fitWidth(); layout(); }).observe(page); // images loading, edits that change heights
  document.fonts?.ready.then(layout);

  const changed = new Set(); // paragraphs whose typing is not in the file yet (an IME still composing)
  let typing = -1; // the paragraph of the typing under way: one undo step per paragraph typed in
  page.addEventListener('input', (e) => {
    const div = e.target.closest('.de-p'); if (!div) return;
    const i = +div.dataset.i;
    if (textOf(div) === items[i].text) { changed.delete(i); return; }
    changed.add(i);
    if (e.isComposing) { layout(); return; } // written when the composition ends
    live(i);
  });
  page.addEventListener('compositionend', (e) => { const div = e.target.closest?.('.de-p'); if (div && changed.has(+div.dataset.i)) live(+div.dataset.i); });
  // Typing into the file at once and the paragraph redrawn as saved (the typed words in TYPED), caret kept.
  function live(i) {
    if (typing !== i) { changed.delete(i); checkpoint(); typing = i; } // the copy before this typing, for Undo
    const it = items[i]; const s = getSelection();
    const at = s.rangeCount && it.div.contains(s.focusNode) ? lenTo(it.div, s.focusNode, s.focusOffset) : null;
    commit(i); changed.delete(i);
    const fresh = renderPara(it.p, i); it.div.replaceWith(fresh); fresh.classList.add('de-changed'); items[i].text = textOf(fresh);
    if (at != null) { const r = document.createRange(); r.setStart(...pointAt(fresh, at)); r.collapse(true); s.removeAllRanges(); s.addRange(r); }
    layout();
  }
  // Neighbouring text runs with the same formatting become one (typing letter by letter would leave a run per letter).
  function mergeRuns(p) {
    const plain = (r) => r?.namespaceURI === W && r.localName === 'r' && [...r.children].every((c) => (c.namespaceURI === W && c.localName === 'rPr') || isPiece(c));
    const fmt = (r) => { const x = kid(r, 'rPr'); return x ? new XMLSerializer().serializeToString(x) : ''; };
    let r = p.firstElementChild;
    while (r) {
      const next = r.nextElementSibling;
      if (plain(r) && plain(next) && fmt(r) === fmt(next)) { r.append(...[...next.children].filter(isPiece)); next.remove(); continue; }
      r = next;
    }
    for (const run of p.getElementsByTagNameNS(W, 'r')) {
      for (const t of [...run.children]) {
        const prev = t.previousElementSibling;
        if (t.localName === 't' && prev?.localName === 't' && t.namespaceURI === W) { prev.textContent += t.textContent; prev.setAttributeNS(XML_NS, 'xml:space', 'preserve'); t.remove(); }
      }
    }
  }
  // A paragraph's typed text into the file: the new characters in TYPED, unless in a text box (its own colours).
  function commit(i) {
    const it = items[i]; const text = textOf(it.div);
    const typed = setText(it.p, text);
    if (markup && !it.div.closest('.de-shape')) for (const [s, e] of typed) for (const run of isolate(it.p, s, e)) SET.color(TYPED)(rPrOf(run));
    mergeRuns(it.p);
    pieces.set(it.p, collect(it.p));
    it.text = text; dirty = true;
  }

  // ---------- undo / redo ----------
  // Before each formatting or shape change, a copy of the document body is kept; typed text is written into the file
  // first, so nothing typed is lost. Undo puts the copy back and redraws the page at the same scroll position.
  const undos = [], redos = [];
  function keepTyping() {
    for (const i of changed) commit(i);
    changed.clear();
  }
  function undoButtons() { $('de-undo').disabled = !undos.length; $('de-redo').disabled = !redos.length; }
  function checkpoint() { keepTyping(); typing = -1; undos.push(body.cloneNode(true)); if (undos.length > 50) undos.shift(); redos.length = 0; undoButtons(); }
  function restore(from, to) {
    if (!from.length) return;
    keepTyping(); typing = -1; to.push(body.cloneNode(true));
    const copy = from.pop(); body.replaceWith(copy); body = copy;
    const scroller = root.querySelector('.de-doc'); const top = scroller.scrollTop;
    select(null); saved = null;
    items.length = 0; divOf.clear(); shapes.length = 0; shapeOf.clear(); counters.clear(); scanPieces();
    page.replaceChildren(); layer.replaceChildren(); renderBlock(body, page); page.append(layer, capture); pagedLayers();
    items.forEach((it) => { it.text = textOf(it.div); }); shapes.forEach(draw);
    scroller.scrollTop = top; dirty = true; undoButtons();
  }
  const undo = () => restore(undos, redos);
  const redo = () => restore(redos, undos);
  $('de-undo').onclick = undo; $('de-redo').onclick = redo;
  // Reset edits (v2.16): back to the file as it opened in one step; Undo brings the edits back.
  $('de-reset').onclick = () => { restore([opened.cloneNode(true)], undos); redos.length = 0; undoButtons(); };

  // ---------- picking and moving shapes ----------
  function select(s) {
    const old = sel; sel = s;
    if (old && old !== s && shapes.includes(old)) draw(old);
    if (s) {
      draw(s);
      const ln = lineOf(s);
      if (/^#[0-9a-f]{6}$/i.test(ln.color)) $('de-lcolor').value = ln.color;
      const pt = ln.width * 0.75; $('de-lwidth').value = String(ln.none ? 0 : WIDTHS.reduce((b, w) => Math.abs(w - pt) < Math.abs(b - pt) ? w : b));
    } else { $('de-lcolor').value = pen.color; $('de-lwidth').value = String(pen.width); }
    $('de-del').disabled = !s;
  }
  const pagePoint = (e) => { const r = page.getBoundingClientRect(); return [(e.clientX - r.left) / zoom, (e.clientY - r.top) / zoom]; };
  const snap = (from, p, shift) => {
    if (!shift) return p;
    const dx = p[0] - from[0], dy = p[1] - from[1]; const ang = Math.round(Math.atan2(dy, dx) / (Math.PI / 4)) * (Math.PI / 4); const len = Math.hypot(dx, dy);
    return [from[0] + len * Math.cos(ang), from[1] + len * Math.sin(ang)];
  };
  function drag(e, onMove, onEnd = () => {}) {
    const start = pagePoint(e); let moved = false;
    const move = (ev) => { const p = pagePoint(ev); if (!moved && Math.hypot(p[0] - start[0], p[1] - start[1]) < 3) return; moved = true; onMove(p, start, ev); };
    // A touch the system takes over (pointercancel) ends the drag where it is.
    const up = () => { removeEventListener('pointermove', move); removeEventListener('pointerup', up); removeEventListener('pointercancel', up); onEnd(moved); layout(); };
    addEventListener('pointermove', move); addEventListener('pointerup', up); addEventListener('pointercancel', up);
  }
  layer.addEventListener('pointerdown', (e) => {
    const el = e.target.closest('.de-shape'); if (!el || e.button !== 0) return;
    const s = shapes.find((x) => x.el === el); if (!s) return;
    if (e.target.closest('.de-txbx')) { select(s); return; } // typing in a text box
    e.preventDefault(); getSelection().removeAllRanges(); root.focus({ preventScroll: true }); select(s);
    const g0 = { ...s.g }; const [a0, b0] = ends(g0); const h = e.target.dataset.h;
    let kept = false; const to = (g) => { if (!kept) { checkpoint(); kept = true; } place(s, g); }; // one undo step per drag
    if (e.target.classList.contains('de-size')) drag(e, (p, st0) => to({ ...g0, w: Math.max(24, g0.w + p[0] - st0[0]), h: Math.max(16, g0.h + p[1] - st0[1]) }));
    else if (h != null) drag(e, (p, _, ev) => to(h === '0' ? boxOf(snap(b0, p, ev.shiftKey), b0) : boxOf(a0, snap(a0, p, ev.shiftKey))));
    else drag(e, (p, st0) => to({ ...g0, x: g0.x + p[0] - st0[0], y: g0.y + p[1] - st0[1] }));
  });
  // The path as dispatched: picking a line or arrow redraws it, so its target is no longer in the page by now.
  page.addEventListener('pointerdown', (e) => { if (!e.composedPath().some((n) => n.matches?.('.de-shape, .de-capture'))) select(null); });
  root.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') { setTool('edit'); select(null); }
    // Ctrl+Z / Ctrl+Y: while there is typing not yet kept, the browser undoes the typing; otherwise these undo changes.
    const k = e.key.toLowerCase();
    if ((e.ctrlKey || e.metaKey) && (k === 'z' || k === 'y') && !(e.target.closest('.de-p') && changed.size) && !e.target.closest('input, select')) {
      e.preventDefault(); if (k === 'y' || e.shiftKey) redo(); else undo();
    }
    if ((e.key === 'Delete' || e.key === 'Backspace') && sel && !e.target.closest('.de-p, input, select')) { e.preventDefault(); removeShape(sel); }
  });
  root.tabIndex = -1; // so Delete reaches the editor after a shape is picked
  $('de-del').onclick = () => { if (sel) removeShape(sel); };

  // ---------- drawing new shapes ----------
  let tool = 'edit';
  function setTool(k) {
    tool = k;
    root.querySelectorAll('[data-tool]').forEach((b) => b.setAttribute('aria-pressed', String(b.dataset.tool === k)));
    page.classList.toggle('de-drawing', k !== 'edit');
    note(k === 'box' ? t('Drag on the page to add a text box, or tap for a standard size', 'ลากบนหน้าเพื่อเพิ่มกล่องข้อความ หรือแตะเพื่อใช้ขนาดมาตรฐาน') : k === 'edit' ? EDIT_NOTE
      : t('Drag on the page to draw · Shift keeps 45° steps · Esc or Edit to stop', 'ลากบนหน้าเพื่อวาด · Shift ให้เอียงทีละ 45° · กด Esc หรือแก้ไขเพื่อหยุด'));
  }
  root.querySelector('.de-bar').addEventListener('click', (e) => { const b = e.target.closest('[data-tool]'); if (b) setTool(b.dataset.tool); });
  capture.addEventListener('pointerdown', (e) => {
    if (e.button !== 0) return;
    e.preventDefault(); root.focus({ preventScroll: true });
    const kind = tool; const a = pagePoint(e); let b = a;
    const ghost = svgEl('svg', { class: 'de-ghost' }); capture.append(ghost);
    drag(e, (p, _, ev) => {
      b = kind === 'box' ? p : snap(a, p, ev.shiftKey);
      ghost.replaceChildren(kind === 'box'
        ? svgEl('rect', { x: Math.min(a[0], b[0]), y: Math.min(a[1], b[1]), width: Math.abs(b[0] - a[0]), height: Math.abs(b[1] - a[1]) })
        : svgEl('line', { x1: a[0], y1: a[1], x2: b[0], y2: b[1], stroke: lineColor(), 'stroke-width': lineWidth() / 0.75 }));
    }, (moved) => {
      ghost.remove();
      if (kind === 'box') {
        const g = moved && Math.abs(b[0] - a[0]) > 20 ? { ...boxOf(a, b), flipH: false, flipV: false } : { x: a[0], y: a[1], w: 200, h: 0 };
        const s = newShape('box', { ...g, h: Math.max(g.h, 30) });
        setTool('edit'); select(s);
        const div = s.el.querySelector('.de-p'); div?.focus();
      } else if (moved && Math.hypot(b[0] - a[0], b[1] - a[1]) > 4) select(newShape(kind, boxOf(a, b)));
    });
  });
  // A preset sets its colour box and applies it; pressing it keeps the text selection.
  root.querySelectorAll('.de-presets').forEach((g) => {
    g.addEventListener('pointerdown', (e) => e.preventDefault());
    g.addEventListener('click', (e) => {
      const b = e.target.closest('[data-c]'); if (!b) return;
      const input = $(g.dataset.for); input.value = b.dataset.c.toLowerCase(); input.dispatchEvent(new Event('change'));
    });
  });
  // A colour or width chosen here is the new default too (No line only applies to the picked shape).
  const onLine = () => {
    const color = $('de-lcolor').value; const width = +$('de-lwidth').value;
    pen.color = color; if (width) pen.width = width;
    if (sel) setLine(sel, color, width);
  };
  $('de-lcolor').addEventListener('change', onLine);
  $('de-lwidth').addEventListener('change', onLine);

  // ---------- text formatting ----------
  // The last text selection on the page is kept, as picking a font or colour takes the focus away from the page.
  let saved = null;
  const hex = (rgb) => { const m = rgb.match(/\d+/g); return m ? `#${m.slice(0, 3).map((n) => (+n).toString(16).padStart(2, '0')).join('')}` : '#000000'; };
  function show(node) { // the toolbar shows the font, size and colour at the selection
    const el = node.nodeType === 1 ? node : node.parentElement; const cs = getComputedStyle(el);
    const font = cs.fontFamily.split(',')[0].replace(/["']/g, '').trim(); const fs = $('de-font');
    if (![...fs.options].some((o) => o.value === font)) fs.append(new Option(font, font));
    fs.value = font;
    const pt = Math.round(parseFloat(cs.fontSize) * 0.75 * 2) / 2; const ss = $('de-size');
    if (![...ss.options].some((o) => +o.value === pt)) ss.append(new Option(pt, pt));
    ss.value = String(pt);
    $('de-color').value = hex(cs.color);
  }
  const onSelection = () => {
    if (!page.isConnected) { document.removeEventListener('selectionchange', onSelection); return; }
    const s = getSelection(); if (!s.rangeCount) return;
    const r = s.getRangeAt(0); const start = r.startContainer.nodeType === 1 ? r.startContainer : r.startContainer.parentElement;
    const p = start?.closest('.de-p'); if (!p || !page.contains(p)) return;
    saved = r.cloneRange(); show(r.startContainer);
    const box = p.closest('.de-shape'); select(box ? shapes.find((x) => x.el === box) : null);
  };
  document.addEventListener('selectionchange', onSelection);
  // Characters before a point in a paragraph (line breaks typed into the page may be <br>).
  function lenTo(div, node, offset) {
    const r = document.createRange(); r.setStart(div, 0); r.setEnd(node, offset);
    let n = 0; const w = document.createTreeWalker(r.cloneContents(), NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT);
    for (let c = w.nextNode(); c; c = w.nextNode()) n += c.nodeType === 3 ? c.length : c.nodeName === 'BR' ? 1 : 0;
    return n;
  }
  function pointAt(div, n) {
    const w = document.createTreeWalker(div, NodeFilter.SHOW_TEXT); let last = null;
    for (let t = w.nextNode(); t; t = w.nextNode()) { if (n <= t.length) return [t, n]; n -= t.length; last = t; }
    return last ? [last, last.length] : [div, div.childNodes.length];
  }
  // Split runs so that characters [s, e) of a paragraph are whole runs of their own; returns those runs.
  function isolate(p, s, e) {
    const cut = (at) => {
      let o = 0;
      for (const el of collect(p)) {
        const t = pieceText(el);
        if (el.localName === 't' && at > o && at < o + t.length) {
          const two = el.cloneNode(false); two.textContent = t.slice(at - o); el.textContent = t.slice(0, at - o);
          for (const x of [el, two]) x.setAttributeNS(XML_NS, 'xml:space', 'preserve');
          el.after(two); return;
        }
        o += t.length;
      }
    };
    cut(s); cut(e);
    const inside = new Set(); let o = 0;
    for (const el of collect(p)) { const n = pieceText(el).length; if (n && o >= s && o + n <= e) inside.add(el); o += n; }
    const runs = [];
    for (const run of new Set([...inside].map((el) => el.parentNode))) {
      const parts = []; let cur = null;
      for (const c of [...run.children]) {
        if (c.namespaceURI === W && c.localName === 'rPr') continue;
        const flag = isPiece(c) ? inside.has(c) : cur ? cur.flag : false; // other run content stays with what is before it
        if (!cur || flag !== cur.flag) parts.push(cur = { flag, nodes: [] });
        cur.nodes.push(c);
      }
      let prev = run;
      parts.forEach((part, k) => {
        let r = run;
        if (k) { r = xml.createElementNS(W, 'w:r'); const rPr = kid(run, 'rPr'); if (rPr) r.append(rPr.cloneNode(true)); r.append(...part.nodes); prev.after(r); prev = r; }
        if (part.flag) runs.push(r);
      });
    }
    return runs;
  }
  function prop(rPr, name, attrs, drop = []) {
    let el = kid(rPr, name);
    if (!el) {
      el = xml.createElementNS(W, `w:${name}`); const at = RPR_ORDER.indexOf(name);
      rPr.insertBefore(el, [...rPr.children].find((c) => c.namespaceURI === W && RPR_ORDER.indexOf(c.localName) > at) || null);
    }
    for (const a of drop) el.removeAttributeNS(W, a);
    for (const [k, v] of Object.entries(attrs)) el.setAttributeNS(W, `w:${k}`, String(v));
  }
  const rPrOf = (parent) => { let r = kid(parent, 'rPr'); if (!r) { r = xml.createElementNS(W, 'w:rPr'); parent.prepend(r); } return r; };
  function markRPr(p) { // the paragraph mark's run properties: the formatting of text typed into an empty paragraph
    let pPr = kid(p, 'pPr'); if (!pPr) { pPr = xml.createElementNS(W, 'w:pPr'); p.prepend(pPr); }
    let r = kid(pPr, 'rPr');
    if (!r) { r = xml.createElementNS(W, 'w:rPr'); pPr.insertBefore(r, kid(pPr, 'sectPr') || kid(pPr, 'pPrChange')); }
    return r;
  }
  const SET = {
    font: (f) => (rPr) => prop(rPr, 'rFonts', { ascii: f, hAnsi: f, eastAsia: f, cs: f }, ['asciiTheme', 'hAnsiTheme', 'eastAsiaTheme', 'cstheme']),
    size: (pt) => (rPr) => { prop(rPr, 'sz', { val: Math.round(pt * 2) }); prop(rPr, 'szCs', { val: Math.round(pt * 2) }); },
    color: (c) => (rPr) => prop(rPr, 'color', { val: c.slice(1).toUpperCase() }, ['themeColor', 'themeTint', 'themeShade']),
  };
  function format(set) {
    const r = saved?.startContainer.isConnected ? saved : null;
    let targets = [];
    if (sel?.kind === 'box' && (!r || r.collapsed || !sel.el.contains(r.startContainer))) {
      // A picked text box with no text selected in it: all its text.
      items.forEach((it, i) => { if (it.div.isConnected && sel.el.contains(it.div)) targets.push({ i, s: 0, e: Infinity, all: true }); });
    } else if (r && !r.collapsed) {
      items.forEach((it, i) => {
        const d = it.div; if (!d.isConnected || !r.intersectsNode(d)) return;
        const s = d.contains(r.startContainer) ? lenTo(d, r.startContainer, r.startOffset) : 0;
        const e = d.contains(r.endContainer) ? lenTo(d, r.endContainer, r.endOffset) : Infinity;
        if (e > s) targets.push({ i, s, e, all: s === 0 && e === Infinity });
      });
    }
    if (!targets.length) { note(t('Select some text first (or pick a text box), then choose the font, size or colour', 'เลือกข้อความก่อน (หรือเลือกกล่องข้อความ) แล้วจึงเลือกแบบอักษร ขนาด หรือสี')); return; }
    checkpoint();
    for (const t of targets) {
      const it = items[t.i];
      if (changed.has(t.i)) { commit(t.i); changed.delete(t.i); } // typed text first, so offsets match the file
      const len = collect(it.p).map(pieceText).join('').length; t.e = Math.min(t.e, len);
      for (const run of isolate(it.p, t.s, t.e)) set(rPrOf(run));
      if (t.all || !len) set(markRPr(it.p));
      pieces.set(it.p, collect(it.p));
      const fresh = renderPara(it.p, t.i); it.div.replaceWith(fresh); items[t.i].text = textOf(fresh);
    }
    dirty = true; layout();
    const first = targets[0], last = targets[targets.length - 1];
    if (!first.all) { // keep the same text selected, for the next change
      const range = document.createRange(); range.setStart(...pointAt(items[first.i].div, first.s)); range.setEnd(...pointAt(items[last.i].div, last.e));
      const s = getSelection(); s.removeAllRanges(); s.addRange(range); saved = range.cloneRange();
    }
  }
  $('de-font').addEventListener('change', (e) => format(SET.font(e.target.value)));
  $('de-size').addEventListener('change', (e) => format(SET.size(+e.target.value)));
  $('de-color').addEventListener('change', (e) => format(SET.color(e.target.value)));

  // Text -> run pieces: w:t for text, w:tab for tabs, w:br for line breaks.
  function textPieces(text) {
    const out = [];
    text.split('\n').forEach((line, li) => {
      if (li) out.push(xml.createElementNS(W, 'w:br'));
      line.split('\t').forEach((part, pi) => {
        if (pi) out.push(xml.createElementNS(W, 'w:tab'));
        if (!part) return;
        const t = xml.createElementNS(W, 'w:t'); t.setAttributeNS(XML_NS, 'xml:space', 'preserve'); t.textContent = part;
        out.push(t);
      });
    });
    return out;
  }
  // Replace a paragraph's text, keeping each run's formatting: only the part between the unchanged start and end is
  // rewritten, and new text takes the formatting of the character before it, as typing in Word does. A paragraph
  // with no text gets a new run in the paragraph mark's formatting. Returns the typed characters' [start, end) ranges.
  function setText(p, text) {
    const old = pieces.get(p) || [];
    if (!old.length) {
      const run = xml.createElementNS(W, 'w:r');
      const markRPr = kid(kid(p, 'pPr'), 'rPr');
      if (markRPr) {
        const rPr = markRPr.cloneNode(true);
        [...rPr.children].filter((c) => c.localName === 'ins' || c.localName === 'del').forEach((c) => c.remove());
        run.append(rPr);
      }
      const added = textPieces(text); run.append(...added); p.append(run);
      pieces.set(p, added);
      return text ? [[0, text.length]] : [];
    }
    const oldText = old.map(pieceText).join('');
    const { keep, ins } = diff(oldText, text);
    const result = []; let off = 0;
    old.forEach((el, n) => {
      const t = pieceText(el); const from = off; off += t.length;
      // This piece's kept characters, each followed by what was typed after it; text typed at the very start of the
      // paragraph goes into the first piece.
      let s = n === 0 ? ins[0] : ''; let same = !s;
      for (let k = from; k < off; k++) { if (keep[k]) s += oldText[k]; else same = false; if (ins[k + 1]) { s += ins[k + 1]; same = false; } }
      if (same) { result.push(el); return; }
      const repl = textPieces(s); el.replaceWith(...repl); result.push(...repl);
    });
    pieces.set(p, result);
    const typed = []; let j = 0; // where the inserted text sits in the new text
    for (let k = 0; k <= oldText.length; k++) { if (ins[k]) { typed.push([j, j + ins[k].length]); j += ins[k].length; } if (k < oldText.length && keep[k]) j++; }
    return typed;
  }
  // Character diff (longest common subsequence): keep[i] says whether old character i stays; ins[i] is the text
  // inserted just before old character i (ins[old.length]: at the end). Long paragraphs fall back to one changed span
  // between the unchanged start and end.
  function diff(a, b) {
    const n = a.length, m = b.length;
    const keep = new Array(n).fill(false); const ins = new Array(n + 1).fill('');
    if ((n + 1) * (m + 1) > 4e6) {
      let s = 0; while (s < n && s < m && a[s] === b[s]) s++;
      let e = 0; while (e < n - s && e < m - s && a[n - 1 - e] === b[m - 1 - e]) e++;
      for (let i = 0; i < n; i++) keep[i] = i < s || i >= n - e;
      ins[n - e] = b.slice(s, m - e);
      return { keep, ins };
    }
    const L = new Uint32Array((n + 1) * (m + 1)); const at = (i, j) => i * (m + 1) + j;
    for (let i = n - 1; i >= 0; i--) for (let j = m - 1; j >= 0; j--) L[at(i, j)] = a[i] === b[j] ? L[at(i + 1, j + 1)] + 1 : Math.max(L[at(i + 1, j)], L[at(i, j + 1)]);
    let i = 0, j = 0;
    while (i < n || j < m) {
      if (i < n && j < m && a[i] === b[j]) { keep[i] = true; i++; j++; }
      else if (j < m && (i === n || L[at(i, j + 1)] >= L[at(i + 1, j)])) { ins[i] += b[j]; j++; }
      else i++;
    }
    return { keep, ins };
  }

  // ---------- pages, headers and footers, page numbers (from TUSU Finance Team Tools' Word editor) ----------
  // The text is laid out in pages as Word does it, block by block: a paragraph or a table row that would cross the
  // bottom margin moves to the next page (with a transform, so typing into the file is untouched); page breaks, "page
  // break before" and section breaks start a new page. The pages are drawn behind the text, inside the page element
  // (so the fit-to-width scale applies to them too). Each page shows its section's header and footer, with PAGE,
  // NUMPAGES, IF and = fields worked out per page. Typing in a header or footer keeps the file's format.
  const GAP = 24; const W15 = (tw) => tw / 15;
  const secSz = () => kid(kid(body, 'sectPr'), 'pgSz');
  const PW = W15(+attr(secSz(), 'w') || 11906); const PG = W15(+attr(secSz(), 'h') || 16838);
  const frames = document.createElement('div'); frames.className = 'de-frames'; frames.setAttribute('aria-hidden', 'true');
  const hfLayer = document.createElement('div'); hfLayer.className = 'de-hfl';
  page.classList.add('de-paged');
  let pgT = 0; // the next layout of the pages, after a pause
  function pagedLayers() { page.prepend(frames); page.append(hfLayer); schedulePaginate(); } // also after Undo redraws the page
  pagedLayers();
  const settingsXml = await text('word/settings.xml');
  const evenOdd = /<w:evenAndOddHeaders(?![^>]*w:val="(?:0|false|off)")/.test(settingsXml || '');
  let relsXml = await zip.file('word/_rels/document.xml.rels')?.async('string') || '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
  let relsChanged = false;
  function addRel(type, target) {
    const ids = [...relsXml.matchAll(/Id="rId(\d+)"/g)].map((m) => +m[1]); const id = `rId${Math.max(0, ...ids) + 1}`;
    relsXml = relsXml.replace('</Relationships>', `<Relationship Id="${id}" Type="${type}" Target="${target}"/></Relationships>`); relsChanged = true; return id;
  }
  async function addContentType(partName, type) {
    let ct = await zip.file('[Content_Types].xml').async('string'); if (ct.includes(`PartName="${partName}"`)) return;
    zip.file('[Content_Types].xml', ct.replace('</Types>', `<Override PartName="${partName}" ContentType="${type}"/></Types>`));
  }
  const relMap = (s) => {
    const out = {};
    for (const m of (s || '').matchAll(/<Relationship\b[^>]*>/g)) {
      const id = /Id="([^"]+)"/.exec(m[0])?.[1]; const tg = /Target="([^"]+)"/.exec(m[0])?.[1];
      if (id && tg && !/TargetMode="External"/.test(m[0])) out[id] = tg.startsWith('/') ? tg.slice(1) : `word/${tg}`.replace(/[^/]+\/\.\.\//g, '');
    }
    return out;
  };
  const parts = new Map(); // zip path -> { path, doc, root, rels, dirty }
  async function loadPart(path) {
    if (!path) return null; if (parts.has(path)) return parts.get(path);
    const s = await zip.file(path)?.async('string'); if (!s) return null;
    const doc = parse(s);
    const part = { path, doc, root: doc.documentElement, rels: relMap(await zip.file(path.replace(/([^/]+)$/, '_rels/$1.rels'))?.async('string')), dirty: false };
    parts.set(path, part); return part;
  }

  // Fields in a header or footer paragraph (complex and simple, nested); the text inside a field is its last result.
  function hfSegments(p) {
    const segs = []; const stack = []; const own = [];
    const run = (r) => {
      for (const c of r.children) {
        if (c.namespaceURI !== W) continue;
        if (c.localName === 'fldChar') {
          const ty = attr(c, 'fldCharType');
          if (ty === 'begin') stack.push({ code: [], res: '', phase: 'code', r, start: r });
          else if (ty === 'separate') { if (stack.length) stack.at(-1).phase = 'res'; }
          else if (ty === 'end') { const f = stack.pop(); if (!f) continue; f.stop = r; if (stack.length) { if (stack.at(-1).phase === 'code') stack.at(-1).code.push(f); } else segs.push({ f }); }
        } else if (c.localName === 'instrText') { if (stack.at(-1)?.phase === 'code') stack.at(-1).code.push(c.textContent); }
        else if (isPiece(c)) { if (stack.length) { if (stack.length === 1 && stack[0].phase === 'res') stack[0].res += pieceText(c); } else { segs.push({ el: c, r }); own.push(c); } }
        else if (c.localName === 'drawing' && !stack.length) segs.push({ img: c, r });
      }
    };
    const visit = (el) => {
      for (const c of el.children) {
        if (c.namespaceURI !== W) continue;
        if (c.localName === 'r') run(c);
        else if (c.localName === 'fldSimple') {
          const f = { code: [attr(c, 'instr') || ''], res: [...c.getElementsByTagNameNS(W, 't')].map((x) => x.textContent).join(''), r: c.getElementsByTagNameNS(W, 'r')[0] || null, start: c, stop: c };
          if (stack.length) { if (stack.at(-1).phase === 'code') stack.at(-1).code.push(f); } else segs.push({ f });
        } else if (['hyperlink', 'ins', 'smartTag', 'sdt', 'sdtContent', 'customXml'].includes(c.localName)) visit(c);
      }
    };
    visit(p); return { segs, pieces: own };
  }
  const FMT_SWITCH = { roman: 'lowerRoman', ROMAN: 'upperRoman', alphabetic: 'lowerLetter', ALPHABETIC: 'upperLetter', arabic: 'decimal', thaiarabic: 'thaiNumbers', thailetter: 'thaiLetters' };
  function calc(s) {
    const tk = s.match(/\d+(?:\.\d+)?|[-+*/()]/g) || []; let i = 0;
    const prim = () => { const x = tk[i++]; if (x === '(') { const v = sum(); i++; return v; } if (x === '-') return -prim(); if (x === '+') return prim(); return x == null ? NaN : +x; };
    const prod = () => { let v = prim(); while (tk[i] === '*' || tk[i] === '/') { const o = tk[i++]; const w = prim(); v = o === '*' ? v * w : v / w; } return v; };
    const sum = () => { let v = prod(); while (tk[i] === '+' || tk[i] === '-') { const o = tk[i++]; const w = prod(); v = o === '+' ? v + w : v - w; } return v; };
    const v = sum(); return Number.isFinite(v) ? v : null;
  }
  function ifValue(s) {
    const tk = [...s.matchAll(/"([^"]*)"|(<=|>=|<>|=|<|>)|([^\s"<>=]+)/g)].map((m) => (m[1] != null ? { s: m[1] } : m[2] ? { op: m[2] } : { s: m[3] }));
    const [a, op, b, yes, no] = tk; if (!a || !op?.op || !b) return '';
    const nx = +a.s; const ny = +b.s; const num = a.s !== '' && b.s !== '' && !Number.isNaN(nx) && !Number.isNaN(ny);
    const c = num ? nx - ny : a.s.localeCompare(b.s);
    const ok = { '=': c === 0, '<>': c !== 0, '<': c < 0, '>': c > 0, '<=': c <= 0, '>=': c >= 0 }[op.op];
    return (ok ? yes?.s : no?.s) ?? '';
  }
  function fieldValue(f, ctx) {
    const code = f.code.map((x) => (typeof x === 'string' ? x : fieldValue(x, ctx))).join('').trim();
    const sw = /\\\*\s*(\w+)/.exec(code); const fmt = (d) => (sw ? FMT_SWITCH[sw[1]] || FMT_SWITCH[sw[1].toLowerCase()] || d : d);
    const bodyCode = code.replace(/\\[*#@!]\s*("[^"]*"|\S+)?/g, '').trim(); const name = (/^\S+/.exec(bodyCode)?.[0] || '').toUpperCase();
    if (name === 'PAGE') return formatNumber(ctx.num, fmt(ctx.fmt));
    if (name === 'NUMPAGES' || name === 'SECTIONPAGES') return formatNumber(ctx.pages, fmt('decimal'));
    if (name.startsWith('=')) { const v = calc(bodyCode.slice(1)); return v == null ? '' : formatNumber(Math.round(v), fmt('decimal')); }
    if (name === 'IF') return ifValue(bodyCode.slice(2));
    return f.res;
  }
  function renderHfPara(p, ctx, part, list) {
    const { p: pp, r: pr } = st.para(p);
    const div = document.createElement('div'); div.className = 'de-hp'; div.dataset.j = list.push(p) - 1;
    div.style.cssText = `${runCss(pr, false)};${paraCss(pp)}`;
    for (const sg of hfSegments(p).segs) {
      const rp = sg.r?.localName === 'r' ? st.run(sg.r, pr) : sg.f?.r ? st.run(sg.f.r, pr) : pr;
      if (sg.f) { const s = document.createElement('span'); s.className = 'de-fld'; s.contentEditable = 'false'; s.textContent = fieldValue(sg.f, ctx); s.style.cssText = runCss(rp, THAI.test(s.textContent)); div.append(s); continue; }
      if (sg.img) { const im = image(sg.img, part.rels); im.contentEditable = 'false'; div.append(im); continue; }
      const tx = pieceText(sg.el);
      if (tx === '\n') { div.append(document.createElement('br')); continue; }
      const s = document.createElement('span'); s.style.cssText = runCss(rp, THAI.test(tx)); s.textContent = tx; div.append(s);
    }
    if (!div.childNodes.length) div.append(document.createElement('br'));
    return div;
  }
  function renderHf(el, ctx, part, list, into) {
    for (const c of el.children) {
      if (c.namespaceURI !== W) continue;
      if (c.localName === 'p') into.append(renderHfPara(c, ctx, part, list));
      else if (c.localName === 'tbl') {
        const tb = document.createElement('table'); tb.className = `de-tbl${st.tableBorders(c) ? ' de-borders' : ''}`; tb.style.width = '100%';
        for (const tr of kids(c, 'tr')) { const row = tb.insertRow(); for (const tc of kids(tr, 'tc')) { const td = row.insertCell(); const w = kid(kid(tc, 'tcPr'), 'tcW'); if (attr(w, 'type') === 'dxa') td.style.width = `${W15(+attr(w, 'w'))}px`; renderHf(tc, ctx, part, list, td); } }
        into.append(tb);
      } else if (c.localName === 'sdt') renderHf(kid(c, 'sdtContent') || c, ctx, part, list, into);
    }
  }

  // Sections, from the blocks on the page: a paragraph holding a sectPr ends its section; the body's ends the last.
  const xmlOf = (el) => (el.classList.contains('de-p') ? items[+el.dataset.i]?.p : null);
  function sections() {
    const blocks = []; const sps = []; let s = 0;
    for (const el of page.children) {
      if (!el.classList.contains('de-p') && !el.classList.contains('de-tbl')) continue;
      blocks.push({ el, sec: s });
      const sp = el.classList.contains('de-p') && kid(kid(xmlOf(el), 'pPr'), 'sectPr'); if (sp) { sps.push(sp); s++; }
    }
    sps.push(kid(body, 'sectPr'));
    const secs = [];
    sps.forEach((sp, i) => {
      const prev = secs[i - 1]; const mar = kid(sp, 'pgMar');
      const refs = (tag) => { const o = { ...(prev?.[tag] || {}) }; for (const r of kids(sp, `${tag}Reference`)) o[attr(r, 'type') || 'default'] = rels[r.getAttributeNS(R, 'id')] || null; return o; };
      const pn = kid(sp, 'pgNumType'); const tp = kid(sp, 'titlePg');
      secs.push({ sp, header: refs('header'), footer: refs('footer'), titlePg: !!tp && on(tp), fmt: attr(pn, 'fmt') || 'decimal', start: attr(pn, 'start') != null ? +attr(pn, 'start') : null,
        type: val(kid(sp, 'type')) || 'nextPage', top: W15(+attr(mar, 'top') || 1440), bottom: W15(+attr(mar, 'bottom') || 1440), hd: W15(+attr(mar, 'header') || 720), fd: W15(+attr(mar, 'footer') || 720),
        left: W15(+attr(mar, 'left') || 1440), right: W15(+attr(mar, 'right') || 1440) });
    });
    return { blocks, secs };
  }
  function breaksOf(el) {
    const p = el.classList.contains('de-p') && xmlOf(el); if (!p) return {};
    const pbb = kid(kid(p, 'pPr'), 'pageBreakBefore');
    return { before: !!pbb && on(pbb), after: [...p.getElementsByTagNameNS(W, 'br')].some((b) => attr(b, 'type') === 'page' && ancestor(b, 'p') === p) };
  }
  let pageList = []; let lastSecs = [];
  function paginate() {
    const { blocks, secs } = sections(); if (!blocks.length) return;
    const units = [];
    blocks.forEach(({ el, sec }, bi) => {
      const newSec = bi > 0 && blocks[bi - 1].sec !== sec && secs[sec].type !== 'continuous';
      const br = breaksOf(el);
      if (el.tagName === 'TABLE') [...el.rows].forEach((tr, ri) => units.push({ el: tr, sec, force: ri === 0 && newSec }));
      else units.push({ el, sec, force: newSec || br.before, after: br.after });
    });
    units.forEach((u) => { u.el.style.transform = ''; u.el.classList?.remove('de-pgtop'); });
    const pr = page.getBoundingClientRect(); const z = pr.width / page.offsetWidth || 1;
    const nat = units.map((u) => { const r = u.el.getBoundingClientRect(); return [(r.top - pr.top) / z, r.height / z]; });
    const stride = PG + GAP; const cTop = (k, s) => k * stride + secs[s].top; const cBot = (k, s) => k * stride + PG - secs[s].bottom;
    const pages = [{ sec: units[0].sec, used: false }]; let k = 0; let push = 0; let breakNext = false;
    units.forEach((u, n) => {
      let top = nat[n][0] + push; const h = nat[n][1];
      const next = () => { k++; pages[k] = { sec: u.sec, used: false }; push += cTop(k, u.sec) - top; top = cTop(k, u.sec); };
      if ((u.force || breakNext) && pages[k].used) next();
      while (pages[k].used && top + h > cBot(k, pages[k].sec) + 0.5) next();
      if (u.el.tagName === 'TR' && !pages[k].used && n && units[n - 1].el.parentElement === u.el.parentElement) u.el.classList.add('de-pgtop');
      pages[k].used = true; breakNext = !!u.after;
      if (push) u.el.style.transform = `translateY(${push}px)`;
    });
    let num = 0;
    pages.forEach((pg, i) => {
      const firstOfSec = i === 0 || pages[i - 1].sec !== pg.sec; const s = secs[pg.sec];
      num = firstOfSec && s.start != null ? s.start : i === 0 ? 1 : num + 1;
      Object.assign(pg, { num, fmt: s.fmt, first: firstOfSec });
    });
    pageList = pages; lastSecs = secs;
    page.style.minHeight = `${pages.length * stride - GAP}px`;
    frames.innerHTML = pages.map((_, i) => `<div class="de-pf" style="top:${i * stride}px;width:${PW}px;height:${PG}px"></div>`).join('');
    drawHf(secs); layout();
  }
  function drawHf(secs) {
    if (hfLayer.querySelector('.de-hfbox:focus')) return; // not while one is being typed in
    const stride = PG + GAP; const total = pageList.length; const editing = page.classList.contains('de-hfedit');
    hfLayer.replaceChildren();
    pageList.forEach((pg, i) => {
      const s = secs[pg.sec]; const type = pg.first && s.titlePg ? 'first' : evenOdd && pg.num % 2 === 0 ? 'even' : 'default';
      for (const kind of ['header', 'footer']) {
        const path = s[kind][type] ?? null;
        const box = document.createElement('div'); box.className = `de-hfbox de-${kind}`; box.dataset.kind = kind; box.dataset.sec = pg.sec; box.dataset.type = type;
        box.style.cssText = `left:${s.left}px;width:${PW - s.left - s.right}px;${kind === 'header' ? `top:${i * stride + s.hd}px` : `top:${i * stride + PG - s.fd}px;transform:translateY(-100%)`}`;
        if (editing) { box.contentEditable = 'plaintext-only'; if (box.contentEditable !== 'plaintext-only') box.contentEditable = 'true'; }
        box.setAttribute('aria-label', `${kind === 'header' ? t('Header', 'หัวกระดาษ') : t('Footer', 'ท้ายกระดาษ')}, ${t('page', 'หน้า')} ${i + 1}`);
        const part = path && parts.get(path);
        if (part) { box.dataset.path = path; const list = []; box._paras = list; renderHf(part.root, { num: pg.num, fmt: pg.fmt, pages: total }, part, list, box); }
        hfLayer.append(box);
      }
    });
  }
  function schedulePaginate() { clearTimeout(pgT); pgT = setTimeout(() => requestAnimationFrame(paginate), 120); }
  for (const path of Object.values(rels)) if (/\/(header|footer)\d*\.xml$/.test(path)) await loadPart(path);
  document.fonts?.ready.then(paginate); paginate();
  // the text changed (typing, Undo, Save redrawing a paragraph): lay the pages out again; not for the layers' own changes
  new MutationObserver((ms) => { if (ms.some((m) => { const el = m.target.nodeType === 1 ? m.target : m.target.parentElement; return !el?.closest('.de-frames, .de-hfl, .de-layer'); })) schedulePaginate(); })
    .observe(page, { childList: true, subtree: true, characterData: true });

  // ---------- typing in headers and footers (their text keeps the file's format) ----------
  const hfBtn = root.querySelector('#de-hf'); const pnBtn = root.querySelector('#de-pn');
  function hfMode(onOff) {
    page.classList.toggle('de-hfedit', onOff); hfBtn.setAttribute('aria-pressed', String(onOff));
    hfBtn.textContent = onOff ? t('Close header & footer', 'ปิดหัว/ท้ายกระดาษ') : t('Header & footer', 'หัว/ท้ายกระดาษ');
    paginate();
    note(onOff ? t('Header & footer: type in them (the text keeps its format). Tap Close header & footer when done.', 'หัว/ท้ายกระดาษ: พิมพ์ได้เลย (ข้อความคงรูปแบบเดิม) แตะ ปิดหัว/ท้ายกระดาษ เมื่อเสร็จ') : EDIT_NOTE);
  }
  hfBtn.onclick = () => hfMode(!page.classList.contains('de-hfedit'));
  hfLayer.addEventListener('keydown', (e) => { if (e.key === 'Escape') { e.preventDefault(); e.target.blur?.(); hfMode(false); } else if (e.key === 'Enter') { e.preventDefault(); document.execCommand('insertLineBreak'); } });
  let hfBusy = Promise.resolve();
  hfLayer.addEventListener('input', (e) => { const b = e.target.closest('.de-hfbox'); if (b) hfBusy = hfBusy.then(() => commitBox(b)).catch((err) => note(err.message)); });
  hfLayer.addEventListener('focusout', (e) => { const b = e.target.closest?.('.de-hfbox'); if (b && !b.contains(e.relatedTarget)) setTimeout(() => { if (!hfLayer.querySelector('.de-hfbox:focus')) paginate(); }, 0); });
  async function commitBox(b) {
    let part = parts.get(b.dataset.path);
    if (!part) { part = await newPart(b.dataset.kind, +b.dataset.sec, b.dataset.type); b.dataset.path = part.path; b._paras = [kid(part.root, 'p')]; }
    const n = getSelection().anchorNode; const d = (n?.nodeType === 3 ? n.parentElement : n)?.closest?.('.de-hp');
    const targets = d && b.contains(d) ? [[b._paras[+d.dataset.j], d]] : !b.querySelector('.de-hp') ? [[b._paras[0], b]] : [];
    for (const [p, div] of targets) if (p) commitHfPara(p, div);
    part.dirty = true; dirty = true;
  }
  function commitHfPara(p, div) {
    const gaps = [[]];
    const walk = (nd) => { for (const c of nd.childNodes) { if (c.nodeType === 3) gaps.at(-1).push(c.data.replace(/ /g, ' ')); else if (c.classList?.contains('de-fld')) gaps.push([]); else if (c.nodeName === 'BR') { if (!(c === nd.lastChild && nd === div)) gaps.at(-1).push('\n'); } else if (!c.classList?.contains('de-img')) walk(c); } };
    walk(div);
    const texts = gaps.map((g) => g.join(''));
    const fileGaps = [[]]; const fields = [];
    hfSegments(p).segs.forEach((s) => { if (s.f) { fields.push(s.f); fileGaps.push([]); } else if (s.el) fileGaps.at(-1).push(s.el); });
    if (fileGaps.length !== texts.length) return;
    texts.forEach((tx, i) => {
      const els = fileGaps[i]; if (tx === els.map(pieceText).join('')) return;
      if (els.length || !fields.length) { pieces.set(p, els); setText(p, tx); return; } // the file's format, not TYPED
      const f = fields[i] || fields[i - 1];
      const run = xml.createElementNS(W, 'w:r'); const rp = f.r?.localName === 'r' && kid(f.r, 'rPr'); if (rp) run.append(rp.cloneNode(true)); run.append(...textPieces(tx));
      if (fields[i]) f.start.before(run); else f.stop.after(run);
    });
  }
  async function newPart(kind, secIndex, type = 'default') {
    let n = 1; while (zip.file(`word/${kind}${n}.xml`)) n++;
    const name = `${kind}${n}.xml`; const tag = kind === 'header' ? 'hdr' : 'ftr';
    zip.file(`word/${name}`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<w:${tag} xmlns:w="${W}" xmlns:r="${R}"><w:p/></w:${tag}>`);
    await addContentType(`/word/${name}`, `application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}+xml`);
    const rid = addRel(`http://schemas.openxmlformats.org/officeDocument/2006/relationships/${kind}`, name); rels[rid] = `word/${name}`;
    const sp = sections().secs[secIndex].sp;
    kids(sp, `${kind}Reference`).filter((r) => (attr(r, 'type') || 'default') === type).forEach((r) => r.remove());
    const ref = xml.createElementNS(W, `w:${kind}Reference`); ref.setAttributeNS(W, 'w:type', type); ref.setAttributeNS(R, 'r:id', rid); sp.prepend(ref);
    const part = await loadPart(`word/${name}`); part.dirty = true; dirty = true; return part;
  }

  // ---------- page numbers: Word fields (PAGE inside IF and = fields), so Word numbers every page itself ----------
  const PN_FMT = [['decimal', '1, 2, 3', 'Arabic'], ['lowerRoman', 'i, ii, iii', 'roman'], ['upperRoman', 'I, II, III', 'ROMAN'], ['lowerLetter', 'a, b, c', 'alphabetic'], ['upperLetter', 'A, B, C', 'ALPHABETIC'], ['thaiNumbers', '๑, ๒, ๓', 'ThaiArabic'], ['none', t('No page numbers', 'ไม่มีเลขหน้า'), '']];
  const PN_MARK = '_TusuPageNumber';
  function pnFields(o) {
    const sw = PN_FMT.find(([f]) => f === o.fmt)[2]; const off = o.first - o.from; const plus = `${off < 0 ? '-' : '+'} ${Math.abs(off)}`;
    const ph = () => ({ code: ['PAGE \\* Arabic'], res: String(o.from) });
    const number = off ? { code: ['= ', ph(), ` ${plus} \\* ${sw}`], res: formatNumber(o.first, o.fmt) } : { code: [`PAGE \\* ${sw}`], res: formatNumber(o.first, o.fmt) };
    const total = !o.total ? null : o.to ? formatNumber(o.first + o.to - o.from, o.fmt) : { code: ['= ', { code: ['NUMPAGES \\* Arabic'], res: '1' }, ` ${plus} \\* ${sw}`], res: '' };
    let seq = [o.before, number, ...(total ? [o.of, total] : [])].filter((x) => x !== '');
    if (o.from > 1) seq = [{ code: ['IF ', ph(), ` >= ${o.from} "`, ...seq, '" ""'], res: formatNumber(o.first, o.fmt) }];
    if (o.to) seq = [{ code: ['IF ', ph(), ` <= ${o.to} "`, ...seq, '" ""'], res: formatNumber(o.first, o.fmt) }];
    return seq;
  }
  function pnRuns(doc, seq) {
    const out = [];
    const mk = (name, attrs = {}, textC = null) => { const e = doc.createElementNS(W, `w:${name}`); Object.entries(attrs).forEach(([k2, v]) => e.setAttributeNS(W, `w:${k2}`, v)); if (textC != null) { e.setAttributeNS(XML_NS, 'xml:space', 'preserve'); e.textContent = textC; } return e; };
    const run = (child) => { const r = mk('r'); r.append(child); out.push(r); };
    const field = (f) => { run(mk('fldChar', { fldCharType: 'begin' })); f.code.forEach((c) => { if (typeof c === 'string') run(mk('instrText', {}, c)); else field(c); }); run(mk('fldChar', { fldCharType: 'separate' })); if (f.res) run(mk('t', {}, f.res)); run(mk('fldChar', { fldCharType: 'end' })); };
    seq.forEach((x) => (typeof x === 'string' ? run(mk('t', {}, x)) : field(x)));
    return out;
  }
  const PN_ONLY = /^[\s\-–—|/·.()]*(?:page|หน้า(?:ที่)?)?[\s\-–—|/·.()]*(?:of|จาก)?[\s\-–—|/·.()]*$/i;
  const codeText = (f) => f.code.map((c) => (typeof c === 'string' ? c : codeText(c))).join('');
  async function applyPageNumbers(o) {
    checkpoint();
    const { secs } = sections(); const kind = o.where;
    if (o.fmt !== 'none') secs.forEach((s) => { const pn = kid(s.sp, 'pgNumType'); if (!pn) return; pn.removeAttributeNS(W, 'start'); pn.removeAttributeNS(W, 'fmt'); if (!pn.attributes.length) pn.remove(); });
    const paths = new Set(); secs.forEach((s) => { if (s[kind].default) paths.add(s[kind].default); if (evenOdd && s[kind].even) paths.add(s[kind].even); });
    if (!secs[0][kind].default && o.fmt !== 'none') paths.add((await newPart(kind, 0)).path);
    let kept = 0;
    for (const path of paths) {
      const part = await loadPart(path); if (!part) continue;
      for (const p of kids(part.root, 'p')) {
        const ours = [...p.getElementsByTagNameNS(W, 'bookmarkStart')].some((b) => attr(b, 'name') === PN_MARK);
        const { segs } = hfSegments(p); const hasPage = segs.some((s) => s.f && /\bPAGE\b/i.test(codeText(s.f)));
        const own = segs.filter((s) => s.el).map((s) => pieceText(s.el)).join('');
        if (ours || (hasPage && PN_ONLY.test(own))) p.remove(); else if (hasPage) kept++;
      }
      if (o.fmt !== 'none') {
        const d = part.doc; const p = d.createElementNS(W, 'w:p'); const pPr = d.createElementNS(W, 'w:pPr');
        const jc = d.createElementNS(W, 'w:jc'); jc.setAttributeNS(W, 'w:val', o.align); pPr.append(jc); p.append(pPr);
        const id = String(2147000000 + Math.floor(Math.random() * 400000));
        const bs = d.createElementNS(W, 'w:bookmarkStart'); bs.setAttributeNS(W, 'w:id', id); bs.setAttributeNS(W, 'w:name', PN_MARK);
        const be = d.createElementNS(W, 'w:bookmarkEnd'); be.setAttributeNS(W, 'w:id', id);
        p.append(bs, ...pnRuns(d, pnFields(o)), be);
        if (kind === 'header') part.root.prepend(p); else part.root.append(p);
      }
      if (!kids(part.root, 'p').length && !kids(part.root, 'tbl').length) part.root.append(part.doc.createElementNS(W, 'w:p'));
      part.dirty = true;
    }
    dirty = true; paginate();
    note(o.fmt === 'none' ? t('Page numbers removed', 'ลบเลขหน้าแล้ว') : t(`Page numbers added in the ${kind}`, `เพิ่มเลขหน้าใน${kind === 'header' ? 'หัวกระดาษ' : 'ท้ายกระดาษ'}แล้ว`) + (kept ? t(` (${kept} line(s) with a number and other text kept)`, ` (คงบรรทัดที่มีเลขหน้าและข้อความอื่นไว้ ${kept} บรรทัด)`) : ''));
  }
  const pop = document.createElement('div'); pop.className = 'de-pop'; pop.hidden = true; pop.setAttribute('role', 'dialog'); pop.setAttribute('aria-labelledby', 'de-pn-title'); root.querySelector('.oe').append(pop);
  function pnDialog() {
    const n = Math.max(1, pageList.length);
    pop.innerHTML = `<form class="de-sform"><div class="de-shead"><b id="de-pn-title">${t('Page numbers', 'เลขหน้า')}</b><button type="button" data-x aria-label="${t('Close', 'ปิด')}">✕</button></div>
      <div class="de-srow"><label>${t('Where', 'ตำแหน่ง')} <select name="where"><option value="footer">${t('Footer', 'ท้ายกระดาษ')}</option><option value="header">${t('Header', 'หัวกระดาษ')}</option></select></label>
        <label>${t('Align', 'จัดแนว')} <select name="align"><option value="center">${t('Centre', 'กึ่งกลาง')}</option><option value="right">${t('Right', 'ขวา')}</option><option value="left">${t('Left', 'ซ้าย')}</option></select></label></div>
      <label>${t('Format', 'รูปแบบ')} <select name="fmt">${PN_FMT.map(([f, l]) => `<option value="${f}">${l}</option>`).join('')}</select></label>
      <label>${t('Text before', 'ข้อความนำหน้า')} <input name="before" type="text" placeholder="หน้า / Page" maxlength="20"></label>
      <label class="de-chk"><input name="total" type="checkbox"> ${t('Show the total (of N)', 'แสดงจำนวนหน้าทั้งหมด (จาก N)')}</label>
      <div class="de-srow"><label>${t('Start on page', 'เริ่มที่หน้า')} <input name="from" type="number" min="1" max="${n}" value="1"></label><label>${t('First number', 'เลขแรก')} <input name="first" type="number" min="0" max="9999" value="1"></label></div>
      <label>${t('Stop after page', 'หยุดหลังหน้า')} <input name="to" type="number" min="1" max="${n}" placeholder="${t('to the end', 'จนจบ')}"></label>
      <p class="de-pnprev" aria-live="polite"></p>
      <p class="de-pnnote">${t('Pages before the start page and after the stop page have no number. Word works the numbers out when it opens the file.', 'หน้าก่อนหน้าเริ่มและหลังหน้าหยุดจะไม่มีเลข Word คำนวณเลขหน้าเองเมื่อเปิดไฟล์')}</p>
      <div class="de-srow"><button type="submit" class="btn primary">${t('Apply', 'ใช้')}</button></div></form>`;
    const f = pop.querySelector('form');
    const close = () => { pop.hidden = true; pop.onkeydown = null; pnBtn.focus(); };
    pop.querySelector('[data-x]').onclick = close;
    pop.onkeydown = (e) => { if (e.key === 'Escape') { e.preventDefault(); close(); } };
    const preview = () => {
      const E = f.elements; const fmt = E.fmt.value; const out = pop.querySelector('.de-pnprev');
      if (fmt === 'none') { out.textContent = t('The page numbers this tool wrote will be taken out.', 'จะลบเลขหน้าที่เครื่องมือนี้ใส่ไว้'); return; }
      const from = Math.max(1, Math.round(+E.from.value || 1)); const first = Math.max(0, Math.round(+E.first.value || 0)); const to = E.to.value ? Math.max(from, Math.round(+E.to.value)) : 0;
      const before = E.before.value.replace(/"/g, ''); const pre = before && !/\s$/.test(before) ? `${before} ` : before;
      const tot = E.total.checked ? `${/[฀-๿]/.test(before) ? ' จาก ' : ' of '}${formatNumber(first + (to || n) - from, fmt)}` : '';
      out.textContent = `${t('On page', 'หน้า')} ${from}: ${pre}${formatNumber(first, fmt)}${tot}`;
    };
    f.addEventListener('input', preview); preview();
    f.onsubmit = async (e) => {
      e.preventDefault(); const E = f.elements; const from = Math.max(1, Math.round(+E.from.value || 1)); const to = E.to.value ? Math.max(from, Math.round(+E.to.value)) : 0;
      const before = E.before.value.replace(/"/g, ''); const thai = /[฀-๿]/.test(before);
      close();
      try { await applyPageNumbers({ where: E.where.value, align: E.align.value, fmt: E.fmt.value, before: before && !/\s$/.test(before) ? `${before} ` : before, total: E.total.checked, of: thai ? ' จาก ' : ' of ', from, first: Math.max(0, Math.round(+E.first.value || 0)), to }); }
      catch (err) { note(err.message); }
    };
    pop.hidden = false;
    const r = pnBtn.getBoundingClientRect(); const bx = root.querySelector('.oe').getBoundingClientRect();
    pop.style.left = `${Math.max(8, Math.min(r.left - bx.left, bx.width - 330))}px`; pop.style.top = `${r.bottom - bx.top + 8}px`;
    f.elements.where.focus();
  }
  pnBtn.onclick = pnDialog;

  // ---------- Edit / Markup: what typed text looks like ----------
  // Markup (as before): typed text is in the feedback colour (TYPED). Edit: it takes the format of the text before it.
  // Kept in this browser.
  function showMode() { root.querySelectorAll('[data-emode]').forEach((b) => b.setAttribute('aria-pressed', String((b.dataset.emode === 'markup') === markup))); }
  root.querySelector('.de-emode').onclick = (e) => {
    const b = e.target.closest('[data-emode]'); if (!b) return; markup = b.dataset.emode === 'markup';
    try { localStorage.setItem('fb-de-markup', markup ? '1' : '0'); } catch { /* not kept */ }
    showMode(); note(markup ? t('Markup: typed text is in the feedback colour', 'มาร์กอัป: ข้อความที่พิมพ์เป็นสีความเห็น') : t('Edit: typed text takes the format of the text before it', 'แก้ไข: ข้อความที่พิมพ์ใช้รูปแบบของข้อความก่อนหน้า'));
  };
  showMode();

  return {
    isDirty: () => changed.size > 0 || dirty,
    async save() {
      for (const i of changed) {
        const it = items[i];
        commit(i);
        const fresh = renderPara(it.p, i); it.div.replaceWith(fresh); // show the paragraph as it is saved
        items[i].text = textOf(fresh);
      }
      changed.clear();
      // A text box that grows with its text is saved at the height it has here; Word fits it again when it is edited.
      for (const s of shapes) if (s.auto && s.g && Math.abs(s.el.offsetHeight - s.g.h) > 1) place(s, { ...s.g, h: s.el.offsetHeight });
      dirty = false;
      let out = new XMLSerializer().serializeToString(xml);
      if (!out.startsWith('<?xml')) out = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n${out}`;
      zip.file('word/document.xml', out);
      for (const pt of parts.values()) { // headers and footers typed in or given page numbers
        if (!pt.dirty) continue;
        let px = new XMLSerializer().serializeToString(pt.doc); if (!px.startsWith('<?xml')) px = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n${px}`;
        zip.file(pt.path, px); pt.dirty = false;
      }
      if (relsChanged) zip.file('word/_rels/document.xml.rels', relsXml);
      return zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
    },
  };
}
