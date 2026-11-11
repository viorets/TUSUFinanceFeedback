// Text boxes written into the PDF as real text (searchable, copyable, sharp), placed where the browser draws them.
// Thai is shaped with the font's own rules (fontkit: GSUB + GPOS mark positioning) and each glyph is placed at its
// shaped offset, because pdf-lib's drawText drops mark positions (tone marks then collide with upper vowels).
//
// Font files: public/fonts/ (Firebase Hosting). Times New Roman and Arial use the PDF's built-in Times and Helvetica
// (same letter widths; nothing to embed). EucrosiaUPC is a Windows font we may not host: Chrome/Edge on a PC can read
// it from the computer (Local Font Access, with the user's permission); elsewhere it falls back to TH Sarabun New.
// Georgia has no file we may embed: those boxes stay a picture (js/pdf-markup.js textPng).
// Fonts from the user's computer (addFontFile): kept in this browser only (IndexedDB), shown with FontFace and
// embedded like the site's own; a file of EucrosiaUPC replaces reading it from the computer (so Safari can use it).

let FONT_DIR = new URL('../fonts/', import.meta.url).href;
// Where the font files are, when this module is not in js/ beside fonts/ (the Feedback site keeps it in js/markup/).
export const setFontDir = (url) => { FONT_DIR = new URL(url, document.baseURI).href.replace(/\/?$/, '/'); };
const FONTKIT = 'https://cdn.jsdelivr.net/npm/@pdf-lib/fontkit@1.1.1/dist/fontkit.umd.min.js';

// family -> files per style (r, b, i, bi); a missing style is made from the nearest one (bold: outline stroke,
// italic: slant), as the browser does on screen. fallback: the font for letters this one lacks.
export const FACES = {
  'LINE Seed Sans TH': { r: 'LINESeedSansTH_Rg.ttf', fallback: 'Sarabun' },
  'Noto Sans Thai': { r: 'NotoSansThai-Regular.ttf', b: 'NotoSansThai-Bold.ttf', fallback: 'Sarabun' },
  Sarabun: { r: 'Sarabun-Regular.ttf', b: 'Sarabun-Bold.ttf', i: 'Sarabun-Italic.ttf', bi: 'Sarabun-BoldItalic.ttf', fallback: 'TH Sarabun New' },
  'TH Sarabun New': { r: 'THSarabunNew-Regular.ttf', b: 'THSarabunNew-Bold.ttf', i: 'THSarabunNew-Italic.ttf', bi: 'THSarabunNew-BoldItalic.ttf', fallback: 'Sarabun' },
  'TH SarabunPSK': { r: 'THSarabunPSK-Regular.ttf', b: 'THSarabunPSK-Bold.ttf', i: 'THSarabunPSK-Italic.ttf', bi: 'THSarabunPSK-BoldItalic.ttf', fallback: 'TH Sarabun New' },
  EucrosiaUPC: { local: ['EucrosiaUPC', 'EucrosiaUPC-Bold', 'EucrosiaUPC-Italic', 'EucrosiaUPC-BoldItalic'], fallback: 'TH Sarabun New' },
  'Angsana New': { local: ['AngsanaNew', 'AngsanaNew-Bold', 'AngsanaNew-Italic', 'AngsanaNew-BoldItalic'], fallback: 'TH Sarabun New' },
  'Times New Roman': { std: ['TimesRoman', 'TimesRomanBold', 'TimesRomanItalic', 'TimesRomanBoldItalic'], fallback: 'TH Sarabun New' },
  Arial: { std: ['Helvetica', 'HelveticaBold', 'HelveticaOblique', 'HelveticaBoldOblique'], fallback: 'TH Sarabun New' },
};
// The CSS font list for a family on screen: the same fallback the PDF uses, so both look alike.
export const fontStack = (family) => {
  const f = FACES[family]; const fb = f?.fallback;
  return `"${family}"${fb ? `, "${fb}"${FACES[fb]?.fallback ? `, "${FACES[fb].fallback}"` : ''}` : ''}, sans-serif`;
};
export const canEmbed = (family) => !!FACES[family];

let fontkitP = null;
export const loadFontkit = () => fontkitP ||= window.fontkit ? Promise.resolve(window.fontkit) : new Promise((res, rej) => {
  const s = document.createElement('script'); s.src = FONTKIT; s.crossOrigin = 'anonymous';
  s.onload = () => res(window.fontkit); s.onerror = () => { fontkitP = null; rej(new Error('Could not load the font tools. Try again.')); }; document.head.append(s);
});
const bytesCache = new Map();
const fetchFont = (file) => {
  if (!bytesCache.has(file)) bytesCache.set(file, fetch(FONT_DIR + file).then((r) => { if (!r.ok) throw new Error(`Font ${file} is missing`); return r.arrayBuffer(); }));
  return bytesCache.get(file);
};
// EucrosiaUPC or Angsana New from this computer (Chrome/Edge, after the user allows "fonts"); null when not possible.
async function localFont(names) {
  try {
    if (!window.queryLocalFonts) return null;
    const all = await window.queryLocalFonts({ postscriptNames: names });
    return all.length ? Object.fromEntries(await Promise.all(all.map(async (f) => [f.postscriptName, await pickFont(await (await f.blob()).arrayBuffer(), f.postscriptName)]))) : null;
  } catch { return null; }
}
// Font collections (.ttc: Angsana New and AngsanaUPC in one file on Windows) -> one font file per font. pdf-lib
// embeds single fonts only, so each font's table directory is written with its tables copied after it.
const isCollection = (buf) => new TextDecoder('latin1').decode(new Uint8Array(buf, 0, 4)) === 'ttcf';
export function splitCollection(buf) {
  if (!isCollection(buf)) return [buf];
  const u8 = new Uint8Array(buf); const dv = new DataView(buf);
  return Array.from({ length: dv.getUint32(8) }, (_, k) => {
    const o = dv.getUint32(12 + 4 * k); const n = dv.getUint16(o + 4); let size = 12 + 16 * n; const tables = [];
    for (let i = 0; i < n; i++) { const p = o + 12 + 16 * i; const t = { p, off: dv.getUint32(p + 8), len: dv.getUint32(p + 12), at: size }; tables.push(t); size += (t.len + 3) & ~3; }
    const out = new Uint8Array(size); const ov = new DataView(out.buffer); out.set(u8.subarray(o, o + 12), 0);
    tables.forEach((t, i) => { out.set(u8.subarray(t.p, t.p + 8), 12 + 16 * i); ov.setUint32(12 + 16 * i + 8, t.at); ov.setUint32(12 + 16 * i + 12, t.len); out.set(u8.subarray(t.off, t.off + t.len), t.at); });
    return out.buffer;
  });
}
// The font of a file or collection with this PostScript name.
async function pickFont(buf, psName) {
  if (!isCollection(buf)) return buf;
  const fk = await loadFontkit();
  return splitCollection(buf).find((b) => { try { return fk.create(new Uint8Array(b)).postscriptName === psName; } catch { return false; } }) || buf;
}

const STYLE = (r) => (r.b && r.i ? 3 : r.b ? 1 : r.i ? 2 : 0);
const KEYS = ['r', 'b', 'i', 'bi'];

// One writer per exported PDF: embeds each face once (subset), draws text boxes. notes: things to tell the user.
export async function textWriter(pdf, PDFLib) {
  const fontkit = await loadFontkit(); pdf.registerFontkit(fontkit);
  const faces = new Map(); const notes = new Set(); const locals = new Map(); // family -> its faces from this computer
  // -> { font, alt, synthB, synthI, std } for family + style, or null if the family cannot be embedded
  async function face(family, st) {
    const key = `${family}|${st}`;
    if (faces.has(key)) return faces.get(key);
    const p = (async () => {
      let f = FACES[family]; if (!f) return null;
      if (f.std) return { std: true, font: await pdf.embedFont(PDFLib.StandardFonts[f.std[st]]), family };
      let bytes; let synthB = false; let synthI = false;
      if (f.local && !f.files) {
        if (!locals.has(family)) locals.set(family, await localFont(f.local));
        const got = locals.get(family);
        bytes = got?.[f.local[st]] || got?.[f.local[st & 1]] || got?.[f.local[0]];
        if (!bytes) { notes.add(`${family} is not available here, so TH Sarabun New was used`); return face(f.fallback, st); }
        synthB = !!(st & 1) && !got[f.local[st]] && !got[f.local[1]]; synthI = !!(st & 2) && !got[f.local[st]];
      } else {
        const src = f.files || f; // a user's font: its bytes; a site font: its file names
        const want = KEYS[st]; const file = src[want] || (st & 1 && src.b) || (st & 2 && src.i) || src.r;
        synthB = !!(st & 1) && !(src[want] || src.b) || (st === 3 && !src.bi && !src.b);
        synthI = !!(st & 2) && !(src[want] || src.i);
        bytes = f.files ? file : await fetchFont(file);
      }
      // two copies: `font` for ordinary glyphs, `alt` for glyphs whose text value is changed (ำ, see place())
      const font = await pdf.embedFont(bytes, { subset: true }); const alt = await pdf.embedFont(bytes, { subset: true });
      return { font, alt, synthB, synthI, family };
    })();
    faces.set(key, p); return p;
  }

  // Draws text box m on page pg. at(x, y) maps page units (top-left origin) to PDF space. Returns false when the box's
  // font cannot be embedded (the caller then draws it as a picture).
  async function draw(pg, m, at, { lineHeight, fontFamily }) {
    let family = fontFamily(m); if (!FACES[family]) return false;
    family = (await face(family, 0))?.family || family; // EucrosiaUPC not on this computer: lay out with its stand-in
    const size = m.size; const lh = size * lineHeight; const stack = fontStack(family);
    const ctx = document.createElement('canvas').getContext('2d');
    const css = (r) => `${r.i ? 'italic ' : ''}${r.b ? 'bold ' : ''}${size}px ${stack}`;
    ctx.font = css({}); const mt = ctx.measureText('Hg');
    const asc = mt.fontBoundingBoxAscent ?? size * 0.8; const desc = mt.fontBoundingBoxDescent ?? size * 0.25;
    const lines = [[]];
    m.runs.forEach((r) => r.t.split('\n').forEach((piece, i) => { if (i) lines.push([]); if (piece) lines.at(-1).push({ ...r, t: piece }); }));
    const color = hexRgb(m.color);
    for (const [i, line] of lines.entries()) {
      const base = m.y + i * lh + (lh - asc - desc) / 2 + asc; // baseline, as CSS places it (and textPng)
      let x = m.x;
      if ((m.align === 'center' || m.align === 'right') && m.wrap) { // centred or right-aligned in the box's width, as on screen
        const lw = line.reduce((s, r) => { ctx.font = css(r); return s + ctx.measureText(r.t).width; }, 0); x += (m.wrap - lw) / (m.align === 'center' ? 2 : 1);
      }
      for (const r of line) {
        ctx.font = css(r); const w = ctx.measureText(r.t).width;
        if (r.hl) { const [p, q] = [at(x, base - asc), at(x + w, base + desc)]; pg.drawRectangle({ x: p.x, y: q.y, width: q.x - p.x, height: p.y - q.y, color: rgbOf(PDFLib, hexRgb(r.hl)), opacity: r.ha ?? 0.45, blendMode: PDFLib.BlendMode?.Multiply }); }
        // split the run where its font lacks a letter (each Thai letter keeps its marks)
        let cx = x;
        for (const part of await coverage(r.t, family, STYLE(r))) {
          ctx.font = css(r); const pw = ctx.measureText(part.t).width;
          await place(pg, part.f, part.t, at(cx, base), size, r.c ? hexRgb(r.c) : color, PDFLib); // a run's own colour (Markup typing)
          cx += pw;
        }
        const th = Math.max(0.5, size / 14);
        const rc = r.c ? hexRgb(r.c) : color;
        if (r.u) { const p = at(x, base + size * 0.1); pg.drawRectangle({ x: p.x, y: p.y - th, width: w, height: th, color: rgbOf(PDFLib, rc) }); }
        if (r.s) { const p = at(x, base - size * 0.3); pg.drawRectangle({ x: p.x, y: p.y - th, width: w, height: th, color: rgbOf(PDFLib, rc) }); }
        x += w;
      }
    }
    return true;
  }

  // Text -> parts { t, f } where f is the face that has every letter of the part (the family, else its fallbacks).
  async function coverage(t, family, st) {
    const out = []; const seg = graphemes(t);
    for (const g of seg) {
      let fam = family; let f = await face(fam, st); let guard = 0;
      while (f && !has(f, g) && guard++ < 3) { fam = FACES[fam].fallback; f = fam ? await face(fam, st) : null; }
      f ||= await face(family, st);
      const last = out.at(-1); if (last && last.f === f) last.t += g; else out.push({ t: g, f });
    }
    return out;
  }
  return { draw, notes };
}

const graphemes = (t) => (Intl.Segmenter ? [...new Intl.Segmenter('th', { granularity: 'grapheme' }).segment(t)].map((s) => s.segment) : [...t]);
const has = (f, g) => {
  if (!g.trim()) return true;
  if (f.std) { const set = f._set ||= new Set(f.font.getCharacterSet()); return [...g].every((c) => set.has(c.codePointAt(0))); }
  const fk = f.font.embedder.font; return [...g].every((c) => fk.hasGlyphForCodePoint(c.codePointAt(0)));
};
const hexRgb = (h) => { const n = parseInt(String(h || '#000').slice(1), 16); return [(n >> 16) & 255, (n >> 8) & 255, n & 255]; };
const rgbOf = (PDFLib, [r, g, b]) => PDFLib.rgb(r / 255, g / 255, b / 255);

// A glyph with the text it copies out as (pdf-lib builds the PDF's text map from glyph.codePoints).
const as = (g, codePoints) => Object.create(g, { codePoints: { value: codePoints } });
// Glyph id -> 4-hex-digit code in the face's subset (what pdf-lib's encodeText does per glyph). The last text
// given for a glyph id is the one the PDF keeps.
function code(font, glyph) {
  const e = font.embedder; const id = e.subset.includeGlyph(glyph);
  e.glyphs[id - 1] = glyph; e.glyphIdMap.set(glyph.id, id); e.glyphCache.invalidate();
  return id.toString(16).padStart(4, '0');
}

// One part of a line, shaped, at baseline point p. ำ: fonts draw it as ํ + า (so a tone mark can sit between);
// the า would then copy out as an extra letter ("น้ำา"). So without a tone mark the font's own ำ glyph is used,
// and with one the tone glyph copies as the tone and the า as ำ (both from the `alt` copy, whose text values differ).
async function place(pg, f, t, p, size, color, PDFLib) {
  const { pushGraphicsState, popGraphicsState, beginText, endText, setFontAndSize, setTextMatrix, showText, PDFHexString,
    setFillingRgbColor, setStrokingRgbColor, setTextRenderingMode, setLineWidth } = PDFLib;
  const [r, g, b] = color.map((v) => v / 255);
  if (f.std) { pg.drawText(t, { x: p.x, y: p.y, size, font: f.font, color: PDFLib.rgb(r, g, b) }); return; }
  const fk = f.font.embedder.font; const run = fk.layout(t);
  const k = size / fk.unitsPerEm; const skew = f.synthI ? 0.2 : 0;
  const keyA = pg.node.newFontDictionary(f.font.name, f.font.ref); const keyB = pg.node.newFontDictionary(f.alt.name, f.alt.ref);
  const ops = [pushGraphicsState(), beginText(), setFillingRgbColor(r, g, b)];
  if (f.synthB) ops.push(setStrokingRgbColor(r, g, b), setLineWidth(size * 0.035), setTextRenderingMode(PDFLib.TextRenderingMode.FillAndOutline));
  let x = p.x; let cur = null;
  const put = (key, hex, dx, dy) => {
    if (cur !== key) { ops.push(setFontAndSize(key, size)); cur = key; }
    ops.push(setTextMatrix(1, 0, skew, 1, x + dx, p.y + dy), showText(PDFHexString.of(hex)));
  };
  const AM = 0x0E33; const AA = 0x0E32;
  // The shaper can leave a glyph without its letter (Sarabun's า after ํ, and then everywhere): every glyph gets
  // its text value set here, from the font's own character table when the shaper gave none.
  // (read from the cmap directly: fontkit caches a glyph with the text of its first lookup, so looking glyphs up
  // here would change what later shaping returns)
  const rev = fk._rev ||= (() => { const m = new Map(); for (const [a, b] of [[0x20, 0x7E], [0xA0, 0xFF], [0x0E01, 0x0E5B], [0x2010, 0x2044]]) for (let c = a; c <= b; c++) { const id = fk._cmapProcessor.lookup(c); if (id && !m.has(id)) m.set(id, c); } return m; })();
  const cps = (g) => (g.codePoints.length ? g.codePoints : rev.has(g.id) ? [rev.get(g.id)] : []);
  for (let i = 0; i < run.glyphs.length; i++) {
    const gl = run.glyphs[i]; const pos = run.positions[i]; const nx = run.glyphs[i + 1];
    const tail = nx && cps(nx).length === 1 && cps(nx)[0] === AA;
    if (cps(gl).includes(AM) && tail) {
      const tones = cps(gl).filter((c) => c !== AM);
      if (!tones.length && fk.hasGlyphForCodePoint(AM)) { // ํ alone: the font's ำ in place of ํ + า
        const am = fk.getGlyph(fk._cmapProcessor.lookup(AM)); put(keyA, code(f.font, as(am, [AM])), 0, 0); x += am.advanceWidth * k; i++; continue;
      }
      const afk = f.alt.embedder.font; // tone (+ ํ) copies as the tone; the า copies as ำ
      put(keyB, code(f.alt, as(afk.getGlyph(gl.id), tones)), pos.xOffset * k, pos.yOffset * k); x += pos.xAdvance * k;
      const np = run.positions[i + 1]; put(keyB, code(f.alt, as(afk.getGlyph(nx.id), [AM])), np.xOffset * k, np.yOffset * k); x += np.xAdvance * k; i++;
      continue;
    }
    put(keyA, code(f.font, as(gl, cps(gl))), pos.xOffset * k, pos.yOffset * k); x += pos.xAdvance * k;
  }
  if (f.synthB) ops.push(setTextRenderingMode(PDFLib.TextRenderingMode.Fill));
  ops.push(endText(), popGraphicsState());
  pg.pushOperators(...ops);
}

// The site's font for a font name found in a PDF: the same family when the site has it, else the nearest
// ("Angsana New" -> TH Sarabun New, "Cambria" -> Times New Roman, "Aptos" -> Arial).
const fkey = (n) => String(n).toLowerCase().replace(/[^a-z0-9]/g, '');
export function siteFont(name) {
  const hit = Object.keys(FACES).find((f) => fkey(f) === fkey(name)); if (hit) return hit;
  return /upc|psk|new$|thai|angsana|cordia|browallia|leelawadee|tahoma|kanit|prompt|mitr|sarabun|th /i.test(name) ? 'TH Sarabun New'
    : /times|georgia|cambria|garamond|serif|book|minion/i.test(name) ? 'Times New Roman' : 'Arial';
}

// ---------- fonts from the user's computer ----------
const FONT_DB = 'tusu-fonts';
const fontDb = () => new Promise((res, rej) => { const q = indexedDB.open(FONT_DB, 1); q.onupgradeneeded = () => q.result.createObjectStore('fonts'); q.onsuccess = () => res(q.result); q.onerror = () => rej(q.error); });
const dbDo = async (mode, fn) => { const db = await fontDb(); return new Promise((res, rej) => { const tx = db.transaction('fonts', mode); const r = fn(tx.objectStore('fonts')); tx.oncomplete = () => res(r?.result); tx.onerror = () => rej(tx.error); }); };
async function registerFont(family, key, bytes) {
  const f = (FACES[family] ||= { user: true, fallback: 'TH Sarabun New' }); f.files ||= {}; f.files[key] = bytes;
  const ff = new FontFace(family, bytes.slice(0), { weight: key === 'b' || key === 'bi' ? 'bold' : 'normal', style: key === 'i' || key === 'bi' ? 'italic' : 'normal' });
  await ff.load(); document.fonts.add(ff);
}
// A .ttf or .otf file -> { family, key } (key: r, b, i, bi from its own style name). The site's own fonts stay as
// they are ({ family, site: true }); EucrosiaUPC and new families are added.
export async function addFontFile(file) {
  const fk = await loadFontkit(); const all = splitCollection(await file.arrayBuffer()); let last = null;
  for (const bytes of all) {
    let font; try { font = fk.create(new Uint8Array(bytes)); } catch { throw new Error(`${file.name} is not a font file (.ttf, .otf or .ttc).`); }
    const family = (font.familyName || file.name.replace(/\.[^.]+$/, '')).trim();
    const ps = String(font.postscriptName || ''); const sub = String(font.subfamilyName || '').toLowerCase();
    // the style from the PostScript name (AngsanaNew-BoldItalic) first: style names can be in another language
    const bold = /-.*bold|bold$/i.test(ps) || (!/-/.test(ps) && /bold/.test(sub)); const ital = /-.*(italic|oblique)/i.test(ps) || (!/-/.test(ps) && /italic|oblique/.test(sub));
    const key = bold && ital ? 'bi' : bold ? 'b' : ital ? 'i' : 'r';
    if (FACES[family] && !FACES[family].user && !FACES[family].local) { last = last || { family, site: true }; continue; }
    await registerFont(family, key, bytes);
    try { await dbDo('readwrite', (s) => s.put(bytes, `${family}|${key}`)); } catch { /* storage off: the font lasts until the page closes */ }
    if (!last || last.site) last = { family, key, more: all.length > 1 ? all.length : 0 };
  }
  return last;
}
// Fonts added before, from this browser's storage. Resolves the families now usable.
export async function loadUserFonts() {
  try {
    const keys = await dbDo('readonly', (s) => s.getAllKeys()) || [];
    for (const k of keys) { const [family, key] = String(k).split('|'); if (FACES[family]?.files?.[key]) continue; const bytes = await dbDo('readonly', (s) => s.get(k)); if (bytes) await registerFont(family, key, bytes); }
  } catch { /* no storage */ }
  return Object.keys(FACES).filter((f) => FACES[f].files);
}
