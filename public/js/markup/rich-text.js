// Small rich-text helpers shared by the PDF text boxes (js/pdf-markup.js) and the Word editor (js/docx-edit.js).

// Typed arrows: when ">" is typed, "-->" before the caret becomes → and "==>" becomes ➜ ("—>" is "-->" after iPad
// smart punctuation). Uses insertText, so Ctrl+Z gives the typed characters back and the formatting around stays.
export const ARROWS = { '-->': '→', '==>': '➜', '—>': '→' };
export function arrowize() {
  const s = window.getSelection(); if (!s || !s.rangeCount || !s.isCollapsed) return false;
  const { startContainer: n, startOffset: o } = s.getRangeAt(0);
  if (n.nodeType !== 3) return false;
  const len = o >= 3 && ARROWS[n.data.slice(o - 3, o)] ? 3 : o >= 2 && ARROWS[n.data.slice(o - 2, o)] ? 2 : 0;
  if (!len) return false;
  const rep = ARROWS[n.data.slice(o - len, o)];
  const r = document.createRange(); r.setStart(n, o - len); r.setEnd(n, o);
  s.removeAllRanges(); s.addRange(r);
  let ok = false; try { ok = document.execCommand('insertText', false, rep); } catch { ok = false; }
  if (!ok) { const c = document.createRange(); c.setStart(n, o); c.collapse(true); s.removeAllRanges(); s.addRange(c); } // refused: caret back
  return ok;
}
export const exec = (cmd, v = null) => { try { return document.execCommand(cmd, false, v); } catch { return false; } };
