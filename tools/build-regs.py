"""Regulations sidebar data (v2.17): the six regulation summaries (.md 01-06 in PPB TEAM\\Regulations), the
กองกิจ practice guide and the Summary tables (tools/regs/*.md) -> public/data/regs.json, read by public/js/regs.js.

    python -I tools/build-regs.py ["C:/Users/.../PPB TEAM/Regulations"]

Markdown here is the plain subset those files use: headings, <a id> anchors, tables, lists, quotes, bold, links.
Sample-document cards come from tools/regs/samples.json; a card shows public/img/regs/p<page>.png when it exists.
"""
import html
import json
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SRC = Path(sys.argv[1] if len(sys.argv) > 1 else r'C:\Users\knowd\OneDrive\TUSU\PPB TEAM\Regulations')
REGS = [  # file, short name, font of the printed original
    ('03-student-activity-regulation-2563.md', 'ข้อบังคับกิจกรรมนักศึกษา 2563', 'TH Sarabun New'),
    ('05-student-activity-budget-regulation-2569.md', 'ระเบียบงบประมาณกิจกรรมนักศึกษา 2569', 'TH Sarabun New'),
    ('04-central-revenue-payment-rates-2567.md', 'ประกาศอัตราการจ่ายเงินรายได้ส่วนกลาง 2567', 'TH Sarabun New'),
    ('01-sports-club-budget-committee.md', 'ประกาศงบชุมนุมกีฬาฯ 2568', 'TH Sarabun New'),
    ('06-independent-group-regulation-2566.md', 'ระเบียบกลุ่มอิสระ 2566', 'TH Sarabun New'),
    ('02-document-review-timeline.md', 'หลักเกณฑ์ระยะเวลาตรวจเอกสาร', 'TH Sarabun PSK'),
]
NO = lambda f: f[:2]  # '03' ...


def inline(s, doc):
    """Escape, then bold / italic / code / links. Links to another .md become in-panel links."""
    s = html.escape(s, quote=False)
    s = re.sub(r'`([^`]+)`', r'<code>\1</code>', s)
    s = re.sub(r'\*\*(.+?)\*\*', r'<b>\1</b>', s)
    s = re.sub(r'(?<![*\w])\*([^*\s][^*]*?)\*(?!\*)', r'<i>\1</i>', s)

    def link(m):
        text, url = m.group(1), html.unescape(m.group(2))
        if re.match(r'https?://', url):
            return f'<a href="{html.escape(url)}" target="_blank" rel="noopener">{text}</a>'
        f = re.match(r'(\d\d)-[\w-]+\.md(?:#([\w-]+))?$', url)
        if f:
            return f'<a href="#" data-reg="{f.group(1)}" data-at="{f.group(2) or ""}">{text}</a>'
        if url.startswith('#'):
            return f'<a href="#" data-reg="{doc}" data-at="{html.escape(url[1:])}">{text}</a>'
        return text
    return re.sub(r'\[([^\]]+)\]\(([^)]+)\)', link, s)


def convert(md, doc, prefix):
    """Markdown subset -> (html, toc). Heading ids: the <a id> before them, else prefix + n."""
    out, toc = [], []
    lines = md.replace('\r', '').split('\n')
    anchor, n, i = None, 0, 0

    def table(rows):
        cells = lambda r: [c.strip() for c in r.strip().strip('|').split('|')]
        head, body = cells(rows[0]), [cells(r) for r in rows[2:]]
        th = ''.join(f'<th>{inline(c, doc)}</th>' for c in head)
        tb = ''.join('<tr>' + ''.join(f'<td>{inline(c, doc)}</td>' for c in r) + '</tr>' for r in body)
        return f'<div class="rg-table"><table><thead><tr>{th}</tr></thead><tbody>{tb}</tbody></table></div>'

    def lst(block):
        html_, stack = '', []  # stack of (indent, tag)
        for ln in block:
            m = re.match(r'^(\s*)([-*]|\d+\.)\s+(.*)$', ln)
            if not m:  # continuation line
                html_ += ' ' + inline(ln.strip(), doc)
                continue
            ind, tag = len(m.group(1)), 'ol' if m.group(2)[0].isdigit() else 'ul'
            while stack and stack[-1][0] > ind:
                html_ += f'</li></{stack.pop()[1]}>'
            if not stack or stack[-1][0] < ind:
                html_ += f'<{tag}><li>'
                stack.append((ind, tag))
            else:
                html_ += '</li><li>'
            html_ += inline(m.group(3), doc)
        while stack:
            html_ += f'</li></{stack.pop()[1]}>'
        return html_

    while i < len(lines):
        ln = lines[i]
        m = re.match(r'^<a id="([^"]+)"></a>\s*$', ln)
        if m:
            anchor = m.group(1); i += 1; continue
        m = re.match(r'^(#{1,4})\s+(.*)$', ln)
        if m:
            level, text = len(m.group(1)), m.group(2).strip()
            n += 1
            hid = f'{prefix}-{anchor}' if anchor else f'{prefix}-h{n}'
            anchor = None
            out.append(f'<h{level} id="{hid}">{inline(text, doc)}</h{level}>')
            if level in (2, 3):
                toc.append({'id': hid, 'title': re.sub(r'<[^>]+>', '', inline(text, doc)), 'level': level})
            i += 1; continue
        if ln.startswith('|'):
            rows = []
            while i < len(lines) and lines[i].startswith('|'):
                rows.append(lines[i]); i += 1
            out.append(table(rows)); continue
        if re.match(r'^\s*([-*]|\d+\.)\s+', ln):
            block = []
            while i < len(lines) and lines[i].strip() and (re.match(r'^\s*([-*]|\d+\.)\s+', lines[i]) or lines[i].startswith('  ')):
                block.append(lines[i]); i += 1
            out.append(lst(block)); continue
        if ln.startswith('>'):
            q = []
            while i < len(lines) and lines[i].startswith('>'):
                q.append(lines[i].lstrip('>').strip()); i += 1
            out.append('<blockquote>' + '<br>'.join(inline(x, doc) for x in q if x) + '</blockquote>'); continue
        if re.match(r'^-{3,}\s*$', ln):
            out.append('<hr>'); i += 1; continue
        if ln.strip() == '[[samples]]':
            out.append(samples()); i += 1; continue
        if not ln.strip():
            i += 1; continue
        para = []
        while i < len(lines) and lines[i].strip() and not re.match(r'^(#|\||>|<a id|\s*([-*]|\d+\.)\s)', lines[i]):
            para.append(lines[i].strip()); i += 1
        out.append('<p>' + ' '.join(inline(x, doc) for x in para) + '</p>')
    return '\n'.join(out), toc


def samples():
    cards = []
    for c in json.loads((ROOT / 'tools/regs/samples.json').read_text(encoding='utf-8')):
        img = f'img/regs/p{c["page"]}.png'
        pic = (f'<a href="{img}" target="_blank" rel="noopener" title="เปิดภาพขนาดเต็ม"><img src="{img}" alt="{html.escape(c["title"])}" loading="lazy"></a>'
               if (ROOT / 'public' / img).exists()
               else '')  # no picture (blank or not exported): the card is text only
        note = f'<span>{html.escape(c["note"])}</span>' if c['note'] else ''
        cards.append(f'<figure class="rg-card">{pic}<figcaption><b>{html.escape(c["title"])}</b>{note}'
                     f'<small>สไลด์ "ข้อมูลภายใน" หน้า {c["page"]}</small></figcaption></figure>')
    return '<div class="rg-cards">' + ''.join(cards) + '</div>'


def strip_nav(md):
    """The .md's own back link and table of contents: the panel has its own."""
    md = re.sub(r'^\[← กลับหน้ารวม\]\([^)]*\)\s*$', '', md, flags=re.M)
    return re.sub(r'^## สารบัญ\s*\n(?:(?!^## |^<a id=).*\n)*', '', md, flags=re.M)


def bold_clauses(h):
    """Regulation mode: "ข้อ 34", "(ข้อ 5–6)" in bold, outside tags."""
    return re.sub(r'(>[^<]*)', lambda m: re.sub(r'(ข้อ\s?\d+(?:\s?\(\d+\))?(?:[–-]\d+)?)', r'<b class="rg-no">\1</b>', m.group(1)), h)


def main():
    regs = []
    for f, short, font in REGS:
        md = strip_nav((SRC / f).read_text(encoding='utf-8'))
        title = re.search(r'^# (.+)$', md, re.M).group(1).strip()
        md = re.sub(r'^# .+$', '', md, count=1, flags=re.M)
        body, toc = convert(md, NO(f), f'r{NO(f)}')
        regs.append({'id': NO(f), 'title': title, 'short': short, 'font': font, 'kind': 'reg', 'html': bold_clauses(body), 'toc': toc})
    md = (ROOT / 'tools/regs/kongkit.md').read_text(encoding='utf-8')
    title = re.search(r'^# (.+)$', md, re.M).group(1).strip()
    body, toc = convert(re.sub(r'^# .+$', '', md, count=1, flags=re.M), 'kk', 'rkk')
    regs.append({'id': 'kk', 'title': title, 'short': title, 'font': 'TH Sarabun New', 'kind': 'guide', 'html': bold_clauses(body), 'toc': toc})

    md = (ROOT / 'tools/regs/summary.md').read_text(encoding='utf-8')
    title = re.search(r'^# (.+)$', md, re.M).group(1).strip()
    body, toc = convert(re.sub(r'^# .+$', '', md, count=1, flags=re.M), 'sum', 's')
    out = {'summary': {'title': title, 'html': body, 'toc': [t for t in toc if t['level'] == 2]}, 'regs': regs}
    dest = ROOT / 'public/data/regs.json'
    dest.parent.mkdir(parents=True, exist_ok=True)
    dest.write_text(json.dumps(out, ensure_ascii=False, separators=(',', ':')), encoding='utf-8')
    print(f'{dest}: {dest.stat().st_size // 1024} KB, {len(regs)} documents, {len(out["summary"]["toc"])} summary topics')


if __name__ == '__main__':
    main()
