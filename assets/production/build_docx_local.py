# -*- coding: utf-8 -*-
"""习题册 HTML -> .docx（连续排版版）

每道题：题号 + 分类 ｜ ① 原题 ｜ ② 解题过程（右侧并排卷面原件）｜ ③ 答案
没有封面、索引、纲要、线索栏、总结栏、错因、订正区；题目连续排布，不再一题一页。

用法：python build_docx_local.py <in.html> <out.docx>
"""
import io
import json
import os
import re
import sys

from bs4 import BeautifulSoup, NavigableString
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from PIL import Image

NL = chr(10)
FONT = '微软雅黑'
MONO = 'Consolas'
LINE = 1.2
BODY = 9.5
CAL = '--cal' in sys.argv
CAP_FILE = os.environ.get('CJT_FIGCAPS') or os.path.join(os.getcwd(), 'figcaps.json')
CAPS = {}
if os.path.exists(CAP_FILE):
    CAPS = {int(k): float(v) for k, v in json.load(io.open(CAP_FILE, encoding='utf-8')).items()}

PAGE_W = 17.4
TEXT_W = 5.6
FIG_W = 11.2
FIG_MAX_H = 14.5
GREY = '555555'
LIGHT = '999999'


def norm(s):
    s = s.replace('\u3000', ' ')
    core = ' '.join(s.split())
    if not core:
        return ' ' if s else ''
    if s[:1].isspace():
        core = ' ' + core
    if s[-1:].isspace():
        core = core + ' '
    return core


def txt(node):
    return norm(node.get_text()) if node is not None else ''


def set_run(r, size=BODY, bold=False, mono=False, color=None):
    r.font.size = Pt(size)
    r.bold = bold
    name = MONO if mono else FONT
    r.font.name = name
    rPr = r._element.get_or_add_rPr()
    rf = rPr.find(qn('w:rFonts'))
    if rf is None:
        rf = OxmlElement('w:rFonts')
        rPr.insert(0, rf)
    for attr in ('w:ascii', 'w:hAnsi', 'w:eastAsia'):
        rf.set(qn(attr), name)
    if color:
        r.font.color.rgb = RGBColor.from_string(color)
    return r


def para(container, size=BODY, space_after=2, space_before=0, line=None, align=None, indent=None):
    p = container.add_paragraph()
    pf = p.paragraph_format
    pf.space_after = Pt(space_after)
    pf.space_before = Pt(space_before)
    pf.line_spacing = LINE if line is None else line
    if align is not None:
        p.alignment = align
    if indent is not None:
        pf.left_indent = Cm(indent)
    return p


PPR_SEQ = (
    'w:pStyle', 'w:keepNext', 'w:keepLines', 'w:pageBreakBefore', 'w:framePr',
    'w:widowControl', 'w:numPr', 'w:suppressLineNumbers', 'w:pBdr', 'w:shd',
    'w:tabs', 'w:suppressAutoHyphens', 'w:kinsoku', 'w:wordWrap',
    'w:overflowPunct', 'w:topLinePunct', 'w:autoSpaceDE', 'w:autoSpaceDN',
    'w:bidi', 'w:adjustRightInd', 'w:snapToGrid', 'w:spacing', 'w:ind',
    'w:contextualSpacing', 'w:mirrorIndents', 'w:suppressOverlap', 'w:jc',
    'w:textDirection', 'w:textAlignment', 'w:textboxTightWrap', 'w:outlineLvl',
    'w:divId', 'w:cnfStyle', 'w:rPr', 'w:sectPr', 'w:pPrChange',
)


def ppr_insert(pPr, el, tag):
    """按 OOXML 规定顺序把 el 插入 pPr。"""
    idx = PPR_SEQ.index(tag)
    pPr.insert_element_before(el, *PPR_SEQ[idx + 1:])


def para_border(p, edge='bottom', sz=6, color='000000'):
    pPr = p._p.get_or_add_pPr()
    b = pPr.find(qn('w:pBdr'))
    if b is None:
        b = OxmlElement('w:pBdr')
        ppr_insert(pPr, b, 'w:pBdr')
    el = OxmlElement('w:' + edge)
    el.set(qn('w:val'), 'single')
    el.set(qn('w:sz'), str(sz))
    el.set(qn('w:space'), '2')
    el.set(qn('w:color'), color)
    b.append(el)


def para_shade(p, fill='F4F5F7'):
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill)
    ppr_insert(pPr, shd, 'w:shd')


def para_keep_next(p):
    p.paragraph_format.keep_with_next = True


def keep_together(body, start, end):
    """让 [start, end) 这段（一道题）尽量不跨页。"""
    from docx.text.paragraph import Paragraph
    kids = list(body.iterchildren())
    for el in kids[start:max(start, end - 1)]:
        tag = el.tag.split('}')[-1]
        if tag == 'p':
            Paragraph(el, None).paragraph_format.keep_with_next = True
        elif tag == 'tbl':
            for tr in el.iter(qn('w:tr')):
                trPr = tr.find(qn('w:trPr'))
                if trPr is None:
                    trPr = OxmlElement('w:trPr')
                    tr.insert(0, trPr)
                if trPr.find(qn('w:cantSplit')) is None:
                    trPr.insert(0, OxmlElement('w:cantSplit'))
            for pp in el.iter(qn('w:p')):
                Paragraph(pp, None).paragraph_format.keep_with_next = True


def table_borders(table, sz=6, color='000000'):
    tblPr = table._tbl.tblPr
    borders = OxmlElement('w:tblBorders')
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        el = OxmlElement('w:' + edge)
        el.set(qn('w:val'), 'single' if sz else 'none')
        el.set(qn('w:sz'), str(sz or 0))
        el.set(qn('w:space'), '0')
        el.set(qn('w:color'), color)
        borders.append(el)
    tblPr.append(borders)


def cell_width(cell, cm):
    cell.width = Cm(cm)
    tcPr = cell._tc.get_or_add_tcPr()
    tcW = tcPr.find(qn('w:tcW'))
    if tcW is None:
        tcW = OxmlElement('w:tcW')
        tcPr.append(tcW)
    tcW.set(qn('w:w'), str(int(cm * 567)))
    tcW.set(qn('w:type'), 'dxa')


def cell_margins(table, top=40, bottom=40, left=60, right=60):
    tblPr = table._tbl.tblPr
    mar = OxmlElement('w:tblCellMar')
    for name, val in (('top', top), ('left', left), ('bottom', bottom), ('right', right)):
        el = OxmlElement('w:' + name)
        el.set(qn('w:w'), str(val))
        el.set(qn('w:type'), 'dxa')
        mar.append(el)
    tblPr.append(mar)


def clear_cell(cell):
    for p in list(cell.paragraphs):
        p._element.getparent().remove(p._element)


def iter_runs(node, bold=False, mono=False):
    for child in node.children:
        if isinstance(child, NavigableString):
            t = norm(str(child))
            if t:
                yield (t, bold, mono)
        elif child.name == 'br':
            yield (NL, bold, mono)
        else:
            b = bold or child.name in ('strong', 'b')
            m = mono or child.name in ('code', 'tt')
            for item in iter_runs(child, b, m):
                yield item


def add_inline(p, node, size=BODY, base_bold=False, color=None):
    merged = []
    for t, b, m in iter_runs(node, base_bold):
        if merged and merged[-1][1] == b and merged[-1][2] == m:
            merged[-1][0] += t
        else:
            merged.append([t, b, m])
    while merged and merged[0][0].strip() == '':
        merged.pop(0)
    while merged and merged[-1][0].strip() == '':
        merged.pop()
    if merged:
        merged[0][0] = merged[0][0].lstrip()
        merged[-1][0] = merged[-1][0].rstrip()
    for t, b, m in merged:
        if t == '':
            continue
        if NL in t:
            for i, part in enumerate(t.split(NL)):
                if i:
                    r = p.add_run()
                    set_run(r, size=size, bold=b, mono=m, color=color)
                    r.add_break()
                if part:
                    set_run(p.add_run(part), size=size, bold=b, mono=m, color=color)
        else:
            set_run(p.add_run(t), size=size, bold=b, mono=m, color=color)
    return p


def build_grid(container, widths, sz=6, color='BFBFBF', margins=(40, 40, 60, 60)):
    t = container.add_table(rows=1, cols=len(widths))
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    t.autofit = False
    table_borders(t, sz=sz, color=color)
    cell_margins(t, top=margins[0], bottom=margins[1], left=margins[2], right=margins[3])
    cells = t.rows[0].cells
    for c, w in zip(cells, widths):
        clear_cell(c)
        cell_width(c, w)
    return t, cells


def build_data_table(container, tb, width_cm=PAGE_W):
    rows = tb.find_all('tr')
    ncol = len(rows[0].find_all(['td', 'th']))
    t, _ = build_grid(container, [3.6] + [(width_cm - 3.6) / (ncol - 1)] * (ncol - 1),
                      sz=6, color='808080', margins=(25, 25, 60, 60))
    for r in rows:
        cells = r.find_all(['td', 'th'])
        row = t.add_row()
        for i, cell in enumerate(cells):
            c = row.cells[i]
            clear_cell(c)
            p = para(c, size=9, space_after=0)
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if i == 0 else WD_ALIGN_PARAGRAPH.CENTER
            add_inline(p, cell, size=9)
    t._tbl.remove(t.rows[0]._tr)
    return t


def add_picture(cell, path, cap, width_cm=FIG_W):
    w, h = Image.open(path).size
    p = para(cell, size=6, space_after=2, align=WD_ALIGN_PARAGRAPH.CENTER)
    p.add_run().add_picture(path, width=Cm(width_cm), height=Cm(width_cm * h / w))
    if cap:
        cp = para(cell, size=7.5, space_after=0, align=WD_ALIGN_PARAGRAPH.CENTER)
        set_run(cp.add_run(cap), size=7.5, color='777777')


def emit_section(container, node, size=BODY):
    st = node.find('span', class_='section-title')
    title = txt(st) if st is not None else ''
    if st is not None:
        p = para(container, size=size + 1, space_after=3, space_before=8)
        set_run(p.add_run(title), size=size + 1, bold=True, color='222222')
        para_border(p, 'bottom', sz=4, color=LIGHT)
        para_keep_next(p)
    answer = title.startswith('③')
    for child in node.children:
        if isinstance(child, NavigableString):
            continue
        if child.name == 'span':
            continue
        if child.name == 'table':
            build_data_table(container, child)
        elif child.name == 'hr':
            p = para(container, size=3, space_after=3)
            para_border(p, 'bottom', sz=4, color=LIGHT)
        elif child.name == 'div':
            p = para(container, size=size, space_after=1, indent=0.1 if answer else None)
            if answer:
                para_shade(p, 'F4F5F7')
            add_inline(p, child, size=size)


def add_footer(doc):
    p = doc.sections[0].footer.paragraphs[0]
    p.paragraph_format.tab_stops.add_tab_stop(Cm(PAGE_W), WD_TAB_ALIGNMENT.RIGHT)
    set_run(p.add_run(chr(9) + '— '), size=8.5, color='777777')
    fld = OxmlElement('w:fldSimple')
    fld.set(qn('w:instr'), 'PAGE')
    r = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    sz = OxmlElement('w:sz')
    sz.set(qn('w:val'), '17')
    rPr.append(sz)
    r.append(rPr)
    t = OxmlElement('w:t')
    t.text = '1'
    r.append(t)
    fld.append(r)
    p._p.append(fld)
    set_run(p.add_run(' —'), size=8.5, color='777777')


def add_header(doc):
    hdr = doc.sections[0].header
    p = hdr.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.tab_stops.add_tab_stop(Cm(PAGE_W), WD_TAB_ALIGNMENT.RIGHT)
    set_run(p.add_run('初中科学典型习题 · 详细解析'), size=8.5, color='888888')
    set_run(p.add_run(chr(9) + '生物分类 · 测量 · 实验探究'), size=8.5, color='AAAAAA')
    para_border(p, 'bottom', sz=4, color='CCCCCC')


def build_title(doc, n):
    p = para(doc, size=17, space_after=2, line=1.1)
    set_run(p.add_run('初中科学典型习题 · 详细解析'), size=17, bold=True)
    p2 = para(doc, size=9.5, space_after=10)
    set_run(p2.add_run('生物分类 · 测量 · 实验探究　全 %d 题' % n), size=9.5, color=GREY)
    para_border(p2, 'bottom', sz=12, color='333333')


def est_cap_cm(blk, page_h=25.4):
    """按「① 原题」的文字量估算卷面图可用高度（cm）。"""
    total = 0.62                      # ① 小标题
    secs = blk.select("div.section")
    for sec in secs:
        st = sec.find("span", class_="section-title")
        title = st.get_text().strip() if st is not None else ""
        if not title.startswith("①"):
            continue
        for child in sec.children:
            if isinstance(child, NavigableString):
                continue
            if child.name == "hr":
                total += 0.25
            elif child.name == "table":
                total += 2.4
            elif child.name == "div":
                t = child.get_text().strip()
                if t:
                    total += max(1, int(len(t) / 54) + 1) * 0.44
    return max(6.0, min(FIG_MAX_H, page_h - 0.95 - total - 1.35))

def build_question(doc, page, base, idx, last=False):
    blk = page.find('div', class_='question-block')
    head = blk.find('div', class_='q-header')
    raw = txt(head.find('h2'))
    raw = re.sub(r'^第\s*(典例|变式)\s*([0-9][^｜]*?)\s*题\s*｜\s*', r'\1 \2 ｜ ', raw)
    raw = re.sub(r'^第\s*([0-9]+)\s*题\s*｜\s*', r'第 \1 题　', raw)
    parts = [x.strip() for x in raw.split('｜')]
    num = parts[0]
    rest = ' · '.join([x for x in parts[1:] if x])

    # 整道题装进一个「行不可分割」的单元格，保证不会跨页断开
    if idx > 1:
        para(doc, size=3, space_after=0)
    box = doc.add_table(rows=1, cols=1)
    box.autofit = False
    table_borders(box, sz=0)
    cell_margins(box, top=0, bottom=0, left=0, right=0)
    trPr = box.rows[0]._tr.get_or_add_trPr()
    trPr.insert(0, OxmlElement('w:cantSplit'))
    c = box.cell(0, 0)
    clear_cell(c)
    cell_width(c, PAGE_W)

    p = para(c, size=13, space_after=4, space_before=6, line=1.15)
    if CAL and idx > 1:
        p.paragraph_format.page_break_before = True
    if '　' in num:                      # 「第 7 题　生物 · 分类检索」
        head_num, tail = num.split('　', 1)
        set_run(p.add_run(head_num), size=13, bold=True)
        rest = (tail + (' · ' if rest else '') + rest).strip()
    else:
        set_run(p.add_run(num), size=13, bold=True)
    if rest:
        set_run(p.add_run('　'), size=11)
        set_run(p.add_run(rest), size=9.5, color=GREY)
    para_border(p, 'bottom', sz=8, color='333333')

    img_cap = 6.0 if CAL else CAPS.get(idx, est_cap_cm(blk))
    main = blk.find('div', class_='main-column')
    for child in main.children:
        if isinstance(child, NavigableString):
            continue
        cls = child.get('class') or []
        if 'zone-label' in cls:
            continue
        if 'figblock' in cls:
            img = child.find('img')
            cap = child.find('div', class_='figcap')
            if img is not None:
                path = os.path.join(base, img.get('src'))
                if os.path.exists(path):
                    ip = para(c, size=6, space_before=4, space_after=1,
                              align=WD_ALIGN_PARAGRAPH.CENTER)
                    w, h = Image.open(path).size
                    wcm = min(FIG_W, FIG_MAX_H * w / float(h))
                    ip.add_run().add_picture(path, width=Cm(wcm), height=Cm(wcm * h / w))
                    if cap is not None and txt(cap):
                        cp = para(c, size=7.5, space_after=2, align=WD_ALIGN_PARAGRAPH.CENTER)
                        set_run(cp.add_run(txt(cap)), size=7.5, color='777777')
        elif 'sec-row' in cls:
            st = child.find('div', class_='sec-text')
            fig = child.find('div', class_='sec-fig')
            fig_w = PAGE_W - TEXT_W - 0.6
            nt, (tl, tr) = build_grid(c, [TEXT_W, fig_w], sz=0, margins=(20, 20, 0, 50))
            if st is not None:
                emit_section(tl, st, size=BODY)
            if fig is not None:
                img = fig.find('img')
                cap = fig.find('div', class_='figcap')
                if img is not None:
                    path = os.path.join(base, img.get('src'))
                    if os.path.exists(path):
                        w, h = Image.open(path).size
                        wcm = min(fig_w - 0.15, img_cap * w / float(h))
                        ip = para(tr, size=6, space_after=2, align=WD_ALIGN_PARAGRAPH.CENTER)
                        ip.add_run().add_picture(path, width=Cm(wcm), height=Cm(wcm * h / w))
                        if cap is not None and txt(cap):
                            cp = para(tr, size=7.5, space_after=0, align=WD_ALIGN_PARAGRAPH.CENTER)
                            set_run(cp.add_run(txt(cap)), size=7.5, color='777777')
        elif 'section' in cls:
            emit_section(c, child, size=BODY)

    if page.find('div', class_='figpage') is not None:
        p = para(c, size=8.5, space_before=7, space_after=2)
        set_run(p.add_run('▸ 卷面原图见下页：整页原图已完整收录（左右两栏并排，左栏读完接右栏）'),
                size=8.5, color='777777')


def build_figpage(doc, node, base):
    """整页长截图：单独一页，左右两栏并排。"""
    pb = para(doc, size=2, space_after=0)
    pb.paragraph_format.page_break_before = True
    cap = node.find('div', class_='figcap')
    if cap is not None and txt(cap):
        cp = para(doc, size=8.5, space_before=4, space_after=6, align=WD_ALIGN_PARAGRAPH.CENTER)
        set_run(cp.add_run(txt(cap)), size=8.5, color='666666')
    imgs = node.find_all('img')
    paths = [os.path.join(base, im.get('src')) for im in imgs]
    paths = [p for p in paths if os.path.exists(p)]
    if not paths:
        return
    widths = [PAGE_W / len(paths) - 0.3] * len(paths)
    t, cells = build_grid(doc, widths, sz=0, margins=(0, 0, 0, 0))
    for cell, path in zip(cells, paths):
        w, h = Image.open(path).size
        wcm = min(widths[0], 23.0 * w / float(h))
        ip = para(cell, size=6, space_after=0, align=WD_ALIGN_PARAGRAPH.CENTER)
        ip.add_run().add_picture(path, width=Cm(wcm), height=Cm(wcm * h / w))


def main():
    if len(sys.argv) < 3:
        print('usage: python build_docx_local.py <in.html> <out.docx>')
        sys.exit(1)
    src = os.path.abspath(sys.argv[1])
    out = os.path.abspath(sys.argv[2])
    base = os.path.dirname(src)
    soup = BeautifulSoup(io.open(src, encoding='utf-8').read(), 'lxml')

    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.top_margin, sec.bottom_margin = Cm(1.6), Cm(1.5)
    sec.left_margin, sec.right_margin = Cm(1.8), Cm(1.8)
    st = doc.styles['Normal']
    st.font.name = FONT
    st.font.size = Pt(BODY)
    st.element.rPr.rFonts.set(qn('w:eastAsia'), FONT)
    st.paragraph_format.space_after = Pt(2)
    st.paragraph_format.line_spacing = LINE
    sec.header_distance = Cm(0.9)
    sec.top_margin = Cm(1.8)
    add_header(doc)
    add_footer(doc)

    pages = [p for p in soup.select('div.page') if p.find('div', class_='question-block')]
    for i, page in enumerate(pages, 1):
        build_question(doc, page, base, i, last=(i == len(pages)))
        fp = page.find('div', class_='figpage')
        if fp is not None:
            build_figpage(doc, fp, base)
    doc.save(out)
    print('DOCX:', out, os.path.getsize(out), '| questions:', len(pages))


if __name__ == '__main__':
    main()








