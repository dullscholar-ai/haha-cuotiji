# -*- coding: utf-8 -*-
"""
错题集排版引擎（哈哈错题集）
--------------------------------
把「题目数据 + 卷面截图」渲染成一份一题一页的错题集 docx。

用法：
    python build_docx.py --content content.json --figs figs --out 错题集.docx
    python build_docx.py ... --fit fit.json --side 01,02,04,12,13 --pdf

content.json 结构见 examples/sample-content.json，字段说明见 references/layout-spec.md。
只依赖 python-docx + Pillow。
"""
import argparse
import io
import json
import os
import sys
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ---------------------------------------------------------------- palette
INK        = RGBColor(0x22, 0x30, 0x3A)
MUTED      = RGBColor(0x6B, 0x7C, 0x88)
ACCENT     = RGBColor(0x1F, 0x5C, 0x74)   # 深青
ACCENT_HEX = '1F5C74'
ACCENT_L   = 'E9F2F6'
WARM       = RGBColor(0xA5, 0x55, 0x25)   # 赭石
WARM_HEX   = 'A55525'
WARM_L     = 'FBF2E9'
LINE_HEX   = 'C9D6DE'
SOFT_HEX   = 'E3EAEE'

BODY_EA   = '微软雅黑'          # 最终版式：正文全套微软雅黑
BODY_ASCII = '微软雅黑'
HEAD_EA   = '微软雅黑'
HEAD_ASCII = 'Segoe UI'

CONTENT_W = 17.4          # cm
CUE_W     = 4.6
NOTES_W   = CONTENT_W - CUE_W
FIG_MAX_H = 10.8          # 图高上限（每题再由 figcaps.json 收窄）
FIG_MAX_W = 11.2

# ---- 运行期配置（由 main() 从命令行注入） -------------------------------
FIGDIR = '.'                    # 截图目录
SIDE = set()                    # 截图与解析并列的题号
SIDE_L = 5.6                    # 并列版式：解析栏宽
SIDE_R = 11.2                   # 并列版式：截图栏宽
FIT = {'fig': {}, 'fs': {}, 'sidew': {}}   # 自动适配参数（fit.py 产出）
OPT = {'multiple': 1.0, 'first_page_number': 1, 'pdf': False}


def load_fit(path):
    """读入 fit.py 产出的适配参数；不存在时用默认值。"""
    if path and os.path.exists(path):
        with io.open(path, encoding='utf-8') as fh:
            data = json.load(fh)
        for k in ('fig', 'fs', 'sidew'):
            FIT.setdefault(k, {}).update(data.get(k, {}))
    return FIT

# ---------------------------------------------------------------- low level xml helpers
def _el(tag, **attrs):
    e = OxmlElement(tag)
    for k, v in attrs.items():
        e.set(qn('w:' + k), str(v))
    return e


def set_run_font(run, size, color=INK, bold=False, italic=False,
                 ea=BODY_EA, ascii_f=BODY_ASCII, spacing=None):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    if color is not None:
        run.font.color.rgb = color
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:ascii'), ascii_f)
    rFonts.set(qn('w:hAnsi'), ascii_f)
    rFonts.set(qn('w:eastAsia'), ea)
    rFonts.set(qn('w:cs'), ascii_f)
    if spacing is not None:
        rPr.append(_el('w:spacing', val=spacing))


def para_fmt(p, before=0, after=2, line=1.34, align=None, left=0.0,
             first=0.0, keep_next=False, keep_lines=False):
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    if line:
        pf.line_spacing = line
    if align is not None:
        pf.alignment = align
    if left:
        pf.left_indent = Cm(left)
    if first:
        pf.first_line_indent = Cm(first)
    pf.keep_with_next = keep_next
    pf.keep_together = keep_lines
    pf.widow_control = True


def border_cell(cell, edge, sz=8, color=LINE_HEX, val='single'):
    tcPr = cell._tc.get_or_add_tcPr()
    borders = tcPr.find(qn('w:tcBorders'))
    if borders is None:
        borders = OxmlElement('w:tcBorders')
        tcPr.append(borders)
    e = borders.find(qn('w:' + edge))
    if e is None:
        e = OxmlElement('w:' + edge)
        borders.append(e)
    if val == 'nil':
        e.set(qn('w:val'), 'nil')
    else:
        e.set(qn('w:val'), val)
        e.set(qn('w:sz'), str(sz))
        e.set(qn('w:space'), '0')
        e.set(qn('w:color'), color)


def shade(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = tcPr.find(qn('w:shd'))
    if shd is None:
        shd = OxmlElement('w:shd')
        tcPr.append(shd)
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill)


def border_para(p, edge='bottom', sz=6, color=LINE_HEX):
    pPr = p._p.get_or_add_pPr()
    pbdr = pPr.find(qn('w:pBdr'))
    if pbdr is None:
        pbdr = OxmlElement('w:pBdr')
        pPr.append(pbdr)
    e = OxmlElement('w:' + edge)
    e.set(qn('w:val'), 'single')
    e.set(qn('w:sz'), str(sz))
    e.set(qn('w:space'), '4')
    e.set(qn('w:color'), color)
    pbdr.append(e)


def table_borders(table, spec, color=LINE_HEX, sz=4):
    tblPr = table._tbl.tblPr
    old = tblPr.find(qn('w:tblBorders'))
    if old is not None:
        tblPr.remove(old)
    borders = OxmlElement('w:tblBorders')
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        e = OxmlElement('w:' + edge)
        if spec.get(edge) is None:
            e.set(qn('w:val'), 'nil')
        else:
            s, c = spec[edge]
            e.set(qn('w:val'), 'single')
            e.set(qn('w:sz'), str(s))
            e.set(qn('w:space'), '0')
            e.set(qn('w:color'), c)
        borders.append(e)
    tblPr.append(borders)


def table_fixed(table, total_cm, widths_cm):
    tblPr = table._tbl.tblPr
    table.autofit = False
    old = tblPr.find(qn('w:tblW'))
    if old is not None:
        tblPr.remove(old)
    tblPr.append(_el('w:tblW', w=int(total_cm * 567), type='dxa'))
    # autofit=False 已由 python-docx 写入 w:tblLayout（type=fixed），此处不再重复添加
    for i, w in enumerate(widths_cm):
        table.columns[i].width = Cm(w)
    for row in table.rows:
        for i, w in enumerate(widths_cm):
            if i < len(row.cells):
                row.cells[i].width = Cm(w)


def table_margins(table, top=0.10, left=0.16, bottom=0.10, right=0.16):
    tblPr = table._tbl.tblPr
    old = tblPr.find(qn('w:tblCellMar'))
    if old is not None:
        tblPr.remove(old)
    mar = OxmlElement('w:tblCellMar')
    for edge, v in (('top', top), ('left', left), ('bottom', bottom), ('right', right)):
        e = OxmlElement('w:' + edge)
        e.set(qn('w:w'), str(int(v * 567)))
        e.set(qn('w:type'), 'dxa')
        mar.append(e)
    tblPr.append(mar)


def no_split(row):
    trPr = row._tr.get_or_add_trPr()
    trPr.append(OxmlElement('w:cantSplit'))


def row_height(row, cm, rule='atLeast'):
    trPr = row._tr.get_or_add_trPr()
    h = OxmlElement('w:trHeight')
    h.set(qn('w:val'), str(int(cm * 567)))
    h.set(qn('w:hRule'), rule)
    trPr.append(h)


# ---------------------------------------------------------------- rich text
import re
TOKEN = re.compile(r'(\{#[^}]*\}|\{\^[^}]*\}|\*\*[^*]+\*\*)')


def add_rich(p, text, size=10, color=INK, bold=False, italic=False,
             ea=BODY_EA, ascii_f=BODY_ASCII):
    for tok in TOKEN.split(text):
        if not tok:
            continue
        if tok.startswith('{#') and tok.endswith('}'):
            r = p.add_run(tok[2:-1]); set_run_font(r, size, color, bold, italic, ea, ascii_f)
            r.font.subscript = True
        elif tok.startswith('{^') and tok.endswith('}'):
            r = p.add_run(tok[2:-1]); set_run_font(r, size, color, bold, italic, ea, ascii_f)
            r.font.superscript = True
        elif tok.startswith('**') and tok.endswith('**'):
            r = p.add_run(tok[2:-2]); set_run_font(r, size, color, True, italic, ea, ascii_f)
        else:
            r = p.add_run(tok); set_run_font(r, size, color, bold, italic, ea, ascii_f)
    return p


def para(parent, text='', size=10, color=INK, bold=False, italic=False,
         ea=BODY_EA, ascii_f=BODY_ASCII, before=0, after=2, line=1.34,
         align=None, left=0.0, first=0.0, keep_next=False, keep_lines=False):
    p = parent.add_paragraph()
    para_fmt(p, before, after, line, align, left, first, keep_next, keep_lines)
    if text:
        add_rich(p, text, size, color, bold, italic, ea, ascii_f)
    return p


def field(paragraph, code, size=8.5, color=MUTED):
    r = paragraph.add_run()
    set_run_font(r, size, color)
    fld = OxmlElement('w:fldChar'); fld.set(qn('w:fldCharType'), 'begin')
    it = OxmlElement('w:instrText'); it.set(qn('xml:space'), 'preserve'); it.text = ' %s ' % code
    sep = OxmlElement('w:fldChar'); sep.set(qn('w:fldCharType'), 'separate')
    t = OxmlElement('w:t'); t.text = '1'
    end = OxmlElement('w:fldChar'); end.set(qn('w:fldCharType'), 'end')
    for e in (fld, it, sep, t, end):
        r._r.append(e)


# ---------------------------------------------------------------- schema order fix
_ORDER = {
    'pPr': ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl',
            'numPr', 'suppressLineNumbers', 'pBdr', 'shd', 'tabs', 'suppressAutoHyphens',
            'kinsoku', 'wordWrap', 'overflowPunct', 'topLinePunct', 'autoSpaceDE',
            'autoSpaceDN', 'bidi', 'adjustRightInd', 'snapToGrid', 'spacing', 'ind',
            'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc', 'textDirection',
            'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr',
            'sectPr', 'pPrChange'],
    'rPr': ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike',
            'dstrike', 'outline', 'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid',
            'vanish', 'webHidden', 'color', 'spacing', 'w', 'kern', 'position', 'sz', 'szCs',
            'highlight', 'u', 'effect', 'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs',
            'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath'],
    'tblPr': ['tblStyle', 'tblpPr', 'tblOverlap', 'bidiVisual', 'tblStyleRowBandSize',
              'tblStyleColBandSize', 'tblW', 'jc', 'tblCellSpacing', 'tblInd', 'tblBorders',
              'shd', 'tblLayout', 'tblCellMar', 'tblLook', 'tblCaption', 'tblDescription'],
    'trPr': ['cnfStyle', 'divId', 'gridBefore', 'gridAfter', 'wBefore', 'wAfter', 'cantSplit',
             'trHeight', 'tblHeader', 'tblCellSpacing', 'jc', 'hidden'],
    'tcPr': ['cnfStyle', 'tcW', 'gridSpan', 'hMerge', 'vMerge', 'tcBorders', 'shd', 'noWrap',
             'tcMar', 'textDirection', 'tcFitText', 'vAlign', 'hideMark'],
}
_NSW = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


def canonicalize(root):
    changed = 0
    for tag, order in _ORDER.items():
        idx = {name: i for i, name in enumerate(order)}
        for el in root.iter(_NSW + tag):
            kids = list(el)
            def key(k):
                name = k.tag.replace(_NSW, '') if isinstance(k.tag, str) else ''
                return idx.get(name, len(order))
            ordered = sorted(kids, key=key)
            if [id(k) for k in ordered] != [id(k) for k in kids]:
                for k in kids:
                    el.remove(k)
                for k in ordered:
                    el.append(k)
                changed += 1
    return changed

# ---------------------------------------------------------------- content
TYPE_ = None  # placeholder


C = []          # 由 load_content() 填充





# ---------------------------------------------------------------- building blocks
def label_para(parent, text, color, fill=None, left_border=True, before=0, after=2.5):
    p = parent.add_paragraph()
    para_fmt(p, before, after, 1.0, None, 0.0, 0.0, True)
    r = p.add_run(text)
    set_run_font(r, 8.5, color, True, False, HEAD_EA, HEAD_ASCII, spacing=24)
    if left_border:
        pPr = p._p.get_or_add_pPr()
        pbdr = OxmlElement('w:pBdr')
        e = OxmlElement('w:left')
        e.set(qn('w:val'), 'single'); e.set(qn('w:sz'), '18')
        e.set(qn('w:space'), '4'); e.set(qn('w:color'), color and ('%02X%02X%02X' % (color[0], color[1], color[2])))
        pbdr.append(e)
        pPr.append(pbdr)
        p.paragraph_format.left_indent = Cm(0.16)
    return p


def add_figure(cell, fname, caption, avail=NOTES_W - 0.75, maxh=None, cap_size=8):
    from PIL import Image
    path = os.path.join(FIGDIR, fname)
    w_px, h_px = Image.open(path).size
    ratio = h_px / w_px
    maxh = FIG_MAX_H if maxh is None else maxh
    width = min(FIG_MAX_W, avail, maxh / ratio)
    height = width * ratio

    holder = cell.add_table(rows=1, cols=1)
    holder.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_fixed(holder, width + 0.30, [width + 0.30])
    table_margins(holder, 0.10, 0.14, 0.10, 0.14)
    table_borders(holder, {'top': (4, LINE_HEX), 'left': (4, LINE_HEX),
                           'bottom': (4, LINE_HEX), 'right': (4, LINE_HEX)})
    c = holder.cell(0, 0)
    p = c.paragraphs[0]
    para_fmt(p, 0, 0, 1.0, WD_ALIGN_PARAGRAPH.CENTER)
    p.add_run().add_picture(path, width=Cm(width), height=Cm(height))

    cap = cell.add_paragraph()
    para_fmt(cap, 2, 6, 1.0, WD_ALIGN_PARAGRAPH.CENTER)
    r = cap.add_run('▲ 原题截图　' + caption)
    set_run_font(r, cap_size, MUTED, False, False, HEAD_EA, HEAD_ASCII, spacing=10)
    return height


def add_grid(cell, spec, avail=NOTES_W - 0.4, fs=1.0):
    widths = spec['widths']
    scale = min(1.0, avail / sum(widths))
    widths = [w * scale for w in widths]
    ncol = len(widths)
    nrow = len(spec['rows']) + (1 if spec['header'] else 0)
    t = cell.add_table(rows=nrow, cols=ncol)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_fixed(t, sum(widths), widths)
    table_margins(t, 0.06, 0.10, 0.06, 0.10)
    table_borders(t, {k: (4, LINE_HEX) for k in
                      ('top', 'left', 'bottom', 'right', 'insideH', 'insideV')})
    r0 = 0
    if spec['header']:
        for j, txt in enumerate(spec['header']):
            c = t.cell(0, j)
            shade(c, ACCENT_L)
            p = c.paragraphs[0]
            para_fmt(p, 1, 1, 1.0, WD_ALIGN_PARAGRAPH.CENTER)
            add_rich(p, txt, round(9 * fs, 2), ACCENT, True, False, HEAD_EA, HEAD_ASCII)
        row_height(t.rows[0], 0.60)
        r0 = 1
    for i, row in enumerate(spec['rows']):
        for j, txt in enumerate(row):
            c = t.cell(r0 + i, j)
            p = c.paragraphs[0]
            para_fmt(p, 1, 1, 1.0,
                     WD_ALIGN_PARAGRAPH.CENTER if j else WD_ALIGN_PARAGRAPH.LEFT)
            add_rich(p, txt, round(9 * fs, 2), INK, False, False, HEAD_EA, HEAD_ASCII)
        row_height(t.rows[r0 + i], 0.58)
    return t


def build_question(doc, item, first=False):
    no = item['no']
    fs = float(FIT.get('fs', {}).get(no, 1.0))
    figh = float(FIT.get('fig', {}).get(no, FIG_MAX_H))
    side = ('ALL' in SIDE) or (no in SIDE)

    def sz(x):
        return round(x * fs, 2)

    def ln(x):
        return max(1.02, round(x * (0.35 + 0.65 * fs), 3))

    # 分页：每题起新页（首题不加分页符）
    brk = doc.add_paragraph()
    para_fmt(brk, 0, 0, 1.0)
    brk.paragraph_format.page_break_before = not first
    r = brk.add_run('')
    set_run_font(r, 2, INK)

    t = doc.add_table(rows=3, cols=1)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_fixed(t, CONTENT_W, [CONTENT_W])
    table_margins(t, 0.10, 0.18, 0.10, 0.18)
    table_borders(t, {'top': (4, LINE_HEX), 'left': (4, LINE_HEX),
                      'bottom': (4, LINE_HEX), 'right': (4, LINE_HEX),
                      'insideH': (4, LINE_HEX)})

    title_cell = t.cell(0, 0)
    shade(title_cell, ACCENT_L)
    border_cell(title_cell, 'bottom', 12, ACCENT_HEX)
    no_split(t.rows[0])
    p = title_cell.paragraphs[0]
    para_fmt(p, 0.5, 0.5, 1.0)
    p.paragraph_format.tab_stops.add_tab_stop(Cm(CONTENT_W - 0.5), WD_TAB_ALIGNMENT.RIGHT)
    r = p.add_run(item['no'])
    set_run_font(r, 16, ACCENT, True, False, HEAD_ASCII, HEAD_ASCII, spacing=20)
    r = p.add_run('\u3000')
    set_run_font(r, 16, ACCENT, True)
    r = p.add_run(item['title'])
    set_run_font(r, 12.5, INK, True, False, HEAD_EA, HEAD_ASCII)
    r = p.add_run('\t')
    set_run_font(r, 9, MUTED)
    r = p.add_run('%s \u00b7 %s' % (item['src'], item['kind']))
    set_run_font(r, 9, MUTED, False, False, HEAD_EA, HEAD_ASCII)

    if item.get('note'):
        pn = title_cell.add_paragraph()
        para_fmt(pn, 0.5, 1, 1.0, keep_next=True)
        add_rich(pn, '\u6ce8\uff1a' + item['note'], sz(8.5), MUTED, False, True)

    # ---- 主笔记区
    notes = t.cell(1, 0)
    notes.vertical_alignment = WD_ALIGN_VERTICAL.TOP
    label_para(notes, '主笔记区  NOTES', ACCENT, after=2)

    p = notes.add_paragraph()
    para_fmt(p, 0, 1, 1.05)
    add_rich(p, '\u3010\u9898\u76ee\u3011', sz(9.5), ACCENT, True, False, HEAD_EA, HEAD_ASCII)

    for kind, txt in item['q']:
        p = notes.add_paragraph()
        if kind == 'p':
            para_fmt(p, 0, sz(2.2), ln(1.30))
            add_rich(p, txt, sz(10.2), INK)
        elif kind == 'sub':
            para_fmt(p, 0, sz(2.2), ln(1.30), None, 0.0)
            add_rich(p, txt, sz(9.8), INK)
        elif kind == 'opt':
            para_fmt(p, 0, sz(1.2), ln(1.26), None, 0.60, -0.0)
            add_rich(p, txt, sz(9.8), INK)
        elif kind == 'deep':
            para_fmt(p, 0, sz(1.2), ln(1.26), None, 0.86)
            add_rich(p, txt, sz(9.2), INK)
        elif kind == 'tbl':
            spec = item['qtable']
            if spec.get('cap'):
                pc = notes.add_paragraph()
                para_fmt(pc, 1, 1, 1.0, None, 0.2)
                add_rich(pc, spec['cap'], sz(9), MUTED, True, False, HEAD_EA, HEAD_ASCII)
            add_grid(notes, spec, fs=fs)
            sp = notes.add_paragraph()
            para_fmt(sp, 0, 1.0, 0.86)

    def render_analysis(cell):
        p = cell.add_paragraph()
        para_fmt(p, 0, 1, 1.05, None, 0.0)
        add_rich(p, '\u3010\u89e3\u6790\u3011', sz(9.5), ACCENT, True, False, HEAD_EA, HEAD_ASCII)
        for kind, txt in item['anal']:
            q = cell.add_paragraph()
            if kind == 'p':
                para_fmt(q, 0, sz(2.2), ln(1.30))
                add_rich(q, txt, sz(9.8), INK)
            else:
                para_fmt(q, 0, sz(1.6), ln(1.26), None, 0.52, -0.42)
                add_rich(q, '\u25a0 ', sz(8.8), ACCENT, False)
                add_rich(q, txt, sz(9.8), RGBColor(0x33, 0x42, 0x4C))

    if side:
        # 解析（左） + 原题截图（右） 并列
        _sr = float(FIT.get('sidew', {}).get(no, SIDE_R))
        _sl = round(17.04 - _sr, 2)
        st = notes.add_table(rows=1, cols=2)
        st.alignment = WD_TABLE_ALIGNMENT.CENTER
        table_fixed(st, _sl + _sr, [_sl, _sr])
        table_margins(st, 0.02, 0.02, 0.02, 0.10)
        table_borders(st, {})
        lc, rc = st.cell(0, 0), st.cell(0, 1)
        lc.vertical_alignment = WD_ALIGN_VERTICAL.TOP
        rc.vertical_alignment = WD_ALIGN_VERTICAL.TOP
        render_analysis(lc)
        for edge in ('top', 'left', 'bottom', 'right'):
            border_cell(rc, edge, 4, LINE_HEX)
        rc_p = rc.paragraphs[0]
        para_fmt(rc_p, 1, 1, 1.0, WD_ALIGN_PARAGRAPH.CENTER)
        from PIL import Image
        _path = os.path.join(FIGDIR, item['fig'][0])
        _w, _h = Image.open(_path).size
        _pw = _sr - 0.45
        _ph = min(_pw * (_h / _w), 17.0)
        _pw2 = _ph / (_h / _w)
        rc_p.add_run().add_picture(_path, width=Cm(_pw2), height=Cm(_ph))
        cap = rc.add_paragraph()
        para_fmt(cap, 1, 0, 1.0, WD_ALIGN_PARAGRAPH.CENTER)
        r = cap.add_run('\u25b2 原题截图')
        set_run_font(r, sz(8), MUTED, False, False, HEAD_EA, HEAD_ASCII, spacing=10)
        cap2 = rc.add_paragraph()
        para_fmt(cap2, 0, 0, 1.0, WD_ALIGN_PARAGRAPH.CENTER)
        r = cap2.add_run(item['fig'][1])
        set_run_font(r, sz(7.5), MUTED, False, False, HEAD_EA, HEAD_ASCII)
    else:
        add_figure(notes, item['fig'][0], item['fig'][1],
                   avail=CONTENT_W - 0.75, maxh=figh, cap_size=sz(8))
        render_analysis(notes)

    if item.get('q2'):
        for kind, txt in item['q2']:
            p = notes.add_paragraph()
            para_fmt(p, 0, sz(2.2), ln(1.30), None, 0.0)
            add_rich(p, txt, sz(9.8), INK)

    tail = notes.add_paragraph()
    para_fmt(tail, 0, 0, 0.9)
    set_run_font(tail.add_run(''), 3, INK)

    # ---- 答案区
    sum_cell = t.cell(2, 0)
    shade(sum_cell, WARM_L)
    border_cell(sum_cell, 'top', 14, WARM_HEX)
    no_split(t.rows[2])
    label_para(sum_cell, '答案区  ANSWER', WARM, after=1.5)
    for txt in item['summ']:
        p = sum_cell.add_paragraph()
        para_fmt(p, 0, sz(1.2), ln(1.26), None, 0.0)
        add_rich(p, txt, sz(9.8), INK)


# ---------------------------------------------------------------- document
def load_content(path):
    """读入题目数据（列表，按题号顺序）。"""
    with io.open(path, encoding='utf-8') as fh:
        data = json.load(fh)
    if isinstance(data, dict):
        data = data.get('questions', [])
    return data


def main(argv=None):
    ap = argparse.ArgumentParser(description='错题集排版引擎（一题一页）')
    ap.add_argument('--content', required=True, help='content.json')
    ap.add_argument('--figs', default='figs', help='截图目录')
    ap.add_argument('--out', required=True, help='输出 docx')
    ap.add_argument('--fit', default=None, help='fit.json（可选）')
    ap.add_argument('--side', default='', help='截图与解析并列的题号，如 01,02,04')
    ap.add_argument('--pdf', action='store_true', help='用 WPS 顺便导出 PDF')
    ap.add_argument('--header', default='错题集 · 订正笔记', help='页眉左侧文字')
    ap.add_argument('--header-right', default='', help='页眉右侧文字')
    args = ap.parse_args(argv)

    global FIGDIR, SIDE, C
    FIGDIR = args.figs
    _side = args.side.strip()
    SIDE = ({x.strip() for x in _side.split(',') if x.strip()}
            if _side else {'ALL'})      # 默认：每题截图都与解析并列
    C = load_content(args.content)
    if not C:
        raise SystemExit('content.json 里没有题目')
    load_fit(args.fit)

    doc = Document()

    st = doc.styles['Normal']
    st.font.name = BODY_ASCII
    st.font.size = Pt(10)
    st.element.rPr.rFonts.set(qn('w:eastAsia'), BODY_EA)
    st.paragraph_format.space_after = Pt(2)
    st.paragraph_format.line_spacing = 1.34

    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(1.8)
    sec.top_margin = Cm(1.55)
    sec.bottom_margin = Cm(1.35)
    sec.header_distance = Cm(0.78)
    sec.footer_distance = Cm(0.68)
    sec.different_first_page_header_footer = False

    hdr = sec.header
    hdr.is_linked_to_previous = False
    hp = hdr.paragraphs[0]
    para_fmt(hp, 0, 0, 1.0)
    hp.paragraph_format.tab_stops.add_tab_stop(Cm(CONTENT_W), WD_TAB_ALIGNMENT.RIGHT)
    r = hp.add_run(args.header)
    set_run_font(r, 8, MUTED, False, False, HEAD_EA, HEAD_ASCII, spacing=16)
    r = hp.add_run('\t')
    set_run_font(r, 8, MUTED)
    r = hp.add_run(args.header_right)
    set_run_font(r, 8, MUTED, False, False, HEAD_EA, HEAD_ASCII, spacing=16)
    border_para(hp, 'bottom', 4, SOFT_HEX)

    ftr = sec.footer
    ftr.is_linked_to_previous = False
    fp = ftr.paragraphs[0]
    para_fmt(fp, 0, 0, 1.0, WD_ALIGN_PARAGRAPH.CENTER)
    r = fp.add_run('\u2014 ')
    set_run_font(r, 8.5, MUTED)
    field(fp, 'PAGE')
    r = fp.add_run(' \u2014')
    set_run_font(r, 8.5, MUTED)

    for i, item in enumerate(C):
        build_question(doc, item, first=(i == 0))

    fixed = canonicalize(doc.element.body)
    for sec_ in doc.sections:
        for part in (sec_.header, sec_.footer, sec_.first_page_header, sec_.first_page_footer):
            try:
                fixed += canonicalize(part._element)
            except Exception:
                pass
    print('canonicalized elements:', fixed)

    doc.core_properties.title = os.path.splitext(os.path.basename(args.out))[0]
    doc.core_properties.comments = '一题一页：题目、原题截图、解析、答案。'
    doc.save(args.out)
    print('saved:', args.out, '| 题目数:', len(C), '| 并列版式:', sorted(SIDE) or '无')

    if args.pdf:
        print('提示：PDF 导出请运行 assets/measure.ps1 -Pdf，或直接用 WPS/Word 另存为 PDF')


if __name__ == '__main__':
    main()





