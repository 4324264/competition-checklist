# -*- coding: utf-8 -*-
import json, os
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

BASE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(BASE)
OUT = os.path.join(ROOT, '单项奖学金竞赛规划')
os.makedirs(OUT, exist_ok=True)

data = json.load(open(os.path.join(BASE, 'data.json'), encoding='utf-8'))
items = data['items']
feat = [i for i in items if i['featured']]

INK = RGBColor(0x1A, 0x17, 0x14)
INK2 = RGBColor(0x4B, 0x44, 0x3B)
INK3 = RGBColor(0x8A, 0x81, 0x74)
RED = RGBColor(0x9C, 0x3A, 0x28)
GREEN = RGBColor(0x2E, 0x5A, 0x4B)

SERIF = '宋体'
SANS = '微软雅黑'

ORDER_P = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl',
           'numPr', 'suppressLineNumbers', 'pBdr', 'shd', 'tabs', 'suppressAutoHyphens',
           'kinsoku', 'wordWrap', 'overflowPunct', 'topLinePunct', 'autoSpaceDE',
           'autoSpaceDN', 'bidi', 'adjustRightInd', 'snapToGrid', 'spacing', 'ind',
           'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc', 'textDirection',
           'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr',
           'sectPr', 'pPrChange']
ORDER_R = ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike',
           'dstrike', 'outline', 'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid',
           'vanish', 'webHidden', 'color', 'spacing', 'w', 'kern', 'position', 'sz', 'szCs',
           'highlight', 'u', 'effect', 'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs',
           'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath']
ORDER_T = ['tblStyle', 'tblpPr', 'tblOverlap', 'bidiVisual', 'tblStyleRowBandSize',
           'tblStyleColBandSize', 'tblW', 'jc', 'tblCellSpacing', 'tblInd', 'tblBorders',
           'shd', 'tblLayout', 'tblCellMar', 'tblLook', 'tblCaption', 'tblDescription']
ORDER_TC = ['cnfStyle', 'tcW', 'gridSpan', 'hMerge', 'vMerge', 'tcBorders', 'shd',
            'noWrap', 'tcMar', 'textDirection', 'tcFitText', 'vAlign', 'hideMark']


def local(el):
    return el.tag.split('}')[-1]


def insert_ordered(parent, el, order):
    name = local(el)
    pos = order.index(name)
    for child in list(parent):
        cn = local(child)
        if cn in order and order.index(cn) > pos:
            child.addprevious(el)
            return
    parent.append(el)


def run(p, text, size=10.5, bold=False, color=INK, font=SERIF, spacing=None):
    r = p.add_run(text)
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = color
    r.font.name = font
    r._element.rPr.rFonts.set(qn('w:eastAsia'), font)
    if spacing is not None:
        rPr = r._element.get_or_add_rPr()
        sp = OxmlElement('w:spacing')
        sp.set(qn('w:val'), str(int(spacing * 20)))
        insert_ordered(rPr, sp, ORDER_R)
    return r


def para(space_before=0, space_after=0, line=1.55, align=None, indent=None, hang=None,
         keep=False, page_break=False):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    pf.line_spacing = line
    if align is not None:
        p.alignment = align
    if indent is not None:
        pf.left_indent = Cm(indent)
    if hang is not None:
        pf.first_line_indent = Cm(-hang)
    pf.keep_with_next = keep
    pf.page_break_before = page_break
    return p


def add_border(p, edge='bottom', val='single', sz=6, color='D5CAB7', space=1):
    pPr = p._p.get_or_add_pPr()
    pbdr = pPr.find(qn('w:pBdr'))
    if pbdr is None:
        pbdr = OxmlElement('w:pBdr')
        insert_ordered(pPr, pbdr, ORDER_P)
    e = OxmlElement('w:' + edge)
    e.set(qn('w:val'), val)
    e.set(qn('w:sz'), str(sz))
    e.set(qn('w:space'), str(space))
    e.set(qn('w:color'), color)
    pbdr.append(e)


def hr(color='D5CAB7', sz=6, space_before=6, space_after=10):
    p = para(space_before=space_before, space_after=space_after)
    p.paragraph_format.line_spacing = Pt(1)
    r = p.add_run('')
    r.font.size = Pt(1)
    add_border(p, 'bottom', 'single', sz, color, 1)


def shade(cell, hexcolor):
    tcPr = cell._tc.get_or_add_tcPr()
    sh = OxmlElement('w:shd')
    sh.set(qn('w:val'), 'clear')
    sh.set(qn('w:color'), 'auto')
    sh.set(qn('w:fill'), hexcolor)
    insert_ordered(tcPr, sh, ORDER_TC)


def cell_margins(table, top=70, bottom=70, left=110, right=110):
    tblPr = table._tbl.tblPr
    mar = OxmlElement('w:tblCellMar')
    for side, v in (('top', top), ('left', left), ('bottom', bottom), ('right', right)):
        e = OxmlElement('w:' + side)
        e.set(qn('w:w'), str(v))
        e.set(qn('w:type'), 'dxa')
        mar.append(e)
    insert_ordered(tblPr, mar, ORDER_T)


def table_borders(table, color='D5CAB7', sz=4, left=None):
    tblPr = table._tbl.tblPr
    b = OxmlElement('w:tblBorders')
    spec = [('top', 4, color), ('bottom', 4, color),
            ('left', 18 if left else 4, left or color), ('right', 4, color),
            ('insideH', 4, color), ('insideV', 4, 'EDE5D6')]
    for edge, s, col in spec:
        e = OxmlElement('w:' + edge)
        if s == 0:
            e.set(qn('w:val'), 'nil')
        else:
            e.set(qn('w:val'), 'single')
            e.set(qn('w:sz'), str(s))
            e.set(qn('w:space'), '0')
            e.set(qn('w:color'), col)
        b.append(e)
    insert_ordered(tblPr, b, ORDER_T)


def cell_text(cell, text, size=9, bold=False, color=INK, font=SANS, align=None, line=1.4):
    cell.text = ''
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.line_spacing = line
    if align is not None:
        p.alignment = align
    run(p, text, size=size, bold=bold, color=color, font=font)
    return p


def set_col_widths(table, widths):
    table.autofit = False
    tblPr = table._tbl.tblPr
    if tblPr.find(qn('w:tblLayout')) is None:
        layout = OxmlElement('w:tblLayout')
        layout.set(qn('w:type'), 'fixed')
        insert_ordered(tblPr, layout, ORDER_T)
    for old in tblPr.findall(qn('w:tblW')):
        tblPr.remove(old)
    tw = OxmlElement('w:tblW')
    tw.set(qn('w:type'), 'dxa')
    tw.set(qn('w:w'), str(int(sum(widths) * 567)))
    insert_ordered(tblPr, tw, ORDER_T)
    for i, w in enumerate(widths):
        table.columns[i].width = Cm(w)
    for row in table.rows:
        for i, w in enumerate(widths):
            row.cells[i].width = Cm(w)


def keep_rows(table, header=True):
    for ri, row in enumerate(table.rows):
        trPr = row._tr.get_or_add_trPr()
        e = OxmlElement('w:cantSplit')
        insert_ordered(trPr, e, ['cnfStyle', 'divId', 'gridBefore', 'gridAfter', 'wBefore',
                                 'wAfter', 'cantSplit', 'trHeight', 'tblHeader',
                                 'tblCellSpacing', 'jc', 'hidden'])
        if header and ri == 0:
            th = OxmlElement('w:tblHeader')
            insert_ordered(trPr, th, ['cnfStyle', 'divId', 'gridBefore', 'gridAfter',
                                      'wBefore', 'wAfter', 'cantSplit', 'trHeight',
                                      'tblHeader', 'tblCellSpacing', 'jc', 'hidden'])


doc = Document()
sec = doc.sections[0]
sec.page_width = Cm(21.0)
sec.page_height = Cm(29.7)
sec.top_margin = Cm(2.3)
sec.bottom_margin = Cm(2.1)
sec.left_margin = Cm(2.4)
sec.right_margin = Cm(2.4)
CW = 21.0 - 4.8

st = doc.styles['Normal']
st.font.name = SERIF
st.font.size = Pt(10.5)
st.font.color.rgb = INK
rPr = st.element.get_or_add_rPr()
rf = rPr.find(qn('w:rFonts'))
if rf is None:
    rf = OxmlElement('w:rFonts')
    insert_ordered(rPr, rf, ORDER_R)
rf.set(qn('w:eastAsia'), SERIF)
st.paragraph_format.line_spacing = 1.55
st.paragraph_format.space_before = Pt(0)
st.paragraph_format.space_after = Pt(0)

# footer
fp = sec.footer.paragraphs[0]
fp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
run(fp, '单项奖学金竞赛规划 · 2026—2027　　　', size=8, color=INK3, font=SANS)
fld = OxmlElement('w:fldSimple')
fld.set(qn('w:instr'), 'PAGE')
r = OxmlElement('w:r')
rPr2 = OxmlElement('w:rPr')
sz = OxmlElement('w:sz'); sz.set(qn('w:val'), '16'); rPr2.append(sz)
col = OxmlElement('w:color'); col.set(qn('w:val'), '8A8174'); rPr2.append(col)
r.append(rPr2)
t = OxmlElement('w:t'); t.text = '1'; r.append(t)
fld.append(r)
fp._p.append(fld)

# ---------- 封面 ----------
p = para(space_before=4, space_after=5)
run(p, '郑州轻工业大学　单项奖学金评定参考项目', size=8.5, color=INK3, font=SANS, spacing=0.8)

p = para(space_before=2, space_after=4, line=1.02)
run(p, '竞赛时间与备赛清单', size=28, bold=True, color=INK, font=SANS, spacing=-0.6)

p = para(space_before=2, space_after=2)
run(p, '2026—2027 学年　·　数字媒体技术　·　编制于 2026 年 10 月', size=9, color=INK2, font=SANS)

hr(space_before=4, space_after=13)

p = para(space_after=7, line=1.72)
run(p, '这份材料把学校 2026 年认定的 ', size=10)
run(p, '328 项', size=10, bold=True)
run(p, '赛事按认定顺序过了一遍。其中 39 项写清了赛程节点、要交的材料和容易出错的地方，'
       '剩下 289 项在网页版里逐条列着，能勾选，能搜索。', size=10)

p = para(space_after=12, line=1.72)
run(p, '时间有两种标注。标「已核实」的，查到了 2026 或 2027 年的官方通知或高校公开文件；'
       '标「按往年」的，是按往届规律推的，报名前必须再看一眼本校通知。', size=10, color=INK2)

# 摘要框
box = doc.add_table(rows=1, cols=1)
box.alignment = WD_TABLE_ALIGNMENT.LEFT
table_borders(box, left='9C3A28')
cell_margins(box, top=140, bottom=140, left=210, right=170)
c = box.rows[0].cells[0]
shade(c, 'FBF7EE')
c.width = Cm(CW)
c.text = ''
p0 = c.paragraphs[0]
p0.paragraph_format.space_after = Pt(8)
p0.paragraph_format.line_spacing = 1
run(p0, '只看三行的话', size=8, color=RED, font=SANS, spacing=1.0)

lines = [
    ('最急的一件事', '蓝桥杯软件赛 12 月 18 日 14:00 关报名，视觉艺术设计赛道到 2027 年 4 月 26 日。'),
    ('最省力的一件事', '中国好创意已经开放，5 月 31 日截稿。拿一份课程作业改一改先投进去。'),
    ('最该盯的三项', '大广赛、中国大学生计算机设计大赛、全国大学生数字媒体科技作品及创意竞赛。'),
]
for i, (k, v) in enumerate(lines):
    p1 = c.add_paragraph()
    p1.paragraph_format.space_after = Pt(5 if i < 2 else 0)
    p1.paragraph_format.line_spacing = 1.55
    p1.paragraph_format.left_indent = Cm(2.4)
    p1.paragraph_format.first_line_indent = Cm(-2.4)
    p1.paragraph_format.tab_stops.add_tab_stop(Cm(2.4))
    run(p1, k, size=9.5, bold=True, color=INK, font=SANS)
    run(p1, '\t', size=9.5)
    run(p1, v, size=9.5, color=INK2)

# ---------- 一 ----------
p = para(space_before=0, space_after=2, keep=True, page_break=True)
run(p, '01', size=9, color=RED, font=SANS, bold=True, spacing=1.2)
run(p, '　现在就能动手的事', size=15, bold=True, color=INK, font=SANS)
p = para(space_after=9)
run(p, '现在是 2026 年 11 月。下面八件事都在报名或收稿窗口里，按截止先后排。', size=9.5, color=INK3, font=SANS)

rows_now = [
    ('1', '蓝桥杯　软件赛', '2026.12.18', '官网注册，选 C/C++ 或 Python 组，缴费。刷近五年真题，寒假每天两小时够用。'),
    ('2', '三创赛　报名', '2026.12 底', '组 3 到 5 人队，先定一个能落地的小生意，寒假写完商业计划书初稿。'),
    ('3', '中国好创意', '2027.5.31', '已经开放。挑一件课程作业改成参赛版，先把位置占上。'),
    ('4', '服务外包大赛', '2026.12 发题', '盯企业命题，寒假把原型做出来，三月初赛，五月决赛。'),
    ('5', '全国大学生英语竞赛', '2026.12 起', '报名费三十元左右，选 C 类。寒假刷完近五年真题就能拿奖。'),
    ('6', '未来设计师 NCDA', '2026.12 开放', '先问清系里哪位老师在收件，这个比赛不接受个人报名。'),
    ('7', '挑战杯　校赛', '2026.10 起', '找一位手上有项目的老师，以成员身份进组，别自己从零起题。'),
    ('8', '互联网+　校赛', '2026.12 启动', '同上，先当团队成员。大一做主创基本走不到省赛。'),
]
t = doc.add_table(rows=len(rows_now) + 1, cols=4)
table_borders(t)
cell_margins(t, top=62, bottom=62, left=95, right=95)
for i, h in enumerate(['#', '赛事', '截止 / 时间', '这一段该做什么']):
    cell_text(t.rows[0].cells[i], h, size=8.5, bold=True, color=INK3, font=SANS)
    shade(t.rows[0].cells[i], 'F3EDE1')
for ri, row in enumerate(rows_now, start=1):
    for ci, v in enumerate(row):
        cell_text(t.rows[ri].cells[ci], v, size=8.5,
                  color=INK if ci in (1, 3) else INK2,
                  font=SANS if ci == 1 else SERIF,
                  align=WD_ALIGN_PARAGRAPH.CENTER if ci == 0 else None)
set_col_widths(t, [0.9, 3.9, 3.0, CW - 7.8])
keep_rows(t)

# ---------- 二 ----------
p = para(space_before=22, space_after=2, keep=True)
run(p, '02', size=9, color=RED, font=SANS, bold=True, spacing=1.2)
run(p, '　一年的节奏', size=15, bold=True, color=INK, font=SANS)
p = para(space_after=9)
run(p, '校赛挤在三月到五月，国赛挤在七月到九月。真正决定成败的是寒假，'
       '大部分作品都是二月做完的，开学只是交件。', size=9.5, color=INK3, font=SANS)

rows_year = [
    ('2026.11', '挑战杯校赛、好创意征稿、蓝桥杯报名中', '只定两三个方向，别铺开。方向定了就开始攒素材。'),
    ('2026.12', '蓝桥杯软件赛报名截止、服务外包发题、NCDA 开放', '把名报完，把寒假要做的作品定下来。'),
    ('2027.01', '大广赛命题发布、米兰设计周启动', '寒假第一周把命题抄下来，逐条对评分点。'),
    ('2027.02', '春节，赛场最安静的一个月', '作品主体必须在这个月做完。开学后没有整块时间。'),
    ('2027.03', '计算机设计大赛校赛、三创赛校赛、大广赛创作', '校赛撞在一起，选两项，别三项都报。'),
    ('2027.04', '蓝桥杯省赛、计算机设计大赛省赛', '备赛冲刺。省赛过了才有国赛。'),
    ('2027.05', '好创意截稿、大广赛提交、英语竞赛决赛、服务外包决赛', '全年最忙的一个月，交件清单提前一周列好。'),
    ('2027.06', '大广赛校赛、NCDA 校内预选、好创意评审', '收尾，同时开始准备期末。'),
    ('2027.07', '大广赛省赛、计算机设计大赛国赛、华灿奖启动', '暑假主攻国赛，顺手启动下一件作品。'),
    ('2027.08', '大广赛国赛、数媒科技作品竞赛报名、计算机设计大赛国赛', '出结果，同时报八月的比赛。'),
    ('2027.09', '数学建模国赛、NCDA 公布、3D 大赛', '开学第一周就比赛，暑假必须留校训练。'),
    ('2027.10', '互联网+ 国赛、挑战杯国赛、3D 大赛国赛', '冲最高级别。看的是前一年攒下的东西。'),
    ('2027.11—12', '互联网+ 总决赛、挑战杯决赛、外研社国赛', '收尾，复盘，定下一年报哪几项。'),
]
t = doc.add_table(rows=len(rows_year) + 1, cols=3)
table_borders(t)
cell_margins(t, top=58, bottom=58, left=95, right=95)
for i, h in enumerate(['时间', '赛场上在发生什么', '你该做什么']):
    cell_text(t.rows[0].cells[i], h, size=8.5, bold=True, color=INK3, font=SANS)
    shade(t.rows[0].cells[i], 'F3EDE1')
for ri, row in enumerate(rows_year, start=1):
    cell_text(t.rows[ri].cells[0], row[0], size=8.5, color=RED, font=SANS, bold=True)
    cell_text(t.rows[ri].cells[1], row[1], size=8.5, color=INK2, font=SANS)
    cell_text(t.rows[ri].cells[2], row[2], size=8.5, color=INK, font=SERIF)
set_col_widths(t, [2.3, 7.0, CW - 9.3])
keep_rows(t)

# ---------- 三 ----------
p = para(space_before=22, space_after=2, keep=True)
run(p, '03', size=9, color=RED, font=SANS, bold=True, spacing=1.2)
run(p, '　39 项重点赛事', size=15, bold=True, color=INK, font=SANS)
p = para(space_after=9)
run(p, '按学校认定表的原始分组和序号排。每条只写四件事：什么时候、准备什么、谁能报、哪里容易错。',
    size=9.5, color=INK3, font=SANS)

GROUP_ORDER = [
    ('单列', '单列项目', '三项都是团队赛，大一先当成员，大二再当负责人。'),
    ('A类', 'A 类竞赛', '加分主力。先挑和专业对口的，别贪多。'),
    ('B类', 'B 类竞赛', '价值在于稳。选一两个能确定拿奖的，不要分散。'),
    ('文体', '文体活动竞赛', '挑能复用课程作业的项目，一份作品投两个地方。'),
    ('精神文明', '精神文明和特殊贡献奖', '要长期项目，临时组队来不及。'),
]


def field(label, value, color=INK2, size=9.5):
    p = para(space_before=0, space_after=2, line=1.58, indent=1.3, hang=1.3)
    p.paragraph_format.tab_stops.add_tab_stop(Cm(1.3))
    run(p, label, size=8, color=INK3, font=SANS)
    run(p, '\t', size=8)
    run(p, value, size=size, color=color)
    return p


marks = '①②③④⑤⑥⑦⑧⑨'
first_group = True
for gkey, gname, gtip in GROUP_ORDER:
    glist = [i for i in feat if i['g'] == gkey]
    if not glist:
        continue
    p = para(space_before=18 if not first_group else 8, space_after=4, keep=True)
    first_group = False
    run(p, gname, size=12, bold=True, color=INK, font=SANS)
    run(p, '　　' + gtip, size=8.5, color=INK3, font=SANS)
    hr(color='1A1714', sz=8, space_before=1, space_after=9)

    for idx, it in enumerate(glist):
        p = para(space_before=9 if idx else 3, space_after=2, keep=True)
        run(p, it['id'].split('-')[0] + '·' + it['no'] + '　', size=8, color=RED, font=SANS, bold=True)
        run(p, it['name'], size=10.5, bold=True, color=INK, font=SANS)
        if it.get('st') == 'now':
            run(p, '　〔报名中〕', size=8, color=GREEN, font=SANS)

        p = para(space_after=5, keep=True)
        run(p, it['award'] + '　/　' + it['dept'], size=8, color=INK3, font=SANS)
        run(p, '　　' + ('时间已核实' if it.get('conf') == '已核实' else '时间按往年推算'),
            size=8, color=GREEN if it.get('conf') == '已核实' else INK3, font=SANS)
        run(p, '　　' + it['win'], size=8, color=INK3, font=SANS)

        nodes = '；'.join('%s %s' % (a, b) for a, b in it['tl'])
        field('节点', nodes).paragraph_format.keep_with_next = True
        prep = '；'.join(marks[i] + it['prep'][i] for i in range(len(it['prep'])))
        field('准备', prep).paragraph_format.keep_with_next = True
        field('规则', it['rule']).paragraph_format.keep_with_next = True
        field('提醒', it['warn'], color=RED)

        if idx < len(glist) - 1:
            p = para(space_before=1, space_after=0, line=1)
            add_border(p, 'bottom', 'dotted', 4, 'D5CAB7', 4)

# ---------- 四 ----------
p = para(space_before=24, space_after=2, keep=True, page_break=True)
run(p, '04', size=9, color=RED, font=SANS, bold=True, spacing=1.2)
run(p, '　三条主线，选一条走', size=15, bold=True, color=INK, font=SANS)
p = para(space_after=11)
run(p, '赛事太多，全报等于全不报。数字媒体技术能走的其实只有三条路，选一条，剩下的当备选。',
    size=9.5, color=INK3, font=SANS)

lines_main = [
    ('作品线', '大广赛　→　计算机设计大赛　→　好创意 / NCDA　→　数媒科技作品竞赛　→　华灿奖',
     '最贴专业，产出是作品集。大广赛三月启动，计算机设计大赛三月校赛，两件作品可以同源。'
     '这条线的东西直接进简历和保研材料，投入产出最高。'),
    ('算法线', '蓝桥杯　→　天梯赛 / CCPC　→　数学建模',
     '起步最早，十二月就能报名。缺点是见效慢，要持续刷题。'
     '如果打算考研或者进开发岗，这条线值得走。'),
    ('创业线', '三创赛　→　服务外包　→　互联网+　→　挑战杯',
     '需要团队和指导老师，大一只能当成员。好处是综合测评里分值高，'
     '而且能认识一批跨专业的人。'),
]
for k, path, note in lines_main:
    p = para(space_before=10, space_after=3, keep=True)
    run(p, k, size=11, bold=True, color=RED, font=SANS)
    run(p, '　' + path, size=9.5, color=INK, font=SANS)
    p = para(space_after=2, line=1.72, indent=0.6)
    run(p, note, size=9.5, color=INK2)

# ---------- 五 ----------
p = para(space_before=26, space_after=2, keep=True)
run(p, '05', size=9, color=RED, font=SANS, bold=True, spacing=1.2)
run(p, '　申报环节容易踩的坑', size=15, bold=True, color=INK, font=SANS)
p = para(space_after=11)
run(p, '获奖只是第一步。材料报上去被退回，等于没参加。下面六条在认定表里写着。',
    size=9.5, color=INK3, font=SANS)

pits = [
    ('申报类别选错，归口部门不受理。',
     '学科类走学生处；名称里带「创业」或「创新创业」的走毕业生就业指导中心；'
     '只带「创新」的走校团委；「挑战杯」两项都走校团委。表末说明第 1 条写得很清楚。'),
    ('团队赛要确认是否发个人证书。',
     '全国大学生智能汽车竞赛、西门子杯在备注里专门标了「团队赛发放个人证书」。'
     '没标的要提前问组委会，否则证书上只有队名，认定时用不上。'),
    ('证书上没有成员姓名的，要补材料。',
     'iFLYTEK 开发者大赛属于这种情况，备注要求提供其他证明材料一并上传系统。'),
    ('一批 B 类按省部级认定。',
     '全国大学生英语竞赛、CADA 日本概念艺术奖、新加坡金沙艺术设计大赛等都标了「按照省部级认定」。'
     '加分比国赛低一档，别把主力时间放进去。'),
    ('十五项不予认定。',
     '包括创新未来设计大赛、圆明园文创设计大赛、大学生就业能力大赛等。'
     '报名前对一眼，别交冤枉报名费。完整名单在网页版最后一组。'),
    ('五月和六月是交件高峰。',
     '大广赛、好创意、NCDA、计算机设计大赛省赛全挤在这两个月，还要应付期末。'
     '三四月就要把作品做到八成，不要留到五月。'),
]
for i, (h, b) in enumerate(pits, 1):
    p = para(space_before=9, space_after=2, line=1.6, indent=0.8, hang=0.8, keep=True)
    run(p, marks[i - 1], size=10, color=RED, font=SANS, bold=True)
    run(p, '　', size=10)
    run(p, h, size=10, bold=True, color=INK, font=SANS)
    p = para(space_after=0, line=1.72, indent=0.8)
    run(p, b, size=9.5, color=INK2)

# ---------- 六 ----------
p = para(space_before=24, space_after=2, keep=True)
run(p, '06', size=9, color=RED, font=SANS, bold=True, spacing=1.2)
run(p, '　交件前对一遍', size=15, bold=True, color=INK, font=SANS)
p = para(space_after=10)
run(p, '每一条都可能让材料退回。交之前逐项过一遍，别在最后一天才发现少东西。',
    size=9.5, color=INK3, font=SANS)

rows_check = [
    ('作品文件', '源文件和导出件都要。只交 JPG 或 MP4，评审想看工程文件时拿不出来，会被判不完整。'),
    ('报名表与承诺书', '全员签名。少一个作者的签名，整组作废，而且是提交后才发现。'),
    ('学生证或学信网截图', '每组只需第一作者上传，但要能看清学号和姓名。'),
    ('查重报告', '挑战杯、数学建模都有重复率要求，一般 15% 以下，自己先查一遍。'),
    ('版权与原创声明', '用到的字体、音乐、素材要有授权。商用字体最容易出事。'),
    ('展板与文件命名', '尺寸、dpi、命名规则错一项就退件。按参赛指南逐条对。'),
    ('提交成功截图', '官网一关就没法补交。截一张带时间和编号的图，存在相册里。'),
    ('纸质材料', '多数学校要求交到指定办公室，只传电子版不算完成。'),
]
t = doc.add_table(rows=len(rows_check) + 1, cols=2)
table_borders(t)
cell_margins(t, top=60, bottom=60, left=95, right=95)
for i, h in enumerate(['要带的东西', '哪里容易出错']):
    cell_text(t.rows[0].cells[i], h, size=8.5, bold=True, color=INK3, font=SANS)
    shade(t.rows[0].cells[i], 'F3EDE1')
for ri, row in enumerate(rows_check, start=1):
    cell_text(t.rows[ri].cells[0], row[0], size=8.5, color=INK, font=SANS, bold=True)
    cell_text(t.rows[ri].cells[1], row[1], size=8.5, color=INK2, font=SERIF)
set_col_widths(t, [3.9, CW - 3.9])
keep_rows(t)

# ---------- 尾 ----------
hr(color='D5CAB7', sz=4, space_before=26, space_after=11)
p = para(space_after=5)
run(p, '剩下的 289 项', size=10, bold=True, color=INK, font=SANS)
p = para(space_after=5, line=1.75)
run(p, 'B 类、文体、精神文明这三组里大部分赛事，要么专业不对口，要么只认省部级，'
       '要么一年只办一次、时间随主办方变。逐条写时间意义不大，写出来也是估的。'
       '这些条目在网页版里全都在，按组排列，可以勾选，可以搜索，点开能看到认定级别和归口部门。',
    size=9.5, color=INK2)
p = para(space_after=0, line=1.75)
run(p, '网页文件　', size=9, color=INK3, font=SANS)
run(p, '竞赛待办清单.html', size=9.5, color=RED, font=SANS)

path = os.path.join(OUT, '竞赛时间与备赛清单.docx')
doc.save(path)
print('saved', path)
