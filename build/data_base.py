# -*- coding: utf-8 -*-
"""构建赛事数据集：抽取 PDF 名录 + 补上时间窗与备赛要点。"""
import json, re, os

BASE = os.path.dirname(os.path.abspath(__file__))
rows = json.load(open(os.path.join(BASE, 'raw_rows.json'), encoding='utf-8'))

GROUP_OF_AWARD = {
    '文体活动优秀奖': '文体',
    '精神文明和特殊贡献奖': '精神文明',
}

def clean_name(s):
    s = s.replace('\u3000', ' ').strip()
    s = re.sub(r'\s+', ' ', s)
    s = s.strip('-－ ')
    return s

items = []
for r in rows:
    while len(r) < 6:
        r.append('')
    no, name, level, award, dept, note = r[0], clean_name(r[1]), r[2], r[3], r[4], r[5]
    if not name:
        continue
    if level == '不予认定的' or '不予认定' in award:
        g = '不予认定'
    elif level == '单列':
        g = '单列'
    elif level in ('A类', 'B类'):
        g = level
    else:
        g = GROUP_OF_AWARD.get(award, '文体')
    items.append({'g': g, 'no': no, 'name': name, 'award': award,
                  'dept': dept, 'note': note})

# 组内按序号重排
def sortkey(it):
    try:
        return int(it['no'])
    except Exception:
        return 999
items.sort(key=lambda it: ({'单列': 0, 'A类': 1, 'B类': 2, '文体': 3,
                            '精神文明': 4, '不予认定': 5}[it['g']], sortkey(it)))

order = {'单列': 0, 'A类': 1, 'B类': 2, '文体': 3, '精神文明': 4, '不予认定': 5}
prefix = {'单列': 'S', 'A类': 'A', 'B类': 'B', '文体': 'W', '精神文明': 'P', '不予认定': 'X'}
counter = {}
for it in items:
    counter[it['g']] = counter.get(it['g'], 0) + 1
    # id 用学校原始序号，和纸质材料对得上
    it['id'] = '%s-%s' % (prefix[it['g']], it['no'])
    it['order'] = order[it['g']] * 10000 + counter[it['g']]
