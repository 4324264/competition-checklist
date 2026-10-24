# -*- coding: utf-8 -*-
import json, os, io

BASE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(BASE)
OUT = os.path.join(ROOT, '单项奖学金竞赛规划')
os.makedirs(OUT, exist_ok=True)

data = json.load(open(os.path.join(BASE, 'data.json'), encoding='utf-8'))
tpl = io.open(os.path.join(BASE, 'template_v2.html'), encoding='utf-8').read()

payload = json.dumps(data, ensure_ascii=False, separators=(',', ':'))
html = tpl.replace('__DATA__', payload)

path = os.path.join(OUT, '竞赛待办清单.html')
io.open(path, 'w', encoding='utf-8').write(html)
print(path, round(len(html.encode('utf-8')) / 1024, 1), 'KB')
