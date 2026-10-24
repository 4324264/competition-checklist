# -*- coding: utf-8 -*-
"""把工作区的产物同步进本仓库，然后提交并推送。
用法：
    python tools/sync.py                  # 自动生成提交消息
    python tools/sync.py "feat: 收紧留白"  # 指定提交消息
"""
import os, shutil, subprocess, sys
from pathlib import Path

PUB = Path(__file__).resolve().parents[1]
WS = PUB.parent

PAIRS = [
    (WS / '单项奖学金竞赛规划' / '竞赛待办清单.html', PUB / 'index.html'),
    (WS / '单项奖学金竞赛规划' / '竞赛时间与备赛清单.docx', PUB / '竞赛时间与备赛清单.docx'),
]
for name in ['data.json', 'data_base.py', 'data.py', 'features.py',
             'make_docx.py', 'render.py', 'template_v2.html', 'raw_rows.json']:
    PAIRS.append((WS / 'build' / name, PUB / 'build' / name))

copied = []
for src, dst in PAIRS:
    if not src.exists():
        continue
    if dst.exists() and src.stat().st_size == dst.stat().st_size and \
       src.read_bytes() == dst.read_bytes():
        continue
    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, dst)
    copied.append(dst.relative_to(PUB).as_posix())


def git(*args, **kw):
    return subprocess.run(['git'] + list(args), cwd=str(PUB),
                          capture_output=True, text=True, **kw)


git('add', '-A')
staged = [l[3:] for l in git('diff', '--cached', '--name-only').stdout.splitlines() if l.strip()]
if not staged:
    print('没有改动，跳过提交')
    sys.exit(0)

msg = sys.argv[1] if len(sys.argv) > 1 else None
if not msg:
    kinds = set()
    for f in staged:
        if f == 'index.html':
            kinds.add('网页')
        elif f.endswith('.docx'):
            kinds.add('文档')
        elif f.startswith('build/'):
            kinds.add('生成脚本')
        elif f.endswith('.md'):
            kinds.add('说明')
        else:
            kinds.add('其他')
    head = '/'.join(sorted(kinds))
    body = '、'.join(staged[:4]) + ('等 %d 个文件' % len(staged) if len(staged) > 4 else '')
    msg = 'chore: 同步%s（%s）' % (head, body)

git('commit', '-q', '-m', msg)
p = git('push', 'origin', 'main')
print('提交：', msg)
print('push:', 'OK' if p.returncode == 0 else p.stderr.strip()[:160])
