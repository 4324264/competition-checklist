# -*- coding: utf-8 -*-
import json, os, sys

BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)

from data_base import items  # noqa: E402
from features import F       # noqa: E402

SPORT = ['体育', '田径', '马拉松', '球', '操', '武术', '健身', '健美', '围棋', '桥牌',
         '射击', '滑雪', '雪雕', '运动', '街舞', '拔河']
INTL = ['国际', '全球', '世界', '亚太', '美国', '日本', '英国', '意大利', '新加坡',
        '韩国', '中英', '丝路', '中德', '两岸', '港澳']
DESIGN = ['设计', '创意', '艺术', '广告', '摄影', '微视频', '短视频', '动漫', '动画',
          '插画', '包装', '视觉', '数媒', '数字媒体', '文创', '影视']
AUTUMN = ['知识竞赛', '答题', '演讲', '征文', '翻译', '词汇', '英语', '数学', '算法',
          '程序设计', '编程', '信息安全', '网络安全', '大数据', '人工智能', '机器人',
          '智能', '建模']


ALGO = ['ICPC', 'CCPC', '程序设计', '算法', '编程', 'ACM']


def season(name):
    if any(k in name for k in ALGO):
        return '秋季', '9—12 月'
    if any(k in name for k in SPORT):
        return '随赛历', '多为 4—6 月或 10—12 月'
    if any(k in name for k in INTL) and any(k in name for k in DESIGN):
        return '全年', '分批征稿，多为 10—3 月'
    if any(k in name for k in INTL):
        return '全年', '分批征稿'
    if any(k in name for k in DESIGN):
        return '春季', '3—6 月'
    if any(k in name for k in AUTUMN):
        return '春秋两季', '3—6 月或 9—12 月'
    return '春季', '3—6 月'


NAME_FIX = {
    'A-4': 'ACM-ICPC 国际大学生程序设计竞赛',
    'A-15': '外研社全国大学生英语系列赛（演讲 / 辩论 / 写作 / 阅读）',
    'A-13': '两岸新锐设计竞赛·华灿奖',
    'A-15': '全国大学生机器人大赛（RoboMaster / RoboCon）',
    'A-19': '中国高校计算机大赛（大数据挑战赛 / 团体程序设计天梯赛 / 移动应用创新赛 / 网络技术挑战赛 / 人工智能创意赛）',
    'A-21': '米兰设计周·中国高校设计学科师生优秀作品展',
    'A-26': '未来设计师·全国高校数字艺术设计大赛',
    'A-31': '“大唐杯”全国大学生新一代信息通信技术大赛',
    'A-36': '全国高校商业精英挑战赛（品牌策划 / 会展 / 国际贸易 / 创新创业 / 会计与商业管理案例）',
    'A-43': '百度之星·程序设计大赛',
    'A-59': '全国大学生电子商务“创新、创意及创业”挑战赛',
    'A-12': '外研社全国大学生英语系列赛（演讲 / 辩论 / 写作 / 阅读）',
    'B-7': '人工智能未来设计大赛·新工科创新数字技能竞赛',
    'B-18': 'G-CROSS 全球大学生创意金星奖',
    'B-25': '长风学霸赛之传奇大战（供应链运营）',
    'B-29': '“学宪法，讲宪法”知识竞赛',
    'B-65': '五一数学建模竞赛',
    'B-82': '“外教社·词达人杯”全国大学生英语词汇能力大赛',
    'B-83': '“傅利叶”杯中国康复人创意大赛',
    'B-101': 'iFLYTEK A.I. 开发者大赛',
    'B-106': '“新华三杯”全国大学生数字技术大赛',
    'B-107': '“外研社·国才杯”“理解当代中国”全国大学生外语能力大赛',
    'B-147': '开放原子大赛 openEuler 应用软件开发赛',
    'B-148': '“新时代·新梦想”河南省大学生创新创业优秀项目选拔赛',
    'B-161': '“三菱电机杯”全国大学生电气与自动化大赛（创新设计赛项）',
    'W-1': '“讲好节能故事”微视频、摄影及征文大赛',
    'W-8': '河南省大学生第二十一届田径运动会',
    'W-23': '河南省高校大学生“中华民族一家亲 同心共筑中国梦”主题演讲比赛',
    'W-50': '河南省学校体育竞赛（乒乓球 / 网球 / 羽毛球）',
    'W-52': '河南省“领航·网络素养新风尚”校园评论大赛',
    'P-1': '“助学·筑梦·铸人”主题活动',
    'P-2': '河南省文明学生、文明宿舍、文明班级、文明社团',
    'P-3': '“活力杯”河南学校共青团基层基础工作大赛（“十佳活力团支部”“十佳魅力团支书”）',
    'X-1': '“阅百年历程，传精神力量，展馆员风采”主题汇演活动',
}

GROUP_META = {
    '单列': ('单列项目', '学校单列认定的三项双创赛事，级别最高'),
    'A类': ('A 类竞赛', '学科与创新创业两条线，加分权重最高'),
    'B类': ('B 类竞赛', '多数按省部级认定，适合积累第一批证书'),
    '文体': ('文体活动竞赛', '艺术、体育、网络文化类'),
    '精神文明': ('精神文明和特殊贡献奖', '按项目金额认定，团队类更划算'),
    '不予认定': ('不予认定', '学校明确不认，不要投入时间'),
}
GROUP_TIP = {
    '单列': '三项都是团队赛，大一先当成员，大二再当负责人。',
    'A类': 'A 类是加分主力。先挑和专业对口的，别贪多。',
    'B类': 'B 类的价值在于"稳"。选一两个能确定拿奖的，不要分散。',
    '文体': '挑能复用课程作业的项目，一份作品投两个地方。',
    '精神文明': '要长期项目，临时组队来不及。',
    '不予认定': '这 15 项学校不认，列在这里只是提醒你别被骗报名费。',
}

out = []
for it in items:
    it = dict(it)
    if it['id'] in NAME_FIX:
        it['name'] = NAME_FIX[it['id']]
    f = F.get((it['g'], it['no']))
    it['featured'] = bool(f)
    if f:
        it.update(f)
    elif it['g'] == '不予认定':
        it['season_label'] = None
        it['season_range'] = None
        it['conf'] = '—'
    else:
        s, rng = season(it['name'])
        it['season_label'] = s
        it['season_range'] = rng
        it['conf'] = '估算'
    out.append(it)

# 重点项缺失检查
missing = [k for k in F if k not in {(i['g'], i['no']) for i in items}]
if missing:
    print('WARN missing:', missing)

stats = {}
for it in out:
    stats[it['g']] = stats.get(it['g'], 0) + 1

data = {
    'meta': {
        'school': '郑州轻工业大学',
        'title': '单项奖学金评定参考项目',
        'year': '2026—2027',
        'major': '数字媒体技术',
        'built': '2026-10-24',
        'total': len(out),
        'featured': sum(1 for i in out if i['featured']),
        'groups': [{'key': k, 'name': GROUP_META[k][0], 'desc': GROUP_META[k][1],
                    'tip': GROUP_TIP[k], 'count': stats.get(k, 0)} for k in GROUP_META],
    },
    'items': out,
}
json.dump(data, open(os.path.join(BASE, 'data.json'), 'w', encoding='utf-8'),
          ensure_ascii=False, indent=1)
print('items', len(out), 'featured', data['meta']['featured'], stats)
