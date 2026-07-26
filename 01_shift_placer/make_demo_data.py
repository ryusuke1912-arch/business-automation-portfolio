# -*- coding: utf-8 -*-
"""
デモ用の入力データを生成する。
※ 生成される氏名・都合はすべて架空のサンプルであり、実在の人物とは関係ありません。
出力: teachers.csv / students.csv / dates.csv
"""
import csv, random, datetime, os

random.seed(20260726)
HERE = os.path.dirname(os.path.abspath(__file__))
SLOTS = ['13:00', '14:30', '17:00', '18:30', '20:00']
WD = '月火水木金土日'

# ---------- 稼働日（日曜休み） ----------
start = datetime.date(2026, 7, 25)
end = datetime.date(2026, 8, 10)
dates = []
d = start
while d <= end:
    if d.weekday() != 6:                       # 日曜は休講
        dates.append(d)
    d += datetime.timedelta(days=1)

with open(os.path.join(HERE, 'dates.csv'), 'w', newline='', encoding='utf-8-sig') as f:
    w = csv.writer(f)
    w.writerow(['日付', '曜日', 'シート名'])
    for d in dates:
        wd = WD[d.weekday()]
        w.writerow([f'{d.month}/{d.day}', wd, f'{d.month}_{d.day}（{wd}）'])

# ---------- 講師（架空） ----------
TEACHERS = ['青木 涼', '伊藤 芽衣', '大西 拓真', '川口 千尋',
            '佐々木 蓮', '田村 結衣', '中島 悠斗', '長谷川 舞']

# 講師ごとの出勤傾向（昼型 / 夜型 / 全日型）
STYLE = {
    '青木 涼': 'all', '伊藤 芽衣': 'day', '大西 拓真': 'night', '川口 千尋': 'all',
    '佐々木 蓮': 'night', '田村 結衣': 'day', '中島 悠斗': 'all', '長谷川 舞': 'night',
}
STYLE_P = {
    'day':   [0.95, 0.95, 0.75, 0.20, 0.05],
    'night': [0.05, 0.25, 0.85, 0.95, 0.90],
    'all':   [0.70, 0.80, 0.85, 0.85, 0.70],
}

with open(os.path.join(HERE, 'teachers.csv'), 'w', newline='', encoding='utf-8-sig') as f:
    w = csv.writer(f)
    w.writerow(['先生', '日付'] + SLOTS)
    for t in TEACHERS:
        p = STYLE_P[STYLE[t]]
        for d in dates:
            row = [t, f'{d.month}/{d.day}']
            off = (random.random() < 0.12)     # 12%の確率で終日休み
            for i in range(5):
                row.append('' if off else ('○' if random.random() < p[i] else ''))
            w.writerow(row)

# ---------- 生徒（架空） ----------
STUDENTS = [
    ('石田 陽菜', '中3'), ('上原 大翔', '中3'), ('岡本 結菜', '中2'), ('片桐 悠真', '中1'),
    ('木下 莉子', '小6'), ('小林 湊',   '小5'), ('斉藤 詩織', '高2'), ('清水 蒼空', '高1'),
    ('鈴木 心春', '中3'), ('高木 陽向', '中2'), ('土屋 美月', '小6'), ('西村 颯太', '中1'),
    ('橋本 咲良', '高3'), ('平野 樹',   '小4'),
]
SUBJ_BY_GRADE = {
    '小4': ['算数', '国語'], '小5': ['算数', '国語'], '小6': ['算数', '国語', '英語'],
    '中1': ['数学', '英語', '国語'], '中2': ['数学', '英語', '理科'],
    '中3': ['数学', '英語', '国語', '理科', '社会'],
    '高1': ['数学', '英語'], '高2': ['数学', '英語'], '高3': ['英語', '国語'],
}
NG_POOL = ['', '', '13:00', '18:30以降', '20:00', '月', '土',
           '火17:00以降', '水18:30以降', '13:00;土', '月;木18:30以降', '8/1']

rows = []
for name, grade in STUDENTS:
    pool = SUBJ_BY_GRADE[grade]
    n_sub = random.choice([1, 2, 2, 3]) if len(pool) >= 3 else random.choice([1, 2])
    subs = random.sample(pool, min(n_sub, len(pool)))
    ng = random.choice(NG_POOL)
    cap = random.choice([2, 2, 2, 3])
    for i, s in enumerate(subs):
        cand = random.sample(TEACHERS, random.choice([1, 2, 2, 3]))
        rows.append({
            '生徒名': name, '学年': grade, '科目': s,
            '回数': random.choice([4, 5, 6, 6, 8]),
            '担当(優先順;区切り)': ';'.join(cand),
            '配分(順/分散)': random.choice(['順', '順', '分散']),
            '1日上限': cap,
            'NG': ng if i == 0 else '',          # NGは生徒の最初の行にだけ書く
        })

cols = ['生徒名', '学年', '科目', '回数', '担当(優先順;区切り)', '配分(順/分散)', '1日上限', 'NG']
with open(os.path.join(HERE, 'students.csv'), 'w', newline='', encoding='utf-8-sig') as f:
    w = csv.DictWriter(f, fieldnames=cols)
    w.writeheader()
    w.writerows(rows)

total = sum(r['回数'] for r in rows)
print(f'稼働日 {len(dates)}日 / 講師 {len(TEACHERS)}名 / 生徒 {len(STUDENTS)}名 '
      f'/ 受講{len(rows)}件・計{total}コマ を生成しました。')
