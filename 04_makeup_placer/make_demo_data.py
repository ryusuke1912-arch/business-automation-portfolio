# -*- coding: utf-8 -*-
"""
振替配置ツール用のサンプル入力を生成する。

生成物 : slots.csv / shifts.csv / lessons.csv / rules.csv / makeups.csv / config.csv
氏名・出勤・受講データはすべて動作確認用の架空データ（乱数シード固定で再現可能）。
"""
import csv, os, random, datetime

HERE = os.path.dirname(os.path.abspath(__file__))
random.seed(20261003)

TEACHERS = ['青木 涼', '井坂 千尋', '梅田 拓真', '遠藤 美咲',
            '大西 悠人', '柏木 奈緒', '桐山 一馬', '倉田 真理']
# 講師ごとの担当科目（振替の科目マッチングに使う）
SPECIAL = {'青木 涼': ['数学', '理科'], '井坂 千尋': ['英語', '国語'],
           '梅田 拓真': ['数学', '英語'], '遠藤 美咲': ['国語', '社会'],
           '大西 悠人': ['理科', '数学'], '柏木 奈緒': ['英語', '社会'],
           '桐山 一馬': ['数学', '国語'], '倉田 真理': ['英語', '理科']}

STUDENTS = [
    ('小4', '相原 陽菜'), ('小5', '飯塚 蒼太'), ('小6', '上原 結衣'), ('小6', '江川 悠真'),
    ('中1', '奥村 一花'), ('中1', '加瀬 大和'), ('中2', '菊池 芽依'), ('中2', 'глебов 無し'),
    ('中2', '國分 湊'), ('中3', '小泉 咲良'), ('中3', '斉藤 陸'), ('中3', '篠原 心春'),
    ('高1', '瀬川 律'), ('高2', '田渕 美月'), ('高2', '津川 悠斗'), ('高3', '寺田 彩花'),
]
STUDENTS[7] = ('中2', '木暮 詩音')          # 名簿の体裁を整える
SUBJECTS = ['国語', '数学', '英語', '理科', '社会']

SLOTS = [
    ('平日', '17:00-18:20', '17:00', '18:20', '通常'),
    ('平日', '18:30-19:50', '18:30', '19:50', '通常'),
    ('平日', '20:00-21:20', '20:00', '21:20', '通常'),
    ('平日', '19:00-19:50', '19:00', '19:50', '映像'),
    ('平日', '20:00-20:50', '20:00', '20:50', '映像'),
    ('土',   '13:00-14:20', '13:00', '14:20', '通常'),
    ('土',   '14:30-15:50', '14:30', '15:50', '通常'),
    ('土',   '17:00-18:20', '17:00', '18:20', '通常'),
    ('土',   '18:30-19:50', '18:30', '19:50', '通常'),
    ('土',   '13:00-13:50', '13:00', '13:50', '映像'),
    ('土',   '14:00-14:50', '14:00', '14:50', '映像'),
]
WDN = '月火水木金土日'


def build_dates():
    """10月の月・水・金・土を開講日とする（火・木・日は休講）"""
    out = []
    for d in range(1, 32):
        dt = datetime.date(2026, 10, d)
        w = WDN[dt.weekday()]
        if w in '月水金土':
            out.append((dt.strftime('%Y-%m-%d'), w))
    return out[:16]


def norm_ok(avail, ds, slot, t):
    return t in avail.get((ds, slot), [])


def main():
    dates = build_dates()

    with open(os.path.join(HERE, 'slots.csv'), 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.writer(f); w.writerow(['区分', 'コマ', '開始', '終了', '枠種別'])
        w.writerows(SLOTS)

    # ---- 出勤表：講師ごとに出やすい曜日と時間帯の傾向をつける ----
    norm_slots = {'平日': [s[1] for s in SLOTS if s[0] == '平日' and s[4] == '通常'],
                  '土':   [s[1] for s in SLOTS if s[0] == '土' and s[4] == '通常']}
    shifts = []
    for t in TEACHERS:
        night = random.random() < 0.4                 # 夜型講師は17:00に入らない
        rate = random.uniform(0.30, 0.60)
        for ds, w in dates:
            if random.random() > rate:
                continue
            kind = '土' if w == '土' else '平日'
            row = {'日付': ds, '曜日': w, '講師': t}
            for s in norm_slots[kind]:
                if night and s.startswith(('13:', '14:', '17:')):
                    continue
                if random.random() < 0.85:
                    row[s] = '○'
            if any(row.get(s) for s in norm_slots[kind]):
                shifts.append(row)
    cols = ['日付', '曜日', '講師'] + sorted({s[1] for s in SLOTS if s[4] == '通常'})
    with open(os.path.join(HERE, 'shifts.csv'), 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.DictWriter(f, fieldnames=cols); w.writeheader()
        for r in shifts:
            w.writerow({c: r.get(c, '') for c in cols})

    # ---- 既存授業：座席のおよそ半分を埋める（振替を差し込む余地を残す） ----
    avail = {}
    for r in shifts:
        for s in [c for c in cols[3:] if r.get(c)]:
            avail.setdefault((r['日付'], s), []).append(r['講師'])
    lessons, busy = [], set()
    for (ds, s), ts in sorted(avail.items()):
        for t in ts:
            for _ in range(random.choice([1, 1, 2, 2])):
                cand = [st for st in STUDENTS if (st[1], ds, s) not in busy]
                if not cand:
                    break
                g, n = random.choice(cand)
                busy.add((n, ds, s))
                lessons.append([ds, s, t, g, n, random.choice(SPECIAL[t]), ''])

    # ---- 生徒別 振替条件 ----
    rows = []
    for i, (g, n) in enumerate(STUDENTS):
        r = {'生徒名': n, 'OK講師': '', 'OK曜日': '', 'NG曜日': '',
             'OKコマ': '', '最早時刻': '', '最終時刻': '', '週上限日数': '', '備考': ''}
        k = [0,8,2,5,4,8,6,1,3,8,5,7,2,8,4,6][i]
        if k == 0:
            r['OK講師'] = ';'.join(random.sample(TEACHERS, 2)); r['備考'] = '担当を2名に限定'
        elif k == 1:
            r['OK曜日'] = random.choice(['月', '水', '金']); r['備考'] = '通塾曜日が固定'
        elif k == 2:
            r['最終時刻'] = '20:00'; r['備考'] = '20:00より前のみ'
        elif k == 3:
            r['最早時刻'] = '18:30'; r['備考'] = '18:30以降のみ'
        elif k == 4:
            r['NG曜日'] = '土'; r['週上限日数'] = '3'; r['備考'] = '土曜不可・週3日まで'
        elif k == 5:
            r['OKコマ'] = '17:00-18:20'; r['備考'] = 'この1コマのみ'
        elif k == 6:
            r['OK講師'] = random.choice(TEACHERS); r['OK曜日'] = random.choice(['水', '土'])
            r['備考'] = '担当1名・曜日も固定（最も厳しい）'
        elif k == 7:
            r['OK講師'] = ';'.join(random.sample(TEACHERS, 2)); r['最早時刻'] = '18:30'
            r['備考'] = '担当を限定・18:30以降のみ'
        else:
            r['OKコマ'] = '20:00-21:20'; r['OK曜日'] = '月;水'
            r['備考'] = '最終コマのみ・月水限定'
        if n == '相原 陽菜':
            r = {'生徒名': n, 'OK講師': '', 'OK曜日': '水', 'NG曜日': '',
                 'OKコマ': '17:00-18:20', '最早時刻': '', '最終時刻': '',
                 '週上限日数': '', '備考': '水曜のこの1コマのみ（通常授業に続けて実施）'}
        rows.append(r)
    with open(os.path.join(HERE, 'rules.csv'), 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0])); w.writeheader(); w.writerows(rows)

    # ---- 座席共有ケース：水曜17:00に毎回在席している児童を1名つくる ----
    #      「その1コマしか受けられない」条件のため、通常授業に続けて実施するほかない
    SHARE_G, SHARE_N = '小4', '相原 陽菜'
    lessons = [l for l in lessons if l[4] != SHARE_N]
    for ds, w in dates:
        if w != '水':
            continue
        ts = [t for t in TEACHERS if norm_ok(avail, ds, '17:00-18:20', t)]
        if ts:
            lessons.append([ds, '17:00-18:20', ts[0], SHARE_G, SHARE_N, '国語', '40'])

    # ---- 振替対象：既存授業の一部を「欠席」に見立てて取り出す ----
    #      欠席したコマは座席を空けるため、lessons.csv からは除外して書き出す
    share_item = [l for l in lessons if l[4] == SHARE_N][:1]
    picked = share_item + random.sample([l for l in lessons if l[4] != SHARE_N], 41)
    rest = [l for l in lessons if l not in picked]
    with open(os.path.join(HERE, 'lessons.csv'), 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.writer(f); w.writerow(['日付', 'コマ', '講師', '学年', '生徒名', '科目', '所要分'])
        w.writerows(rest)
    with open(os.path.join(HERE, 'makeups.csv'), 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.writer(f)
        w.writerow(['学年', '生徒名', '科目', '欠席日', '欠席コマ', '所要分', '元担当', '期限', '備考'])
        for ds, s, t, g, n, sub, dur in picked:
            dl = '2026-10-20' if random.random() < 0.2 else ''
            w.writerow([g, n, sub, ds, s, dur, t, dl, '模試直しのため期限あり' if dl else ''])

    with open(os.path.join(HERE, 'config.csv'), 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.writer(f); w.writerow(['項目', '値', '説明'])
        w.writerows([
            ['座席定員', '2', '1講師が1コマで同時に見られる人数（1対2）'],
            ['連続とみなす間隔(分)', '15', 'この分数以内に隣接していれば「連続」と判定'],
            ['配置不可講師', '倉田 真理', '振替を入れない講師（; 区切りで複数可）'],
        ])

    print(f'サンプル入力を生成しました（{HERE}）')
    print(f'  開講日 {len(dates)}日 / 講師 {len(TEACHERS)}名 / 生徒 {len(STUDENTS)}名')
    print(f'  既存授業 {len(rest)}コマ / 振替対象 {len(picked)}件')


if __name__ == '__main__':
    main()
