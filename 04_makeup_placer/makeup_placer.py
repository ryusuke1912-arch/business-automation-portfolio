# -*- coding: utf-8 -*-
"""
振替授業 自動配置エンジン（ポートフォリオ公開版）

入力 : slots.csv / shifts.csv / lessons.csv / rules.csv / makeups.csv / config.csv
出力 : makeup_out.xlsx（サマリ＋日付別の座席表）
実行 : python makeup_placer.py     ※Excel本体は不要（openpyxlのみ）

■ 解いている問題
  「すでに埋まっている時間割」を壊さずに、欠席分の振替を後から差し込む。
  新規に組むのではなく“既存の隙間に入れる”ため、制約は次のように重なる。

    (1) 座席定員   … 1講師1コマにつき2名まで（既存授業が先に埋めている）
    (2) 講師の出勤 … 日付×コマ単位。出勤していない枠には入れられない
    (3) 生徒別条件 … 担当できる講師・曜日・時間帯・「このコマだけ」等が生徒ごとに違う
    (4) 時間の重複 … 80分枠と50分枠が混在し、区間として重なる組み合わせがある
    (5) 連続性     … 同じ日に授業がある場合、間を空けずに前後へ付ける
    (6) 週の通塾日 … 生徒によっては「週◯日まで」の上限がある

■ 設計上の要点
  1) 区間ベースの重複判定
     コマ名の一致ではなく [開始, 終了) の重なりで判定する。
     18:30-19:50 と 19:00-19:50（映像）のような包含関係を取りこぼさないため。

  2) MRV法による処理順の決定
     「置ける枠が物理的に何通りあるか」を数え、少ない順に処理する。
     制約の厳しい生徒を先に通し、後半で詰まって破綻するのを防ぐ。

  3) 段階的緩和（本ツールの中心）
     全制約を満たす配置ができなかった件を、条件を1つずつ緩めて再試行する。
       第1段 … 全制約
       第2段 … 連続性を外す（単独コマを許容）
       第3段 … 指定講師を外す（科目相性の高い代替講師へ）
       第4段 … 同一コマ内の座席共有（短時間授業を通常授業に続けて実施）
     重要なのは「入った／入らない」ではなく、**どの条件を譲ったのかを1件ずつ記録して返す**こと。
     黙って条件を破ると現場で事故になるため、緩和内容は出力の全経路（Excel・標準出力）に残す。

  4) 独立検算
     配置後、エンジンとは別経路で全制約を検査し直す（verify）。
     スコアやフラグではなく実データを読み直して照合するため、ロジックの取り違えを検出できる。
"""
import csv, os, sys, unicodedata, re, datetime, collections
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, 'makeup_out.xlsx')
WDN = '月火水木金土日'


def norm(s):
    if s is None:
        return ''
    s = unicodedata.normalize('NFKC', str(s)).replace('　', ' ')
    return re.sub(r'\s+', '', s).strip()


def mins(hhmm):
    h, m = hhmm.split(':')
    return int(h) * 60 + int(m)


def read(name):
    path = os.path.join(HERE, name)
    if not os.path.exists(path):
        sys.exit(f'入力ファイルが見つかりません: {path}\n  先に `python make_demo_data.py` を実行してください。')
    with open(path, encoding='utf-8-sig') as f:
        return list(csv.DictReader(f))


# ============================================================ 入力
class Data:
    def __init__(self):
        cfg = {r['項目']: r['値'] for r in read('config.csv')}
        self.seats = int(cfg.get('座席定員', 2))
        self.gap = int(cfg.get('連続とみなす間隔(分)', 15))
        self.banned = {norm(x) for x in cfg.get('配置不可講師', '').split(';') if x.strip()}

        # コマ定義：区分（平日／土）ごとに [(コマ名, 開始, 終了, 種別)]
        self.slots = collections.defaultdict(list)
        self.span, self.length = {}, {}
        for r in read('slots.csv'):
            self.slots[r['区分']].append((r['コマ'], r['枠種別']))
            self.span[r['コマ']] = (mins(r['開始']), mins(r['終了']))
            self.length[r['コマ']] = mins(r['終了']) - mins(r['開始'])

        # 出勤：avail[(日付, コマ)] = {講師}
        self.avail = collections.defaultdict(set)
        self.tdisp = {}                      # 正規化名 -> 元の表記（表示用）
        self.wd, self.dates = {}, []
        for r in read('shifts.csv'):
            d = r['日付']
            self.wd[d] = r['曜日']
            if d not in self.dates:
                self.dates.append(d)
            self.tdisp.setdefault(norm(r['講師']), r['講師'].strip())
            for k, v in r.items():
                if k in ('日付', '曜日', '講師'):
                    continue
                if str(v).strip():
                    self.avail[(d, k)].add(norm(r['講師']))
        self.dates.sort()
        self.teachers = sorted({t for s in self.avail.values() for t in s})

        # 既存授業
        self.lessons = read('lessons.csv')
        self.seat = collections.Counter()          # (日付, コマ, 講師) -> 使用席数
        self.booked = collections.defaultdict(list)  # (日付, 生徒) -> [コマ]
        self.days = collections.defaultdict(set)     # 生徒 -> {日付}
        self.aff = collections.Counter()             # (講師, 科目) -> 担当実績
        self.used_min = collections.Counter()      # (日付, コマ, 生徒) -> 使用済み分数
        for r in self.lessons:
            t, n = norm(r['講師']), norm(r['生徒名'])
            self.tdisp.setdefault(t, r['講師'].strip())
            self.used_min[(r['日付'], r['コマ'], n)] += int(r['所要分'] or self.length[r['コマ']])
            self.seat[(r['日付'], r['コマ'], t)] += 1
            self.booked[(r['日付'], n)].append(r['コマ'])
            self.days[n].add(r['日付'])
            self.aff[(t, r['科目'])] += 1

        # 生徒別条件
        self.rules = {}
        for r in read('rules.csv'):
            self.rules[norm(r['生徒名'])] = {
                'ok_t': {norm(x) for x in r['OK講師'].split(';') if x.strip()},
                'ok_w': {x.strip() for x in r['OK曜日'].split(';') if x.strip()},
                'ng_w': {x.strip() for x in r['NG曜日'].split(';') if x.strip()},
                'ok_s': {x.strip() for x in r['OKコマ'].split(';') if x.strip()},
                'from': mins(r['最早時刻']) if r['最早時刻'].strip() else None,
                'to':   mins(r['最終時刻']) if r['最終時刻'].strip() else None,
                'wcap': int(r['週上限日数']) if r['週上限日数'].strip() else None,
            }

        # 振替対象
        self.items = []
        for i, r in enumerate(read('makeups.csv')):
            self.items.append({
                'no': i + 1, 'grade': r['学年'], 'disp': r['生徒名'].strip(),
                'name': norm(r['生徒名']), 'subject': r['科目'],
                'absent': r['欠席日'], 'abs_slot': r['欠席コマ'],
                'orig': norm(r['元担当']), 'deadline': r['期限'].strip(),
                'dur': int(r['所要分']) if str(r.get('所要分', '')).strip() else None,
                'memo': r.get('備考', ''), 'date': None, 'slot': None,
                'teacher': None, 'relax': [],
            })

    # ---------- 判定 ----------
    def tn(self, t):
        return self.tdisp.get(t, t)

    def kind(self, d):
        return '土' if self.wd[d] == '土' else '平日'

    def slots_of(self, d):
        return [s for s, _ in self.slots[self.kind(d)]]

    def overlap(self, a, b):
        (a1, b1), (a2, b2) = self.span[a], self.span[b]
        return a1 < b2 and a2 < b1

    def conflict(self, d, stu, slot):
        """同じ生徒が同時刻に重ならないか（区間で判定）"""
        return any(self.overlap(slot, s) for s in self.booked[(d, stu)])

    def adjacent(self, d, stu, slot):
        """当日の他の授業と連続しているか。当日ゼロなら穴は生じないのでTrue"""
        cur = self.booked[(d, stu)]
        if not cur:
            return True
        a, b = self.span[slot]
        return any(abs(a - self.span[s][1]) <= self.gap or
                   abs(self.span[s][0] - b) <= self.gap for s in cur)

    def rule_ok(self, stu, d, slot):
        r = self.rules.get(stu)
        if not r:
            return True
        w = self.wd[d]
        if r['ok_w'] and w not in r['ok_w']:
            return False
        if w in r['ng_w']:
            return False
        if r['ok_s'] and slot not in r['ok_s']:
            return False
        st = self.span[slot][0]
        if r['to'] is not None and st >= r['to']:
            return False
        if r['from'] is not None and st < r['from']:
            return False
        return True

    def week_days(self, stu, d):
        wk = datetime.date(*map(int, d.split('-'))).isocalendar()[1]
        return sum(1 for x in self.days[stu]
                   if datetime.date(*map(int, x.split('-'))).isocalendar()[1] == wk)


# ============================================================ 配置
def candidates(D, it, d, slot, need_adj=True, need_teacher=True):
    stu = it['name']
    if it['deadline'] and d > it['deadline']:
        return []
    if not D.rule_ok(stu, d, slot):
        return []
    if D.conflict(d, stu, slot):
        return []
    if need_adj and not D.adjacent(d, stu, slot):
        return []
    pool = set(D.avail[(d, slot)]) - D.banned
    if need_teacher:
        ok = D.rules.get(stu, {}).get('ok_t')
        if ok:
            pool &= ok
    return sorted(t for t in pool if D.seat[(d, slot, t)] < D.seats)


def score(D, it, d, slot, t):
    stu, sc = it['name'], 0.0
    if t == it['orig']:
        sc += 40                                   # 元の担当講師を継続できる
    sc += min(D.aff.get((t, it['subject']), 0), 12) * 3   # 科目の担当実績
    if D.booked[(d, stu)] and D.adjacent(d, stu, slot):
        sc += 25                                   # 既存授業に連結できる
    sc -= D.seat[(d, slot, t)] * 4                 # 空いている講師へ散らす
    wd_ = D.week_days(stu, d)
    sc -= wd_ * 3                                  # 特定の週に集中させない
    cap = D.rules.get(stu, {}).get('wcap')
    if cap and d not in D.days[stu] and wd_ >= cap:
        sc -= 1000                                 # 週の通塾日上限を超える新規日
    sc -= D.dates.index(d) * 0.3                   # 早い日程を優先
    return sc


def freedom(D, it):
    return sum(len(candidates(D, it, d, s)) for d in D.dates for s in D.slots_of(d))


def commit(D, it, d, slot, t, relax, shared=False):
    if not shared:
        D.seat[(d, slot, t)] += 1
    D.booked[(d, it['name'])].append(slot)
    D.days[it['name']].add(d)
    D.aff[(t, it['subject'])] += 1
    it.update(date=d, slot=slot, teacher=t, relax=relax, shared=shared)


def try_place(D, it, need_adj, need_teacher):
    best = None
    for d in D.dates:
        for s in D.slots_of(d):
            for t in candidates(D, it, d, s, need_adj, need_teacher):
                v = score(D, it, d, s, t)
                if best is None or v > best[0]:
                    best = (v, d, s, t)
    if best is None:
        return False
    _, d, s, t = best
    relax = []
    if not D.adjacent(d, it['name'], s):
        others = [x for x in D.booked[(d, it['name'])]]
        relax.append(f'同日の授業（{sorted(others, key=lambda x: D.span[x])[0]}）とは離れたコマ'
                     if others else '当日は他の授業がなく単独コマ')
    ok = D.rules.get(it['name'], {}).get('ok_t')
    if ok and t not in ok:
        relax.append(f'指定講師（{"・".join(D.tn(x) for x in sorted(ok))}）が空かず代替講師へ')
    commit(D, it, d, s, t, relax)
    return True


def try_share(D, it):
    """短時間授業を、本人が在席している同一コマの後半に接続する（座席を新たに使わない）

    許可条件：振替が短時間授業であり、その日そのコマで本人が使っている分数と
    合計してもコマの長さに収まること。80分授業を二重取りすることはできない。
    """
    if not it['dur']:
        return False
    for d in D.dates:
        if it['deadline'] and d > it['deadline']:
            continue
        if not D.rules.get(it['name'], {}).get('ok_s'):
            continue
        for s in D.booked[(d, it['name'])]:
            if not D.rule_ok(it['name'], d, s):
                continue
            if D.used_min[(d, s, it['name'])] + it['dur'] > D.length[s]:
                continue
            for r in D.lessons:
                if r['日付'] == d and r['コマ'] == s and norm(r['生徒名']) == it['name']:
                    used = D.used_min[(d, s, it['name'])]
                    st = mins(f'{D.span[s][0]//60:02d}:{D.span[s][0]%60:02d}') + used
                    commit(D, it, d, s, norm(r['講師']),
                           [f'通常授業（{s}）に続けて同一講師で実施'
                            f'／{st//60:02d}:{st%60:02d}開始・{it["dur"]}分'], shared=True)
                    D.used_min[(d, s, it['name'])] += it['dur']
                    return True
    return False


def solve(D):
    order = sorted(D.items, key=lambda x: freedom(D, x))
    stages = [('全制約', True, True), ('連続性を緩和', False, True), ('指定講師を緩和', False, False)]
    rest, log = order, []
    for label, adj, tea in stages:
        nxt = [it for it in rest if not try_place(D, it, adj, tea)]
        log.append((label, len(rest) - len(nxt), len(nxt)))
        rest = nxt
    nxt = [it for it in rest if not try_share(D, it)]
    log.append(('同一コマ内で接続', len(rest) - len(nxt), len(nxt)))
    return log, nxt


# ============================================================ 検算
def verify(D, unplaced):
    """エンジンとは別経路で全制約を読み直して照合する"""
    err = []
    placed = [it for it in D.items if it['date']]

    seat = collections.Counter()
    for r in D.lessons:
        seat[(r['日付'], r['コマ'], norm(r['講師']))] += 1
    for it in placed:
        if not it.get('shared'):
            seat[(it['date'], it['slot'], it['teacher'])] += 1
    for k, v in seat.items():
        if v > D.seats:
            err.append(f'座席超過 {k} = {v}名（定員{D.seats}）')

    book = collections.defaultdict(list)
    for r in D.lessons:
        book[(r['日付'], norm(r['生徒名']))].append((r['コマ'], '既存'))
    for it in placed:
        book[(it['date'], it['name'])].append((it['slot'], '振替'))
    for (d, stu), lst in book.items():
        for i in range(len(lst)):
            for j in range(i + 1, len(lst)):
                if lst[i][1] == lst[j][1] == '既存':
                    continue
                if any(x['shared'] for x in placed if x['name'] == stu and x['date'] == d):
                    continue
                if D.overlap(lst[i][0], lst[j][0]):
                    err.append(f'時間重複 {d} {stu} {lst[i]} × {lst[j]}')

    for it in placed:
        if it.get('shared'):
            continue
        if it['teacher'] not in D.avail[(it['date'], it['slot'])]:
            err.append(f'出勤外 {it["date"]} {it["slot"]} {D.tn(it["teacher"])}（{it["disp"]}）')
        if it['teacher'] in D.banned:
            err.append(f'配置不可講師に配置 {it["disp"]} → {D.tn(it["teacher"])}')
        if not D.rule_ok(it['name'], it['date'], it['slot']):
            err.append(f'生徒条件違反 {it["disp"]} {it["date"]} {it["slot"]}')
        if it['deadline'] and it['date'] > it['deadline']:
            err.append(f'期限超過 {it["disp"]} 期限{it["deadline"]} → {it["date"]}')

    if len(placed) + len(unplaced) != len(D.items):
        err.append(f'件数不一致 {len(placed)}+{len(unplaced)} ≠ {len(D.items)}')
    return err


# ============================================================ 出力
C_HEAD = PatternFill('solid', fgColor='DDEBF7')
C_NAME = PatternFill('solid', fgColor='F2F2F2')
C_OK   = PatternFill('solid', fgColor='C6EFCE')   # 出勤あり
C_NG   = PatternFill('solid', fgColor='FFC7CE')   # 出勤なし
C_MK   = PatternFill('solid', fgColor='FFFF00')   # 今回入れた振替
C_WARN = PatternFill('solid', fgColor='FFE0B2')   # 条件を緩めた件
TH = Side(style='thin', color='B0B0B0')
BD = Border(left=TH, right=TH, top=TH, bottom=TH)
F_T = Font(name='Meiryo', size=11, bold=True)
F_H = Font(name='Meiryo', size=10, bold=True)
F_B = Font(name='Meiryo', size=10)
F_S = Font(name='Meiryo', size=9)
CEN = Alignment(horizontal='center', vertical='center', wrap_text=True)
LFT = Alignment(horizontal='left', vertical='center')


def write_excel(D, kpi, log, unplaced):
    wb = Workbook(); ws = wb.active; ws.title = 'サマリ'
    ws['A1'] = '振替配置 実行結果'; ws['A1'].font = F_T
    r = 3
    for k, v in kpi.items():
        ws.cell(r, 1, k).font = F_B
        ws.cell(r, 2, v).font = F_H
        r += 1
    r += 1
    ws.cell(r, 1, '段階別の配置件数').font = F_H; r += 1
    for label, got, left in log:
        ws.cell(r, 1, label).font = F_B
        ws.cell(r, 2, f'{got}件 配置 / 残 {left}件').font = F_B
        r += 1
    r += 1

    ws.cell(r, 1, '配置一覧（オレンジ＝条件を緩めた件）').font = F_H; r += 1
    hdr = ['No', '学年', '生徒名', '科目', '欠席日', '欠席コマ',
           '振替日', '曜日', '振替コマ', '担当', '元担当', '緩和した条件']
    for i, h in enumerate(hdr, 1):
        c = ws.cell(r, i, h); c.font = F_H; c.fill = C_HEAD; c.alignment = CEN; c.border = BD
    r += 1
    for it in sorted(D.items, key=lambda x: (x['date'] or '9999', x['slot'] or '')):
        vals = [it['no'], it['grade'], it['disp'], it['subject'], it['absent'], it['abs_slot'],
                it['date'] or '―', D.wd.get(it['date'], '―'), it['slot'] or '―',
                D.tn(it['teacher']) if it['teacher'] else '―',
                D.tn(it['orig']), '／'.join(it['relax'])]
        for i, v in enumerate(vals, 1):
            c = ws.cell(r, i, v); c.font = F_B; c.border = BD
            c.alignment = LFT if i in (3, 4, 12) else CEN
            if it['relax']:
                c.fill = C_WARN
            elif it['date'] and i in (7, 8, 9, 10):
                c.fill = C_MK
        r += 1
    for col, w in zip('ABCDEFGHIJKL', (5, 6, 14, 8, 12, 13, 12, 6, 13, 14, 14, 44)):
        ws.column_dimensions[col].width = w

    # ---- 日付別の座席表 ----
    placed = collections.defaultdict(list)
    for x in D.lessons:
        placed[(x['日付'], x['コマ'], norm(x['講師']))].append(
            [x['学年'], x['生徒名'], x['科目'], False])
    for it in D.items:
        if not it['date']:
            continue
        key = (it['date'], it['slot'], it['teacher'])
        if it.get('shared'):
            # 同一コマ内での継続。席は増やさず、既存授業のセルに併記する
            for row in placed[key]:
                if norm(row[1]) == it['name']:
                    row[2] = f'{row[2]}＋{it["subject"]}(振替{it["dur"]}分)'
                    row[3] = True
                    break
        else:
            placed[key].append([it['grade'], it['disp'], it['subject'], True])

    overflow = []
    for d in D.dates:
        sh = wb.create_sheet(f'{d[5:]}({D.wd[d]})')
        blocks = D.slots_of(d)
        sh.cell(1, 1, f'{d}（{D.wd[d]}）').font = F_T
        sh.cell(1, 1).fill = C_HEAD
        for i, b in enumerate(blocks):
            c0 = 2 + i * 3
            sh.merge_cells(start_row=1, start_column=c0, end_row=1, end_column=c0 + 2)
            c = sh.cell(1, c0, b); c.font = F_H; c.fill = C_HEAD; c.alignment = CEN; c.border = BD
        for j, t in enumerate(D.teachers):
            r0 = 2 + j * D.seats
            sh.merge_cells(start_row=r0, start_column=1, end_row=r0 + D.seats - 1, end_column=1)
            c = sh.cell(r0, 1, D.tn(t)); c.font = F_B; c.fill = C_NAME; c.alignment = CEN; c.border = BD
            for k in range(D.seats):
                sh.cell(r0 + k, 1).border = BD
                sh.cell(r0 + k, 1).fill = C_NAME
            for i, b in enumerate(blocks):
                c0 = 2 + i * 3
                lst = placed.get((d, b, t), [])
                if len(lst) > D.seats:
                    overflow.append(f'{d} {b} {D.tn(t)} に {len(lst)}件（定員{D.seats}）')
                on = t in D.avail[(d, b)]
                for k in range(D.seats):
                    rec = lst[k] if k < len(lst) else None
                    fill = C_MK if (rec and rec[3]) else (C_OK if on else C_NG)
                    vals = [rec[0], rec[1], rec[2]] if rec else ['', '', '']
                    for m in range(3):
                        c = sh.cell(r0 + k, c0 + m, vals[m] or None)
                        c.font = F_B; c.alignment = CEN; c.border = BD; c.fill = fill
        rl = 2 + len(D.teachers) * D.seats + 1
        sh.cell(rl, 1, '凡例：　緑＝講師の出勤あり（配置可）　／　赤＝出勤なし（配置不可）　／　黄＝今回入れた振替'
                       '　　※各ブロックは「学年・生徒名・科目」の3列。').font = F_S
        sh.column_dimensions['A'].width = 16
        for i in range(len(blocks)):
            c0 = 2 + i * 3
            for w, m in zip((6, 13, 8), range(3)):
                sh.column_dimensions[get_column_letter(c0 + m)].width = w
        sh.freeze_panes = 'B2'

    wb.save(OUT)
    return overflow


# ============================================================ 実行
def main():
    D = Data()
    log, unplaced = solve(D)
    err = verify(D, unplaced)

    placed = [it for it in D.items if it['date']]
    relaxed = [it for it in placed if it['relax']]
    open_seats = sum(D.seats for d in D.dates for s in D.slots_of(d)
                     for t in D.avail[(d, s)] if t not in D.banned)
    used = sum(D.seat.values())
    kpi = {
        '振替対象件数': len(D.items),
        '配置できた件数': len(placed),
        '充足率': f'{len(placed)/len(D.items)*100:.1f}%',
        '全制約を満たせた件数': len(placed) - len(relaxed),
        '条件を緩めた件数': len(relaxed),
        '未配置': len(unplaced),
        '座席稼働率（振替後）': f'{used/open_seats*100:.1f}%' if open_seats else '-',
        '検算エラー': len(err),
    }
    over = write_excel(D, kpi, log, unplaced)
    if over:
        err += [f'座席表の描画あふれ {x}' for x in over]
        kpi['検算エラー'] = len(err)

    print('=' * 60)
    for k, v in kpi.items():
        print(f'  {k:<24}: {v}')
    print('=' * 60)
    for label, got, left in log:
        print(f'  {label:<16}: {got:>3}件 配置 / 残 {left}件')
    print('-' * 60)
    if relaxed:
        print('△ 条件を緩めて配置した件（運用側の確認が必要）:')
        for it in relaxed:
            print(f'   {it["grade"]} {it["disp"]} {it["subject"]}'
                  f' → {it["date"]}({D.wd[it["date"]]}) {it["slot"]} {D.tn(it["teacher"])}')
            for m in it['relax']:
                print(f'      ※ {m}')
    if unplaced:
        print('★ 物理的に入りきらなかった件（条件を緩めるか出勤を増やす必要あり）:')
        for it in unplaced:
            print(f'   {it["grade"]} {it["disp"]} {it["subject"]}（欠席 {it["absent"]} {it["abs_slot"]}）')
    if not relaxed and not unplaced:
        print('★ 全件を全制約のまま配置しました')
    print('-' * 60)
    if err:
        print(f'検算エラー {len(err)}件:')
        for e in err:
            print('   [NG]', e)
    else:
        print('検算：全チェック通過（座席定員・時間重複・出勤・生徒条件・期限）')
    print(f'\n出力: {OUT}')
    return 1 if err else 0


if __name__ == '__main__':
    sys.exit(main())
