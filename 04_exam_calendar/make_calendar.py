"""入試カレンダー（教室掲示用）を Excel で作る。

使い方:
  python make_calendar.py                      → 空欄のカレンダー（手書き用）
  python make_calendar.py 入試日程.csv          → 日付のマスに「学年 名字（学校名）」を自動で書き込む

CSV の列: 日付,学年,名字,学校名,区分   （日付は 2027/1/16 や 1/16 の形。月だけなら年は自動で補う）
各月 A4横1枚で印刷できる。最後のシートは日付順の一覧。
"""
import csv
import sys
import calendar
import datetime as dt
from collections import defaultdict
from pathlib import Path

from openpyxl import Workbook
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.page import PageMargins

# 対象期間（入試シーズン）
MONTHS = [(2026, 12), (2027, 1), (2027, 2), (2027, 3)]
# 祝日（2026年12月〜2027年3月）
HOLIDAYS = {
    dt.date(2027, 1, 1): '元日', dt.date(2027, 1, 11): '成人の日',
    dt.date(2027, 2, 11): '建国記念の日', dt.date(2027, 2, 23): '天皇誕生日',
    dt.date(2027, 3, 21): '春分の日', dt.date(2027, 3, 22): '振替休日',
}
GRADE_ORDER = {'小6': 0, '中3': 1, '高3': 2}
FONT = 'ＭＳ ゴシック'
thin = Side(style='thin', color='000000')
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
GRAY = PatternFill('solid', fgColor='D9D9D9')
LIGHT = PatternFill('solid', fgColor='F2F2F2')
RED = 'C00000'
BLUE = '1F4E99'


def norm_grade(g):
    g = g.strip().translate(str.maketrans('０１２３４５６７８９', '0123456789'))
    return g.replace('年', '').replace('中学', '中').replace('高校', '高').replace('小学', '小')


def parse_date(s):
    s = s.strip().translate(str.maketrans('０１２３４５６７８９／', '0123456789/')).replace('-', '/')
    parts = [int(p) for p in s.split('/') if p]
    if len(parts) == 3:
        return dt.date(*parts)
    m, d = parts
    return dt.date(2026 if m >= 4 else 2027, m, d)   # 年なし：4月以降は2026年、1〜3月は2027年


def load(path):
    rows = []
    with open(path, encoding='utf-8-sig') as f:
        for r in csv.DictReader(f):
            if not (r.get('日付') or '').strip():
                continue
            rows.append({
                'date': parse_date(r['日付']), 'grade': norm_grade(r.get('学年', '')),
                'name': r.get('名字', '').strip(), 'school': r.get('学校名', '').strip(),
                'kind': r.get('区分', '').strip(),
            })
    rows.sort(key=lambda r: (r['date'], GRADE_ORDER.get(r['grade'], 9), r['name']))
    return rows


def month_sheet(wb, y, m, by_day):
    ws = wb.create_sheet(f'{y}年{m}月')
    ws.sheet_view.showGridLines = False
    for c in 'ABCDEFG':
        ws.column_dimensions[c].width = 26
    ws.merge_cells('A1:G1')
    ws['A1'] = f'入試カレンダー　{y}年{m}月'
    ws['A1'].font = Font(name=FONT, size=20, bold=True)
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 34
    ws.merge_cells('A2:G2')
    ws['A2'] = '表記：学年 名字：学校名 区分　／　S1御影教室'
    ws['A2'].font = Font(name=FONT, size=10)
    ws['A2'].alignment = Alignment(horizontal='right')

    for i, w in enumerate('日月火水木金土'):
        c = ws.cell(row=3, column=i + 1, value=w)
        c.font = Font(name=FONT, size=12, bold=True,
                      color=RED if i == 0 else BLUE if i == 6 else '000000')
        c.fill, c.border = GRAY, BOX
        c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[3].height = 20

    weeks = calendar.Calendar(firstweekday=6).monthdatescalendar(y, m)
    row_h = 590 / len(weeks)          # A4横1枚に収まるよう週の数で高さを決める
    for wi, week in enumerate(weeks):
        r = 4 + wi
        ws.row_dimensions[r].height = row_h
        for di, d in enumerate(week):
            c = ws.cell(row=r, column=di + 1)
            c.border = BOX
            c.alignment = Alignment(vertical='top', wrap_text=True)
            if d.month != m:
                c.fill = LIGHT
                continue
            head = f'{d.day}'
            if d in HOLIDAYS:
                head += f'　{HOLIDAYS[d]}'
            lines = [head]
            for e in by_day.get(d, []):
                detail = ' '.join(x for x in (e['school'], e['kind']) if x)
                lines.append(f"{e['grade']} {e['name']}" + (f'：{detail}' if detail else ''))
            # 日付（と祝日名）だけ色を付け、名前は黒で書く
            red = di == 0 or d in HOLIDAYS
            size = 11 if len(lines) <= 5 else 9 if len(lines) <= 7 else 8
            day_font = InlineFont(rFont=FONT, sz=size + 2, b=True,
                                  color=RED if red else BLUE if di == 6 else '000000')
            parts = [TextBlock(day_font, head)]
            if len(lines) > 1:
                parts.append(TextBlock(InlineFont(rFont=FONT, sz=size), '\n' + '\n'.join(lines[1:])))
            c.value = CellRichText(parts)
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(left=0.4, right=0.4, top=0.4, bottom=0.4)
    ws.print_options.horizontalCentered = True


def list_sheet(wb, rows):
    ws = wb.create_sheet('一覧')
    head = ['日付', '曜日', '学年', '名字', '学校名', '区分']
    widths = [14, 6, 7, 14, 30, 14]
    for i, (h, w) in enumerate(zip(head, widths)):
        c = ws.cell(row=1, column=i + 1, value=h)
        c.font, c.fill, c.border = Font(name=FONT, bold=True), GRAY, BOX
        c.alignment = Alignment(horizontal='center')
        ws.column_dimensions[chr(65 + i)].width = w
    for ri, e in enumerate(rows, start=2):
        vals = [e['date'].strftime('%Y/%m/%d'), '月火水木金土日'[e['date'].weekday()],
                e['grade'], e['name'], e['school'], e['kind']]
        for ci, v in enumerate(vals):
            c = ws.cell(row=ri, column=ci + 1, value=v)
            c.font, c.border = Font(name=FONT), BOX
            if ci in (0, 1, 2):
                c.alignment = Alignment(horizontal='center')
    ws.freeze_panes = 'A2'
    ws.print_title_rows = '1:1'
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True


def main():
    src = sys.argv[1] if len(sys.argv) > 1 else None
    here = Path(__file__).parent
    out = Path(sys.argv[2]) if len(sys.argv) > 2 else here / ('入試カレンダー.xlsx' if src else '入試カレンダー_空欄.xlsx')
    rows = load(src) if src else []
    by_day = defaultdict(list)
    for e in rows:
        by_day[e['date']].append(e)
    wb = Workbook()
    wb.remove(wb.active)
    for y, m in MONTHS:
        month_sheet(wb, y, m, by_day)
    if rows:
        list_sheet(wb, rows)
        outside = [e for e in rows if (e['date'].year, e['date'].month) not in MONTHS]
        if outside:
            print(f'※ カレンダー期間外の日程が {len(outside)} 件あります（一覧シートには載っています）')
    wb.save(out)
    print(f'作成: {out}（{len(rows)}件）')


if __name__ == '__main__':
    main()
