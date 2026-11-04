"""自習室 来室計画表（川内さん用）を A4縦1枚で印刷できる Excel に出力する。"""
import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.page import PageMargins

START, END = dt.date(2026, 11, 5), dt.date(2026, 11, 18)
WEEK = "月火水木金土日"
FONT = "游ゴシック"

thin = Side(style="thin", color="333333")
thick = Side(style="medium", color="000000")
box = Border(left=thin, right=thin, top=thin, bottom=thin)
center = Alignment(horizontal="center", vertical="center", wrap_text=True)
left = Alignment(horizontal="left", vertical="center", wrap_text=True)
head_fill = PatternFill("solid", fgColor="E6E6E6")
sat_fill = PatternFill("solid", fgColor="EAF2FB")
sun_fill = PatternFill("solid", fgColor="FBECEC")

wb = Workbook()
ws = wb.active
ws.title = "自習計画表"

# 列幅（A:日付 B:曜日 C:時間 D:科目・内容 E:○× F:サイン）
for col, w in zip("ABCDEF", [9, 6, 22, 30, 7, 13]):
    ws.column_dimensions[col].width = w

# タイトル
ws.merge_cells("A1:F1")
ws["A1"] = "自習室 来室計画表"
ws["A1"].font = Font(name=FONT, size=20, bold=True)
ws["A1"].alignment = center
ws.row_dimensions[1].height = 40

# 氏名欄・テスト開始
ws.merge_cells("A3:C3")
ws["A3"] = "氏名　川内"
ws["A3"].font = Font(name=FONT, size=14, bold=True)
ws["A3"].alignment = left
for c in "ABC":
    ws[f"{c}3"].border = Border(bottom=thick)
ws.merge_cells("D3:F3")
ws["D3"] = "テスト開始：11月19日（木）"
ws["D3"].font = Font(name=FONT, size=14, bold=True)
ws["D3"].alignment = center
for c in "DEF":
    ws[f"{c}3"].border = Border(left=thick if c == "D" else None, right=thick if c == "F" else None,
                                top=thick, bottom=thick)
ws.row_dimensions[2].height = 8
ws.row_dimensions[3].height = 30

# 説明
ws.merge_cells("A4:F4")
ws["A4"] = ("テストまでの2週間、自習室に「いつ来て、何を勉強するか」を自分で決めて書きましょう。"
            "来られない日は時間欄に「×」と書いてください。")
ws["A4"].font = Font(name=FONT, size=10)
ws["A4"].alignment = left
ws.row_dimensions[4].height = 34

# 表の見出し
headers = ["日付", "曜日", "来る時間 〜 帰る時間", "勉強する科目・内容", "来た\n○×", "講師\n確認サイン"]
for i, h in enumerate(headers, 1):
    cell = ws.cell(row=5, column=i, value=h)
    cell.font = Font(name=FONT, size=10, bold=True)
    cell.alignment = center
    cell.fill = head_fill
    cell.border = box
ws.row_dimensions[5].height = 30

# 日付行
r = 6
d = START
while d <= END:
    wd = WEEK[d.weekday()]
    vals = [f"{d.month}/{d.day}", wd, "　：　 〜 　：　", "", "", ""]
    fill = sat_fill if wd == "土" else sun_fill if wd == "日" else None
    for i, v in enumerate(vals, 1):
        cell = ws.cell(row=r, column=i, value=v)
        cell.border = box
        cell.alignment = left if i == 4 else center
        cell.font = Font(name=FONT, size=12 if i <= 2 else 11, bold=i <= 2,
                         color=("1A4FA0" if wd == "土" else "B22222" if wd == "日" else "000000")
                         if i == 2 else ("666666" if i == 3 else "000000"))
        if fill:
            cell.fill = fill
    ws.row_dimensions[r].height = 36
    r += 1
    d += dt.timedelta(days=1)

# 決意欄
r += 1
ws.row_dimensions[r - 1].height = 10
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
ws.cell(row=r, column=1, value="テストに向けての決意（ひとこと）").font = Font(name=FONT, size=11, bold=True)
ws.cell(row=r, column=1).alignment = Alignment(vertical="top", indent=1)
ws.row_dimensions[r].height = 22
ws.merge_cells(start_row=r + 1, start_column=1, end_row=r + 1, end_column=6)
ws.row_dimensions[r + 1].height = 50
for rr in (r, r + 1):
    for c in range(1, 7):
        ws.cell(row=rr, column=c).border = Border(
            left=thick if c == 1 else None, right=thick if c == 6 else None,
            top=thick if rr == r else None, bottom=thick if rr == r + 1 else None)
last = r + 1

# 印刷設定：A4縦・1ページに収める
ws.print_area = f"A1:F{last}"
ws.page_setup.paperSize = ws.PAPERSIZE_A4
ws.page_setup.orientation = "portrait"
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 1
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.print_options.horizontalCentered = True
ws.page_margins = PageMargins(left=0.5, right=0.5, top=0.6, bottom=0.5, header=0.2, footer=0.2)
ws.oddFooter.right.text = "S1御影教室"
ws.sheet_view.showGridLines = False

wb.save("自習計画表_川内.xlsx")
print("saved", last, "rows")
