# -*- coding: utf-8 -*-
"""
build_template — assets/template.xlsx(ハウススタイル見本ブック)を再生成する。
スタイルを改良したらこのファイルを編集して再実行(テンプレの単一ソース)。

    python build_template.py
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
from corp_xlsx import *

wb = new_book()

# ============================================================ ① サマリ(KPI+チャート)
ws = wb.active
ws.title = "サマリ"
setup_sheet(ws)
r = sheet_title(ws, "[資料タイトル] 業績サマリ", sub="[作成部署 / 作成日を記載]", span=13)

kpi_row(ws, "B5", [
    ("売上収益",   1500, NUM, "+5.6%"),
    ("営業利益",    120, NUM, "+12.0%"),
    ("営業利益率", 0.08, PCT, "+0.5pt"),
])

unit_note(ws, "M9", "(単位:億円)")

# チャートの元データは「データ」シートに置く(提出シートに生データを混ぜない)
ws_d = wb.create_sheet("データ")
setup_sheet(ws_d, tab_color=GREY, gridlines=True)
sheet_title(ws_d, "チャート元データ(作業用シート)", span=4)
info = write_table(ws_d, "B4",
                   ["年度", "売上収益", "営業利益"],
                   [["FY22", 1180, 62], ["FY23", 1290, 84], ["FY24", 1350, 95],
                    ["FY25", 1420, 107], ["FY26予", 1500, 120]],
                   formats=[None, NUM, NUM], widths=[10, 12, 12], freeze=False)
add_chart(ws, "bar", "B10", ws_d, info, value_fmt=PLAIN, width=22, height=9)

# ============================================================ ② 明細(スタイル表の見本)
ws2 = wb.create_sheet("明細")
setup_sheet(ws2, tab_color=STEEL)
r = sheet_title(ws2, "[明細タイトル] セグメント別内訳", span=7)
unit_note(ws2, "H4", "(単位:億円)")
write_table(ws2, "B5",
            ["セグメント", "FY24実績", "FY25実績", "FY26計画", "増減", "増減率"],
            [
                ["[セグメント1]", 4120, 4300, 4400, 100, 0.023],
                ["[セグメント2]", 3890, 3820, 3700, -120, -0.031],
                ["[セグメント3]", 2100, 2250, 2400, 150, 0.067],
                ["[セグメント4]", 890, 940, 1000, 60, 0.064],
                ["合計", 11000, 11310, 11500, 190, 0.017],
            ],
            formats=[None, NUM, NUM, NUM, NUM, PCT],
            widths=[20, 12, 12, 12, 10, 10],
            total_row=True, autofilter=False)

# ============================================================ ③ 書式ガイド(数値書式の見本)
ws3 = wb.create_sheet("書式ガイド")
setup_sheet(ws3, tab_color=GREY, gridlines=True)
r = sheet_title(ws3, "数値書式の見本(このシートは提出時に削除)", span=6)
write_table(ws3, "B4",
            ["用途", "定数", "正の例", "負の例"],
            [
                ["整数(標準)", "NUM", 1234, -56],
                ["小数1桁", "NUM1", 12.3, -4.5],
                ["割合", "PCT", 0.056, -0.012],
                ["円→百万円表示", "YEN_MM", 123456789, -9876543],
                ["日付", "DATE", "=DATE(2026,7,3)", ""],
            ],
            formats=[None, None, None, None],
            widths=[18, 10, 16, 16], freeze=False, zebra=True)
# 例セルに実際の書式を適用
for rr, fmt in [(5, NUM), (6, NUM1), (7, PCT), (8, YEN_MM), (9, DATE)]:
    ws3.cell(row=rr, column=4).number_format = fmt
    ws3.cell(row=rr, column=5).number_format = fmt
    ws3.cell(row=rr, column=4).alignment = Alignment(horizontal="right")
    ws3.cell(row=rr, column=5).alignment = Alignment(horizontal="right")

out = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "assets", "template.xlsx"))
os.makedirs(os.path.dirname(out), exist_ok=True)
wb.save(out)
print("saved:", out)

# プレビュー(Excel COM → PDF → PNG)
try:
    from export_preview import export_pngs, contact_sheet
    files = export_pngs(out)
    sheet = contact_sheet(files, os.path.join(os.path.dirname(out), "template-preview.png"))
    print("preview:", sheet)
except Exception as e:
    print("preview skipped:", e)
