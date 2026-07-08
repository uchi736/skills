---
name: corp-excel
description: >
  自社ブランドの Excel 資料(xlsx)を作る・直すときに必ず使う。業績サマリ・セグメント別
  明細・KPI一覧・提出用集計表など。同梱の Python ライブラリ(openpyxl ベース)が
  ハウススタイル——游ゴシック、青ヘッダ表、ゼブラ、△赤字の財務書式、KPIタイル、
  自社カラーのチャート、A4横・幅1ページ印刷、©フッター——を強制し、Excel COM で
  実画像検証まで行う。トリガー: 自社, Excel, xlsx, エクセル, 集計表, 業績, 明細, 管理表。
---

# 自社 Excel ハウススタイル

**役割分担**: xlsx の機械的な生成は同梱の `scripts/corp_xlsx.py`(openpyxl ベース)が担い、
このスキルは「自社の Excel としてどう見せるか」を固定する。pandas の使い方や openpyxl の
一般論はここには書かない。

## 鉄則(5つ)

1. **ゼロから書式を組まない** — 必ず `scripts/corp_xlsx.py` のヘルパーで書く。
   `new_book()` → `setup_sheet()` → `sheet_title()` → `write_table()` / `kpi_row()` / `add_chart()`。
2. **検証せずに納品しない** — 生成後、`scripts/export_preview.py`(Excel COM → PDF → PNG)で
   実画像を出して全ページ目視し、[references/checklist.md](references/checklist.md) と照合。
3. **数値書式は定数を使う** — `NUM`(整数・負値△赤)/ `NUM1` / `PCT` / `YEN_MM` / `DATE`。
   自前の書式文字列を発明しない。詳細は [references/formats.md](references/formats.md)。
4. **提出シートと作業シートを分ける** — チャート元データ・中間計算は「データ」シート
   (タブ灰色・枠線あり)へ。提出シート(タブ青・枠線なし)に生データを混ぜない。
5. **単位を必ず書く** — 表の右上に `unit_note()`。列見出しにも必要なら「(億円)」を併記。

## 作業手順

```
① 構成を決める   シート構成(サマリ/明細/データ)と各表の列設計
② スクリプトを書く  scratchpad に build_book.py を作りヘルパーAPIで記述
③ 生成           PYTHONUTF8=1 で実行
④ 検証           python scripts/export_preview.py out.xlsx → contact sheet を Read で目視
⑤ 修正ループ      checklist.md を満たすまで ②〜④
⑥ 納品           ユーザー指定場所(既定: Desktop)へコピー
```

最小コード:

```python
import sys; sys.path.insert(0, os.path.expanduser("~/.claude/skills/corp-excel/scripts"))
from corp_xlsx import *
wb = new_book()
ws = wb.active; ws.title = "サマリ"
setup_sheet(ws)                                   # 枠線OFF・A4横・幅1頁・©フッター
sheet_title(ws, "FY2026 業績サマリ", sub="経営企画部 / 2026-07-03", span=13)
kpi_row(ws, "B5", [("売上収益", 1500, NUM, "+5.6%"),
                   ("営業利益", 120, NUM, "+12.0%"),
                   ("営業利益率", 0.08, PCT, "+0.5pt")])
unit_note(ws, "M9", "(単位:億円)")
ws_d = wb.create_sheet("データ"); setup_sheet(ws_d, tab_color=GREY, gridlines=True)
info = write_table(ws_d, "B4", ["年度", "売上収益"], [["FY25", 1420], ["FY26予", 1500]],
                   formats=[None, NUM], widths=[10, 12], freeze=False)
add_chart(ws, "bar", "B10", ws_d, info)           # サマリにデータシート参照のチャート
ws2 = wb.create_sheet("明細"); setup_sheet(ws2, tab_color=STEEL)
sheet_title(ws2, "セグメント別内訳", span=7)
write_table(ws2, "B5", ["セグメント", "FY25", "FY26計画", "増減率"],
            [["航空・宇宙", 4300, 4400, 0.023], ["資源・エネ", 3820, 3700, -0.031],
             ["合計", 8120, 8100, -0.002]],
            formats=[None, NUM, NUM, PCT], widths=[20, 12, 12, 10], total_row=True)
wb.save("out.xlsx")
```

## スタイル規定(ヘルパーが自動適用)

| 要素 | 規定 |
|---|---|
| フォント | 游ゴシック 11pt(既定で全セル)。純黒でなく `#262626` |
| 表ヘッダ | BLUE `#004EA2` 塗り・白太字・中央・下罫線 NAVY medium |
| 1列目 | 行見出し扱い: 淡青塗り・太字(`first_col_header=False` で解除) |
| ゼブラ | 偶数行 `#F0F5F8` |
| 合計行 | `total_row=True`: 二重上罫線 NAVY+太字+淡青塗り(最終行が合計であること) |
| 負値 | `△123` 赤(数値書式が自動処理。**文字列で"−"を書かない**) |
| KPIタイル | 上端に太い青罫線+淡背景。補足は `+`→緑 / `△`→赤 に自動色分け |
| チャート | 系列色 BLUE→CYANSUB→TEAL→…、横グリッドのみ淡灰、凡例下部 |
| シート | A列は幅2.5の余白列。内容はB2から。タブ色=提出:青/明細:鋼青/作業:灰 |
| 印刷 | A4横・幅1ページ・©フッター+ページ番号(setup_sheet が設定) |

## よくある失敗(ここが一番効く)

- 数値を**文字列**で入れる("1,234" と書く)→ 数値のまま入れて書式(NUM)に任せる。
  ソート・集計・チャートが壊れる。
- 負値をハイフンやマイナス記号の文字列で表現 → 書式の △ 赤に任せる。
- % を 5.6 のような数で入れる → **0.056 の比率**で入れて PCT を当てる。
- 提出シートに枠線(グリッド線)が残る → `setup_sheet()` を通せば消える。通し忘れ注意。
- 単位の書き忘れ / 「予」「計画」の明示漏れ → IR・報告資料として通らない。
- セル結合の乱用 → 結合は KPI タイルとタイトル行だけ。表の中では結合しない
  (ソート・フィルタ・参照が壊れる)。
- 色の発明 → パレット(corp_xlsx.py の定数)以外を使わない。赤は負値専用。
- 「Sheet1」のまま / タブ色なし → シート名は日本語で意味を持たせ、タブ色で種別を示す。
- マージン計算せずチャートを表の上に置く → チャートは9cm≈17行。重なりはプレビューで必ず確認。

## ファイル構成

```
scripts/corp_xlsx.py        ヘルパーAPI本体(色・書式・表・KPI・チャートの単一ソース)
scripts/export_preview.py  Excel COM → PDF → PNG の実機検証
scripts/build_template.py  assets のテンプレを再生成(スタイル改良時はここを直す)
assets/template.xlsx   見本ブック(サマリ/データ/明細/書式ガイド)
assets/template-preview.png 実機レンダリング一覧
references/formats.md      数値書式の完全リファレンス
references/checklist.md    納品前チェックリスト
```

パワポ側のハウススタイルは [corp-slides](../corp-slides/SKILL.md) スキル(配色は共通)。
