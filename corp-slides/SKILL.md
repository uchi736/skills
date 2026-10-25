---
name: corp-slides
description: >
  自社ブランドのプレゼン資料(パワーポイント/pptx)を新規に作るときに必ず使う。
  公式テンプレート(同梱・git管理外)のマスターを土台に、Deck API で表紙・目次・中扉・要旨・
  箇条書き・対比・概念図(番号付き凡例+線画アイコンの構成図)・関係図・結論ツリー・関門図・手順・
  グラフ(全幅・図上注記・左右の問い chart_pair)・論証テーブル・数値目標の表・評価マトリクス・
  2軸マトリクス・ロードマップ・ガント・壁→打ち手・結びを組み、PowerPoint COM の実画像と
  機械検査で検証してから納品する。既存デッキの編集・解析は公式 pptx スキルを使う。
  トリガー: 自社, スライド, パワポ, pptx, プレゼン, 説明資料, 検討資料, 提案資料, IR資料。
---

# corp-slides — 自社ハウススタイルの新規デッキ

公式テンプレートの上に、決まった型だけでスライドを組む。座標・色・フォントは
`scripts/corp_pptx.py` の `Deck` が持つので、手で置かない。

## 読む順番

1. この SKILL.md(手順・鉄則・テンプレート)
2. 全型の見本 PNG: `assets/template-16x9_pptx_png/スライドN.PNG`(下の「型と見本」表で番号を引く)。
   各枚の出典の上に「用途 / 避ける → 代替の型」の 1 行がある。番号が — の型は同じ型の変種で見本は無い
   (引数は layouts.md)。無ければ `python scripts/build_template.py 16:9` で作る
3. 型の引数と上限: [references/layouts.md](references/layouts.md)
4. グラフ・表を使うとき: [references/charts-tables.md](references/charts-tables.md)
5. 検証のとき: [references/checklist.md](references/checklist.md)
6. 余裕があれば実物の見本: `assets/ideal/index.md`(before/after)、`assets/ideal/abeam/README.md`、
   `assets/ideal/ir-fy2026/README.md`。**見本と鉄則が食い違うときは鉄則が優先**

使わないとき: 既存 pptx の編集・解析・テキスト抽出は公式 `pptx` スキル。corp-slides は新規作成専用。

## 鉄則

1. **型だけで組む** — `Deck` の型メソッドと部品で作る。型に無い図は部品
   (`slide`/`lead`/`panel`/`numbered_list`/`node_icon`/`arrow_h` など)で組み、`rect`/`text` の手置きは最後の手段。
2. **1 スライド 1 キーメッセージ** — タイトルの下に ▍青バー+太字の lead を 1 文。
   lead は**読み手に持ち帰らせる判断**を 1 文で(「〜する」「〜が要る」「〜に届かない」)。説明や件数
   (「〜は 4 系統」)で終えない。「〜について」「〜の分析」は不可。途中に「。」のある 2 文は警告。
   全角 40 字以内なら 20pt の 1 行、長くても 90 字・2 行まで(超えると警告)。list を渡すとエラー。
   lead は内容の全部を言う(表+下段なら両方、グループを列挙するなら全部)。
3. **要旨は固定構成** — 冒頭の exec_summary は 目的 → 背景 → 実施内容 → 結果 → 結論。
   「■ 目的」の見出し行の下に本文を置く縦のラダー形(青いチップにしない)。結論は他の型の結論帯と同じ
   淡青の帯に中央揃えの青太字 18pt(ラベルなし。濃紺の帯にしない)。
   資料の種類が違っても順と見出しを変えない。お願い帯(依頼・確認の帯)は付けない。
4. **文字は図の中でも 12pt 以上** — 本文 14〜16pt、見出し 14pt 以上。10〜11pt は出典・注記・単位・
   ページ番号だけ。`check_deck.py` が 12pt 未満を検出する。
5. **数字は意味とセットで置く** — kpi タイル(大数字だけの箱)は廃止。数値は
   表(指標・測り方・しきい値・現状)か、関門表(gate_table: 関門 × 測り方 / 合格ライン / 判定者)か、グラフ。
6. **グラフは全幅、読み取りは図の上** — 右に「ポイント」欄を置かない。結論は lead、注目点は
   `highlight`(主役の棒)/ `bracket`(範囲)/ `callouts`(真上の一文)/ `target`(目標の破線)。
   問いが 2 つなら `chart_pair`(Q: 見出し+下線で左右)。
7. **箇条書きは使ってよい** — 並列の事実・方針・条件は content が一番速い。
   関係・流れ・比較・数値が主役のときに図にする。同じ型の 3 連続は避ける。
8. **AI くさい見た目を避ける** — 薄青の角丸カードの横並び、入れ子の角丸、チップの羅列、
   手順でもないのに等幅の箱を一列、毎枚の結論帯、15〜25 字の断片の羅列、
   縦の番号丸を細い線でつないだステッパー、「壁 / 打ち手」のようなラベル語の対、小さな灰色の ▶。
   事実の格子(関門・工程・課題)は塗りのない軽い表のほうが読める。
   代わりに 領域は `panel`(BLUE ヘッダ帯+白本体+GREY 細枠)、上に表や図が既にある下段は
   `section_block`(見出し+下線の軽い段組。青帯を重ねない)、柱は `header_box`、
   仕組みは「左 panel に番号付き凡例 ①〜⑤+右 panel に線画アイコンの構成図(badge で番号対応)」。
9. **色はパレットだけ** — 見出し帯は全て BLUE で揃える。NAVY は主役の棒・推奨列・番号バッジ・ガントの
   key だけ(帯や結論には使わない。結論帯は全て淡青)。アイコンは BLUE 1 色。赤は負値・未達だけ。
   純黒は使わない(本文 INK #262626。タイトルも INK に上書きされる)。フォントは游ゴシック。
10. **検証は欠陥ハント** — 実画像を全ページ見て、問題はあると仮定して探す。直したら再レンダリング。
    ユーザーへの確認依頼は挟まず、直し切ってから納品する。

## 作業手順

```
① 題材を読む     元資料(docx/pdf/xlsx)から、事実・数字・章番号を拾う。数字を作らない。
                  元資料に無い数字を例として置くときは (例)、案なら (案) と書き、source= に出所を書く
② 設計図を書く   下の形式で全スライドを決めてから書き始める(ここで考え終える)
③ build_deck.py  scratchpad に下のテンプレートを複製して埋める
④ 生成           python build_deck.py(save() が出す「注意」は全て該当ページを見る)
⑤ 検証           export_preview.py で PNG → 一覧 → 1 枚ずつ Read。check_deck.py で機械検査
⑥ 修正ループ     ③〜⑤ を、注意 0・検査 0・目視の欠陥 0 になるまで回す
⑦ 納品           Desktop(指定があればそこ)へコピー。同名があれば上書きせず _v2 などの別名
```

### ② 設計図の形式

| No | 型 | タイトル(体言止め・1 行) | lead(主張の 1 文) | 中身(データ・出典) |
|---|---|---|---|---|
| 1 | cover | 資料名 | — | 日付・部署 |
| 2 | exec_summary | 本資料の要旨 | (結論帯が兼ねる) | 目的/背景/実施内容/結果/結論 |
| 3 | toc | 目次 | — | 章 3〜6 |
| 5 | 概念図 | ○○の仕組み | 仕組みの結論 | 凡例 5 項目+構成図 |

- 資料全体で一番言いたいことを 30 字以内で先に書く。各スライドの lead はそれを支える。
- 各章の頭に図を 1 枚置く。図解は内容の形で選ぶ: 仕組み・置き場 → 概念図 / layers、流れ → cycle_flow /
  swimlane / step_flow、根拠・分解 → tree(縦 = reason_tree / 横)、目標との差 → goal_bars / chart(target)、
  計画と現在地 → timeline(now=) / gantt(today=)。save() が「図解の無い章」と「表が半分超」を警告する。
- 中扉は本文 10 枚以上か 4 章以上のときだけ。各章の本文は 3 枚以上(少なければ章を束ねるか目次だけにする)。
- 結びの直前に推奨か次の一手(動詞・担当・期日)の本文を置き、結び(message)は完全な文+次の一手+連絡手段。
- 1 スライドに主張が 2 つあれば分ける。迷ったら layouts.md 末尾の「型選択ガイド」。

### takeaway(淡青の結論帯)を付ける基準

- lead と**違う情報**(担当・条件・次の判断)を足せるときだけ。lead の言い換えなら付けない。
- デッキ全体の 1/3 以下、2 枚連続させない。exec_summary・reason_tree・message には付けない。
- 「付けない枚が多くてばらばらに見える」は問題ではない。付ける基準が「足せる情報があるか」なので、無い枚には無いのが正しい。

### ④ build_deck.py テンプレート(そのまま複製して埋める)

```python
# -*- coding: utf-8 -*-
import os, sys
sys.path.insert(0, os.path.expanduser("~/.claude/skills/corp-slides/scripts"))
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
from corp_pptx import Deck

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "deck.pptx")
SRC = "○○仕様書 rev1.0"          # 出典の共通部分(各スライドで章番号を足す)
d = Deck()                          # 16:9。公式テンプレが無ければ中立ダミーで出る

d.cover("資料名", subtitle="副題(課題か目的)", label="○○会議(第N回)", date="2026年10月20日", dept="○○部")
d.exec_summary("本資料の要旨",
    purpose="この資料で何を決めるか(1〜2 文)",
    background=["なぜ今か。数字を入れた完全文", "2 本目"],
    actions=["何をどう進めたか / 進めるか"],
    results=["決めたこと・条件・判断点(計画段階なら『本資料で決めたこと』)"],
    conclusion="結論の 1 文(2 行まで)")
d.toc([("背景", "補足"), ("提案", "補足"), ("計画", "補足")])   # 最初の目次は current を渡さない
d.section(1, "背景")                                                # 中扉が章一覧と現在地を示す

# 概念図(章頭の 1 枚): 左 panel に番号付き凡例、右 panel に構成図
s = d.slide("○○の仕組み")
y0 = d.lead(s, "仕組みの結論を 1 文で")
src = SRC + " 第2章"
ph = d.body_bot(src) - y0
bx, ly, bw, bh = d.panel(s, d.MX, y0, 4.9, ph, "構成要素")
d.numbered_list(s, bx, ly, bw, [("見出し", "説明は 1 行 24 字以内")] * 5, max_y=ly + bh)
rx, rw = d.MX + 5.3, d.CW - 5.3
_, ry, rw_in, rh_in = d.panel(s, rx, y0, rw, ph, "構成図")
xs, (r1, r2) = d.concept_grid(rx, ry, rw, rh_in, cols=3, rows=2)   # 列の中心 x と 2 行の y
c1, c2, c3 = xs
d.node_icon(s, c1, r1, "doc", "入力", "補足", badge=1)
d.arrow_h(s, c1 + 0.45, c2 - 0.45, r1, label="受け渡し")
d.node_icon(s, c2, r1, "chat", "処理", "補足", badge=2)            # アイコンは BLUE 1 色・番号はラベルの左
d.arrow_h(s, c2 + 0.45, c3 - 0.45, r1)
d.node_icon(s, c3, r1, "laptop", "出力", "補足", badge=3)
d.arrow_h(s, c3 - 0.45, c1 + 0.45, r2, label="戻り", dash=True)   # 戻りの流れは破線
d.source(s, src)

d.content("提供範囲", lead="主張の 1 文",
          bullets=[("親(子を持つ項目)", ["子の文", ("子2", ["例:孫は例示に"])]),
                   ("親2", ["子の文"])],
          source=SRC + " 第1章")
d.chart("月次の推移", "bar", ["4月", "5月", "6月"], [("件数", [820, 1050, 1340])],
        lead="主張の 1 文", unit="件", subtitle="何の図か+範囲(例)",
        highlight=2, callouts=[(2, "注目点の一文")], source=SRC + " 7.1")
d.table("目標と現状", ["指標", "測り方", "しきい値", "現状"],
        [["根拠が正しい率", "週次の抜き取り", "90% 以上", "92%"],
         ["防御率", "攻撃テストセット", "95% 以上", "△93%"]],
        lead="主張の 1 文", col_widths=[3, 3, 2, 2], font_size=14, highlight_row=1,
        source=SRC + " 7.2")
d.message("結論の完全な文(何を・どうするか。40 字以内)",
          sub="次の一手の 1 文(動詞で終え、担当と期限を入れる)",
          contact=["お問い合わせ:○○部(Teams: #channel)"])
d.save(OUT)
print("saved:", OUT)
```

型ごとの全引数は layouts.md。見本の実コードは `scripts/build_template.py`(30 枚分。迷ったら該当箇所を写す)。

### ⑤ 検証コマンド

PowerShell:
```powershell
$env:PYTHONUTF8 = "1"; $S = "$HOME\.claude\skills\corp-slides\scripts"
python build_deck.py
python "$S\export_preview.py" deck.pptx      # deck_pptx_png\スライドN.PNG と deck_ALL.png
python "$S\check_deck.py" deck.pptx          # issues: 0 になるまで直す
```
Git Bash:
```bash
S=~/.claude/skills/corp-slides/scripts
PYTHONUTF8=1 python build_deck.py
PYTHONUTF8=1 python $S/export_preview.py deck.pptx
PYTHONUTF8=1 python $S/check_deck.py deck.pptx
```

1. `deck_ALL.png` で全体の並び・密度・型の偏りを見る
2. `スライドN.PNG` を **1 枚ずつ全部** Read で見る(一覧の縮小画像では文字の欠陥が見えない)
3. checklist.md の「視覚QA」に沿って欠陥を列挙し、直して再生成する
4. 重要な資料は、PNG を Explore サブエージェントに渡して先入観のない目で検査させる
   (依頼文は checklist.md 末尾)

### ⑦ 納品

- 生成物は scratchpad で作り、完成版だけ Desktop にコピーする。
- 同名ファイルがある・ユーザーが手で直したファイルは上書きしない(`_v2` などの別名)。
- PowerPoint で開いていると保存できない(save() が止まって知らせる)。別名で保存する。
- 報告は「ファイルの場所・枚数・構成の要点・(例)として置いた数字」を短く。確認依頼はしない。

## 型と見本(見本 PNG の番号)

| 場面 | 型 | 見本 |
|---|---|---|
| 表紙 / 要旨 / 目次 / 型の選び方 / 中扉 / 結び | `cover` / `exec_summary` / `toc` / `table`(選定表) / `section` / `message` | 1 / 2 / 3 / 4 / 5 / 30 |
| 並列の事実・方針・条件 | `content`(■●□ のラダー) | 6 |
| 見出しの強い 4〜6 グループ(項目ごとに観点が違う。同じ観点で書けるなら table。「表は堅い」は理由にならない) | `content(style="blocks")` | 7 |
| 変化(前後)・案の対比・並列の柱 | `boxes(arrow=True / False)`(柱は arrow=False。同じ属性が並ぶなら表)。到達点・土台の行は `goal=` / `base=`、箱ごとの効果の数字は `figures=` | 8 |
| 仕組み・データの流れ(章頭) | 概念図: `panel`+`numbered_list`+`concept_grid`/`zone_cols`+`node_icon`/`arrow_h`/`arrow_v` | 10 |
| 役割分担・受け渡し(矢印にラベル) | `boxes(arrows=[...], band=)`(= `flow`。8 の箱に矢印ラベルと土台) | — |
| 結論と根拠 2〜3 | `tree(direction="down")`(= `reason_tree`) | 11 |
| 関門・条件・合格基準 | `gate_table`(軽い表+※注。下段は関門の直後のことだけ) | 12 |
| 手順そのもの | `step_flow`+`section_block` | 13 |
| 循環する工程(改善の輪・PDCA・運用ループ) | `cycle_flow`(2 段の周回路。工程カード+▶。中央には何も置かない) | — |
| システム構成・階層(何がどの層にあるか) | `layers`(層の帯+要素の箱。`side=` で横断の帯) | 14 |
| 担当 × 工程の流れ(誰が・いつ・何を) | `swimlane`(箱を流れの順に。矢印は自動) | 15 |
| 目標 ← 要因 ← 具体策 / 課題の分解 | `tree`(横。direction="right"。根の箱 → 枝(数値)→ 葉を「項目名 \| 数値」の表の行に。列見出しは `levels=`) | — |
| アイコンの種類 | `icon`(13 種・線画) | 16 |
| 推移・量の差 | `chart("bar")`(値軸なし・bracket に根拠の数字) | 18 |
| 構成の推移 | `chart("stack")`(合計・直接ラベル・share_of・forecast_from) | 19 |
| 問いが 2 つ・目標線 | `chart_pair`(`target=`) | 20 |
| 増減の要因 | `chart("waterfall")` | 21 |
| 課題→現状→打ち手→担当 / 対象 × 観点 | `table`(論証・`zebra=False`。別の群は `sections=` で下段に) | 22 |
| 指標・しきい値・現状 | `table`(数値。未達は △ と `highlight_row`) | 23 |
| 目標に対する現在地(指標 4 つ以下) | `goal_bars`(起点→現在→目標の物差し+差。5 つ以上は 23 の表) | — |
| 候補の選定 | `matrix_eval` | 24 |
| 立ち位置・絞り込み | `matrix_2x2`(点+ラベル) | 25 |
| 粗いフェーズ(+現在地) | `timeline`(`now=`) | 27 |
| 詳細な行程(+現在) | `gantt`(月の等間隔+`axis_label=`+`today=`。style は main / option / key) | 28 |
| 手順+壁と打ち手 | `process`(表の形: No・工程・壁・打ち手) | 29 |

図解は編集できるネイティブ図形で組む。上の型で表せない自由な図(型の比較・境界・ふるい・工程の格子など)は
`html_shapes.html_diagram` で HTML/CSS で組み、ネイティブ図形(矩形・文字・コネクタ矢印・線画アイコン)に変換する
(画像にしないので PowerPoint で直せる。layouts.md「HTML 図解」)。HTML→PNG の `figure` は写真のような
特殊な絵だけの最後の手段(使ったら .html を必ず残す)。

**表の様式は 2 系統で固定**: 数値と論証の表(table・matrix_eval)は濃色ヘッダ+1 列目の淡い塗り。
図解・計画の中の格子(gate_table・process・gantt・swimlane・型の選び方)は塗りのない罫線だけ。1 枚の中で混ぜない。

## カラーパレット(この表が正)

| 定数 | HEX | 用途 |
|---|---|---|
| NAVY | `#01357A` | 主役の棒・推奨列・番号バッジ・ガントの key だけ(帯・結論には使わない) |
| BLUE | `#004EA2` | **全ての見出し帯**(パネル・対比の箱・柱・関係図・表のヘッダ)。帯の色はこれ一択 |
| BLUE_TX | `#00549A` | 見出しテキスト・注記 |
| STEEL | `#3377AE` | 矢印・補助 |
| CYAN | `#00BCEB` | ▶・アクセント線 |
| CYANSUB | `#66C5DC` | 第 2 系列 |
| TEAL | `#00BB99` | 増加(少量) |
| RED | `#C00000` | 負値・未達(△)専用 |
| BG / PALE / LTBLUE / PERI | `#F0F5F8` / `#E8F2FA` / `#D6ECF7` / `#C5D0E6` | 淡背景・表の 1 列目・強調行と結論帯・表紙帯 |
| INK / INK2 / GREY | `#262626` / `#4D4D4D` / `#A6A6A6` | 本文・補足・罫線 |

## 階層マーカー(全型共通)

```
▍ キーメッセージ(lead。青バー+太字 20pt / 長いと 18pt。1 スライド 1 つ)
■ 親(子を持つ項目。太字)
　　● 子
　　　　□ 孫(グレー。「例:○○」の置き場)
```

- 親=■ は子を持つ項目だけ。子を持たない並列は ● で並べる。・✓◆► や数字マーカーを混ぜない。
- ■ は体言止めの名詞句、● は 1 項目 1 文(〜する)、□ は補足 1 行で 1 枚 1 か所。
- content で ■ を使うと ■18/●16/□14pt 基準(入らないときだけ 17→16→15→14)。● だけなら 16pt 基準。
  箱・段組・表の中は ■16/●14 基準。それでも入らなければスライドを分ける。

## よくある失敗

- 型の説明を lead に書く(「並列の柱は濃色ヘッダの箱で立てる」)→ lead は**内容の主張**。型の説明は不要。
- lead と中身・結論帯が食い違う(lead「7 月から目標達成」なのにグラフは 8 月から)→ 数字を突き合わせる。
- lead が上半分しか言っていない(関門の表+下段の運用があるのに関門だけ)→ 1 文で両方を言う。
- 条件・関門を色箱と矢印で描く → `gate_table`。事実の格子は塗りのない軽い表のほうが読める。
- 他の案件の表をそのまま流用する / 実在 OSS のライセンスを断定する → 題材の資料から書く。候補名は匿名化してよい。
- 1 文字だけの行・語の途中の折り返し → 柱・箱・カードは語の切れ目で自動改行する。残るなら文を削るか `\n` で切る。
- 判定者の欄に「LLM 判定」「判定モデル」(道具)を書く → 判定者は人か会議体。道具は「何で測る」に。
- 本文 2 枚の章に中扉を置く → 章を束ねるか目次だけにする(save() が警告)。
- 結びがキャッチコピー(「A で B、C で D」)→ 何を・どうするかの完全な文+次の一手(担当・期限)。
- 進捗・中間レビューの資料に「ゴールに対する現在地」が無い → goal_bars か timeline(now=)/gantt(today=) を 1 枚。
- 表ばかりで図解が無い → 仕組みは layers、流れは cycle_flow / swimlane、分解は tree に置き換えられないか見直す。
- 箱に短文 1 行だけ・下半分が空 → 項目を足すか箱を小さく。箱には「役割+具体例・数字」を入れる。
- 目次 7 章以上 → `toc` がエラー。章を束ねる。
- PowerPoint で開いたまま保存 → 別名で保存。
- 4:3 で作る → 公式は 16:9(14.347×8.071 inch)。4:3 は依頼されたときだけ `Deck(aspect="4:3")`。

## ファイル構成

```
SKILL.md                         この手順書
references/layouts.md            型カタログ(引数・上限・使い分け・概念図の部品)
references/charts-tables.md      グラフ・表の規定
references/checklist.md          納品前チェックリスト+別視点検査の依頼文
scripts/corp_pptx.py             Deck API(色・型・座標の単一ソース)
scripts/build_template.py        全型の見本デッキ(30 枚)を再生成。実コードの見本
scripts/export_preview.py        PowerPoint COM で PNG 化+一覧
scripts/check_deck.py            機械検査(プレースホルダ・空枠・純黒・タイトルの継承色・12pt 未満・「〜について」)
scripts/html_shapes.py           HTML 図解を編集できるネイティブ図形に変換(html_diagram / place_html。ヘッドレス Edge で測る)
scripts/html2png.py              HTML 図解を PNG に(figure 型。編集性を捨てる特殊ケースだけ)
scripts/make_ideal_kit.py        納品デッキから理想見本キット(BEFORE/AFTER の対)を作る
scripts/make_ideal_pairs.py      直したキットを assets/ideal/ の PNG ペアに追記する
assets/fig-template.html         HTML 図解の雛形(ハウス色・部品入り)
assets/brand/                    公式テンプレ format-16x9.pptx と brand.json(社名等)(git 管理外)
assets/template-16x9*.pptx       見本デッキ(生成物・git 管理外)
assets/template-16x9_pptx_png/   見本の実機 PNG(スライドN.PNG)
assets/template-preview-16x9.png 見本の一覧
assets/ideal/                    実物の見本(社内資料・第三者 PDF を含むため git 管理外・公開しない)
```

- 依存: Python 3.10+、python-pptx、Pillow。PNG 検証は Windows+PowerPoint+pywin32。
  PowerPoint が無い環境でも pptx の生成はできる(export_preview は理由を出して止まる)。
- 公式テンプレがあればロゴ・表紙/扉/最終ページの意匠・コピーライト・ページ番号は本物が付く。
  無ければ中立ダミー(LOGO・© Your Company)で同じ型が出る。テーマの和文フォントは游ゴシックに自動で直る。
- 公開リポジトリに載るのはコードとドキュメントだけ(assets/brand・assets/ideal・生成物は .gitignore)。
