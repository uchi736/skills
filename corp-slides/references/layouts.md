# レイアウトカタログ

すべて `scripts/corp_pptx.py` の `Deck` メソッド。見本は `assets/template-16x9_pptx_png/スライドN.PNG`
(番号は SKILL.md「型と見本」表)、実コードは `scripts/build_template.py`。
キャンバスは 14.347×8.071 inch。本文の安全領域は x: `MX`(0.87)〜`MX+CW`(≈13.47)、
y: `BODY_TOP`(≈1.30)〜`BODY_BOT`(H-0.62)。座標を書くときは定数と `body_bot()` を使う。

## 共通スロット

どのスライドも **タイトル → ▍lead → 内容 →(takeaway)→ 注記・出典** の縦一本。

| 引数 | 中身 |
|---|---|
| 表の様式 | 2 系統で固定。数値・論証の表(table・matrix_eval)=濃色ヘッダ+1 列目の淡い塗り。図解・計画の中の格子(gate_table・process・gantt・swimlane)=塗りのない罫線だけ。1 枚の中で混ぜない |
| `lead=` | ▍青バー+太字で、読み手に持ち帰らせる判断を 1 文(文字列のみ。説明・件数で終えない)。全角 40 字以内は 20pt・1 行、それ以上は 18pt・2 行まで。2 文(途中に「。」)や「〜について」で終わる文は警告。戻り値の本文開始 y は 1 行≈2.23 / 2 行≈2.55 / lead なし≈2.15 |
| `takeaway=` | 内容の下の淡青の結論帯(14pt)。lead と違う情報を足せるときだけ。連続した 2 枚、デッキの 1/3 超は save() が警告 |
| `note=` / `source=` | 「注:」「出典:」の行(10pt)。出典 H-0.71、注 H-0.92。本文は自動でその上で止まる |
| `subtitle=` | グラフの定義行「何の数字か+集計の単位+期間(年度)」(13pt) |

- `d.note("...")` はスピーカーノート(引数の note= とは別物)。
- 生成中の注意(はみ出し見込み・泣き別れ・章の枚数・結論帯の連続など)は `d.warnings` に溜まり、`save()` が表示する。
  表示されたページは必ず PNG で見る。
- 文字列の中の `\n` は改行になる(箇条の切る位置を自分で決めたいとき)。柱・箱・カード・ロードマップの箇条は、
  1 行に入らない文を「、。」や助詞の直後で自動で改行する(語の途中で折れない)。

---

## 基本の型

### cover — 表紙(見本 1 番)
```python
d.cover(title, subtitle=None, label=None, date=None, dept=None, company=None)
```
- title は全角 20 字目安(30 字を超えると 2 行になり警告)。タイトルか subtitle のどちらかに課題・目的を入れる
  (「○○の全体像」だけで終えない)。
- label は会議名と回(「生成AI勉強会(第3回)」。社外向けは「○○株式会社 御中」)。
- date と dept(作成部署)は版の管理に要るので省かない(無いと警告)。company 省略時は brand.json の社名。ページ番号なし。

### exec_summary — 要旨(冒頭の 1 枚。見本 2 番)
```python
d.exec_summary(title, conclusion, purpose=None, background=None, actions=None, results=None,
               lead=None, labels=None, source=None, note=None)
# purpose/background/actions/results: 文字列 または ["文", ...](● の箇条)。None の行は省く
# conclusion: 最下段の淡青の結論帯(中央揃えの青太字 18pt・2 行まで)。キーメッセージを兼ねる
# labels: 見出しを変えるときだけ {"purpose": "狙い"}。既定 目的/背景/実施内容/結果/結論
```
- **目的 → 背景 → 実施内容 → 結果 → 結論** の固定構成。資料の種類が違っても順と見出しを変えない。
- 縦のラダー形: 「■ 目的」の見出し行、その下に本文(1 文なら平文、複数なら ● 箇条)、項目の間は細い罫線。
  結論は濃紺の帯でなく、他の型の結論帯(takeaway)と同じ淡青 LTBLUE の帯に中央揃えの青太字 18pt(ラベルなし)。
- 本文は 16pt 基準、入らないときだけ 15→14pt。各行 40〜70 字の完全文で数字を入れる。箇条は各 2〜3 本。
  列挙は「・」でつなぎ、半角スペースで区切らない。
- 「結果」は計画段階なら「本資料で決めたこと」(条件・判断点・行程)。
- `ask=`(お願い帯)・`bullets=`・`reasons=` は廃止(渡すと TypeError)。

### toc / section / message(見本 3・5・30 番)
```python
d.toc(items, current=None, title="目次", pages=True)   # items: ["章"] か [("章", "補足")]。3〜6 章
d.section(number, title, nav=True)                      # 中扉。目次の章一覧を再掲し、現在の章を強調
d.message(text, sub=None, contact=None)                 # 結び
```
- toc: 最初の目次では `current` を渡さない(中扉が現在地を示す)。`pages=True` で各章の開始ページが右端に入る
  (save() 時に埋まる)。章名は中扉と一字一句同じにする(section() が照合して警告)。
- section: 章番号と章名の下に全章の一覧を再掲し、現在の章だけ淡青の帯+青太字で示す(全体のどこにいるか)。
  中扉は本文 10 枚以上か 4 章以上のときに使い、各章の本文は 3 枚以上にする(少なければ save() が警告。
  満たさなければ章を束ねるか目次 1 枚で済ませる)。`nav=False` で一覧を消せる。
- message: text は結論か依頼の**完全な文**(何を・どうするか。24 字以内 28pt / 40 字以内 24pt。
  対句やキャッチコピーにしない)。sub は次の一手の 1 文(動詞で終え、担当と期限を入れる。無いと警告)。
  contact は部署名と連絡手段(「ヘルプデスク改善チーム(Teams: #faq-bot)」)。
  message の直前の本文には、推奨か次の一手(動詞で始め、担当と期日を書く。table の「アクション|担当|期日」、
  timeline、gantt など)を置く。

### content — 箇条書き(見本 6 番)
```python
d.content(title, lead=None, bullets=None, source=None, style="ladder", note=None, takeaway=None)
# bullets: ["文"] または [("親", ["子", ("子2", ["孫"])])] — 3 階層まで
```
- 並列の事実・方針・条件を読ませる型。**箇条書きは悪くない**。図にするのは関係・流れ・比較・数値が主役のとき。
- ladder: ■親(太字)→ ●子 → □孫(グレー)。■ があれば ■18/●16/□14pt 基準、入らないときだけ 17→16→15→14。
  ● だけのリストは 16pt 基準。■ は ▍ の左端に揃え、■ のグループの間は項目間の約 2 倍空く。
- 書き方: ■ は体言止めの名詞句で揃える。● は常体の文(〜する / 〜だ)で揃え、1 項目 1 文。
  □ は ● の補足(例・条件・注のどれか 1 つ)1 行に限り、1 枚 1 か所まで。
- 上限: ■ は 3〜4 個(5 個以上は 2 枚に分けるか blocks)。■ を使わない ● だけなら 6 個。超えると警告。
- `style="blocks"`(見本 7 番): 青太字の見出し(18pt)+本文(16pt)の段組。グループ間は細罫線。
  4 グループ以上は 2 列で、**左上→右上→左下→右下**の順に並び、行の高さが揃う。6 グループまで、各 2〜4 行
  (行数の差が 2 以上なら警告)。lead でグループを列挙するなら全部入れる(入らないなら数で言う)。
- **blocks か table かの判断**(3 つの場合):
  1. 項目の中身を同じ観点で書ける(6 つの対象がどれも「管理の内容」「確認の契機」)→ **table(対象 × 観点)**。
     列で縦に見比べられる。
  2. 項目ごとに観点が違う(汎用 LLM はライセンス、埋め込みはナレッジの作り直し、と軸がばらばら)→ **blocks**。
     表にすると列が立たない。見本 7 番がこれ。
  3. 同じ観点で書けるのに blocks にしたくなる → たいてい「表は堅い」という見た目の理由で、読みやすさの理由ではない。
     table にする。見せたいのが「項目の中身」でなく「項目の間の比較」なら、blocks は仕事をしていない。
  観点の揃う群と揃わない群が混ざるなら、揃う群を table、残りを `sections=` の下段に分け、lead で両方を言う。

### boxes — 対比・前後・並列の柱・受け渡し(見本 8 番)
```python
d.boxes(title, boxes, lead=None, arrow=False, source=None, note=None, takeaway=None, arrows=None, band=None,
        goal=None, base=None, figures=None)
# boxes: [(見出し, 小見出し|None, [項目...]), ...] 2〜3 個
# arrows=["ラベル", ...](len-1 個)と band="共通の土台" を渡すと、ラベル付き矢印の受け渡し図(= flow)
# goal=(ラベル, 文) / base=(ラベル, 文): 柱の上の「到達点」の行 / 柱の下の「土台」の行(矢印は付かない)
# figures=[(ラベル, "前 → 後")|None, ...]: 箱ごとの数値帯(箱の最下段。「→」の後の値だけ大きく太く)
```
- 変化(Before/After)は `arrow=True`(細く高い ▶)、並列の案は False。
- 箱は BLUE 見出し帯(17pt)+白本体+GREY 細枠。小見出しは青太字 16pt+細い下線。本文は全箱共通で 16→15→14pt。
- 各箱 4〜5 項目で揃え、文末の形(〜する / 体言止め)も揃える。箱の下半分が空くなら項目を足す。
- 並列の柱(3 本柱・3 観点)は `arrow=False`・小見出しなしで同じ型(見た目が同じなので型は分けない)。
  柱ごとに中身の種類が違うときだけで、各柱に同じ属性が同じ順で並ぶなら gate_table か table(数値)。
- **到達点 → 柱 → 土台**(中期計画の骨格など「柱が到達点を支え、共通の基盤の上に立つ」)は `goal=` と `base=`。
  どちらもラベル(青太字 12pt)+文(太字 14.5pt)の 1 行と GREY の細い罫線だけで、帯・箱・三角屋根にはしない
  (帯で挟む案・屋根の案と比べて選定。2026-10-25)。文は 1 行に収める(超えると警告)。
- **数値帯(`figures=`)**は、箱ごとの効果・実績を数字で見せたいとき(時間削減効果、件数の増減など)だけ付ける。
  値は「前 → 後」の形で書くと後の値だけ 22pt 太字になる(△/- 始まりは赤、+/▲ 始まりは緑)。
  数字が主役でない並列には付けない(無い箱は None)。到達点・土台・数値帯・結論帯を全部載せると本文が入らない
  (警告が出る)ので、どれかを lead か注記に回す。
- `header_boxes(s, y, bot, [(柱, [文...]), ...])` は boxes(arrow=False) と同じ見た目の部品(他の部品と組むときだけ)。

---

## 図解の型

### 概念図(仕組み・全体像。各章の頭に 1 枚。見本 10 番)

**既定形**: 左 4.9in の `panel` に番号付き凡例 ①〜⑤、右の `panel` に線画アイコンの構成図。
番号は図のラベルの左にインラインで付き(「❷ 評価」)、凡例と対応させる。

```python
s = d.slide("○○の仕組み")
y0 = d.lead(s, "仕組みの結論(40 字以内)")
src = "出典"
ph = d.body_bot(src) - y0
bx, ly, bw, bh = d.panel(s, d.MX, y0, 4.9, ph, "構成要素")
d.numbered_list(s, bx, ly, bw, [("見出し", "説明は 1 行 23 字以内の完全文")] * 5, max_y=ly + bh)
rx, rw = d.MX + 5.3, d.CW - 5.3
_, ry, rw_in, rh_in = d.panel(s, rx, y0, rw, ph, "構成図", sub="補足")
xs, (r1, r2) = d.concept_grid(rx, ry, rw, rh_in, cols=4, rows=2)   # 列の中心 x と 2 行の y
c1, c2, c3, c4 = xs
d.zone_cols(s, rx, ry, rw, xs, 0, 1, r1, "境界の名前")                 # 列 0〜1 を囲む破線枠
d.node_icon(s, c1, r1, "doc", "入力", "補足", badge=1)
d.arrow_h(s, c1 + 0.45, c2 - 0.45, r1, label="渡すもの")
d.node_icon(s, c2, r1, "db", "蓄積", "補足", badge=2)
d.arrow_v(s, c2, r1 + 0.98, r2 - 0.42)
d.node_icon(s, c2, r2, "laptop", "出力", "補足", badge=3)
d.arrow_h(s, c2 - 0.45, c1 + 0.45, r2, label="戻り", dash=True)       # 戻り・フィードバックは破線
d.source(s, src)
```

- 凡例は 5 項目まで、説明は 1 行(23 字)の完全文(〜する)。`max_y` を渡すと項目間を詰め、はみ出せば警告。
- **図のラベルの語を、凡例の見出しか説明文に入れる**(凡例「監視」↔ 図「トレース」なら、説明に「トレースに記録する」)。
- アイコンは **BLUE 1 色**(NAVY は番号バッジ専用。渡すと置き換えて警告)。意味は見本 16 番に従う
  (利用者=person、画面・アプリ=laptop、LLM・対話=chat、蓄積=db、文書=doc、検査・防御=shield、集計・評価=chart、
  循環=cycle)。同じ対象には同じアイコンを使う。
- 主の流れは実線、戻り・フィードバックは `dash=True` の破線(1 枚に 2 本まで)。構造がループなら矢印を往復で描く。
- 座標は `concept_grid`(列 3〜4・行 1〜2)で決め、手計算の比率を書かない。3 列なら中心間隔 2.4in 以上、
  4 列は `lab_w=1.7`。ラベルは 6 字・補足は 8 字以内(はみ出すと警告)。

### 部品カタログ

| 部品 | 用途 |
|---|---|
| `slide(title)` / `lead(s, text)` / `body_bot(source, note, takeaway)` / `source(s, source, note)` / `takeaway_band(s, text, source, note)` | 型に無いスライドを組む土台。本文の上端は lead の戻り値、下端は body_bot |
| `panel(s,x,y,w,h,title,sub=,tone=)` | 領域を囲む既定の部品(BLUE 帯+白文字+GREY 細枠)。戻り値は内側 (x,y,w,h) |
| `numbered_list(s,x,y,w,items,size=12,gap=0.12,max_y=)` | ①〜⑤ の見出し 14pt+説明 12pt。戻り値は下端 y |
| `concept_grid(rx,ry,rw,rh,cols,rows=2)` | 右 panel の内側を格子に切り、ノード中心の (xs, ys) を返す |
| `zone_cols(s,rx,ry,rw,xs,c0,c1,row_y,label=)` | 列 c0〜c1 のノードを囲む破線枠(辺は隣の列との中点) |
| `node_icon(s,cx,cy,kind,label,sub=,badge=,lab_w=)` | 線画アイコン(BLUE)+ラベル 12pt 太字+補足 12pt。badge はラベルの左 |
| `arrow_h(s,x1,x2,y,label=,below=,dash=)` / `arrow_v(s,x,y1,y2,label=,dash=)` | 線+▶ の矢印(左向き・上向きも可)。ラベル 12pt。dash=True は戻りの流れ |
| `dashed_zone(s,x,y,w,h,label=)` | 破線の境界(環境・組織の範囲)。列で囲むなら zone_cols |
| `badge(s,cx,cy,n)` | 番号バッジ(12pt) |
| `header_boxes(s,y,bot,pillars)` / `header_box(...)` | BLUE 見出し帯+白本体の柱(並べる / 1 本) |
| `icon(s,kind,x,y,size,color,style="line")` | 13 種: person laptop chat doc db search gear shield cloud chart book cycle network(見本 16 番)。線幅 2pt |
| `step_flow(s,x,y,w,h,steps,intro=)` | 番号付き手順。太い上帯+番号+見出し 18pt+説明 14pt(2 行以内。切る位置は \n)。h は 2.0 以上。**手順そのものが主役のときだけ**。intro と 5 段は同時に置かない(説明が 1 行 7 字になる)。下段は section_block(見本 13 番) |
| `light_table(s,x,y,w,col_heads,rows,bold_rows=)` | 塗りのない軽い表(関門表の中身)。行高はセルの折り返しに合わせる。戻り値は下端 y |
| `section_block(s,x,y,w,max_y,heading,items)` | 見出し+下線の軽い段組(下段用)。items は gate_table と同じ。見出しに段の番号を入れる |
| `inline_steps(s,x,y,w,steps,trailing=)` | 短い手順を「1. 記録　2. …」と 1 段落に |
| `stage_flow(s,x,y,w,h,stages)` | 状態の推移(done / now / dead(×) / plan(破線))。関門・条件は gate_table |
| `col_label(s,x,y,w,text,sub=,align=)` | 見出し+下線。上に表や図がある下段の区切り、chart_pair の列見出し |
| `node()` / `edge()` | 丸ノードと線のミニグラフ(関係の網) |
| `card(..., dash=True)` | 破線の枠だけ。塗りカードを横に並べない |
| `tag_row()` / `zone()` | 非推奨(チップの羅列・淡色ゾーンは AI くさい)。互換のため残す |
| `tri_right` / `tri_down` / `_line_seg` / `_dash_line` | ▶ ▼ と任意の線分・破線 |

### flow — 役割分担・受け渡し(見本なし。boxes(arrows=) と同じ。見本 8 番の箱に矢印ラベルと土台が付く)
```python
d.flow(title, nodes, arrows=None, lead=None, band=None, source=None, note=None)
# nodes: [(見出し, 小見出し|None, [項目...]), ...] 2〜3 個 / arrows: 矢印上のラベル(12 字以内・14pt)
# band: 下に敷く共通の土台の一文(他の型の結論帯と同じ高さ)
```
- 矢印に意味がある関係は arrows でラベルを付ける(boxes と同じ箱。型としては 1 つ)。
  本文は全箱共通で 16→13pt。項目は 13〜19 字で、文末の形(〜する)を揃える。

### reason_tree — 結論と根拠(見本 11 番。tree(direction="down") と同じ)
```python
d.reason_tree(title, conclusion, reasons, lead=None, source=None, note=None)
# reasons: [(見出し, 本文 または ["箇条", ...], 根拠の一文|None), ...] 2〜3 個
# 同じものを d.tree(title, root, branches=[(見出し, [本文...], 根拠), ...], direction="down") でも呼べる
```
- lead と takeaway は渡さない(上の結論帯が兼ねる。takeaway は警告)。
- 結論帯は他の型の結論帯と同じ淡青+青太字(1 行なら 18pt、2 行なら 16pt。3 行は警告)。
- 本文は全カード共通で 16→15→14pt。箇条は 1 行 14 字前後が目安で、長い文は語の切れ目で自動改行する
  (3 行になるなら警告)。根拠には数字を入れる。

### gate_table — 関門・条件の表(見本 12 番)
```python
d.gate_table(title, gates, rows, lead=None, footnote=None, sections=None, source=None,
             note=None, bold_row=None, takeaway=None, on_fail=None)
# gates: 列見出し ["① 日本語の精度", "② 更新による劣化", "③ 権限外の再生"](2〜4)
# rows: [(行見出し, [関門ごとの値...]), ...] 例 何で測る / 合格ライン / 判定者(2〜4 行)
# footnote: 表の直下の ※ 注(判定の条件・通過後の流れ)
# on_fail: 不合格のときの扱い。文字列なら ※注の末尾に足す / リストなら 4 行目「不合格なら」。無いと警告
# sections: 下段の軽い段組 [(見出し, [items]), ...] 1〜2 個
#   items: "文"(● 箇条) / ("steps", ["記録", "切り分け", ...], "補足の一文") / ("p", "平文")
```
- 関門・合格基準・提供条件は、色箱と矢印の関門図ではなく **塗りのない軽い表**(見出し行の下に青罫線、行間は細線)。
  「合格ライン」の行は太字。**表が主役**: 既定は表+※注(+結論帯)だけで 1 枚(行高が自動で広がる)。
  sections は関門を通った直後に起きることで、他の枚(process など)に無いものだけ。次頁への案内や、別の枚と重複する
  手順は置かない。※注は 1 行に収める(長くなるなら on_fail の行か結論帯へ)。
- 書き方: 関門の見出しは「何を確かめるか」の名詞句で揃える(社内略語を使わない)。判定者には責任を負う人か
  会議体を 1 つだけ書き、LLM や判定モデルは「何で測る」の行に入れる(道具が判定者にあると警告)。
  合格ラインは「以上 / 以下 / 未満」で向きを書く。セルは 1 行 15 字が目安(3 行になると警告)。
- lead は表と下段の両方を 1 文で言う(「3 つの関門を通してから出し、出した後は 1 営業日で止める」)。
- 下段は `section_block`(見出し+下線)で軽く。短い手順は `("steps", [...])` で「1. 記録　2. …」と 1 段落に並べる。
- `stage_flow` は状態の推移(done / now / dead / plan)を見せるときだけ残す。

### cycle_flow — サイクル図(見本なし)
```python
d.cycle_flow(title, steps, lead=None, center=None, source=None, note=None, takeaway=None)
# steps: [(見出し 8 字以内, 説明 1〜2 行), ...] 3〜6(1 → n の順に周回)   center: 省略可の短い語(文字だけ)
```
- 改善の輪・PDCA・運用ループなど、lead で「輪・回す」と言う内容。**2 段の周回路**: 上段を左→右、下段を右→左に
  工程カード(白地+細枠。番号丸+見出し+説明)を並べ、間は ▶、左右の端は淡青の弧で上下をつなぐ
  (3 工程は 1 段+下辺の戻り、4 は 2+2、5 は 3+2、6 は 3+3)。
- 説明はカードの中(12〜14pt・2 行まで)。見出しの言い換えでなく、成果物・判断・数字を書く。
- 中央(`center=`)は既定で置かない(目的は lead が言う。図形も置かない — ユーザー選定 2026-10-25)。lead に無い情報
  (「毎週 1 周」のような回し方)を足したいときだけ短い語を渡すと、段の間に文字だけで入る。
- 旧 `detail_title=` は受け取るだけで使わない(渡すと警告)。

### layers — レイヤー図(見本 14 番)
```python
d.layers(title, layers, lead=None, side=None, source=None, note=None, takeaway=None)
# layers: [(層の名前, [要素, ...]), ...] 2〜5(上→下)。要素は "名前" か ("名前", "hl"|"plan")
# side: 全層をまたぐ縦の帯の名前(例「監視・トレース」)
```
- 何がどの層にあるかの「置き場」を見せる(データの流れは概念図)。要素は各層 2〜5 個、文字は 14pt。
- 主役は hl(BLUE 塗り)、計画は plan(破線)。意味は note に書く(「濃い箱=主題、破線=計画」)。
- side の帯は横書き(幅 1.9in)。名前は 8 字以内。

### swimlane — スイムレーン(見本 15 番)
```python
d.swimlane(title, lanes, stages, cells, lead=None, source=None, note=None, takeaway=None)
# lanes: 担当 2〜4(行)   stages: 工程 3〜6(列)
# cells: [(担当 index, 工程 index, 文, "hl"?), ...] を流れの順に。1 枡 1 箱
```
- 誰が・どの段で・何をするかを 1 枚に。連続する箱は自動でつなぐ(同じ担当で右へ=実線、担当をまたぐ=カギ線、
  左へ戻る=破線)。判定点は hl。役割の説明だけなら flow、工程だけなら step_flow。

### tree — ロジックツリー(縦=見本 11 番。横は見本なし)
```python
d.tree(title, root, branches, lead=None, source=None, note=None, takeaway=None, highlight=None,
       direction="right", levels=None)
# root: 根の文(2 行まで)か (論点, 補足)   branches: [(枝の見出し, [葉...]), ...] 2〜5。葉は各 4 まで(全体 12 まで)
#   枝は (見出し, 数値, [葉...]) でもよい(数値は枝の箱に太字で出る。△/- 始まりは赤)
#   葉は "文" か (項目名, 数値・説明)(横ツリーでは「項目名 | 数値」の表の行になる)
# direction="right": 横(目標 ← 要因 ← 具体策、課題の分解) / "down": 縦(結論 1 つと根拠 2〜3 のカード = reason_tree)
# highlight: 強調する枝 index のリスト。(枝 index, 葉 index) で葉だけの強調もできる(lead で名指しするものと合わせる)
# levels: (根の列見出し, 枝の列見出し, 葉の列見出し)。例 ("論点", "要因(影響額)", "具体的事象(数値)")。省略時は描かない
```
- 横ツリーの見た目: 根は濃紺の直角の箱、枝は淡い箱+左の青バー+数値、葉は罫線で区切った表の行。接続線は GREY の細線。
  枝の間に「+」「×」の記号は置かない(ユーザー選定 2026-10-25)。足し算か掛け算かは lead か levels の列見出しで言う。
- 目標 ← 要因 ← 具体策、課題 → 分解、方針 → 施策。枝と葉には数字(寄与・影響額)を入れる。数値の無い分解は content で書く。
- 結論と根拠 2〜3 を縦に並べるなら reason_tree、分解を横に広げるなら tree。

### HTML 図解 — HTML/CSS で組み、編集できる図形にする(scripts/html_shapes.py)
```python
from html_shapes import html_diagram, place_html
html_diagram(d, title, body_html, css="", lead=None, source=None, note=None, takeaway=None, keep="figs/p6.html")
place_html(d, s, body_html, x, y, w, h, css="")      # 部品として他の型と同じ枚に置く
```
- 型に無い自由な図(3 案の構造の比較、社内外の境界、判定のふるい、担当 × 工程の格子、計算の段など)に使う。
  ヘッドレス Edge で本文領域と同じ大きさ(inch)に描画し、位置・色・枠・文字の行を測ってネイティブ図形に置き直す。
  文字の折り返し位置はブラウザの結果をそのまま再現する(PowerPoint 側で折り返さないよう幅に余裕を持たせる)。
- 書き方: 文字は pt(本文 13.5〜14pt、補足 12pt が下限)、配置は flex / grid / position:absolute(in 単位)。
  背景色か枠のある要素 → 矩形(border-radius 50% は楕円、片側の枠はその辺の線、dashed は破線)。文字を持つブロック → テキストボックス。
  `<i class="ico" data-icon="db">` → 線画アイコン。`<x-wire from="#a" to="#b" route="h|v|hv|vh" label="…" dash head="both">`
  → コネクタ矢印(折れ線はフリーフォーム)。端は "#id:r:0.3" で辺と位置を指定できる。::before の文字は拾わない。
- 既定クラス: `.box` `.hl`(BLUE 塗り)`.plan`(破線)`.head`(BLUE 帯)`.zone`(破線の境界)`.badge` `.lbl` `.sub` `.cap`。
  色は CSS 変数(var(--blue) など。corp_pptx のパレットと同じ値)。
- 1 文字だけの行は `<br>` で切る位置を決めて直す。12pt 未満・図の外へのはみ出し・箱からのはみ出しは save() の注意に出る。
- keep= に .html を残す(ブラウザで開くと矢印まで描いた見た目になる。直したら再生成)。

### figure — 画像(最後の手段)
```python
d.figure(title, image, lead=None, source=None, note=None)
```
- 写真や複雑な曲線など、図形に置き直せない絵だけ。編集できる図は上の HTML 図解で作る。
  `python scripts/html2png.py fig.html --size 1600x680 --scale 2` で PNG 化。.html を必ず残す。

---

## 数値の型

### chart / chart_pair — グラフ(見本 17〜20 番)
```python
d.chart(title, kind, categories, series, lead=None, unit=None, value_fmt="#,##0",
        source=None, subtitle=None, highlight=None, bracket=None, callouts=None,
        ymin=None, ymax=None, target=None, takeaway=None, note=None,
        value_axis="auto", totals=True, direct_labels=True, forecast_from=None, share_of=None)
# kind: "bar" | "stack" | "line" | "waterfall"
# highlight: 主役の棒 index(単系列 bar)  bracket: (i0, i1, "範囲の名前と根拠の数字")(bar だけ)
# callouts: [(index, "一文"), ...] 2〜3 個  target: (値, "目標 90% 以上") 破線(ラベルに向きを書く)
# value_axis: "auto"=値ラベルが全部ある図(単系列 bar・合計付き stack)では値軸と目盛線を消す
# stack: totals=合計を棒の上に / direct_labels=系列名を最後の棒の右に(凡例なし) / share_of=i で区分 i に「値(xx%)」
#        forecast_from=i で i 以降を「計画」として縦の破線で区切る
d.chart_pair(title, specs, lead=None, source=None, note=None, takeaway=None)
# specs: [dict(head="Q: 問い", subtitle=, kind=, categories=, series=, unit=, ymin=, ymax=,
#              highlight=, bracket=, callouts=, target=), ...] 2 つ。head は必須。値軸は出す
```
- **全幅**。右の「ポイント」欄は廃止(`points=`/`sowhat=` は TypeError)。定義行(subtitle)に ▍は付かない。
- lead の主張には、図の上に根拠の数字(差・比・平均)を 1 つ置く(bracket か callout)。lead の言い換えにしない。
- 率の折れ線は `ymin`/`ymax` を明示(例 60〜100)。目標がある率は `target=` を付け、ラベルに「以上 / 以下」。
- stack の lead は合計か一番下の区分について書く(2 つの区分を論じるなら chart_pair か 2 枚)。
  callout は区間ラベルにある数字を繰り返さない。
- waterfall の lead は期首→期末の純増減と主な要因を言う。少ないほど良い指標(件数・コスト)では、
  改善の要因に callout で評価を書く。
- 詳細は charts-tables.md。

### table — 表(見本 22・23 番)
```python
d.table(title, headers, rows, lead=None, col_widths=None, first_col_header=True,
        font_size=13, unit=None, source=None, highlight_row=None, note=None, takeaway=None, zebra=True)
```
- `sections=[(見出し, [items]), ...]` で表の下に軽い段組を置ける(gate_table と同じ部品)。表に入れると列が立たない
  別の群(「版は持たない 2 つ」など)を分けるときに使い、lead は表と下段の両方を言う。
- **論証テーブル**(zebra=False): 「課題|現状|打ち手|担当」。課題は足りない状態を言う 12 字以内の短句、
  現状は観察した事実と数字だけ(評価語を書かない)、打ち手は動詞で終え到達値を添える、担当は右端の独立した列。
  核心の行を `highlight_row`。lead は表全体の結論を 1 文で(強調行が主役である理由を言う)。
- **数値の表**(kpi タイルの代わり。zebra=True): 「指標|測り方|しきい値|現状|差」。未達は `△` で赤字+`highlight_row`。
  しきい値には向き(以上 / 以下)を書く。
- 文字は 13〜14pt(12pt 未満は 12pt に上げて警告)。行は 7 行まで。セルの折り返しは 2 行以内
  (4 行以上になる見込みと高さ超過は警告)。
- 1 列目は PALE 塗り+太字。数値の列は過半が数値なら列ごと右寄せ。

### matrix_eval — 評価マトリクス(見本 24 番)
```python
d.matrix_eval(title, options, criteria, ratings, lead=None, verdicts=None,
              verdict_header="位置づけ", recommend=None, source=None, note=None, takeaway=None, legend=None)
# options: 候補(列・4 まで) criteria: 比較軸(行・6 まで) ratings[i][j]: 候補 i × 軸 j
# 値は "◎○△×" / 短文 / ["箇条", ...] を混在可  verdicts: 最下段の結論  recommend: 推奨列
# legend: 記号の凡例。None なら使った記号に合わせて自動(◎ 最適 ○ 良い △ 課題あり × 不可)
```
- 比較軸は ◎ が良い向きに書く(「コスト」でなく「コストの低さ」)。文字のセルと行見出しは 14pt。
- 全軸◎の候補を作らない。推奨案の △× を隠さない。採点が主観なら source に書く。

### matrix_2x2 — 2 軸(見本 25 番)
```python
d.matrix_2x2(title, x_axis, y_axis, items, lead=None, quadrants=None, emphasize=None,
             source=None, note=None, takeaway=None, highlight=None)
# x_axis/y_axis: (低, 高, 軸名)。横軸の端は括弧で定義を添えてよい("高(既存の検索基盤で作れる)")
# items: [(名前, x, y)] 0〜1・6 個まで  quadrants: [左下, 右下, 左上, 右上]
# emphasize: 強調する象限 index(1 つ。淡青に塗る)  highlight: 強調する項目名(省略時は強調象限の項目)
```
- 項目は正確な座標の**点+ラベル**(ピルにしない)。強調=BLUE の点+青太字、他=GREY の点+細字。
- 絞り込みの説明に使い、落ちた候補も残す。中心線の近く(±0.03)、ラベルの重なり、強調なしは警告。
- 象限名は名前か打ち手のどちらかに揃え、言いさしにしない。縦軸名は和文で短く(縦書きになる)。
  位置が目安なら note= に「位置は説明用の目安」と軸の定義を書く。

### goal_bars — ゴールに対する現在地(見本なし)
```python
d.goal_bars(title, items, lead=None, source=None, note=None, takeaway=None)
# items: [(指標, 現在値, 目標値, 単位, 起点値|None, 補足|None, "lower"?), ...] 2〜6
```
- 「目標に対してどこにいるか」を問われる報告(中間レビュー・進捗)の型。起点→現在→目標の物差しに、
  塗り=現在、濃紺の縦線=目標、右端に目標との差(達成 / 不足。未達は △ 赤)。補足に時点を書く。
- 低いほど良い指標は "lower"。物差しは同じ左→右で、目標が「上限」、現在値に「上限内 / 超過」、差は「余裕 / 超過」と
  出るので向きで迷わない。lower の行は最後にまとめる。
- 使い分け: 指標が 4 つ以下で目標との差を見せるなら goal_bars、5 つ以上で測り方としきい値まで並べるなら table(数値)。
  推移まで見せるなら chart(line/bar)+target。

### kpi — 廃止
- `d.kpi(...)` は TypeError。数値は table(数値の表)・gate_table・chart に意味とセットで置く。

---

## 計画の型

### timeline — 粗いフェーズ(見本 27 番)
```python
d.timeline(title, phases, lead=None, source=None, note=None, takeaway=None)
# phases: [(期間, 見出し, [説明...]), ...] 2〜4 個。説明は各 3〜4 個
```
- 説明は全フェーズ共通で 16→14pt、語の切れ目で自動改行。各フェーズの最後の行に出口条件
  (何ができたら次へ進むか)を書くと強い。期間に年度を入れる。
- 進捗の報告では `now=(フェーズ index, 0〜1 の位置, "現在地(10 月)")` で現在地の印(▼)を打つ。

### gantt — 詳細な行程(見本 28 番)
```python
d.gantt(title, columns, lanes, phases=None, milestones=None, lead=None, source=None,
        note=None, label_w=1.75, takeaway=None, col_weights=None, axis_label=None)
# columns: 時間軸(等しい単位。12 か月以下なら月)  axis_label: 年度か起点("2027年度"。無いと警告)
# phases: [(帯ラベル, 開始列, 終了列)]  lanes: [(レーン名, [(label, 開始列, 終了列, style), ...])] 5 本まで
# style: main(BLUE・確定) / option(白地+青破線・判断次第) / key(NAVY・主役の 1 本)
# milestones: [(列位置 float, "判断名(時期)")] — 縦の破線+下の ◆
# today: 現在の列位置(float)。進捗の報告では青の縦線(注に「青線=現在」が付く)
```
- 旧 style(solid→main / plan→option / dark→key。mid・light は main に寄せて警告)。意味は注に自動で付く。
- 束ね列+col_weights は最後の手段(警告)。バーのラベルはバーの幅に入る長さに(超えると警告)。
- レーンは同じ粒度の作業の流れか担当で切る。節目は判断点だけ(2〜3 個)。全作業の WBS は付録に回す。

### process — 壁→打ち手(見本 29 番)
```python
d.process(title, steps, lead=None, source=None, note=None, takeaway=None, loop=False)
# steps: [(工程の一文, (壁の一文, 打ち手の一文) | None), ...] 3〜6 個
# loop=True: 表の下に「工程 N の結果を工程 1 の変更に戻す」の 1 行(lead で「輪・回す」と言うなら付ける)
```
- **表の形**: 列は No | 工程 | 壁(詰まるところ) | 打ち手。塗りなし・横罫だけ(関門表と同じ軽さ)。
  壁の無い工程は「—」。打ち手は青太字。縦の番号丸をつないだステッパーや「壁 / 打ち手」のラベル対、
  小さな ▶ は UI 部品の見た目で AI くさいのでやめた(ユーザー評価 2026-10-24)。
- 全工程に壁を付けない(付けると警告)。セルは 2 行以内(3 行で警告)。
- 書き方: 壁は「なぜその工程が詰まるか」の原因(〜できない / 〜が分からない / 〜が古くなる)。
  判定結果(未達・NG)や起きうる結果(〜のおそれ)は壁にしない。打ち手は動詞の終止形で終え(警告あり)、
  頻度・期限・件数のどれかを 1 つ入れる。

---

## 標準構成

```
cover → exec_summary → toc → section(1) → 概念図 → 本文 ×n → section(2) → … → 推奨・次の一手 → message
```

- 本文は結論が先、根拠が後。章の 1 枚目に仕組みの図か結論の数値(グラフ・数値の表・関門表)。
- 進捗・中間レビューでは「ゴールに対する現在地」を 1 枚置く(goal_bars、または timeline の now= / gantt の today=)。
- save() が「図解・グラフの無い章」と「表が本文の半分超」を警告する。図解は編集できるネイティブ図形で組み、
  型に無い図は HTML 図解(html_diagram)で。HTML→PNG の figure は最後の手段。
- 中扉は本文 10 枚以上か 4 章以上のときに使い、各章の本文は 3 枚以上。満たさなければ目次 1 枚で済ませる。
- 結びの直前に推奨か次の一手(動詞・担当・期日)を置き、message はその結論と次の一手を繰り返す。

## 型選択ガイド

### まず構造で考える

| 構造 | 中身 | 主な型 |
|---|---|---|
| 仕組み | 構成・データの流れ・置き場・全体像 | 概念図 / layers / flow |
| 順序 | 時系列・手順・循環・担当の流れ | timeline / gantt / cycle_flow / swimlane / step_flow / process |
| 比較 | 前後・案の対比・候補選定 | boxes / matrix_eval |
| 分類 | 立ち位置・柱 | matrix_2x2 / boxes(arrow=False) / content(blocks) |
| 問題解決 | 課題→打ち手・課題の分解 | table(論証) / process / tree |
| ピラミッド | 結論+根拠・目標←要因 | reason_tree / tree / chart → 根拠スライド |
| 現在地 | 目標に対する進捗 | goal_bars / timeline(now=) / gantt(today=) |

### 内容 → 型の逆引き

| 内容がこれなら | 使う型 |
|---|---|
| 冒頭の要旨 | exec_summary |
| 仕組み・アプローチ・データの流れ | 概念図(panel+numbered_list+concept_grid+node_icon) |
| システム構成・階層(何がどの層にあるか) | layers |
| 循環する工程(改善の輪・PDCA) | cycle_flow |
| 担当 × 工程の流れ(誰が・いつ・何を) | swimlane |
| 目標 ← 要因 ← 具体策 / 課題の分解 | tree |
| 目標に対する現在地(進捗の報告) | goal_bars |
| 並列の事実・方針・条件・提供範囲 | content |
| 見出しの強い 4〜6 グループ(本文の種類がまちまち) | content(blocks) |
| グループごとに同じ属性が同じ順で並ぶ | table(対象 × 属性。zebra=False) |
| 状態の変化 | boxes(arrow=True) |
| 中身の種類が違う並列の柱 | boxes(arrow=False) |
| 到達点 ← 柱 ← 土台(柱が目標を支え、基盤の上に立つ) | boxes(goal=, base=) |
| 柱ごとの効果・実績を数字で見せる | boxes(figures=) |
| 柱ごとに同じ属性が並ぶ | gate_table か table(数値) |
| 役割分担・受け渡し | flow |
| 強み・選定理由 | reason_tree |
| 提供条件・関門・合否の基準 | gate_table |
| 重要な数値・目標と現状 | table(数値の表)/ gate_table |
| 推移・構成・増減 | chart(bar / stack / waterfall) |
| 2 つの問いを並べる | chart_pair |
| 課題と打ち手 | table(論証・zebra=False・担当列) |
| 候補を比べて選ぶ | matrix_eval |
| 絞り込み・立ち位置 | matrix_2x2 |
| 粗い計画 / 詳細な行程 | timeline / gantt |
| 手順と詰まりどころ / 改善の輪 | process(loop=True) |

- 同じ型が 2 枚続いたら、次は別の型にできないか考える。3 連続は避ける。
- 箱の中身が薄い(短文 1 行・下半分が空)なら、箱を減らすか項目を足す。箱には「役割+具体例・数字」。
