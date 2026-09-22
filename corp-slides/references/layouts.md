# レイアウトカタログ(全17型)

すべて `scripts/corp_pptx.py` の `Deck` メソッド。**この17型の組み合わせだけでデッキを構成する**。
各メソッドは公式テンプレの対応レイアウト(表紙/目次/扉/本文ページ/最終ページ)にスライドを
追加するため、ロゴ・タイトル罫線・フッター・ページ番号は自動で付く。型に合わない要求が来たら、
まず最も近い型に内容を再分解できないか考える。それでも無理な場合のみ `Deck` の低レベルAPI
(`rect`/`text`)で自作し、その場合も `_slide(LAYOUT_BODY, title=...)` から始めて公式の
枠組みを維持する。キャンバスは 14.347×8.071 inch。本文の安全領域は
x: `MX`(0.87)〜13.47 / y: `BODY_TOP`(1.30)〜`BODY_BOT`(7.45)。

## 基本骨格と共通スロット(全型共通)

どのスライドも **「タイトル → ▍キーメッセージ(lead=) → 内容 → 軽い結論帯(takeaway=)」**
の縦一本に載る:

- `lead=` — ▍青バー+太字のキーメッセージ。必須(1スライド1つ)。
- `takeaway=` — 内容の下に敷く淡青の「軽い結論帯」。締めの一文が言えるスライドにだけ置く
  (content/boxes/kpi/chart/table/timeline/matrix系/gantt/process/reason_tree 対応。
  内容は自動で上に詰まり、注記・出典とも衝突しない)。flow は同等の `band=` を持つ。
- `note=` / `source=` — 「注:」「出典:」の2行フッター。
- chart系はさらに `subtitle=`(定義行)と `sowhat=`(構造化パネル)を持つ(⑦参照)。

各型とも `d.note("...")` を直後に呼ぶとスピーカーノートを付けられる(引数の note= とは別物)。

---

## ① cover — 表紙

```python
d.cover(title, subtitle=None, label=None, date=None, dept=None, company="[会社名]")
```
- **用途**: 1枚目専用。
- title は全角20字以内。長ければ subtitle に分割する。
- label は「決算説明会」「経営会議資料」などの資料種別。
- ページ番号なし(仕様)。

## ② toc — 目次

```python
d.toc(items, current=None, title="目次")
# items: ["章名"] または [("章名", "補足")]
```
- **用途**: 3章以上あるとき。2章以下なら目次は省略してよい。
- current(1始まり)で現在章をハイライト。**章の頭ごとに再掲する場合は current を進める**。
- 章は最大6個。7個以上は資料の構成を見直す。

## ③ section — 中扉

```python
d.section(number, title)
```
- **用途**: 章の切れ目。番号は目次と一致させる。

## ④ content — 箇条書き

```python
d.content(title, lead=None, bullets=None, source=None, style="ladder",
          note=None, takeaway=None)
# bullets: ["文"] または [("親", ["子1", ("子2", ["孫"])])] — 最大3階層
```
- **用途**: 主張+根拠の標準型。迷ったらこれ。
- 階層はハウスラダー **親■太字16pt → 子●14pt(全角2字下げ) → 孫□12ptグレー(全角4字下げ)**
  に自動統一(SKILL.md参照)。孫は「例:○○」の置き場。
- 収まらない時だけ 16→15→14pt に自動縮小(下限14)。それでも溢れる分量はスライドを分ける。
- `style="blocks"` で見出し段組(マーカーなし・罫線区切り・4グループ以上は自動2カラム)。
- lead は必ず書く。親は最大6個、子は親1つにつき最大3個。1項目2行以内。

## ⑤ boxes — ボックス対比(2〜3列)

```python
d.boxes(title, boxes, lead=None, arrow=False)
# boxes: [(見出し, 小見出し|None, [項目...]), ...] 2〜3個
```
- **用途**: 対比(Before/After、A案B案)、並列(3本柱)。
- arrow=True は「変化・遷移」のときだけ(対比には付けない)。
- 矢印の意匠は右向き三角形 ▶(`tri_right()`)。ブロック矢印には戻さない(ハウス好み)。
- 小見出し(sub)はボックス内の**タグ(角丸チップ)**として描画される。全幅の帯にはしない(ハウス好み)。
- 項目は各ボックス最大5個。列ごとの項目数はなるべく揃える。

## ⑥ kpi — 数値ハイライト(2〜4タイル)

```python
d.kpi(title, kpis, lead=None, source=None, emphasize=None, takeaway=None)
# kpis: [(ラベル, 値, 単位, 補足|None), ...]  emphasize: 主役タイルのindex(濃紺反転)
```
- **用途**: 決算ハイライト、目標値、実績サマリ。
- 値は文字列。**先頭 `+`/`▲` → 緑(TEAL)、`-`/`△`/`▼` → 赤** に自動色分け。中立は青。
- 値は6文字以内が最も映える(自動縮小はあるが7文字超は避ける)。
- 細かい数値の羅列は ⑧ table へ。

## ⑦ chart — ネイティブグラフ

```python
d.chart(title, kind, categories, series, lead=None, unit=None,
        points=None, value_fmt="#,##0", source=None, subtitle=None,
        highlight=None, bracket=None, sowhat=None, sowhat_intro=None,
        takeaway=None, note=None)
# kind: "bar"(集合縦棒) | "stack"(積上げ) | "line"(折れ線) | "waterfall"(滝)
# series: [(系列名, [値...]), ...] — waterfall は1系列で [期首, ±増減..., 期末(None可)]
```
- **用途**: 推移(bar/line)、構成比の推移(stack)、増減の要因分解(waterfall)。
- **強調の語彙(単色チャートの作法)**: `subtitle=` 定義行(何の図か+軸・単位) /
  `highlight=` 主役の棒だけ濃紺・他は淡青(単系列bar) / `bracket=(i0,i1,ラベル)` 範囲の括り注記 /
  `sowhat=[(ラベル,主張,説明|None),...]`+`sowhat_intro="読み方"` 右1/3の構造化So Whatパネル。
- 系列色は自動(BLUE→CYANSUB→TEAL→…)。**指定順=強調順**。主役の系列を最初に置く。
- 単系列barは自動でデータラベル(太字)が付く。stackは白ラベル。
- sowhat か points を必ず付ける。グラフだけ貼って読み手に解釈を委ねない。
- カテゴリは最大8個、系列は最大4個。超えるならデータを集約する。詳細は charts-tables.md。

## ⑧ table — ネイティブ表

```python
d.table(title, headers, rows, lead=None, col_widths=None,
        first_col_header=True, font_size=12, unit=None, source=None,
        highlight_row=None)
```
- **用途**: 正確な数値一覧、および**論証テーブル**(コンサル資料の主力型)。
- **論証テーブル**: 「課題|現状|打ち手」の3列で、行ごとに左→右で論理を完結させる。
  課題列に対策を書かない・現状列に意見を書かない(列の役割を混ぜない)。
  現状は事実と数字だけ、打ち手には「誰がやるか」を入れる。核心行は highlight_row で強調。
- 数値セルは自動右寄せ、`△`/`-` 始まりは自動赤字。
- `highlight_row=index`(0始まり)で推奨・注目行を淡青+太字で強調できる。
- 行は最大10行(超えるなら Excel 添付にし、スライドはサマリに)。
- col_widths は比率指定(例 `[3,2,2,2,2]`)。1列目(項目名)を広めに。

## ⑨ timeline — ロードマップ(2〜4フェーズ)

```python
d.timeline(title, phases, lead=None)
# phases: [(期間, 見出し, [説明...]), ...]
```
- **用途**: 中期計画、導入ステップ。
- 色は時系列で濃くなる(自動)。フェーズ見出しは10字以内、説明は各3個まで。

## ⑩ message — メッセージ / 結び

```python
d.message(text, sub=None, contact=None)
```
- **用途**: 資料の締め、または章をまたぐ強いメッセージの独立表示。
- text は24字以内で一文。**言い切る**(体言止め可、疑問形・弱気表現は避ける)。
- contact を渡すと結びスライドになる。

## ⑪ matrix_eval — 評価マトリクス(コンサル比較表)

```python
d.matrix_eval(title, options, criteria, ratings, lead=None,
              verdicts=None, verdict_header="位置づけ", recommend=None, source=None)
# options: 列(候補・最大4)  criteria: 行(比較軸)  ratings: 候補ごとの値リスト(ratings[i][j]=候補i×軸j)
# 値は "◎○△×" の記号 / 短文テキスト / ["箇条",...] のリストを行ごとに混在できる
# verdicts: 最下段の結論行  recommend: 強調する候補列 index(0始まり)
```
- **用途**: 候補の選定・比較・優先度付け。**選定・比較の議論は箇条書きでなくこれを使う**。
- **候補を列に立て、比較軸を行に並べる**(競合比較・ベンダー比較の定石形。ハウス好み)。
  推奨列はヘッダがNAVY・列全体が淡青ハイライトになる。
- 判断は ◎(青・最良)/○/△/× の4段記号、事実(数値・仕様・期間)は短文テキストの行にして
  混在させる。「概要」「概算投資」等の事実行+記号の評価行+「留意点」の箇条行、が定番構成。
- options は最大4列、criteria は6行程度まで。テキストセルは2行以内、箇条は各3個まで。
- verdicts で「だから何か」を必ず言う。採点が主観・目安の場合は source に根拠を明記する。
- **○×比較の信用を守る規律**: 全軸◎の候補を作らない(比較として信用されない)。
  結論に合わせて評価軸を後出ししない。推奨案の×や△を隠さず、カバー策を verdicts か lead で言う。

## ⑫ matrix_2x2 — 2軸マトリクス(ポジショニング)

```python
d.matrix_2x2(title, x_axis, y_axis, items, lead=None,
             quadrants=None, emphasize=None, source=None)
# x_axis/y_axis: (低ラベル, 高ラベル, 軸名)   items: [(名前, x, y)] x,yは0.0〜1.0
# quadrants: [左下, 右下, 左上, 右上] の象限ラベル   emphasize: 強調象限 index
```
- **用途**: 立ち位置の整理(候補・競合・技術の位置づけ)、多数の候補からの**絞り込みの説明**。
- 軸は「議論を分ける独立な2観点」を選ぶ。items は最大6個(重なるなら座標をずらす)。
- 絞り込みに使うときは、**落ちた候補も図の中に残す**(消すと「最初から決まっていた」と見える)。
- emphasize で「主戦場」を1象限だけ示す。全象限を強調しない。
- 座標が主観の場合は source に「位置は説明用の目安」と明記する。

## ⑬ flow — 関係図(ラベル付き矢印)

```python
d.flow(title, nodes, arrows=None, lead=None, band=None, source=None)
# nodes: [(見出し, 小見出し|None, [項目...]), ...] 2〜3個
# arrows: 矢印上のラベル(len(nodes)-1個)   band: 下に敷く共通基盤バーの文言
```
- **用途**: 役割分担・直列/連携・受け渡しの構造。**矢印に意味がある関係は boxes でなくこれ**。
- arrows のラベルは14字以内(2行まで自動折返し)。band は「両者に共通する土台」を一文で。
- 矢印の意匠は右向き三角形 ▶(`tri_right()`)。ブロック矢印には戻さない(ハウス好み)。
- 小見出し(sub)はボックス内の**タグ(角丸チップ)**として描画される。全幅の帯にはしない(ハウス好み)。

## ⑮ exec_summary — エグゼクティブサマリ

```python
d.exec_summary(title, conclusion, reasons, ask=None, lead=None, source=None, note=None)
# conclusion: 結論の一文(濃紺帯・白抜き)  reasons: [(見出し, 本文, 参照頁|None), ...] 2〜3個
# ask: (タグ, 本文, 補足|None) — 末尾の依頼・ネクストアクション帯。省略可
```
- **用途**: 冒頭の「本日の要旨」1枚。結論帯 → 理由01〜03カード → お願い帯の3段。
- lead 省略時は結論帯がキーメッセージを兼ねる。理由本文は2〜3行、参照頁(詳細 P.xx)を付けると強い。

## ⑯ reason_tree — 結論→根拠ツリー

```python
d.reason_tree(title, conclusion, reasons, lead=None, source=None, note=None)
# reasons: [(見出し, 本文, 根拠の一文|None), ...] 2〜3個
```
- **用途**: 強み・選定理由の構造化。結論帯から2〜3カードへ分岐、各カード下部に「根拠」チップ。
- 根拠には数字を入れる(例「有資格者120名が在籍」)。根拠の無いカードが混ざると全体が弱る。

## ⑰ gantt — スイムレーン・ガント

```python
d.gantt(title, columns, lanes, phases=None, milestones=None, lead=None,
        source=None, note=None, takeaway=None)
# columns: 時間軸ラベル(等幅)   phases: [(帯ラベル, 開始列, 終了列)](0始まり・両端含む)
# lanes: [(レーン名, [(バーlabel, 開始列, 終了列, style), ...]), ...]
#   style: "solid"=青 | "dark"=濃紺 | "mid"=スチール | "light"=淡青 | "plan"=白+破線(計画)
# milestones: [(列位置float, ラベル)] — グリッド下の●
```
- **用途**: 行程の詳細(フェーズ×ワークストリーム×節目)。粗いフェーズ提示だけなら ⑨ timeline。
- 破線=計画・未確定の意匠は全型共通。マイルストーンは判断ポイントに絞る(2〜3個)。

## ⑱ process — 壁→打ち手プロセス

```python
d.process(title, steps, lead=None, source=None, note=None, takeaway=None)
# steps: [(工程の一文, (壁の一文, 打ち手の一文) | None), ...] 3〜6個
```
- **用途**: 手順の説明を「どこで標準が破れ、何を当てるか」の論証に変える。導入計画の定番。
- 壁が無い工程は None のまま出す。全工程に壁を付けない(壁だらけ=計画になっていない)。

---

## 概念図(絵解き)とアイコン — ネイティブ部品で「編集可能な図解」を組む

**「○○:アプローチ」図解(この手法はどう動くか)を各章の頭に1枚置くのを既定とする。**
テキストの型だけで構成した資料は「図が1個もない」未完成品になりがち — 概念図は後付けの
飾りではなく設計図(手順①)の段階で割り当てる。部品はすべてネイティブ図形なので、
成果物は PowerPoint 上で全要素を編集できる。

### 部品カタログ(corp_pptx.py 実装済み)

| 部品 | 用途 |
|---|---|
| `step_flow(s,x,y,w,h,steps,intro=)` | 番号付き処理フロー ①→②→…。手法の動作説明の定番。intro=(ラベル,実例)で入力カードを先頭に |
| `stage_flow(s,x,y,w,h,stages,label=)` | 状態付き段階パイプライン。state: done=青 / now=現行 / dead=×破線(棄却。次矢印に「切替」自動) / plan=青破線 |
| `card(s,x,y,w,h,fill=,line=,dash=)` | 角丸カード(ヘアライン枠)。図解の基本コンテナ。dash=計画・棄却 |
| `zone(s,x,y,w,h,title)` | 見出し付きゾーン枠。「構築/検索」のような対構造に |
| `node()` / `edge()` | ノードリンク図(グラフ・関係)。dash=True が「欠けた要素・未接続」の表現 |
| `icon(s,kind,x,y,size,color)` | 幾何アイコン: book/cycle/network/search/db/doc/gear。淡青の円(OVAL)に載せると映える |
| `_tag(cx=で中央)` / `tag_row()` | 角丸チップ。前提・位置づけ・補足の列挙 |
| `tri_right` / `tri_down` / `_line_seg` | 矢印(▶/▼)と任意角度の線分 |

### アプローチ図解の定石(3段構成)

`_slide(LAYOUT_BODY)` + `_lead()` から始めて公式ヘッダを維持し、上・中・下の3段で組む:

```python
s = d._slide(LAYOUT_BODY, title="○○:アプローチ")
y0 = d._lead(s, "仕組みの結論を一文で。")
# 上段: 処理の流れ(どちらか)
d.step_flow(s, d.MX, y0+0.05, d.CW, 1.55,
            steps=[("抽出","説明"),("参照","説明"),("拡張","説明"),("検索","説明")],
            intro=("質問(例)", "「実物の例文」"))     # or d.stage_flow(... state付き段階)
# 中段: card×2〜3(実例・分担・注記) または zone×2 + node/edge のミニグラフ
# 下段: まとめ帯 card(fill=LTBLUE) や課題チップ card(dash=True)
```

- 座標感覚: 本文エリアは y0(≈1.9)〜BODY_BOT(7.45)、幅 CW≈12.6。上段1.5 / 中段2.5 / 下段0.5 が目安
- 文字サイズ: ゾーン見出し11 bold / カード内タイトル10〜11.5 bold / 本文9〜9.5 / 図中注記8
- その他の定番構図: 屋根+柱(役割分担)、積層+深さ軸(類型)、二段ネット+すり抜け(補完関係)

### 規律

- 絵解きは「仕組み・全体像を掴ませる」章頭の1枚が主戦場。論証は表とグラフで行い、連打しない
- **箱には「役割名+具体例(または数値)」の2層を詰める** — 大きな箱に短文一行だけの
  「AI図解」は禁止。中身が薄いなら箱を減らして密度を上げる
  (悪例: 横並び5箱に一行ずつ / 良例: エージェント⇄環境の往復矢印+具体例+目的式のループ図)
- 構造がループ・往復なら矢印の往復で描く。直列の step_flow に無理に開かない
- アイコン・色は単色(BLUE系/白)・幾何のみ。多色・立体・絵文字・外部画像は使わない

### HTML図解ルート(figure型) — 編集性を捨ててよい特殊ケースだけの代替

概念図の既定は上記ネイティブ部品(編集可能)。HTMLに逃がすのは、ネイティブで表現しづらい
ビジュアルを編集性を犠牲にしてでも入れたい時だけ:

1. `assets/fig-template.html` を複製して図を組む(ハウストークン・基本部品入り。既定 1600×680)
2. `python scripts/html2png.py fig.html --size 1600x680 --scale 2` で PNG 化(Edge headless)
3. `d.figure(title, png, lead=..., source=...)` で本文エリア中央に配置(比率維持)
- **.html がその図の編集可能なマスター**。デッキの生成スクリプトと同じ場所に必ず残し、
  修正は HTML を直して再レンダリングする(スライド上の画像は直接編集できない)
- ハウス規律は HTML 内でも同じ: 色はトークンのみ / 影・グラデ・多色禁止 / 游ゴシック /
  SVG アイコンは stroke=BLUE の線画(lucide 風)で統一
- 使い分け: 表・グラフ・比較表など「編集されうる型もの」はネイティブ継続。
  HTML に逃がすのは「概念図・絵解き」だけ(編集性より表現力を優先する場面)。

## 標準構成(推奨の並び)

```
cover → toc → section(1) → 本文(content/boxes/kpi/chart/table/timeline/matrix/flow)×n
      → section(2) → 本文×n → … → message(結び)
```

- 本文は「結論が先、根拠が後」。章の1枚目に kpi または chart で結論の数値を見せると強い。

## 型選択ガイド(ワンパターン防止)

### まず「構造」で考える(6分類)

型を選ぶ前に、言いたいことがどの構造かを決める。構造が決まれば型はほぼ決まる:

| 構造 | 中身 | 主な型 |
|---|---|---|
| 順序 | 時系列・プロセス・手順 | gantt / process / timeline / flow |
| 比較 | As-is/To-be・案の対比・競合 | boxes(arrow) / matrix_eval |
| 分類 | カテゴリ整理・立ち位置(MECE) | matrix_2x2 / table |
| 問題解決 | 課題→打ち手→効果 | flow / boxes |
| 因果 | 原因→結果・So What | flow / content(leadで示唆) |
| ピラミッド | 結論+根拠2〜3 | kpi・chart(章頭で結論の数値) → 根拠スライド |

### 設計4原則(全型共通)

1. **縦横** — 表・マトリクス・timelineは、縦軸と横軸が何を意味するかを必ず決めてから作る
2. **非冗長** — 同じ語を各行で繰り返さない。繰り返す語は軸・見出し・列名に括り出す
3. **二次元** — 1枚に3軸以上を詰めない。多ければスライドを分けるか左右分割する
4. **視覚ボリューム** — 数値の大小は長さ・面積の大小に反映する(chartが担う。表の数字の羅列で済ませない)

### 頻度の目安(コンサル実物資料の実測分布より)

一流の実物資料は「派手な図解」ではなく地味な型の使い倒しでできている。
ただし**各章の頭の概念図は既定**(SKILL.md 手順①d)であり、この頻度論は「連打するな」の意味で読む:

- **論証テーブルが主力(実測45%)** — 課題整理・施策一覧は、まず table を
  「課題|現状|打ち手」の3列で書く。行ごとに論理が完結し、行単位で議論できるのが強み。
- **概念図(絵解き)は章頭に1枚ずつ** — 仕組み・全体像を掴ませる場面専用。
  本文中で連打するとテンプレ感が出る。論証は表・グラフ・matrix に任せる。
- **フレームワーク(matrix_2x2等)は「絞り込みの瞬間」専用** — 見せびらかしに使わない。
- **チャートは必ず sowhat/points とセット** — 飾りで貼らない。
- **中扉は全体の1割** — 長い資料では惜しまない。迷子防止に効く。

### 内容パターン→型の逆引き

内容のパターンから型を選ぶ。**箇条書き(content)は最後の選択肢**:

| 内容がこれなら | 使う型 |
|---|---|
| 資料冒頭の要旨(結論+理由+お願い) | **exec_summary** |
| 強み・選定理由の構造 | **reason_tree** |
| 手法・仕組み・アプローチの動作説明 | **概念図**(step_flow / stage_flow / zone+node — 上の「概念図」参照) |
| 候補を選ぶ・比べる・優先度を付ける | **matrix_eval** |
| 複数の対象の立ち位置・すみ分けを示す | **matrix_2x2** |
| 役割分担・受け渡し・直列の関係 | **flow** |
| 状態の変化(Before/After) | boxes(arrow=True) |
| 独立した並列の柱・案 | boxes |
| 重要な数値そのもの | kpi |
| 推移・比較・構成比 | chart |
| 正確な数値一覧 | table(+highlight_row) |
| 時間軸の計画(粗いフェーズ) | timeline |
| 行程の詳細(フェーズ×担当×節目) | **gantt** |
| 手順+どこで詰まり何を当てるか | **process** |
| 上記のどれでもない主張+根拠 | content |

- **同じ型が2枚続いたら**、次の1枚は図解型(matrix/flow/2x2/chart/timeline)にできないか必ず検討する。
- 3枚以上の連続は禁止(特に content 連打)。
