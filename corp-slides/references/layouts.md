# レイアウトカタログ(全10型)

すべて `scripts/corp_pptx.py` の `Deck` メソッド。**この10型の組み合わせだけでデッキを構成する**。
各メソッドは公式テンプレの対応レイアウト(表紙/目次/扉/本文ページ/最終ページ)にスライドを
追加するため、ロゴ・タイトル罫線・フッター・ページ番号は自動で付く。型に合わない要求が来たら、
まず最も近い型に内容を再分解できないか考える。それでも無理な場合のみ `Deck` の低レベルAPI
(`rect`/`text`)で自作し、その場合も `_slide(LAYOUT_BODY, title=...)` から始めて公式の
枠組みを維持する。キャンバスは 14.347×8.071 inch。本文の安全領域は
x: `MX`(0.87)〜13.47 / y: `BODY_TOP`(1.30)〜`BODY_BOT`(7.45)。

各型とも `d.note("...")` を直後に呼ぶとスピーカーノートを付けられる。

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
d.content(title, lead=None, bullets=None, source=None)
# bullets: ["文"] または [("親", ["子1", ("子2", ["孫"])])] — 最大3階層
```
- **用途**: 主張+根拠の標準型。迷ったらこれ。
- 階層マーカーは全型共通のラダー **●→□→‐** に自動統一される(SKILL.md参照)。
- lead は必ず書く(このスライドの結論1〜2行)。lead なしの箇条書きは避ける。
- 親は最大6個、子は親1つにつき最大3個。第3階層は本当に必要なときだけ。
- 1項目2行以内。あふれたらスライドを分ける。

## ⑤ boxes — ボックス対比(2〜3列)

```python
d.boxes(title, boxes, lead=None, arrow=False)
# boxes: [(見出し, 小見出し|None, [項目...]), ...] 2〜3個
```
- **用途**: 対比(Before/After、A案B案)、並列(3本柱)。
- arrow=True は「変化・遷移」のときだけ(対比には付けない)。
- 項目は各ボックス最大5個。列ごとの項目数はなるべく揃える。

## ⑥ kpi — 数値ハイライト(2〜4タイル)

```python
d.kpi(title, kpis, lead=None, source=None)
# kpis: [(ラベル, 値, 単位, 補足|None), ...]
```
- **用途**: 決算ハイライト、目標値、実績サマリ。
- 値は文字列。**先頭 `+`/`▲` → 緑(TEAL)、`-`/`△`/`▼` → 赤** に自動色分け。中立は青。
- 値は6文字以内が最も映える(自動縮小はあるが7文字超は避ける)。
- 細かい数値の羅列は ⑧ table へ。

## ⑦ chart — ネイティブグラフ

```python
d.chart(title, kind, categories, series, lead=None, unit=None,
        points=None, value_fmt="#,##0", source=None)
# kind: "bar"(集合縦棒) | "stack"(積上げ) | "line"(折れ線)
# series: [(系列名, [値...]), ...]
```
- **用途**: 推移(bar/line)、構成比の推移(stack)。
- 系列色は自動(BLUE→CYANSUB→TEAL→…)。**指定順=強調順**。主役の系列を最初に置く。
- 単系列barは自動でデータラベル(太字)が付く。stackは白ラベル。
- points(右のポイント欄)推奨。グラフだけ貼って読み手に解釈を委ねない。
- カテゴリは最大8個、系列は最大4個。超えるならデータを集約する。
- 詳細規定は charts-tables.md。

## ⑧ table — ネイティブ表

```python
d.table(title, headers, rows, lead=None, col_widths=None,
        first_col_header=True, font_size=12, unit=None, source=None)
```
- **用途**: 正確な数値一覧。
- 数値セルは自動右寄せ、`△`/`-` 始まりは自動赤字。
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

---

## 標準構成(推奨の並び)

```
cover → toc → section(1) → 本文(content/boxes/kpi/chart/table/timeline)×n
      → section(2) → 本文×n → … → message(結び)
```

- 本文は「結論が先、根拠が後」。章の1枚目に kpi または chart で結論の数値を見せると強い。
- 同じ型が3枚以上連続したら型を変えて緩急をつける(content 連打が最も単調)。
