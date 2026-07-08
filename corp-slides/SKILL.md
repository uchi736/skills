---
name: corp-slides
description: >
  自社ブランドのプレゼン資料(パワーポイント/pptx)を作る・直すときに必ず使う。
  自社公式テンプレート(同梱)のマスターを土台に、Deck API が表紙・目次・中扉・箇条書き・
  対比ボックス・KPI・グラフ・表・ロードマップ・メッセージの10型を提供。本物の自社ロゴ・
  公式フッター・扉/最終ページの意匠はマスター継承で自動付与され、PowerPoint COM で
  実画像検証まで行う。トリガー: 自社, スライド, パワポ, pptx, プレゼン, 決算説明会, IR資料。
---

# 自社 スライド ハウススタイル

**土台は公式テンプレート**(`assets/brand/format-16x9.pptx` / `format-4x3.pptx` — 社内正規の
自社フォーマット。**git管理外**)。Deck API はこの公式マスターの上にスライドを組むため、
**自社ロゴ(ベクター)・表紙/扉の意匠・コピーライト・ページ番号は本物が自動で付く**。
このスキルはその上に、2026年IR資料の表現(游ゴシック・青基調・KPI/チャート/表の型)で
コンテンツを載せる。一般的なスライドの作り方はここには書かない。

## テンプレート解決とpush安全設計

- `Deck(aspect="16:9")`(既定)/ `Deck(aspect="4:3")` で判型を選ぶ。座標はすべて
  W/H比率で計算されるため、**どの型メソッドも両判型で同じAPI**。
- テンプレは `assets/brand/format-<判型>.pptx` を探し、**あれば公式マスター継承(本物)、
  無ければ中立ダミー**(「LOGO」表記・「© Your Company」・自前描画の同型意匠)で生成する。
- `assets/brand/` と生成物(*.pptx / *.png)は `.gitignore` 済み。**リポジトリに載るのは
  コードとドキュメントだけ**なので、このスキルはそのまま push してよい。
- 新しいマシンでは `assets/brand/` に公式ファイルを2つ置くだけで本物モードに戻る
  (4:3の公式テンプレを入手したら `assets/brand/format-4x3.pptx` に置く)。
- push後の見え方の確認: `Deck(brand=False)` または `python scripts/build_template.py 16:9 dummy`。
- ブランドテンプレ側の前提: レイアウト名「表紙/目次/扉/本文ページ/最終ページ」を持つこと。

## 鉄則(5つ)

1. **ゼロから作らない** — 必ず `scripts/corp_pptx.py` の `Deck` API で組む。座標・色・フォントを
   手置きしない。10型は [references/layouts.md](references/layouts.md) 参照。
2. **検証せずに納品しない** — 生成後、必ず `scripts/export_preview.py` で PowerPoint 実機レンダリング
   画像を出して**全ページ目視**し、[references/checklist.md](references/checklist.md) と照合する。
3. **パレット外の色を足さない** — 色は下表のみ。赤は負値専用、装飾に使わない。
4. **1スライド1メッセージ** — 各スライドに lead(結論1〜2行)。詰め込むより分割。
5. **プレースホルダを残さない** — `[ ]` 付きテキスト、`202X年`、ロゴ位置の確認まで済ませる。

## 作業手順

```
① 構成を決める   資料の目的 → 章立て → 各スライドに10型を割り当てる(layouts.md)
② スクリプトを書く  scratchpad に build_deck.py を作り Deck API で記述
③ 生成           PYTHONUTF8=1 で実行(パスは sys.path.insert でこのスキルの scripts/ を追加)
④ 検証           python scripts/export_preview.py out.pptx → contact sheet を Read で目視
⑤ 修正ループ      checklist.md の項目を満たすまで ②〜④ を繰り返す
⑥ 納品           ユーザーの指定場所(既定: Desktop)に .pptx をコピー
```

最小コード:

```python
import sys; sys.path.insert(0, os.path.expanduser("~/.claude/skills/corp-slides/scripts"))
from corp_pptx import Deck
d = Deck()
d.cover("タイトル", subtitle="サブ", label="決算説明会", date="2026年7月3日", dept="○○部")
d.toc(["背景", "提案", "効果"], current=1)
d.section(1, "背景")
d.content("課題", lead="結論。", bullets=["要点1", ("要点2", ["子A", "子B"])])
d.kpi("実績", kpis=[("売上", "1,234", "億円", "前年比+5%"), ("営業利益率", "+8.2", "%", None)])
d.chart("売上推移", "bar", ["FY24","FY25","FY26予"], [("売上", [1350,1420,1500])],
        unit="億円", points=["3期連続増収", "FY26は過去最高を計画"])
d.table("セグメント別", ["セグメント","FY25","FY26計画","増減"],
        [["航空・宇宙","4,120","4,400","+280"], ["資源・エネルギー","3,890","3,700","△190"]],
        unit="億円", col_widths=[3,2,2,2])
d.timeline("ロードマップ", [("FY2026","基盤整備",["施策1"]), ("FY2027〜","全社展開",["施策2"])])
d.message("技術と人財で、次の成長へ", contact=["お問い合わせ:○○部"])
d.save("out.pptx")
```

チャート・表の詳細規定(系列色順・データラベル・単位・△赤字など)は
[references/charts-tables.md](references/charts-tables.md)。

## カラーパレット(実測値。この表が正)

| 定数 | HEX | 用途 |
|---|---|---|
| NAVY | `#01357A` | 最強の強調・表紙帯上端・ロゴ |
| BLUE | `#004EA2` | 見出し帯・箱ヘッダ・アンダーライン(主役) |
| BLUE_TX | `#00549A` | 見出しテキスト |
| STEEL | `#3377AE` | 目次番号・補助 |
| CYAN | `#00BCEB` | アクセントバー・矢印 |
| CYANSUB | `#66C5DC` | サブ見出し帯 |
| TEAL | `#00BB99` | ポジティブ数値(少量) |
| RED | `#C00000` | 負値(△)専用 |
| BG / LTBLUE / PERI | `#F0F5F8` / `#D6ECF7` / `#C5D0E6` | 淡背景・ハイライト・表紙帯 |
| INK / INK2 / GREY | `#262626` / `#4D4D4D` / `#A6A6A6` | 本文・注釈・罫線(純黒禁止) |

フォント: **游ゴシック**(見出しBold/本文Regular)のみ。ロゴプレースホルダだけ Times New Roman。

## 階層マーカー(全型共通・変更禁止)

箇条書きの階層は、資料全体で必ずこの3段ラダーに統一する(定義は `corp_pptx.py` の `bullet()` が唯一):

```
● 第1階層   (●=青 #004EA2、本文 15pt 基準)
　□ 第2階層  (□=スチールブルー、-2pt)
　　‐ 第3階層 (‐=スチールブルー、-3pt、本文は補足色 #4D4D4D)
```

- content/boxes/timeline/チャートのポイント欄すべて同じラダー(サイズだけ文脈で縮小)。
- ネストは `("親", ["子", ("子2", ["孫"])])` の形で渡せば自動でこの記法になる。
- **・、✓、◆、►、数字マーカーなどを勝手に混ぜない**。第4階層が要る時点で構成を見直す。

## よくある失敗(ここが一番効く)

- 紫グラデ・原色・ネオン・絵文字・影付き文字 → **使わない**。自社は端正な青基調。
- 階層マーカーを勝手に変える(・や✓を混ぜる、階層ごとの記号が資料内でぶれる)
  → 上記ラダー以外は使わない。`bullet()`/`nest_paras()` を通せば統一される。
- 本文系スライドを白紙レイアウトで作る → 本文ページレイアウト(Deck が自動選択)以外では
  ロゴ・タイトル罫線・フッターが付かない。**必ず Deck の型メソッドで作る**。
- 本文を純黒で打つ / フォント混在(明朝・メイリオ) → INK(#262626)+游ゴシックに統一。
- グラフに単位なし・桁区切りなし・「予」表記なし → IR資料として通らない。
- 2軸グラフ・円グラフ多用 → 作らない(誤読の元)。構成比は stack か表で。
- プレビューを見ずに文字あふれのまま納品 → export_preview.py での目視は省略不可。
- 公式マスターのロゴ・帯・フッターを削除/改変する → 触らない(マスター継承のまま)。
- 4:3で作る → 判型は公式の 16:9(14.347×8.071 inch)。座標は Deck の定数
  (`MX`/`CW`/`BODY_TOP`/`BODY_BOT`)基準で書く。

## ファイル構成

```
.gitignore                     brand/ と生成物を除外(push安全の要)
scripts/corp_pptx.py            Deck API 本体(色・型・座標・ダミー意匠の単一ソース)
scripts/export_preview.py      PowerPoint COM 実機レンダリング検証
scripts/build_template.py      見本デッキを再生成(16:9 / 4:3 / 強制ダミー)
assets/brand/format-16x9.pptx  公式自社フォーマット 16:9(git管理外・編集禁止)
assets/brand/format-4x3.pptx   公式自社フォーマット 4:3(入手したらここに置く)
assets/template-*.pptx     10型の見本(生成物・git管理外)
assets/template-preview-*.png  実機レンダリング一覧(生成物・git管理外)
references/layouts.md          型カタログと使い分け・上限値
references/charts-tables.md    チャート・表の詳細規定
references/checklist.md        納品前チェックリスト
```

公式マスター継承により、ロゴ・表紙/扉/最終ページ・
コピーライト(公式文言)・ページ番号は**本物**が自動で付く。テーマの和文フォントは
ロード時に ＭＳ Ｐゴシック → 游ゴシック へ自動パッチされる(2026年IR実務に合わせた現代化)。
