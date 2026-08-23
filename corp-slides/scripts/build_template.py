# -*- coding: utf-8 -*-
"""
build_template — 全10レイアウトの型見本を再生成する(テンプレの単一ソース)。

    python build_template.py                # 16:9(brand解決) + 4:3(brand解決) + ダミー16:9 の3本
    python build_template.py 16:9           # 指定判型のみ(brand解決)
    python build_template.py 16:9 dummy     # 強制ダミー(push後の見え方の確認用)

出力: assets/template-<aspect>[-dummy].pptx + 実機プレビュー(すべてgit管理外)
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
from corp_pptx import Deck


def _build_scenes(d):
    # ① 表紙
    d.cover("「メインタイトルを入力」", subtitle="サブメッセージ",
            label="[決算説明会などのラベル]", date="202X年X月X日", dept="[部署名]")
    # ② 目次
    d.toc([("[章タイトル1]", None), ("[章タイトル2]", "[補足]"), ("[章タイトル3]", None)],
          current=1)
    # ③ 中扉
    d.section(1, "[セクションタイトル]")
    # ④ 本文・箇条書き(3階層 ●→□→‐)
    d.content("[スライドタイトル:箇条書き型]",
              lead="[リード文:このスライドで言いたい結論を1〜2行で]",
              bullets=[
                  "[要点1:1項目は2行以内に収める]",
                  ("[要点2:子項目を持てる]",
                   ["[子項目A]", ("[子項目B:さらに孫も持てる]", ["[孫項目:補足の補足]"])]),
                  "[要点3]",
              ],
              source="[出典が必要な場合に記載]")
    # ⑤ ボックス対比(矢印つき2列)
    d.boxes("[スライドタイトル:対比型]",
            lead="[2つの状態・観点を並べて対比する型。arrow=True で変化を表す]",
            boxes=[
                ("[導入前]", "[小見出し]", ["[項目1]", "[項目2]", "[項目3]"]),
                ("[導入後]", "[小見出し]", ["[項目1]", "[項目2]", "[項目3]"]),
            ], arrow=True)
    # ⑥ KPIハイライト
    d.kpi("[スライドタイトル:KPI型]",
          lead="[重要な数値を大きく見せる型。+/△の符号で自動的に緑/赤になる]",
          kpis=[
              ("[指標名]", "1,234", "億円", "[前年比+X%]"),
              ("[指標名]", "+12.3", "%", "[補足]"),
              ("[指標名]", "△45", "億円", "[補足]"),
          ])
    # ⑦ チャート(単系列棒+ポイント欄)
    d.chart("[スライドタイトル:グラフ型]", "bar",
            categories=["FY22", "FY23", "FY24", "FY25", "FY26予"],
            series=[("売上収益", [1180, 1290, 1350, 1420, 1500])],
            lead="[グラフで推移や比較を見せる型。右のポイント欄は省略可]",
            unit="億円",
            points=["[読み取るべきポイント1]", "[ポイント2]", "[ポイント3]"])
    # ⑧ 表
    d.table("[スライドタイトル:表型]",
            headers=["[項目]", "FY24実績", "FY25実績", "FY26計画", "増減"],
            rows=[
                ["[行見出し1]", "1,234", "1,350", "1,500", "+150"],
                ["[行見出し2]", "567", "590", "620", "+30"],
                ["[行見出し3]", "89", "76", "70", "△6"],
                ["[行見出し4]", "45", "52", "60", "+8"],
            ],
            lead="[数値の一覧を正確に見せる型。数値は自動で右寄せ・△は赤]",
            unit="億円",
            col_widths=[3, 2, 2, 2, 2])
    # ⑨ ロードマップ
    d.timeline("[スライドタイトル:ロードマップ型]",
               lead="[時間軸で計画を見せる型。フェーズは2〜4個]",
               phases=[
                   ("FY2026", "[フェーズ1見出し]", ["[施策1]", "[施策2]"]),
                   ("FY2027", "[フェーズ2見出し]", ["[施策1]", "[施策2]"]),
                   ("FY2028〜", "[フェーズ3見出し]", ["[施策1]", "[施策2]"]),
               ])
    # ⑪ 評価マトリクス
    d.matrix_eval("[スライドタイトル:評価マトリクス型]",
                  lead="[候補×基準を◎○△×で採点し、推奨行を強調する型。選定・比較の議論はこれ]",
                  options=["[候補A]", "[候補B]", "[候補C]"],
                  criteria=["[基準1]", "[基準2]", "[基準3]", "[基準4]"],
                  ratings=[["◎", "◎", "○", "△"],
                           ["○", "△", "◎", "○"],
                           ["△", "○", "○", "◎"]],
                  verdicts=["[推奨:理由を一言]", "[次点:理由]", "[見送り:理由]"],
                  recommend=0,
                  source="[採点根拠の資料名]")
    # ⑫ 2軸マトリクス
    d.matrix_2x2("[スライドタイトル:2軸マトリクス型]",
                 lead="[2つの独立な観点で立ち位置を整理する型。主戦場を1象限だけ強調]",
                 x_axis=("低", "高", "[横軸の名前]"),
                 y_axis=("低", "高", "[縦軸の名前]"),
                 items=[("[対象A]", 0.80, 0.30), ("[対象B]", 0.60, 0.55), ("[対象C]", 0.30, 0.80)],
                 quadrants=["[左下の意味]", "[右下の意味:主戦場]", "[左上の意味]", "[右上の意味]"],
                 emphasize=1)
    # ⑬ 関係図
    d.flow("[スライドタイトル:関係図型]",
           lead="[役割分担・受け渡しの構造。矢印に意味がある関係は boxes でなくこれ]",
           nodes=[
               ("[主体A]", "[Aの問い]", ["[Aの役割1]", "[Aの役割2]", "[Aの成果物]"]),
               ("[主体B]", "[Bの問い]", ["[Bの役割1]", "[Bの役割2]", "[Bの成果物]"]),
           ],
           arrows=["[矢印の意味(14字以内)]"],
           band="[両者に共通する土台を一文で]")
    # ⑩ メッセージ / 結び
    d.message("[一番伝えたいメッセージを一文で]",
              sub="[補足メッセージ(省略可)]",
              contact=["お問い合わせ:[部署名]", "[メールアドレス等]"])
    return d


def build(aspect="16:9", brand=True):
    d = _build_scenes(Deck(aspect=aspect, brand=brand))
    tag = aspect.replace(":", "x") + ("" if d._brand else "-dummy")
    out = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "assets",
                                       f"template-{tag}.pptx"))
    d.save(out)
    print(f"saved: {out}  [{aspect} / {'brand' if d._brand else 'dummy'}]")
    try:
        from export_preview import export_pngs, contact_sheet
        files = export_pngs(out)
        sheet = contact_sheet(files, os.path.join(os.path.dirname(out),
                                                  f"template-preview-{tag}.png"))
        print("preview:", sheet)
    except Exception as e:
        print("preview skipped:", e)


if __name__ == "__main__":
    aspect = sys.argv[1] if len(sys.argv) > 1 else None
    force_dummy = len(sys.argv) > 2 and sys.argv[2] == "dummy"
    if aspect:
        build(aspect, brand=not force_dummy)
    else:
        build("16:9")                 # brandがあれば本物
        build("4:3")                  # brandがあれば本物、無ければダミー
        build("16:9", brand=False)    # push後の見え方(強制ダミー)
