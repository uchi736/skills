# -*- coding: utf-8 -*-
"""
corp_pptx — 自社ハウススタイルのプレゼン生成ライブラリ(corp-slides スキル同梱)v3

■ テンプレートの解決(push安全設計)
    assets/brand/format-16x9.pptx / format-4x3.pptx   ← 公式テンプレ(git管理外。ロゴ・意匠入り)
    → 存在すればそれを土台にする(本物のロゴ・フッター・扉デザインをマスター継承)
    → 無ければ中立なダミー(自前描画・「LOGO」表記・「© Your Company」)で同じ型を生成
    ブランド資産をリポジトリに含めなくても、コードだけで動作する。

■ 判型
    Deck()                # 16:9(既定)
    Deck(aspect="4:3")    # 4:3
    座標はすべて W/H 比率で計算されるため、両判型で同じAPIが使える。

使い方(最小):
    import os, sys; sys.path.insert(0, os.path.expanduser("~/.claude/skills/corp-slides/scripts"))
    from corp_pptx import Deck
    d = Deck()
    d.cover("メインタイトル", subtitle="サブ", label="決算説明会", date="2026年7月", dept="○○部")
    d.toc(["背景", "提案", "効果"], current=1)
    d.section(1, "背景")
    d.content("タイトル", lead="結論。", bullets=["要点1", ("親", ["子", ("子2", ["孫"])])])
    d.message("締めのメッセージ")
    d.save("out.pptx")

検証は export_preview.py(PowerPoint COM)で実画像を出して行うこと。
"""
import os, copy, re, warnings
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE, PP_PLACEHOLDER
from pptx.oxml.ns import qn
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION

# ============================================================ パレット(自社公開IR資料から実測)
NAVY   = "#01357A"  # コーポレート濃紺(最強の強調)
BLUE   = "#004EA2"  # メインブルー(箱ヘッダ・KPI値・チャート主系列)
BLUE_TX= "#00549A"  # 見出し系テキスト
STEEL  = "#3377AE"  # スチールブルー(目次番号・第2階層)
CYAN   = "#00BCEB"  # 明シアン(アクセントバー・矢印)
CYANSUB= "#66C5DC"  # サブ見出し帯
TEAL   = "#00BB99"  # ティール緑(ポジティブ数値・少量)
RED    = "#C00000"  # ネガティブ数値(△)専用。装飾に使わない
BG     = "#F0F5F8"  # 淡背景ボックス
LTBLUE = "#D6ECF7"  # 選択ハイライト(目次・強調行・結論帯)
PALE   = "#E8F2FA"  # 表の先頭列・評価マトリクスの位置づけ行(LTBLUE より淡い)
PERI   = "#C5D0E6"  # 帯
INK    = "#262626"  # 本文(純黒は使わない)
INK2   = "#4D4D4D"  # 補足・注釈
GREY   = "#A6A6A6"  # 罫線・弱い文字
GRID   = "#E3E7EC"  # チャートのグリッド線
YU     = "Yu Gothic"

SERIES_COLORS = [BLUE, CYANSUB, TEAL, STEEL, NAVY, "#66A6EA", PERI, CYAN]
STACK_COLORS = [BLUE, "#66A6EA", PERI, GREY]   # 積上げ専用: 主役が濃く上へ淡く、最後の「その他」は灰
RAMP = ["#66A6EA", "#2070C7", BLUE, NAVY]

_NUMERIC_RE = re.compile(r"^[+\-△▲]?[\d,.]+\s*(%|円|億円|百万円|人|件|pt)?$")

_ASSETS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "assets")
BRAND_FILES = {"16:9": os.path.join(_ASSETS, "brand", "format-16x9.pptx"),
               "4:3":  os.path.join(_ASSETS, "brand", "format-4x3.pptx")}
DUMMY_SIZES = {"16:9": (13.333, 7.5), "4:3": (10.0, 7.5)}

# ダミーモードの中立表記(push されるのはこの文字列だけ)
DUMMY_LOGO   = "LOGO"
DUMMY_FOOTER = "© Your Company All Rights Reserved."
BRAND_CFG_FILE = os.path.join(_ASSETS, "brand", "brand.json")


def load_brand_cfg():
    """assets/brand/brand.json(git管理外)があれば自社固有値を読む。無ければ中立既定。

    brand.json 例: {"company": "株式会社 ○○", "footer": "© ○○ ...", "label": "○○"}
    """
    import json
    try:
        with open(BRAND_CFG_FILE, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}

# 公式レイアウト名 → 用途(ブランドテンプレ側のレイアウト名と一致させること)
LAYOUT_COVER   = "表紙"
LAYOUT_TOC     = "目次"
LAYOUT_SECTION = "扉"
LAYOUT_BODY    = "本文ページ"
LAYOUT_LAST    = "最終ページ"


def _jlen(t):
    """全角 1・半角 0.5 で数えた文字幅(行数の見積もり用)。改行は 0。"""
    return sum(0.0 if c == "\n" else (1.0 if ord(c) > 0x2E80 else 0.5) for c in str(t))


CHAR_EM = 0.98   # 游ゴシックの実測: 全角 1 字 ≈ 0.93〜1.0em(混植で少し詰まる)。表のセルはほぼ 1em


def _cpl(w, pt, pad=0.0):
    """幅 w(inch)に pt の文字が何字入るか(全角換算・較正済み)。"""
    return max(int((w - pad) * 72 / (pt * CHAR_EM)), 6)


def _nlines(t, cpl):
    """文字列 t(\n 可)が cpl 字/行で何行になるか。"""
    return sum(max(1, -(-int(_jlen(seg) + 0.999) // cpl)) for seg in str(t).split("\n"))


_SENT_END_VERB = re.compile(r"(る|す|う|く|ぐ|つ|ぬ|ぶ|む|い|た|だ|ない|ます|です)$")


def _rgb(h):
    return RGBColor.from_string(h.lstrip("#"))

_ALIGN = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}
_ANCHOR = {"top": MSO_ANCHOR.TOP, "middle": MSO_ANCHOR.MIDDLE, "bottom": MSO_ANCHOR.BOTTOM}


def _set_font(run, size, bold, color, name=YU):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = name
    run.font.color.rgb = _rgb(color)
    rPr = run._r.get_or_add_rPr()
    rPr.set("lang", "ja-JP")          # 和文の禁則処理(行頭の句読点・拗音、行末の開き括弧)を効かせる
    rPr.set("altLang", "en-US")
    for tag in ("a:latin", "a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {})
            rPr.append(el)
        el.set("typeface", name)


def _patch_theme(prs):
    """テーマの和文フォントを游ゴシックへ(ブランド:MS Pゴシック置換 / ダミー:Calibri置換)"""
    for part in prs.part.package.iter_parts():
        pn = str(part.partname)
        if "theme" in pn and pn.endswith(".xml"):
            try:
                blob = part.blob.decode("utf-8")
                blob = blob.replace('typeface="ＭＳ Ｐゴシック"', f'typeface="{YU}"')
                blob = blob.replace('typeface="Calibri Light"', f'typeface="{YU}"')
                blob = blob.replace('typeface="Calibri"', f'typeface="{YU}"')
                blob = blob.replace('<a:ea typeface=""/>', f'<a:ea typeface="{YU}"/>')
                part._blob = blob.encode("utf-8")
            except Exception:
                pass
    # マスター・レイアウト(タイトル・フッター・ページ番号)に直書きされた書体も游ゴシックへ
    old = {"ＭＳ Ｐゴシック", "ＭＳ ゴシック", "MS PGothic", "MS Gothic", "Calibri", "Calibri Light"}
    for m in prs.slide_masters:
        for owner in [m] + list(m.slide_layouts):
            for el in owner._element.iter(qn("a:latin"), qn("a:ea"), qn("a:cs")):
                if el.get("typeface") in old:
                    el.set("typeface", YU)


class Deck:
    """ハウススタイルのデッキビルダー。1メソッド=1スライド(型)。

    aspect : "16:9"(既定) | "4:3"
    brand  : True(既定)=assets/brand のテンプレがあれば使う / False=強制ダミー
    """

    def __init__(self, aspect="16:9", brand=True, template=None):
        if aspect not in DUMMY_SIZES:
            raise ValueError('aspect は "16:9" か "4:3"')
        self.aspect = aspect
        path = template or (BRAND_FILES[aspect] if brand else None)
        self._brand = bool(path and os.path.exists(path))
        self.warnings = []      # 生成中の注意(はみ出し見込み等)。save() でまとめて表示
        if template and not self._brand:
            raise FileNotFoundError(f"template が見つからない: {template}")
        if brand and not self._brand:
            self._warn(f"公式テンプレ {path} が無いため中立ダミーで生成する(ロゴ・公式フッター無し)")
        if self._brand:
            self.prs = Presentation(path)
            self._clear_slides()
            self._layouts = {lay.name: lay for m in self.prs.slide_masters
                             for lay in m.slide_layouts}
        else:
            self.prs = Presentation()
            w, h = DUMMY_SIZES[aspect]
            self.prs.slide_width = Inches(w)
            self.prs.slide_height = Inches(h)
            self._blank = self.prs.slide_layouts[6]
        _patch_theme(self.prs)
        self.brand_cfg = load_brand_cfg()
        # ---- 判型に依存しない相対ジオメトリ(基準: 公式16:9 = 14.347x8.071)
        self.W = self.prs.slide_width / 914400
        self.H = self.prs.slide_height / 914400
        self.MX = 0.87 if self._brand else 0.061 * self.W
        self.CW = self.W - 2 * self.MX
        self.BODY_TOP = 0.161 * self.H
        self.BODY_BOT = self.H - 0.62
        self._page = 0
        self._takeaway_pages = []     # 結論帯を敷いたページ(連続・多すぎを save() で警告)
        self._toc_items = None        # toc() に渡した章名(中扉の一覧・章名の照合に使う)
        self._toc_page_boxes = []     # 目次の各章のページ番号欄 (章index, textbox)
        self._chapters = []           # [[章番号, 中扉ページ, 本文枚数], ...]
        self.body_font = self.brand_cfg.get("body_font")   # 例 "Yu Gothic Medium"(未指定なら游ゴシック Regular)
        self._page_kind = {}          # ページ → text / box / table / chart / diagram(章ごとの図解の有無を見る)

    # ---------------------------------------------------------------- 基盤
    def _kind(self, kind):
        self._page_kind[self._page] = kind

    def _warn(self, msg):
        """生成中の注意を記録する(save() で一覧表示。検証で必ず該当スライドを見る)。"""
        tag = f"[p{self._page}] {msg}" if getattr(self, "_page", 0) else msg
        self.warnings.append(tag)
        # 表示は save() でまとめて行う(二重に出さない)

    def _clear_slides(self):
        sldIdLst = self.prs.slides._sldIdLst
        for sldId in list(sldIdLst):
            rId = sldId.get(qn("r:id"))
            self.prs.part.drop_rel(rId)
            sldIdLst.remove(sldId)

    def _slide(self, layout_name, title=None):
        self._page += 1
        if layout_name == LAYOUT_BODY and self._chapters:
            self._chapters[-1][2] += 1
        if self._brand:
            lay = self._layouts[layout_name]
            s = self.prs.slides.add_slide(lay)
            # python-pptx はフッター/ページ番号プレースホルダを複製しないため手動継承
            for ph in lay.placeholders:
                if ph.placeholder_format.type in (PP_PLACEHOLDER.FOOTER,
                                                  PP_PLACEHOLDER.SLIDE_NUMBER):
                    el = copy.deepcopy(ph._element)
                    s.shapes._spTree.append(el)
                    cnv = el.find(".//" + qn("p:cNvPr"))   # id 重複は PowerPoint の修復ダイアログの元
                    if cnv is not None:
                        cnv.set("id", str(s.shapes._next_shape_id))
            if title is not None and s.shapes.title is not None:
                tf = s.shapes.title.text_frame
                tf.text = title
                for p in tf.paragraphs:          # テンプレの tx1(=#000000)を上書き。サイズ・位置は継承
                    for r in p.runs:
                        r.font.color.rgb = _rgb(INK)
            # 未使用の空プレースホルダを除去(編集画面の「テキストを入力」残り防止)
            for ph in list(s.placeholders):
                if ph.placeholder_format.type in (PP_PLACEHOLDER.FOOTER,
                                                  PP_PLACEHOLDER.SLIDE_NUMBER):
                    continue
                if ph.has_text_frame and ph.text_frame.text.strip():
                    continue
                ph._element.getparent().remove(ph._element)
            return s
        # ---- ダミーモード: 中立の意匠を自前で描く
        s = self.prs.slides.add_slide(self._blank)
        W, H = self.W, self.H
        if layout_name == LAYOUT_COVER:
            self.rect(s, 0, 0.510 * H, W, 0.062 * H, PERI)
            self.rect(s, 0, 0.510 * H, W, 0.006 * H, NAVY)
            self.text(s, W - 2.4, 0.510 * H, 1.9, 0.062 * H,
                      [[(DUMMY_LOGO, 22, True, GREY)]], align="right", anchor="middle")
            if title:
                self.text(s, self.MX, 0.365 * H, W - 2 * self.MX, 0.12 * H,
                          [[(title, 28, True, INK)]])
        elif layout_name in (LAYOUT_TOC, LAYOUT_SECTION):
            self.rect(s, 0, 0.883 * H, W, 0.045 * H, PERI)
            self.text(s, W - 2.2, 0.870 * H, 1.7, 0.06 * H,
                      [[(DUMMY_LOGO, 18, True, GREY)]], align="right", anchor="middle")
            if title:
                y = 0.065 * H if layout_name == LAYOUT_TOC else 0.400 * H
                size = 24 if layout_name == LAYOUT_TOC else 28
                self.text(s, self.MX, y, W - 2 * self.MX, 0.1 * H,
                          [[(title, size, True, INK)]])
        elif layout_name == LAYOUT_LAST:
            self.text(s, 0, 0.44 * H, W, 0.12 * H,
                      [[(DUMMY_LOGO, 40, True, GREY)]], align="center")
        else:  # 本文ページ
            if title:
                self.text(s, self.MX, 0.045 * H, W - 2 * self.MX - 1.3, 0.08 * H,
                          [[(title, 22, True, INK)]])
            self.rect(s, self.MX, 0.128 * H, self.CW, 0.004 * H, BLUE)
            self.text(s, W - self.MX - 1.3, 0.032 * H, 1.3, 0.06 * H,
                      [[(DUMMY_LOGO, 16, True, GREY)]], align="right")
        # フッター(表紙はページ番号なし)
        self.text(s, self.MX, H - 0.36, 0.5 * W, 0.3,
                  [[(DUMMY_FOOTER, 8, False, GREY)]])
        if layout_name != LAYOUT_COVER:
            self.text(s, W - self.MX - 0.8, H - 0.38, 0.8, 0.3,
                      [[(str(self._page), 10, False, INK2)]], align="right")
        return s

    def rect(self, s, x, y, w, h, fill, line=None, lw=1.0, shape=MSO_SHAPE.RECTANGLE,
             adj=None, dash=False):
        sp = s.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
        sp.shadow.inherit = False
        if adj is not None:
            try:
                sp.adjustments[0] = adj
            except Exception:
                pass
        if fill is None:
            sp.fill.background()
        else:
            sp.fill.solid()
            sp.fill.fore_color.rgb = _rgb(fill)
        if line is None:
            sp.line.fill.background()
        else:
            sp.line.color.rgb = _rgb(line)
            sp.line.width = Pt(lw)
            if dash:
                from pptx.enum.dml import MSO_LINE
                sp.line.dash_style = MSO_LINE.DASH
        return sp

    def _tag(self, s, x, y, text, size=12, cx=None):
        """タグ(角丸チップ)。ボックス内のサブ見出しはこの形で示す(全幅の帯にしない — ハウス好み)。
        cx を渡すと中央揃えで配置。戻り値は描画幅。"""
        t = str(text)
        tw = 0.30 + sum(0.16 if ord(c) > 0x2E80 else 0.088 for c in t) * (size / 11)
        if cx is not None:
            x = cx - tw / 2
        self.rect(s, x, y, tw, 0.34, "#FFFFFF", line=GREY, lw=1.0,
                  shape=MSO_SHAPE.ROUNDED_RECTANGLE, adj=0.5)
        self.text(s, x, y, tw, 0.34, [[(t, size, False, INK2)]],
                  align="center", anchor="middle")
        return tw

    def _line_seg(self, s, x1, y1, x2, y2, color, wpt=1.5):
        """任意角度の線分(細長い矩形を回転させて描く)。アイコン・概念図用。"""
        import math
        L = max(math.hypot(x2 - x1, y2 - y1), 0.01)
        th = wpt / 72.0
        sp = self.rect(s, (x1 + x2) / 2 - L / 2, (y1 + y2) / 2 - th / 2, L, th, color)
        sp.rotation = math.degrees(math.atan2(y2 - y1, x2 - x1))
        return sp

    def _dash_line(self, s, x1, y1, x2, y2, color=STEEL, wpt=1.5, seg=0.10):
        """破線(線分の連なり)。戻り・フィードバックの流れ、見込みの区切りに使う。"""
        import math
        L = math.hypot(x2 - x1, y2 - y1)
        nseg = max(int(L / seg), 2)
        for k in range(0, nseg, 2):
            t0, t1 = k / nseg, min((k + 1) / nseg, 1.0)
            self._line_seg(s, x1 + (x2 - x1) * t0, y1 + (y2 - y1) * t0,
                           x1 + (x2 - x1) * t1, y1 + (y2 - y1) * t1, color, wpt)

    def icon(self, s, kind, x, y, size=0.5, color=BLUE, style="line"):
        """簡易アイコン(単色・幾何図形のみ。多色・立体・絵文字・外部画像は使わない)。
        kind : book(辞書・規程) | cycle(学習・循環) | network(グラフ・関係) | search(検索) |
               db(データ) | doc(文書) | gear(処理) | person(利用者) | laptop(端末・画面) |
               chat(対話・LLM) | shield(認証・防御) | cloud(外部サービス) | chart(集計・評価)
        style: "line"(線画・既定。図解は全部これ) | "fill"(塗り。淡青の円に載せるときだけ明示)
        """
        ln = (style == "line")
        W = "#FFFFFF"

        def shp(px, py, w, h, shape, rot=0, lw=2.0, adj=None):
            if ln:
                sp = self.rect(s, px, py, w, h, None, line=color, lw=lw, shape=shape, adj=adj)
            else:
                sp = self.rect(s, px, py, w, h, color, shape=shape, adj=adj)
            if rot:
                sp.rotation = rot
            return sp

        def stroke(px, py, w, h):          # 塗りでは白抜き線、線画では色の線
            self.rect(s, px, py, w, h, color if ln else W)

        if kind == "book":
            shp(x, y + size * 0.10, size, size * 0.78, MSO_SHAPE.ROUNDED_RECTANGLE, adj=0.15)
            stroke(x + size * 0.47, y + size * 0.16, size * 0.06, size * 0.66)
            for i in range(2):
                yy = y + size * (0.34 + 0.20 * i)
                stroke(x + size * 0.13, yy, size * 0.24, 0.016)
                stroke(x + size * 0.63, yy, size * 0.24, 0.016)
        elif kind == "cycle":
            if ln:
                # 300° の円弧(上に切れ目)+線端の矢じり。枠の上下中央に収まる
                from lxml import etree
                sp = s.shapes.add_shape(MSO_SHAPE.ARC, Inches(x + size * 0.08), Inches(y + size * 0.08),
                                        Inches(size * 0.84), Inches(size * 0.84))
                sp.shadow.inherit = False
                try:
                    sp.adjustments[0] = 300 * 60000 / 100000     # 開始角 300°
                    sp.adjustments[1] = 240 * 60000 / 100000     # 終了角 240°(時計回りに 300°)
                except Exception:
                    pass
                sp.fill.background()
                sp.line.color.rgb = _rgb(color)
                sp.line.width = Pt(2.0)
                ln_el = sp.line._get_or_add_ln()
                for old in ln_el.findall(qn("a:tailEnd")):
                    ln_el.remove(old)
                te = etree.SubElement(ln_el, qn("a:tailEnd"))
                te.set("type", "triangle"); te.set("w", "med"); te.set("len", "med")
            else:
                shp(x, y, size, size, MSO_SHAPE.CIRCULAR_ARROW)
        elif kind == "network":
            pts = [(0.50, 0.12), (0.12, 0.58), (0.88, 0.58), (0.50, 0.88)]
            for a, b in [(0, 1), (0, 2), (1, 2), (1, 3), (2, 3)]:
                self._line_seg(s, x + pts[a][0] * size, y + pts[a][1] * size,
                               x + pts[b][0] * size, y + pts[b][1] * size, color, 2.0)
            d0 = size * 0.30
            for px, py in pts:
                if ln:
                    self.rect(s, x + px * size - d0 / 2, y + py * size - d0 / 2, d0, d0,
                              W, line=color, lw=2.0, shape=MSO_SHAPE.OVAL)
                else:
                    self.rect(s, x + px * size - d0 / 2, y + py * size - d0 / 2, d0, d0,
                              color, shape=MSO_SHAPE.OVAL)
        elif kind == "search":
            d0 = size * 0.66
            self.rect(s, x, y, d0, d0, None, line=color, lw=2.0, shape=MSO_SHAPE.OVAL)
            self._line_seg(s, x + d0 * 0.88, y + d0 * 0.88, x + size, y + size, color, 2.4)
        elif kind == "db":
            shp(x + size * 0.12, y, size * 0.76, size, MSO_SHAPE.CAN)
        elif kind == "doc":
            shp(x + size * 0.12, y, size * 0.76, size, MSO_SHAPE.FOLDED_CORNER)
            for i in range(3):
                yy = y + size * (0.30 + 0.18 * i)
                stroke(x + size * 0.26, yy, size * 0.40, 0.016)
        elif kind == "person":
            hd = size * 0.38
            shp(x + (size - hd) / 2, y, hd, hd, MSO_SHAPE.OVAL)
            try:
                shp(x + size * 0.10, y + size * 0.48, size * 0.80, size * 0.52,
                    MSO_SHAPE.ROUND_2_SAME_RECTANGLE, adj=0.5)
            except Exception:
                shp(x + size * 0.10, y + size * 0.48, size * 0.80, size * 0.52,
                    MSO_SHAPE.ROUNDED_RECTANGLE, adj=0.5)
        elif kind == "laptop":
            shp(x + size * 0.05, y + size * 0.08, size * 0.90, size * 0.62, MSO_SHAPE.RECTANGLE)
            self.rect(s, x, y + size * 0.76, size, size * 0.05, color)   # 台座は細い帯
        elif kind == "chat":
            shp(x, y + size * 0.08, size, size * 0.66, MSO_SHAPE.ROUNDED_RECTANGLE, adj=0.3)
            sp = shp(x + size * 0.16, y + size * 0.70, size * 0.22, size * 0.22,
                     MSO_SHAPE.RIGHT_TRIANGLE, lw=2.0)
            sp.rotation = 180
            for i in range(2):
                stroke(x + size * 0.22, y + size * (0.30 + 0.16 * i), size * 0.56, 0.016)
        elif kind == "shield":
            # 盾: 上辺が少し山なり、下が尖る多角形(フリーフォーム)+中にチェック
            pts = [(0.50, 0.00), (0.92, 0.14), (0.92, 0.50), (0.80, 0.76), (0.50, 1.00),
                   (0.20, 0.76), (0.08, 0.50), (0.08, 0.14)]
            ff = s.shapes.build_freeform(Inches(x + size * pts[0][0]), Inches(y + size * pts[0][1]))
            ff.add_line_segments([(Inches(x + size * px), Inches(y + size * py)) for px, py in pts[1:]],
                                 close=True)
            sp = ff.convert_to_shape()
            sp.shadow.inherit = False
            if ln:
                sp.fill.background()
                sp.line.color.rgb = _rgb(color)
                sp.line.width = Pt(2.0)
            else:
                sp.fill.solid()
                sp.fill.fore_color.rgb = _rgb(color)
                sp.line.fill.background()
            ck = color if ln else W
            self._line_seg(s, x + size * 0.32, y + size * 0.50, x + size * 0.46, y + size * 0.64, ck, 2.0)
            self._line_seg(s, x + size * 0.46, y + size * 0.64, x + size * 0.70, y + size * 0.36, ck, 2.0)
        elif kind == "cloud":
            shp(x, y + size * 0.12, size, size * 0.72, MSO_SHAPE.CLOUD)
        elif kind == "chart":
            bw = size * 0.22
            for i, hh in enumerate((0.45, 0.72, 0.95)):
                shp(x + size * (0.06 + 0.34 * i), y + size * (1 - hh), bw, size * hh - 0.02,
                    MSO_SHAPE.RECTANGLE, lw=2.0)
            self.rect(s, x, y + size - 0.02, size, 0.02, color)
        else:
            shp(x, y, size, size, MSO_SHAPE.GEAR_6)

    # ---------------------------------------------------------------- 概念図の部品(全て編集可能なネイティブ図形)
    def card(self, s, x, y, w, h, fill="#FFFFFF", line=GREY, lw=0.75, dash=False, adj=0.06):
        """角丸カード。主な用途は破線の境界(dash=True。部署別環境など)と関門図の段。
        領域を囲むのは panel()。塗りカードを横に並べない(AI くさい)。"""
        return self.rect(s, x, y, w, h, fill, line=line, lw=lw,
                         shape=MSO_SHAPE.ROUNDED_RECTANGLE, adj=adj, dash=dash)

    def zone(self, s, x, y, w, h, title):
        """[非推奨] 淡色の角丸ゾーン。新規では panel() を使う(互換のため残す)。"""
        self.card(s, x, y, w, h, fill=BG)
        self.text(s, x + 0.22, y + 0.16, w - 0.44, 0.32, [[(title, 14, True, BLUE_TX)]])

    def tri_down(self, s, cx, y, w=0.30, h=0.26, color=CYAN):
        sp = self.rect(s, cx - w / 2, y, w, h, color, shape=MSO_SHAPE.ISOSCELES_TRIANGLE)
        sp.rotation = 180
        return sp

    def tag_row(self, s, x, y, texts, size=12):
        """タグを横に並べる。戻り値は右端x。"""
        for t in texts:
            x += self._tag(s, x, y, t, size=size) + 0.15
        return x

    # ---------------------------------------------------------------- コンサル実物の語彙(2026-10-15。assets/ideal/abeam 参照)
    def col_label(self, s, x, y, w, text, size=14, sub=None, align="center"):
        """領域の見出し+細い下線(軽い区切り)。上に表や図が既にあって青帯のパネルを重ねたくない
        ときの下段の段組、chart_pair の列見出しに使う。align="left" で左寄せ。戻り値は下線の下の y。"""
        runs = [(text, size, True, INK)]
        if sub:
            runs.append(("　" + sub, 12, False, INK2))
        self.text(s, x, y, w, 0.36, [runs], align=align, anchor="bottom")
        self.rect(s, x, y + 0.40, w, 0.014, GREY)
        return y + 0.56

    def inline_steps(self, s, x, y, w, steps, size=14, trailing=None):
        """短い手順を 1 段落に「1. 記録　2. 原因の切り分け　3. …」と並べる(箱の手順図にしない)。
        trailing: 続けて置く補足の一文(INK2)。戻り値は下端の y。"""
        runs = []
        for i, st in enumerate(steps, 1):
            if i > 1:
                runs.append(("　", size, False, INK))
            runs.append((f"{i}. ", size, True, INK))
            runs.append((str(st), size, False, INK))
        paras = [runs]
        txt = "　".join(f"{i}. {st}" for i, st in enumerate(steps, 1))
        cpl = max(int((w - 0.2) / (size / 72)), 8)
        h = size / 72 * 1.45 * max(1, -(-int(_jlen(txt) + 0.999) // cpl)) + 0.05
        if trailing:
            paras.append([(str(trailing), size, False, INK2)])
            h += size / 72 * 1.45 * max(1, -(-int(_jlen(trailing) + 0.999) // cpl)) + 0.12
        self.text(s, x, y, w, h, paras, space_after=6, line_spacing=1.3)
        return y + h

    def _section_need(self, w, items, size):
        """section_block の必要高さの見積もり。"""
        cpl = _cpl(w, size, pad=0.2)
        need = 0.48
        bullets = []
        def flush_b():
            nonlocal need, bullets
            if bullets:
                need += self._fit_check(bullets, w, 99, size, "", warn=False)
                bullets = []
        for it in items:
            if isinstance(it, (tuple, list)) and it and it[0] == "steps":
                flush_b()
                txt = "　".join(f"{i}. {st}" for i, st in enumerate(it[1], 1))
                need += size / 72 * 1.45 * _nlines(txt, cpl) + 0.05
                if len(it) > 2 and it[2]:
                    need += size / 72 * 1.45 * _nlines(it[2], cpl) + 0.12
            elif isinstance(it, (tuple, list)) and it and it[0] == "p":
                flush_b()
                need += size / 72 * 1.45 * _nlines(it[1], cpl) + 0.08
            else:
                bullets.append(it)
        flush_b()
        return need

    def section_block(self, s, x, y, w, max_y, heading, items, size=14):
        """見出し+下線(col_label 左寄せ)の下に内容を置く軽い段組。items の要素:
        "文" → ● の箇条 / ("steps", [...], "補足"|None) → inline_steps / ("p", "文") → 平文(INK2)。
        入らない見込みなら 13pt に落とす。戻り値は下端の y。それでも max_y を超える見込みなら警告。"""
        if self._section_need(w, items, size) > max_y - y and size > 13:
            size = 13
        yy = self.col_label(s, x, y, w, heading, size=size + 1, align="left") - 0.08
        bullets = []

        def flush():
            nonlocal yy, bullets
            if bullets:
                need = self._fit_check(bullets, w, max_y - yy, size, f"段組「{heading}」", warn=False)
                self.text(s, x, yy, w, max(need, 0.3), self.nest_paras(bullets, size=size),
                          space_after=6, hanging=True)
                yy += need
                bullets = []
        for it in items:
            if isinstance(it, (tuple, list)) and it and it[0] == "steps":
                flush()
                yy = self.inline_steps(s, x, yy, w, it[1], size=size,
                                       trailing=(it[2] if len(it) > 2 else None))
            elif isinstance(it, (tuple, list)) and it and it[0] == "p":
                flush()
                cpl = max(int((w - 0.2) / (size / 72)), 8)
                h = size / 72 * 1.45 * max(1, -(-int(_jlen(it[1]) + 0.999) // cpl)) + 0.08
                self.text(s, x, yy, w, h, [[(str(it[1]), size, False, INK2)]], line_spacing=1.3)
                yy += h
            else:
                bullets.append(it)
        flush()
        if yy > max_y + 0.05:
            self._warn(f"段組「{heading}」が {yy - max_y:.2f}in はみ出す見込み。項目を減らす")
        return yy

    @staticmethod
    def _soft_break(text, cpl):
        """1 行に入らない文に、語の切れ目(、。や助詞の直後)で改行を 1 つ入れる。
        数字と単位(以上・以下・件・%)の間、英数字の途中では切らない。候補が無ければそのまま返す。"""
        t = str(text)
        if "\n" in t or _jlen(t) <= cpl:
            return t
        NO_NEXT = ("以上", "以下", "以内", "未満", "件", "%", "％", "ポイント", "回", "人", "円", "分", "時間")
        best = None
        for i in range(1, len(t)):
            prev, nxt = t[i - 1], t[i:]
            ok = prev in "、。" or (prev in "のをにでとはがへやも" and nxt[:1] not in "、。」）)" and
                                    not nxt.startswith(NO_NEXT))
            if not ok:
                continue
            if prev.isascii() and nxt[:1].isascii():
                continue
            if _jlen(t[:i]) <= cpl - 0.5 and _jlen(nxt) <= cpl:
                best = i
        return t if best is None else t[:best] + "\n" + t[best:]

    def header_box(self, s, x, y, w, h, title, body=None, items=None, fill=BLUE, size=16,
                   head_h=0.42, align="center"):
        """見出し帯(BLUE。panel・boxes と同じ色で揃える)+白本体(細枠)の箱。薄青の角丸カードの代わり。
        body: 本文(size pt) / items: ["文", ...](● の箇条書き。1 行に入らない文は語の切れ目で改行する)。
        戻り値は本文領域 (x, y, w, h)。並べるときは header_boxes() を使う(共通の文字サイズ・高さ)。"""
        self.rect(s, x, y, w, head_h, fill)
        self.text(s, x + 0.1, y, w - 0.2, head_h, [[(title, max(14, size + 1), True, "#FFFFFF")]],
                  align=align, anchor="middle")
        self.rect(s, x, y + head_h, w, h - head_h, "#FFFFFF", line=GREY, lw=0.75)
        bx, by, bw, bh = x + 0.18, y + head_h + 0.14, w - 0.36, h - head_h - 0.26
        if items:
            cpl = _cpl(bw, size, pad=size / 72 * 1.1)
            fixed = []
            for it in items:
                t2 = self._soft_break(it, cpl)
                if "\n" not in t2 and _jlen(t2) > cpl and _jlen(t2) % cpl <= 1.5:
                    self._warn(f"柱「{title}」の項目が語の途中で折り返す見込み: {str(it)[:16]}…")
                fixed.append(t2)
            paras = [self.bullet(t2, 1, size) for t2 in fixed]
            self.text(s, bx, by, bw, bh, paras, space_after=8, line_spacing=1.3, hanging=True)
        elif body:
            self.text(s, bx, by, bw, bh, [[(body, size, False, INK)]], line_spacing=1.3)
        return bx, by, bw, bh

    def header_boxes(self, s, y, bot, pillars, gap=0.30, sizes=(16, 15, 14)):
        """並列の柱を等幅で並べる。pillars: [(見出し, [項目...]), ...] 2〜4 本。
        文字サイズは全柱で共通(入る最大)。高さは bot まで。柱ごとの中身の種類が違う並列に使う
        (同じ属性が同じ順に並ぶなら gate_table か table にする)。戻り値は下端 y。"""
        n = len(pillars)
        if not 2 <= n <= 4:
            raise ValueError("header_boxes は 2〜4 本")
        w = (self.CW - gap * (n - 1)) / n
        h = bot - y
        bw, bh = w - 0.36, h - 0.42 - 0.26
        sz = min(self._fit_size(list(items), bw, bh, sizes, f"柱「{t}」") for t, items in pillars)
        rows = []
        for t, items in pillars:
            cpl = _cpl(bw, sz, pad=sz / 72 * 1.1)
            rows.append(sum(_nlines(self._soft_break(it, cpl), cpl) for it in items))
        if max(rows) - min(rows) >= 3:
            self._warn(f"柱の行数の差が大きい({min(rows)}〜{max(rows)} 行)。項目数を揃える")
        need = max(rows) * sz / 72 * 1.3 + 8 / 72 * max(len(it) for _, it in pillars)
        if need < bh * 0.5:
            self._warn("柱の中身が本体の半分に届かない。項目を足すか、箱を低くするか、型を変える")
        for i, (t, items) in enumerate(pillars):
            self.header_box(s, self.MX + i * (w + gap), y, w, h, t, items=list(items), size=sz)
        return bot

    def panel(self, s, x, y, w, h, title, sub=None, tone="dark", head_h=0.46):
        """領域パネル: ヘッダ帯+細枠の本体。左右・上下の領域を囲む既定の部品(ユーザー指定 2026-10-15)。
        tone: "dark"(BLUE 帯+白文字。既定。箱ヘッダと同じ語彙) | "light"(BG 帯+青文字)
        戻り値は本体の内側 (x, y, w, h)。この中に numbered_list や図を置く。"""
        if tone == "light":
            fill, tcol, scol = BG, BLUE_TX, INK2
        else:
            fill, tcol, scol = BLUE, "#FFFFFF", LTBLUE
        self.rect(s, x, y + head_h, w, h - head_h, "#FFFFFF", line=GREY, lw=0.75)
        self.rect(s, x, y, w, head_h, fill)
        runs = [(title, 14, True, tcol)]
        if sub:
            runs.append(("　" + sub, 12, False, scol))
        self.text(s, x + 0.2, y, w - 0.4, head_h, [runs], align="center", anchor="middle")
        return x + 0.22, y + head_h + 0.16, w - 0.44, h - head_h - 0.30

    def badge(self, s, cx, cy, n, r=0.17, fill=NAVY):
        """番号バッジ ①②…。numbered_list の番号を図の上に打って対応づける。"""
        self.rect(s, cx - r, cy - r, 2 * r, 2 * r, fill, shape=MSO_SHAPE.OVAL)
        self.text(s, cx - r, cy - r, 2 * r, 2 * r, [[(str(n), 12, True, "#FFFFFF")]],
                  align="center", anchor="middle")

    def numbered_list(self, s, x, y, w, items, size=12, gap=0.12, rule=True, start=1, max_y=None):
        """番号付き見出し+説明の縦並び(左に凡例、右に図、の左側)。items: [(見出し, 説明), ...]
        見出し size+2pt 太字、説明 size pt の完全文(1 行 24 字以内が目安)。戻り値は下端の y。
        max_y を渡すと、入りきらないときは項目間を詰め(下限 0)、それでも超えたら警告する。"""
        cpl = max(int((w - 0.55) * 72 / size), 10)
        last = start + len(items) - 1
        if max_y is not None and len(items) > 1:
            need = sum(0.42 + size / 72 * 1.3 * max(1, -(-int(_jlen(dsc) + 0.999) // cpl)) + 0.08
                       for _, dsc in items)
            room = max_y - y - need
            if room < gap * (len(items) - 1):
                gap = max(0.0, room / (len(items) - 1))
        for i, (head, desc) in enumerate(items, start):
            self.badge(s, x + 0.19, y + 0.19, i)
            self.text(s, x + 0.50, y + 0.01, w - 0.50, 0.36, [[(head, size + 2, True, INK)]],
                      anchor="middle")
            lines = max(1, -(-int(_jlen(desc) + 0.999) // cpl))
            dh = size / 72 * 1.3 * lines + 0.08
            self.text(s, x + 0.50, y + 0.42, w - 0.50, dh, [[(desc, size, False, INK2)]],
                      line_spacing=1.3)
            y += 0.42 + dh + gap
            if rule and i < last:
                self.rect(s, x + 0.50, y - gap / 2, w - 0.50, 0.012, GRID)
        if max_y is not None and y - gap > max_y + 0.02:
            self._warn(f"numbered_list が枠を {y - gap - max_y:.2f}in 超える。項目か説明を減らす")
        return y

    # ---------------------------------------------------------------- 構成図の部品(概念図の右側)
    def node_icon(self, s, cx, cy, kind, label, sub=None, color=BLUE, badge=None, size=0.60,
                  lab_w=2.0, badge_side=None):
        """線画アイコン(BLUE)+ラベル(12pt 太字)+補足(12pt)。cx,cy はアイコンの中心。
        badge を渡すと番号をラベルの左にインラインで付ける(「❷ 評価」の並び。位置は全ノード共通)。
        アイコンの色は BLUE だけ(NAVY は番号バッジ専用。渡すと BLUE に置き換えて警告)。"""
        if color == NAVY:
            self._warn(f"アイコン「{label}」の色 NAVY は使わない(BLUE に置き換えた)")
            color = BLUE
        elif color != BLUE:
            self._warn(f"アイコン「{label}」の色が BLUE 以外。図のアイコンは 1 色に揃える")
        if badge_side is not None:
            self._warn("node_icon の badge_side は廃止(番号はラベルの左に付く)")
        self.icon(s, kind, cx - size / 2, cy - size / 2, size, color, style="line")
        ly = cy + size / 2 + 0.05
        if badge:
            tw_ = _jlen(label) * 12 / 72 * CHAR_EM + 0.06
            sw_ = (_jlen(sub) * 12 / 72 * CHAR_EM + 0.06) if sub else 0.0
            unit_w = 0.34 + 0.08 + tw_
            if 0.42 + max(tw_, sw_) > lab_w + 0.05:
                self._warn(f"ノード「{label}」のラベルか補足が幅 {lab_w}in に入らない。短くするか lab_w を広げる")
            x0 = cx - unit_w / 2
            self.badge(s, x0 + 0.17, ly + 0.15, badge)
            tx = x0 + 0.34 + 0.08
            paras = [[(label, 12, True, INK)]]
            if sub:
                paras.append([(sub, 12, False, INK2)])
            self.text(s, tx, ly, max(tw_, sw_) + 0.25,
                      0.62 if sub else 0.36, paras, align="left", space_after=0, line_spacing=1.15)
        else:
            paras = [[(label, 12, True, INK)]]
            if sub:
                paras.append([(sub, 12, False, INK2)])
            self.text(s, cx - lab_w / 2, ly, lab_w, 0.62 if sub else 0.36, paras,
                      align="center", space_after=0, line_spacing=1.15)

    def arrow_h(self, s, x1, x2, y, label=None, color=STEEL, lab_w=2.0, below=False, dash=False):
        """水平の矢印(線+▶)。x1→x2(x2<x1 なら左向き)。label は線の上(below=True で下)に 12pt。
        dash=True は戻り・フィードバックの流れ(主の流れは実線)。"""
        seg = self._dash_line if dash else (lambda s_, a, b, c, d, col, w_: self._line_seg(s_, a, b, c, d, col, w_))
        if x2 >= x1:
            seg(s, x1, y, x2 - 0.16, y, color, 1.5)
            self.tri_right(s, x2 - 0.20, y - 0.10, 0.20, 0.20, color)
        else:
            seg(s, x1, y, x2 + 0.16, y, color, 1.5)
            sp = self.rect(s, x2, y - 0.10, 0.20, 0.20, color, shape=MSO_SHAPE.ISOSCELES_TRIANGLE)
            sp.rotation = 270
        if label:
            ly = y + 0.08 if below else y - 0.36
            self.text(s, (x1 + x2) / 2 - lab_w / 2, ly, lab_w, 0.28,
                      [[(label, 12, False, INK2)]], align="center")

    def arrow_v(self, s, x, y1, y2, label=None, color=STEEL, dash=False):
        """垂直の矢印(線+▼/▲)。y1→y2。label は線の右に 12pt。dash=True は戻りの流れ。"""
        seg = self._dash_line if dash else (lambda s_, a, b, c, d, col, w_: self._line_seg(s_, a, b, c, d, col, w_))
        if y2 >= y1:
            seg(s, x, y1, x, y2 - 0.16, color, 1.5)
            self.tri_down(s, x, y2 - 0.20, 0.20, 0.20, color)
        else:
            seg(s, x, y1, x, y2 + 0.16, color, 1.5)
            self.rect(s, x - 0.10, y2, 0.20, 0.20, color, shape=MSO_SHAPE.ISOSCELES_TRIANGLE)
        if label:
            self.text(s, x + 0.12, (y1 + y2) / 2 - 0.14, 2.0, 0.28, [[(label, 12, False, INK2)]])

    def concept_grid(self, rx, ry, rw, rh, cols, rows=2):
        """右 panel の内側 (rx, ry, rw, rh) を cols 列 × rows 行の格子に切り、ノード中心の
        (xs, ys) を返す。手計算の比率をやめ、どのデッキでも同じ間隔になる。"""
        xs = [rx + rw * (i + 0.5) / cols for i in range(cols)]
        top, bot = ry + 0.95, ry + rh - 0.95
        ys = [top] if rows == 1 else [top + (bot - top) * j / (rows - 1) for j in range(rows)]
        return xs, ys

    def zone_cols(self, s, rx, ry, rw, xs, c0, c1, row_y, label=None, has_sub=True):
        """列 c0〜c1 のノードを囲む破線枠。左右の辺は隣の列との中点(端は panel の内側 0.18in)、
        上下はアイコンとラベル・補足を含む高さにする。"""
        left = (xs[c0 - 1] + xs[c0]) / 2 if c0 > 0 else rx + 0.18
        right = (xs[c1] + xs[c1 + 1]) / 2 if c1 < len(xs) - 1 else rx + rw - 0.18
        top = row_y - 0.72
        bot = row_y + 0.35 + (0.62 if has_sub else 0.36) + 0.20
        self.dashed_zone(s, left, top, right - left, bot - top, label)

    def dashed_zone(self, s, x, y, w, h, label=None):
        """破線の境界(部署別環境・閉域など)。label は左上に 12pt。"""
        self.card(s, x, y, w, h, fill=None, line=GREY, lw=1.2, dash=True, adj=0.03)
        if label:
            self.text(s, x + 0.14, y + 0.06, w - 0.28, 0.3, [[(label, 12, False, INK2)]])

    def node(self, s, cx, cy, label, r=0.38, fill=BLUE, dash=False, size=12):
        """ノードリンク図の丸ノード。dash=True は「欠けている要素・未接続」の表現。"""
        if dash:
            self.rect(s, cx - r, cy - r, r * 2, r * 2, "#FFFFFF", line=GREY, lw=1.2,
                      shape=MSO_SHAPE.OVAL, dash=True)
            col = INK2
        else:
            self.rect(s, cx - r, cy - r, r * 2, r * 2, fill, shape=MSO_SHAPE.OVAL)
            col = "#FFFFFF"
        self.text(s, cx - r, cy - r, r * 2, r * 2,
                  [[(ln, size, True, col)] for ln in str(label).split("\n")],
                  align="center", anchor="middle", space_after=0, line_spacing=1.1)

    def edge(self, s, x1, y1, x2, y2, label=None, dash=False, color=STEEL):
        """ノード間の線(先にedge、後からnodeを描くと重なりが綺麗)。dash=弱い/欠けた関係。"""
        import math
        if dash:
            L = math.hypot(x2 - x1, y2 - y1)
            nseg = max(int(L / 0.09), 2)
            for k in range(0, nseg, 2):
                t0, t1 = k / nseg, min((k + 1) / nseg, 1.0)
                self._line_seg(s, x1 + (x2 - x1) * t0, y1 + (y2 - y1) * t0,
                               x1 + (x2 - x1) * t1, y1 + (y2 - y1) * t1, GREY, 1.5)
        else:
            self._line_seg(s, x1, y1, x2, y2, color, 1.5)
        if label:
            self.text(s, (x1 + x2) / 2 - 0.9, (y1 + y2) / 2 - 0.26, 1.8, 0.24,
                      [[(label, 12, False, INK2)]], align="center")

    def step_flow(self, s, x, y, w, h, steps, intro=None, last_hl=True):
        """番号付き手順(1→2→…)。**手順そのものが主役のときだけ**使う。仕組みの説明は
        概念図(panel+numbered_list+構成図)にする。

        steps: [(見出し 6 字以内, 説明 2 行以内), ...] 3〜5 個
        intro: (小ラベル, 本文) — 先頭に置く入力の箱。省略可
        last_hl: 最終ステップの上帯を濃紺にする(成果の位置)
        箱は白本体+GREY 0.75pt の細枠+上の青帯(薄青の角丸カードは使わない)。
        """
        aw = 0.30
        n = len(steps)
        iw = 1.8 if intro else 0
        cw_ = (w - iw - (n - 1 + (1 if intro else 0)) * aw) / n
        if intro and n >= 5:
            self._warn("step_flow で intro と 5 段を同時に置くと説明が 1 行 7 字になる。intro を外して入力は lead で言うか、4 段以下にする")
        cpl_d = _cpl(cw_ - 0.32, 14)
        for t, desc in steps:
            segs = str(desc).split("\n")
            if any(_jlen(sg) > cpl_d for sg in segs):
                self._warn(f"手順「{t}」の説明が 1 行({cpl_d} 字)を超える行がある。改行の位置を直すか短くする")
            if _nlines(desc, cpl_d) >= 3:
                self._warn(f"手順「{t}」の説明が 3 行以上になる見込み。2 行以内にする")
        cx = x
        if intro:
            self.rect(s, cx, y, iw, h, "#FFFFFF", line=STEEL, lw=1.2)
            self.text(s, cx + 0.16, y + 0.30, iw - 0.32, 0.28, [[(intro[0], 12, False, INK2)]])
            self.text(s, cx + 0.16, y + 0.62, iw - 0.32, h - 0.76,
                      [[(intro[1], 16, True, INK)]], line_spacing=1.3)
            cx += iw
            self.tri_right(s, cx + 0.07, y + h / 2 - 0.22, 0.16, 0.44, CYAN)
            cx += aw
        for i, (t, desc) in enumerate(steps):
            top = NAVY if (last_hl and i == n - 1) else BLUE
            self.rect(s, cx, y, cw_, h, "#FFFFFF", line=GREY, lw=0.75)
            self.rect(s, cx, y, cw_, 0.22, top)                       # 太い上帯
            self.rect(s, cx + 0.16, y + 0.40, 0.40, 0.40, top, shape=MSO_SHAPE.OVAL)
            self.text(s, cx + 0.16, y + 0.40, 0.40, 0.40,
                      [[(str(i + 1), 14, True, "#FFFFFF")]], align="center", anchor="middle")
            self.text(s, cx + 0.64, y + 0.38, cw_ - 0.74, 0.44, [[(t, 18, True, BLUE_TX)]],
                      anchor="middle")
            self.text(s, cx + 0.16, y + 0.96, cw_ - 0.32, h - 1.06,
                      [[(desc, 14, False, INK2)]], line_spacing=1.3)
            cx += cw_
            if i < n - 1:
                self.tri_right(s, cx + 0.07, y + h / 2 - 0.22, 0.16, 0.44, CYAN)
                cx += aw

    def stage_flow(self, s, x, y, w, h, stages, label=None):
        """状態付きステージパイプライン(学習手順・移行計画などの段階図)。

        stages: [(タイトル, サブ(改行は\\n), state), ...]
          state: "done"=青 | "now"=スチール青(現行) | "dead"=白+灰破線+×(棄却) | "plan"=白+青破線(計画)
        dead ステージから出る矢印には「切替」ラベルが自動で付く。
        """
        if label:
            self.text(s, x, y - 0.32, w, 0.28, [[(label, 12, False, INK2)]])
        aw = 0.36
        n = len(stages)
        cw_ = (w - (n - 1) * aw) / n
        cx = x
        for i, (t, sub, state) in enumerate(stages):
            if state == "dead":
                self.card(s, cx, y, cw_, h, line=GREY, lw=1.2, dash=True)
                self.text(s, cx, y + 0.15, cw_, 0.32,
                          [[(t + " ", 13, True, INK2), ("×", 15, True, GREY)]], align="center")
                tcol = INK2
            elif state == "plan":
                self.card(s, cx, y, cw_, h, line=STEEL, lw=1.2, dash=True)
                self.text(s, cx, y + 0.15, cw_, 0.32, [[(t, 14, True, BLUE_TX)]], align="center")
                tcol = BLUE_TX
            else:
                self.card(s, cx, y, cw_, h, fill=(STEEL if state == "now" else BLUE), line=None)
                self.text(s, cx, y + 0.15, cw_, 0.32, [[(t, 14, True, "#FFFFFF")]], align="center")
                tcol = "#FFFFFF"
            if sub:
                self.text(s, cx + 0.10, y + 0.52, cw_ - 0.20, h - 0.62,
                          [[(ln, 12, False, tcol)] for ln in str(sub).split("\n")],
                          align="center", space_after=2, line_spacing=1.25)
            cx += cw_
            if i < n - 1:
                if state == "dead":
                    self.text(s, cx - 0.30, y + h / 2 - 0.44, aw + 0.60, 0.24,
                              [[("切替", 12, True, BLUE_TX)]], align="center")
                self.tri_right(s, cx + 0.08, y + h / 2 - 0.11, aw - 0.16, 0.22,
                               GREY if stages[i + 1][2] == "plan" else CYAN)
                cx += aw

    def figure(self, title, image, lead=None, source=None, note=None):
        """⑭ 図版スライド。HTML図解(scripts/html2png.py でPNG化)などの画像を
        本文エリア中央に比率維持で配置する。
        編集可能なマスター(.html)は画像と同じ場所に必ず残すこと(画像は直接編集不可)。"""
        from PIL import Image as _PILImage
        s = self._slide(LAYOUT_BODY, title=title)
        self._kind("diagram")
        y0 = self._lead(s, lead)
        iw, ih = _PILImage.open(image).size
        top = y0 + 0.10
        bot = self._body_bot(source, note) - 0.10
        r = min(self.CW / iw, (bot - top) / ih)
        w, h = iw * r, ih * r
        s.shapes.add_picture(image, Inches(self.MX + (self.CW - w) / 2),
                             Inches(top + (bot - top - h) / 2), Inches(w), Inches(h))
        self._source(s, source, note)
        return s

    def tri_right(self, s, x, y, w, h, fill):
        """右向き三角形(▶)。図解の「流れ」はブロック矢印ではなくこれで示す(ハウス好み)。
        x,y,w,h は見た目のバウンディングボックス(回転を吸収して配置する)。"""
        cx, cy = x + w / 2, y + h / 2
        sp = self.rect(s, cx - h / 2, cy - w / 2, h, w, fill,
                       shape=MSO_SHAPE.ISOSCELES_TRIANGLE)
        sp.rotation = 90
        return sp

    def _hanging_indent(self, tb):
        """箇条書き段落にぶら下げインデントを付け、折返し行を本文の頭に揃える。
        bullet()/nest_paras() が作る「先頭の全角空白 run + マーカー run + 本文 run」を前提に、
        空白 run を外して段落の marL/indent に置き換える(それ以外の段落は触らない)。"""
        for p in tb.text_frame.paragraphs:
            runs = p.runs
            if not runs:
                continue
            first = runs[0].text
            lvl = 1
            if first and first.strip() == "" and "　" in first:
                lvl = 2 if len(first) <= 2 else 3
                runs[0]._r.getparent().remove(runs[0]._r)
                runs = p.runs
            if not runs or runs[0].text not in ("■", "●", "□"):
                continue
            size = max((r.font.size.pt if r.font.size else 12) for r in runs)
            marker_w = size / 72 * 1.05                 # マーカー+半角空白2つ分(約1字)
            base = {1: 0.0, 2: 2 * size / 72, 3: 4 * size / 72}[lvl]   # 元の全角空白ぶん
            pPr = p._p.get_or_add_pPr()
            pPr.set("marL", str(int(Inches(base + marker_w))))
            pPr.set("indent", str(-int(Inches(marker_w))))

    def text(self, s, x, y, w, h, paras, align="left", anchor="top",
             space_after=8, line_spacing=1.15, wrap=True, hanging=False):
        """paras: [[(text,size,bold,color[,font]), ...], ...]
        hanging=True で箇条書き(bullet()/nest_paras() の段落)にぶら下げインデントを付ける。"""
        tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = wrap
        tf.vertical_anchor = _ANCHOR[anchor]
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        for i, para in enumerate(paras):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = _ALIGN[align]
            p.space_after = Pt(space_after)
            p.line_spacing = line_spacing
            for run_spec in para:
                txt, size, bold, color = run_spec[:4]
                name = run_spec[4] if len(run_spec) > 4 else YU
                if name == YU and not bold and self.body_font:
                    name = self.body_font
                for k, piece in enumerate(str(txt).split("\n")):
                    if k > 0:
                        br = p._p.add_br()
                        rpr = br.get_or_add_rPr()
                        rpr.set("sz", str(int(size * 100)))
                        rpr.set("lang", "ja-JP")
                    r = p.add_run()
                    r.text = piece
                    _set_font(r, size, bold, color, name)
        if hanging:
            self._hanging_indent(tb)
        return tb

    def _lead(self, s, lead, y=None):
        """キーメッセージ(リード)。▍青バー+太字の 1 文。1 スライド 1 つ。
        全角 40 字以内は 20pt(1 行)。それを超えると 18pt(最大 2 行、目安 90 字)。
        戻り値は本文の開始 y(1 行≈2.23in / 2 行≈2.55in / lead 無し≈2.15in)。"""
        y = self.BODY_TOP if y is None else y
        if not lead:
            return y + 0.85
        if isinstance(lead, (list, tuple)):
            raise TypeError("lead は 1 文の str。● の複数本リードは使わない(1 スライド 1 キーメッセージ)")
        lead = str(lead)
        if lead.rstrip("。").count("。") >= 1:
            self._warn(f"リードが 2 文になっている。1 文で言い切る: {lead[:24]}…")
        if re.search(r"(について|の分析|の概要)$", lead.rstrip("。")):
            self._warn(f"リードが主題名で終わっている。主張を言い切る: {lead[:24]}…")
        size = 20 if _jlen(lead) <= 40 else 18
        line_h = size / 72 * 1.35
        cpl = _cpl(self.CW - 0.35, size)                     # 全角換算の 1 行文字数
        lines = max(1, -(-int(_jlen(lead) + 0.999) // cpl))
        if lines > 2:
            self._warn(f"リードが {lines} 行になる見込み(2 行まで)。短くする: {lead[:20]}…")
        self.rect(s, self.MX + 0.02, y + 0.13, 0.07, line_h * 0.82 * min(lines, 2) + (0.0 if lines == 1 else 0.05), BLUE)
        self.text(s, self.MX + 0.24, y + 0.10, self.CW - 0.35, 1.1,
                  [[(lead, size, True, INK)]])
        return y + 0.10 + line_h * lines + 0.45

    def lead(self, s, lead, y=None):
        """公開版の _lead(概念図などを部品で組む枚で使う)。戻り値は本文の開始 y。"""
        return self._lead(s, lead, y)

    # ---------------------------------------------------------------- 階層マーカー(唯一の定義)
    # ハウスラダー: 親(子持ち)=■太字 → 子「●」(親−2pt) → 孫「□」(親−4pt・グレー)。
    # content(ladder) で ■ を使うときは ■18/●16/□14 基準(入らなければ 17/16 基準へ)。
    # 他の型(箱・段組・表)は ■16/●14/□12 基準。リードは ▍青バー+太字で別格(_lead が描く)。
    def bullet(self, text, level=1, size=16, bold=False, heading=False):
        if heading:
            return [("■", max(8, round(size * 10 / 16)), True, INK),
                    ("  " + text, size, True, INK)]
        if level <= 1:
            return [("●", max(8, round(size * 9 / 16)), False, INK),
                    ("  " + text, size, bold, INK)]
        if level == 2:
            s = max(12, size - 2)
            return [("　　", s, False, INK),
                    ("●", max(8, round(s * 9 / 14)), False, INK),
                    ("  " + text, s, bold, INK)]
        s = max(12, size - 4)
        return [("　　　　", s, False, INK2),
                ("□", max(8, round(s * 10 / 12)), False, INK2),
                ("  " + text, s, False, INK2)]

    def nest_paras(self, items, size=16, level=1):
        """ネストしたリストを段落列に変換。items: ["文", ("親", ["子", ("子2", ["孫"])]), ...]
        親(子を持つ項目)は ■太字、子は「●」、孫は「□」。"""
        paras = []
        for it in items:
            if (isinstance(it, (tuple, list)) and len(it) == 2
                    and isinstance(it[1], (list, tuple))):
                head, kids = it
                if level == 1:
                    paras.append(self.bullet(head, 1, size, heading=True))
                else:
                    paras.append(self.bullet(head, level, size))
                paras.extend(self.nest_paras(kids, size, min(level + 1, 3)))
            else:
                paras.append(self.bullet(it, level, size))
        return paras

    def _ladder(self, s, y, bottom, bullets, sizes=None, x_pad=0.02):
        """ハウスラダーの箇条書きを y〜bottom に描く。
        ■(子を持つ親)があれば ■18/●16/□14 基準、無ければ ● 16pt 基準。収まらない時だけ段階縮小。
        ■ のグループの間は項目間の約 2 倍空ける。■ は MX+0.02(▍ の左端)に揃える。
        上限(■ 4 個・● だけなら 6 個・□ は 1 枚 1 か所)を超えたら警告する。戻り値は使用高さの見積もり。"""
        heads = [it for it in bullets if isinstance(it, (tuple, list)) and len(it) == 2
                 and isinstance(it[1], (list, tuple))]
        if sizes is None:
            sizes = ((18, 10), (17, 9), (16, 8), (15, 7), (14, 6)) if heads else ((16, 9), (15, 7), (14, 6))
        if len(heads) > 4:
            self._warn(f"■ の見出しが {len(heads)} 個。4 個以内にする(5 個以上は 2 枚に分けるか blocks)")
        if not heads and len(bullets) > 6:
            self._warn(f"● が {len(bullets)} 個。6 個以内にする")
        n_grand = sum(1 for h in heads for k in h[1]
                      if isinstance(k, (tuple, list)) and len(k) == 2 and isinstance(k[1], (list, tuple))
                      for _ in k[1])
        if n_grand > 1:
            self._warn(f"□(孫)が {n_grand} 行。孫は補足 1 行・1 枚 1 か所まで")
        avail = bottom - y
        est, paras, sp = 0.0, [], sizes[-1][1]
        for sz, sp in sizes:
            paras = self.nest_paras(bullets, size=sz)
            est = 0.0
            for p in paras:
                psz = max(r[1] for r in p)
                txt = "".join(r[0] for r in p)
                cpl = _cpl(self.CW - 2 * x_pad, psz, pad=0.4)
                est += psz / 72 * 1.2 * _nlines(txt, cpl) + sp / 72
            est += max(len(heads) - 1, 0) * sp / 72          # グループ間の追加余白
            if est * 1.12 <= avail:
                break
        if est * 1.12 > avail:
            self._warn(f"箇条書きが入りきらない見込み(必要 {est:.1f}in / 余白 {avail:.1f}in)。項目を減らすか 2 枚に分ける")
        tb = self.text(s, self.MX + x_pad, y, self.CW - 2 * x_pad, avail, paras,
                       space_after=sp, hanging=True)
        first = True
        for p in tb.text_frame.paragraphs:        # 2 つ目以降の ■ の前にグループ間の余白
            runs = p.runs
            if runs and runs[0].text == "■":
                if not first:
                    p.space_before = Pt(sp)
                first = False
        return est

    def note(self, text):
        s = self.prs.slides[-1]
        s.notes_slide.notes_text_frame.text = text

    def _source(self, s, source, note=None):
        if note:
            self.text(s, self.MX, self.H - 0.92, self.CW, 0.3,
                      [[("注:" + note, 10, False, INK2)]])
        if source:
            self.text(s, self.MX, self.H - 0.71, self.CW, 0.3,
                      [[("出典:" + source, 10, False, INK2)]])

    def _body_bot(self, source=None, note=None):
        """本文の下端。出典・注記があればその行の上で止める(takeaway 予約済みなら更に低い方)。"""
        b = self.BODY_BOT
        if note:
            b = min(b, self.H - 1.06)
        elif source:
            b = min(b, self.H - 0.86)
        return b

    # 部品で組む枚(概念図・柱・関門図)のための公開 API。座標は d.MX / d.CW / 戻り値の y で決める
    def slide(self, title):
        """本文ページを 1 枚追加して返す(公式ヘッダ・ロゴ・フッター付き)。部品で組む図解の枚に使う。"""
        s = self._slide(LAYOUT_BODY, title=title)
        self._kind("diagram")
        return s

    def body_bot(self, source=None, note=None, takeaway=False):
        """本文の下端 y。takeaway=True なら結論帯の分(0.72in)も空ける。"""
        return self._body_bot(source, note) - (0.72 if takeaway else 0.0)

    def source(self, s, source=None, note=None):
        """出典(H-0.71)と注記(H-0.92)の行を描く。"""
        self._source(s, source, note)

    def takeaway_band(self, s, takeaway, source=None, note=None):
        """結論帯を本文の最下段(body_bot の直上)に描く。"""
        self._takeaway_band(s, self._body_bot(source, note) - 0.72 - 0.12 + 0.0, takeaway)

    # ---------------------------------------------------------------- ① 表紙
    def cover(self, title, subtitle=None, label=None, date=None, dept=None, company=None):
        """表紙。title は全角 20 字目安(30 字を超えると 2 行になり label に迫る)。タイトルか subtitle の
        どちらかに課題・目的を入れる。label は会議名と回(社外向けは「○○株式会社 御中」)。
        date と dept(作成部署)は版の管理に要るので省かない(無いと警告)。"""
        s = self._slide(LAYOUT_COVER, title=title)
        H = self.H
        if _jlen(title) > 30:
            self._warn("表紙タイトルが 2 行になる見込み(全角 20 字目安)。副題に分ける")
        if not date:
            self._warn("表紙に日付が無い(版の取り違えの元)")
        if not dept:
            self._warn("表紙に作成部署(dept)が無い。社内資料は部署か発表者を書く")
        if company is None:
            company = self.brand_cfg.get("company") if self._brand else "Your Company"
        if label:
            self.text(s, self.MX, 0.316 * H, 0.8 * self.W, 0.4, [[(label, 15, False, INK)]])
        if subtitle:
            self.text(s, self.MX, 0.558 * H, 0.85 * self.W, 0.5,
                      [[("― " + subtitle + " ―", 17, True, BLUE_TX)]])
        if date:
            self.text(s, self.MX, 0.651 * H, 7.0, 0.4, [[(date, 14, False, INK)]])
        if company:
            self.text(s, self.MX, 0.768 * H, 9.0, 0.4, [[(company, 16, True, INK)]])
        if dept:
            self.text(s, self.MX, 0.820 * H, 9.0, 0.4, [[(dept, 14, False, INK)]])
        return s

    # ---------------------------------------------------------------- ② 目次
    def toc(self, items, current=None, title="目次", pages=True):
        """目次。items: ["章"] か [("章", "補足")]。3〜6 章。章名は中扉と一字一句同じにする
        (section() が照合する)。current は再掲するときだけ(最初の目次では渡さない。中扉が現在地を示す)。
        pages=True で各章の開始ページを右端に入れる(save() 時に埋まる)。"""
        if len(items) > 6:
            raise ValueError("目次は 6 章まで。章を束ねる")
        if len(items) < 3:
            self._warn("章が 2 つ以下なら目次は省いてよい")
        s = self._slide(LAYOUT_TOC, title=title)
        self._toc_items = [str(it[0] if isinstance(it, (tuple, list)) else it) for it in items]
        x = 0.237 * self.W
        y = 0.242 * self.H
        step = 0.118 * self.H
        if len(items) >= 6:               # 6 章は下の帯に掛かるので、開始を上げて間隔を詰める
            y = 0.20 * self.H
            step = min(step, (0.80 * self.H - y) / (len(items) - 1))
        right = self.W - self.MX - 0.3
        for i, it in enumerate(items, 1):
            name, sub = (it if isinstance(it, (tuple, list)) else (it, None))
            if current == i:
                hl_w = min(right - x + 0.28,
                           0.9 + _jlen(name) * 23 / 72 * CHAR_EM + (_jlen(sub) * 14 / 72 * CHAR_EM + 0.5 if sub else 0) + 0.8)
                self.rect(s, x - 0.28, y - 0.14, hl_w, 0.72, LTBLUE)
            runs = [(f"{i}.", 21, True, STEEL), ("   " + name, 23, True, BLUE_TX)]
            if sub:
                runs.append(("  ― " + sub, 14, False, STEEL))
            self.text(s, x, y, right - x - 1.0, 0.5, [runs])
            if pages:
                tb = self.text(s, right - 1.0, y + 0.10, 1.0, 0.4, [[("", 14, False, INK2)]], align="right")
                self._toc_page_boxes.append((i, tb))
            y += step
        if self._brand:   # 目次レイアウトにはページ番号の枠が無いので本文と同じ位置に描く
            self.text(s, self.W - self.MX - 0.55, self.H - 0.45, 0.55, 0.25,
                      [[(str(self._page), 10, False, INK2)]], align="right")
        return s

    # ---------------------------------------------------------------- ③ 中扉
    def section(self, number, title, nav=True):
        """中扉。章番号と章名(目次と同じ表記)。nav=True なら目次の章一覧を再掲して現在の章を強調する
        (全体のどこにいるかを示す)。中扉を挟む章は本文 3 枚以上にする(少なければ save() が警告)。"""
        s = self._slide(LAYOUT_SECTION, title=f"{number}. {title}")
        self._chapters.append([number, self._page, 0])
        items = self._toc_items
        if items:
            if number < 1 or number > len(items):
                self._warn(f"中扉の章番号 {number} が目次({len(items)} 章)と合わない")
            elif items[number - 1] != str(title):
                self._warn(f"中扉「{title}」が目次「{items[number - 1]}」と一字一句同じでない")
        if nav and items and len(items) >= 2:
            x, y = self.MX, 0.54 * self.H
            step = (0.052 if len(items) <= 4 else 0.044) * self.H
            for i, name in enumerate(items, 1):
                cur = (i == number)
                if cur:
                    self.rect(s, x - 0.18, y - 0.06, min(7.0, 1.2 + _jlen(name) * 16 / 72 * CHAR_EM + 0.6),
                              0.40, LTBLUE)
                self.text(s, x, y, 8.0, 0.30, [[(f"{i}.", 16, True, STEEL if cur else INK2),
                          ("   " + name, 16, cur, BLUE_TX if cur else INK2)]])
                y += step
        return s

    # ---------------------------------------------------------------- ④ 本文・箇条書き
    def content(self, title, lead=None, bullets=None, source=None, style="ladder",
                note=None, takeaway=None):
        """bullets: ["文"] または [("親", ["子1", ("子2", ["孫"])])] — 最大3階層

        style: "ladder"(既定) — ■→●→□ のハウスラダー。並列の事実・方針・条件はこれが一番速い。
               "blocks" — 親=見出し行(青太字)+子=本文行のブロック段組。グループ間は細罫線。
               4 グループ以上は 2 列(左上→右上→左下→右下の行優先で、行の高さを揃える)。
               6 グループまで。lead でグループを列挙するなら全部入れる(入らないなら数で言う)。
        """
        s = self._slide(LAYOUT_BODY, title=title)
        self._kind("text")
        _tb = self._reserve_takeaway(takeaway, source, note)
        y = self._lead(s, lead)
        if style == "ladder" or not bullets:
            if bullets:
                self._ladder(s, y, self._body_bot(source, note), bullets)
            self._finish_takeaway(s, _tb, takeaway)
            self._source(s, source, note)
            return s
        groups = []
        for it in bullets:
            if (isinstance(it, (tuple, list)) and len(it) == 2
                    and isinstance(it[1], (list, tuple))):
                groups.append((str(it[0]), list(it[1])))
            else:
                groups.append((str(it), []))
        n = len(groups)
        if n > 6:
            self._warn(f"段組が {n} グループ。6 以内にする(2 枚に分けるか表にする)")
        kid_counts = [len(k) for _, k in groups]
        if kid_counts and max(kid_counts) - min(kid_counts) >= 2:
            self._warn(f"段組のグループごとの行数の差が大きい({min(kid_counts)}〜{max(kid_counts)} 行)。揃える")
        two_col = n >= 4
        gap_col = 0.70
        tw = (self.CW - gap_col) / 2 - 0.10 if two_col else min(self.CW - 0.1, 10.6)

        ks = 14   # 子の文字サイズ。入るなら 16pt に上げる(下で決める)

        def g_height(kids, width):
            k_ = ks / 14
            cpl = max(int((width - 0.4) / (ks / 72)), 10)
            h = 0.40 * k_
            for k in kids:
                kk, gks = (k if isinstance(k, (tuple, list)) else (k, []))
                h += 0.32 * k_ * max(1, -(-int(_jlen(kk) + 0.999) // cpl)) + 0.06
                for gk in gks:
                    h += 0.30 * k_ * max(1, -(-int(_jlen(gk) + 0.999) // cpl)) + 0.04
            return h + 0.24

        def draw_group(x, gy, width, head, kids, rule, rule_y=None):
            # 見出しは青太字の文字だけ(▍はリード専用なのでタックを付けない)。子は lead の本文と同じ頭に揃える
            k_ = ks / 14
            self.text(s, x, gy, width, 0.36 * k_, [[(head, ks + 2, True, BLUE_TX)]])
            yy = gy + 0.44 * k_
            cpl = max(int((width - 0.4) / (ks / 72)), 10)
            for k in kids:
                kk, gks = (k if isinstance(k, (tuple, list)) else (k, []))
                self.text(s, x + 0.22, yy, width - 0.22, 0.6 * k_,
                          [[(str(kk), ks, False, INK)]], line_spacing=1.3)
                yy += 0.32 * k_ * max(1, -(-int(_jlen(kk) + 0.999) // cpl)) + 0.06
                for gk in gks:
                    self.text(s, x + 0.40, yy, width - 0.40, 0.5 * k_,
                              [[("□  " + str(gk), ks - 1, False, INK2)]])
                    yy += 0.30 * k_ * max(1, -(-int(_jlen(gk) + 0.999) // cpl)) + 0.04
            if rule:
                self.rect(s, x, (rule_y if rule_y is not None else yy + 0.10), width, 0.012, GRID)
            return yy + 0.26

        _avail = self._body_bot(source, note) - 0.15 - y
        ks = 16
        pairs = [groups[i:i + 2] for i in range(0, n, 2)]          # 2 列は行優先(Z 順)
        if two_col:
            _need = sum(max(g_height(k, tw) for _, k in pr) for pr in pairs)
        else:
            _need = sum(g_height(k, tw) for _, k in groups)
        if _need > _avail:
            ks = 14
        if two_col:
            avail = _avail
            row_hs = [max(g_height(k, tw) for _, k in pr) for pr in pairs]
            total = sum(row_hs)
            if total > avail + 0.05:
                self._warn("段組(blocks)が入りきらない見込み。グループか行を減らす")
            extra = (min(0.45, max(0.0, (avail - total) / max(len(pairs) - 1, 1) * 0.5))
                     if len(pairs) > 1 else 0.0)
            yy = y
            for pi, pr in enumerate(pairs):
                last_row = (pi == len(pairs) - 1)
                rule_y = yy + row_hs[pi] - 0.16 if not last_row else None
                for ci, (head, kids) in enumerate(pr):
                    x = self.MX + 0.02 + ci * (tw + 0.10 + gap_col)
                    draw_group(x, yy, tw, head, kids, not last_row, rule_y)
                yy += row_hs[pi] + extra
        else:
            avail = self._body_bot(source, note) - 0.15 - y
            total = sum(g_height(k, tw) for _, k in groups)
            if total > avail + 0.05:
                self._warn("段組(blocks)が入りきらない見込み。グループか行を減らす")
            extra = (min(0.30, max(0.0, (avail - total) / max(n - 1, 1) * 0.6))
                     if n > 1 else 0.0)
            yy = y
            for gi, (head, kids) in enumerate(groups):
                yy = draw_group(self.MX + 0.02, yy, tw, head, kids, gi < n - 1) + extra
        self._finish_takeaway(s, _tb, takeaway)
        self._source(s, source, note)
        return s

    # ---------------------------------------------------------------- ⑤ ボックス対比(2〜3列)
    def boxes(self, title, boxes, lead=None, arrow=False, source=None, note=None,
              takeaway=None, arrows=None, band=None):
        """箱の列(2〜3)。boxes: [(見出し, 小見出し|None, [項目...]), ...]
        arrow=True : 前後の変化(▶ を間に置く)。False: 並列の案・柱
        arrows=[...]: 矢印にラベルを付けた受け渡し(役割分担)。band= は下の共通の土台。この 2 つを渡すと flow と同じ描画
        各箱 4〜5 項目で揃える。同じ属性が同じ順で並ぶなら table にする。"""
        if arrows or band:
            return self.flow(title, boxes, arrows=arrows, lead=lead, band=band, source=source, note=note)
        s = self._slide(LAYOUT_BODY, title=title)
        self._kind("box")
        _tb = self._reserve_takeaway(takeaway, source, note)
        y0 = self._lead(s, lead) - 0.15
        n = len(boxes)
        gap = 0.58 if arrow else 0.30
        w = (self.CW - gap * (n - 1)) / n
        bot = self._body_bot(source, note) - 0.10
        # 文字サイズは全ての箱で共通(各箱で入る最大のうち最小)
        bsz = min(self._fit_size(it, w - 0.60, bot - (y0 + 0.55 + 0.20 + (0.60 if sb else 0)) - 0.17,
                                 (16, 15, 14), f"箱「{hd}」") for hd, sb, it in boxes)
        for i, (head, sub, items) in enumerate(boxes):
            x = self.MX + i * (w + gap)
            self.rect(s, x, y0, w, 0.55, BLUE)
            self.text(s, x, y0, w, 0.55, [[(head, 17, True, "#FFFFFF")]],
                      align="center", anchor="middle")
            yy = y0 + 0.55
            self.rect(s, x, yy, w, bot - yy, "#FFFFFF", line=GREY, lw=0.75)  # 白本体+細枠(パネルと同じ語彙)
            iy = yy + 0.20
            if sub:
                # サブ見出し: 青太字 16pt+細い下線(チップにしない)
                self.text(s, x + 0.30, iy, w - 0.60, 0.36, [[(sub, 16, True, BLUE_TX)]])
                self.rect(s, x + 0.30, iy + 0.42, w - 0.60, 0.012, GRID)
                iy += 0.60
            self.text(s, x + 0.30, iy, w - 0.60, bot - iy - 0.17,
                      self.nest_paras(self._soft_items(items, w - 0.60, bsz), size=bsz),
                      space_after=8, hanging=True)
            if arrow and i < n - 1:
                self.tri_right(s, x + w + (gap - 0.30) / 2, (y0 + bot) / 2 - 0.42,
                               0.30, 0.85, CYAN)   # 細く高い▶(ユーザー手直し 2026-10-16)
        self._finish_takeaway(s, _tb, takeaway)
        self._source(s, source, note)
        return s

    # ---------------------------------------------------------------- ⑥ kpi(廃止)
    def kpi(self, *a, **k):
        """廃止(2026-10-15)。大数字だけでは何の数字か伝わらない。"""
        raise TypeError("kpi は廃止。数値は table(指標・測り方・しきい値・現状)か "
                        "stage_flow(各段に測り方+しきい値)に置く")

    def _soft_items(self, items, w, size, pad=0.0):
        """箇条の文字列項目に _soft_break を掛ける(入れ子はそのまま)。"""
        cpl = _cpl(w, size, pad=pad + size / 72 * 1.1)
        out = []
        for it in items:
            if isinstance(it, (tuple, list)):
                out.append(it)
            else:
                out.append(self._soft_break(it, cpl))
        return out

    def _fit_size(self, items, w, h, sizes, where):
        """sizes(大→小)のうち、items が幅 w・高さ h に入る最大の pt を返す。最小でも入らなければ警告。"""
        def _flat(xs):
            for it in xs:
                if isinstance(it, (tuple, list)):
                    yield str(it[0])
                    if len(it) > 1 and isinstance(it[1], (tuple, list)):
                        yield from _flat(it[1])
                else:
                    yield str(it)

        def _orphan(sz):   # 最終行が 1〜1.5 字だけになる項目があるか(泣き別れ)
            cpl = _cpl(w, sz, pad=0.3)
            for t in _flat(items):
                for seg in str(t).split("\n"):
                    L = _jlen(seg)
                    if L > cpl and 0 < (L % cpl) <= 1.5:
                        return True
            return False

        fits = [sz for sz in sizes if self._fit_check(items, w, h, sz, where, warn=False) <= h + 0.05]
        for sz in fits:            # 入るサイズのうち、泣き別れが出ない最大
            if not _orphan(sz):
                return sz
        if fits:
            return fits[0]
        self._fit_check(items, w, h, sizes[-1], where)
        return sizes[-1]

    def _fit_check(self, items, w, h, size, where, warn=True):
        """箇条書き items が幅 w・高さ h に size pt で入るか見積もり、入らなければ警告する。"""
        paras = self.nest_paras(items, size=size)
        need = 0.0
        for p in paras:
            psz = max(r[1] for r in p)
            txt = "".join(r[0] for r in p)
            cpl = _cpl(w, psz, pad=0.3)
            need += psz / 72 * 1.25 * _nlines(txt, cpl) + 8 / 72
        if warn and need > h + 0.05:
            self._warn(f"{where} の文字が入りきらない見込み(必要 {need:.1f}in / 枠 {h:.1f}in)。項目か文字を減らす")
        return need

    # ---------------------------------------------------------------- ⑦ ネイティブチャート
    # プロット領域の近似パディング(左, 右, 上, 下)。図上注記の座標に使う
    _CH_PAD = (0.55, 0.12, 0.14, 0.42)

    @staticmethod
    def _nice_ceil(v):
        import math
        if v <= 0:
            return 1.0
        e = 10 ** math.floor(math.log10(v))
        for m in (1, 1.2, 1.5, 2, 2.5, 3, 4, 5, 6, 8, 10):
            if m * e >= v - 1e-9:
                return m * e
        return 10 * e

    def _native_chart(self, s, x, y, w, h, kind, categories, series, value_fmt="#,##0",
                      highlight=None, headroom=1.0, ymin=None, ymax=None, value_axis="show",
                      need_top=0.0, pad_r=None, legend=None):
        """ネイティブグラフを (x,y,w,h) に描く。プロット領域は manualLayout で固定するので、
        戻り値の fx/fy は正確な座標になる(注記・目標線・直接ラベルの位置決めに使う)。
        value_axis: "show"=値軸と目盛線を出す / "hide"=消す(値ラベルが全部あるとき)
        need_top  : 軸を消すときに棒の上に要る高さ(inch。値ラベル+注記)。上限はここから計算する
        戻り値: (chart, fx(i), fy(v), tops, geo) geo={"l","r","t","b","slot","bar_half","lo","hi"}"""
        from lxml import etree
        from pptx.enum.chart import XL_TICK_MARK, XL_MARKER_STYLE
        data = CategoryChartData()
        data.categories = categories
        for name, vals in series:
            data.add_series(name, vals)
        xl = {"bar": XL_CHART_TYPE.COLUMN_CLUSTERED,
              "stack": XL_CHART_TYPE.COLUMN_STACKED,
              "line": XL_CHART_TYPE.LINE}[kind]
        gf = s.shapes.add_chart(xl, Inches(x), Inches(y), Inches(w), Inches(h), data)
        c = gf.chart
        c.has_title = False
        c.font.size = Pt(12)
        c.font.name = YU
        c.font.color.rgb = _rgb(INK)
        show_legend = (len(series) > 1) if legend is None else bool(legend)
        c.has_legend = show_legend
        if c.has_legend:
            c.legend.position = XL_LEGEND_POSITION.BOTTOM
            c.legend.include_in_layout = False
            c.legend.font.size = Pt(12)
        plot = c.plots[0]
        try:
            plot.gap_width = 60 if kind != "line" else 150
        except Exception:
            pass
        colors = STACK_COLORS if kind == "stack" else SERIES_COLORS
        for i, ser in enumerate(plot.series):
            col = _rgb(colors[i % len(colors)])
            if kind == "line":
                ser.format.line.color.rgb = col
                ser.format.line.width = Pt(2.5)
                ser.smooth = False
                ser.marker.style = XL_MARKER_STYLE.CIRCLE
                ser.marker.size = 7
                ser.marker.format.fill.solid()
                ser.marker.format.fill.fore_color.rgb = col
                ser.marker.format.line.color.rgb = col
            else:
                ser.format.fill.solid()
                ser.format.fill.fore_color.rgb = col
                ser.format.line.fill.background()
        if highlight is not None and not (kind == "bar" and len(series) == 1):
            self._warn("highlight は単系列の bar だけ(stack は一番下の区分と share_of で主役を示す)")
        if kind == "bar" and len(series) == 1 and highlight is not None:
            # 主役の棒だけ濃紺、他は淡青(単色チャートの1点強調)
            hl = {highlight} if isinstance(highlight, int) else set(highlight)
            ser0 = plot.series[0]
            for pi in range(len(categories)):
                pt = ser0.points[pi]
                pt.format.fill.solid()
                pt.format.fill.fore_color.rgb = _rgb(NAVY if pi in hl else RAMP[0])
        if kind == "bar" and len(series) == 1:
            plot.has_data_labels = True
            dl = plot.data_labels
            dl.number_format = value_fmt
            dl.number_format_is_linked = False
            dl.position = XL_LABEL_POSITION.OUTSIDE_END
            dl.font.size = Pt(12); dl.font.bold = True
            dl.font.color.rgb = _rgb(INK); dl.font.name = YU
        elif kind == "stack":
            light = {CYANSUB, PERI, CYAN, "#66A6EA", TEAL, GREY}   # 明るい系列は濃い文字(白だと読めない)
            for i, ser in enumerate(plot.series):
                col = colors[i % len(colors)]
                dls = ser.data_labels
                dls.show_value = True
                dls.number_format = value_fmt
                dls.number_format_is_linked = False
                dls.font.size = Pt(12); dls.font.bold = True; dls.font.name = YU
                dls.font.color.rgb = _rgb(INK if col in light else "#FFFFFF")
        cat = c.category_axis
        cat.has_major_gridlines = False
        cat.format.line.color.rgb = _rgb(GREY)
        cat.tick_labels.font.size = Pt(12)
        cat.tick_labels.font.name = YU
        cat.major_tick_mark = XL_TICK_MARK.NONE          # Excel 既定の外向き目盛りを消す
        cat.minor_tick_mark = XL_TICK_MARK.NONE
        val = c.value_axis
        n = max(len(categories), 1)
        tops = []
        for i in range(n):
            vs = [vals[i] for _, vals in series if i < len(vals) and vals[i] is not None]
            tops.append((sum(vs) if kind == "stack" else max(vs)) if vs else 0.0)
        if kind in ("bar", "stack") and ymin not in (None, 0):
            self._warn("棒グラフは 0 始まり固定(差の誇張防止)。ymin を無視した")
            ymin = None
        lo = 0.0 if ymin is None else float(ymin)
        vmax = max(tops + [1e-9])
        # プロット領域(内側)のパディング。値軸を消すときは左を詰める
        pad_l = 0.55 if value_axis == "show" else 0.12
        pad_r_ = 0.12 if pad_r is None else pad_r
        pad_t = 0.14
        pad_b = 0.42 + (0.32 if c.has_legend else 0.0)
        plot_l, plot_r = x + pad_l, x + w - pad_r_
        plot_t, plot_b = y + pad_t, y + h - pad_b
        plot_h = max(plot_b - plot_t, 0.5)
        if value_axis == "show":
            if ymax is not None and headroom > 1.0 and (float(ymax) - vmax) < 0.12 * (float(ymax) - lo):
                self._warn(f"ymax={ymax} では注記の余白が足りない(最大 {vmax:g})。ymax を上げるか注記を減らす")
            hi = float(ymax) if ymax is not None else self._nice_ceil(vmax * headroom)
        else:
            # 軸を消す: 上限は「棒の上に要る高さ」から正確に計算(きりの良い数にしない。目盛りは出ないため)
            need = min(max(need_top, 0.0), plot_h * 0.6)
            hi = float(ymax) if ymax is not None else vmax * plot_h / (plot_h - need)
            if ymax is not None and (float(ymax) - vmax) / max(float(ymax) - lo, 1e-9) * plot_h < need_top - 0.02:
                self._warn(f"ymax={ymax} では値ラベル・注記の余白が足りない。ymax を上げるか注記を減らす")
        val.minimum_scale = lo
        val.maximum_scale = hi
        val.tick_labels.font.size = Pt(12)
        val.tick_labels.font.name = YU
        val.tick_labels.number_format = value_fmt
        val.tick_labels.number_format_is_linked = False
        val.format.line.fill.background()
        val.major_tick_mark = XL_TICK_MARK.NONE
        if value_axis == "show":
            val.has_major_gridlines = True
            val.major_gridlines.format.line.color.rgb = _rgb(GRID)
            val.major_gridlines.format.line.width = Pt(0.75)
            rng = max(hi - lo, 1e-9)
            import math
            e = 10 ** math.floor(math.log10(rng / 6))
            cands = [m * e for m in (1, 2, 2.5, 5, 10, 20)]
            val.major_unit = min(cc for cc in cands if cc >= rng / 6 - 1e-9)   # 目盛りは 5〜6 本
        else:
            val.has_major_gridlines = False
            val.visible = False
        # プロット領域を固定(manualLayout inner)。自動配置によるずれを無くし fx/fy を正確にする
        pa = plot._element.getparent()
        lay = pa.find(qn("c:layout"))
        if lay is None:
            lay = etree.SubElement(pa, qn("c:layout"))
            pa.remove(lay)
            pa.insert(0, lay)
        for ch_ in list(lay):
            lay.remove(ch_)
        ml = etree.SubElement(lay, qn("c:manualLayout"))
        for tag, v in (("c:layoutTarget", "inner"), ("c:xMode", "edge"), ("c:yMode", "edge"),
                       ("c:x", f"{pad_l / w:.5f}"), ("c:y", f"{pad_t / h:.5f}"),
                       ("c:w", f"{(w - pad_l - pad_r_) / w:.5f}"), ("c:h", f"{(h - pad_t - pad_b) / h:.5f}")):
            etree.SubElement(ml, qn(tag)).set("val", v)
        slot = (plot_r - plot_l) / n
        bar_half = slot / (1 + (plot.gap_width if kind != "line" else 150) / 100) / 2

        def fx(i):
            return plot_l + (plot_r - plot_l) * (i + 0.5) / n

        def fy(v):
            return plot_b - (plot_b - plot_t) * (float(v) - lo) / max(hi - lo, 1e-9)
        geo = {"l": plot_l, "r": plot_r, "t": plot_t, "b": plot_b, "slot": slot,
               "bar_half": bar_half, "lo": lo, "hi": hi}
        return c, fx, fy, tops, geo

    def _chart_body(self, s, x, y, w, h, kind, categories, series, value_fmt="#,##0",
                    highlight=None, bracket=None, callouts=None, ymin=None, ymax=None,
                    target=None, value_axis="auto", totals=True, direct_labels=True,
                    forecast_from=None, share_of=None):
        """グラフ本体+図上の注記。chart と chart_pair の共通部。
        bracket は図の上の専用の帯(0.42in)に描く(bar だけ)。target=(値, "目標 90% 以上") は破線。
        value_axis: "auto"=値ラベルが全部ある図(単系列 bar・合計付き stack)では軸と目盛線を消す
        totals/direct_labels/share_of/forecast_from は stack 向け(docstring は chart() を参照)。"""
        from lxml import etree
        if kind == "waterfall":
            if any(v is not None for v in (highlight, bracket, target, ymin, ymax)):
                self._warn("waterfall では highlight / bracket / target / ymin / ymax は使えない(無視した)")
            cw = min(w, 9.5)
            self._waterfall(s, x + (w - cw) / 2, y, cw, h, categories, series[0][1], callouts=callouts)
            return
        n = max(len(categories), 1)
        single_bar = (kind == "bar" and len(series) == 1)
        if value_axis == "auto":
            value_axis = "hide" if (single_bar or (kind == "stack" and totals)) else "show"
        top_pad = 0.42 if (bracket and kind == "bar") else 0.0
        if forecast_from is not None:
            top_pad += 0.30
        headroom = 1.25 if callouts else 1.0
        need_top = 0.0
        if value_axis == "hide":
            need_top = 0.30 + (0.42 if callouts else 0.0)
        pad_r = None
        dl_labels = None
        if kind == "stack" and direct_labels:
            names = [str(nm) for nm, _ in series]
            lab_w = max(_jlen(nm) for nm in names) * 12 / 72 * CHAR_EM + 0.30
            if lab_w <= 1.8:
                pad_r = lab_w + 0.10
                dl_labels = names
            else:
                self._warn("stack の系列名が長く直接ラベルにできない。凡例に戻した(系列名を短くする)")
        cy, chh = y + top_pad, h - top_pad
        c, fx, fy, tops, geo = self._native_chart(s, x, cy, w, chh, kind, categories, series, value_fmt,
                                                  highlight, headroom, ymin, ymax, value_axis, need_top,
                                                  pad_r, legend=(False if dl_labels else None))
        slot, bh = geo["slot"], geo["bar_half"]
        if bracket and kind == "bar":
            i0, i1, blab = bracket
            bx0, bx1 = fx(i0) - slot * 0.40, fx(i1) + slot * 0.40
            self.text(s, bx0 - 0.8, y, (bx1 - bx0) + 1.6, 0.26,
                      [[(blab, 12, True, BLUE_TX)]], align="center")
            by = y + 0.30
            self.rect(s, bx0, by, bx1 - bx0, 0.018, BLUE_TX)
            self.rect(s, bx0, by, 0.018, 0.09, BLUE_TX)
            self.rect(s, bx1 - 0.018, by, 0.018, 0.09, BLUE_TX)
        elif bracket:
            self._warn("bracket は棒グラフ(bar)だけ。line/stack では描かない")
        if forecast_from is not None:                 # 実績 | 計画 の区切り(縦の破線+上の帯のラベル)
            if not 1 <= forecast_from < n:
                self._warn("forecast_from はカテゴリ index 1〜n-1")
            else:
                dx = (fx(forecast_from - 1) + fx(forecast_from)) / 2
                self._dash_line(s, dx, geo["t"], dx, geo["b"], GREY, 0.75)
                band_y = y + (0.42 if (bracket and kind == "bar") else 0.0)
                self.text(s, dx - 1.3, band_y, 1.2, 0.26, [[("実績", 12, False, INK2)]], align="right")
                self.text(s, dx + 0.10, band_y, 1.2, 0.26, [[("計画", 12, False, INK2)]], align="left")
                if callouts and any(i in (forecast_from - 1, forecast_from) for i, _ in callouts):
                    self._warn("境目の両隣の棒に注記がある。区切りの破線と重なる可能性を目視で確かめる")
        tline = None
        if target is not None:
            tv, tlab = (target if isinstance(target, (tuple, list)) else (target, None))
            tline = fy(tv)
            x0, x1 = geo["l"], geo["r"]
            self._dash_line(s, x0, tline, x1, tline, INK2, 1.25)
            if tlab:
                if not re.search(r"(以上|以下|未満|超|上限|下限|まで|内)", str(tlab)):
                    self._warn(f"目標線のラベル「{tlab}」に向き(以上 / 以下)が無い")
                # ラベルは線の右端。右 1/3 に注記があれば左端に回す。データがその端で線より上なら線の下に
                right_busy = any(i >= n * 2 / 3 - 0.01 for i, _ in (callouts or []))
                end_v = (tops[0] if right_busy else tops[-1]) if tops else tv
                ly = tline + 0.06 if fy(end_v) < tline else tline - 0.30
                if right_busy:
                    self.text(s, x0 + 0.05, ly, 2.0, 0.26, [[(tlab, 12, True, INK2)]], align="left")
                else:
                    self.text(s, x1 - 2.0, ly, 2.0, 0.26, [[(tlab, 12, True, INK2)]], align="right")
        if kind == "stack":
            plot = c.plots[0]
            cum = [[0.0] * n for _ in series]        # 区間の下端
            run_ = [0.0] * n
            for si, (_, vals) in enumerate(series):
                for j in range(n):
                    v = float(vals[j] or 0) if j < len(vals) else 0.0
                    cum[si][j] = run_[j]
                    run_[j] += v
            # 合計(棒の上)
            if totals:
                for j in range(n):
                    self.text(s, fx(j) - slot * 0.45, fy(tops[j]) - 0.30, slot * 0.9, 0.28,
                              [[(f"{tops[j]:,.0f}", 12, True, INK)]], align="center", anchor="bottom")
            # 薄い区間(0.22in 未満)のラベルは棒の右外へ。share_of の区分は「値(xx%)」の文字にする
            thin = {}
            for si, (ser, (_, vals)) in enumerate(zip(plot.series, series)):
                for j in range(n):
                    v = float(vals[j] or 0) if j < len(vals) else 0.0
                    seg_h = fy(cum[si][j]) - fy(cum[si][j] + v)
                    if share_of is not None and si == share_of and v > 0 and tops[j] > 0:
                        pt = ser.points[j]
                        tf = pt.data_label.text_frame
                        tf.text = f"{v:,.0f}({v / tops[j] * 100:.0f}%)"
                        for p_ in tf.paragraphs:
                            for r_ in p_.runs:
                                _set_font(r_, 12, True, "#FFFFFF" if STACK_COLORS[si % 4] == BLUE else INK)
                    if 0 < v and seg_h < 0.22:
                        dls_el = ser._element.find(qn("c:dLbls"))
                        if dls_el is not None:
                            d_el = etree.SubElement(dls_el, qn("c:dLbl"))
                            etree.SubElement(d_el, qn("c:idx")).set("val", str(j))
                            etree.SubElement(d_el, qn("c:delete")).set("val", "1")
                            dls_el.remove(d_el)
                            dls_el.insert(0, d_el)
                        thin[(si, j)] = (v, fy(cum[si][j] + v / 2))
            gap_w = slot - 2 * bh
            for (si, j), (v, yc) in thin.items():
                if j == n - 1 and dl_labels:
                    continue                           # 最後の棒は系列名と一緒に書く
                if gap_w < 0.5:
                    self._warn(f"stack の薄い区間({series[si][0]} {v:g})のラベルを置く余白が無い(消した)")
                    continue
                self.rect(s, fx(j) + bh + 0.06, yc - 0.07, 0.14, 0.14, STACK_COLORS[si % len(STACK_COLORS)])
                self.text(s, fx(j) + bh + 0.26, yc - 0.13, gap_w - 0.28, 0.26,
                          [[(f"{v:,.0f}", 12, False, INK2)]], align="left", anchor="middle")
            # 系列名を最後の棒の右に直接書く(凡例の代わり)
            if dl_labels:
                j = n - 1
                labs = []
                for si, (nm, vals) in enumerate(series):
                    v = float(vals[j] or 0) if j < len(vals) else 0.0
                    yc = fy(cum[si][j] + v / 2)
                    txt = nm if (si, j) not in thin else f"{nm} {v:,.0f}"
                    labs.append([yc, txt, si])
                labs.sort(key=lambda t: -t[0])          # 下から上へ
                for k in range(1, len(labs)):
                    if labs[k - 1][0] - labs[k][0] < 0.30:
                        labs[k][0] = labs[k - 1][0] - 0.30
                for yc, txt, si in labs:
                    y_true = fy(cum[si][j] + (float(series[si][1][j] or 0)) / 2)
                    lx = fx(j) + bh + 0.12
                    if abs(yc - y_true) > 0.02:          # 押し上げたものだけ引き出し線
                        self._line_seg(s, fx(j) + bh + 0.02, y_true, lx - 0.02, yc, GREY, 0.75)
                    self.text(s, lx, yc - 0.14, pad_r - 0.10, 0.28, [[(txt, 12, True, INK)]],
                              align="left", anchor="middle")
        for idx, text in (callouts or []):
            top_y = fy(tops[idx])
            lift = 0.70 if (single_bar or (kind == "stack" and totals)) else 0.42   # 値ラベル(合計)の上
            ty = max(top_y - lift, cy + 0.02)
            tw = max(2.4, slot * 1.6)
            if kind == "line" and tline is not None and ty - 0.04 < tline < ty + 0.34:
                ty = top_y + 0.14                                           # 目標線と重なるなら点の下へ
            tb = self.text(s, fx(idx) - tw / 2, ty, tw, 0.30, [[(text, 12, True, BLUE_TX)]], align="center")
            if kind != "line" and tline is not None and ty - 0.04 < tline < ty + 0.34:
                tb.fill.solid()                                             # 線の上に白地で載せる(棒から離さない)
                tb.fill.fore_color.rgb = _rgb("#FFFFFF")

    def chart(self, title, kind, categories, series, lead=None, unit=None,
              value_fmt="#,##0", source=None, subtitle=None, highlight=None, bracket=None,
              callouts=None, ymin=None, ymax=None, target=None, takeaway=None, note=None,
              points=None, sowhat=None, sowhat_intro=None, value_axis="auto", totals=True,
              direct_labels=True, forecast_from=None, share_of=None):
        """⑦ グラフ(全幅)。結論はリードで言い切り、注目点は図の上に書く。
        subtitle: 定義行「何の数字か+集計の単位+期間(年度)」(13pt)
        callouts: [(カテゴリ index, 注記の一文), ...] — その棒(点)の真上に 12pt 青太字。lead の根拠の数字を 1 つ置く
        target  : (値, "目標 90% 以上") — しきい値の破線。ラベルに向き(以上 / 以下)を書く
        value_axis: "auto"(既定)=値ラベルが全部ある図(単系列 bar・合計付き stack)では値軸と目盛線を消す。
                    "show" / "hide" で固定
        stack 向け: totals=True で各棒の上に合計 / direct_labels=True で系列名を最後の棒の右に直接書く(凡例なし)
                    share_of=i で区分 i のラベルを「3,600(63%)」の形に / forecast_from=i で i 以降を「計画」として区切る
        points= / sowhat= は廃止(右の「ポイント」欄は使わない)。渡すとエラー"""
        if points or sowhat or sowhat_intro:
            raise TypeError("points/sowhat は廃止。結論は lead、注目点は callouts/bracket/highlight に書く")
        s = self._slide(LAYOUT_BODY, title=title)
        self._kind("chart")
        y0 = self._lead(s, lead)
        cw = self.CW
        if subtitle:
            # 定義行: 何のグラフか+単位・期間(▍はリード専用なので付けない)
            self.text(s, self.MX, y0, cw - 2.3, 0.3, [[(subtitle, 13, True, INK)]])
            if unit:
                self.text(s, self.MX + cw - 2.2, y0 + 0.02, 2.2, 0.3,
                          [[(f"(単位:{unit})", 11, False, INK2)]], align="right")
            ch_y = y0 + 0.46
        else:
            if unit:
                self.text(s, self.MX + cw - 2.2, y0 - 0.05, 2.2, 0.3,
                          [[(f"(単位:{unit})", 11, False, INK2)]], align="right")
            ch_y = y0 + 0.25
        bot_lim = self._body_bot(source, note) - 0.10
        ch_h = bot_lim - ch_y - (0.64 if takeaway else 0.0)
        self._chart_body(s, self.MX, ch_y, cw, ch_h, kind, categories, series, value_fmt,
                         highlight, bracket, callouts, ymin, ymax, target, value_axis, totals,
                         direct_labels, forecast_from, share_of)
        self._takeaway_band(s, ch_y + ch_h, takeaway)
        self._source(s, source, note)
        return s

    def chart_pair(self, title, specs, lead=None, source=None, note=None, takeaway=None):
        """⑦' 2 つの問いを左右に並べるグラフ(コンサル実物の作法)。
        specs: [dict(head="Q: 問い(必須)", subtitle=, kind=, categories=, series=, unit=,
                     value_fmt=, highlight=, bracket=, callouts=, ymin=, ymax=, target=), ...] 2 つ
        各列は col_label(見出し 14pt+下線)の下にグラフ。結論はリードで言い切る。"""
        s = self._slide(LAYOUT_BODY, title=title)
        self._kind("chart")
        y0 = self._lead(s, lead)
        n = max(len(specs), 1)
        gap = 0.45
        cw = (self.CW - gap * (n - 1)) / n
        bot = self._body_bot(source, note) - 0.10 - (0.64 if takeaway else 0.0)
        for i, sp in enumerate(specs):
            if not sp.get("head"):
                raise ValueError("chart_pair は各列に head(Q: 問い)が必須")
            x = self.MX + i * (cw + gap)
            ry = self.col_label(s, x, y0, cw, sp["head"], sub=sp.get("subtitle", sp.get("sub")))
            if sp.get("unit"):
                self.text(s, x + cw - 2.0, ry - 0.10, 2.0, 0.26,
                          [[(f"(単位:{sp['unit']})", 11, False, INK2)]], align="right")
            ch_y = ry + 0.16
            self._chart_body(s, x, ch_y, cw, bot - ch_y, sp["kind"], sp["categories"], sp["series"],
                             sp.get("value_fmt", "#,##0"), sp.get("highlight"), sp.get("bracket"),
                             sp.get("callouts"), sp.get("ymin"), sp.get("ymax"), sp.get("target"),
                             sp.get("value_axis", "show"), sp.get("totals", True),
                             sp.get("direct_labels", True), sp.get("forecast_from"), sp.get("share_of"))
        self._takeaway_band(s, bot, takeaway)
        self._source(s, source, note)
        return s

    def _takeaway_band(self, s, y, takeaway):
        """図解の下に敷く「まとめ帯」(takeaway が無ければ何もしない)"""
        if not takeaway:
            return
        self._takeaway_pages.append(self._page)
        self.rect(s, self.MX, y + 0.18, self.CW, 0.50, LTBLUE)
        self.text(s, self.MX + 0.28, y + 0.18, self.CW - 0.56, 0.50,
                  [[(takeaway, 14, True, BLUE_TX)]], anchor="middle")

    def _reserve_takeaway(self, takeaway, source=None, note=None):
        """軽い結論帯(takeaway)の場所を本文下に確保する。全型共通スロット。
        型メソッド冒頭で _tb = self._reserve_takeaway(...) と呼び、
        末尾で self._finish_takeaway(s, _tb, takeaway)。戻り値は元の BODY_BOT。"""
        if not takeaway:
            return None
        orig = self.BODY_BOT
        bot = orig
        if note:
            bot = min(bot, self.H - 1.06)
        elif source:
            bot = min(bot, self.H - 0.86)
        self.BODY_BOT = bot - 0.72
        return orig

    def _finish_takeaway(self, s, orig, takeaway):
        if orig is None:
            return
        self._takeaway_band(s, self.BODY_BOT - 0.12, takeaway)
        self.BODY_BOT = orig

    def _waterfall(self, s, x, y, w, h, categories, values, callouts=None):
        """滝グラフ(ネイティブ図形で描画)。増減プロセスの分解(利益ブリッジ等)専用。

        values: [期首実額, ±増減..., 期末実額(None なら自動計算)]。期首・期末は正の実額。
        要因は影響の大きい順(時系列・工程の順は例外)で 8 本まで。5% 未満の要因は「その他」にまとめる。
        色は 期首/期末=BLUE、増加=TEAL、減少=RED。ラベルは +35 / △20(増加の文字は INK。TEAL は薄い)。
        全バーにデータラベルが付くため値軸・グリッドは引かない。callouts=[(index, 一文)] で棒の上に注記。
        """
        vals = list(values)
        if vals[-1] is None:
            vals[-1] = vals[0] + sum(vals[1:-1])
        elif abs(vals[0] + sum(vals[1:-1]) - vals[-1]) > 1e-6:
            self._warn(f"滝グラフの期末 {vals[-1]:g} が期首+増減の合計 {vals[0] + sum(vals[1:-1]):g} と合わない")
        if vals[0] is None or vals[0] <= 0 or vals[-1] <= 0:
            raise ValueError("滝グラフは正の期首・期末が必要: [期首, ±増減..., 期末(None可)]")
        if len(vals) - 2 > 8:
            self._warn(f"滝グラフの要因が {len(vals) - 2} 本。8 本以内にし、小さい要因は「その他」にまとめる")
        n = len(vals)
        levels, cum = [(0.0, float(vals[0]))], float(vals[0])
        for d in vals[1:-1]:
            b, t = (cum, cum + d) if d >= 0 else (cum + d, cum)
            levels.append((b, t))
            cum += d
        levels.append((0.0, float(vals[-1])))
        ymax = max(t for _, t in levels) * 1.18
        lab_h, cat_h = (0.60 if callouts else 0.26), 0.34
        plot_top = y + lab_h
        plot_h = h - lab_h - cat_h
        base_y = plot_top + plot_h

        def py(v):
            return plot_top + plot_h * (1 - v / ymax)

        bw = w / (n + (n - 1) * 0.5)     # ギャップはバー幅の1/2
        step = bw * 1.5
        for i, (b, t) in enumerate(levels):
            bx = x + i * step
            is_total = (i == 0 or i == n - 1)
            d = vals[i]
            fill = BLUE if is_total else (TEAL if d >= 0 else RED)
            self.rect(s, bx, py(t), bw, max(py(b) - py(t), 0.02), fill)
            if is_total:
                lab, lc = f"{vals[i]:,.0f}", INK
            elif d >= 0:
                lab, lc = f"+{d:,.0f}", INK          # TEAL の文字は白地で薄い。向きは「+」で示す
            else:
                lab, lc = f"△{abs(d):,.0f}", RED
            self.text(s, bx - step * 0.25, py(t) - 0.26, bw + step * 0.5, 0.24,
                      [[(lab, 12, True, lc)]], align="center", anchor="bottom")
            self.text(s, bx - step * 0.25, base_y + 0.06, bw + step * 0.5, cat_h - 0.06,
                      [[(str(categories[i]), 12, False, INK)]], align="center")
            if i < n - 1:  # 累積レベルを次の棒へつなぐ細線
                lev = t if (is_total or d >= 0) else b
                self.rect(s, bx + bw, py(lev) - 0.005, step - bw, 0.01, GREY)
        for idx, text in (callouts or []):
            b_, t_ = levels[idx]
            bx = x + idx * step
            tw = max(2.4, step * 1.6)
            self.text(s, bx + bw / 2 - tw / 2, py(t_) - 0.26 - 0.32, tw, 0.30,
                      [[(text, 12, True, BLUE_TX)]], align="center")
        self.rect(s, x, base_y - 0.005, w, 0.01, GREY)

    # ---------------------------------------------------------------- ⑧ ネイティブ表
    def table(self, title, headers, rows, lead=None, col_widths=None,
              first_col_header=True, font_size=13, unit=None, source=None,
              highlight_row=None, note=None, takeaway=None, zebra=True, sections=None):
        """表。zebra=True(既定): 偶数行を淡く塗る(数値の表向き)。zebra=False: 塗りをやめて行間に
        細い横罫を引く(論証テーブル向き。縦の白線も消える)。見出し行 BLUE・1 列目 PALE・強調行 LTBLUE は共通。
        sections: 表の下に置く軽い段組 [(見出し, [items]), ...] 1〜2 個(items は section_block と同じ)。
                  表に入れると列が立たない別の群(例「版は持たない 2 つ」)を分けて置く。lead は表と下段の両方を言う。
        セルの折り返しを見積もり、4 行以上になるセルや高さ超過は警告する。"""
        s = self._slide(LAYOUT_BODY, title=title)
        self._kind("table")
        _tb = self._reserve_takeaway(takeaway, source, note)
        y0 = self._lead(s, lead)
        if unit:
            self.text(s, self.MX, y0 - 0.05, self.CW, 0.3,
                      [[(f"(単位:{unit})", 10, False, INK2)]], align="right")
            ty = y0 + 0.26
        else:
            ty = y0 + 0.08
        ncols, nrows = len(headers), len(rows) + 1
        if font_size < 12:
            self._warn(f"表の文字 {font_size}pt は小さすぎる。12pt に上げた(行が多いなら 2 枚に分ける)")
            font_size = 12
        avail_h = self._body_bot(source, note) - 0.10 - ty
        sec_h = 0.0
        if sections:
            n_sec = len(sections)
            sw = (self.CW - 0.5 * (n_sec - 1)) / n_sec
            sec_h = max(self._section_need(sw, items, 14) for _, items in sections) + 0.30
            avail_h -= sec_h
        row_h = min(0.58, avail_h / nrows)
        if row_h < 0.34:
            self._warn(f"表の行が {nrows} 行で詰まりすぎ(行高 {row_h:.2f}in)。行を減らすか 2 枚に分ける")
        # 折り返しの見積もり(列幅 × 文字サイズ)。4 行以上のセル、表の高さ超過を警告
        _w = [self.CW * cw_ / sum(col_widths) for cw_ in col_widths] if col_widths else [self.CW / ncols] * ncols
        est_h = row_h
        for row in rows:
            mx_lines = 1
            for cji, val in enumerate(row[:ncols]):
                nl = _nlines(str(val), _cpl(_w[cji], font_size, pad=0.16))
                if nl >= 4:
                    self._warn(f"表のセルが {nl} 行になる見込み: {str(val)[:18]}…(列を広げるか文を削る)")
                mx_lines = max(mx_lines, nl)
            est_h += max(row_h, font_size / 72 * 1.3 * mx_lines + 0.10)
        if est_h > avail_h + 0.05:
            self._warn(f"表の高さが本文領域を {est_h - avail_h:.2f}in 超える見込み。行を減らすか文を削る")
        gf = s.shapes.add_table(nrows, ncols, Inches(self.MX), Inches(ty),
                                Inches(self.CW), Inches(row_h * nrows))
        tbl = gf.table
        tbl.first_row = False
        tbl.horz_banding = False
        if col_widths:
            total = sum(col_widths)
            for i, cwd in enumerate(col_widths):
                tbl.columns[i].width = Emu(int(Inches(self.CW) * cwd / total))
        for r in range(nrows):
            tbl.rows[r].height = Inches(row_h)
        body_size = font_size

        def put(cell, txt, size, bold, color, fill, align):
            cell.fill.solid()
            cell.fill.fore_color.rgb = _rgb(fill if fill else "#FFFFFF")
            cell.margin_left = Inches(0.08)
            cell.margin_right = Inches(0.08)
            cell.margin_top = cell.margin_bottom = Inches(0.02)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf = cell.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = _ALIGN[align]
            r = p.add_run()
            r.text = str(txt)
            _set_font(r, size, bold, color)

        for cji, htxt in enumerate(headers):
            put(tbl.cell(0, cji), htxt, body_size, True, "#FFFFFF", BLUE, "center")
            if not zebra:
                self._cell_borders(tbl.cell(0, cji), bottom=None, top=None)
        def _isnum(t):
            return bool(_NUMERIC_RE.match(str(t).replace(",", "").replace(" ", "")))
        # 列の揃え: 本文セルの過半が数値なら、その列は全セル右揃え(「2 営業日」等の混在で揺れない)
        col_num = []
        for cji in range(ncols):
            vals = [r[cji] for r in rows if cji < len(r)]
            col_num.append(bool(vals) and sum(_isnum(v) for v in vals) * 2 > len(vals))
        for ri, row in enumerate(rows, start=1):
            zfill = (BG if (zebra and ri % 2 == 0) else None)
            is_hl = (highlight_row is not None and ri - 1 == highlight_row)
            last = (ri == nrows - 1)
            for cji, val in enumerate(row):
                txt = str(val)
                is_num = bool(_NUMERIC_RE.match(txt.replace(",", "").replace(" ", "")))
                color = INK
                bold = is_hl
                fill = LTBLUE if is_hl else zfill
                align = "right" if (col_num[cji] and cji > 0) else "left"
                if cji == 0 and first_col_header:
                    bold = True
                    fill = LTBLUE if is_hl else PALE
                    align = "left"
                if txt.startswith(("△", "-", "−", "▼")) and is_num:
                    color = RED
                put(tbl.cell(ri, cji), txt, body_size, bold, color, fill, align)
                if not zebra:    # 塗りの代わりに横罫(最終行の下は GREY で閉じる)。縦線なし
                    self._cell_borders(tbl.cell(ri, cji), bottom=((GREY, 0.75) if last else (GRID, 0.75)))
        if sections:
            n_sec = len(sections)
            sw = (self.CW - 0.5 * (n_sec - 1)) / n_sec
            sy = ty + row_h * nrows + 0.28
            bot_s = self._body_bot(source, note) - 0.05
            for i, (head, items) in enumerate(sections):
                self.section_block(s, self.MX + i * (sw + 0.5), sy, sw, bot_s, head, items)
        self._finish_takeaway(s, _tb, takeaway)
        self._source(s, source, note)
        return s

    # ---------------------------------------------------------------- 関門表(gate_table)
    @staticmethod
    def _cell_borders(cell, bottom=None, top=None):
        """セルの罫線。bottom/top は (色, pt) か None(線なし)。左右は常に線なし。"""
        from lxml import etree
        tcPr = cell._tc.get_or_add_tcPr()
        for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
            for el in tcPr.findall(qn(tag)):
                tcPr.remove(el)
        specs = [("a:lnL", None), ("a:lnR", None), ("a:lnT", top), ("a:lnB", bottom)]
        for i, (tag, spec) in enumerate(specs):
            el = etree.Element(qn(tag))
            if spec is None:
                el.set("w", "0")
                etree.SubElement(el, qn("a:noFill"))
            else:
                el.set("w", str(int(Pt(spec[1]))))
                sf = etree.SubElement(el, qn("a:solidFill"))
                etree.SubElement(sf, qn("a:srgbClr")).set("val", spec[0].lstrip("#"))
            tcPr.insert(i, el)

    def light_table(self, s, x, y, w, col_heads, rows, label_w=2.4, size=15, head_size=16,
                    bold_rows=(), row_h=0.52):
        """塗りのない軽い表(関門表などの事実の格子)。先頭列は行見出し(INK2)、
        見出し行の下に BLUE の罫線、行間は GRID の細線。縦線なし。戻り値は下端の y。
        rows: [(行見出し, [セル...]), ...]。bold_rows に入れた行 index のセルは太字。"""
        ncols = len(col_heads) + 1
        nrows = len(rows) + 1
        cw = (w - label_w) / len(col_heads)
        heights = [row_h]
        for lab, cells in rows:
            mx = 1
            for c, v in enumerate(cells):
                nl = _nlines(str(v), _cpl(cw, size, pad=0.22))
                mx = max(mx, nl)
            if mx >= 3:
                self._warn(f"関門表のセルが {mx} 行になる見込み: {str(cells)[:24]}…(短くするか関門を減らす)")
            heights.append(max(row_h, size / 72 * 1.25 * mx + 0.16))
        gf = s.shapes.add_table(nrows, ncols, Inches(x), Inches(y), Inches(w), Inches(sum(heights)))
        tbl = gf.table
        tbl.first_row = False
        tbl.horz_banding = False
        tbl.columns[0].width = Inches(label_w)
        for c in range(1, ncols):
            tbl.columns[c].width = Inches(cw)
        for r in range(nrows):
            tbl.rows[r].height = Inches(heights[r])

        def put(cell, txt, sz, bold, color, bottom, top=None):
            cell.fill.background()
            cell.margin_left = Inches(0.10)
            cell.margin_right = Inches(0.10)
            cell.margin_top = cell.margin_bottom = Inches(0.03)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf = cell.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            for k, piece in enumerate(str(txt).split(chr(10))):
                if k > 0:
                    br = p._p.add_br()
                    br.get_or_add_rPr().set("sz", str(int(sz * 100)))
                run = p.add_run()
                run.text = piece
                _set_font(run, sz, bold, color)
            self._cell_borders(cell, bottom=bottom, top=top)

        put(tbl.cell(0, 0), "", head_size, True, INK, (BLUE, 1.5))
        for c, hd in enumerate(col_heads, 1):
            put(tbl.cell(0, c), hd, head_size, True, INK, (BLUE, 1.5))
        for r, (lab, cells) in enumerate(rows, 1):
            last = (r == nrows - 1)
            bb = (GREY, 0.75) if last else (GRID, 0.75)
            put(tbl.cell(r, 0), lab, size - 1, False, INK2, bb)
            for c, v in enumerate(cells, 1):
                v2 = self._soft_break(v, _cpl(cw, size, pad=0.22))   # 語の切れ目で改行(1 字だけの行を避ける)
                put(tbl.cell(r, c), v2, size, (r - 1) in bold_rows, INK, bb)
        return y + sum(heights)

    def gate_table(self, title, gates, rows, lead=None, footnote=None, sections=None,
                   source=None, note=None, bold_row=None, takeaway=None, on_fail=None):
        """関門・条件・合格基準を「関門 × 測り方 / 合格ライン / 判定者」の軽い表にする型
        (ユーザーの理想形 2026-10-21。色箱+矢印の関門図はこれに置き換える)。

        gates   : 列見出し ["① 素の精度", "② 回帰なし", "③ 記憶の再生"](2〜4)
        rows    : [(行見出し, [関門ごとの値...]), ...] 例 何で測る / 合格ライン / 判定者(2〜4 行)
        bold_row: 太字にする行 index(既定: 「合格」「しきい」「基準」「条件」を含む行)
        footnote: 表の直下の ※ 注(校正の条件、通過後の流れなど。1〜2 行)
        sections: 下段の軽い段組 [(見出し, [items]), ...] 1〜2 個。items は section_block と同じ
                  ("文" → ●箇条 / ("steps", [...], 補足) → 1. 2. 3. の一文手順 / ("p", "文") → 平文)
                  **関門を通った直後に起きること(通報を受けたとき・版管理など)だけ**に使う。次頁への案内や
                  関門と別の手順を置かない。無ければ省く(表だけで 1 枚。行高が自動で広がる)
        on_fail : 不合格のときの扱い。文字列なら ※注の末尾に足す(例「落ちた版は改修して①から再評価する」)。
                  リストなら関門ごとの戻し先として 4 行目「不合格なら」を足す。無いと警告。
        書き方: 関門の見出しは「何を確かめるか」の名詞句で揃える(社内略語を使わない)。判定者には責任を
        負う人か会議体を 1 つだけ書き、LLM や判定モデルは「何で測る」の行に入れる。合格ラインは
        「以上 / 以下 / 未満」で向きを書く。
        lead は上の表と下の段組の両方を 1 文で言う(例「3 つの関門を通してから出し、出した後は 1 営業日で止める」)。"""
        s = self._slide(LAYOUT_BODY, title=title)
        self._kind("table")
        _tb = self._reserve_takeaway(takeaway, source, note)
        y0 = self._lead(s, lead)
        rows = [tuple(r) for r in rows]
        if isinstance(on_fail, (list, tuple)):
            rows = rows + [("不合格なら", list(on_fail))]
        elif on_fail:
            footnote = (str(footnote).rstrip("。") + "。" if footnote else "") + str(on_fail)
        else:
            self._warn("関門表に不合格のときの扱い(on_fail)が無い。差し戻し先か再判定を書く")
        for lab, cells in rows:
            if "判定" in str(lab):
                for v in cells:
                    if any(k in str(v) for k in ("モデル", "LLM", "ツール", "スクリプト")):
                        self._warn(f"判定者「{v}」は道具。人か会議体を書き、道具は『何で測る』に入れる")
        if bold_row is None:
            bold_row = next((i for i, (lab, _) in enumerate(rows)
                             if any(k in str(lab) for k in ("合格", "しきい", "基準", "条件"))), None)
        bold = (bold_row,) if bold_row is not None else ()
        row_h = 0.52
        if not sections:
            # 表が主役: 本文の高さに合わせて行高を広げる(上限 0.78in)
            fn_h = (12 / 72 * 1.4 * _nlines("※ " + str(footnote), _cpl(self.CW - 0.2, 12)) + 0.14) if footnote else 0.0
            avail = self._body_bot(source, note) - 0.25 - (y0 + 0.05) - fn_h
            row_h = max(0.52, min(0.78, avail / (len(rows) + 1)))
        yy = self.light_table(s, self.MX, y0 + 0.05, self.CW, gates, rows, bold_rows=bold, row_h=row_h)
        if footnote:
            fn = str(footnote)
            if not fn.startswith("※"):
                fn = "※ " + fn
            cpl = max(int((self.CW - 0.2) / (12 / 72)), 10)
            lines = max(1, -(-int(_jlen(fn) + 0.999) // cpl))
            fh = 12 / 72 * 1.4 * lines + 0.06
            self.text(s, self.MX, yy + 0.08, self.CW, fh, [[(fn, 12, False, INK2)]], line_spacing=1.3)
            yy += 0.08 + fh
        bot = self._body_bot(source, note) - 0.05
        if sections:
            n = len(sections)
            gap = 0.5
            w = (self.CW - gap * (n - 1)) / n
            sy = yy + 0.28
            if bot - sy < 1.2:
                self._warn(f"下段の段組の高さが {bot - sy:.2f}in しかない。行か注を減らす")
            for i, (head, items) in enumerate(sections):
                self.section_block(s, self.MX + i * (w + gap), sy, w, bot, head, items)
        self._finish_takeaway(s, _tb, takeaway)
        self._source(s, source, note)
        return s

    # ---------------------------------------------------------------- ⑪ 評価マトリクス(コンサル型)
    def matrix_eval(self, title, options, criteria, ratings, lead=None,
                    verdicts=None, verdict_header="位置づけ", recommend=None,
                    source=None, note=None, takeaway=None, legend=None):
        """選定・評価マトリクス(コンサル比較表)。候補を列に立て、左に比較軸の行を並べる。

        options : 列(候補)の名前リスト(最大4)
        criteria: 行(比較軸)の名前リスト
        ratings : 候補ごとの値リスト(ratings[i][j] = 候補i×軸j)。
                  値は "◎"/"○"/"△"/"×" の記号、短文テキスト、["箇条", ...] のリストを混在可
        verdicts: 最下段の位置づけ・結論行(省略可)
        recommend: 強調する候補列の index(0始まり)
        legend  : 記号の凡例。None なら記号を使った表に自動で付ける(「◎ 最適 ○ 良い △ 課題あり × 不可」)。
                  False で消す。文字列で差し替え。比較軸は「コストの低さ」のように ◎ が良い向きに書く。
        """
        s = self._slide(LAYOUT_BODY, title=title)
        self._kind("table")
        _tb = self._reserve_takeaway(takeaway, source, note)
        y0 = self._lead(s, lead)
        ty = y0 + 0.15
        ncols = 1 + len(options)
        nrows = 1 + len(criteria) + (1 if verdicts else 0)
        syms_used = {str(v) for col in ratings for v in col if str(v) in ("◎", "○", "△", "×")}
        if legend is None and syms_used:
            parts = [("◎", "最適"), ("○", "良い"), ("△", "課題あり"), ("×", "不可")]
            legend = "　".join(f"{k} {v}" for k, v in parts if k in syms_used)
        leg_h = 0.34 if legend else 0.0
        h_avail = self._body_bot(source, note) - 0.15 - ty - leg_h
        head_h = 0.50
        row_h = max(0.42, min(1.05, (h_avail - head_h) / (nrows - 1)))
        gf = s.shapes.add_table(nrows, ncols, Inches(self.MX), Inches(ty),
                                Inches(self.CW), Inches(head_h + row_h * (nrows - 1)))
        tbl = gf.table
        tbl.first_row = False
        tbl.horz_banding = False
        lab_w = 2.1
        opt_w = (self.CW - lab_w) / len(options)
        for i, wd in enumerate([lab_w] + [opt_w] * len(options)):
            tbl.columns[i].width = Emu(int(Inches(wd)))
        tbl.rows[0].height = Inches(head_h)
        for r in range(1, nrows):
            tbl.rows[r].height = Inches(row_h)

        def put(cell, paras, fill):
            cell.fill.solid()
            cell.fill.fore_color.rgb = _rgb(fill if fill else "#FFFFFF")
            cell.margin_left = cell.margin_right = Inches(0.08)
            cell.margin_top = cell.margin_bottom = Inches(0.04)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf = cell.text_frame
            tf.word_wrap = True
            for k, (runs, align) in enumerate(paras):
                p = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
                p.alignment = _ALIGN[align]
                p.space_after = Pt(2)
                p.line_spacing = 1.1
                for txt, size, bold, color in runs:
                    r = p.add_run()
                    r.text = str(txt)
                    _set_font(r, size, bold, color)

        SYM = {"◎": (BLUE, True), "○": (INK, False), "△": (INK2, False),
               "×": (INK2, False), "－": (INK2, False), "-": (INK2, False)}

        def body_paras(val):
            if isinstance(val, (list, tuple)):
                return [([("●", 9, False, INK), ("  " + str(it), 14, False, INK)], "left")
                        for it in val]
            v = str(val)
            if v in SYM:
                color, bold = SYM[v]
                return [([(v, 22, bold, color)], "center")]
            return [([(v, 14, False, INK)], "center")]

        # ヘッダ行(候補)。左上コーナーは空白のまま
        put(tbl.cell(0, 0), [([("", 11, True, "#FFFFFF")], "center")], None)
        for j, opt in enumerate(options, start=1):
            is_rec = (recommend is not None and j - 1 == recommend)
            put(tbl.cell(0, j), [([(opt, 14, True, "#FFFFFF")], "center")],
                NAVY if is_rec else BLUE)
        for i, cri in enumerate(criteria):
            ri = i + 1
            zebra = BG if ri % 2 == 0 else None
            put(tbl.cell(ri, 0), [([(cri, 14, True, "#FFFFFF")], "center")], BLUE)
            for j in range(len(options)):
                is_rec = (recommend is not None and j == recommend)
                put(tbl.cell(ri, j + 1), body_paras(ratings[j][i]),
                    LTBLUE if is_rec else zebra)
        if verdicts:
            ri = nrows - 1
            put(tbl.cell(ri, 0), [([(verdict_header, 14, True, "#FFFFFF")], "center")], BLUE)
            for j, v in enumerate(verdicts):
                is_rec = (recommend is not None and j == recommend)
                put(tbl.cell(ri, j + 1),
                    [([(v, 14, is_rec, BLUE_TX if is_rec else INK)], "center")],
                    LTBLUE if is_rec else PALE)
        if legend:
            self.text(s, self.MX, ty + head_h + row_h * (nrows - 1) + 0.08, self.CW, 0.26,
                      [[(str(legend), 12, False, INK2)]], align="right")
        self._finish_takeaway(s, _tb, takeaway)
        self._source(s, source, note)
        return s

    # ---------------------------------------------------------------- ⑫ 2軸マトリクス(ポジショニング)
    def matrix_2x2(self, title, x_axis, y_axis, items, lead=None,
                   quadrants=None, emphasize=None, source=None, note=None,
                   takeaway=None, highlight=None):
        """2軸ポジショニングマップ(点+ラベル)。

        x_axis / y_axis: (低ラベル, 高ラベル, 軸名)。横軸の端には括弧で定義を添えてよい
                         (例 "高(既存の検索基盤で作れる)")。縦軸の端は短い語。軸名は和文で短く(縦書きになる)
        items    : [(名前, x, y)] x,y は 0.0〜1.0(右上が高・高)。6 個まで。中心線の近く(±0.03)は警告
        quadrants: [左下, 右下, 左上, 右上] の象限名(名前か打ち手に揃え、言いさしにしない)
        emphasize: 強調する象限 index(0=左下,1=右下,2=左上,3=右上)。その象限だけ淡青に塗る
        highlight: 強調する項目名のリスト(省略時は emphasize 象限に入る項目)。強調=BLUE の点+青太字、他=GREY の点+細字
        位置が目安なら note= に「位置は説明用の目安」と軸の定義を書く。
        """
        s = self._slide(LAYOUT_BODY, title=title)
        self._kind("diagram")
        _tb = self._reserve_takeaway(takeaway, source, note)
        y0 = self._lead(s, lead)
        px = self.MX + 1.05
        py = y0 + 0.05
        pw = self.CW - 1.05
        ph = self._body_bot(source, note) - 0.45 - py
        if len(items) > 6:
            self._warn(f"2 軸マトリクスの項目が {len(items)} 個。6 個以内にする")
        # 象限(強調象限だけ淡青)。十字と外枠は GREY の実線
        if emphasize is not None:
            qx, qy = [(0, 1), (1, 1), (0, 0), (1, 0)][emphasize]
            self.rect(s, px + qx * pw / 2, py + qy * ph / 2, pw / 2, ph / 2, LTBLUE)
        self.rect(s, px, py, pw, ph, None, line=GREY, lw=1.0)
        self.rect(s, px + pw / 2 - 0.006, py, 0.012, ph, GREY)
        self.rect(s, px, py + ph / 2 - 0.006, pw, 0.012, GREY)
        if quadrants:
            pos = [(px + 0.15, py + ph - 0.42, "left"), (px + pw / 2 + 0.15, py + ph - 0.42, "left"),
                   (px + 0.15, py + 0.12, "left"), (px + pw / 2 + 0.15, py + 0.12, "left")]
            for qi, q in enumerate(quadrants):
                if q:
                    qx_, qy_, al = pos[qi]
                    bold = (emphasize == qi)
                    self.text(s, qx_, qy_, pw / 2 - 0.3, 0.35,
                              [[(q, 12, bold, BLUE_TX if bold else INK2)]], align=al)
        # 軸ラベル: 横軸は 低 ─ 軸名 ─ 高、縦軸は 大(左上の外)/ 小(左下の外)+縦書きの軸名
        lo_x, hi_x, name_x = x_axis
        lo_y, hi_y, name_y = y_axis
        end_w = min(3.2, pw * 0.30)
        for t in (lo_x, hi_x):
            if _jlen(t) * 12 / 72 * CHAR_EM > end_w:
                self._warn(f"横軸の端ラベル「{t}」が長すぎる(幅 {end_w:.1f}in)。定義を短くする")
        self.text(s, px, py + ph + 0.08, end_w, 0.3, [[(lo_x, 12, False, INK2)]])
        self.text(s, px + pw - end_w, py + ph + 0.08, end_w, 0.3,
                  [[(hi_x, 12, False, INK2)]], align="right")
        self.text(s, px + end_w, py + ph + 0.08, pw - 2 * end_w, 0.32,
                  [[(name_x, 12, True, INK)]], align="center")
        self.text(s, self.MX - 0.1, py - 0.02, 1.05, 0.30, [[(hi_y, 12, False, INK2)]], align="right")
        self.text(s, self.MX - 0.1, py + ph - 0.30, 1.05, 0.30, [[(lo_y, 12, False, INK2)]], align="right")
        nh = _jlen(name_y) * 12 / 72 * 1.15 + 0.2
        if nh > ph - 0.9:
            self._warn(f"縦軸名「{name_y}」が長い。縦書きで端の語にかかる(短くする)")
        ytb = self.text(s, px - 0.45 - nh / 2, py + ph / 2 - 0.20, nh, 0.40,
                        [[(name_y, 12, True, INK)]], align="center", anchor="middle", wrap=False)
        ytb.rotation = 270   # 下から上に読む(PDF 変換でも崩れない)
        # 項目: 正確な座標に点、右にラベル。強調は BLUE+青太字、他は GREY+細字
        hl = set(highlight) if highlight else None
        q_of = lambda ix, iy: (0 if ix < 0.5 else 1) + (2 if iy >= 0.5 else 0)
        placed = []
        any_hl = False
        for item in items:
            name, ix, iy = item[:3]
            is_hl = (name in hl) if hl is not None else (emphasize is not None and q_of(ix, iy) == emphasize)
            any_hl = any_hl or is_hl
            if abs(ix - 0.5) < 0.03 or abs(iy - 0.5) < 0.03:
                self._warn(f"項目「{name}」が中心線の上(±0.03)。どちらの象限か決めて座標を動かす")
            cx = px + ix * pw
            cy = py + (1 - iy) * ph
            d = 0.20
            self.rect(s, cx - d / 2, cy - d / 2, d, d, BLUE if is_hl else GREY, shape=MSO_SHAPE.OVAL)
            tw = _jlen(name) * 13 / 72 * CHAR_EM + 0.15
            right_edge = px + pw if ix >= 0.5 else px + pw / 2
            if cx + 0.18 + tw > right_edge - 0.05:       # 象限の右端を越えるなら点の左に右揃え
                lx, al = cx - 0.18 - tw, "right"
            else:
                lx, al = cx + 0.18, "left"
            ly = cy - 0.16
            for (ox, oy, ow) in placed:                   # 近すぎるラベルは警告
                if abs(ly - oy) < 0.30 and (lx < ox + ow and ox < lx + tw):
                    self._warn(f"項目「{name}」のラベルが他の項目と重なる。座標をずらす")
            placed.append((lx, ly, tw))
            self.text(s, lx, ly, tw, 0.32, [[(name, 13, is_hl, BLUE_TX if is_hl else INK)]],
                      align=al, anchor="middle")
        if not any_hl:
            self._warn("2 軸マトリクスで強調する項目が無い(emphasize か highlight を渡す)")
        self._finish_takeaway(s, _tb, takeaway)
        self._source(s, source, note)
        return s

    # ---------------------------------------------------------------- ⑬ 関係図(ラベル付き矢印)
    def flow(self, title, nodes, arrows=None, lead=None, band=None, source=None, note=None):
        """ノード(箱)をラベル付き矢印でつなぐ関係図。直列の役割分担などに使う。

        nodes : [(見出し, 小見出し|None, [項目...]), ...] 2〜3個
        arrows: 矢印の上に載せるラベル ["課題を引き渡す", ...](len(nodes)-1 個)
        band  : 下に敷く共通基盤バーの文言(省略可)
        """
        s = self._slide(LAYOUT_BODY, title=title)
        self._kind("diagram")
        y0 = self._lead(s, lead) - 0.10
        n = len(nodes)
        gap = 1.60 if arrows else 1.20   # 箱を広く、隙間は狭く(ユーザー手直し 2026-10-16)
        w = (self.CW - gap * (n - 1)) / n
        band_h = 0.50                      # 他の型の結論帯と同じ高さ
        bot = self._body_bot(source, note) - (band_h + 0.30 if band else 0.10)
        bsz = min(self._fit_size(it, w - 0.60, bot - (y0 + 0.55 + 0.20 + (0.60 if sb else 0)) - 0.17,
                                 (16, 15, 14, 13), f"箱「{hd}」") for hd, sb, it in nodes)
        for i, (head, sub, items) in enumerate(nodes):
            x = self.MX + i * (w + gap)
            self.rect(s, x, y0, w, 0.55, BLUE)
            self.text(s, x, y0, w, 0.55, [[(head, 17, True, "#FFFFFF")]],
                      align="center", anchor="middle")
            yy = y0 + 0.55
            self.rect(s, x, yy, w, bot - yy, "#FFFFFF", line=GREY, lw=0.75)  # 白本体+細枠(boxes と同じ寸法)
            iy = yy + 0.20
            if sub:   # サブ見出し: 青太字+細い下線(チップにしない)
                self.text(s, x + 0.30, iy, w - 0.60, 0.36, [[(sub, 16, True, BLUE_TX)]])
                self.rect(s, x + 0.30, iy + 0.42, w - 0.60, 0.012, GRID)
                iy += 0.60
            self.text(s, x + 0.30, iy, w - 0.60, bot - iy - 0.17,
                      self.nest_paras(self._soft_items(items, w - 0.60, bsz), size=bsz),
                      space_after=8, hanging=True)
            if i < n - 1:
                ax = x + w + 0.10
                aw = gap - 0.20
                ay = (y0 + bot) / 2
                tw = 0.30
                self.tri_right(s, ax + (aw - tw) / 2, ay - 0.42, tw, 0.85, CYAN)   # 細く高い▶
                if arrows and i < len(arrows) and arrows[i]:
                    # ラベルはギャップ内に収める(2行まで折返し可。箱への重なり防止)
                    self.text(s, ax - 0.20, ay - 1.20, aw + 0.40, 0.72,
                              [[(arrows[i], 14, True, BLUE_TX)]],
                              align="center", anchor="bottom")
        if band:   # 他の型の結論帯と同じ見た目(左寄せ)
            self.rect(s, self.MX, bot + 0.22, self.CW, band_h, LTBLUE)
            self.text(s, self.MX + 0.28, bot + 0.22, self.CW - 0.56, band_h,
                      [[(band, 14, True, BLUE_TX)]], anchor="middle")
        self._source(s, source, note)
        return s

    # ---------------------------------------------------------------- ⑨ ロードマップ
    def timeline(self, title, phases, lead=None, source=None, note=None,
                 takeaway=None, now=None):
        """ロードマップ(2〜4 フェーズ)。phases: [(期間, 見出し, [説明...]), ...]
        now: 現在地の印。(フェーズ index, 0〜1 の位置, "現在地(10 月)") か index だけ。
        進捗の報告では必ず付け、ゴールに対してどこにいるかを示す。"""
        s = self._slide(LAYOUT_BODY, title=title)
        self._kind("diagram")
        _tb = self._reserve_takeaway(takeaway, source, note)
        y0 = self._lead(s, lead)
        n = len(phases)
        colors = RAMP[-n:] if n <= len(RAMP) else [RAMP[min(i, len(RAMP) - 1)] for i in range(n)]
        overlap = 0.24
        w = (self.CW + overlap * (n - 1)) / n
        ch_y = y0 + 0.18
        _bw = w - overlap - 0.10
        _dh = self._body_bot(source, note) - (ch_y + 1.10) - 0.70
        dsz = min(self._fit_size(d_, _bw - 0.10, _dh, (16, 15, 14), f"フェーズ「{h_}」")
                  for _, h_, d_ in phases)   # 全フェーズ共通のサイズ
        for i, (period, head, descs) in enumerate(phases):
            x = self.MX + i * (w - overlap)
            self.rect(s, x, ch_y, w, 0.88, colors[i],
                      shape=(MSO_SHAPE.CHEVRON if i > 0 else MSO_SHAPE.PENTAGON), adj=0.28)
            self.text(s, x + (0.30 if i > 0 else 0.12), ch_y, w - 0.42, 0.88,
                      [[(period, 15, True, "#FFFFFF")]], align="center", anchor="middle")
            bx = self.MX + i * (w - overlap) + (0.10 if i > 0 else 0)
            bw = w - overlap - 0.10
            by = ch_y + 1.10
            self.text(s, bx + 0.06, by, bw, 0.45, [[(head, 16, True, BLUE_TX)]])
            self.rect(s, bx + 0.06, by + 0.46, min(bw - 0.2, _jlen(head) * 16 / 72 * 1.02 + 0.08),
                      0.035, CYAN)
            dh = self._body_bot(source, note) - by - 0.70
            self.text(s, bx + 0.06, by + 0.66, bw - 0.10, dh,
                      self.nest_paras(self._soft_items(descs, bw - 0.10, dsz), size=dsz),
                      space_after=8, hanging=True)
        if now is not None:
            if isinstance(now, (tuple, list)):
                idx = int(now[0])
                frac = float(now[1]) if len(now) > 2 else 0.5
                lab = now[-1] if isinstance(now[-1], str) else "現在地"
            else:
                idx, frac, lab = int(now), 0.5, "現在地"
            idx = max(0, min(idx, n - 1))
            mxp = self.MX + idx * (w - overlap) + 0.30 + (w - overlap - 0.30) * frac
            self.tri_down(s, mxp, ch_y - 0.20, 0.26, 0.18, NAVY)
            self.text(s, mxp - 1.5, ch_y - 0.50, 3.0, 0.28, [[(lab, 12, True, NAVY)]], align="center")
        self._finish_takeaway(s, _tb, takeaway)
        self._source(s, source, note)
        return s

    # ---------------------------------------------------------------- ⑩ メッセージ / 結び
    def message(self, text, sub=None, contact=None):
        """結び。text は結論か依頼の完全な文(何を・どうするか。24 字以内 28pt / 40 字以内 24pt。
        対句やキャッチコピーにしない)。sub は次の一手の 1 文(動詞で終え、担当と期限を入れる)。
        contact は部署名と連絡手段(メール・Teams など)。"""
        s = self._slide(LAYOUT_LAST)
        L = _jlen(text)
        size = 28 if L <= 24 else 24
        if L > 40:
            self._warn("結びの文が 40 字を超えている。2 行以内に削る")
        if not sub:
            self._warn("結びに次の一手(sub)が無い。動詞で終わる 1 文(担当・期限入り)を足す")
        paras = [[(text, size, True, BLUE_TX)]]
        if sub:
            paras.append([(sub, 16, False, INK)])
        self.text(s, 0.10 * self.W, 0.217 * self.H, 0.80 * self.W, 1.5,
                  paras, align="center", space_after=12)
        if contact:
            self.text(s, 0.10 * self.W, 0.787 * self.H, 0.80 * self.W, 1.0,
                      [[(c, 12, False, INK2)] for c in contact],
                      align="center", space_after=4)
        return s

    # ---------------------------------------------------------------- 保存
# ---------------------------------------------------------------- ⑮ エグゼクティブサマリ
    SUMMARY_LABELS = {"purpose": "目的", "background": "背景", "actions": "実施内容",
                      "results": "結果", "conclusion": "結論"}

    def exec_summary(self, title, conclusion, purpose=None, background=None, actions=None,
                     results=None, lead=None, labels=None, source=None, note=None,
                     ask=None, bullets=None, reasons=None, reason_head=None):
        """⑮ 要旨: 目的 → 背景 → 実施内容 → 結果 → 結論(濃紺帯) の固定構成(ユーザー指定)。

        purpose / background / actions / results:
            文字列(1〜2 文)、または ["文", ...](● 2〜3 本、各 40〜70 字)。None の項目は行ごと省く
        conclusion: 結論の一文(最下段の淡青の帯に中央揃えの太字 18pt。2 行・約 90 字まで)。キーメッセージを兼ねる
        lead      : 渡さない(結論帯が兼ねる)
        labels    : 見出しを変えたいときだけ {"purpose": "狙い", ...} で上書き
        本文は 16pt。収まらないときだけ 15→14pt に縮める(それでも溢れたら警告。文を削る)。
        ask= / bullets= / reasons= は廃止(お願い帯・理由カードは使わない)。渡すとエラー。
        """
        if ask or bullets or reasons or reason_head:
            raise TypeError("exec_summary の ask/bullets/reasons は廃止。目的→背景→実施内容→結果→結論 で書く")
        s = self._slide(LAYOUT_BODY, title=title)
        self._kind("text")
        y0 = self._lead(s, lead) if lead else self.BODY_TOP + 0.15
        lab = dict(self.SUMMARY_LABELS, **(labels or {}))
        bot = self._body_bot(source, note) - 0.12
        cpl = max(int((self.CW - 0.6) / (18 / 72)) - 2, 20)      # 18pt 全角換算の 1 行文字数
        raw_lines = max(1, -(-int(_jlen(conclusion) + 0.999) // cpl))
        if raw_lines > 2:
            self._warn("要旨の結論が 3 行以上になる見込み。約 90 字以内に削る")
        n_lines = min(2, raw_lines)
        band_h = 0.42 + 0.36 * n_lines          # 結論帯(他の型の結論帯と同じ淡青・中央揃え)の高さ

        rows = [(lab["purpose"], purpose), (lab["background"], background),
                (lab["actions"], actions), (lab["results"], results)]
        rows = [(l, v) for l, v in rows if v]
        # 縦のラダー形(ユーザー手直し 2026-10-22): 「■ 目的」の見出し行の下に本文(全角 1 字下げ)。
        # 平文は ● なし、複数なら ● の箇条。項目の間は細い罫線
        tx = self.MX + 0.30
        tw = self.CW - 0.30
        head_h = 0.40
        avail = (bot - band_h - 0.22) - y0
        heights, size = [], 16
        for size in (16, 15, 14):
            lh = size / 72 * 1.3
            heights = []
            for _, v in rows:
                items = list(v) if isinstance(v, (list, tuple)) else [v]
                cpl_r = max(int((tw - (0.35 if isinstance(v, (list, tuple)) else 0.0)) / (size / 72)), 10)
                nl = sum(max(1, -(-int(_jlen(it) + 0.999) // cpl_r)) for it in items)
                heights.append(head_h + nl * lh + (len(items) - 1) * 0.06 + 0.26)
            if sum(heights) <= avail:
                break
        if sum(heights) > avail + 0.05:
            self._warn("要旨の本文が入りきらない見込み。各項目を 2 本・各 60 字以内に削る")
        extra = min(0.18, max(0.0, (avail - sum(heights)) / max(len(rows), 1)))
        yy = y0
        for i, ((l, v), h) in enumerate(zip(rows, heights)):
            h += extra
            self.text(s, self.MX, yy, self.CW, head_h, [[("■ " + str(l), size, True, INK)]],
                      anchor="middle")
            if isinstance(v, (list, tuple)) and len(v) > 1:
                paras = [self.bullet(str(it), 1, size) for it in v]
                self.text(s, tx, yy + head_h, tw, h - head_h - 0.10, paras,
                          space_after=4, line_spacing=1.25, hanging=True)
            else:
                v1 = v[0] if isinstance(v, (list, tuple)) else v
                self.text(s, tx, yy + head_h, tw, h - head_h - 0.10, [[(str(v1), size, False, INK)]],
                          line_spacing=1.3)
            yy += h
            if i < len(rows) - 1:
                self.rect(s, self.MX, yy - 0.06, self.CW, 0.012, GRID)
        # 結論: 他の型の結論帯(takeaway)と同じ淡青の帯。ラベルなし・中央揃え・青太字 18pt
        # (濃紺の帯は AI 感。ユーザー指定 2026-10-21/22)
        by = bot - band_h
        self.rect(s, self.MX, by, self.CW, band_h, LTBLUE)
        self.text(s, self.MX + 0.40, by, self.CW - 0.80, band_h, [[(conclusion, 18, True, BLUE_TX)]],
                  line_spacing=1.25, align="center", anchor="middle")
        self._source(s, source, note)
        return s

    # ---------------------------------------------------------------- ⑯ 結論→根拠ツリー
    def reason_tree(self, title, conclusion, reasons, lead=None, source=None,
                    note=None, evidence_tag="根拠", takeaway=None):
        """結論帯から2〜3カードへ分岐する論証ツリー(強み・選定理由の構造化)。

        reasons: [(見出し, 本文(文字列 or ["箇条", ...]), 根拠の一文|None), ...]
        結論帯は他の型の結論帯と同じ淡青+青太字(18pt。2 行なら 16pt)。lead と takeaway は渡さない。
        本文は全カード共通で 16→15→14pt。箇条は 1 行 14 字前後が目安(長ければ \n で切る)。
        根拠はカード下部に「根拠」+青太字(数字を入れると強い)。
        """
        if takeaway:
            self._warn("reason_tree に takeaway は付けない(結論帯が兼ねる)")
            takeaway = None
        s = self._slide(LAYOUT_BODY, title=title)
        self._kind("diagram")
        _tb = self._reserve_takeaway(takeaway, source, note)
        y0 = self._lead(s, lead) if lead else self.BODY_TOP + 0.15
        cw_c = min(self.CW * 0.62, 7.8)
        cx0 = self.MX + (self.CW - cw_c) / 2
        c_lines = _nlines(conclusion, _cpl(cw_c - 0.4, 18))
        c_size, ch = (18, 0.62) if c_lines <= 1 else (16, 0.95)
        if _nlines(conclusion, _cpl(cw_c - 0.4, 16)) > 2:
            self._warn("結論ツリーの結論が 3 行になる見込み。短くする")
        self.rect(s, cx0, y0, cw_c, ch, LTBLUE)
        self.text(s, cx0 + 0.2, y0, cw_c - 0.4, ch,
                  [[(conclusion, c_size, True, BLUE_TX)]], align="center", anchor="middle",
                  line_spacing=1.2)
        n = len(reasons)
        gap = 0.34
        w = (self.CW - gap * (n - 1)) / n
        cy = y0 + ch + 0.48
        mid = self.MX + self.CW / 2
        xs = [self.MX + i * (w + gap) + w / 2 for i in range(n)]
        self.rect(s, mid - 0.006, y0 + ch, 0.012, 0.24, GREY)
        self.rect(s, xs[0], y0 + ch + 0.24, xs[-1] - xs[0], 0.012, GREY)
        for xc in xs:
            self.rect(s, xc - 0.006, y0 + ch + 0.24, 0.012, cy - (y0 + ch + 0.24), GREY)
        bot = self._body_bot(source, note) - 0.12
        # 本文の文字サイズは全カード共通(各カードで入る最大のうち最小)
        _bsz = 16
        for r in reasons:
            head, body, ev = (list(r) + [None] * 3)[:3]
            bh_ = bot - cy - (1.85 if ev else 1.05)
            if isinstance(body, (list, tuple)):
                _bsz = min(_bsz, self._fit_size(list(body), w - 0.52, bh_, (16, 15, 14), f"カード「{head}」"))
            else:
                _bsz = min(_bsz, self._fit_size([str(body)], w - 0.30, bh_ / 1.15, (16, 15, 14), f"カード「{head}」"))
        for i, r in enumerate(reasons):
            head, body, ev = (list(r) + [None] * 3)[:3]
            x = self.MX + i * (w + gap)
            self.rect(s, x, cy, w, bot - cy, "#FFFFFF", line=GREY, lw=0.75)
            self.rect(s, x, cy, w, 0.70, BLUE)   # 見出し帯は他の型と同じ BLUE(淡青は結論帯だけ 2026-10-25)
            self.rect(s, x + 0.26, cy + 0.18, 0.34, 0.34, "#FFFFFF", shape=MSO_SHAPE.OVAL)
            self.text(s, x + 0.26, cy + 0.18, 0.34, 0.34,
                      [[(str(i + 1), 12, True, BLUE)]], align="center", anchor="middle")
            self.text(s, x + 0.72, cy + 0.18, w - 0.95, 0.34, [[(head, 16, True, "#FFFFFF")]],
                      anchor="middle")
            bh_ = bot - cy - (1.85 if ev else 1.05)
            bsz = _bsz
            if isinstance(body, (list, tuple)):
                cpl = _cpl(w - 0.52, bsz, pad=bsz / 72 * 1.1)
                fixed = []
                for it in body:
                    t2 = self._soft_break(it, cpl)          # 語の切れ目で改行(泣き別れ防止)
                    if "\n" not in t2 and _jlen(t2) > cpl and _jlen(t2) % cpl <= 1.5:
                        self._warn(f"カード「{head}」の箇条が語の途中で折り返す見込み: {str(it)[:14]}…(短くする)")
                    if _nlines(t2, cpl) > 2:
                        self._warn(f"カード「{head}」の箇条が 3 行になる見込み: {str(it)[:14]}…(短くする)")
                    fixed.append(t2)
                self.text(s, x + 0.26, cy + 0.90, w - 0.52, bh_,
                          self.nest_paras(fixed, size=bsz), space_after=8, hanging=True)
            else:
                self.text(s, x + 0.26, cy + 0.90, w - 0.52, bh_,
                          [[(str(body), bsz, False, INK)]], line_spacing=1.45)
            if ev:
                ey = bot - 0.92
                self.rect(s, x + 0.26, ey, w - 0.52, 0.012, GRID)
                self.text(s, x + 0.26, ey + 0.14, w - 0.52, 0.28,
                          [[(evidence_tag, 12, True, INK2)]])
                self.text(s, x + 0.26, ey + 0.44, w - 0.52, 0.40,
                          [[(ev, 15, True, BLUE_TX)]])
        self._finish_takeaway(s, _tb, takeaway)
        self._source(s, source, note)
        return s

    # ---------------------------------------------------------------- ⑰ スイムレーン・ガント
    def gantt(self, title, columns, lanes, phases=None, milestones=None,
              lead=None, source=None, note=None, label_w=1.75, takeaway=None, col_weights=None,
              axis_label=None, today=None):
        """フェーズ帯+レーン別バー+節目の行程表(timeline の詳細版)。

        columns   : 時間軸の列ラベル。等しい単位(12 か月以下なら月)で切る。束ね列+col_weights は最後の手段
        axis_label: 年度か起点(例 "2027年度")。列見出しの行の左に出す。無いと警告
        phases    : [(帯ラベル, 開始列, 終了列), ...](0始まり・両端含む)。省略可
        lanes     : [(レーン名, [(バーlabel, 開始列, 終了列, style), ...]), ...] 5 本まで
                    style: "main"=BLUE(既定・確定した作業) | "option"=白地+青破線(判断次第・未確定) |
                           "key"=NAVY(主役の 1 本。2 本以上は警告)
                    旧名 solid→main / plan→option / dark→key。mid・light は main に寄せる(警告)
        milestones: [(列位置(float・列単位), ラベル), ...] — 縦の破線+下の ◆。ラベルは「判断名+時期」
        today     : 現在の列位置(float)。進捗の報告では青の縦線で現在地を示す(注に「青線=現在」が付く)
        注には、使った style の意味を自動で足す(例「破線=判断次第で実施」)。
        """
        s = self._slide(LAYOUT_BODY, title=title)
        self._kind("diagram")
        _tb = self._reserve_takeaway(takeaway, source, note)
        y0 = self._lead(s, lead)
        gx = self.MX + label_w
        gw = self.CW - label_w
        n = len(columns)
        wts = list(col_weights) if col_weights else [1.0] * n
        if col_weights and len(set(wts)) > 1:
            self._warn("ガントの列幅が不揃い(col_weights)。列は等しい単位で切るのが基本")
        if not axis_label:
            self._warn("ガントに年度か起点(axis_label)が無い。「2027年度」のように書く")
        if len(lanes) > 5:
            self._warn(f"ガントのレーンが {len(lanes)} 本。5 本以内にする")
        edges = [gx]
        for wt in wts:
            edges.append(edges[-1] + gw * wt / sum(wts))

        def cx_(pos):                       # 列位置(float)→ x
            i = min(int(pos), n - 1)
            return edges[i] + (edges[i + 1] - edges[i]) * (pos - i)
        y = y0
        if phases:
            for lab, c0, c1 in phases:
                x0 = edges[c0] + 0.03
                x1 = edges[c1 + 1] - 0.03
                self.rect(s, x0, y, x1 - x0, 0.38, BG)
                self.text(s, x0, y, x1 - x0, 0.38, [[(lab, 12, True, BLUE_TX)]],
                          align="center", anchor="middle")
            y += 0.48
        if axis_label:
            if _jlen(axis_label) * 12 / 72 * CHAR_EM > label_w - 0.2:
                self._warn(f"axis_label「{axis_label}」がレーン名の列幅に入らない。短くする")
            self.text(s, self.MX, y, label_w - 0.15, 0.28, [[(str(axis_label), 12, True, INK2)]])
        for i, c in enumerate(columns):
            self.text(s, edges[i], y, edges[i + 1] - edges[i], 0.28, [[(str(c), 12, False, INK2)]],
                      align="center")
        y += 0.38
        bot = self._body_bot(source, note) - (0.62 if milestones else 0.10)
        for e in edges:
            self.rect(s, e - 0.006, y, 0.012, bot - y, GRID)
        if milestones:                       # 節目の縦の破線(バーの背面)
            for pos, _ in milestones:
                mx_ = cx_(pos) if pos < n else edges[-1]
                self._dash_line(s, mx_, y, mx_, bot, INK2, 1.0)
        if today is not None:                # 現在地(青の縦線。バーの背面)
            tx_ = cx_(float(today)) if float(today) < n else edges[-1]
            self.rect(s, tx_ - 0.012, y, 0.024, bot - y, BLUE)
        lane_h = (bot - y) / len(lanes)
        if lane_h < 0.56:
            self._warn(f"ガントのレーンが多すぎる(レーン高 {lane_h:.2f}in)。レーンを減らす")
        STYLES = {"main":   (BLUE, "#FFFFFF", None, False),
                  "key":    (NAVY, "#FFFFFF", None, False),
                  "option": ("#FFFFFF", BLUE_TX, BLUE, True)}
        ALIAS = {"solid": "main", "dark": "key", "plan": "option", "mid": "main", "light": "main"}
        used, n_key = set(), 0
        for li, (name, bars) in enumerate(lanes):
            ly = y + li * lane_h
            if li > 0:
                self.rect(s, self.MX, ly - 0.006, self.CW, 0.012, GRID)
            self.text(s, self.MX, ly, label_w - 0.15, lane_h,
                      [[(name, 13, True, INK)]], anchor="middle")
            for (lab, c0, c1, style) in bars:
                st = ALIAS.get(style, style)
                if style in ("mid", "light"):
                    self._warn(f"バー「{lab}」の style {style} は廃止。main に寄せた(意味の違いは option か key で)")
                if st not in STYLES:
                    raise ValueError(f"gantt の style は main / option / key: {style}")
                used.add(st)
                if st == "key":
                    n_key += 1
                bx0 = edges[c0] + 0.06
                bx1 = edges[c1 + 1] - 0.06
                bh = min(0.40, lane_h - 0.16)
                by = ly + lane_h / 2 - bh / 2
                fill, tcol, line, dash = STYLES[st]
                self.rect(s, bx0, by, bx1 - bx0, bh, fill, line=line, lw=1.2,
                          dash=dash, shape=MSO_SHAPE.ROUNDED_RECTANGLE, adj=0.12)
                need_w = _jlen(lab) * 12 / 72 * 1.05 + 0.16
                if need_w > bx1 - bx0:
                    self._warn(f"バーのラベル「{lab}」がバーの幅({bx1 - bx0:.1f}in)に入らない。語を縮めるかバーを延ばす")
                self.text(s, bx0 + 0.05, by, bx1 - bx0 - 0.10, bh,
                          [[(lab, 12, True, tcol)]], align="center", anchor="middle")
        if n_key >= 2:
            self._warn(f"ガントの主役(key)のバーが {n_key} 本。1 本に絞る")
        self.rect(s, self.MX, bot - 0.006, self.CW, 0.012, GRID)
        if milestones:
            for pos, lab in milestones:
                mx_ = cx_(pos) if pos < n else edges[-1]
                self.rect(s, mx_ - 0.10, bot + 0.06, 0.20, 0.20, BLUE, shape=MSO_SHAPE.DIAMOND)
                if mx_ - 1.6 < self.MX:          # 端の点はラベルを点に寄せる
                    self.text(s, self.MX, bot + 0.30, 3.2, 0.28,
                              [[(lab, 12, True, INK)]])
                elif mx_ + 1.6 > self.W - self.MX:
                    self.text(s, self.W - self.MX - 3.2, bot + 0.30, 3.2, 0.28,
                              [[(lab, 12, True, INK)]], align="right")
                else:
                    self.text(s, mx_ - 1.6, bot + 0.30, 3.2, 0.28,
                              [[(lab, 12, True, INK)]], align="center")
        legend = []
        if "option" in used:
            legend.append("破線=判断次第で実施")
        if "key" in used:
            legend.append("濃紺=主役の作業")
        if milestones:
            legend.append("◆=判断の節目")
        if today is not None:
            legend.append("青線=現在")
        if legend:
            note = (str(note).rstrip("。") + "。" if note else "") + "、".join(legend)
        self._finish_takeaway(s, _tb, takeaway)
        self._source(s, source, note)
        return s

    # ---------------------------------------------------------------- ⑱ 壁→打ち手プロセス
    def process(self, title, steps, lead=None, source=None, note=None,
                takeaway=None, loop=False):
        """⑱ 壁→打ち手プロセス(表の形)。列は No | 工程 | 壁(詰まるところ) | 打ち手。

        steps: [(工程の一文 30 字以内, (壁の一文, 打ち手の一文) | None), ...] 3〜6 個
        壁が無い工程は None(セルは「—」)。全工程に壁を付けない。
        loop=True: 表の下に「工程 N の結果を工程 1 の変更に戻す」の 1 行(lead で「輪」「回す」と言うなら付ける)。
        書き方: 壁は「なぜその工程が詰まるか」の原因(〜できない / 〜が分からない / 〜が古くなる)。
        判定結果(未達・NG)や起きうる結果は壁にしない。打ち手は動詞の終止形で終え、頻度・期限・件数を 1 つ入れる。
        縦の番号丸をつないだステッパーや、ラベル語の対(壁 / 打ち手)は UI 部品の見た目で AI くさい(ユーザー評価 2026-10-24)。
        """
        from lxml import etree  # noqa: F401
        s = self._slide(LAYOUT_BODY, title=title)
        self._kind("table")
        _tb = self._reserve_takeaway(takeaway, source, note)
        y = self._lead(s, lead) + 0.05
        if not loop and lead and re.search(r"(輪|循環|サイクル|回す)", str(lead)):
            self._warn("lead が「輪・回す」と言っているが loop=False。loop=True にするか lead を直す")
        rows = []
        for st in steps:
            if (isinstance(st, (tuple, list)) and len(st) == 2
                    and (st[1] is None or isinstance(st[1], (tuple, list)))):
                rows.append((str(st[0]), st[1]))
            else:
                rows.append((str(st), None))
        if len(rows) > 6:
            self._warn(f"工程が {len(rows)} 個。6 工程以内にする")
        if all(g for _, g in rows) and len(rows) >= 3:
            self._warn("全工程に壁が付いている。詰まる工程だけに絞る")
        for _, g in rows:
            if g and not _SENT_END_VERB.search(str(g[1]).rstrip("。")):
                self._warn(f"打ち手「{g[1]}」が動詞で終わっていない(〜する / 〜足す の形にする)")
        heads = ["", "工程", "壁(詰まるところ)", "打ち手"]
        widths = [0.45, self.CW * 0.40, self.CW * 0.25, 0.0]
        widths[3] = self.CW - sum(widths[:3])
        size = 15
        loop_h = 0.40 if loop else 0.0
        bot = self._body_bot(source, note) - 0.10 - loop_h
        cells_all = []
        heights = [0.46]
        for i, (txt, gap) in enumerate(rows):
            cells = [str(i + 1), txt, (gap[0] if gap else "—"), (gap[1] if gap else "—")]
            # 1 行に入らない文は語の切れ目で改行(語の途中・1 字だけの行を避ける)
            cells = [cells[0]] + [self._soft_break(c, _cpl(wj, size, pad=0.24)) if c != "—" else c
                                  for c, wj in zip(cells[1:], widths[1:])]
            cells_all.append(cells)
            mx = max(_nlines(c, _cpl(wj, size, pad=0.24)) for c, wj in zip(cells[1:], widths[1:]))
            if mx >= 3:
                self._warn(f"工程 {i + 1} のセルが {mx} 行になる見込み。文を短くする")
            heights.append(max(0.56, size / 72 * 1.3 * mx + 0.24))
        total = sum(heights)
        avail = bot - y
        if total > avail + 0.05:
            self._warn(f"工程の表が本文領域を {total - avail:.2f}in 超える見込み。工程か文を減らす")
        elif total < avail:                      # 余りは行に配る(上限 0.95in)
            extra = min(0.95, (avail - 0.46) / len(rows))
            heights = [0.46] + [min(max(h, extra), 0.95) for h in heights[1:]]
        nrows, ncols = len(rows) + 1, 4
        gf = s.shapes.add_table(nrows, ncols, Inches(self.MX), Inches(y), Inches(self.CW), Inches(sum(heights)))
        tbl = gf.table
        tbl.first_row = False
        tbl.horz_banding = False
        for c, wj in enumerate(widths):
            tbl.columns[c].width = Inches(wj)
        for r, h in enumerate(heights):
            tbl.rows[r].height = Inches(h)

        def put(cell, txt, sz, bold, color, bottom, align="left"):
            cell.fill.background()
            cell.margin_left = Inches(0.10)
            cell.margin_right = Inches(0.10)
            cell.margin_top = cell.margin_bottom = Inches(0.04)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf = cell.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = _ALIGN[align]
            for k, piece in enumerate(str(txt).split(chr(10))):
                if k > 0:
                    br = p._p.add_br()
                    br.get_or_add_rPr().set("sz", str(int(sz * 100)))
                run = p.add_run()
                run.text = piece
                _set_font(run, sz, bold, color)
            self._cell_borders(cell, bottom=bottom)

        for c, hd in enumerate(heads):
            put(tbl.cell(0, c), hd, 14, True, INK2, (BLUE, 1.5))
        for r, cells in enumerate(cells_all, 1):
            bb = (GREY, 0.75) if r == nrows - 1 else (GRID, 0.75)
            put(tbl.cell(r, 0), cells[0], 14, True, STEEL, bb, align="center")
            put(tbl.cell(r, 1), cells[1], size, False, INK, bb)
            if cells[2] == "—":
                put(tbl.cell(r, 2), "—", size, False, GREY, bb, align="center")
                put(tbl.cell(r, 3), "—", size, False, GREY, bb, align="center")
            else:
                put(tbl.cell(r, 2), cells[2], size, False, INK2, bb)
                put(tbl.cell(r, 3), cells[3], size, True, BLUE_TX, bb)
        if loop:
            self.text(s, self.MX, y + sum(heights) + 0.10, self.CW, 0.28,
                      [[(f"工程 {len(rows)} で足した結果は、次の工程 1 の変更に戻して回す", 12, False, INK2)]])
        self._finish_takeaway(s, _tb, takeaway)
        self._source(s, source, note)
        return s

    # ---------------------------------------------------------------- 図解の型(追加 2026-10-24)
    def cycle_flow(self, title, steps, lead=None, center=None, source=None, note=None, takeaway=None,
                   detail_title="各工程の成果物と判断"):
        """サイクル図(改善の輪・PDCA・運用ループ)。左に円環(番号+見出し)、右に番号付きの欄。
        steps : [(見出し 6 字以内, 右欄の 1 行 24 字以内), ...] 3〜6 個(順に時計回り)
                右欄は見出しの言い換えでなく、別の情報(成果物・判断・詰まりどころと打ち手)を書く
        detail_title: 右欄の見出し(例「各工程の詰まりどころと打ち手」)
        center: 円の中に置く短い語(例「版ごとに 1 周」)。伝わらないなら省く
        輪・循環・繰り返しを lead で言う内容はこれ(表+「戻す」の 1 行より流れが見える)。"""
        import math
        from lxml import etree
        if not 3 <= len(steps) <= 6:
            raise ValueError("cycle_flow は 3〜6 工程")
        s = self._slide(LAYOUT_BODY, title=title)
        self._kind("diagram")
        _tb = self._reserve_takeaway(takeaway, source, note)
        y0 = self._lead(s, lead)
        bot = self._body_bot(source, note) - 0.10
        n = len(steps)
        cx = self.MX + 3.4
        cy = (y0 + bot) / 2
        R = max(1.0, min(1.35, (bot - y0) / 2 - 1.1))
        nr = 0.27
        pts = []
        for i in range(n):
            deg = (270 + 360 * i / n) % 360
            th = math.radians(deg)
            pts.append((cx + R * math.cos(th), cy + R * math.sin(th), deg))
        gdeg = math.degrees(math.asin(min(0.99, (nr + 0.12) / R)))
        for i in range(n):                    # ノード間の弧(時計回り)+矢じり
            a0 = (pts[i][2] + gdeg) % 360
            a1 = (pts[(i + 1) % n][2] - gdeg) % 360
            sp = s.shapes.add_shape(MSO_SHAPE.ARC, Inches(cx - R), Inches(cy - R), Inches(2 * R), Inches(2 * R))
            sp.shadow.inherit = False
            try:
                sp.adjustments[0] = a0 * 0.6
                sp.adjustments[1] = a1 * 0.6
            except Exception:
                pass
            sp.fill.background()
            sp.line.color.rgb = _rgb(STEEL)
            sp.line.width = Pt(2.0)
            ln = sp.line._get_or_add_ln()
            for old in ln.findall(qn("a:tailEnd")):
                ln.remove(old)
            te = etree.SubElement(ln, qn("a:tailEnd"))
            te.set("type", "triangle"); te.set("w", "med"); te.set("len", "med")
        lw_lab = 1.6
        for i, ((px, py, deg), (head, _)) in enumerate(zip(pts, steps)):
            if _jlen(head) > 7:
                self._warn(f"サイクル図の見出し「{head}」が長い(6 字以内)。説明は右の欄に書く")
            self.rect(s, px - nr, py - nr, 2 * nr, 2 * nr, BLUE, shape=MSO_SHAPE.OVAL)
            self.text(s, px - nr, py - nr, 2 * nr, 2 * nr, [[(str(i + 1), 14, True, "#FFFFFF")]],
                      align="center", anchor="middle")
            c, sn = math.cos(math.radians(deg)), math.sin(math.radians(deg))
            if c > 0.35:
                self.text(s, px + nr + 0.10, py - 0.18, lw_lab, 0.36, [[(head, 14, True, INK)]], anchor="middle")
            elif c < -0.35:
                self.text(s, px - nr - 0.10 - lw_lab, py - 0.18, lw_lab, 0.36, [[(head, 14, True, INK)]],
                          align="right", anchor="middle")
            elif sn < 0:
                self.text(s, px - lw_lab / 2, py - nr - 0.44, lw_lab, 0.36, [[(head, 14, True, INK)]],
                          align="center", anchor="bottom")
            else:
                self.text(s, px - lw_lab / 2, py + nr + 0.08, lw_lab, 0.36, [[(head, 14, True, INK)]],
                          align="center")
        if center:
            self.text(s, cx - R + 0.45, cy - 0.45, 2 * R - 0.9, 0.9, [[(center, 14, True, BLUE_TX)]],
                      align="center", anchor="middle", line_spacing=1.15)
        rx = self.MX + 6.8
        rw = self.CW - 6.8
        bx, ly, bw, bh = self.panel(s, rx, y0, rw, bot - y0, detail_title)
        self.numbered_list(s, bx, ly, bw, [(h, d) for h, d in steps], max_y=ly + bh)
        self._finish_takeaway(s, _tb, takeaway)
        self._source(s, source, note)
        return s

    def layers(self, title, layers, lead=None, side=None, source=None, note=None, takeaway=None):
        """レイヤー図(システム構成・階層構造)。上から順に層を帯で重ね、各層の要素を箱で並べる。
        layers: [(層の名前, [要素, ...]), ...] 2〜5(上→下)。要素は "名前" か ("名前", "hl"|"plan")
                hl=主役(BLUE 塗り・白文字) / plan=計画(破線)。要素は各層 2〜5 個
        side  : 全層をまたぐ縦の帯の名前(例「監視・トレース」)。省略可
        概念図(データの流れ)と違い、何がどの層にあるかの「置き場」を見せる型。"""
        if not 2 <= len(layers) <= 5:
            raise ValueError("layers は 2〜5 層")
        s = self._slide(LAYOUT_BODY, title=title)
        self._kind("diagram")
        _tb = self._reserve_takeaway(takeaway, source, note)
        y0 = self._lead(s, lead)
        bot = self._body_bot(source, note) - 0.10
        n = len(layers)
        gap, lab_w = 0.16, 1.9
        side_w = 1.9 if side else 0.0
        row_h = min(1.25, (bot - y0 - gap * (n - 1)) / n)
        x_body = self.MX + lab_w + 0.12
        w_body = self.CW - lab_w - 0.12 - (side_w + 0.16 if side else 0.0)
        for i, (name, items) in enumerate(layers):
            y = y0 + i * (row_h + gap)
            self.rect(s, self.MX, y, lab_w, row_h, BLUE)
            self.text(s, self.MX + 0.10, y, lab_w - 0.20, row_h, [[(name, 14, True, "#FFFFFF")]],
                      align="center", anchor="middle")
            self.rect(s, x_body, y, w_body, row_h, BG)
            m = max(len(items), 1)
            if m > 5:
                self._warn(f"層「{name}」の要素が {m} 個。5 個以内にする")
            ig, pad = 0.18, 0.18
            bw = (w_body - 2 * pad - ig * (m - 1)) / m
            bh = min(row_h - 0.30, 0.90)
            by = y + (row_h - bh) / 2
            for j, it in enumerate(items):
                nm, st = (it if isinstance(it, (tuple, list)) else (it, None))
                bx = x_body + pad + j * (bw + ig)
                if st == "hl":
                    fill, col, line, dash, bold = BLUE, "#FFFFFF", None, False, True
                elif st == "plan":
                    fill, col, line, dash, bold = "#FFFFFF", INK2, GREY, True, False
                else:
                    fill, col, line, dash, bold = "#FFFFFF", INK, GREY, False, False
                self.rect(s, bx, by, bw, bh, fill, line=line, lw=0.75, dash=dash)
                cpl = _cpl(bw, 14, pad=0.2)
                t2 = self._soft_break(nm, cpl)
                if _nlines(t2, cpl) > 2:
                    self._warn(f"要素「{nm}」が 3 行になる見込み。短くするか要素を減らす")
                self.text(s, bx + 0.08, by, bw - 0.16, bh, [[(t2, 14, bold, col)]],
                          align="center", anchor="middle", line_spacing=1.15)
        if side:
            sx = self.MX + self.CW - side_w
            sh = n * row_h + (n - 1) * gap
            self.rect(s, sx, y0, side_w, sh, "#FFFFFF", line=GREY, lw=0.75)
            cpl_s = _cpl(side_w, 14, pad=0.3)
            self.text(s, sx + 0.12, y0 + 0.2, side_w - 0.24, sh - 0.4,
                      [[(self._soft_break(side, cpl_s), 14, True, BLUE_TX)]],
                      align="center", anchor="middle", line_spacing=1.2)
        self._finish_takeaway(s, _tb, takeaway)
        self._source(s, source, note)
        return s

    def swimlane(self, title, lanes, stages, cells, lead=None, source=None, note=None, takeaway=None):
        """スイムレーン(担当 × 工程の流れ)。行=担当(2〜4)、列=工程(3〜6)。
        cells: [(担当 index, 工程 index, 文, "hl"?), ...] を流れの順に。連続する箱は矢印で自動的につなぐ
               (同じ担当で右へ=実線、担当をまたぐ=カギ線、左へ戻る=破線)。判定点は "hl"(BLUE 塗り)。
        誰が・どの段で・何をするかを 1 枚で見せる型。役割の説明だけなら flow。"""
        if not 2 <= len(lanes) <= 4 or not 3 <= len(stages) <= 6:
            raise ValueError("swimlane は担当 2〜4 × 工程 3〜6")
        s = self._slide(LAYOUT_BODY, title=title)
        self._kind("diagram")
        _tb = self._reserve_takeaway(takeaway, source, note)
        y0 = self._lead(s, lead)
        bot = self._body_bot(source, note) - 0.10
        lab_w, head_h = 1.7, 0.42
        gx = self.MX + lab_w
        gw = self.CW - lab_w
        col_w = gw / len(stages)
        top = y0 + head_h
        lane_h = (bot - top) / len(lanes)
        for j, st in enumerate(stages):
            self.text(s, gx + j * col_w, y0, col_w, head_h - 0.06, [[(st, 14, True, INK2)]],
                      align="center", anchor="bottom")
        self.rect(s, gx, top - 0.02, gw, 0.014, BLUE)
        for i, ln_ in enumerate(lanes):
            ly = top + i * lane_h
            self.rect(s, self.MX, ly + 0.05, lab_w - 0.12, lane_h - 0.10, BLUE)
            self.text(s, self.MX + 0.08, ly, lab_w - 0.28, lane_h, [[(ln_, 14, True, "#FFFFFF")]],
                      align="center", anchor="middle")
            if i > 0:
                self.rect(s, gx, ly - 0.006, gw, 0.012, GRID)
        for j in range(1, len(stages)):
            self.rect(s, gx + j * col_w - 0.006, top, 0.012, bot - top, GRID)
        self.rect(s, gx, bot - 0.006, gw, 0.012, GRID)
        bw = col_w - 0.34
        bh = min(0.95, lane_h - 0.34)
        boxes, seen = [], set()
        for c in cells:
            li, sj, txt = int(c[0]), int(c[1]), str(c[2])
            st = c[3] if len(c) > 3 else None
            if (li, sj) in seen:
                self._warn(f"スイムレーンの同じ枡(担当 {li}・工程 {sj})に 2 つ目の箱「{txt[:10]}」。1 枡 1 箱にする")
                continue
            seen.add((li, sj))
            bx = gx + sj * col_w + (col_w - bw) / 2
            by = top + li * lane_h + (lane_h - bh) / 2
            hl = (st == "hl")
            self.rect(s, bx, by, bw, bh, BLUE if hl else "#FFFFFF", line=(None if hl else GREY), lw=0.75)
            cpl = _cpl(bw, 13, pad=0.2)
            t2 = self._soft_break(txt, cpl)
            if _nlines(t2, cpl) > 3:
                self._warn(f"箱「{txt[:10]}」が 4 行になる見込み。短くする")
            self.text(s, bx + 0.08, by, bw - 0.16, bh, [[(t2, 13, hl, "#FFFFFF" if hl else INK)]],
                      align="center", anchor="middle", line_spacing=1.15)
            boxes.append((bx, by, bw, bh, li, sj))
        for a, b in zip(boxes, boxes[1:]):
            ax, ay, aw, ah, al, asj = a
            bx_, by_, bw_, bh_, bl, bsj = b
            acx, acy = ax + aw / 2, ay + ah / 2
            bcx, bcy = bx_ + bw_ / 2, by_ + bh_ / 2
            if al == bl and bsj > asj:
                self.arrow_h(s, ax + aw, bx_, acy)
            elif asj == bsj:
                if bl > al:
                    self.arrow_v(s, acx, ay + ah, by_)
                else:
                    self.arrow_v(s, acx, ay, by_ + bh_)
            elif bsj > asj:
                mx_ = (ax + aw + bx_) / 2
                self._line_seg(s, ax + aw, acy, mx_, acy, STEEL, 1.5)
                self._line_seg(s, mx_, acy, mx_, bcy, STEEL, 1.5)
                self.arrow_h(s, mx_, bx_, bcy)
            else:
                mx_ = (bx_ + bw_ + ax) / 2
                self._dash_line(s, ax, acy, mx_, acy, STEEL, 1.5)
                self._dash_line(s, mx_, acy, mx_, bcy, STEEL, 1.5)
                self.arrow_h(s, mx_, bx_ + bw_, bcy, dash=True)
        self._finish_takeaway(s, _tb, takeaway)
        self._source(s, source, note)
        return s

    def tree(self, title, root, branches, lead=None, source=None, note=None, takeaway=None, highlight=None,
             direction="right"):
        """ロジックツリー。direction="right": 横(目標 ← 要因 ← 具体策 / 課題 → 分解。枝 2〜5・葉は各 4)
        direction="down": 縦(結論 1 つと根拠 2〜3 のカード。= reason_tree。枝は (見出し, [本文...], 根拠の一文|None))
        root     : 根の文(2 行まで。\n で改行)
        branches : [(枝の見出し, [葉の文, ...]), ...]。down のときは 3 つ目に根拠の一文を置ける
        highlight: 強調する枝 index のリスト(淡青)。lead で名指しする枝と合わせる"""
        if direction == "down":
            reasons = [(b[0], list(b[1]), (b[2] if len(b) > 2 else None)) for b in branches]
            return self.reason_tree(title, root, reasons, lead=lead, source=source, note=note)
        if not 2 <= len(branches) <= 5:
            raise ValueError("tree の枝は 2〜5")
        s = self._slide(LAYOUT_BODY, title=title)
        self._kind("diagram")
        _tb = self._reserve_takeaway(takeaway, source, note)
        y0 = self._lead(s, lead)
        bot = self._body_bot(source, note) - 0.10
        n = len(branches)
        rw_, bw_ = 2.5, 3.2
        trunk_x = self.MX + rw_ + 0.40
        bx = trunk_x + 0.40
        lx = bx + bw_ + 0.35
        lw_ = self.MX + self.CW - lx
        cpl = _cpl(lw_ - 0.1, 13, pad=0.3)
        sizes = []
        for name, leaves in branches:
            if len(leaves) > 4:
                self._warn(f"枝「{name}」の葉が {len(leaves)} 個。4 個以内にする")
            rows = max(1, sum(_nlines(self._soft_break(l, cpl), cpl) for l in leaves))
            sizes.append(max(0.72, rows * 0.30 + 0.34))
        gap = 0.18
        avail = bot - y0
        total = sum(sizes) + gap * (n - 1)
        if total > avail + 0.05:
            self._warn(f"ツリーが本文領域を {total - avail:.2f}in 超える見込み。葉を減らす")
            sc = (avail - gap * (n - 1)) / max(sum(sizes), 1e-9)
            sizes = [h * sc for h in sizes]
        else:
            extra = (avail - total) / n
            sizes = [h + extra for h in sizes]
        ys, y = [], y0
        for h in sizes:
            ys.append((y, h))
            y += h + gap
        root_h = min(1.4, avail * 0.45)
        root_y = y0 + (avail - root_h) / 2
        self.rect(s, self.MX, root_y, rw_, root_h, BLUE)
        self.text(s, self.MX + 0.12, root_y, rw_ - 0.24, root_h, [[(root, 15, True, "#FFFFFF")]],
                  align="center", anchor="middle", line_spacing=1.2)
        rcy = root_y + root_h / 2
        cys = [yy + h / 2 for yy, h in ys]
        self._line_seg(s, self.MX + rw_, rcy, trunk_x, rcy, GREY, 1.25)
        self._line_seg(s, trunk_x, min(cys[0], rcy), trunk_x, max(cys[-1], rcy), GREY, 1.25)
        hl = set(highlight or [])
        for i, ((yy, h), cy, (name, leaves)) in enumerate(zip(ys, cys, branches)):
            self._line_seg(s, trunk_x, cy, bx, cy, GREY, 1.25)
            box_h = min(h, 0.9)
            box_y = cy - box_h / 2
            self.rect(s, bx, box_y, bw_, box_h, LTBLUE if i in hl else "#FFFFFF", line=GREY, lw=0.75)
            self.text(s, bx + 0.12, box_y, bw_ - 0.24, box_h, [[(name, 14, True, BLUE_TX if i in hl else INK)]],
                      anchor="middle", line_spacing=1.15)
            if leaves:
                self._line_seg(s, bx + bw_, cy, lx - 0.10, cy, GREY, 1.25)
                paras = self.nest_paras(self._soft_items(list(leaves), lw_ - 0.1, 13), size=13)
                self.text(s, lx, yy, lw_ - 0.1, h, paras, space_after=4, hanging=True, anchor="middle")
        self._finish_takeaway(s, _tb, takeaway)
        self._source(s, source, note)
        return s

    def goal_bars(self, title, items, lead=None, source=None, note=None, takeaway=None):
        """ゴールに対する現在地(指標ごとの進捗バー)。「目標に対してどこにいるか」を問われる報告の型。
        items: [(指標, 現在値, 目標値, 単位, 起点値|None, 補足|None, "lower"?), ...] 2〜6
               バーは起点(無ければ 0)→目標の物差し。塗り=現在値、濃紺の縦線=目標、右端に目標との差(未達は △ 赤)。
               "lower" は低いほど良い指標(応答時間・拒否率)。"""
        if not 2 <= len(items) <= 6:
            raise ValueError("goal_bars は 2〜6 指標")
        s = self._slide(LAYOUT_BODY, title=title)
        self._kind("chart")
        _tb = self._reserve_takeaway(takeaway, source, note)
        y0 = self._lead(s, lead)
        bot = self._body_bot(source, note) - 0.10
        n = len(items)
        lab_w, gap_w, hdr = 3.0, 1.6, 0.34
        tx = self.MX + lab_w
        tw = self.CW - lab_w - gap_w
        self.text(s, self.MX, y0, lab_w, hdr, [[("指標", 12, True, INK2)]], anchor="bottom")
        self.text(s, tx, y0, tw, hdr, [[("起点 → 現在 → 目標", 12, True, INK2)]], anchor="bottom")
        self.text(s, tx + tw, y0, gap_w, hdr, [[("目標との差", 12, True, INK2)]], align="right", anchor="bottom")
        self.rect(s, self.MX, y0 + hdr + 0.02, self.CW, 0.014, BLUE)
        top = y0 + hdr + 0.16
        row_h = min(1.05, (bot - top) / n)

        def fmt(v):
            return f"{v:g}"
        for i, it in enumerate(items):
            name, cur, goal, unit, base, memo, direc = (list(it) + [None] * 7)[:7]
            unit = unit or ""
            cur, goal = float(cur), float(goal)
            lower = (direc == "lower")
            ry = top + i * row_h
            paras = [[(name, 14, True, INK)]]
            if memo:
                paras.append([(memo, 12, False, INK2)])
            self.text(s, self.MX, ry, lab_w - 0.2, row_h, paras, anchor="middle", space_after=2)
            bar_h = 0.34
            by = ry + row_h / 2 - bar_h / 2
            if lower:
                # 低いほど良い: 物差しは他の行と同じ左→右(0 → 大)。目標は「上限」、現在値に上限内 / 超過を添える
                lo = 0.0
                hi = max(cur, goal, float(base) if base is not None else 0.0) * 1.15
                fx = lambda v, lo=lo, hi=hi: tx + tw * (v - lo) / max(hi - lo, 1e-9)
            else:
                lo = float(base) if base is not None else 0.0
                if lo >= min(cur, goal):
                    lo = min(cur, goal) * 0.85
                hi = max(goal, cur) * 1.08
                fx = lambda v, lo=lo, hi=hi: tx + tw * (v - lo) / max(hi - lo, 1e-9)
            self.rect(s, tx, by, tw, bar_h, BG)
            self.rect(s, tx, by, max(fx(cur) - tx, 0.02), bar_h, BLUE)
            gx_ = fx(goal)
            self.rect(s, gx_ - 0.02, by - 0.12, 0.04, bar_h + 0.24, NAVY)
            self.text(s, gx_ - 1.3, by - 0.40, 2.6, 0.26,
                      [[(f"{'上限' if lower else '目標'} {fmt(goal)}{unit}", 12, True, NAVY)]], align="center")
            if lower:
                cur_txt = f"現在 {fmt(cur)}{unit}({'上限内' if cur <= goal else '超過'})"
            else:
                cur_txt = f"現在 {fmt(cur)}{unit}"
            cw_ = _jlen(cur_txt) * 14 / 72 * CHAR_EM + 0.25
            if fx(cur) + cw_ + 0.05 <= tx + tw and not (fx(cur) < gx_ < fx(cur) + cw_ + 0.1):
                self.text(s, fx(cur) + 0.08, by, cw_, bar_h, [[(cur_txt, 14, True, INK)]], anchor="middle")
            else:
                self.text(s, fx(cur) - cw_ - 0.08, by, cw_, bar_h, [[(cur_txt, 14, True, "#FFFFFF")]],
                          align="right", anchor="middle")
            if base is not None:
                bx_ = fx(float(base))
                if lower:    # 起点は物差しの途中にあるので、灰の目盛りで示す
                    self.rect(s, bx_ - 0.01, by - 0.06, 0.02, bar_h + 0.12, GREY)
                    self.text(s, bx_ - 1.0, by + bar_h + 0.04, 2.0, 0.24,
                              [[(f"起点 {fmt(float(base))}{unit}", 12, False, INK2)]], align="center")
                else:
                    self.text(s, tx, by + bar_h + 0.04, 2.4, 0.24, [[(f"起点 {fmt(float(base))}{unit}", 12, False, INK2)]])
            d = (goal - cur) if lower else (cur - goal)
            du = "pt" if unit == "%" else unit
            if d >= 0:
                dtxt, dcol = f"+{fmt(d)}{du} {'余裕' if lower else '達成'}", INK
            else:
                dtxt, dcol = f"△{fmt(abs(d))}{du} {'超過' if lower else '不足'}", RED
            self.text(s, tx + tw + 0.1, ry, gap_w - 0.1, row_h, [[(dtxt, 16, True, dcol)]],
                      align="right", anchor="middle")
        self._finish_takeaway(s, _tb, takeaway)
        self._source(s, source, note)
        return s

    def _finalize_checks(self):
        """保存前の全体検査: 結論帯の連続・多すぎ、本文の少ない章、目次のページ番号の記入。"""
        tp = sorted(set(self._takeaway_pages))
        for a, b in zip(tp, tp[1:]):
            if b == a + 1:
                self._warn_global(f"結論帯が p{a}・p{b} と連続している(2 枚続けない)")
        total = len(self.prs.slides)
        if total and len(tp) > total / 3 + 0.01:
            self._warn_global(f"結論帯が {len(tp)} 枚/{total} 枚で 1/3 を超えている。lead と違う情報を足せる枚だけに絞る")
        for num, pg, cnt in self._chapters:
            if cnt < 3:
                self._warn_global(f"第 {num} 章(p{pg})の本文が {cnt} 枚しかない。章を束ねるか中扉を省く")
        kinds = self._page_kind
        chs = self._chapters
        for k, (num, pg, cnt) in enumerate(chs):
            end = chs[k + 1][1] - 1 if k + 1 < len(chs) else total
            body = [kinds.get(p) for p in range(pg + 1, end + 1) if kinds.get(p)]
            if body and not any(x in ("diagram", "chart", "box") for x in body):   # 対比の箱も図解に数える
                self._warn_global(f"第 {num} 章に図解・グラフが 1 枚も無い(表と箇条書きだけ)。"
                                  "仕組みは概念図・layers、流れは cycle_flow・swimlane、目標との差は goal_bars にできないか見直す")
        body_all = [x for x in kinds.values() if x]
        nt = sum(1 for x in body_all if x == "table")
        if len(body_all) >= 5 and nt > 0.5 * len(body_all):
            self._warn_global(f"表の枚数が {nt}/{len(body_all)} 枚で半分を超えている。同じ属性の並びでない内容は図解にできないか見直す")
        for idx, tb in self._toc_page_boxes:
            pg = next((p for n_, p, _ in self._chapters if n_ == idx), None)
            if pg is None:
                tb._element.getparent().remove(tb._element)
            else:
                tb.text_frame.paragraphs[0].runs[0].text = f"p.{pg}"
        self._toc_page_boxes = []

    def _warn_global(self, msg):
        self.warnings.append(msg)

    def save(self, path):
        """保存する。PowerPoint で開いていると保存できないので分かるように止める。
        生成中の注意(self.warnings)があれば一覧を表示する(検証で該当ページを必ず見る)。"""
        self._finalize_checks()
        try:
            self.prs.save(path)
        except PermissionError:
            raise SystemExit(f"保存できない: {path} が PowerPoint で開かれている。閉じるか別名で保存する")
        if self.warnings:
            print(f"注意 {len(self.warnings)} 件(該当ページを必ず目視する):")
            for w in self.warnings:
                print("  -", w)
        return path
