# -*- coding: utf-8 -*-
"""
corp_pptx — 自社ハウススタイルのプレゼン生成ライブラリ(corp-slides スキル同梱)v3

■ テンプレートの解決(push安全設計)
    assets/brand/format-16x9.pptx / format-4x3.pptx   ← 公式テンプレ(git管理外。ロゴ・意匠入り)
    → 存在すればそれを土台にする(本物のロゴ・フッター・扉デザインをマスター継承)
    → 無ければ中立なダミー(自前描画・「LOGO」表記・「© Your Company」)で同じ10型を生成
    ブランド資産をリポジトリに含めなくても、コードだけで動作する。

■ 判型
    Deck()                # 16:9(既定)
    Deck(aspect="4:3")    # 4:3
    座標はすべて W/H 比率で計算されるため、両判型で同じAPIが使える。

使い方(最小):
    import sys; sys.path.insert(0, os.path.expanduser("~/.claude/skills/corp-slides/scripts"))
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
import os, copy, re
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
LTBLUE = "#D6ECF7"  # 選択ハイライト(目次)
PERI   = "#C5D0E6"  # 帯
INK    = "#262626"  # 本文(純黒は使わない)
INK2   = "#4D4D4D"  # 補足・注釈
GREY   = "#A6A6A6"  # 罫線・弱い文字
GRID   = "#E3E7EC"  # チャートのグリッド線
YU     = "Yu Gothic"

SERIES_COLORS = [BLUE, CYANSUB, TEAL, STEEL, NAVY, "#66A6EA", PERI, CYAN]
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

    # ---------------------------------------------------------------- 基盤
    def _clear_slides(self):
        sldIdLst = self.prs.slides._sldIdLst
        for sldId in list(sldIdLst):
            rId = sldId.get(qn("r:id"))
            self.prs.part.drop_rel(rId)
            sldIdLst.remove(sldId)

    def _slide(self, layout_name, title=None):
        self._page += 1
        if self._brand:
            lay = self._layouts[layout_name]
            s = self.prs.slides.add_slide(lay)
            # python-pptx はフッター/ページ番号プレースホルダを複製しないため手動継承
            for ph in lay.placeholders:
                if ph.placeholder_format.type in (PP_PLACEHOLDER.FOOTER,
                                                  PP_PLACEHOLDER.SLIDE_NUMBER):
                    s.shapes._spTree.append(copy.deepcopy(ph._element))
            if title is not None and s.shapes.title is not None:
                s.shapes.title.text_frame.text = title
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
                          [[(title, 28, True, NAVY)]])
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

    def rect(self, s, x, y, w, h, fill, line=None, lw=1.0, shape=MSO_SHAPE.RECTANGLE, adj=None):
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
        return sp

    def text(self, s, x, y, w, h, paras, align="left", anchor="top",
             space_after=8, line_spacing=1.15, wrap=True):
        """paras: [[(text,size,bold,color[,font]), ...], ...]"""
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
                r = p.add_run()
                r.text = txt
                _set_font(r, size, bold, color, name)
        return tb

    def _lead(self, s, lead, y=None):
        y = self.BODY_TOP if y is None else y
        if not lead:
            return y + 0.85
        self.text(s, self.MX + 0.05, y + 0.12, self.CW - 0.15, 0.9,
                  [[("●", 13, True, BLUE), ("  " + lead, 15, False, INK)]])
        chars_per_line = int(self.CW * 4.0)
        lines = max(1, -(-len(lead) // chars_per_line))
        return y + 0.12 + 0.42 * lines + 0.42

    # ---------------------------------------------------------------- 階層マーカー(唯一の定義)
    # 全型共通の3段ラダー: ●(Lv1) → 　□(Lv2) → 　　‐(Lv3)。他のマーカーは使わない。
    def bullet(self, text, level=1, size=15, bold=False):
        if level <= 1:
            return [("●", int(size * 0.85), True, BLUE), ("  " + text, size, bold, INK)]
        if level == 2:
            s = max(10, size - 2)
            return [("　□", int(s * 0.9), True, STEEL), ("  " + text, s, bold, INK)]
        s = max(9, size - 3)
        return [("　　‐", s, False, STEEL), ("  " + text, s, bold, INK2)]

    def nest_paras(self, items, size=15, level=1):
        """ネストしたリストを段落列に変換。items: ["文", ("親", ["子", ("子2", ["孫"])]), ...]"""
        paras = []
        for it in items:
            if (isinstance(it, (tuple, list)) and len(it) == 2
                    and isinstance(it[1], (list, tuple))):
                head, kids = it
                paras.append(self.bullet(head, level, size))
                paras.extend(self.nest_paras(kids, size, min(level + 1, 3)))
            else:
                paras.append(self.bullet(it, level, size))
        return paras

    def bullet_para(self, t, size=15, color=INK, bold=False):
        """旧API互換。新規コードは bullet()/nest_paras() を使う"""
        return self.bullet(t, 1, size, bold)

    def note(self, text):
        s = self.prs.slides[-1]
        s.notes_slide.notes_text_frame.text = text

    def _source(self, s, source):
        if source:
            self.text(s, self.MX, self.H - 0.77, self.CW, 0.3,
                      [[("出典:" + source, 9, False, INK2)]])

    # ---------------------------------------------------------------- ① 表紙
    def cover(self, title, subtitle=None, label=None, date=None, dept=None, company=None):
        s = self._slide(LAYOUT_COVER, title=title)
        H = self.H
        if company is None:
            company = self.brand_cfg.get("company", "[会社名]")
        if label:
            self.text(s, self.MX, 0.316 * H, 0.8 * self.W, 0.4, [[(label, 15, False, INK)]])
        if subtitle:
            self.text(s, self.MX, 0.558 * H, 0.85 * self.W, 0.5,
                      [[("― " + subtitle + " ―", 17, True, BLUE_TX)]])
        if date:
            self.text(s, self.MX, 0.651 * H, 7.0, 0.4, [[(date, 14, False, INK)]])
        self.text(s, self.MX, 0.768 * H, 9.0, 0.4, [[(company, 16, True, INK)]])
        if dept:
            self.text(s, self.MX, 0.820 * H, 9.0, 0.4, [[(dept, 13, False, INK)]])
        return s

    # ---------------------------------------------------------------- ② 目次
    def toc(self, items, current=None, title="目次"):
        s = self._slide(LAYOUT_TOC, title=title)
        x = 0.237 * self.W
        y = 0.242 * self.H
        step = 0.118 * self.H
        for i, it in enumerate(items, 1):
            name, sub = (it if isinstance(it, (tuple, list)) else (it, None))
            if current == i:
                hl_w = min(self.W - x - 0.9,
                           1.3 + len(name) * 0.34 + (len(sub) * 0.20 if sub else 0) + 0.8)
                self.rect(s, x - 0.28, y - 0.14, hl_w, 0.72, LTBLUE)
            runs = [(f"{i}.", 21, True, STEEL), ("   " + name, 23, True, BLUE_TX)]
            if sub:
                runs.append(("  ― " + sub, 14, True, STEEL))
            self.text(s, x, y, self.W - x - 0.9, 0.5, [runs])
            y += step
        return s

    # ---------------------------------------------------------------- ③ 中扉
    def section(self, number, title):
        return self._slide(LAYOUT_SECTION, title=f"{number}. {title}")

    # ---------------------------------------------------------------- ④ 本文・箇条書き
    def content(self, title, lead=None, bullets=None, source=None):
        """bullets: ["文"] または [("親", ["子1", ("子2", ["孫"])])] — 最大3階層"""
        s = self._slide(LAYOUT_BODY, title=title)
        y = self._lead(s, lead)
        paras = self.nest_paras(bullets or [], size=15)
        if paras:
            self.text(s, self.MX + 0.15, y, self.CW - 0.3, self.BODY_BOT - y,
                      paras, space_after=12)
        self._source(s, source)
        return s

    # ---------------------------------------------------------------- ⑤ ボックス対比(2〜3列)
    def boxes(self, title, boxes, lead=None, arrow=False):
        s = self._slide(LAYOUT_BODY, title=title)
        y0 = self._lead(s, lead) - 0.15
        n = len(boxes)
        gap = 0.58 if arrow else 0.30
        w = (self.CW - gap * (n - 1)) / n
        bot = self.BODY_BOT - 0.10
        for i, (head, sub, items) in enumerate(boxes):
            x = self.MX + i * (w + gap)
            self.rect(s, x, y0, w, 0.55, BLUE)
            self.text(s, x, y0, w, 0.55, [[(head, 17, True, "#FFFFFF")]],
                      align="center", anchor="middle")
            yy = y0 + 0.55
            if sub:
                self.rect(s, x + 0.30, yy + 0.14, w - 0.60, 0.42, CYANSUB)
                self.text(s, x + 0.30, yy + 0.14, w - 0.60, 0.42,
                          [[(sub, 12, True, "#FFFFFF")]], align="center", anchor="middle")
                yy += 0.66
            self.rect(s, x, yy, w, bot - yy, BG)
            self.text(s, x + 0.30, yy + 0.20, w - 0.60, bot - yy - 0.35,
                      self.nest_paras(items, size=12), space_after=8)
            if arrow and i < n - 1:
                self.rect(s, x + w + 0.08, (y0 + bot) / 2 - 0.20, gap - 0.16, 0.40,
                          CYAN, shape=MSO_SHAPE.RIGHT_ARROW)
        return s

    # ---------------------------------------------------------------- ⑥ KPIハイライト
    def kpi(self, title, kpis, lead=None, source=None):
        s = self._slide(LAYOUT_BODY, title=title)
        y0 = self._lead(s, lead) - 0.05
        n = len(kpis)
        gap = 0.32
        w = (self.CW - gap * (n - 1)) / n
        h = min(3.60, self.BODY_BOT - 0.35 - y0)
        for i, item in enumerate(kpis):
            label, value, unit, note = (list(item) + [None] * 4)[:4]
            x = self.MX + i * (w + gap)
            self.rect(s, x, y0, w, h, BG)
            self.rect(s, x, y0, w, 0.07, BLUE)
            self.text(s, x + 0.2, y0 + 0.30, w - 0.4, 0.5,
                      [[(label, 14, True, INK)]], align="center")
            vcolor = BLUE
            if str(value).startswith(("+", "▲")):
                vcolor = TEAL
            elif str(value).startswith(("-", "−", "△", "▼")):
                vcolor = RED
            vsize = 40 if len(str(value)) <= 6 else (32 if len(str(value)) <= 9 else 26)
            runs = [(str(value), vsize, True, vcolor)]
            if unit:
                runs.append((" " + unit, 16, True, INK2))
            self.text(s, x + 0.2, y0 + h / 2 - 0.42, w - 0.4, 1.0, [runs],
                      align="center", anchor="middle")
            if note:
                self.text(s, x + 0.2, y0 + h - 0.62, w - 0.4, 0.5,
                          [[(note, 11, False, INK2)]], align="center")
        self._source(s, source)
        return s

    # ---------------------------------------------------------------- ⑦ ネイティブチャート
    def chart(self, title, kind, categories, series, lead=None, unit=None,
              points=None, value_fmt="#,##0", source=None):
        s = self._slide(LAYOUT_BODY, title=title)
        y0 = self._lead(s, lead)
        cw = self.CW * 0.64 if points else self.CW
        if unit:
            self.text(s, self.MX, y0 - 0.05, 3.0, 0.3,
                      [[(f"(単位:{unit})", 10, False, INK2)]])
        ch_y = y0 + 0.25
        ch_h = self.BODY_BOT - 0.10 - ch_y
        data = CategoryChartData()
        data.categories = categories
        for name, vals in series:
            data.add_series(name, vals)
        xl = {"bar": XL_CHART_TYPE.COLUMN_CLUSTERED,
              "stack": XL_CHART_TYPE.COLUMN_STACKED,
              "line": XL_CHART_TYPE.LINE}[kind]
        gf = s.shapes.add_chart(xl, Inches(self.MX), Inches(ch_y), Inches(cw), Inches(ch_h), data)
        c = gf.chart
        c.has_title = False
        c.font.size = Pt(10)
        c.font.name = YU
        c.font.color.rgb = _rgb(INK)
        c.has_legend = len(series) > 1
        if c.has_legend:
            c.legend.position = XL_LEGEND_POSITION.BOTTOM
            c.legend.include_in_layout = False
        plot = c.plots[0]
        try:
            plot.gap_width = 60 if kind != "line" else 150
        except Exception:
            pass
        for i, ser in enumerate(plot.series):
            col = _rgb(SERIES_COLORS[i % len(SERIES_COLORS)])
            if kind == "line":
                ser.format.line.color.rgb = col
                ser.format.line.width = Pt(2.5)
                ser.smooth = False
            else:
                ser.format.fill.solid()
                ser.format.fill.fore_color.rgb = col
                ser.format.line.fill.background()
        if kind == "bar" and len(series) == 1:
            plot.has_data_labels = True
            dl = plot.data_labels
            dl.number_format = value_fmt
            dl.number_format_is_linked = False
            dl.position = XL_LABEL_POSITION.OUTSIDE_END
            dl.font.size = Pt(10); dl.font.bold = True
            dl.font.color.rgb = _rgb(INK); dl.font.name = YU
        elif kind == "stack":
            plot.has_data_labels = True
            dl = plot.data_labels
            dl.number_format = value_fmt
            dl.number_format_is_linked = False
            dl.font.size = Pt(9); dl.font.bold = True
            dl.font.color.rgb = _rgb("#FFFFFF"); dl.font.name = YU
        cat = c.category_axis
        cat.has_major_gridlines = False
        cat.format.line.color.rgb = _rgb(GREY)
        cat.tick_labels.font.size = Pt(10)
        cat.tick_labels.font.name = YU
        val = c.value_axis
        if kind != "line":
            val.minimum_scale = 0  # 棒グラフの0始まりは省略不可(差の誇張防止)
        val.has_major_gridlines = True
        val.major_gridlines.format.line.color.rgb = _rgb(GRID)
        val.major_gridlines.format.line.width = Pt(0.75)
        val.format.line.fill.background()
        val.tick_labels.font.size = Pt(9)
        val.tick_labels.font.name = YU
        val.tick_labels.number_format = value_fmt
        val.tick_labels.number_format_is_linked = False
        if points:
            px = self.MX + cw + 0.32
            pw = self.W - self.MX - px
            self.rect(s, px, ch_y, pw, 0.46, NAVY)
            self.text(s, px, ch_y, pw, 0.46, [[("ポイント", 13, True, "#FFFFFF")]],
                      align="center", anchor="middle")
            self.rect(s, px, ch_y + 0.46, pw, ch_h - 0.46, BG)
            self.text(s, px + 0.22, ch_y + 0.72, pw - 0.44, ch_h - 0.95,
                      self.nest_paras(points, size=12), space_after=10)
        self._source(s, source)
        return s

    # ---------------------------------------------------------------- ⑧ ネイティブ表
    def table(self, title, headers, rows, lead=None, col_widths=None,
              first_col_header=True, font_size=12, unit=None, source=None):
        s = self._slide(LAYOUT_BODY, title=title)
        y0 = self._lead(s, lead)
        if unit:
            self.text(s, self.MX, y0 - 0.05, self.CW, 0.3,
                      [[(f"(単位:{unit})", 10, False, INK2)]], align="right")
        ty = y0 + 0.26
        ncols, nrows = len(headers), len(rows) + 1
        row_h = min(0.48, (self.BODY_BOT - 0.10 - ty) / nrows)
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
        for ri, row in enumerate(rows, start=1):
            zebra = BG if ri % 2 == 0 else None
            for cji, val in enumerate(row):
                txt = str(val)
                is_num = bool(_NUMERIC_RE.match(txt.replace(",", "").replace(" ", "")))
                color = INK
                bold = False
                fill = zebra
                align = "right" if is_num and cji > 0 else "left"
                if cji == 0 and first_col_header:
                    bold = True
                    fill = LTBLUE if ri % 2 == 0 else "#E8F2FA"
                    align = "left"
                if txt.startswith(("△", "-", "−", "▼")) and is_num:
                    color = RED
                put(tbl.cell(ri, cji), txt, body_size, bold, color, fill, align)
        self._source(s, source)
        return s

    # ---------------------------------------------------------------- ⑨ ロードマップ
    def timeline(self, title, phases, lead=None):
        s = self._slide(LAYOUT_BODY, title=title)
        y0 = self._lead(s, lead)
        n = len(phases)
        colors = RAMP[-n:] if n <= len(RAMP) else [RAMP[min(i, len(RAMP) - 1)] for i in range(n)]
        overlap = 0.24
        w = (self.CW + overlap * (n - 1)) / n
        ch_y = y0 + 0.18
        for i, (period, head, descs) in enumerate(phases):
            x = self.MX + i * (w - overlap)
            self.rect(s, x, ch_y, w, 0.88, colors[i],
                      shape=(MSO_SHAPE.CHEVRON if i > 0 else MSO_SHAPE.PENTAGON), adj=0.28)
            self.text(s, x + (0.30 if i > 0 else 0.12), ch_y, w - 0.42, 0.88,
                      [[(period, 15, True, "#FFFFFF")]], align="center", anchor="middle")
            bx = self.MX + i * (w - overlap) + (0.10 if i > 0 else 0)
            bw = w - overlap - 0.10
            by = ch_y + 1.10
            self.text(s, bx + 0.06, by, bw, 0.45, [[(head, 14, True, BLUE_TX)]])
            self.rect(s, bx + 0.06, by + 0.44, min(1.2, bw - 0.2), 0.035, CYAN)
            self.text(s, bx + 0.06, by + 0.62, bw - 0.10, self.BODY_BOT - by - 0.6,
                      self.nest_paras(descs, size=11), space_after=6)
        return s

    # ---------------------------------------------------------------- ⑩ メッセージ / 結び
    def message(self, text, sub=None, contact=None):
        s = self._slide(LAYOUT_LAST)
        size = 28 if len(text) <= 24 else 23
        paras = [[(text, size, True, NAVY)]]
        if sub:
            paras.append([(sub, 15, False, INK2)])
        self.text(s, 0.10 * self.W, 0.217 * self.H, 0.80 * self.W, 1.5,
                  paras, align="center", space_after=12)
        if contact:
            self.text(s, 0.10 * self.W, 0.787 * self.H, 0.80 * self.W, 1.0,
                      [[(c, 12, False, INK2)] for c in contact],
                      align="center", space_after=4)
        return s

    # ---------------------------------------------------------------- 保存
    def save(self, path):
        self.prs.save(path)
        return path
