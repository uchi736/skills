# -*- coding: utf-8 -*-
"""
make_ideal_kit — 納品デッキから「理想見本キット」を作る(PowerPoint COM)。

    python make_ideal_kit.py <deck.pptx> <kit.pptx> <枚番号:カテゴリ> [<枚番号:カテゴリ> ...]
    例: python make_ideal_kit.py 説明資料.pptx キット.pptx 2:要旨 5:概念図 15:表の図解版

- 指定した枚だけ残し、各枚を複製して「現状(BEFORE)」「理想版(AFTER)」の対にする
- AFTER の枚をユーザーが PowerPoint 上で直す(中身は全部ネイティブ図形)
- 直したキットは make_ideal_pairs.py で assets/ideal/ の PNG ペアに変換する
- 対応づけ用の目印は図形名 "IDEAL_MARKER|<ROLE>|<カテゴリ>" に持たせる(見た目のラベルは別)
- 出力先に同名ファイルがあれば止まる(上書きしない)
"""
import os
import re
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from export_preview import _winpath  # noqa: E402

MSO_HORIZONTAL = 1
PP_ALIGN_CENTER = 2


def rgb(hexstr):
    """'#RRGGBB' → PowerPoint COM の RGB 長整数(R + G*256 + B*65536)。"""
    h = hexstr.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return r + g * 256 + b * 65536


NAVY, GREY, WHITE = rgb("#01357A"), rgb("#A6A6A6"), rgb("#FFFFFF")


def add_marker(slide, width, role, cat):
    label = "現状(BEFORE)" if role == "BEFORE" else "理想版(AFTER) この枚を直してください"
    # タイトル帯の右側(ロゴの左)に置く。罫線とリード文の間は余白が足りず本文に重なる
    tb = slide.Shapes.AddTextbox(MSO_HORIZONTAL, width - 440, 27, 310, 22)
    tb.Name = f"IDEAL_MARKER|{role}|{cat}"
    tb.Fill.Visible = True
    tb.Fill.Solid()
    tb.Fill.ForeColor.RGB = GREY if role == "BEFORE" else NAVY
    tb.Line.Visible = False
    tf = tb.TextFrame
    tf.WordWrap = True
    tf.MarginTop = tf.MarginBottom = 2
    tr = tf.TextRange
    tr.Text = label
    tr.Font.Name = "Yu Gothic"
    tr.Font.Size = 11
    tr.Font.Bold = True
    tr.Font.Color.RGB = WHITE
    tr.ParagraphFormat.Alignment = PP_ALIGN_CENTER


def main(src, dst, picks):
    src, dst = _winpath(src), _winpath(dst)
    if not os.path.isfile(src):
        raise SystemExit(f"デッキがありません: {src}")
    if os.path.exists(dst):
        raise SystemExit(f"出力先が既にあります(上書きしない): {dst}")
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        raise SystemExit("pywin32 がありません(pip install pywin32)。Windows+PowerPoint が必要")
    tmp = os.path.splitext(dst)[0] + "_tmp.pptx"
    shutil.copy(src, tmp)
    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("PowerPoint.Application")
    try:
        app.DisplayAlerts = 1
    except Exception:
        pass
    pres = app.Presentations.Open(tmp, False, False, False)
    try:
        keep = {n for n, _ in picks}
        total = pres.Slides.Count
        bad = sorted(n for n in keep if n < 1 or n > total)
        if bad:
            raise SystemExit(f"枚番号が範囲外: {bad}(全{total}枚)")
        for i in range(total, 0, -1):
            if i not in keep:
                pres.Slides(i).Delete()
        order = sorted(picks)                      # 元デッキの順に並ぶ
        width = pres.PageSetup.SlideWidth
        for k, (_, cat) in enumerate(order):
            idx = 2 * k + 1
            pres.Slides(idx).Duplicate()           # idx+1 に複製が入る
            add_marker(pres.Slides(idx), width, "BEFORE", cat)
            add_marker(pres.Slides(idx + 1), width, "AFTER", cat)
        pres.SaveAs(dst)
        print(f"kit: {dst}  slides: {pres.Slides.Count}  pairs: {len(order)}")
    finally:
        pres.Close()
        try:
            if app.Presentations.Count == 0:
                app.Quit()
        except Exception:
            pass
        if os.path.exists(tmp):
            os.remove(tmp)


def parse_picks(args):
    picks, seen = [], set()
    for a in args:
        m = re.fullmatch(r"\s*(\d+)\s*[:：]\s*(.+?)\s*", a)
        if not m:
            raise SystemExit(f"指定の形が違う: {a!r}(枚番号:カテゴリ。例 5:概念図)")
        n = int(m.group(1))
        if n in seen:
            raise SystemExit(f"枚番号 {n} が重複している")
        seen.add(n)
        picks.append((n, m.group(2)))
    return picks


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if len(sys.argv) < 4:
        raise SystemExit(__doc__)
    main(sys.argv[1], sys.argv[2], parse_picks(sys.argv[3:]))
