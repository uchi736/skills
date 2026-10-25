# -*- coding: utf-8 -*-
r"""
export_preview — 生成した pptx を PowerPoint 本体(COM)でレンダリングして PNG 化する検証ツール。

    python export_preview.py deck.pptx [outdir]

- 各スライドの実画像(フォント・チャート・表も本物)を outdir(既定 <deck>_pptx_png)に
  スライド1.PNG, スライド2.PNG ... として出力する
- 全スライドを 2 列に並べた一覧(<deck>_ALL.png)も出力する
- PowerPoint / pywin32 が無い環境では理由を表示して終わる(生成物自体は有効)
- パスは Windows 形式に正規化する(Git Bash の /c/Users/... を渡しても C:\c を作らない)
"""
import glob
import os
import re
import sys

_PNG_RE = re.compile(r"^(スライド|Slide)(\d+)\.png$", re.IGNORECASE)


def _winpath(p):
    r"""Git Bash / MSYS 形式 (/c/Users/...) を C:\Users\... に直してから絶対パス化する。"""
    p = str(p)
    m = re.match(r"^/([a-zA-Z])/(.*)$", p)
    if m and os.name == "nt":
        p = f"{m.group(1).upper()}:\\" + m.group(2).replace("/", "\\")
    return os.path.abspath(p)


def _slide_pngs(outdir):
    found = {}
    for f in os.listdir(outdir):
        m = _PNG_RE.match(f)
        if m:
            found[os.path.join(outdir, f).lower()] = (int(m.group(2)), os.path.join(outdir, f))
    return [p for _, p in sorted(found.values())]


def export_pngs(pptx_path, outdir=None, width=1440):
    pptx_path = _winpath(pptx_path)
    if not os.path.isfile(pptx_path):
        raise SystemExit(f"ファイルがありません: {pptx_path}")
    outdir = _winpath(outdir) if outdir else os.path.splitext(pptx_path)[0] + "_pptx_png"
    os.makedirs(outdir, exist_ok=True)
    for old in _slide_pngs(outdir):          # 前回のスライド画像だけ消す(他のファイルは触らない)
        try:
            os.remove(old)
        except OSError:
            pass
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        raise SystemExit("pywin32 がありません(pip install pywin32)。PNG 検証は Windows+PowerPoint が必要")
    pythoncom.CoInitialize()
    try:
        app = win32com.client.DispatchEx("PowerPoint.Application")
    except Exception as e:
        raise SystemExit(f"PowerPoint を起動できません: {e}")
    try:
        app.DisplayAlerts = 1                # ppAlertsNone
    except Exception:
        pass
    pres = None
    try:
        pres = app.Presentations.Open(pptx_path, True, False, False)  # ReadOnly, Untitled, WithWindow
        sw, sh = pres.PageSetup.SlideWidth, pres.PageSetup.SlideHeight
        pres.Export(outdir, "PNG", width, int(round(width * sh / sw)))
    except Exception as e:
        raise SystemExit(f"PowerPoint でのレンダリングに失敗: {e}\n"
                         "(同じファイルを PowerPoint で開いている・保護ビュー・パスの文字を確認)")
    finally:
        if pres is not None:
            try:
                pres.Close()
            except Exception:
                pass
        try:
            if app.Presentations.Count == 0:
                app.Quit()
        except Exception:
            pass
    return _slide_pngs(outdir)


def contact_sheet(files, out_path, cols=2, thumb_w=760):
    """スライド画像を cols 列に並べた一覧 PNG。番号を左上に振る。"""
    from PIL import Image, ImageDraw
    if not files:
        return None
    thumbs = []
    for f in files:
        with Image.open(f) as im:
            r = thumb_w / im.width
            thumbs.append(im.convert("RGB").resize((thumb_w, int(im.height * r))))
    th = max(t.height for t in thumbs)
    rows = -(-len(thumbs) // cols)
    pad = 10
    sheet = Image.new("RGB", (cols * thumb_w + (cols + 1) * pad, rows * th + (rows + 1) * pad), "#666666")
    dr = ImageDraw.Draw(sheet)
    for i, t in enumerate(thumbs):
        r, c = divmod(i, cols)
        x, y = pad + c * (thumb_w + pad), pad + r * (th + pad)
        sheet.paste(t, (x, y))
        dr.rectangle([x, y, x + 34, y + 20], fill="#1F2D3D")
        dr.text((x + 6, y + 4), str(i + 1), fill="#FFFFFF")
    out_path = _winpath(out_path)
    sheet.save(out_path)
    return out_path


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if len(sys.argv) < 2 or sys.argv[1] in ("-h", "--help"):
        raise SystemExit(__doc__)
    pptx = _winpath(sys.argv[1])
    files = export_pngs(pptx, sys.argv[2] if len(sys.argv) > 2 else None)
    print(f"exported {len(files)} slides -> {os.path.dirname(files[0]) if files else '-'}")
    sheet = contact_sheet(files, os.path.splitext(pptx)[0] + "_ALL.png")
    print("contact sheet:", sheet)
