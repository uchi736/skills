# -*- coding: utf-8 -*-
"""
export_preview — 生成した pptx を PowerPoint 本体(COM)でレンダリングしてPNG化する検証ツール。

    python export_preview.py deck.pptx [outdir]

- 各スライドの実画像(PowerPointの実レンダリング=フォント・チャート・表も本物)を outdir に出力
- 全スライドを2列グリッドに並べた contact sheet(<deck>_ALL.png)も出力
- PowerPoint が無い環境ではエラーメッセージのみ(生成物自体は有効)
"""
import sys, os, glob, re


def export_pngs(pptx_path, outdir=None, width=1440):
    pptx_path = os.path.abspath(pptx_path)
    if outdir is None:
        outdir = os.path.splitext(pptx_path)[0] + "_pptx_png"  # xlsx側と衝突しない名前
    os.makedirs(outdir, exist_ok=True)
    # Windows はファイル名の大文字小文字を区別しないため set で重複除去する
    for old in set(glob.glob(os.path.join(outdir, "*.PNG")) + glob.glob(os.path.join(outdir, "*.png"))):
        try:
            os.remove(old)
        except FileNotFoundError:
            pass
    import win32com.client
    app = win32com.client.Dispatch("PowerPoint.Application")
    pres = app.Presentations.Open(pptx_path, True, False, False)  # ReadOnly, Untitled, WithWindow=False
    try:
        pres.Export(os.path.abspath(outdir), "PNG", width, int(width * 9 / 16))
    finally:
        pres.Close()
        try:
            if app.Presentations.Count == 0:
                app.Quit()
        except Exception:
            pass
    files = glob.glob(os.path.join(outdir, "*.PNG")) + glob.glob(os.path.join(outdir, "*.png"))

    def keynum(p):
        m = re.search(r"(\d+)", os.path.basename(p))
        return int(m.group(1)) if m else 0
    files = sorted(set(files), key=keynum)
    return files


def contact_sheet(files, out_path, cols=2, thumb_w=760):
    from PIL import Image
    thumbs = []
    for f in files:
        im = Image.open(f)
        r = thumb_w / im.width
        thumbs.append(im.resize((thumb_w, int(im.height * r))))
    if not thumbs:
        return None
    th = thumbs[0].height
    rows = -(-len(thumbs) // cols)
    pad = 10
    sheet = Image.new("RGB", (cols * thumb_w + (cols + 1) * pad,
                              rows * th + (rows + 1) * pad), "#666666")
    for i, t in enumerate(thumbs):
        r, c = divmod(i, cols)
        sheet.paste(t, (pad + c * (thumb_w + pad), pad + r * (th + pad)))
    sheet.save(out_path)
    return out_path


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    pptx = sys.argv[1]
    outdir = sys.argv[2] if len(sys.argv) > 2 else None
    files = export_pngs(pptx, outdir)
    print(f"exported {len(files)} slides")
    sheet = contact_sheet(files, os.path.splitext(os.path.abspath(pptx))[0] + "_ALL.png")
    print("contact sheet:", sheet)
