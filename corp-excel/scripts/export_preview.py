# -*- coding: utf-8 -*-
"""
export_preview — 生成した xlsx を Excel 本体(COM)で PDF 化し、PNG に変換する検証ツール。

    python export_preview.py book.xlsx

- Excel の実レンダリング(フォント・書式・チャート・印刷レイアウトすべて本物)で確認できる
- 各ページPNG + 全ページ縦連結の contact sheet(<book>_ALL.png)を出力
"""
import sys, os


def export_pngs(xlsx_path, outdir=None, dpi=110):
    xlsx_path = os.path.abspath(xlsx_path)
    if outdir is None:
        outdir = os.path.splitext(xlsx_path)[0] + "_xlsx_png"  # pptx側と衝突しない名前
    os.makedirs(outdir, exist_ok=True)
    pdf_path = os.path.join(outdir, "_render.pdf")
    import win32com.client
    xl = win32com.client.Dispatch("Excel.Application")
    xl.DisplayAlerts = False
    wb = xl.Workbooks.Open(xlsx_path, ReadOnly=True)
    try:
        wb.ExportAsFixedFormat(0, pdf_path)  # 0 = xlTypePDF
    finally:
        wb.Close(False)
        try:
            if xl.Workbooks.Count == 0:
                xl.Quit()
        except Exception:
            pass
    import fitz
    doc = fitz.open(pdf_path)
    files = []
    for i, page in enumerate(doc):
        p = os.path.join(outdir, f"page{i+1:02d}.png")
        page.get_pixmap(dpi=dpi).save(p)
        files.append(p)
    doc.close()
    return files


def contact_sheet(files, out_path, max_w=900):
    from PIL import Image
    imgs = [Image.open(f) for f in files]
    if not imgs:
        return None
    scaled = []
    for im in imgs:
        r = max_w / im.width
        scaled.append(im.resize((max_w, int(im.height * r))))
    pad = 12
    H = sum(im.height for im in scaled) + pad * (len(scaled) + 1)
    sheet = Image.new("RGB", (max_w + 2 * pad, H), "#666666")
    y = pad
    for im in scaled:
        sheet.paste(im, (pad, y))
        y += im.height + pad
    sheet.save(out_path)
    return out_path


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    xlsx = sys.argv[1]
    files = export_pngs(xlsx)
    print(f"exported {len(files)} pages")
    sheet = contact_sheet(files, os.path.splitext(os.path.abspath(xlsx))[0] + "_ALL.png")
    print("contact sheet:", sheet)
