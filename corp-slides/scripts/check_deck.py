# -*- coding: utf-8 -*-
"""
check_deck — 納品前の機械検査(目視の取りこぼし防止)。

    python check_deck.py out.pptx

検査すること(1 件でもあれば終了コード 1):
- 再オープンできるか(壊れていないか)
- プレースホルダらしき文字列: [ ] / XXXX / 202X / ipsum / TODO / ○○(表・図形の中も見る)
- 空のプレースホルダ(PNG には写らず、編集画面でだけ「テキストを入力」と出る)
- 純黒 #000000 の文字(本文は #262626)
- 12pt 未満の文字(出典・注記・脚注・ページ番号・コピーライト以外。表とグラフ内の文字も含む)
- 「〜について」で終わるリード(主張を言い切っていない)
- タイトルの色が未指定(テンプレの純黒 #000000 を継承する疑い)
"""
import re
import sys

from pptx import Presentation
from pptx.util import Pt

PH_RE = re.compile(r"\[[^\]]*\]|XXXX|202X|ipsum|TODO|○○")
SMALL_OK = re.compile(r"^(出典|注|※|Copyright|©|\(単位|（単位)")


def _texts(shape):
    """shape 内の (text, runs) を、グループ・表も含めて列挙する。"""
    if shape.shape_type == 6:                       # GROUP
        for sh in shape.shapes:
            yield from _texts(sh)
        return
    if shape.has_text_frame:
        for p in shape.text_frame.paragraphs:
            yield "".join(r.text for r in p.runs), p.runs
    if getattr(shape, "has_table", False) and shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                for p in cell.text_frame.paragraphs:
                    yield "".join(r.text for r in p.runs), p.runs


def check(path):
    prs = Presentation(path)
    issues = []
    for i, s in enumerate(prs.slides, 1):
        for sh in s.shapes:
            for text, runs in _texts(sh):
                if not text.strip():
                    continue
                if PH_RE.search(text):
                    issues.append((i, "プレースホルダらしき文字", text[:50]))
                small = False
                for r in runs:
                    try:
                        if r.font.color and r.font.color.type is not None and str(r.font.color.rgb) == "000000":
                            issues.append((i, "純黒の文字", r.text[:40]))
                    except Exception:
                        pass
                    if r.font.size is not None and r.font.size < Pt(12) and r.text.strip() \
                            and r.text.strip() not in ("●", "■", "□", "・"):
                        small = True
                if small and not SMALL_OK.match(text.strip()) and not text.strip().isdigit():
                    issues.append((i, "12pt 未満の文字", text[:50]))
            if sh.has_text_frame and sh.text_frame.text.strip().endswith("について"):
                issues.append((i, "「〜について」で終わる文(言い切っていない)", sh.text_frame.text[:50]))
            if sh.is_placeholder and "TITLE" in str(sh.placeholder_format.type) and sh.has_text_frame:
                for p in sh.text_frame.paragraphs:        # タイトルはテンプレの tx1(#000000)を継承しやすい
                    for r in p.runs:
                        if r.text.strip() and (r.font.color is None or r.font.color.type is None):
                            issues.append((i, "タイトルの色が未指定(テンプレの純黒を継承する疑い)", r.text[:40]))
                            break
        for ph in s.placeholders:
            t = str(ph.placeholder_format.type)
            if "FOOTER" in t or "SLIDE_NUMBER" in t or "DATE" in t:
                continue
            if not ph.has_text_frame or not ph.text_frame.text.strip():
                issues.append((i, "空のプレースホルダ", t))
    return len(prs.slides), issues


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if len(sys.argv) < 2:
        raise SystemExit(__doc__)
    n, issues = check(sys.argv[1])
    for i, kind, txt in issues:
        print(f"p{i}: {kind}: {txt!r}")
    print(f"slides: {n}  issues: {len(issues)}")
    sys.exit(1 if issues else 0)
