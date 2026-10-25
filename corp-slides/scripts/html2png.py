# -*- coding: utf-8 -*-
"""html2png — HTML 図解をヘッドレス Edge / Chrome で PNG にする(figure 型の相棒)。

    python html2png.py fig.html [out.png] [--size 1600x680] [--scale 2]

- 図のキャンバスは <body> の width/height で決める(assets/fig-template.html を複製して使う)
- --size は「幅x高さ」(x / × / X どれでも可)。--scale 2 で 2 倍解像度(貼っても滲まない)
- .html がその図の「編集可能なマスター」。PNG と同じ場所に必ず残す
"""
import argparse
import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path


def _candidates():
    pf = os.environ.get("ProgramFiles", "C:/Program Files")
    pf86 = os.environ.get("ProgramFiles(x86)", "C:/Program Files (x86)")
    la = os.environ.get("LOCALAPPDATA", "")
    paths = [os.path.join(pf86, "Microsoft", "Edge", "Application", "msedge.exe"),
             os.path.join(pf, "Microsoft", "Edge", "Application", "msedge.exe"),
             os.path.join(pf, "Google", "Chrome", "Application", "chrome.exe"),
             os.path.join(pf86, "Google", "Chrome", "Application", "chrome.exe")]
    if la:
        paths += [os.path.join(la, "Google", "Chrome", "Application", "chrome.exe"),
                  os.path.join(la, "Microsoft", "Edge", "Application", "msedge.exe")]
    for name in ("msedge", "chrome", "google-chrome", "chromium", "chromium-browser"):
        w = shutil.which(name)
        if w:
            paths.append(w)
    return paths


def render(html_path, out_png=None, width=1600, height=680, scale=2):
    html = Path(html_path).resolve()
    if not html.is_file():
        raise SystemExit(f"HTML がありません: {html}")
    out = Path(out_png).resolve() if out_png else html.with_suffix(".png")
    exe = next((p for p in _candidates() if p and os.path.exists(p)), None)
    if exe is None:
        raise SystemExit("ヘッドレス描画に使える Edge / Chrome が見つからない(インストール場所を確認)")
    profile = tempfile.mkdtemp(prefix="corp_slides_headless_")   # 毎回まっさら(既存ブラウザと衝突しない)
    try:
        r = subprocess.run(
            [exe, "--headless", "--disable-gpu", "--hide-scrollbars",
             f"--user-data-dir={profile}", f"--screenshot={out}",
             f"--window-size={width},{height}", f"--force-device-scale-factor={scale}",
             "--default-background-color=FFFFFFFF", html.as_uri()],
            capture_output=True, timeout=90)
    except subprocess.TimeoutExpired:
        raise SystemExit("ブラウザが 90 秒で終わらなかった(HTML の外部読み込みを確認)")
    finally:
        shutil.rmtree(profile, ignore_errors=True)
    if not out.exists():
        err = (r.stderr or b"").decode("utf-8", "replace")[-400:]
        raise SystemExit(f"PNG が生成されなかった: {out}\n{err}")
    return str(out)


def _size(v):
    m = re.fullmatch(r"\s*(\d+)\s*[xX×]\s*(\d+)\s*", v)
    if not m:
        raise argparse.ArgumentTypeError("--size は 1600x680 の形で指定する")
    return int(m.group(1)), int(m.group(2))


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    ap = argparse.ArgumentParser(description="HTML 図解を PNG にする")
    ap.add_argument("html")
    ap.add_argument("out", nargs="?")
    ap.add_argument("--size", type=_size, default=(1600, 680))
    ap.add_argument("--scale", type=int, default=2)
    a = ap.parse_args()
    print("saved:", render(a.html, a.out, a.size[0], a.size[1], a.scale))
