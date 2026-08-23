# -*- coding: utf-8 -*-
"""html2png — HTML図解をヘッドレスEdge/ChromeでPNG化する(figure型の相棒)。

    python html2png.py fig.html [out.png] [--size 1600x680] [--scale 2]

- 図のキャンバスは <body> の width/height で決める(assets/fig-template.html 参照)
- --scale 2 で2倍解像度(スライド貼付で滲まない)
- .html がその図の「編集可能なマスター」。PNGと同じ場所に必ず残すこと
"""
import os, sys, subprocess, tempfile

CANDIDATES = [
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
]


def render(html_path, out_png=None, width=1600, height=680, scale=2):
    html_path = os.path.abspath(html_path)
    if out_png is None:
        out_png = os.path.splitext(html_path)[0] + ".png"
    out_png = os.path.abspath(out_png)
    exe = next((p for p in CANDIDATES if os.path.exists(p)), None)
    if exe is None:
        raise RuntimeError("ヘッドレス描画に使える Edge/Chrome が見つからない")
    profile = os.path.join(tempfile.gettempdir(), "corp_slides_headless_profile")
    url = "file:///" + html_path.replace("\\", "/")
    subprocess.run(
        [exe, "--headless", "--disable-gpu", "--hide-scrollbars",
         f"--user-data-dir={profile}", f"--screenshot={out_png}",
         f"--window-size={width},{height}", f"--force-device-scale-factor={scale}",
         "--default-background-color=FFFFFF", url],
        check=True, capture_output=True, timeout=60)
    if not os.path.exists(out_png):
        raise RuntimeError("PNG が生成されなかった: " + out_png)
    return out_png


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    args, pos, size, scale = sys.argv[1:], [], "1600x680", 2
    i = 0
    while i < len(args):
        if args[i] == "--size":
            size = args[i + 1]; i += 2
        elif args[i] == "--scale":
            scale = int(args[i + 1]); i += 2
        else:
            pos.append(args[i]); i += 1
    w, h = (int(v) for v in size.lower().split("x"))
    print("saved:", render(pos[0], pos[1] if len(pos) > 1 else None, w, h, scale))
