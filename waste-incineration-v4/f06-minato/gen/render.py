# -*- coding: utf-8 -*-
"""LibreOffice（専用プロファイル file:///tmp/lo_F06）で PDF にし、pdftoppm で PNG にする。"""
import os
import shutil
import subprocess
import sys

from xlhelp import OUT, WORK

LO_PROFILE = "file:///tmp/lo_F06"


def to_pdf(path, outdir=WORK):
    os.makedirs(outdir, exist_ok=True)
    subprocess.run(["soffice", f"-env:UserInstallation={LO_PROFILE}", "--headless", "--convert-to", "pdf",
                    "--outdir", outdir, path], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                   timeout=600)
    pdf = os.path.join(outdir, os.path.splitext(os.path.basename(path))[0] + ".pdf")
    return pdf


def npages(pdf):
    out = subprocess.run(["pdfinfo", pdf], capture_output=True, text=True).stdout
    for line in out.splitlines():
        if line.startswith("Pages:"):
            return int(line.split()[1])
    return 0


def trim(path, pad=30):
    from PIL import Image, ImageChops
    im = Image.open(path).convert("RGB")
    bg = Image.new("RGB", im.size, (255, 255, 255))
    bb = ImageChops.difference(im, bg).getbbox()
    if bb:
        l, t, r, b = bb
        im.crop((max(0, l - pad), max(0, t - pad), min(im.width, r + pad), min(im.height, b + pad))).save(path)


def to_png(pdf, pages, stem, dpi=110, crop=False):
    """pages: 1 始まりのページ番号のリスト。出力は out/<stem>.png（1 枚）または <stem>_p1.png …"""
    res = []
    multi = len(pages) > 1
    for i, p in enumerate(pages, start=1):
        name = f"{stem}_p{i}.png" if multi else f"{stem}.png"
        tmp = os.path.join(WORK, "pp")
        subprocess.run(["pdftoppm", "-png", "-r", str(dpi), "-f", str(p), "-l", str(p), "-singlefile", pdf, tmp],
                       check=True)
        shutil.move(tmp + ".png", os.path.join(OUT, name))
        if crop:
            trim(os.path.join(OUT, name))
        res.append(name)
    return res


if __name__ == "__main__":
    pdf = to_pdf(sys.argv[1])
    print(pdf, npages(pdf))
