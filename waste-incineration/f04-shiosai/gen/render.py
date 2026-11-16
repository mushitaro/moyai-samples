# -*- coding: utf-8 -*-
"""LibreOffice（専用プロファイル /tmp/lo_F04）で PDF にし、pdftoppm で PNG にする。"""
import os, subprocess, shutil, glob, sys
from xl import OUT, BUILD

PROFILE = "file:///tmp/lo_F04"

def to_pdf(path, outdir=BUILD):
    os.makedirs(outdir, exist_ok=True)
    subprocess.run(["soffice", "-env:UserInstallation=" + PROFILE, "--headless", "--convert-to", "pdf",
                    "--outdir", outdir, path], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=300)
    return os.path.join(outdir, os.path.splitext(os.path.basename(path))[0] + ".pdf")

def npages(pdf):
    out = subprocess.run(["pdfinfo", pdf], capture_output=True, text=True).stdout
    for l in out.splitlines():
        if l.startswith("Pages:"): return int(l.split()[1])
    return 0

def png(path, pages=None, dpi=80):
    """path: 納品ファイル（xlsx または pdf）。pages: 1 始まりのページ番号のリスト。返り値：PNG のファイル名"""
    pdf = path if path.endswith(".pdf") else to_pdf(path)
    n = npages(pdf)
    pages = pages or [1]
    pages = [p for p in pages if p <= n]
    stem = os.path.basename(path)
    names = []
    for i, p in enumerate(pages, 1):
        name = stem + (".png" if len(pages) == 1 else "_p%d.png" % i)
        tmp = os.path.join(BUILD, "tmp_png")
        subprocess.run(["pdftoppm", "-png", "-r", str(dpi), "-f", str(p), "-l", str(p), "-singlefile", pdf, tmp], check=True)
        os.makedirs(os.path.join(BUILD, "png"), exist_ok=True); shutil.move(tmp + ".png", os.path.join(BUILD, "png", name))
        names.append(name)
    return names, n
