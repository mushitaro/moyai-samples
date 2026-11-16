# -*- coding: utf-8 -*-
"""LibreOffice（自分専用プロファイル）で PDF → pdftoppm で PNG"""
import os, subprocess, shutil, glob

PROFILE = "file:///tmp/lo_F09"
HERE = os.path.dirname(os.path.abspath(__file__))
WORK = os.path.join(HERE, "..", "tmp", "render")

def to_pdf(src, infilter=None):
    os.makedirs(WORK, exist_ok=True)
    stem = os.path.splitext(os.path.basename(src))[0]
    pdf = os.path.join(WORK, stem + ".pdf")
    if src.lower().endswith(".pdf"):
        return src
    if os.path.exists(pdf): os.remove(pdf)
    cmd = ["soffice", f"-env:UserInstallation={PROFILE}", "--headless"]
    if infilter: cmd.append(f"--infilter={infilter}")
    cmd += ["--convert-to", "pdf", "--outdir", WORK, src]
    subprocess.run(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=300, check=False)
    return pdf

def npages(pdf):
    out = subprocess.run(["pdfinfo", pdf], capture_output=True, text=True).stdout
    for line in out.splitlines():
        if line.startswith("Pages:"):
            return int(line.split()[1])
    return 0

def csv_view(src, enc):
    """CSV を表計算ソフトで開いたときの見た目を確かめるため、そのままの値で一時 xlsx に置く（納品物は CSV のまま）"""
    import csv
    from openpyxl import Workbook
    from openpyxl.styles import Font, Border, Side
    from openpyxl.utils import get_column_letter
    os.makedirs(WORK, exist_ok=True)
    with open(src, encoding=enc, newline="") as fp:
        rows = list(csv.reader(fp))
    wb = Workbook(); ws = wb.active
    th = Side(style="thin", color="BBBBBB")
    wmax = {}
    for r, row in enumerate(rows, 1):
        for c, v in enumerate(row, 1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.font = Font(name="IPAGothic", size=8)
            if r > 1: cell.border = Border(left=th, right=th, top=th, bottom=th)
            if r > 1: wmax[c] = max(wmax.get(c, 4), min(40, sum(2 if ord(ch) > 0x2000 and not (0xFF61 <= ord(ch) <= 0xFF9F) else 1 for ch in v) + 1))
    for c, w in wmax.items():
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.page_setup.orientation = "landscape"; ws.page_setup.paperSize = 8
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
    out = os.path.join(WORK, os.path.basename(src) + ".view.xlsx")
    wb.save(out)
    return out

def to_png(src, pages, outdir, infilter=None, dpi=110):
    if src.lower().endswith(".csv"):
        enc = "utf-8-sig" if infilter and infilter.endswith(",76") else "cp932"
        view = csv_view(src, enc)
        pdf = to_pdf(view)
        base0 = src
    else:
        pdf = to_pdf(src, infilter)
    n = npages(pdf)
    base = os.path.basename(src)
    outs = []
    pages = [p for p in pages if p <= n] or [1]
    for i, p in enumerate(pages):
        name = base + (".png" if len(pages) == 1 else f"_p{p}.png")
        tmp = os.path.join(WORK, "pg")
        subprocess.run(["pdftoppm", "-r", str(dpi), "-png", "-f", str(p), "-l", str(p), "-singlefile", pdf, tmp], check=True)
        shutil.move(tmp + ".png", os.path.join(outdir, name))
        outs.append(name)
    return outs, n
