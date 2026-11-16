# -*- coding: utf-8 -*-
import os
import subprocess
import glob
import shutil
from openpyxl.styles import Font, Border, Side, Alignment, PatternFill

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
OUT = os.path.join(ROOT, "files")
WORK = os.path.join(ROOT, "work")
ANS = os.path.join(ROOT, "answer", "files")
LO_PROFILE = "file:///tmp/lo_F07"
FN = "【架空】かざはやエコプラザ"

FONT = "IPAGothic"
PFONT = "IPAPGothic"


def font(sz=9, b=False, color=None, name=FONT, i=False):
    kw = dict(name=name, size=sz, bold=b, italic=i)
    if color:
        kw["color"] = color
    return Font(**kw)


thin = Side(style="thin", color="000000")
hair = Side(style="hair", color="000000")
med = Side(style="medium", color="000000")
dot = Side(style="dotted", color="000000")
dbl = Side(style="double", color="000000")
NONE = Side(style=None)


def box(l=thin, r=thin, t=thin, b=thin):
    return Border(left=l, right=r, top=t, bottom=b)


ALL_THIN = box()
C = Alignment(horizontal="center", vertical="center", wrap_text=True)
CL = Alignment(horizontal="left", vertical="center", wrap_text=False)
CLW = Alignment(horizontal="left", vertical="center", wrap_text=True)
CT = Alignment(horizontal="left", vertical="top", wrap_text=True)
CR = Alignment(horizontal="right", vertical="center")


def fill(hexrgb, pattern="solid"):
    if pattern == "solid":
        return PatternFill("solid", fgColor=hexrgb)
    return PatternFill(pattern, fgColor=hexrgb, bgColor="FFFFFF")


def setp(wb, title, author, subject=""):
    wb.properties.title = title
    wb.properties.creator = author
    wb.properties.subject = subject or title
    wb.properties.lastModifiedBy = author


def outpath(name):
    return os.path.join(OUT, name)


def to_pdf(src, outdir=None, filt=None):
    outdir = outdir or WORK
    os.makedirs(outdir, exist_ok=True)
    conv = "pdf" if not filt else f"pdf:{filt}"
    cmd = ["soffice", f"-env:UserInstallation={LO_PROFILE}", "--headless", "--convert-to", conv, "--outdir", outdir, src]
    subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=300)
    base = os.path.splitext(os.path.basename(src))[0]
    return os.path.join(outdir, base + ".pdf")


def pdf_to_png(pdf, out_stem, pages, dpi=80):
    """pages: 1 始まりのページ番号リスト。PNG は out/ へ『<帳票ファイル名>.png』『_p1』…で置く。"""
    res = []
    multi = len(pages) > 1
    for i, p in enumerate(pages, 1):
        tmp = os.path.join(WORK, "pngtmp")
        subprocess.run(["pdftoppm", "-f", str(p), "-l", str(p), "-r", str(dpi), "-png", "-singlefile", pdf, tmp], check=True)
        name = f"{out_stem}_p{i}.png" if multi else f"{out_stem}.png"
        shutil.move(tmp + ".png", os.path.join(OUT, name))
        res.append(name)
    return res


def pdf_pages(pdf):
    r = subprocess.run(["pdfinfo", pdf], capture_output=True, text=True)
    for line in r.stdout.splitlines():
        if line.startswith("Pages:"):
            return int(line.split()[1])
    return 0


# ---------------------------------------------------------------- 正解の記録
REC = {}


def rec(file, kind, fields, r=None, sheet=None, seq=None):
    d = {"kind": kind, "fields": fields}
    if sheet is not None:
        d["sheet"] = sheet
    if r is not None:
        d["r"] = r
    if seq is not None:
        d["seq"] = seq
    REC.setdefault(file, {"plan": False, "records": []})["records"].append(d)


def set_plan(file, flag=True):
    REC.setdefault(file, {"plan": flag, "records": []})["plan"] = flag


def save_answers():
    import json
    os.makedirs(ANS, exist_ok=True)
    for f, v in REC.items():
        stem = os.path.splitext(f)[0]
        with open(os.path.join(ANS, stem + ".json"), "w", encoding="utf-8") as fh:
            json.dump({"file": f, "plan": v["plan"], "records": v["records"]}, fh, ensure_ascii=False, indent=0)
