# -*- coding: utf-8 -*-
import os, subprocess, glob, shutil, datetime
from openpyxl.styles import Border, Side, Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(BASE, "files")
PNG = os.path.join(BASE, "build", "png_check")
BUILD = os.path.join(BASE, "build")
LO_PROFILE = "file:///tmp/lo_F03"
FIXED_DT = datetime.datetime(2026, 11, 13, 9, 0, 0)

thin = Side(style="thin")
hair = Side(style="hair")
med = Side(style="medium")
dot = Side(style="dotted")
NONE = Side(style=None)


def box(ws, r1, c1, r2, c2, side=thin, inner=thin, inner_h=None):
    inner_h = inner if inner_h is None else inner_h
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            ws.cell(r, c).border = Border(
                left=side if c == c1 else inner,
                right=side if c == c2 else inner,
                top=side if r == r1 else inner_h,
                bottom=side if r == r2 else inner_h)


def font(name="IPAGothic", size=9, bold=False, color=None, italic=False):
    return Font(name=name, size=size, bold=bold, color=color, italic=italic)


def fill(hexrgb):
    return PatternFill("solid", fgColor=hexrgb)


CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
LEFT_TOP = Alignment(horizontal="left", vertical="top", wrap_text=True)
RIGHT = Alignment(horizontal="right", vertical="center")
NOWRAP = Alignment(horizontal="left", vertical="center", wrap_text=False)


def put(ws, r, c, v, f=None, a=None, fl=None, nf=None):
    cell = ws.cell(r, c, v)
    if f is not None:
        cell.font = f
    if a is not None:
        cell.alignment = a
    if fl is not None:
        cell.fill = fl
    if nf is not None:
        cell.number_format = nf
    return cell


def widths(ws, ws_widths):
    for i, w in enumerate(ws_widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def page(ws, orient="landscape", paper="A3", fit_w=1, fit_h=0, area=None, titles=None, margins=None, scale=None):
    ws.page_setup.orientation = orient
    ws.page_setup.paperSize = {"A3": ws.PAPERSIZE_A3, "A4": ws.PAPERSIZE_A4, "B4": 12, "B5": 13}[paper]
    if scale:
        ws.sheet_properties.pageSetUpPr.fitToPage = False
        ws.page_setup.scale = scale
    else:
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = fit_w
        ws.page_setup.fitToHeight = fit_h
    if area:
        ws.print_area = area
    if titles:
        ws.print_title_rows = titles
    m = margins or (0.4, 0.4, 0.5, 0.5)
    ws.page_margins.left, ws.page_margins.right, ws.page_margins.top, ws.page_margins.bottom = m


def props(wb, title, creator, subject="", desc=""):
    wb.properties.title = title
    wb.properties.creator = creator
    wb.properties.lastModifiedBy = creator
    wb.properties.subject = subject
    wb.properties.description = desc
    wb.properties.created = FIXED_DT
    wb.properties.modified = FIXED_DT


def to_pdf(path, outdir=None):
    outdir = outdir or os.path.join(BUILD, "pdf")
    os.makedirs(outdir, exist_ok=True)
    subprocess.run(["soffice", "-env:UserInstallation=" + LO_PROFILE, "--headless",
                    "--convert-to", "pdf", "--outdir", outdir, path],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=600)
    return os.path.join(outdir, os.path.splitext(os.path.basename(path))[0] + ".pdf")


def pdf_pages(pdf):
    out = subprocess.run(["pdfinfo", pdf], capture_output=True, text=True).stdout
    for line in out.splitlines():
        if line.startswith("Pages:"):
            return int(line.split()[1])
    return 0


def to_png(pdf, stem, pages, dpi=110):
    """pdf の指定ページを out/ に PNG で書く。stem は拡張子込みの帳票ファイル名。"""
    res = []
    n = pdf_pages(pdf)
    pages = [p for p in pages if p <= n]
    tmp = os.path.join(BUILD, "png")
    os.makedirs(tmp, exist_ok=True)
    for i, p in enumerate(pages, start=1):
        pre = os.path.join(tmp, "x")
        for f in glob.glob(pre + "*"):
            os.remove(f)
        subprocess.run(["pdftoppm", "-r", str(dpi), "-f", str(p), "-l", str(p), "-png", pdf, pre], check=True)
        got = sorted(glob.glob(pre + "*.png"))[0]
        name = stem + (".png" if len(pages) == 1 else "_p%d.png" % i)
        os.makedirs(PNG, exist_ok=True)
        shutil.move(got, os.path.join(PNG, name))
        res.append(name)
    return res


def find_pages(pdf, text, maxn=1):
    """text を含むページ番号（1 始まり）"""
    n = pdf_pages(pdf)
    hits = []
    for p in range(1, n + 1):
        t = subprocess.run(["pdftotext", "-f", str(p), "-l", str(p), "-layout", pdf, "-"],
                           capture_output=True, text=True).stdout
        if text in t.replace(" ", ""):
            hits.append(p)
            if len(hits) >= maxn:
                break
    return hits
