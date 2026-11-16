# -*- coding: utf-8 -*-
import os, subprocess, glob, shutil, datetime as dt
from openpyxl.styles import Font, Border, Side, Alignment, PatternFill
from openpyxl.utils import get_column_letter

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(BASE, "files")
ANS = os.path.join(BASE, "answer", "files")
WORK = os.path.join(BASE, "work")
LO_PROFILE = "file:///tmp/lo_F11"
FONT = "IPAGothic"
FONTP = "IPAPGothic"

thin = Side(style="thin", color="000000")
hair = Side(style="hair", color="000000")
med = Side(style="medium", color="000000")
dot = Side(style="dotted", color="000000")
B_ALL = Border(left=thin, right=thin, top=thin, bottom=thin)
B_HAIR = Border(left=thin, right=thin, top=hair, bottom=hair)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
LEFT_TOP = Alignment(horizontal="left", vertical="top", wrap_text=True)
RIGHT = Alignment(horizontal="right", vertical="center")

def F(size=9, bold=False, color="000000", name=FONT, strike=False, italic=False):
    return Font(name=name, size=size, bold=bold, color=color, strike=strike, italic=italic)

def fill(hex6):
    return PatternFill("solid", fgColor=hex6)

def setw(ws, widths):
    for k, v in widths.items():
        ws.column_dimensions[k].width = v

def box(ws, rng, border=B_ALL):
    for row in ws[rng]:
        for c in row:
            c.border = border

def cell(ws, ref, value=None, font=None, align=None, border=None, fill_=None, fmt=None):
    c = ws[ref]
    if value is not None:
        c.value = value
    c.font = font or F()
    if align: c.alignment = align
    if border: c.border = border
    if fill_: c.fill = fill_
    if fmt: c.number_format = fmt
    return c

def props(wb, title, creator, subject=""):
    wb.properties.title = title
    wb.properties.creator = creator
    wb.properties.subject = subject
    wb.properties.lastModifiedBy = creator
    wb.properties.created = dt.datetime(2026, 11, 13, 9, 0, 0)
    wb.properties.modified = dt.datetime(2026, 11, 13, 9, 0, 0)

def page(ws, orientation="landscape", paper=9, fit_w=1, fit_h=0, title_rows=None, margins=(0.4, 0.4, 0.5, 0.5), area=None):
    """paper 9=A4, 8=A3"""
    ws.page_setup.orientation = orientation
    ws.page_setup.paperSize = paper
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = fit_w
    ws.page_setup.fitToHeight = fit_h
    l, r, t, b = margins
    ws.page_margins.left, ws.page_margins.right = l, r
    ws.page_margins.top, ws.page_margins.bottom = t, b
    if title_rows:
        ws.print_title_rows = title_rows
    if area:
        ws.print_area = area

def render(path, pages=None, dpi=110, prefix=None):
    """LibreOffice（自分専用プロファイル）で PDF にし、pdftoppm で PNG。
    pages: 1 始まりの頁番号のリスト（None なら全頁）。返り値は PNG のパス"""
    os.makedirs(WORK, exist_ok=True)
    stem = os.path.splitext(os.path.basename(path))[0]
    if path.lower().endswith(".pdf"):
        pdf = path
    else:
        subprocess.run(["soffice", "-env:UserInstallation=" + LO_PROFILE, "--headless",
                        "--convert-to", "pdf", "--outdir", WORK, path],
                       check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=300)
        pdf = os.path.join(WORK, stem + ".pdf")
    n = int([l for l in subprocess.run(["pdfinfo", pdf], capture_output=True, text=True).stdout.splitlines() if l.startswith("Pages:")][0].split()[1])
    pages = pages or list(range(1, n + 1))
    pages = [p for p in pages if p <= n]
    outs = []
    for i, p in enumerate(pages, 1):
        tmp = os.path.join(WORK, "tmp_%s" % i)
        subprocess.run(["pdftoppm", "-png", "-r", str(dpi), "-f", str(p), "-l", str(p), "-singlefile", pdf, tmp], check=True)
        name = (stem + ".png") if len(pages) == 1 else ("%s_p%d.png" % (stem, i))
        os.makedirs(os.path.join(WORK, "png"), exist_ok=True)
        dst = os.path.join(WORK, "png", name)
        shutil.move(tmp + ".png", dst)
        outs.append(dst)
    return outs, n

def render_csv(path, dpi=100):
    """CSV（Shift_JIS）を表計算で開いた見た目で確かめる：一時 xlsx（列幅自動・A3 横・幅 1 頁）にして PDF→PNG"""
    import csv
    from openpyxl import Workbook
    stem = os.path.splitext(os.path.basename(path))[0]
    os.makedirs(WORK, exist_ok=True)
    with open(path, "rb") as f:
        rows = list(csv.reader(f.read().decode("cp932").splitlines()))
    wb = Workbook(); ws = wb.active
    for r in rows:
        ws.append(r)
    for col in ws.columns:
        L = max(len(str(c.value or "")) for c in list(col)[1:]) if len(list(col)) > 1 else 8
        ws.column_dimensions[col[0].column_letter].width = min(max(4, L * 1.7), 48)
    for row in ws.iter_rows():
        for c in row:
            c.font = F(8)
    page(ws, "landscape", 8, 1, 0)
    tmp = os.path.join(WORK, stem + ".xlsx")
    wb.save(tmp)
    return render(tmp, [1], dpi)
