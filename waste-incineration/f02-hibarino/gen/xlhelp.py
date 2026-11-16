# -*- coding: utf-8 -*-
import os
import subprocess
import shutil
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(BASE, "work", "png")
os.makedirs(OUT, exist_ok=True)
WORK = os.path.join(BASE, "work")
LO_PROFILE = "file:///tmp/lo_F02"

thin = Side(style="thin")
hair = Side(style="hair")
med = Side(style="medium")
dot = Side(style="dotted")
B_ALL = Border(left=thin, right=thin, top=thin, bottom=thin)
B_HAIR = Border(left=hair, right=hair, top=hair, bottom=hair)


def font(size=9, bold=False, name="IPAGothic", color=None):
    return Font(name=name, size=size, bold=bold, color=color)


def fill(hex_):
    return PatternFill("solid", fgColor=hex_)


CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
LEFT_TOP = Alignment(horizontal="left", vertical="top", wrap_text=True)
RIGHT = Alignment(horizontal="right", vertical="center")


def put(ws, r, c, v, f=None, a=None, b=None, fi=None, nf=None):
    cell = ws.cell(row=r, column=c, value=v)
    if f is not None:
        cell.font = f
    if a is not None:
        cell.alignment = a
    if b is not None:
        cell.border = b
    if fi is not None:
        cell.fill = fi
    if nf is not None:
        cell.number_format = nf
    return cell


def box(ws, r1, c1, r2, c2, border=B_ALL):
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            ws.cell(row=r, column=c).border = border


def outline(ws, r1, c1, r2, c2, side=med):
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = ws.cell(row=r, column=c)
            b = cell.border
            cell.border = Border(left=side if c == c1 else b.left, right=side if c == c2 else b.right,
                                 top=side if r == r1 else b.top, bottom=side if r == r2 else b.bottom)


def props(wb, title, author, created=None):
    wb.properties.title = title
    wb.properties.creator = author
    wb.properties.lastModifiedBy = author
    if created is not None:
        wb.properties.created = created
        wb.properties.modified = created


def page(ws, landscape=True, paper=8, fit_w=1, fit_h=0, area=None, titles=None, margins=(0.4, 0.4, 0.5, 0.5),
         footer=None):
    """paper: 9=A4, 8=A3"""
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.page_setup.paperSize = paper
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = fit_w
    ws.page_setup.fitToHeight = fit_h
    if area:
        ws.print_area = area
    if titles:
        ws.print_title_rows = titles
    l, r, t, b = margins
    ws.page_margins.left = l
    ws.page_margins.right = r
    ws.page_margins.top = t
    ws.page_margins.bottom = b
    if footer:
        ws.oddFooter.center.text = footer


def render(path, pages=None, stem=None, dpi=80):
    """LibreOffice（専用プロファイル）で PDF → pdftoppm で PNG。pages: 1 始まりのページ番号のリスト。"""
    os.makedirs(WORK, exist_ok=True)
    stem = stem or os.path.splitext(os.path.basename(path))[0]
    if path.lower().endswith(".pdf"):
        pdf = path
    else:
        subprocess.run(["soffice", f"-env:UserInstallation={LO_PROFILE}", "--headless", "--convert-to", "pdf",
                        "--outdir", WORK, path], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        pdf = os.path.join(WORK, os.path.splitext(os.path.basename(path))[0] + ".pdf")
    n = npages(pdf)
    pages = pages or [1]
    pages = [p for p in pages if p <= n]
    outs = []
    for p in pages:
        base = os.path.join(WORK, f"tmp_{p}")
        subprocess.run(["pdftoppm", "-r", str(dpi), "-f", str(p), "-l", str(p), "-png", "-singlefile", pdf, base], check=True)
        name = f"{stem}.png" if len(pages) == 1 else f"{stem}_p{p}.png"
        shutil.move(base + ".png", os.path.join(OUT, name))
        outs.append(name)
    return outs, n


def npages(pdf):
    out = subprocess.run(["pdfinfo", pdf], capture_output=True, text=True).stdout
    for line in out.splitlines():
        if line.startswith("Pages:"):
            return int(line.split()[1])
    return 0


def normalize_xlsx(path):
    """同じ命令で同じファイルになるよう、zip の時刻と core.xml の modified を固定する。"""
    import re
    import zipfile
    with zipfile.ZipFile(path) as z:
        items = [(i.filename, z.read(i.filename)) for i in z.infolist()]
    out = []
    for name, data in items:
        if name == "docProps/core.xml":
            s = data.decode("utf-8")
            m = re.search(r"<dcterms:created[^>]*>([^<]+)</dcterms:created>", s)
            if m:
                s = re.sub(r"(<dcterms:modified[^>]*>)[^<]+(</dcterms:modified>)", lambda x: x.group(1) + m.group(1) + x.group(2), s)
            data = s.encode("utf-8")
        out.append((name, data))
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in out:
            zi = zipfile.ZipInfo(name, date_time=(2026, 11, 13, 0, 0, 0))
            zi.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(zi, data)
