# -*- coding: utf-8 -*-
import datetime as dt
import os

from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(BASE, "files")
WORK = os.path.join(BASE, "work")

FONT_G = "IPAGothic"     # 旧い様式（等幅ゴシック）
FONT_P = "IPAPGothic"    # 新しい様式（プロポーショナル）

thin = Side(style="thin", color="000000")
hair = Side(style="hair", color="000000")
med = Side(style="medium", color="000000")
dot = Side(style="dotted", color="000000")
B_ALL = Border(left=thin, right=thin, top=thin, bottom=thin)
B_HAIR = Border(left=hair, right=hair, top=hair, bottom=hair)


def font(name=FONT_G, size=9, bold=False, color="000000", italic=False, u=None):
    return Font(name=name, size=size, bold=bold, color=color, italic=italic, underline=u)


def fill(hexcolor):
    return PatternFill("solid", fgColor=hexcolor)


AL_C = Alignment(horizontal="center", vertical="center", wrap_text=True)
AL_L = Alignment(horizontal="left", vertical="center", wrap_text=False)
AL_LW = Alignment(horizontal="left", vertical="center", wrap_text=True)
AL_LT = Alignment(horizontal="left", vertical="top", wrap_text=True)
AL_R = Alignment(horizontal="right", vertical="center")


def put(ws, r, c, v, f=None, al=None, b=None, fl=None, nf=None):
    cell = ws.cell(row=r, column=c, value=v)
    if f is not None:
        cell.font = f
    if al is not None:
        cell.alignment = al
    if b is not None:
        cell.border = b
    if fl is not None:
        cell.fill = fl
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


def merge(ws, r1, c1, r2, c2, v=None, f=None, al=AL_C, b=B_ALL, fl=None):
    if v is not None or f is not None:
        put(ws, r1, c1, v, f, al, None, fl)
    if b is not None:
        box(ws, r1, c1, r2, c2, b)
    if fl is not None:
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                ws.cell(row=r, column=c).fill = fl
    if (r1, c1) != (r2, c2):
        ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)


def note(cell, text, author):
    cm = Comment(text, author)
    cm.width = 260
    cm.height = 90
    cell.comment = cm


def props(wb, title, author, subject="", created=None, modified_by=None, company=None):
    wb.properties.title = title
    wb.properties.creator = author
    wb.properties.subject = subject
    wb.properties.lastModifiedBy = modified_by or author
    if created:
        wb.properties.created = created
        wb.properties.modified = created
    wb.properties.keywords = "【架空】"


def page(ws, landscape=True, paper=9, fit_w=1, fit_h=0, area=None, titles=None, margins=(0.4, 0.4, 0.5, 0.5),
         center=True, header=None, footer=None):
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.page_setup.paperSize = paper  # 9=A4, 8=A3
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = fit_w
    ws.page_setup.fitToHeight = fit_h
    if area:
        ws.print_area = area
    if titles:
        ws.print_title_rows = titles
    l, r, t, b = margins
    ws.page_margins.left, ws.page_margins.right, ws.page_margins.top, ws.page_margins.bottom = l, r, t, b
    ws.print_options.horizontalCentered = center
    if footer:
        ws.oddFooter.center.text = footer
        ws.oddFooter.center.size = 8
    if header:
        ws.oddHeader.right.text = header
        ws.oddHeader.right.size = 8


WAREKI = [(dt.date(2019, 5, 1), "R", 2018), (dt.date(1989, 1, 8), "H", 1988)]


def fy_wa(fy, style="H"):
    """年度 → 和暦の年度表記"""
    if fy >= 2019:
        n = fy - 2018
        return "R元" if n == 1 else f"R{n}"
    return f"H{fy - 1988}"


def fy_wa_long(fy):
    if fy >= 2019:
        n = fy - 2018
        return "令和元年度" if n == 1 else f"令和{n}年度"
    return f"平成{fy - 1988}年度"


def d_wa(d, sep="."):
    if isinstance(d, str):
        d = dt.date.fromisoformat(d)
    if d >= dt.date(2019, 5, 1):
        y = f"R{d.year - 2018}"
    else:
        y = f"H{d.year - 1988}"
    return f"{y}{sep}{d.month:02d}{sep}{d.day:02d}"


def d_wa_long(d):
    if isinstance(d, str):
        d = dt.date.fromisoformat(d)
    if d >= dt.date(2019, 5, 1):
        n = d.year - 2018
        y = "令和元" if n == 1 else f"令和{n}"
    else:
        y = f"平成{d.year - 1988}"
    return f"{y}年{d.month}月{d.day}日"


WD = "月火水木金土日"
