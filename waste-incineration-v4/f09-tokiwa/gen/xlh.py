# -*- coding: utf-8 -*-
"""openpyxl の小道具"""
import datetime as dt
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter as L

FONT = "IPAGothic"
PFONT = "IPAPGothic"
thin = Side(style="thin", color="000000")
hair = Side(style="hair", color="000000")
med = Side(style="medium", color="000000")
dot = Side(style="dotted", color="000000")
B_ALL = Border(left=thin, right=thin, top=thin, bottom=thin)
B_HAIR = Border(left=thin, right=thin, top=hair, bottom=hair)

def fill(c):
    return PatternFill("solid", fgColor=c)

def f(size=9, bold=False, name=FONT, color="000000", italic=False):
    return Font(name=name, size=size, bold=bold, color=color, italic=italic)

def al(h="left", v="center", wrap=False, rot=0, shrink=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap, text_rotation=rot, shrink_to_fit=shrink)

def put(ws, r, c, v, font=None, align=None, border=None, fl=None, fmt=None):
    cell = ws.cell(row=r, column=c, value=v)
    if font is not None: cell.font = font
    if align is not None: cell.alignment = align
    if border is not None: cell.border = border
    if fl is not None: cell.fill = fl
    if fmt is not None: cell.number_format = fmt
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

def merge(ws, r1, c1, r2, c2, v=None, font=None, align=None, border=B_ALL, fl=None):
    if border is not None:
        box(ws, r1, c1, r2, c2, border)
    if (r1, c1) != (r2, c2):
        ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
    cell = ws.cell(row=r1, column=c1)
    if v is not None: cell.value = v
    if font is not None: cell.font = font
    if align is not None: cell.alignment = align
    if fl is not None:
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                ws.cell(row=r, column=c).fill = fl
    return cell

def widths(ws, d):
    for k, w in d.items():
        ws.column_dimensions[k if isinstance(k, str) else L(k)].width = w

def page(ws, orient="landscape", paper=8, fit_w=1, fit_h=0, area=None, titles=None, margins=(0.4, 0.4, 0.5, 0.5), center=True):
    """paper: 9=A4, 8=A3"""
    ws.page_setup.orientation = orient
    ws.page_setup.paperSize = paper
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = fit_w
    ws.page_setup.fitToHeight = fit_h
    if area: ws.print_area = area
    if titles: ws.print_title_rows = titles
    ws.page_margins.left, ws.page_margins.right, ws.page_margins.top, ws.page_margins.bottom = margins
    ws.print_options.horizontalCentered = center

def props(wb, title, creator, created=None, modified_by=None):
    wb.properties.title = title
    wb.properties.creator = creator
    wb.properties.lastModifiedBy = modified_by or creator
    d = created or dt.datetime(2026, 11, 13, 9, 0)
    wb.properties.created = d
    wb.properties.modified = d

WAREKI = [(dt.date(2019, 5, 1), "令和", 2018), (dt.date(1989, 1, 8), "平成", 1988)]

def wareki_fy(fy, short=False):
    if fy >= 2019:
        n = fy - 2018
        return (f"R{n}" if short else f"令和{n if n > 1 else '元'}年度")
    n = fy - 1988
    return (f"H{n}" if short else f"平成{n}年度")

def wareki_date(d, style="kanji"):
    for start, era, base in WAREKI:
        if d >= start:
            y = d.year - base
            if style == "kanji":
                return f"{era}{y}年{d.month}月{d.day}日"
            if style == "H":
                return f"{'H' if era == '平成' else 'R'}{y}.{d.month}.{d.day}"
    return d.isoformat()

YOUBI = "月火水木金土日"
