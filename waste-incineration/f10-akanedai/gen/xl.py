# -*- coding: utf-8 -*-
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter as L
import datetime as dt
F = 'IPAGothic'
def font(sz=9, b=False, color=None, u=None):
    return Font(name=F, size=sz, bold=b, color=color, underline=u)
thin = Side(style='thin', color='000000'); hair = Side(style='hair', color='000000'); med = Side(style='medium', color='000000')
dot = Side(style='dotted', color='000000')
def box(t=thin, l=thin, r=thin, b=thin): return Border(top=t, left=l, right=r, bottom=b)
B = box(); NB = Border()
def fill(c): return PatternFill('solid', fgColor=c)
AC = Alignment(horizontal='center', vertical='center', wrap_text=True)
AL = Alignment(horizontal='left', vertical='center', wrap_text=False)
ALW = Alignment(horizontal='left', vertical='center', wrap_text=True)
AR = Alignment(horizontal='right', vertical='center')
def put(ws, r, c, v, f=None, a=None, b=None, fl=None, nf=None):
    x = ws.cell(r, c, v)
    if f: x.font = f
    if a: x.alignment = a
    if b: x.border = b
    if fl: x.fill = fl
    if nf: x.number_format = nf
    return x
def merge(ws, r1, c1, r2, c2, v=None, f=None, a=AC, b=B, fl=None):
    if v is not None or f: put(ws, r1, c1, v, f, a, None, fl)
    if (r1, c1) != (r2, c2): ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
    if b:
        for rr in range(r1, r2 + 1):
            for cc in range(c1, c2 + 1):
                ws.cell(rr, cc).border = b
    if fl:
        for rr in range(r1, r2 + 1):
            for cc in range(c1, c2 + 1):
                ws.cell(rr, cc).fill = fl
def widths(ws, ws_list):
    for i, w in enumerate(ws_list, 1): ws.column_dimensions[L(i)].width = w
def page(ws, land=True, paper=9, fitw=1, fith=0, titles=None, area=None, margins=(0.4, 0.4, 0.5, 0.5), footer=True, header=None):
    ws.page_setup.orientation = 'landscape' if land else 'portrait'
    ws.page_setup.paperSize = paper   # 9=A4 8=A3
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = fitw; ws.page_setup.fitToHeight = fith
    if titles: ws.print_title_rows = titles
    if area: ws.print_area = area
    ws.page_margins.left, ws.page_margins.right, ws.page_margins.top, ws.page_margins.bottom = margins
    if footer: ws.oddFooter.center.text = '&P / &N'
    if header: ws.oddHeader.right.text = header
def props(wb, title, author):
    wb.properties.title = title; wb.properties.creator = author
    wb.properties.created = dt.datetime(2026, 11, 13, 9, 0, 0); wb.properties.modified = dt.datetime(2026, 11, 13, 9, 0, 0)
    wb.properties.lastModifiedBy = author
def strip(s): return (s or '').replace('【架空】', '')
