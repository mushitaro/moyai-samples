# -*- coding: utf-8 -*-
"""openpyxl の小さな道具（罫線・印刷設定・文書のプロパティ）"""
import os, datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter as CL

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.abspath(os.path.join(HERE, "..", "files"))
BUILD = os.path.abspath(os.path.join(HERE, "..", "build"))

# ---- 正解の記録 ----
REC = {}      # stem -> {"file", "plan", "records": []}
LINKS = []    # (file, sheet, r/seq, wid)
CUR = {"stem": None}
def rec_open(fname, plan=False):
    stem = os.path.splitext(fname)[0]
    REC[stem] = {"file": fname, "plan": plan, "records": []}
    CUR["stem"] = stem
    return stem
def rec(sheet, r, kind, fields, works=(), seq=None, image=None):
    st = REC[CUR["stem"]]
    d = {"sheet": sheet}
    if seq is not None: d = {"seq": seq}
    else: d["r"] = r
    d["kind"] = kind
    if image: d["image"] = image
    d["fields"] = {k: v for k, v in fields.items() if v not in (None, "", [], {})}
    st["records"].append(d)
    for w in works:
        LINKS.append({"file": st["file"], "sheet": sheet, ("seq" if seq is not None else "r"): (seq if seq is not None else r), "work": w.wid})
    return d

thin = Side(style="thin"); hair = Side(style="hair"); med = Side(style="medium"); dot = Side(style="dotted")
B_ALL = Border(left=thin, right=thin, top=thin, bottom=thin)
B_HAIR = Border(left=thin, right=thin, top=hair, bottom=hair)

def F(size=10, bold=False, color=None, name="IPAGothic", italic=False, u=None):
    return Font(name=name, size=size, bold=bold, color=color, italic=italic, underline=u)

def fill(hexrgb):
    return PatternFill("solid", fgColor=hexrgb)

AC = Alignment(horizontal="center", vertical="center", wrap_text=True)
AL = Alignment(horizontal="left", vertical="center", wrap_text=True)
ALT = Alignment(horizontal="left", vertical="top", wrap_text=True)
AR = Alignment(horizontal="right", vertical="center")

def new_wb(title, author, created=None, modified_by=None):
    wb = Workbook()
    wb.properties.title = title
    wb.properties.creator = author
    wb.properties.lastModifiedBy = modified_by or author
    wb.properties.subject = "【架空】試験用サンプル"
    if created:
        wb.properties.created = created
        wb.properties.modified = created
    return wb

def put(ws, r, c, v, font=None, al=None, border=None, fl=None, fmt=None):
    cell = ws.cell(row=r, column=c, value=v)
    if font: cell.font = font
    if al: cell.alignment = al
    if border: cell.border = border
    if fl: cell.fill = fl
    if fmt: cell.number_format = fmt
    return cell

def box(ws, r1, c1, r2, c2, border=B_ALL):
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            ws.cell(row=r, column=c).border = border

def merge(ws, r1, c1, r2, c2, v=None, font=None, al=AC, border=B_ALL, fl=None):
    if v is not None or font:
        put(ws, r1, c1, v, font, al, None, fl)
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            if border: ws.cell(row=r, column=c).border = border
            if fl: ws.cell(row=r, column=c).fill = fl
    if (r1, c1) != (r2, c2):
        ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
    ws.cell(row=r1, column=c1).alignment = al

def widths(ws, ws_list):
    for i, w in enumerate(ws_list, 1):
        ws.column_dimensions[CL(i)].width = w

def page(ws, landscape=True, paper=9, fit_w=1, fit_h=0, area=None, titles=None, margins=(0.4, 0.4, 0.5, 0.5), center=True, scale=None):
    """paper: 9=A4 8=A3"""
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.page_setup.paperSize = paper
    if scale:
        ws.page_setup.scale = scale
        ws.sheet_properties.pageSetUpPr.fitToPage = False
    else:
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = fit_w
        ws.page_setup.fitToHeight = fit_h
    l, r, t, b = margins
    ws.page_margins.left, ws.page_margins.right, ws.page_margins.top, ws.page_margins.bottom = l, r, t, b
    ws.page_margins.header = 0.2; ws.page_margins.footer = 0.2
    if area: ws.print_area = area
    if titles: ws.print_title_rows = titles
    ws.print_options.horizontalCentered = center

def comment(ws, r, c, text, author):
    cm = Comment(text, author)
    cm.width = 220; cm.height = 80
    ws.cell(row=r, column=c).comment = cm

def save(wb, fname):
    os.makedirs(OUT, exist_ok=True)
    p = os.path.join(OUT, fname)
    wb.save(p)
    return p

EQLINKS = []  # (file, sheet, r, [mid])
def eqlink(sheet, r, mids):
    EQLINKS.append({"file": REC[CUR["stem"]]["file"], "sheet": sheet, "r": r, "mids": list(mids)})
