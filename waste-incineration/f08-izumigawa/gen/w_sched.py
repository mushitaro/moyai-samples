# -*- coding: utf-8 -*-
"""工程表。m1（2005–2018、メーカーの定期整備・部分改良・補修請負）：区域ごとに 1 シート、日単位、予定（青）・実績（赤）の 2 段、班・人員。
g1（2019–）：m1 の表をコピーして簡素化。1 作業 1 行の塗り、実績欄に日付、「班」列は残るが空。区域（号炉の停止）ごとに 1 シート。"""
import datetime as dt
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from openpyxl import Workbook  # noqa: E402
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side  # noqa: E402
from openpyxl.utils import get_column_letter  # noqa: E402
from common import FAC_K, FONT, LOCAL, MAKER, set_props, wareki_year  # noqa: E402
from model import is_off  # noqa: E402
from rec import Recs, fname, path, short  # noqa: E402

thin, hair = Side(style="thin"), Side(style="hair")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
F = Font(name=FONT, size=9)
FB = Font(name=FONT, size=9, bold=True)
CE = Alignment(horizontal="center", vertical="center", wrap_text=True)
LE = Alignment(horizontal="left", vertical="center", wrap_text=True)
PLAN = PatternFill("solid", fgColor="9BC2E6")
ACT = PatternFill("solid", fgColor="F4B183")
SUN = PatternFill("solid", fgColor="F2F2F2")
GRAY = PatternFill("solid", fgColor="D9D9D9")
WD = "月火水木金土日"
CREWN = {"機械": 8, "製缶": 4, "築炉": 4, "電気": 2, "足場": 4, "運転": 0}


def days(a, b):
    d = a
    while d <= b:
        yield d
        d += dt.timedelta(days=1)


def _setup(ws, last_col, last_row, a4=False):
    ws.print_title_rows = "5:6"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4 if a4 else ws.PAPERSIZE_A3
    ws.page_setup.fitToWidth = ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_area = f"A1:{last_col}{last_row}"


def m1_sheet(ws, job, area, R):
    c = job["contract"]
    ws["A1"] = f"{FAC_K}　{c['name'].replace('いずみ川清掃センター ', '')}"
    ws["A1"].font = Font(name=FONT, size=12, bold=True)
    ws["A2"] = f"{area['label']}　工 程 表"
    ws["A2"].font = Font(name=FONT, size=14, bold=True)
    s0, s1 = area["span"]
    ws["A3"] = f"{s0.year}年{s0.month}月{s0.day}日〜{s1.month}月{s1.day}日　（実績記入）"
    ws["A3"].font = F
    dl = list(days(s0, s1))
    c0 = 6
    for j, h in enumerate(["No", "作業項目", "班", "人員", "区分"], 1):
        x = ws.cell(row=5, column=j, value=h)
        ws.merge_cells(start_row=5, start_column=j, end_row=6, end_column=j)
        x.font, x.alignment, x.border, x.fill = FB, CE, BOX, GRAY
    for k, d in enumerate(dl):
        x = ws.cell(row=5, column=c0 + k, value=f"{d.month}/{d.day}" if (d.day == 1 or k == 0) else d.day)
        x.font, x.alignment, x.border = F, CE, BOX
        w = ws.cell(row=6, column=c0 + k, value=WD[d.weekday()])
        w.font, w.alignment, w.border = F, CE, BOX
        if is_off(d):
            x.fill = w.fill = SUN
        ws.column_dimensions[get_column_letter(c0 + k)].width = 3.6
    for col, wdt in zip("ABCDE", (4, 36, 5, 5, 5)):
        ws.column_dimensions[col].width = wdt
    r = 7
    for t in area["tasks"]:
        ws.cell(row=r, column=1, value=t["no"])
        ws.cell(row=r, column=2, value=t["name"])
        ws.cell(row=r, column=3, value=t["han"])
        ws.cell(row=r, column=4, value=CREWN.get(t["han"]) or None)
        ws.cell(row=r, column=5, value="予定")
        ws.cell(row=r + 1, column=5, value="実績")
        for j in (1, 2, 3, 4):
            ws.merge_cells(start_row=r, start_column=j, end_row=r + 1, end_column=j)
        for k, d in enumerate(dl):
            for rr, (a, b, fill) in ((r, (t["ps"], t["pe"], PLAN)), (r + 1, (t["as_"], t["ae"], ACT))):
                x = ws.cell(row=rr, column=c0 + k)
                x.border = Border(left=hair, right=hair, top=thin if rr == r else hair, bottom=hair if rr == r else thin)
                if is_off(d):
                    x.fill = SUN
                if a and a <= d <= b and not (is_off(d) and t["han"] != "運転"):
                    x.fill = fill
        for rr in (r, r + 1):
            for j in range(1, 6):
                x = ws.cell(row=rr, column=j)
                x.font, x.border = F, BOX
                x.alignment = LE if j == 2 else CE
        R.add(ws.title, r, "item", {"no": t["no"], "name": t["name"], "team": t["han"], "crew": CREWN.get(t["han"]) or None,
                                    "planStart": t["ps"], "planEnd": t["pe"], "actualStart": t["as_"], "actualEnd": t["ae"],
                                    "area": area["label"]}, [w["wid"] for w in t["works"]])
        r += 2
    ws.cell(row=r + 1, column=2, value="凡例　■ 予定（青）　■ 実績（赤）　網掛け：日曜・祝日（休工）").font = F
    ws.cell(row=r + 2, column=2, value=f"{MAKER}　現場代理人 m1").font = F
    ws.freeze_panes = "F7"
    _setup(ws, get_column_letter(c0 + len(dl) - 1), r + 2)


def g1_sheet(ws, job, area, R):
    c = job["contract"]
    ws["A1"] = f"{FAC_K}　{c['name'].replace('いずみ川清掃センター ', '')}"
    ws["A1"].font = Font(name=FONT, size=12, bold=True)
    ws["A2"] = "工 程 表" + (f"　（{area['label']}）" if len(job["areas"]) > 1 else "")
    ws["A2"].font = Font(name=FONT, size=14, bold=True)
    s0, s1 = area["span"]
    ws["A3"] = (f"工期　{_wa(job['cdate'])}〜{_wa(job['climit'])}　（現場作業 {s0.month}/{s0.day}〜{s1.month}/{s1.day}）")
    ws["A3"].font = F
    dl = list(days(s0, s1))
    c0 = 5
    for j, h in enumerate(["No", "作業項目", "班", "実績"], 1):
        x = ws.cell(row=5, column=j, value=h)
        ws.merge_cells(start_row=5, start_column=j, end_row=6, end_column=j)
        x.font, x.alignment, x.border, x.fill = FB, CE, BOX, GRAY
    for k, d in enumerate(dl):
        x = ws.cell(row=5, column=c0 + k, value=f"{d.month}/{d.day}" if k == 0 or d.day == 1 else d.day)
        x.font, x.alignment, x.border = F, CE, BOX
        w = ws.cell(row=6, column=c0 + k, value=WD[d.weekday()])
        w.font, w.alignment, w.border = F, CE, BOX
        if is_off(d):
            x.fill = w.fill = SUN
        ws.column_dimensions[get_column_letter(c0 + k)].width = 4.2
    for col, wdt in zip("ABCD", (4, 34, 5, 22)):
        ws.column_dimensions[col].width = wdt
    r = 7
    for k, t in enumerate(area["tasks"], 1):
        ws.cell(row=r, column=1, value=k)
        ws.cell(row=r, column=2, value=t["name"])
        if t["as_"]:
            txt = f"{t['as_'].month}/{t['as_'].day}" if t["as_"] == t["ae"] else \
                f"{t['as_'].month}/{t['as_'].day}〜{t['ae'].month}/{t['ae'].day}"
            note = t.get("note") or ""
            ws.cell(row=r, column=4, value=("済 " + txt + (" " + note if note else "")).strip())
        elif t["ps"]:
            ws.cell(row=r, column=4, value="使わず")
        for kk, d in enumerate(dl):
            x = ws.cell(row=r, column=c0 + kk)
            x.border = BOX
            if is_off(d):
                x.fill = SUN
            a, b = (t["ps"], t["pe"]) if t["ps"] else (t["as_"], t["ae"])
            if a and a <= d <= b and not (is_off(d) and "運転" not in t["name"]):
                x.fill = PLAN if t["ps"] else ACT
        for j in range(1, 5):
            x = ws.cell(row=r, column=j)
            x.font, x.border = F, BOX
            x.alignment = LE if j in (2, 4) else CE
        ws.row_dimensions[r].height = 20
        R.add(ws.title, r, "item", {"no": k, "name": t["name"], "planStart": t["ps"], "planEnd": t["pe"],
                                    "actualText": ws.cell(row=r, column=4).value, "area": area["label"]},
              [w["wid"] for w in t["works"]])
        r += 1
    ws.cell(row=r + 1, column=2, value="凡例　━ 予定　━ 実績").font = F
    ws.cell(row=r + 2, column=2, value="塗り：予定（追加作業は赤）。実績は右欄に記入").font = F
    ws.cell(row=r + 3, column=2, value=f"{LOCAL}　職長 g1").font = F
    _setup(ws, get_column_letter(c0 + len(dl) - 1), r + 3, a4=True)


def _wa(d):
    if d >= dt.date(2019, 5, 1):
        n = d.year - 2018
        return f"令和{'元' if n == 1 else n}年{d.month}月{d.day}日"
    return f"平成{d.year - 1988}年{d.month}月{d.day}日"


def write(job):
    c = job["contract"]
    g1 = job["vendor"] == "g1"
    fn = fname("工程表", short(c["name"]), "g1様式（地元業者・簡素化）" if g1 else "m1様式（メーカー・区域別日単位）", "xlsx")
    R = Recs(fn)
    wb = Workbook()
    for k, a in enumerate(job["areas"]):
        ws = wb.active if k == 0 else wb.create_sheet()
        ws.title = a["label"][:20]
        (g1_sheet if g1 else m1_sheet)(ws, job, a, R)
    set_props(wb, f"{FAC_K} {c['name']} 工程表")
    wb.properties.creator = (f"{LOCAL}（架空） g1" if g1 else f"{MAKER}（架空） m1")
    wb.save(path(fn))
    R.save()
    return fn
