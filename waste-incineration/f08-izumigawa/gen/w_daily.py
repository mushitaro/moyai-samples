# -*- coding: utf-8 -*-
"""作業日報（全日分）。
m1a（メーカー、2005–2009）：1 日 1 枚・横型 A4。左に作業内容（区域・作業・協力会社・人数）、右に協力会社別の出面表、下に記事。
m1b（メーカー、2010–2018）：1 日 1 枚・縦型 A4（見本の様式）。作業行ごとに職種別の人数、下に出面計・明日の予定・特記事項。
g1（地元業者、2019–）：1 週 1 枚・横型、1 日 1 行（見本の様式）。"""
import datetime as dt
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from openpyxl import Workbook  # noqa: E402
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side  # noqa: E402
from common import FAC_K, FONT, LOCAL, MAKER, set_props  # noqa: E402
from rec import Recs, fname, path, short  # noqa: E402

thin = Side(style="thin")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
F = Font(name=FONT, size=10)
FS = Font(name=FONT, size=9)
FB = Font(name=FONT, size=10, bold=True)
CE = Alignment(horizontal="center", vertical="center", wrap_text=True)
LE = Alignment(horizontal="left", vertical="center", wrap_text=True)
LT = Alignment(horizontal="left", vertical="top", wrap_text=True)
GRAY = PatternFill("solid", fgColor="E7E6E6")
WD = "月火水木金土日"
COMPANY = {"機械": "みのり工業", "製缶": "みのり工業", "築炉": "さつき築炉", "電気": "あさひ電設", "足場": "かえで足場",
           "作業員": "ひばり興業"}
PRIMARY = [("足場", "足場"), ("ボイラ", "製缶"), ("スートブロワ", "製缶"), ("エコノマイザ", "製缶"),
           ("耐火物", "築炉"), ("電源断", "電気"), ("盤", "電気")]


def box(ws, r1, c1, r2, c2, value=None, font=F, align=CE, fill=None):
    if (r1, c1) != (r2, c2):
        ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
    c = ws.cell(row=r1, column=c1, value=value)
    c.font, c.alignment = font, align
    for rr in range(r1, r2 + 1):
        for cc in range(c1, c2 + 1):
            ws.cell(row=rr, column=cc).border = BOX
            if fill:
                ws.cell(row=rr, column=cc).fill = fill
    return c


def wa_date(d):
    return f"平成{d.year - 1988}年{d.month}月{d.day}日（{WD[d.weekday()]}）"


def trades(it, crew, sample):
    ln, eq, txt, n = it[0], it[1], it[2], it[3]
    if sample:
        key = eq + txt
        main = next((t for k, t in PRIMARY if k in key and crew.get(t)), "機械")
    else:
        main = it[5] if crew.get(it[5]) else "機械"
    a = min(n, max(1, crew.get(main, 0)))
    rest = n - a
    parts = [f"{main}{a}"]
    if rest:
        parts.append(f"{'作業員' if crew.get('作業員') else '機械'}{rest}")
    return "・".join(parts)


def _a4(ws, last, portrait):
    ws.page_setup.orientation = "portrait" if portrait else "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_area = last


# ---------------------------------------------------------------- m1a（2005–2009）
def m1a_day(ws, job, d, v, R):
    for col, w in zip("ABCDEFGHIJ", (8, 40, 14, 6, 3, 22, 7, 3, 10, 10)):
        ws.column_dimensions[col].width = w
    box(ws, 1, 1, 1, 10, "【架空】作　業　日　報", Font(name=FONT, size=16, bold=True))
    box(ws, 2, 1, 2, 1, "件名", FB, fill=GRAY)
    box(ws, 2, 2, 2, 4, job["contract"]["name"].replace("いずみ川清掃センター ", ""), F, LE)
    box(ws, 2, 6, 2, 6, wa_date(d), F)
    box(ws, 2, 7, 2, 7, "天候", FB, fill=GRAY)
    box(ws, 2, 9, 2, 10, v["weather"], F)
    for c, h in ((1, "区域"), (2, "作業内容"), (3, "協力会社"), (4, "人数")):
        box(ws, 4, c, 4, c, h, FS, fill=GRAY)
    r = 5
    for it in v["items"]:
        main = it[5] if it[5] in COMPANY else "機械"
        vals = [it[0], it[2] if it[2].startswith(it[1]) else f"{it[1]}：{it[2]}", COMPANY[main], it[3]]
        for c, val in enumerate(vals, 1):
            box(ws, r, c, r, c, val, F, LE if c == 2 else CE)
        ws.row_dimensions[r].height = 28
        R.add(ws.title, r, "item", {"date": d, "weather": v["weather"], "area": it[0], "work": vals[1],
                                    "company": vals[2], "people": it[3]}, it[7])
        r += 1
    while r < 14:
        for c in range(1, 5):
            box(ws, r, c, r, c, None)
        ws.row_dimensions[r].height = 28
        r += 1
    # 右：協力会社別 出面
    box(ws, 4, 6, 4, 7, "協力会社別 出面", FS, fill=GRAY)
    box(ws, 4, 9, 4, 9, "人数", FS, fill=GRAY)
    box(ws, 4, 10, 4, 10, "時間", FS, fill=GRAY)
    comp = {}
    for k, n in v["crew"].items():
        nm = COMPANY.get(k, "みのり工業")
        comp[nm] = comp.get(nm, 0) + n
    rr = 5
    box(ws, rr, 6, rr, 7, "鷹羽（直）", F, LE)
    box(ws, rr, 9, rr, 9, 2, F)
    box(ws, rr, 10, rr, 10, "8:00-17:00", FS)
    rr += 1
    for nm, n in comp.items():
        box(ws, rr, 6, rr, 7, nm, F, LE)
        box(ws, rr, 9, rr, 9, n, F)
        box(ws, rr, 10, rr, 10, "8:00-17:00" if d.weekday() != 5 else "8:00-12:00", FS)
        rr += 1
    box(ws, rr, 6, rr, 7, "計", FB)
    box(ws, rr, 9, rr, 9, 2 + sum(comp.values()), FB)
    box(ws, 15, 1, 15, 1, "記事", FB, fill=GRAY)
    box(ws, 15, 2, 16, 10, (v.get("note") or "") + ("" if not v.get("note") else "。") + f"翌日：{v.get('tomorrow', '')}", F, LT)
    box(ws, 18, 8, 18, 9, "代理人", FS, fill=GRAY)
    box(ws, 18, 10, 18, 10, "市", FS, fill=GRAY)
    box(ws, 19, 8, 19, 9, "m1", F)
    box(ws, 19, 10, 19, 10, "i1", F)
    ws.cell(row=20, column=1, value=f"{MAKER}　いずみ川作業所").font = FS
    _a4(ws, "A1:J20", False)


# ---------------------------------------------------------------- m1b（2010–2018）
def m1b_day(ws, job, d, v, R):
    for col, w in zip("ABCDEFGHIJ", (5, 9, 14, 12, 12, 12, 8, 8, 8, 8)):
        ws.column_dimensions[col].width = w
    areas = sorted({it[0] for it in v["items"]}, key=lambda x: ["1号炉", "2号炉", "共通"].index(x) if x in ("1号炉", "2号炉", "共通") else 9)
    box(ws, 1, 1, 1, 10, "【架空】作 業 日 報", Font(name=FONT, size=16, bold=True))
    box(ws, 2, 1, 2, 2, "工事件名", FB, fill=GRAY)
    box(ws, 2, 3, 2, 10, f"{FAC_K}　{job['contract']['name'].replace('いずみ川清掃センター ', '')}（{'・'.join(areas)}）", F, LE)
    box(ws, 3, 1, 3, 2, "日付", FB, fill=GRAY)
    box(ws, 3, 3, 3, 5, f"平成{d.year - 1988}年{d.month}月{d.day}日（{WD[d.weekday()]}）", F)
    box(ws, 3, 6, 3, 6, "天候", FB, fill=GRAY)
    box(ws, 3, 7, 3, 7, v["weather"], F)
    box(ws, 3, 8, 3, 8, "作業時間", FB, fill=GRAY)
    box(ws, 3, 9, 3, 10, "8:00〜17:00" if d.weekday() != 5 else "8:00〜12:00", F)
    box(ws, 5, 1, 5, 10, "作 業 内 容", FB, fill=GRAY)
    for c1, c2, h in ((1, 1, "No"), (2, 2, "号炉"), (3, 3, "設備・機器"), (4, 6, "作業内容"),
                      (7, 8, "人員（職種）"), (9, 9, "進捗"), (10, 10, "備考")):
        box(ws, 6, c1, 6, c2, h, FS, fill=GRAY)
    r = 7
    for k, it in enumerate(v["items"], 1):
        tr = trades(it, v["crew"], v.get("sample"))
        box(ws, r, 1, r, 1, k, F)
        box(ws, r, 2, r, 2, it[0], F)
        box(ws, r, 3, r, 3, it[1], F, LE)
        box(ws, r, 4, r, 6, it[2], F, LE)
        box(ws, r, 7, r, 8, tr, FS, LE)
        box(ws, r, 9, r, 9, it[4], F)
        box(ws, r, 10, r, 10, None, F)
        ws.row_dimensions[r].height = 30
        R.add(ws.title, r, "item", {"date": d, "weather": v["weather"], "no": k, "area": it[0], "equipment": it[1],
                                    "work": it[2], "crewText": tr, "progress": it[4]}, it[7])
        r += 1
    while r < 14:
        for c1, c2 in ((1, 1), (2, 2), (3, 3), (4, 6), (7, 8), (9, 9), (10, 10)):
            box(ws, r, c1, r, c2, None)
        ws.row_dimensions[r].height = 30
        r += 1
    tot = sum(v["crew"].values())
    hrs = 4 if d.weekday() == 5 else 8
    box(ws, r, 1, r, 3, "出面 計", FB, fill=GRAY)
    box(ws, r, 4, r, 10, f"{tot} 名（" + "・".join(f"{k}{n}" for k, n in v["crew"].items() if n) +
        f"）　延べ {tot * hrs} h　※掛け持ちがあるため各行の計とは一致しない", F, LE)
    r += 1
    box(ws, r, 1, r, 2, "明日の予定", FB, fill=GRAY)
    box(ws, r, 3, r, 10, v.get("tomorrow", ""), F, LE)
    ws.row_dimensions[r].height = 30
    r += 1
    box(ws, r, 1, r + 1, 2, "特記事項\n（発見・指示）", FB, fill=GRAY)
    box(ws, r, 3, r + 1, 10, v.get("note") or "特になし", F, LT)
    r += 2
    box(ws, r, 1, r, 2, "安全", FB, fill=GRAY)
    box(ws, r, 3, r, 10, "KY実施　朝礼 8:00　災害 0 件　酸欠・一酸化炭素 測定実施（入炉前）", F, LE)
    r += 2
    box(ws, r, 6, r, 7, "現場代理人", FS, fill=GRAY)
    box(ws, r, 8, r, 8, "市確認", FS, fill=GRAY)
    box(ws, r, 9, r, 10, "受付", FS, fill=GRAY)
    box(ws, r + 1, 6, r + 2, 7, "m1", F)
    box(ws, r + 1, 8, r + 2, 8, "i1" if d.year < 2014 or (d.year == 2014 and d.month < 4) else "i2", F)
    box(ws, r + 1, 9, r + 2, 10, None, F)
    ws.cell(row=r + 4, column=1, value=MAKER).font = FS
    _a4(ws, f"A1:J{r + 4}", True)


# ---------------------------------------------------------------- g1
def g1_week(ws, job, rows, wk, R, cum0, use_count):
    widths = [11, 5, 52, 14, 7.5, 12, 7, 8, 8]
    for k, w in enumerate(widths):
        ws.column_dimensions[chr(ord("A") + k)].width = w
    box(ws, 1, 1, 1, 9, f"【架空】作業日報（週報）　第{wk}週", Font(name=FONT, size=14, bold=True))
    box(ws, 2, 1, 2, 1, "工事名", FB, fill=GRAY)
    box(ws, 2, 2, 2, 5, f"{FAC_K}　{job['contract']['name'].replace('いずみ川清掃センター ', '')}", F, LE)
    box(ws, 2, 6, 2, 6, "請負者", FB, fill=GRAY)
    box(ws, 2, 7, 2, 9, f"{LOCAL}　職長 g1", F, LE)
    heads = ["月日", "天候", "作業内容", "場所", "人員", "時間", "延べ(h)"] + (["取替本数", "累計"] if use_count else ["", ""])
    for j, h in enumerate(heads, 1):
        box(ws, 4, j, 4, j, h or None, FB, fill=GRAY)
    cum = cum0
    prev_place = None
    r = 5
    for (d, w, n, h, txt, place, noz, wids, area) in rows:
        cum += noz or 0
        vals = [f"{d.month}/{d.day}（{WD[d.weekday()]}）", w, txt, "〃" if place == prev_place else place, f"{n}名",
                "8:00~17:00" if h == 8 else "8:00~18:30", n * h]
        if use_count:
            vals += [f"{noz}本" if noz else "—", cum if cum else "—"]
        else:
            vals += [None, None]
        prev_place = place
        for j, val in enumerate(vals, 1):
            box(ws, r, j, r, j, val, F, LE if j == 3 else CE)
        ws.row_dimensions[r].height = 34
        f = {"date": d, "weather": w, "work": txt, "place": vals[3], "people": n, "hours": vals[5], "manHours": n * h,
             "manDays": n, "manDaysText": f"{n}名"}
        if use_count:
            f.update({"count": noz, "cumulative": cum})
        R.add(ws.title, r, "item", f, wids)
        r += 1
    while r < 11:
        for j in range(1, 10):
            box(ws, r, j, r, j, None)
        ws.row_dimensions[r].height = 34
        r += 1
    tot = sum(x[2] * x[3] for x in rows)
    box(ws, r, 1, r, 4, "週　計", FB, fill=GRAY)
    box(ws, r, 5, r, 5, f"{sum(x[2] for x in rows)}人日", F)
    box(ws, r, 6, r, 6, None)
    box(ws, r, 7, r, 7, f"{tot:g}", FB)
    if use_count:
        box(ws, r, 8, r, 8, f"{sum((x[6] or 0) for x in rows)}本", F)
        box(ws, r, 9, r, 9, cum, FB)
    else:
        box(ws, r, 8, r, 8, None)
        box(ws, r, 9, r, 9, None)
    R.add(ws.title, r, "subtotal", {"manDays": sum(x[2] for x in rows), "manDaysText": f"{sum(x[2] for x in rows)}人日",
                                    "manHours": tot})
    r += 1
    box(ws, r, 1, r + 1, 1, "記事", FB, fill=GRAY)
    notes = [x for x in rows if "追加" in x[4] or "合格" in x[4] or "漏れなし" in x[4]]
    note = "。".join(f"{x[0].month}/{x[0].day} {x[4][:40]}" for x in notes[:2]) or "特になし"
    if job["cid"] == "L20230":
        note = ("10/18 風箱内に砂たまり有（ヘッダ継手から漏れ）。市 i2 と協議、予備日 10/25 に清掃することで了承"
                if wk == 1 else "10/25 風箱清掃・継手パッキン交換 2 箇所（追加）。10/27 通風試験 良好、2号炉 引渡し 15:30")
    box(ws, r, 2, r + 1, 9, note, F, LT)
    r += 3
    box(ws, r, 8, r, 8, "職長", FS, fill=GRAY)
    box(ws, r, 9, r, 9, "市 確認", FS, fill=GRAY)
    box(ws, r + 1, 8, r + 2, 8, "g1", F)
    box(ws, r + 1, 9, r + 2, 9, "i2", F)
    _a4(ws, f"A1:I{r + 2}", False)
    return cum


def write(job):
    c = job["contract"]
    if job["vendor"] == "g1":
        form = "g1様式（地元業者・1週1枚）"
    elif c["fy"] <= 2009:
        form = "m1様式A（メーカー・H17〜H21・横型）"
    else:
        form = "m1様式B（メーカー・H22〜・縦型）"
    fn = fname("作業日報", short(c["name"]), form, "xlsx")
    R = Recs(fn)
    wb = Workbook()
    first = True
    if job["vendor"] == "g1":
        rows = job["daily"]
        weeks = {}
        for row in rows:
            mon = row[0] - dt.timedelta(days=row[0].weekday())
            weeks.setdefault(mon, []).append(row)
        cum = 0
        for k, (mon, rs) in enumerate(sorted(weeks.items()), 1):
            ws = wb.active if first else wb.create_sheet()
            first = False
            ws.title = f"{rs[0][0].month}.{rs[0][0].day}-{rs[-1][0].month}.{rs[-1][0].day}"
            cum = g1_week(ws, job, rs, k, R, cum, job["cid"] == "L20230")
    else:
        for d in sorted(job["days"]):
            ws = wb.active if first else wb.create_sheet()
            first = False
            ws.title = f"{d.year % 100:02d}.{d.month}.{d.day}"
            (m1a_day if c["fy"] <= 2009 else m1b_day)(ws, job, d, job["days"][d], R)
    set_props(wb, f"{FAC_K} {c['name']} 作業日報")
    wb.properties.creator = (f"{LOCAL}（架空） g1" if job["vendor"] == "g1" else f"{MAKER}（架空） m1")
    wb.save(path(fn))
    R.save()
    return fn
