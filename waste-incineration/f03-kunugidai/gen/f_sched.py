# -*- coding: utf-8 -*-
"""工程表（全年度）。2001〜2014：m1 様式（1 工事 1 枚・予定のみ・線）。2015〜：m2 様式（年度 1 枚に全契約、予定／実績の塗り）。"""
import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Alignment
from openpyxl.comments import Comment
import xlhelp as X
from rec import Rec
from model import workday, agent, wdays
from world import wareki

F = X.font
WD = "月火水木金土日"
KUBUN = {"市": "停止", "鳶": "仮設", "A": "燃焼", "A2": "燃焼", "築炉": "燃焼", "B": "ボイラ", "B2": "ボイラ", "C": "排ガス", "C2": "排ガス",
         "D": "通風", "D2": "通風", "E": "電計", "E2": "電計", "全": "試運転", "H": "補修", "F": "共通", "T": "タービン"}
CREWN = {"A": "機械A", "A2": "機械A", "B": "ボイラ", "B2": "ボイラ", "C": "排ガス", "C2": "排ガス", "D": "機械B", "D2": "機械B",
         "E": "電計", "E2": "電計", "築炉": "築炉", "鳶": "鳶", "市": "市", "全": "全班", "H": "補修班", "F": "共通", "T": "タービン"}


def _dmy(d):
    return d.isoformat()


def m1(c, path, fname):
    rec = Rec(fname)
    wb = Workbook()
    ws = wb.active
    ws.title = "工程表"
    X.props(wb, "【架空】%s 作業工程表" % c.name, "汐見重工環境 %s" % agent(c.fy), "工程表", "m1 様式（1 工事 1 枚・予定）")
    tasks = c.tasks
    a0 = min(t.plan[0] for t in tasks if t.plan) - dt.timedelta(days=2)
    b0 = max(t.plan[1] for t in tasks if t.plan) + dt.timedelta(days=3)
    dl = [a0 + dt.timedelta(days=i) for i in range((b0 - a0).days + 1)]
    c0 = 6
    X.widths(ws, [4, 7, 40, 7, 1.5] + [3.2] * len(dl) + [12])
    X.put(ws, 1, 1, "【架空】作　業　工　程　表（予定）", F(size=16, bold=True), X.CENTER)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=c0 + len(dl))
    X.put(ws, 3, 1, "工事名", F(size=10), X.CENTER)
    ws.merge_cells("A3:B3")
    X.put(ws, 3, 3, c.name, F(size=11, bold=True), X.LEFT)
    X.put(ws, 4, 1, "工　期", F(size=10), X.CENTER)
    ws.merge_cells("A4:B4")
    wy = lambda d: "%s年%d月%d日" % (("平成%d" % (d.year - 1988)) if d.year < 2019 else "令和%d" % (d.year - 2018), d.month, d.day)
    X.put(ws, 4, 3, "%s ～ %s" % (wy(c.start), wy(c.end)), F(size=10), X.LEFT)
    cd = c.cdate - dt.timedelta(days=10)
    for rr, t in ((3, "発注者　くぬぎ台市"), (4, "受注者　汐見重工環境株式会社"), (5, "現場代理人　%s　　H%d.%d.%d 作成" % (agent(c.fy), cd.year - 1988, cd.month, cd.day))):
        X.put(ws, rr, c0 + 18, t, F(size=10), Alignment(horizontal="left", vertical="center", wrap_text=False))
        ws.merge_cells(start_row=rr, start_column=c0 + 18, end_row=rr, end_column=c0 + 30)
    ws.merge_cells(start_row=3, start_column=3, end_row=3, end_column=c0 + 12)
    ws.merge_cells(start_row=4, start_column=3, end_row=4, end_column=c0 + 12)
    X.box(ws, 3, 1, 4, c0 + 12)
    rec.add("工程表", 3, "parent", {"工事名": c.name, "工期開始": _dmy(c.start), "工期終了": _dmy(c.end), "現場代理人": agent(c.fy)})
    hr = 7
    for i, h in enumerate(["No", "区分", "作　業　項　目", "担当"], 1):
        X.put(ws, hr, i, h, F(bold=True), X.CENTER)
        ws.merge_cells(start_row=hr, start_column=i, end_row=hr + 2, end_column=i)
    for j, d in enumerate(dl):
        cc = c0 + j
        if d.day == 1 or j == 0:
            X.put(ws, hr, cc, "%d月" % d.month, F(size=9, bold=True), X.LEFT)
        X.put(ws, hr + 1, cc, d.day, F(size=8), X.CENTER)
        X.put(ws, hr + 2, cc, WD[d.weekday()], F(size=8, color="FF0000" if not workday(d) else None), X.CENTER)
    X.put(ws, hr, c0 + len(dl), "備　考", F(bold=True), X.CENTER)
    ws.merge_cells(start_row=hr, start_column=c0 + len(dl), end_row=hr + 2, end_column=c0 + len(dl))
    r = hr + 3
    prev = None
    n = 0
    for t in tasks:
        if not t.plan:
            continue    # m1 様式は予定だけ（追加作業は載らない）
        n += 1
        kub = getattr(t, "kubun", None) or KUBUN.get(t.crew, "")
        if t.crew == "市" and n > 2:
            kub = "復旧"
        X.put(ws, r, 1, n, F(size=9), X.CENTER)
        X.put(ws, r, 2, kub if kub != prev else "", F(size=9), X.CENTER)
        prev = kub
        X.put(ws, r, 3, t.name, F(size=9), X.LEFT)
        crew = t.crew if c.sample else CREWN.get(t.crew, t.crew)
        X.put(ws, r, 4, crew, F(size=8), X.CENTER)
        a, b = t.plan
        for j, d in enumerate(dl):
            cc = c0 + j
            if not workday(d):
                ws.cell(r, cc).fill = X.fill("F2F2F2")
            if a <= d <= b:
                ch = "━"
                if d == a:
                    ch = "┣"
                if d == b:
                    ch = "┫" if a != b else "◆"
                if not workday(d) and t.crew != "市":
                    ch = "┄"
                X.put(ws, r, cc, ch, F(size=10), X.CENTER)
        note = (t.note or "").replace("◇", "")
        if "水圧" in t.name:
            note = note or "市立会"
            X.put(ws, r, c0 + (b - dl[0]).days, "◇", F(size=10, color="C00000"), X.CENTER)
        X.put(ws, r, c0 + len(dl), note, F(size=8), X.LEFT)
        ws.row_dimensions[r].height = 18
        rec.add("工程表", r, "item", {"No": n, "区分": ws.cell(r, 2).value, "作業項目": t.name, "担当": crew,
                                    "予定開始": _dmy(a), "予定終了": _dmy(b), "備考": note}, works=t.works)
        r += 1
    while r < hr + 3 + 25:
        r += 1
    X.box(ws, hr, 1, r - 1, c0 + len(dl), side=X.med, inner=X.thin, inner_h=X.dot)
    X.box(ws, hr, 1, hr + 2, c0 + len(dl), side=X.med, inner=X.thin)
    X.put(ws, r + 1, 1, "注）日曜日・祝日は休工。┄ は休工日をまたぐ作業。炉の立下げ・立上げは市の運転による。", F(size=9), X.LEFT)
    ws.merge_cells(start_row=r + 1, start_column=1, end_row=r + 1, end_column=c0 + 20)
    X.page(ws, "landscape", "A3", 1, 1)
    wb.save(path)
    rec.save()


def m2(fy, contracts, path, fname, model):
    rec = Rec(fname)
    wb = Workbook()
    ws = wb.active
    ws.title = "%s定修工程" % wareki(fy, True)
    X.props(wb, "【架空】%s くぬぎ台環境センター 定期整備工程表" % wareki(fy), "汐見重工環境 m2", "工程表", "m2 様式（H27〜、全契約を 1 枚、予定／実績）")
    cs = [c for c in contracts if c.kind in ("teishu", "kyotsu", "hoshu")]
    a0 = min(min(t.a for t in c.tasks) for c in cs)
    a0 = dt.date(a0.year, a0.month, 1)
    b0 = max(max(t.b for t in c.tasks) for c in cs) + dt.timedelta(days=3)
    dl = [a0 + dt.timedelta(days=i) for i in range((b0 - a0).days + 1)]
    c0 = 8
    X.widths(ws, [9, 4, 34, 6, 4, 0.8, 0.8] + [2.25] * len(dl) + [16])
    ce = c0 + len(dl)
    X.put(ws, 1, 1, "【架空】%s　くぬぎ台環境センター　定期整備 工程表（%s）" % (wareki(fy), "・".join(
        ["%d号炉" % c.unit for c in cs if c.kind == "teishu"] + ["全炉停止時共通設備"] + (["補修工事"] if any(c.kind == "hoshu" for c in cs) else []))),
          F(size=15, bold=True), X.NOWRAP)
    done = "" if fy == 2026 else "　実績：%s 記入" % ("R%d.12.%d" % (fy - 2018, 16) if fy >= 2019 else "H%d.12.16" % (fy - 1988))
    X.put(ws, 2, 1, "汐見重工環境株式会社 くぬぎ台作業所　現場代理人 m2　　予定：%s 作成%s%s" % (
        ("R%d.8.20" % (fy - 2018)) if fy >= 2019 else "H%d.8.20" % (fy - 1988), done,
        "　（R8.10.30 時点：1号炉のみ実績記入）" if fy == 2026 else ""), F(size=10), X.NOWRAP)
    hr = 4
    for i, h in enumerate(["契約", "No", "作業項目", "班", "予/実"], 1):
        X.put(ws, hr, i, h, F(bold=True, size=9), X.CENTER, X.fill("DDEBF7"))
        ws.merge_cells(start_row=hr, start_column=i, end_row=hr + 2, end_column=i)
    for j, d in enumerate(dl):
        cc = c0 + j
        if d.day == 1 or j == 0:
            X.put(ws, hr, cc, "%d月" % d.month, F(size=9, bold=True), X.NOWRAP, X.fill("DDEBF7"))
        else:
            ws.cell(hr, cc).fill = X.fill("DDEBF7")
        X.put(ws, hr + 1, cc, d.day, F(size=7), X.CENTER, X.fill("D9D9D9" if not workday(d) else "DDEBF7"))
        X.put(ws, hr + 2, cc, WD[d.weekday()], F(size=7, color="C00000" if not workday(d) else None), X.CENTER,
              X.fill("D9D9D9" if not workday(d) else "DDEBF7"))
    X.put(ws, hr, ce, "備考", F(bold=True, size=9), X.CENTER, X.fill("DDEBF7"))
    ws.merge_cells(start_row=hr, start_column=ce, end_row=hr + 2, end_column=ce)
    r = hr + 3
    # 運転状態の帯
    spans = {}
    for c in cs:
        if c.kind == "teishu":
            st = [t for t in c.tasks if t.crew == "市"]
            if st:
                spans.setdefault("%d号炉" % c.unit, []).append((st[0].a, st[-1].b))
        if c.kind == "kyotsu" and c.zentei:
            for u in (1, 2, 3):
                spans.setdefault("%d号炉" % u, []).append((c.zentei[0] - dt.timedelta(days=1), c.zentei[1] + dt.timedelta(days=1)))
    X.put(ws, r, 1, "運転状態", F(bold=True, size=9), X.CENTER, X.fill("FFF2CC"))
    ws.merge_cells(start_row=r, start_column=1, end_row=r + 2, end_column=2)
    for k, nm in enumerate(["1号炉", "2号炉", "3号炉"]):
        X.put(ws, r + k, 3, nm + "（停止＝灰）", F(size=8), X.LEFT, X.fill("FFF2CC"))
        ws.merge_cells(start_row=r + k, start_column=3, end_row=r + k, end_column=5)
        for j, d in enumerate(dl):
            if any(a <= d <= b for a, b in spans.get(nm, [])):
                ws.cell(r + k, c0 + j).fill = X.fill("7F7F7F")
        sp = spans.get(nm, [])
        rec.add(ws.title, r + k, "item", {"区分": "運転状態", "号炉": nm, "停止期間": "、".join("%s〜%s" % (_dmy(a), _dmy(b)) for a, b in sp)})
    X.box(ws, r, 1, r + 2, ce, inner=X.hair)
    r += 3
    color = {"teishu": [("DDEBF7", "2F5597"), ("E2EFDA", "548235"), ("EDE1F5", "7030A0")], "kyotsu": ("FCE4D6", "C55A11"), "hoshu": ("FFF2CC", "BF8F00")}
    order = sorted(cs, key=lambda c: (min(t.a for t in c.tasks), c.kind != "teishu"))
    for c in order:
        col = color["teishu"][c.unit - 1] if c.kind == "teishu" else color[c.kind]
        short = c.name.split("くぬぎ台環境センター")[-1].strip()
        X.put(ws, r, 1, "%s　（工期 %d/%d〜%d/%d）" % (short, c.start.month, c.start.day, c.end.month, c.end.day), F(bold=True, size=10), X.NOWRAP, X.fill("F2F2F2"))
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
        for j, d in enumerate(dl):
            ws.cell(r, c0 + j).fill = X.fill("F2F2F2")
            if c.start <= d <= c.end:
                ws.cell(r, c0 + j).border = X.Border(bottom=X.med)
        rec.add(ws.title, r, "parent", {"契約": short, "工期開始": _dmy(c.start), "工期終了": _dmy(c.end)})
        r += 1
        r0 = r
        for i, t in enumerate(c.tasks, 1):
            X.put(ws, r, 2, i, F(size=8), X.CENTER)
            X.put(ws, r, 3, t.name, F(size=8, color="C00000" if t.plan is None else None), X.LEFT)
            X.put(ws, r, 4, t.crew, F(size=8), X.CENTER)
            X.put(ws, r, 5, "予", F(size=7), X.CENTER)
            X.put(ws, r + 1, 5, "実", F(size=7), X.CENTER)
            for k in (2, 3, 4):
                ws.merge_cells(start_row=r, start_column=k, end_row=r + 1, end_column=k)
            planned_only = getattr(c, "planned_only", False)
            ac = None if planned_only else (t.act or t.plan)
            for j, d in enumerate(dl):
                cc = c0 + j
                hol = not workday(d) and t.crew not in ("市", "T")
                if t.plan and t.plan[0] <= d <= t.plan[1]:
                    ws.cell(r, cc).fill = X.fill(col[0] if not hol else "EDEDED")
                if ac and ac[0] <= d <= ac[1]:
                    ws.cell(r + 1, cc).fill = X.fill(("C00000" if t.plan is None else col[1]) if not hol else "A6A6A6")
            rem = ""
            if t.act and t.plan:
                dlt = (t.act[1] - t.plan[1]).days
                rem = t.note or ("完了 %+d 日" % dlt if dlt else "")
            if t.plan is None:
                rem = t.note or "追加・変更契約"
            if t.tags:
                ws.cell(r, 3).comment = Comment("対象：" + "、".join(t.tags[:8]), "m2")
            X.put(ws, r, ce, rem, F(size=7), X.LEFT)
            ws.merge_cells(start_row=r, start_column=ce, end_row=r + 1, end_column=ce)
            ws.row_dimensions[r].height = 10.5
            ws.row_dimensions[r + 1].height = 10.5
            f = {"No": i, "作業項目": t.name, "班": t.crew, "備考": rem}
            if t.plan:
                f.update({"予定開始": _dmy(t.plan[0]), "予定終了": _dmy(t.plan[1])})
            if ac:
                f.update({"実績開始": _dmy(ac[0]), "実績終了": _dmy(ac[1])})
            rec.add(ws.title, r, "item", f, works=t.works, r_end=r + 1)
            r += 2
        lab = {"teishu": "%d号炉" % c.unit, "kyotsu": "共通", "hoshu": "補修"}[c.kind]
        X.put(ws, r0, 1, lab, F(bold=True, size=10), X.CENTER)
        ws.merge_cells(start_row=r0, start_column=1, end_row=r - 1, end_column=1)
        X.box(ws, r0 - 1, 1, r - 1, ce, inner=X.hair)
    X.box(ws, hr, 1, hr + 2, ce, side=X.med, inner=X.thin)
    ky = [c for c in cs if c.kind == "kyotsu" and c.zentei]
    if ky:
        z = ky[0].zentei
        X.put(ws, r + 1, 1, "※ 全炉停止 %d/%d〜%d/%d。" % (z[0].month, z[0].day, z[1].month, z[1].day), F(size=9), X.NOWRAP)
    ws.freeze_panes = "H7"
    ws.print_title_cols = "A:E"
    X.page(ws, "landscape", "A3", 1, 0, area="A1:%s%d" % (X.get_column_letter(ce), r + 2), titles="4:6", margins=(0.3, 0.3, 0.4, 0.4))
    wb.save(path)
    rec.save()
