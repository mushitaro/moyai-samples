# -*- coding: utf-8 -*-
"""工程表（全年度）：sv1 様式 2009–2016／sv2 2017–2023（sv1 のコピー）／2024 様式 2024–／メーカー様式（基幹改良）"""
import datetime as dt
from openpyxl import Workbook
import world as W, works as K
from xlh import *
from rec import Rec

COMP = {W.SUB1: "北嶺", W.REFR: "丸越築炉", W.MAKER: "鷹羽"}

def segments(dates, join_gap=2):
    ds = sorted(set(dates)); segs = []
    for d in ds:
        if segs and (d - segs[-1][1]).days <= join_gap: segs[-1][1] = d
        else: segs.append([d, d])
    return segs

def in_segs(d, segs):
    return any(a <= d <= b for a, b in segs)

def approvals(ws, r, c, labels, names):
    for i, (lb, nm) in enumerate(zip(labels, names)):
        put(ws, r, c + i * 2, lb, f(8), al("center"), B_ALL)
        ws.merge_cells(start_row=r, start_column=c + i * 2, end_row=r, end_column=c + i * 2 + 1)
        box(ws, r, c + i * 2, r, c + i * 2 + 1)
        merge(ws, r + 1, c + i * 2, r + 2, c + i * 2 + 1, nm, f(9), al("center"), B_ALL)

def wid(c, t):
    return [f"{c.cid}:{t.tid}:{x}" for x in t.tags] if t.tags else [f"{c.cid}:{t.tid}"]

def plan_segs(t):
    if t.steps:
        ds = [s[0] for s in t.steps if not s[1].startswith("【追加】")]
        return segments(ds) if ds else [list(t.plan)]
    return [list(t.plan)]

def plan_override(c):
    if c.cid == "T2015-2":
        return {"T04": [(dt.date(2015, 10, 8), dt.date(2015, 10, 14)), (dt.date(2015, 10, 16), dt.date(2015, 10, 17)), (dt.date(2015, 10, 19), dt.date(2015, 10, 19))],
                "T13": [(dt.date(2015, 10, 20), dt.date(2015, 10, 24))], "T00": [(dt.date(2015, 10, 5), dt.date(2015, 10, 7))],
                "T14": [(dt.date(2015, 10, 26), dt.date(2015, 10, 30))]}
    if c.cid == "T2025-1":
        return {"T05": [(dt.date(2025, 5, 19), dt.date(2025, 5, 26))], "T13": [(dt.date(2025, 6, 2), dt.date(2025, 6, 6))],
                "T00": [(dt.date(2025, 5, 12), dt.date(2025, 5, 14))], "T12": [(dt.date(2025, 5, 28), dt.date(2025, 5, 31))],
                "T01": [(dt.date(2025, 5, 15), dt.date(2025, 5, 15)), (dt.date(2025, 5, 30), dt.date(2025, 5, 30))]}
    out = {}
    for t in c.tasks:
        if t.owner == "SPC" or not t.steps:
            out[t.tid] = [tuple(t.plan)]
    return out

def date_cols(c, pad_before=4, pad_after=3):
    a = c.start - dt.timedelta(days=pad_before)
    b = c.end + dt.timedelta(days=pad_after)
    if c.cid == "T2015-2":
        a, b = dt.date(2015, 10, 1), dt.date(2015, 10, 31)
    if c.cid == "T2025-1":
        a, b = dt.date(2025, 5, 8), dt.date(2025, 6, 10)
    return list(W.daterange(a, b))

def month_header(ws, r, dc0, days):
    months = {}
    for i, d in enumerate(days):
        months.setdefault((d.year, d.month), []).append(i)
    first = True
    for (y, m), idx in months.items():
        txt = f"{y}年{m}月" if first else f"{m}月"
        first = False
        merge(ws, r, dc0 + idx[0], r, dc0 + idx[-1], txt, f(9, True), al("center"), B_ALL, fill("D9D9D9"))

# ============================================================== sv1 / sv2（2009–2023）
def sv1(c, path, fname, form="sv1"):
    rec = Rec(fname)
    wb = Workbook(); ws = wb.active; ws.title = "工程表"
    days = date_cols(c)
    dc0 = 5; nc = dc0 + len(days)
    who = "sv1" if form == "sv1" else "sv2"
    ttl = f"{W.FAC}　{c.short}　工程表" if c.kind != "dcs" else f"{W.FAC}　{c.short}　工程表（現地工事）"
    merge(ws, 1, 1, 1, nc - 8, ttl, f(16, True), al("left"), None)
    merge(ws, 2, 1, 2, nc - 8, f"工期：{c.start.year}年{c.start.month}月{c.start.day}日〜{c.end.year}年{c.end.month}月{c.end.day}日　発注：{W.SPC}　施工：{c.contractor}", f(9), al("left"), None)
    approvals(ws, 1, nc - 6, ["承認", "確認", "作成"], ["所長", "GL" if form == "sv1" else "sv2", who])
    put(ws, 4, 1, "No.", f(9, True), al("center"), B_ALL, fill("D9D9D9"))
    for i, h in enumerate(["作業項目", "対象機器", "業者" if form == "sv1" else "施工"]):
        merge(ws, 4, 2 + i, 6, 2 + i, h, f(9, True), al("center"), B_ALL, fill("D9D9D9"))
    ws.merge_cells("A4:A6")
    month_header(ws, 4, dc0, days)
    for i, d in enumerate(days):
        fl = fill("F2DCDB") if d.weekday() == 6 else (fill("EDEDED") if d.weekday() == 5 else fill("D9D9D9"))
        put(ws, 5, dc0 + i, d.day, f(8, True), al("center"), B_ALL, fl)
        put(ws, 6, dc0 + i, YOUBI[d.weekday()], f(8), al("center"), B_ALL, fl)
    merge(ws, 4, nc, 6, nc, "備考", f(9, True), al("center"), B_ALL, fill("D9D9D9"))
    r = 7
    put(ws, r, 1, "", None, None, B_ALL)
    lab = f"{c.furn}号炉　運転状態" if c.furn else "運転状態"
    merge(ws, r, 2, r, 4, lab, f(9, True), al("left"), B_ALL)
    for i, d in enumerate(days): put(ws, r, dc0 + i, None, None, None, B_ALL)
    for d, t in c.milestones:
        if d in days:
            put(ws, r, dc0 + days.index(d), t[0], f(10, True, color="C00000"), al("center"), B_ALL)
    ms = " → ".join(f"{t[1:]} {d.month}/{d.day}" for d, t in c.milestones)
    put(ws, r, nc, ms, f(8), al("left", shrink=True), B_ALL)
    rec.add("item", {"name": lab, "milestones": {t[1:]: d.isoformat() for d, t in c.milestones}}, ws.title, r)
    r += 1
    po = plan_override(c)
    prevc = None
    for no, t in enumerate(c.tasks, 1):
        tags = "、".join(W.BY_TAG[x].name for x in t.tags) if t.tags else ""
        put(ws, r, 1, no, f(8), al("center"), B_ALL)
        put(ws, r, 2, t.name, f(8), al("left", shrink=True), B_ALL)
        put(ws, r, 3, tags, f(7), al("left", shrink=True), B_ALL)
        comp = COMP.get(t.company, "SPC")
        put(ws, r, 4, comp if no == 1 or comp != prevc else "〃", f(8), al("center"), B_ALL)
        prevc = comp
        segs = po.get(t.tid) or plan_segs(t)
        factory = c.cid == "T2015-2" and t.tid == "T06"
        for i, d in enumerate(days):
            cell = put(ws, r, dc0 + i, None, None, None, B_ALL)
            if d.weekday() == 6: cell.fill = fill("F2DCDB")
            if in_segs(d, segs) and d.weekday() != 6: cell.fill = fill("404040" if form == "sv1" else "2F5597")
            if factory and dt.date(2015, 10, 12) <= d <= dt.date(2015, 10, 14): cell.fill = fill("A6A6A6")
        note = {"T00": "SPC 運転班"}.get(t.tid, "")
        if c.cid == "T2015-2":
            note = {"T06": "10/12〜14 スクリュー肉盛（工場）", "T04": "定期事業者検査 10/14 予定", "T03": "丸越築炉 別途発注",
                    "T00": "SPC 運転班", "T14": "SPC 運転班、協力会社 立会"}.get(t.tid, "")
        elif t.company == W.REFR:
            note = "丸越築炉 別途発注"
        elif "定期事業者検査" in t.name:
            note = "定期事業者検査あり"
        put(ws, r, nc, note, f(8), al("left", shrink=True), B_ALL)
        rec.add("item", {"no": no, "name": t.name, "equipment": tags, "company": comp,
                         "plan": [[a.isoformat(), b.isoformat()] for a, b in segs], "memo": note}, ws.title, r, wid(c, t))
        r += 1
    for _ in range(3):
        for cc in range(1, nc + 1):
            cell = put(ws, r, cc, None, None, None, B_ALL)
            if dc0 <= cc < nc and days[cc - dc0].weekday() == 6: cell.fill = fill("F2DCDB")
        r += 1
    put(ws, r + 1, 2, "※日曜は休工。足場・開放は SPC の炉内温度確認後に着手。作業時間 8:00〜17:00。", f(8))
    mk = c.start - dt.timedelta(days=17)
    put(ws, r + 2, 2, f"作成 {mk.year}/{mk.month}/{mk.day}　{who}（保全グループ）", f(8))
    outline(ws, 4, 1, r - 1, nc)
    widths(ws, {"A": 4, "B": 30, "C": 26, "D": 8})
    for i in range(len(days)): ws.column_dimensions[L(dc0 + i)].width = 3.2
    ws.column_dimensions[L(nc)].width = 26
    for rr in range(8, r): ws.row_dimensions[rr].height = 20
    page(ws, "landscape", 8, 1, 1, f"A1:{L(nc)}{r + 2}")
    props(wb, f"{W.FAC} {c.short} 工程表（{c.fy}）", f"{who}（【架空】ときわ北部エコサービス 保全グループ）", dt.datetime(mk.year, mk.month, mk.day, 15))
    wb.save(path); rec.save()

# ============================================================== 2024 様式
def s2024(c, path, fname):
    rec = Rec(fname)
    wb = Workbook(); ws = wb.active; ws.title = "工程表"
    days = date_cols(c)
    dc0 = 7; nc = dc0 + len(days)
    merge(ws, 1, 1, 1, nc - 8, f"{W.FAC}　{c.fy}年度　{c.short}　工程表（予定・実績）", f(16, True), al("left"), None)
    merge(ws, 2, 1, 2, nc - 8, f"工期：{c.start.year}年{c.start.month}月{c.start.day}日〜{c.end.month}月{c.end.day}日　第2期運営委託 年度整備（上限内）　施工：{W.SUB1}（職長 x1）", f(9), al("left"), None)
    approvals(ws, 1, nc - 6, ["承認", "確認", "作成"], ["所長", "sv2", "sv2"])
    for i, h in enumerate(["No.", "班", "作業項目", "対象機器", "業者", "予実"]):
        merge(ws, 4, 1 + i, 6, 1 + i, h, f(9, True), al("center"), B_ALL, fill("D9D9D9" if i not in (1, 5) else "DDEBF7"))
    month_header(ws, 4, dc0, days)
    for i, d in enumerate(days):
        fl = fill("F2DCDB") if d.weekday() == 6 else (fill("EDEDED") if d.weekday() == 5 else fill("D9D9D9"))
        put(ws, 5, dc0 + i, d.day, f(8, True), al("center"), B_ALL, fl)
        put(ws, 6, dc0 + i, YOUBI[d.weekday()], f(8), al("center"), B_ALL, fl)
    merge(ws, 4, nc, 6, nc, "備考", f(9, True), al("center"), B_ALL, fill("D9D9D9"))
    r = 7
    merge(ws, r, 1, r, 6, f"{c.furn}号炉 炉状態", f(9, True), al("left"), B_ALL, fill("DDEBF7"))
    s, e = c.start, c.end
    su = [d for d, t in c.milestones if "点火" in t][0]
    bands = [(s, s + dt.timedelta(days=2), "停止・冷却", "BDD7EE"), (s + dt.timedelta(days=3), su - dt.timedelta(days=4), "開放・点検整備", "FFF2CC"),
             (su - dt.timedelta(days=3), su - dt.timedelta(days=1), "復旧", "E2EFDA"), (su, e, "立上げ・試運転", "FCE4D6")]
    for i, d in enumerate(days): put(ws, r, dc0 + i, None, None, None, B_ALL)
    for a, b, t, col in bands:
        if a in days and b in days:
            merge(ws, r, dc0 + days.index(a), r, dc0 + days.index(b), t, f(8, True), al("center", shrink=True), B_ALL, fill(col))
    other = {"T2024-1": "2号炉は運転継続", "T2024-2": "1号炉は運転継続", "T2025-1": "2号炉は運転継続（9月から基幹改良）",
             "T2026-2": "1号炉は運転継続（9月から基幹改良）"}.get(c.cid, "")
    put(ws, r, nc, other, f(8), al("left", shrink=True), B_ALL)
    rec.add("item", {"name": f"{c.furn}号炉 炉状態", "bands": {t: [a.isoformat(), b.isoformat()] for a, b, t, _ in bands}, "memo": other}, ws.title, r)
    r += 1
    merge(ws, r, 1, r, 6, "節目", f(8), al("left"), B_ALL)
    for i, d in enumerate(days): put(ws, r, dc0 + i, None, None, None, B_ALL)
    for d, t in c.milestones:
        if d in days: put(ws, r, dc0 + days.index(d), "▼", f(9, True, color="C00000"), al("center"), B_ALL)
    ms = "・".join(f"{t[1:]} {d.month}/{d.day}" for d, t in c.milestones)
    put(ws, r, nc, ms, f(8), al("left"), B_ALL)
    rec.add("item", {"name": "節目", "milestones": {t[1:]: d.isoformat() for d, t in c.milestones}}, ws.title, r)
    r += 1
    po = plan_override(c)
    from contracts import B24
    for no, t in enumerate(c.tasks, 1):
        tags = "、".join(W.BY_TAG[x].name for x in t.tags) if t.tags else ""
        comp = COMP.get(t.company, "SPC")
        team = t.team if c.cid == "T2025-1" else B24.get(t.team, t.team)
        if t.team == "-": team = "-"
        planned = po.get(t.tid) or plan_segs(t)
        actual = segments([s_[0] for s_ in t.steps], 1) if t.steps else [list(t.plan)]
        if t.tid in ("T00",) or t.owner == "SPC":
            actual = [list(t.plan)]
        for k, kind in enumerate(("予定", "実績")):
            if k == 0:
                merge(ws, r, 1, r + 1, 1, no, f(8), al("center"), B_ALL)
                merge(ws, r, 2, r + 1, 2, team, f(8), al("center"), B_ALL)
                merge(ws, r, 3, r + 1, 3, t.name, f(8), al("left", wrap=True), B_ALL)
                merge(ws, r, 4, r + 1, 4, tags, f(7), al("left", wrap=True), B_ALL)
                merge(ws, r, 5, r + 1, 5, comp, f(8), al("center"), B_ALL)
            put(ws, r + k, 6, kind, f(7), al("center"), B_ALL)
            segs = planned if k == 0 else actual
            for i, d in enumerate(days):
                cell = put(ws, r + k, dc0 + i, None, None, None, B_ALL if kind == "実績" else Border(left=thin, right=thin, top=thin, bottom=dot))
                if d.weekday() == 6: cell.fill = fill("F2DCDB")
                if in_segs(d, segs) and d.weekday() != 6:
                    cell.fill = fill("BDD7EE") if kind == "予定" else fill("1F4E79")
        note = ""
        if c.cid == "T2025-1":
            note = {"T05": "ケージ曲がり 3 本交換で 1 日延長（5/27 完了）", "T03": "丸越築炉。最小補修（R8 基幹改良で全面更新）",
                    "T11": "共通機器（2号機で受入継続）", "T04": "定期事業者検査 5/21", "T00": "SPC 運転", "T13": "SPC 運転、北嶺 立会"}.get(t.tid, "")
        else:
            if any(s_[1].startswith("【追加】") for s_ in t.steps):
                note = "追加補修で 1 日延長"
            elif t.company == W.REFR: note = "丸越築炉"
            elif t.owner == "SPC": note = "SPC 運転"
            elif t.tags and W.BY_TAG[t.tags[0]].furn == 0: note = "共通機器"
        merge(ws, r, nc, r + 1, nc, note, f(7), al("left", wrap=True), B_ALL)
        base = {"no": no, "team": team, "name": t.name, "equipment": tags, "company": comp, "memo": note}
        rec.add("item", {**base, "row": "予定", "plan": [[a.isoformat(), b.isoformat()] for a, b in planned]}, ws.title, r, wid(c, t))
        rec.add("item", {**base, "row": "実績", "actual": [[a.isoformat(), b.isoformat()] for a, b in actual]}, ws.title, r + 1, wid(c, t))
        r += 2
    outline(ws, 4, 1, r - 1, nc)
    put(ws, r + 1, 3, "凡例　予定：薄青　実績：濃青　日曜休工。実績は日報（北嶺 x1）から sv2 が転記。", f(8))
    mk = c.start - dt.timedelta(days=21); fin = c.end + dt.timedelta(days=3)
    put(ws, r + 2, 3, f"作成 {mk.year}/{mk.month}/{mk.day}（予定）　実績記入 {fin.year}/{fin.month}/{fin.day}　sv2（保全グループ）", f(8))
    widths(ws, {"A": 4, "B": 5, "C": 26, "D": 24, "E": 7, "F": 4})
    for i in range(len(days)): ws.column_dimensions[L(dc0 + i)].width = 2.9
    ws.column_dimensions[L(nc)].width = 24
    page(ws, "landscape", 8, 1, 1, f"A1:{L(nc)}{r + 2}")
    props(wb, f"{W.FAC} {c.fy}年度 {c.short} 工程表", "sv2（【架空】ときわ北部エコサービス 保全グループ）", dt.datetime(fin.year, fin.month, fin.day, 17))
    wb.save(path); rec.save()

# ============================================================== メーカー（基幹改良）
def maker(c, path, fname):
    rec = Rec(fname)
    wb = Workbook(); ws = wb.active; ws.title = "実施工程表"
    w0 = c.start - dt.timedelta(days=c.start.weekday())
    nweeks = ((max(c.end, c.plan_end) - w0).days // 7) + 2
    weeks = [w0 + dt.timedelta(weeks=i) for i in range(nweeks)]
    wc0 = 6; nc = wc0 + len(weeks)
    inprog = c.status == "施工中"
    ver = "第4版" if not inprog else "第3版"
    made = "令和8年1月20日" if not inprog else "令和8年11月9日"
    merge(ws, 1, 1, 1, nc, "【架空】ときわ北部クリーンセンター基幹的設備改良工事（令和7〜9年度）", f(14, True), al("center"), None)
    merge(ws, 2, 1, 2, nc, f"実 施 工 程 表　（{'令和7' if c.fy == 2025 else '令和8'}年度施工分：{c.furn}号炉）　{ver}", f(13, True), al("center"), None)
    merge(ws, 3, 1, 3, nc, f"発注者：{W.UNION}　受注者：{W.MAKER}　現場代理人 m1　作成 {made}", f(9), al("right"), None)
    for i, h in enumerate(["工種", "作業", "対象機器（タグ）", "区分"]):
        merge(ws, 5, 1 + i, 7, 1 + i, h, f(9, True), al("center", wrap=True), B_ALL, fill("DEEAF6"))
    ws.column_dimensions["E"].hidden = True
    put(ws, 5, 5, "担当", f(8), al("center"), B_ALL)
    months = {}
    for i, w in enumerate(weeks): months.setdefault((w.year, w.month), []).append(i)
    for (y, m), idxs in months.items():
        merge(ws, 5, wc0 + idxs[0], 5, wc0 + idxs[-1], f"R{y - 2018}.{m}月", f(9, True), al("center"), B_ALL, fill("DEEAF6"))
    for i, w in enumerate(weeks):
        put(ws, 6, wc0 + i, f"{w.month}/{w.day}", f(7), al("center", rot=90), B_ALL, fill("DEEAF6"))
        put(ws, 7, wc0 + i, f"{i + 1}w", f(7), al("center"), B_ALL, fill("DEEAF6"))
    ws.row_dimensions[6].height = 30
    merge(ws, 5, nc, 7, nc, "備考", f(9, True), al("center"), B_ALL, fill("DEEAF6"))
    r = 8
    merge(ws, r, 1, r, 4, "主要節目", f(9, True), al("left"), B_ALL, fill("FFF2CC"))
    for i, w in enumerate(weeks): put(ws, r, wc0 + i, None, None, None, B_ALL, fill("FFF2CC"))
    for d, t in c.milestones:
        i = (d - w0).days // 7
        put(ws, r, wc0 + i, t[0], f(10, True, color="C00000"), al("center"), B_ALL, fill("FFF2CC"))
    txt = " / ".join(f"{d.month}/{d.day} {t[1:]}" for d, t in c.milestones)
    put(ws, r, nc, txt, f(7), al("left", wrap=True), B_ALL, fill("FFF2CC"))
    rec.add("item", {"name": "主要節目", "milestones": {t[1:]: d.isoformat() for d, t in c.milestones}}, ws.title, r)
    ws.row_dimensions[r].height = 36
    r += 1
    prev = None
    for n_, (kou, name, tags, attr, ps, pe, as_, ae, crew) in enumerate(c.tasks):
        a_end = min(ae, W.TODAY) if inprog else ae
        started = as_ <= W.TODAY
        for k in range(2):
            if k == 0:
                put(ws, r, 1, kou if kou != prev else None, f(8, True), al("center"), Border(left=thin, right=thin, top=thin if kou != prev else None))
                put(ws, r + 1, 1, None, None, None, Border(left=thin, right=thin))
                prev = kou
                merge(ws, r, 2, r + 1, 2, name, f(8), al("left", wrap=True), B_ALL)
                merge(ws, r, 3, r + 1, 3, "\n".join(tags) if tags else "—", f(7), al("center", wrap=True), B_ALL)
            put(ws, r + k, 4, "予定" if k == 0 else "実績", f(7), al("center"), B_ALL)
            a, b = (ps, pe) if k == 0 else (as_, a_end)
            for i, w in enumerate(weeks):
                we = w + dt.timedelta(days=6)
                cell = put(ws, r + k, wc0 + i, None, None, None, Border(left=hair, right=hair, top=thin if k == 0 else dot, bottom=dot if k == 0 else thin))
                if k == 1 and not started: continue
                if a <= we and b >= w:
                    full = a <= w and b >= we
                    if k == 0:
                        cell.value = "━" if full else "─"; cell.font = f(9, True, color="2F5597"); cell.alignment = al("center")
                    else:
                        cell.fill = fill("2F5597" if (as_, ae) == (ps, pe) else "C55A11")
        note = ""
        if (as_, ae) != (ps, pe) and started:
            note = f"予定 {ps.month}/{ps.day}〜{pe.month}/{pe.day} → 実績 {as_.month}/{as_.day}〜" + (f"{ae.month}/{ae.day}" if ae <= W.TODAY else "（施工中）")
        if "過熱器 据付" in name and c.fy == 2025: note += "（管材納期遅れ）"
        if inprog and started and ae > W.TODAY and not note: note = "施工中"
        merge(ws, r, nc, r + 1, nc, note, f(7), al("left", wrap=True), B_ALL)
        wk = [f"{c.cid}:K{n_:02d}:{t}" for t in tags] or [f"{c.cid}:K{n_:02d}"]
        base = {"kou": kou, "name": name, "equipment": " ".join(tags), "memo": note}
        rec.add("item", {**base, "row": "予定", "plan": [ps.isoformat(), pe.isoformat()]}, ws.title, r, wk)
        if started:
            rec.add("item", {**base, "row": "実績", "actual": [as_.isoformat(), a_end.isoformat() if (not inprog or ae <= W.TODAY) else None]}, ws.title, r + 1, wk)
        r += 2
    box(ws, 9, 1, r - 1, 1, Border(left=thin, right=thin, bottom=None))
    outline(ws, 5, 1, r - 1, nc)
    put(ws, r + 1, 2, "凡例　予定：━（青線）　実績：■青＝予定どおり　■橙＝予定から変更　◆節目　▲事象", f(8))
    put(ws, r + 2, 2, "年末年始休工 R7.12.29〜R8.1.3。工程変更は変更契約（第1回 R7.12.5）による。" if c.fy == 2025 else "実績は R8.11.13 まで記入。", f(8))
    widths(ws, {"A": 8, "B": 26, "C": 9, "D": 4.5, "E": 6})
    for i in range(len(weeks)): ws.column_dimensions[L(wc0 + i)].width = 3.6
    ws.column_dimensions[L(nc)].width = 30
    page(ws, "landscape", 8, 1, 1, f"A1:{L(nc)}{r + 2}")
    props(wb, f"{W.FAC} 基幹的設備改良工事 実施工程表（{'令和7' if c.fy == 2025 else '令和8'}年度施工分）", "m1（【架空】鷹羽エンジニアリング 工事部）",
          dt.datetime(2026, 1, 20, 10) if not inprog else dt.datetime(2026, 11, 9, 10))
    wb.save(path); rec.save()
