# -*- coding: utf-8 -*-
"""工程表：sv1 様式（2015 2号炉定期整備）／2024 様式（2025 1号炉定期整備）／メーカー様式（基幹改良 R7 施工分）"""
import datetime as dt
from openpyxl import Workbook
from openpyxl.comments import Comment
import world as W, works as K
from xlh import *

def segments(dates, join_gap=2):
    ds = sorted(set(dates))
    segs = []
    for d in ds:
        if segs and (d - segs[-1][1]).days <= join_gap:
            segs[-1][1] = d
        else:
            segs.append([d, d])
    return segs

def in_segs(d, segs):
    return any(a <= d <= b for a, b in segs)

def approvals(ws, r, c, labels, names):
    for i, (lb, nm) in enumerate(zip(labels, names)):
        put(ws, r, c + i * 2, lb, f(8), al("center"), B_ALL)
        ws.merge_cells(start_row=r, start_column=c + i * 2, end_row=r, end_column=c + i * 2 + 1)
        box(ws, r, c + i * 2, r, c + i * 2 + 1)
        merge(ws, r + 1, c + i * 2, r + 2, c + i * 2 + 1, nm, f(9), al("center"), B_ALL)

# ============================================================== sv1 様式 2015
def sv1_2015(path):
    c = W.CONTRACTS["K1"]
    wb = Workbook(); ws = wb.active; ws.title = "工程表"
    days = list(W.daterange(dt.date(2015, 10, 1), dt.date(2015, 10, 31)))
    dc0 = 5; nc = dc0 + len(days)
    merge(ws, 1, 1, 1, 28, f"{W.FAC}　2号炉定期整備工事　工程表", f(16, True), al("left"), None)
    merge(ws, 2, 1, 2, 28, f"工期：2015年10月5日〜2015年10月30日（炉停止 10/5、点火 10/26）　発注：{W.SPC}　施工：{W.SUB1}", f(9), al("left"), None)
    approvals(ws, 1, nc - 6, ["承認", "確認", "作成"], ["所長", "GL", "sv1"])
    put(ws, 4, 1, "No.", f(9, True), al("center"), B_ALL, fill("D9D9D9"))
    for i, h in enumerate(["作業項目", "対象機器", "業者"]):
        merge(ws, 4, 2 + i, 6, 2 + i, h, f(9, True), al("center"), B_ALL, fill("D9D9D9"))
    ws.merge_cells("A4:A6")
    merge(ws, 4, dc0, 4, nc - 1, "2015年10月", f(9, True), al("center"), B_ALL, fill("D9D9D9"))
    for i, d in enumerate(days):
        fl = fill("F2DCDB") if d.weekday() == 6 else (fill("EDEDED") if d.weekday() == 5 else fill("D9D9D9"))
        put(ws, 5, dc0 + i, d.day, f(8, True), al("center"), B_ALL, fl)
        put(ws, 6, dc0 + i, YOUBI[d.weekday()], f(8), al("center"), B_ALL, fl)
    merge(ws, 4, nc, 6, nc, "備考", f(9, True), al("center"), B_ALL, fill("D9D9D9"))
    # 炉の状態（マイルストーン）
    r = 7
    put(ws, r, 1, "", None, None, B_ALL); merge(ws, r, 2, r, 4, "2号炉　運転状態", f(9, True), al("left"), B_ALL)
    for i, d in enumerate(days):
        put(ws, r, dc0 + i, None, None, None, B_ALL)
    for d, t in K.K1_MILESTONES:
        put(ws, r, dc0 + days.index(d), t[0], f(10, True, color="C00000"), al("center"), B_ALL)
    put(ws, r, nc, "停止 10/5 → 点火 10/26 → 定格 10/30", f(8), al("left"), B_ALL)
    r += 1
    plan_over = {"T04": [(dt.date(2015, 10, 8), dt.date(2015, 10, 14)), (dt.date(2015, 10, 16), dt.date(2015, 10, 17)), (dt.date(2015, 10, 19), dt.date(2015, 10, 19))],
                 "T13": [(dt.date(2015, 10, 20), dt.date(2015, 10, 24))], "T00": [(dt.date(2015, 10, 5), dt.date(2015, 10, 7))],
                 "T14": [(dt.date(2015, 10, 26), dt.date(2015, 10, 30))]}
    no = 0
    for t in K.K1:
        no += 1
        tags = "、".join(W.BY_TAG[x].name for x in t.tags) if t.tags else ""
        put(ws, r, 1, no, f(8), al("center"), B_ALL)
        put(ws, r, 2, t.name, f(8), al("left", shrink=True), B_ALL)
        put(ws, r, 3, tags, f(7), al("left", shrink=True), B_ALL)
        comp = {W.SUB1: "北嶺", W.REFR: "丸越築炉"}.get(t.company, "SPC")
        put(ws, r, 4, comp if no == 1 or comp != prevc else "〃", f(8), al("center"), B_ALL)
        prevc = comp
        if t.tid in plan_over:
            segs = plan_over[t.tid]
        else:
            ds = [s[0] for s in t.steps if not s[1].startswith("【追加】")]
            segs = segments(ds)
        factory = t.tid == "T06"
        for i, d in enumerate(days):
            cell = put(ws, r, dc0 + i, None, None, None, B_ALL)
            if d.weekday() == 6:
                cell.fill = fill("F2DCDB")
            if in_segs(d, segs) and d.weekday() != 6:
                cell.fill = fill("404040")
            if factory and dt.date(2015, 10, 12) <= d <= dt.date(2015, 10, 14):
                cell.fill = fill("A6A6A6")
        note = {"T06": "10/12〜14 スクリュー肉盛（工場）", "T04": "定期事業者検査 10/14 予定", "T03": "丸越築炉 別途発注",
                "T00": "SPC 運転班", "T14": "SPC 運転班、協力会社 立会"}.get(t.tid, "")
        put(ws, r, nc, note, f(8), al("left", shrink=True), B_ALL)
        r += 1
    # 余白行（追記用）
    for _ in range(3):
        for cc in range(1, nc + 1):
            cell = put(ws, r, cc, None, None, None, B_ALL)
            if 5 <= cc < nc and days[cc - dc0].weekday() == 6: cell.fill = fill("F2DCDB")
        r += 1
    put(ws, r + 1, 2, "※日曜は休工。足場・開放は SPC の炉内温度確認後に着手。作業時間 8:00〜17:00。", f(8))
    put(ws, r + 2, 2, "作成 2015/9/18　sv1（保全グループ）", f(8))
    outline(ws, 4, 1, r - 1, nc)
    widths(ws, {"A": 4, "B": 30, "C": 26, "D": 8})
    for i in range(len(days)): ws.column_dimensions[L(dc0 + i)].width = 3.2
    ws.column_dimensions[L(nc)].width = 26
    for rr in range(8, r): ws.row_dimensions[rr].height = 20
    page(ws, "landscape", 8, 1, 1, f"A1:{L(nc)}{r + 2}")
    props(wb, f"{W.FAC} 2号炉定期整備工事 工程表（2015）", "sv1（【架空】ときわ北部エコサービス 保全グループ）", dt.datetime(2015, 9, 18, 15, 0))
    wb.save(path)

# ============================================================== 2024 様式（sv2）2025 1号炉
def sv2_2025(path):
    c = W.CONTRACTS["K2"]
    wb = Workbook(); ws = wb.active; ws.title = "工程表"
    days = list(W.daterange(dt.date(2025, 5, 8), dt.date(2025, 6, 10)))
    # 列：A No / B 班(追加) / C 作業項目 / D 対象機器 / E 業者 / F 予実(追加) / G.. 日
    dc0 = 7; nc = dc0 + len(days)
    merge(ws, 1, 1, 1, 30, f"{W.FAC}　2025年度　1号炉定期整備工事　工程表（予定・実績）", f(16, True), al("left"), None)
    merge(ws, 2, 1, 2, 30, f"工期：2025年5月12日〜6月6日　第2期運営委託 年度整備（上限内）　施工：{W.SUB1}（職長 x1）", f(9), al("left"), None)
    approvals(ws, 1, nc - 6, ["承認", "確認", "作成"], ["所長", "sv2", "sv2"])
    for i, h in enumerate(["No.", "班", "作業項目", "対象機器", "業者", "予実"]):
        merge(ws, 4, 1 + i, 6, 1 + i, h, f(9, True), al("center"), B_ALL, fill("D9D9D9" if i not in (1, 5) else "DDEBF7"))
    m5 = [d for d in days if d.month == 5]; m6 = [d for d in days if d.month == 6]
    merge(ws, 4, dc0, 4, dc0 + len(m5) - 1, "2025年5月", f(9, True), al("center"), B_ALL, fill("D9D9D9"))
    merge(ws, 4, dc0 + len(m5), 4, nc - 1, "6月", f(9, True), al("center"), B_ALL, fill("D9D9D9"))
    for i, d in enumerate(days):
        fl = fill("F2DCDB") if d.weekday() == 6 else (fill("EDEDED") if d.weekday() == 5 else fill("D9D9D9"))
        put(ws, 5, dc0 + i, d.day, f(8, True), al("center"), B_ALL, fl)
        put(ws, 6, dc0 + i, YOUBI[d.weekday()], f(8), al("center"), B_ALL, fl)
    merge(ws, 4, nc, 6, nc, "備考", f(9, True), al("center"), B_ALL, fill("D9D9D9"))
    # 炉状態の帯（2024 で追加）
    r = 7
    merge(ws, r, 1, r, 6, "1号炉 炉状態", f(9, True), al("left"), B_ALL, fill("DDEBF7"))
    bands = [(dt.date(2025, 5, 12), dt.date(2025, 5, 14), "停止・冷却", "BDD7EE"), (dt.date(2025, 5, 15), dt.date(2025, 5, 29), "開放・点検整備", "FFF2CC"),
             (dt.date(2025, 5, 30), dt.date(2025, 6, 1), "復旧", "E2EFDA"), (dt.date(2025, 6, 2), dt.date(2025, 6, 6), "立上げ・試運転", "FCE4D6")]
    for i, d in enumerate(days): put(ws, r, dc0 + i, None, None, None, B_ALL)
    for a, b, t, col in bands:
        ia, ib = days.index(a), days.index(b)
        merge(ws, r, dc0 + ia, r, dc0 + ib, t, f(8, True), al("center", shrink=True), B_ALL, fill(col))
    put(ws, r, nc, "2号炉は運転継続（9月から基幹改良）", f(8), al("left", shrink=True), B_ALL)
    r += 1
    r_m = r
    merge(ws, r, 1, r, 6, "節目", f(8), al("left"), B_ALL)
    for i, d in enumerate(days): put(ws, r, dc0 + i, None, None, None, B_ALL)
    for d, t in K.K2_MILESTONES:
        put(ws, r, dc0 + days.index(d), "▼", f(9, True, color="C00000"), al("center"), B_ALL)
    put(ws, r, nc, "停止 5/12・点火 6/2・定格 6/6", f(8), al("left"), B_ALL)
    r += 1
    plan_over = {"T05": [(dt.date(2025, 5, 19), dt.date(2025, 5, 26))], "T13": [(dt.date(2025, 6, 2), dt.date(2025, 6, 6))],
                 "T00": [(dt.date(2025, 5, 12), dt.date(2025, 5, 14))], "T12": [(dt.date(2025, 5, 28), dt.date(2025, 5, 31))],
                 "T01": [(dt.date(2025, 5, 15), dt.date(2025, 5, 15)), (dt.date(2025, 5, 30), dt.date(2025, 5, 30))]}
    no = 0
    for t in K.K2:
        no += 1
        tags = "、".join(W.BY_TAG[x].name for x in t.tags) if t.tags else ""
        comp = {W.SUB1: "北嶺", W.REFR: "丸越築炉"}.get(t.company, "SPC")
        for k, kind in enumerate(("予定", "実績")):
            if k == 0:
                merge(ws, r, 1, r + 1, 1, no, f(8), al("center"), B_ALL)
                merge(ws, r, 2, r + 1, 2, t.team, f(8), al("center"), B_ALL)
                merge(ws, r, 3, r + 1, 3, t.name, f(8), al("left", wrap=True), B_ALL)
                merge(ws, r, 4, r + 1, 4, tags, f(7), al("left", wrap=True), B_ALL)
                merge(ws, r, 5, r + 1, 5, comp, f(8), al("center"), B_ALL)
            put(ws, r + k, 6, kind, f(7), al("center"), B_ALL)
            if kind == "予定":
                segs = plan_over.get(t.tid) or segments([s[0] for s in t.steps])
            else:
                segs = segments([s[0] for s in t.steps], 1) if t.steps else ([] if t.owner == "SPC" and t.tid == "T13" else [t.actual])
                if t.tid in ("T00", "T13"):
                    segs = [t.plan]
            for i, d in enumerate(days):
                cell = put(ws, r + k, dc0 + i, None, None, None, B_ALL if kind == "実績" else Border(left=thin, right=thin, top=thin, bottom=dot))
                if d.weekday() == 6: cell.fill = fill("F2DCDB")
                if in_segs(d, segs) and d.weekday() != 6:
                    if kind == "予定":
                        cell.fill = fill("BDD7EE")
                    else:
                        cell.fill = fill("1F4E79")
            if k == 0:
                note = {"T05": "ケージ曲がり 3 本交換で 1 日延長（5/27 完了）", "T03": "丸越築炉。最小補修（R8 基幹改良で全面更新）",
                        "T11": "共通機器（2号機で受入継続）", "T04": "定期事業者検査 5/21", "T00": "SPC 運転", "T13": "SPC 運転、北嶺 立会"}.get(t.tid, "")
                merge(ws, r, nc, r + 1, nc, note, f(7), al("left", wrap=True), B_ALL)
        r += 2
    outline(ws, 4, 1, r - 1, nc)
    put(ws, r + 1, 3, "凡例　予定：薄青　実績：濃青　日曜休工。実績は日報（北嶺 x1）から sv2 が転記。", f(8))
    put(ws, r + 2, 3, "作成 2025/4/21（予定）　実績記入 2025/6/9　sv2（保全グループ）", f(8))
    widths(ws, {"A": 4, "B": 5, "C": 26, "D": 24, "E": 7, "F": 4})
    for i in range(len(days)): ws.column_dimensions[L(dc0 + i)].width = 2.9
    ws.column_dimensions[L(nc)].width = 24
    page(ws, "landscape", 8, 1, 1, f"A1:{L(nc)}{r + 2}")
    props(wb, f"{W.FAC} 2025年度 1号炉定期整備工事 工程表", "sv2（【架空】ときわ北部エコサービス 保全グループ）", dt.datetime(2025, 6, 9, 17, 30))
    wb.save(path)

# ============================================================== メーカー様式（基幹改良）
def maker_k3(path):
    wb = Workbook(); ws = wb.active; ws.title = "実施工程表"
    w0 = dt.date(2025, 8, 18)
    weeks = [w0 + dt.timedelta(weeks=i) for i in range(23)]
    wc0 = 6; nc = wc0 + len(weeks)
    merge(ws, 1, 1, 1, nc, "【架空】ときわ北部クリーンセンター基幹的設備改良工事（令和7〜9年度）", f(14, True), al("center"), None)
    merge(ws, 2, 1, 2, nc, "実 施 工 程 表　（令和7年度施工分：2号炉）　第4版", f(13, True), al("center"), None)
    merge(ws, 3, 1, 3, nc, f"発注者：{W.UNION}　受注者：{W.MAKER}　現場代理人 m1　作成 令和8年1月20日", f(9), al("right"), None)
    for i, h in enumerate(["工種", "作業", "対象機器（タグ）", "区分"]):
        merge(ws, 5, 1 + i, 7, 1 + i, h, f(9, True), al("center", wrap=True), B_ALL, fill("DEEAF6"))
    ws.column_dimensions["E"].hidden = True   # 版を重ねる途中で使わなくなった「担当」列
    put(ws, 5, 5, "担当", f(8), al("center"), B_ALL)
    months = {}
    for i, w in enumerate(weeks):
        months.setdefault((w.year, w.month), []).append(i)
    for (y, m), idxs in months.items():
        merge(ws, 5, wc0 + idxs[0], 5, wc0 + idxs[-1], f"{'R7' if y == 2025 else 'R8'}.{m}月", f(9, True), al("center"), B_ALL, fill("DEEAF6"))
    for i, w in enumerate(weeks):
        put(ws, 6, wc0 + i, f"{w.month}/{w.day}", f(7), al("center", rot=90), B_ALL, fill("DEEAF6"))
        put(ws, 7, wc0 + i, f"{i + 1}w", f(7), al("center"), B_ALL, fill("DEEAF6"))
    ws.row_dimensions[6].height = 30
    merge(ws, 5, nc, 7, nc, "備考", f(9, True), al("center"), B_ALL, fill("DEEAF6"))
    r = 8
    merge(ws, r, 1, r, 4, "主要節目", f(9, True), al("left"), B_ALL, fill("FFF2CC"))
    for i, w in enumerate(weeks):
        put(ws, r, wc0 + i, None, None, None, B_ALL, fill("FFF2CC"))
    for d, t in K.K3_EVENTS:
        i = (d - w0).days // 7
        put(ws, r, wc0 + i, t[0], f(10, True, color="C00000"), al("center"), B_ALL, fill("FFF2CC"))
    put(ws, r, nc, " / ".join(f"{d.month}/{d.day} {t[1:]}" for d, t in K.K3_EVENTS), f(7), al("left", wrap=True), B_ALL, fill("FFF2CC"))
    ws.row_dimensions[r].height = 36
    r += 1
    prev = None
    for kou, name, tags, attr, ps, pe, as_, ae, crew in K.K3:
        for k in range(2):
            if k == 0:
                put(ws, r, 1, kou if kou != prev else None, f(8, True), al("center"), Border(left=thin, right=thin, top=thin if kou != prev else None))
                put(ws, r + 1, 1, None, None, None, Border(left=thin, right=thin))
                prev = kou
                merge(ws, r, 2, r + 1, 2, name, f(8), al("left", wrap=True), B_ALL)
                merge(ws, r, 3, r + 1, 3, "\n".join(tags) if tags else "—", f(7), al("center", wrap=True), B_ALL)
            put(ws, r + k, 4, "予定" if k == 0 else "実績", f(7), al("center"), B_ALL)
            a, b = (ps, pe) if k == 0 else (as_, ae)
            for i, w in enumerate(weeks):
                we = w + dt.timedelta(days=6)
                cell = put(ws, r + k, wc0 + i, None, None, None, Border(left=hair, right=hair, top=thin if k == 0 else dot, bottom=dot if k == 0 else thin))
                if a <= we and b >= w:
                    full = a <= w and b >= we
                    if k == 0:
                        cell.fill = fill("FFFFFF"); cell.value = "━" if full else "─"
                        cell.font = f(9, True, color="2F5597"); cell.alignment = al("center")
                    else:
                        cell.fill = fill("2F5597" if (as_, ae) == (ps, pe) else "C55A11")
            if k == 0:
                note = ""
                if (as_, ae) != (ps, pe):
                    note = f"予定 {ps.month}/{ps.day}〜{pe.month}/{pe.day} → 実績 {as_.month}/{as_.day}〜{ae.month}/{ae.day}"
                if "過熱器 据付" in name:
                    note += "（管材納期遅れ）"
                merge(ws, r, nc, r + 1, nc, note, f(7), al("left", wrap=True), B_ALL)
        r += 2
    for i, w in enumerate(weeks):
        if any(h >= w and h <= w + dt.timedelta(days=6) for h in K.K3_HOLIDAYS):
            ws.cell(row=7, column=wc0 + i).value = "休"
    box(ws, 9, 1, r - 1, 1, Border(left=thin, right=thin, bottom=None))
    outline(ws, 5, 1, r - 1, nc)
    put(ws, r + 1, 2, "凡例　予定：━（青線）　実績：■青＝予定どおり　■橙＝予定から変更　◆節目　▲事象", f(8))
    put(ws, r + 2, 2, "年末年始休工 R7.12.29〜R8.1.3。工程変更は変更契約（第1回 R7.12.5）による。", f(8))
    widths(ws, {"A": 8, "B": 26, "C": 9, "D": 4.5, "E": 6})
    for i in range(len(weeks)): ws.column_dimensions[L(wc0 + i)].width = 3.6
    ws.column_dimensions[L(nc)].width = 30
    page(ws, "landscape", 8, 1, 1, f"A1:{L(nc)}{r + 2}")
    props(wb, f"{W.FAC} 基幹的設備改良工事 実施工程表（令和7年度施工分）", "m1（【架空】鷹羽エンジニアリング 工事部）", dt.datetime(2026, 1, 20, 10, 0))
    wb.save(path)
