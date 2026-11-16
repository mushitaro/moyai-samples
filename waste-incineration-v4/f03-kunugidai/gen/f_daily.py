# -*- coding: utf-8 -*-
"""作業日報（全年度）。メーカー：m1 様式（〜H25）／m2 様式（H26〜）。市直営：d2 週報（H29〜）。d1 手書きは f_scan.py。"""
import datetime as dt
from openpyxl import Workbook
import xlhelp as X
from rec import Rec
from model import agent, city_person, WORKS
from world import wareki

F = X.font
WD = "月火水木金土日"
TRADE = {"鳶": "鳶工", "機械A": "機械工", "機械B": "機械工", "ボイラ": "溶接工", "排ガス": "機械工", "築炉": "築炉工", "電計": "電工・計装工",
         "全": "機械工", "A": "機械工", "A2": "機械工", "B": "溶接工", "B2": "溶接工", "C": "機械工", "C2": "機械工", "D": "機械工",
         "D2": "機械工", "E": "電工・計装工", "E2": "電工・計装工", "H": "機械工", "F": "機械工", "T": "機械工", "全班": "機械工"}


def _wa(d):
    if d.year >= 2019 and not (d.year == 2019 and d.month < 5):
        return "令和%d年%d月%d日（%s）" % (d.year - 2018, d.month, d.day, WD[d.weekday()])
    return "平成%d年%d月%d日（%s）" % (d.year - 1988, d.month, d.day, WD[d.weekday()])


def _tm(h):
    return {6: "8:00〜15:00", 4: "8:00〜12:00", 8: "8:00〜17:00", 9: "8:00〜18:00", 10: "8:00〜19:00"}.get(int(h), "8:00〜17:00")


def m1(c, path, fname):
    rec = Rec(fname)
    wb = Workbook()
    wb.remove(wb.active)
    X.props(wb, "【架空】%s 作業日報" % c.name, "汐見重工環境 %s" % agent(c.fy), "作業日報", "m1 様式")
    tasks = {t.id: t for t in c.tasks}
    fixed = getattr(c, "daily_fixed_m1", {}) or {}
    a0 = min(c.daily) if c.daily else c.start
    for d, v in sorted(c.daily.items()):
        sh = "%d.%d" % (d.month, d.day)
        ws = wb.create_sheet(sh)
        X.widths(ws, [4, 46, 7, 13, 2, 10, 7])
        X.put(ws, 1, 1, "【架空】作　業　日　報", F(size=16, bold=True), X.CENTER)
        ws.merge_cells("A1:G1")
        X.put(ws, 3, 1, "工事名", F(size=9), X.CENTER)
        X.put(ws, 3, 2, c.name, F(size=10), X.LEFT)
        ws.merge_cells("B3:D3")
        wx = fixed[d]["weather"] if d in fixed else v["weather"]
        X.put(ws, 3, 6, "天候", F(size=9), X.CENTER)
        X.put(ws, 3, 7, wx, F(size=10), X.CENTER)
        X.put(ws, 4, 1, "日付", F(size=9), X.CENTER)
        X.put(ws, 4, 2, _wa(d), F(size=10), X.LEFT)
        ws.merge_cells("B4:D4")
        no = (d - c.start).days + 1
        X.put(ws, 4, 6, "No.", F(size=9), X.CENTER)
        X.put(ws, 4, 7, no, F(size=10), X.CENTER)
        X.box(ws, 3, 1, 4, 4)
        X.box(ws, 3, 6, 4, 7)
        head = {"工事名": c.name, "日付": d.isoformat(), "天候": wx, "日報No": no}
        for i, h in enumerate(["No", "作　業　内　容", "人員", "時間"], 1):
            X.put(ws, 6, i, h, F(bold=True, size=9), X.CENTER, X.fill("F2F2F2"))
        r = 7
        trades = {}
        if d in fixed:
            rows = [(tid, crew, txt, n, tm) for tid, crew, txt, n, tm in fixed[d]["rows"]]
            safe = fixed[d]["safe"]
        else:
            rows = [(tid, crew, ("%s　%s" % (place, txt)), n, _tm(h)) for tid, crew, place, txt, n, h, prog in v["rows"]]
            safe = "高所作業 安全帯使用" if d.day % 3 == 0 else ("炉内入炉前の酸欠測定（O2 20.9%）" if d.day % 3 == 1 else "KY 実施")
        for i, (tid, crew, txt, n, tm) in enumerate(rows[:12], 1):
            X.put(ws, r, 1, i, F(size=9), X.CENTER)
            X.put(ws, r, 2, txt, F(size=10), X.LEFT)
            X.put(ws, r, 3, "%d名" % n, F(size=10), X.CENTER)
            X.put(ws, r, 4, tm, F(size=9), X.CENTER)
            ws.row_dimensions[r].height = 22
            tr = TRADE.get(crew, "機械工")
            trades[tr] = trades.get(tr, 0) + n
            f = dict(head)
            f.update({"No": i, "作業内容": txt, "人員": n, "時間": tm})
            rec.add(sh, r, "item", f, works=tasks[tid].works if tid in tasks else None)
            r += 1
        while r < 19:
            ws.row_dimensions[r].height = 22
            r += 1
        X.box(ws, 6, 1, 18, 4, side=X.med, inner=X.thin)
        X.put(ws, 6, 6, "職　種", F(bold=True, size=9), X.CENTER, X.fill("F2F2F2"))
        X.put(ws, 6, 7, "人員", F(bold=True, size=9), X.CENTER, X.fill("F2F2F2"))
        rr = 7
        for t in ["現場代理人", "鳶工", "機械工", "溶接工", "築炉工", "電工・計装工", "安全管理"]:
            n = 1 if t in ("現場代理人", "安全管理") else trades.get(t, 0)
            X.put(ws, rr, 6, t, F(size=9), X.LEFT)
            X.put(ws, rr, 7, n if n else "", F(size=10), X.CENTER)
            rr += 1
        tot = sum(trades.values()) + 2
        X.put(ws, rr, 6, "合　計", F(bold=True, size=9), X.CENTER)
        X.put(ws, rr, 7, tot, F(bold=True, size=10), X.CENTER)
        rec.add(sh, rr, "total", {"日付": d.isoformat(), "合計": tot})
        X.box(ws, 6, 6, rr, 7, side=X.med)
        X.put(ws, 20, 1, "安全", F(size=9), X.CENTER)
        X.put(ws, 20, 2, safe, F(size=10), X.LEFT)
        ws.merge_cells("B20:D20")
        X.put(ws, 21, 1, "特記", F(size=9), X.CENTER)
        note = "" if d in fixed else v.get("note", "")
        if d in fixed and d == dt.date(2013, 5, 22):
            note = "BF 4室 破損ろ布 2本は予備品で交換（市 支給予備）。残りの予備 8本。"
        X.put(ws, 21, 2, note, F(size=10), X.LEFT)
        ws.merge_cells("B21:D22")
        ws.merge_cells("A21:A22")
        X.box(ws, 20, 1, 22, 4)
        if note:
            rec.add(sh, 21, "item", {"日付": d.isoformat(), "特記": note})
        X.put(ws, 24, 4, "現場代理人", F(size=9), X.CENTER)
        X.put(ws, 24, 6, "市 確認", F(size=9), X.CENTER)
        ws.merge_cells("F24:G24")
        X.put(ws, 25, 4, agent(c.fy), F(size=11), X.CENTER)
        X.put(ws, 25, 6, city_person(c.fy), F(size=11), X.CENTER)
        ws.merge_cells("F25:G26")
        ws.merge_cells("D25:D26")
        X.box(ws, 24, 4, 26, 4)
        X.box(ws, 24, 6, 26, 7)
        X.put(ws, 28, 1, "汐見重工環境株式会社　くぬぎ台作業所", F(size=9), X.LEFT)
        ws.merge_cells("A28:D28")
        X.page(ws, "portrait", "A4", 1, 1)
    wb.save(path)
    rec.save()


def m2(c, path, fname):
    rec = Rec(fname)
    wb = Workbook()
    wb.remove(wb.active)
    X.props(wb, "【架空】%s 作業日報" % c.name, "汐見重工環境 m2", "作業日報", "m2 様式（H26〜）")
    tasks = {t.id: t for t in c.tasks}
    tno = {t.id: i for i, t in enumerate(c.tasks, 1)}
    for d, v in sorted(c.daily.items()):
        sh = "%02d%02d" % (d.month, d.day)
        ws = wb.create_sheet(sh)
        X.widths(ws, [4, 6, 6, 26, 50, 5, 5, 6, 6])
        X.put(ws, 1, 1, "【架空】作業日報", F(size=15, bold=True), X.NOWRAP)
        no = (d - c.start).days + 1
        X.put(ws, 1, 6, "日報 No.%d" % no, F(size=10), X.RIGHT)
        ws.merge_cells("F1:I1")
        X.put(ws, 2, 1, "工事名", F(size=9), X.CENTER, X.fill("F2F2F2"))
        ws.merge_cells("A2:B2")
        X.put(ws, 2, 3, c.name, F(size=10, bold=True), X.LEFT)
        ws.merge_cells("C2:E2")
        X.put(ws, 2, 6, "天候", F(size=9), X.CENTER, X.fill("F2F2F2"))
        X.put(ws, 2, 7, v["weather"], F(size=10), X.CENTER)
        ws.merge_cells("G2:I2")
        X.put(ws, 3, 1, "日付", F(size=9), X.CENTER, X.fill("F2F2F2"))
        ws.merge_cells("A3:B3")
        X.put(ws, 3, 3, "%d/%d/%d（%s）　作業時間 8:00〜17:00" % (d.year, d.month, d.day, WD[d.weekday()]), F(size=10), X.LEFT)
        ws.merge_cells("C3:E3")
        z = c.zentei
        stopped = bool(z and z[0] <= d <= z[1])
        if c.kind == "teishu":
            st = "%d号炉 停止中" % c.unit
        elif c.kind == "hoshu":
            st = "%s停止中" % ("%d号炉 " % c.unit if c.unit else "")
        else:
            st = "全炉停止" if stopped else "各炉 運転（順次 定修）"
        X.put(ws, 3, 6, "炉", F(size=9), X.CENTER, X.fill("F2F2F2"))
        X.put(ws, 3, 7, st, F(size=8), X.LEFT)
        ws.merge_cells("G3:I3")
        X.box(ws, 2, 1, 3, 9)
        head = {"工事名": c.name, "日付": d.isoformat(), "天候": v["weather"], "日報No": no}
        hdr = ["No", "工程", "班", "作業箇所", "作業内容", "人数", "時間", "工数\n(人時)", "進捗\n(%)"]
        for i, h in enumerate(hdr, 1):
            X.put(ws, 5, i, h, F(bold=True, size=9), X.CENTER, X.fill("DDEBF7"))
        ws.row_dimensions[5].height = 24
        r = 6
        totp = toth = 0
        bycrew = {}
        rows = v["rows"]
        for i, (tid, crew, place, txt, n, h, prog) in enumerate(rows, 1):
            X.put(ws, r, 1, i, F(size=9), X.CENTER)
            kno = ("No.%d" % tno[tid]) if tid in tno else "－"
            X.put(ws, r, 2, kno, F(size=8), X.CENTER)
            X.put(ws, r, 3, crew, F(size=9), X.CENTER)
            X.put(ws, r, 4, place, F(size=8), X.LEFT)
            X.put(ws, r, 5, txt, F(size=9, color="C00000" if txt.startswith("【追加】") else None), X.LEFT)
            X.put(ws, r, 6, n, F(size=9), X.CENTER)
            X.put(ws, r, 7, h, F(size=9), X.CENTER)
            X.put(ws, r, 8, n * h, F(size=9), X.RIGHT)
            X.put(ws, r, 9, prog if prog is not None else "", F(size=9), X.CENTER)
            ws.row_dimensions[r].height = 27
            totp += n
            toth += n * h
            bycrew[crew] = bycrew.get(crew, 0) + n
            f = dict(head)
            f.update({"No": i, "工程": kno, "班": crew, "作業箇所": place, "作業内容": txt, "人数": n, "時間": h,
                      "manDays": round(n * h / 8.0, 2), "manDaysText": "%d人時" % (n * h), "進捗": prog})
            rec.add(sh, r, "item", f, works=tasks[tid].works if tid in tasks else None)
            r += 1
        while r < 14:
            ws.row_dimensions[r].height = 20
            r += 1
        X.put(ws, r, 4, "現場管理（代理人・安全）", F(size=8), X.LEFT)
        X.put(ws, r, 6, 2, F(size=9), X.CENTER)
        X.put(ws, r, 7, 8, F(size=9), X.CENTER)
        X.put(ws, r, 8, 16, F(size=9), X.RIGHT)
        rec.add(sh, r, "item", {"日付": d.isoformat(), "作業箇所": "現場管理（代理人・安全）", "人数": 2, "時間": 8, "manDays": 2, "manDaysText": "16人時"})
        r += 1
        X.put(ws, r, 5, "合　計", F(bold=True, size=9), X.RIGHT)
        X.put(ws, r, 6, totp + 2, F(bold=True, size=9), X.CENTER)
        X.put(ws, r, 8, toth + 16, F(bold=True, size=9), X.RIGHT)
        rec.add(sh, r, "total", {"日付": d.isoformat(), "人数": totp + 2, "manDays": round((toth + 16) / 8.0, 2), "manDaysText": "%d人時" % (toth + 16)})
        X.box(ws, 5, 1, r, 9, side=X.med, inner=X.thin)
        r += 1
        X.put(ws, r, 1, "班別", F(size=8), X.CENTER, X.fill("F2F2F2"))
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        X.put(ws, r, 3, "　".join("%s %d名" % (k, v2) for k, v2 in bycrew.items()), F(size=8), X.LEFT)
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=9)
        X.box(ws, r, 1, r, 9)
        r += 1
        X.put(ws, r, 1, "特記\n事項", F(size=9), X.CENTER, X.fill("F2F2F2"))
        ws.merge_cells(start_row=r, start_column=1, end_row=r + 2, end_column=2)
        X.put(ws, r, 3, v.get("note", ""), F(size=9), X.LEFT_TOP)
        ws.merge_cells(start_row=r, start_column=3, end_row=r + 2, end_column=9)
        if v.get("note"):
            rec.add(sh, r, "item", {"日付": d.isoformat(), "特記事項": v["note"]})
        X.box(ws, r, 1, r + 2, 9)
        r += 3
        X.put(ws, r, 1, "安全", F(size=9), X.CENTER, X.fill("F2F2F2"))
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        X.put(ws, r, 3, "KY・TBM 実施／災害 なし／ヒヤリハット なし", F(size=9), X.LEFT)
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
        X.put(ws, r, 6, "代理人", F(size=8), X.CENTER)
        X.put(ws, r, 7, "m2", F(size=9), X.CENTER)
        X.put(ws, r, 8, "市確認", F(size=8), X.CENTER)
        X.put(ws, r, 9, city_person(c.fy), F(size=9), X.CENTER)
        X.box(ws, r, 1, r, 9)
        X.put(ws, r + 1, 1, "汐見重工環境株式会社 くぬぎ台作業所", F(size=8), X.LEFT)
        ws.merge_cells(start_row=r + 1, start_column=1, end_row=r + 1, end_column=5)
        X.page(ws, "landscape", "A4", 1, 1, margins=(0.5, 0.5, 0.5, 0.5))
    wb.save(path)
    rec.save()


# ---------------------------------------------------------------- 直営
ROUT = {"pump_s": ["定期 点検（絶縁 {m}MΩ）、異常なし", "グランド 増締め、滴下 止まる", "軸受 給脂、振動 {v}mm/s"],
        "pump_m": ["振動測定 {v}mm/s、異常なし", "軸受 給脂", "グランド 点検"],
        "conv": ["チェーン 給油・張り 点検", "スクリュー 異音 点検、異常なし", "落下物 清掃"],
        "fan": ["Vベルト 張り 点検", "軸受 給脂、温度 {t}℃", "振動 {v}mm/s"],
        "comp": ["ドレン 排出・フィルタ 清掃", "運転時間 確認、油量 点検"],
        "misc": ["外観点検、異常なし", "給油・清掃"], "inst_a": ["指示 確認・電極 洗浄", "ゼロ点 確認"],
        "hyd": ["作動油 量・温度 点検", "フィルタ 差圧 確認"], "feeder": ["切出し 状況 点検、詰まりなし", "給油"]}
WORKT = {"first": ["分解、{w}（部品確認）", "点検・不具合箇所 確認、{w} 準備", "取外し・型取り（{w}）"],
         "mid": ["{w}", "{w}、溶接・加工"], "last": ["組立・試運転、{w} 完了", "{w}、漏れなし 確認"]}


def direct_text(r, rng):
    if r.get("text"):
        return r["text"]
    if r["phase"] == "routine":
        return rng.choice(ROUT.get(r["cls"], ROUT["misc"])).format(m=rng.choice([20, 50, 100, 200]), v=round(rng.uniform(0.8, 2.6), 1),
                                                                     t=rng.randint(35, 58))
    w = r["what"] or "補修"
    return rng.choice(WORKT[r["phase"] if r["phase"] in WORKT else "mid"]).format(w=w)


def d2_weekly(fy, model, path, fname, parts_by_wid):
    import random
    from world import SEED
    rng = random.Random("%s-d2-%d" % (SEED, fy))
    rec = Rec(fname)
    wb = Workbook()
    wb.remove(wb.active)
    X.props(wb, "【架空】くぬぎ台環境センター 直営補修作業報告（週報）%s" % wareki(fy), "くぬぎ台市 環境センター 直営補修班 d2",
            "作業日報", "d2 様式（H29〜）1 週 1 枚")
    rows = [r for r in model.direct if (r["date"].year if r["date"].month >= 4 else r["date"].year - 1) == fy]
    weeks = {}
    for r in rows:
        mon = r["date"] - dt.timedelta(days=r["date"].weekday())
        weeks.setdefault(mon, []).append(r)
    for mon in sorted(weeks):
        lst = sorted(weeks[mon], key=lambda r: (r["date"], r["wid"] is None))
        sat = mon + dt.timedelta(days=5)
        sh = "%d.%d-%d.%d" % (mon.month, mon.day, sat.month, sat.day)
        ws = wb.create_sheet(sh)
        X.widths(ws, [11, 9, 22, 54, 5, 6, 10])
        X.put(ws, 1, 1, "【架空】直営補修 作業報告（週報）", F(size=14, bold=True), X.NOWRAP)
        X.put(ws, 1, 6, "くぬぎ台環境センター", F(size=9), X.RIGHT)
        ws.merge_cells("F1:G1")
        X.put(ws, 2, 1, "期間：%s〜%d月%d日（土）" % (_wa(mon), sat.month, sat.day), F(size=10), X.LEFT)
        ws.merge_cells("A2:D2")
        X.put(ws, 2, 6, "班長 d2", F(size=10), X.RIGHT)
        ws.merge_cells("F2:G2")
        X.put(ws, 3, 1, "班員：t1・t2・t3・t4（交替勤務明けの者を除く）", F(size=9), X.LEFT)
        ws.merge_cells("A3:D3")
        for i, h in enumerate(["月日", "タグNo", "機器", "作業内容・結果", "人数", "時間", "部品・控No"], 1):
            X.put(ws, 5, i, h, F(bold=True, size=9), X.CENTER, X.fill("E2EFDA"))
        r = 6
        prev = None
        tot = 0
        for it in lst[:14]:
            e = model.tag[it["tag"]]
            d = it["date"]
            txt = direct_text(it, rng)
            dtxt = "%d/%d（%s）" % (d.month, d.day, WD[d.weekday()])
            X.put(ws, r, 1, dtxt if d != prev else "", F(size=9), X.CENTER)
            prev = d
            X.put(ws, r, 2, it["tag"] or None, F(size=8), X.CENTER)
            nm = it.get("eqname") or (("%d号炉 " % e.unit if e.unit else "") + e.name)
            if not it.get("showtag", True):
                it = dict(it, tag="")
            X.put(ws, r, 3, nm, F(size=9), X.LEFT)
            X.put(ws, r, 4, txt, F(size=9), X.LEFT)
            X.put(ws, r, 5, it["n"], F(size=9), X.CENTER)
            X.put(ws, r, 6, it["h"], F(size=9), X.CENTER, nf="0.0")
            pn = parts_by_wid.get(it["wid"], "") if it["phase"] == "first" else ""
            X.put(ws, r, 7, pn, F(size=8), X.CENTER)
            ws.row_dimensions[r].height = 30
            tot += it["n"] * it["h"]
            rec.add(sh, r, "item", {"月日": d.isoformat(), "タグNo": it["tag"] or None, "機器": nm, "作業内容": txt, "人数": it["n"], "時間": it["h"],
                                    "manDays": round(it["n"] * it["h"] / 8.0, 2), "manDaysText": "%s人×%s時間" % (it["n"], it["h"]), "部品控No": pn},
                    works=[it["wid"]])
            r += 1
        while r < 16:
            ws.row_dimensions[r].height = 22
            r += 1
        X.put(ws, r, 4, "延べ時間（人×時間）", F(size=9), X.RIGHT)
        X.put(ws, r, 6, tot, F(size=9, bold=True), X.CENTER, nf="0.0")
        rec.add(sh, r, "total", {"延べ時間": tot})
        X.box(ws, 5, 1, r, 7, side=X.med, inner=X.thin)
        r += 1
        X.put(ws, r, 1, "所見・連絡", F(size=9), X.CENTER, X.fill("F2F2F2"))
        X.put(ws, r, 2, "", F(size=9), X.LEFT_TOP)
        ws.merge_cells(start_row=r, start_column=2, end_row=r + 1, end_column=7)
        ws.merge_cells(start_row=r, start_column=1, end_row=r + 1, end_column=1)
        X.box(ws, r, 1, r + 1, 7)
        X.put(ws, r + 3, 5, "施設係", F(size=8), X.CENTER)
        X.put(ws, r + 3, 6, city_person(fy), F(size=9), X.CENTER)
        X.put(ws, r + 3, 7, "班長 d2", F(size=9), X.CENTER)
        X.box(ws, r + 3, 5, r + 3, 7)
        X.page(ws, "landscape", "A4", 1, 1)
    wb.save(path)
    rec.save()
