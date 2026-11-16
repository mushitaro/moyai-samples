# -*- coding: utf-8 -*-
"""作業日報（協力会社 野分工業）。
2016：鷹羽指定の横型（午前・午後の○、工数欄なし）。2017〜2024：w1 の 1日1枚 縦型（2024 は w2 がそのまま使用）。
2025〜：w2 の 1週1枚 横型（週報式）。工程（gen_contract）の実績日から作る。"""
import datetime as dt
import random
from openpyxl import Workbook
from openpyxl.utils import get_column_letter as L
import world as W
import works as WK
import contracts as K
import gen_contract as G
from daily_src import NOTE_2024, NOTE_2025, KY, WEATHER
from xlhelp import *
from rec import Rec, path


def seed_of(c):
    return {(2024, 1): 24, (2025, 2): 25}.get((c["y"], c["f"]), c["y"] * 10 + c["f"])


def fname(c):
    sty = "2016様式" if c["y"] == 2016 else ("1日1枚" if c["y"] <= 2024 else "週1枚")
    return f"【架空】ひばり野クリーンセンター_作業日報_{c['y']}年度{c['f']}号炉定期整備工事_{sty}.xlsx"


def works_of(t, y):
    return [w for w in (WK.wid(W.eq(nm, f), y) for nm, f in t[7]) if w]


def day_notes(c, rows):
    """日付 → 特記。見本の 2 契約は手書きの特記。それ以外は工程から（入炉許可・立会・遅れ・発見）。"""
    y, f = c["y"], c["f"]
    if (y, f) == (2024, 1):
        return dict(NOTE_2024)
    if (y, f) == (2025, 2):
        return dict(NOTE_2025)
    notes = {}
    days = sorted(rows)
    work_days = [x for x in days if any(l[0][3] != "SPC運転" for l in rows[x])]
    if work_days:
        notes[work_days[0]] = f"SPC 運転課：{f}号炉 冷却中。縁切り板 挿入 SPC 立会"
    for t in c["tasks"]:
        ds = [x for x in K.task_days(t, y) if x.weekday() != 6]
        if not ds:
            continue
        if "UT" in t[2]:
            notes[ds[-1]] = (notes.get(ds[-1], "") + " " + G.FINDINGS.get((y, f), f"UT 水管 最小 {G.boiler_min(f, y):.1f}mm 良")).strip()
        if "定期事業者検査" in t[2]:
            notes[ds[0]] = (notes.get(ds[0], "") + " ボイラ定期事業者検査（開放）実施、指摘なし").strip()
        if "水圧試験" in t[2]:
            notes[ds[0]] = (notes.get(ds[0], "") + " ボイラ水圧試験 4.5MPa 30分 漏れなし（SPC・検査員 立会）").strip()
        if len(t) > 10 and t[10]:
            notes[ds[-1]] = (notes.get(ds[-1], "") + f" {t[2].split(' ')[0]}：{t[10]}").strip()
        if t[2].startswith("負荷運転"):
            notes[ds[-1]] = (notes.get(ds[-1], "") + " 定格負荷 確認、SPC へ引渡し").strip()
    if (y, f) not in ((2016, 1), (2016, 2)) and G.FINDINGS.get((y, f)) and not any("UT" in t[2] for t in c["tasks"]):
        notes[work_days[len(work_days) // 2]] = G.FINDINGS[(y, f)]
    return notes


def place_of(t, f):
    if t[1] == "共通" or t[1].startswith("共通"):
        return f"{f}炉 全般" if t[1] == "共通" else t[1].replace("共通（", "").replace("設備）", "")
    p = f"{f}炉 " + t[1].replace("設備", "")
    if "給水ポンプ" in t[2]:
        p = f"{f}炉 給水ポンプ室"
    elif "送風機" in t[2] or "通風機" in t[2]:
        p = f"{f}炉 通風機室"
    elif t[3] == "電気班":
        p = f"{f}炉 電気室"
    return p


# ---------------------------------------------------------------- 2016 様式
def build_2016(c):
    fn = fname(c)
    R = Rec(fn)
    y, f = c["y"], c["f"]
    rows = K.daily_rows(c["tasks"], y, seed_add=seed_of(c))
    notes = day_notes(c, rows)
    r_ = random.Random(W.SEED + y * 7 + f)
    wb = Workbook()
    wb.remove(wb.active)
    props(wb, f"【架空】ひばり野クリーンセンター {c['title']} 作業日報", "野分工業 職長 w1", dt.datetime(y, 12, 1, 9, 0))
    for x, lines in rows.items():
        work = [l for l in lines if l[0][3] != "SPC運転"]
        if not work:
            continue
        sh = f"{x.month}-{x.day}"
        ws = wb.create_sheet(sh)
        for col, w in {1: 14, 2: 46, 3: 6, 4: 6, 5: 6, 6: 30}.items():
            ws.column_dimensions[L(col)].width = w
        ws.merge_cells("A1:F1")
        put(ws, 1, 1, "【架空】作業日報", font(15, True), CENTER)
        put(ws, 2, 1, "工事名", font(9, True), CENTER, B_ALL, fill("EDEDED"))
        ws.merge_cells("B2:C2")
        put(ws, 2, 2, c["title"], font(9), LEFT, B_ALL)
        put(ws, 2, 4, "日付", font(9, True), CENTER, B_ALL, fill("EDEDED"))
        ws.merge_cells("E2:F2")
        put(ws, 2, 5, f"平成{y - 1988}年{x.month}月{x.day}日（{K.WD[x.weekday()]}）", font(9), LEFT, B_ALL)
        wea = r_.choice(WEATHER)
        put(ws, 3, 1, "天候", font(9, True), CENTER, B_ALL, fill("EDEDED"))
        put(ws, 3, 2, wea, font(9), LEFT, B_ALL)
        put(ws, 3, 4, "記入", font(9, True), CENTER, B_ALL, fill("EDEDED"))
        ws.merge_cells("E3:F3")
        put(ws, 3, 5, "野分工業㈱ 梶本", font(9), LEFT, B_ALL)
        for i, h in enumerate(["作業班", "作業内容", "午前", "午後", "人員", "備考"], 1):
            put(ws, 5, i, h, font(9, True), CENTER, B_ALL, fill("DDEBF7"))
        r = 6
        tm = 0
        for (t, desc, men, hrs) in work:
            pm = "○"
            am = "○" if hrs >= 8 else ""
            put(ws, r, 1, t[3], font(9), CENTER, B_ALL)
            put(ws, r, 2, desc, font(9), LEFT, B_ALL)
            put(ws, r, 3, am or None, font(9), CENTER, B_ALL)
            put(ws, r, 4, pm, font(9), CENTER, B_ALL)
            put(ws, r, 5, men, font(9), CENTER, B_ALL)
            put(ws, r, 6, "残業 1.5h" if hrs > 8 else None, font(8), LEFT, B_ALL)
            R.add(sh, r, "item", {"日付": x.isoformat(), "天候": wea, "作業班": t[3], "作業内容": desc,
                                  "午前": am or None, "午後": pm, "人員": men, "備考": "残業 1.5h" if hrs > 8 else None,
                                  "manDays": round(men * hrs / 8, 2), "manDaysText": f"{men}人", "特記": notes.get(x)},
                  works_of(t, y))
            tm += men
            r += 1
        put(ws, r, 1, "計", font(9, True), CENTER, B_ALL, fill("F2F2F2"))
        put(ws, r, 5, tm, font(9, True), CENTER, B_ALL, fill("F2F2F2"))
        R.add(sh, r, "total", {"日付": x.isoformat(), "人員": tm})
        r += 2
        put(ws, r, 1, "連絡事項", font(9, True), CENTER, B_ALL, fill("EDEDED"))
        ws.merge_cells(start_row=r, start_column=2, end_row=r + 1, end_column=6)
        put(ws, r, 2, notes.get(x, ""), font(9), LEFT_TOP, B_ALL)
        box(ws, r, 1, r + 1, 6)
        page(ws, landscape=True, paper=9, fit_w=1, fit_h=1, area=f"A1:F{r + 1}")
    wb.save(path(fn))
    return fn


# ---------------------------------------------------------------- 2017〜2024 w1 様式（1日1枚）
def build_daily(c):
    fn = fname(c)
    R = Rec(fn)
    y, f = c["y"], c["f"]
    w2 = y >= 2024
    name = "森野" if w2 else "梶本"
    rows = K.daily_rows(c["tasks"], y, seed_add=seed_of(c))
    notes = day_notes(c, rows)
    r_ = random.Random(W.SEED + (241 if (y, f) == (2024, 1) else y * 7 + f))
    wb = Workbook()
    wb.remove(wb.active)
    last = max(rows)
    props(wb, f"【架空】ひばり野クリーンセンター {c['title']} 作業日報", f"野分工業 職長 {'w2' if w2 else 'w1'}",
          dt.datetime(last.year, last.month, last.day, 17, 0) + dt.timedelta(days=1))
    spc = "小鳥遊" if y >= 2022 else "早瀬"
    for x, lines in rows.items():
        work = [l for l in lines if l[0][3] != "SPC運転"]
        if not work:
            continue
        sh = f"{x.month}.{x.day:02d}"
        ws = wb.create_sheet(sh)
        for col, w in {1: 11, 2: 14, 3: 38, 4: 5, 5: 6, 6: 7, 7: 7}.items():
            ws.column_dimensions[L(col)].width = w
        ws.merge_cells("A1:G1")
        put(ws, 1, 1, "【架空】　作　業　日　報", font(16, True), CENTER)
        put(ws, 3, 1, "工事件名", font(9, True), CENTER, B_ALL, fill("F2F2F2"))
        ws.merge_cells("B3:C3")
        put(ws, 3, 2, c["title"], font(9), LEFT, B_ALL)
        put(ws, 3, 4, "日付", font(9, True), CENTER, B_ALL, fill("F2F2F2"))
        ws.merge_cells("E3:G3")
        put(ws, 3, 5, f"{x.year}年{x.month}月{x.day}日（{K.WD[x.weekday()]}）", font(9), CENTER, B_ALL)
        put(ws, 4, 1, "施設", font(9, True), CENTER, B_ALL, fill("F2F2F2"))
        ws.merge_cells("B4:C4")
        put(ws, 4, 2, f"【架空】ひばり野クリーンセンター　{f}号炉", font(9), LEFT, B_ALL)
        put(ws, 4, 4, "天候", font(9, True), CENTER, B_ALL, fill("F2F2F2"))
        ws.merge_cells("E4:G4")
        wea = r_.choice(WEATHER)
        put(ws, 4, 5, wea, font(9), CENTER, B_ALL)
        put(ws, 5, 1, "作業時間", font(9, True), CENTER, B_ALL, fill("F2F2F2"))
        ws.merge_cells("B5:C5")
        long_ = any(l[3] > 8 for l in work)
        put(ws, 5, 2, "8:00 〜 17:00" + ("（一部 18:30 まで）" if long_ else ""), font(9), LEFT, B_ALL)
        put(ws, 5, 4, "職長", font(9, True), CENTER, B_ALL, fill("F2F2F2"))
        ws.merge_cells("E5:G5")
        put(ws, 5, 5, name, font(9), CENTER, B_ALL)
        for i, h in enumerate(["班", "作業場所", "作業内容", "人員", "時間", "工数", "進捗"], 1):
            put(ws, 7, i, h, font(9, True), CENTER, B_ALL, fill("DDEBF7"))
        r = 8
        prevplace = None
        tm = th = 0
        nt = notes.get(x, "")
        for (t, desc, men, hrs) in work:
            place = place_of(t, f)
            wds = [d for d in K.task_days(t, y) if d.weekday() != 6]
            k = wds.index(x) + 1 if x in wds else 0
            prog = round(100 * k / len(wds)) if wds else 0
            pv = "〃" if place == prevplace else place
            put(ws, r, 1, t[3], font(9), CENTER, B_ALL)
            put(ws, r, 2, pv, font(9), CENTER if pv == "〃" else LEFT, B_ALL)
            prevplace = place
            put(ws, r, 3, desc, font(9), LEFT, B_ALL)
            put(ws, r, 4, men, font(9), CENTER, B_ALL)
            put(ws, r, 5, hrs, font(9), CENTER, B_ALL, None, "0.0")
            put(ws, r, 6, round(men * hrs, 1), font(9), RIGHT, B_ALL, None, "0.0")
            pg = f"{prog}%" if prog < 100 else "完"
            put(ws, r, 7, pg, font(9), CENTER, B_ALL)
            ws.row_dimensions[r].height = 26 if len(desc) > 22 else 18
            R.add(sh, r, "item", {"日付": x.isoformat(), "天候": wea, "班": t[3], "作業場所": pv, "作業内容": desc, "人員": men,
                                  "時間": hrs, "工数": round(men * hrs, 1), "進捗": pg, "manDays": round(men * hrs / 8, 2),
                                  "manDaysText": f"{men * hrs:.1f}", "特記": nt or None}, works_of(t, y))
            tm += men
            th += men * hrs
            r += 1
        while r < 20:
            box(ws, r, 1, r, 7)
            r += 1
        put(ws, r, 1, "合　計", font(9, True), CENTER, B_ALL, fill("F2F2F2"))
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        put(ws, r, 4, tm, font(9, True), CENTER, B_ALL)
        put(ws, r, 6, round(th, 1), font(9, True), RIGHT, B_ALL, None, "0.0")
        box(ws, r, 1, r, 7)
        R.add(sh, r, "total", {"日付": x.isoformat(), "人員": tm, "工数": round(th, 1)})
        r += 2
        put(ws, r, 1, "特記事項", font(9, True), CENTER, B_ALL, fill("F2F2F2"))
        ws.merge_cells(start_row=r, start_column=2, end_row=r + 2, end_column=7)
        put(ws, r, 2, nt, font(9), LEFT_TOP, B_ALL)
        ws.merge_cells(start_row=r, start_column=1, end_row=r + 2, end_column=1)
        box(ws, r, 1, r + 2, 7)
        r += 3
        put(ws, r, 1, "安全(KY)", font(9, True), CENTER, B_ALL, fill("F2F2F2"))
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
        put(ws, r, 2, KY[x.toordinal() % len(KY)], font(9), LEFT, B_ALL)
        box(ws, r, 1, r, 7)
        r += 2
        for j, h in enumerate(("職長", "SPC確認")):
            put(ws, r, 5 + j, h, font(8), CENTER, B_ALL)
            put(ws, r + 1, 5 + j, name if h == "職長" else spc, font(9, color="C00000"), CENTER, B_ALL)
        ws.row_dimensions[r + 1].height = 30
        put(ws, r, 1, "野分工業株式会社", font(9))
        page(ws, landscape=False, paper=9, fit_w=1, fit_h=1, area=f"A1:G{r + 1}", footer="野分工業(株)　梶本")
    wb.save(path(fn))
    return fn


# ---------------------------------------------------------------- 2025〜 週報式
def build_weekly(c):
    fn = fname(c)
    R = Rec(fn)
    y, f = c["y"], c["f"]
    rows = K.daily_rows(c["tasks"], y, seed_add=seed_of(c))
    notes = day_notes(c, rows)
    r_ = random.Random(W.SEED + (251 if (y, f) == (2025, 2) else y * 7 + f))
    wb = Workbook()
    wb.remove(wb.active)
    last = max(rows)
    props(wb, f"【架空】ひばり野クリーンセンター {c['title']} 作業日報（週報式）", "野分工業 職長 w2",
          dt.datetime(last.year, last.month, last.day, 17, 30) + dt.timedelta(days=6))
    first = min(rows)
    w0 = first - dt.timedelta(days=first.weekday())
    weeks = []
    while w0 <= last:
        weeks.append(w0)
        w0 += dt.timedelta(days=7)
    wk = 0
    for w0 in weeks:
        days = [w0 + dt.timedelta(days=i) for i in range(6)]
        active = [t for t in c["tasks"] if t[3] != "SPC運転" and any(d in days for d in K.task_days(t, y))]
        if not active:
            continue
        wk += 1
        sh = f"第{wk}週"
        ws = wb.create_sheet(sh)
        for col, w in {1: 4, 2: 10, 3: 24, 4: 8}.items():
            ws.column_dimensions[L(col)].width = w
        C0 = 5
        for i in range(6):
            ws.column_dimensions[L(C0 + 2 * i)].width = 3.6
            ws.column_dimensions[L(C0 + 2 * i + 1)].width = 15
        CS = C0 + 12
        ws.column_dimensions[L(CS)].width = 7
        ws.column_dimensions[L(CS + 1)].width = 6
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=CS + 1)
        put(ws, 1, 1, "【架空】　作 業 日 報（週）", font(16, True), CENTER)
        ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=2)
        put(ws, 3, 1, "工事件名", font(9, True), CENTER, B_ALL, fill("F2F2F2"))
        ws.merge_cells(start_row=3, start_column=3, end_row=3, end_column=8)
        put(ws, 3, 3, c["title"], font(9), LEFT, B_ALL)
        put(ws, 3, 9, "期間", font(9, True), CENTER, B_ALL, fill("F2F2F2"))
        ws.merge_cells(start_row=3, start_column=10, end_row=3, end_column=12)
        put(ws, 3, 10, f"{days[0]:%Y/%m/%d}（月）〜{days[-1]:%m/%d}（土）", font(9), CENTER, B_ALL)
        put(ws, 3, 13, "職長", font(9, True), CENTER, B_ALL, fill("F2F2F2"))
        put(ws, 3, 14, "森野", font(9), CENTER, B_ALL)
        ws.merge_cells(start_row=4, start_column=1, end_row=4, end_column=2)
        put(ws, 4, 1, "施設", font(9, True), CENTER, B_ALL, fill("F2F2F2"))
        ws.merge_cells(start_row=4, start_column=3, end_row=4, end_column=8)
        extra = "（タービン・空冷復水器を含む）" if any(t[3] == "タービン班" for t in c["tasks"]) else ""
        put(ws, 4, 3, f"【架空】ひばり野クリーンセンター　{f}号炉{extra}", font(9), LEFT, B_ALL)
        box(ws, 3, 1, 4, 8)
        Rr = 6
        for col, h in ((1, "No"), (2, "設備"), (3, "作業項目"), (4, "班")):
            ws.merge_cells(start_row=Rr, start_column=col, end_row=Rr + 1, end_column=col)
            put(ws, Rr, col, h, font(9, True), CENTER, B_ALL, fill("E2EFDA"))
        for i, x in enumerate(days):
            col = C0 + 2 * i
            ws.merge_cells(start_row=Rr, start_column=col, end_row=Rr, end_column=col + 1)
            hol = x in (dt.date(2025, 10, 13), dt.date(2026, 5, 4), dt.date(2026, 5, 5), dt.date(2026, 5, 6), dt.date(2026, 10, 12))
            put(ws, Rr, col, f"{x.month}/{x.day}（{K.WD[x.weekday()]}）" + ("祝" if hol else ""), font(9, True), CENTER, B_ALL, fill("E2EFDA"))
            put(ws, Rr + 1, col, "人", font(8, True), CENTER, B_ALL, fill("E2EFDA"))
            put(ws, Rr + 1, col + 1, "作業内容", font(8, True), CENTER, B_ALL, fill("E2EFDA"))
        ws.merge_cells(start_row=Rr, start_column=CS, end_row=Rr + 1, end_column=CS)
        put(ws, Rr, CS, "週工数\n(人h)", font(8, True), CENTER, B_ALL, fill("E2EFDA"))
        ws.merge_cells(start_row=Rr, start_column=CS + 1, end_row=Rr + 1, end_column=CS + 1)
        put(ws, Rr, CS + 1, "進捗", font(8, True), CENTER, B_ALL, fill("E2EFDA"))
        r = Rr + 2
        daymen = {x: 0 for x in days}
        dayh = {x: 0.0 for x in days}
        for n, t in enumerate(active, 1):
            put(ws, r, 1, n, font(8), CENTER, B_ALL)
            put(ws, r, 2, t[1].replace("設備", ""), font(8), LEFT, B_ALL)
            put(ws, r, 3, t[2], font(8), LEFT, B_ALL)
            put(ws, r, 4, t[3], font(8), CENTER, B_ALL)
            wh = 0.0
            dd = {}
            for i, x in enumerate(days):
                col = C0 + 2 * i
                hit = [l for l in rows.get(x, []) if l[0][0] == t[0]]
                if hit:
                    _, desc, men, hrs = hit[0]
                    txt = desc + ("（残業）" if hrs > 8 else "")
                    put(ws, r, col, men, font(8), CENTER, B_ALL)
                    put(ws, r, col + 1, txt, font(8), LEFT, B_ALL)
                    wh += men * hrs
                    daymen[x] += men
                    dayh[x] += men * hrs
                    dd[x.isoformat()] = {"人": men, "作業内容": txt}
                else:
                    put(ws, r, col, None, None, None, B_ALL)
                    put(ws, r, col + 1, None, None, None, B_ALL, fill("F7F7F7"))
            td = [d for d in K.task_days(t, y) if d.weekday() != 6]
            done = sum(1 for d in td if d <= days[-1])
            prog = round(100 * done / len(td)) if td else 0
            pg = "完" if prog >= 100 else f"{prog}%"
            put(ws, r, CS, round(wh, 1), font(8), RIGHT, B_ALL, None, "0.0")
            put(ws, r, CS + 1, pg, font(8), CENTER, B_ALL)
            ws.row_dimensions[r].height = 22
            R.add(sh, r, "item", {"週": f"{days[0]}〜{days[-1]}", "設備": t[1].replace("設備", ""), "作業項目": t[2], "班": t[3],
                                  "days": dd, "週工数": round(wh, 1), "進捗": pg, "manDays": round(wh / 8, 2),
                                  "manDaysText": f"{wh:.1f}"}, works_of(t, y))
            r += 1
        for lab, vals, nf in (("日計 人員", daymen, "0"), ("日計 工数(人h)", dayh, "0.0")):
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
            put(ws, r, 1, lab, font(8, True), CENTER, B_ALL, fill("F2F2F2"))
            for i, x in enumerate(days):
                col = C0 + 2 * i
                ws.merge_cells(start_row=r, start_column=col, end_row=r, end_column=col + 1)
                put(ws, r, col, round(vals[x], 1) if vals[x] else None, font(8, True), CENTER, B_ALL, fill("F2F2F2"), nf)
            put(ws, r, CS, round(sum(dayh.values()), 1) if "工数" in lab else None, font(8, True), RIGHT, B_ALL, fill("F2F2F2"), "0.0")
            put(ws, r, CS + 1, None, None, None, B_ALL, fill("F2F2F2"))
            R.add(sh, r, "subtotal", {"区分": lab, "days": {x.isoformat(): round(vals[x], 1) for x in days if vals[x]},
                                       **({"週工数": round(sum(dayh.values()), 1)} if "工数" in lab else {})})
            r += 1
        for lab in ("天候", "特記", "安全(KY)", "SPC確認"):
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
            put(ws, r, 1, lab, font(8, True), CENTER, B_ALL, fill("F2F2F2"))
            dd = {}
            for i, x in enumerate(days):
                col = C0 + 2 * i
                ws.merge_cells(start_row=r, start_column=col, end_row=r, end_column=col + 1)
                val = None
                if lab == "天候":
                    val = r_.choice(WEATHER) if daymen[x] else None
                elif lab == "特記":
                    val = notes.get(x)
                elif lab == "安全(KY)":
                    val = KY[x.toordinal() % len(KY)] if daymen[x] else None
                elif lab == "SPC確認":
                    val = "小鳥遊" if daymen[x] else None
                put(ws, r, col, val, font(8, color="C00000" if lab == "SPC確認" else None), LEFT_TOP if lab == "特記" else CENTER, B_ALL)
                if val:
                    dd[x.isoformat()] = val
            ws.merge_cells(start_row=r, start_column=CS, end_row=r, end_column=CS + 1)
            ws.row_dimensions[r].height = 36 if lab == "特記" else 16
            if dd:
                R.add(sh, r, "item", {"区分": lab, "days": dd})
            r += 1
        put(ws, r + 1, 1, "野分工業株式会社　※毎日 当日の欄を記入し、土曜に SPC へ提出（2025年度から週1枚に変更）", font(8))
        page(ws, landscape=True, paper=8, fit_w=1, fit_h=1, area=f"A1:{L(CS + 1)}{r + 1}")
    wb.save(path(fn))
    return fn


def build(c):
    if c["y"] == 2016:
        return build_2016(c)
    if c["y"] <= 2024:
        return build_daily(c)
    return build_weekly(c)
