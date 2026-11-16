# -*- coding: utf-8 -*-
"""作業日報（全年度）
H16〜H17 m1 班別 1 枚 ／ H18〜H25 m1 工程×班 の表（1 日 1 枚）／ H26〜R2 m2 縦型（R2 m3 が検温欄追加）／ R3〜 m3 縦型拡張
タービン検査は協力会社（東雲）の作業週報
"""
import datetime as dt
import os
from collections import defaultdict

from openpyxl import Workbook
from openpyxl.styles import Border, Font
from openpyxl.utils import get_column_letter as CL

import ans
import core
import texts
from xlhelp import *

D = dt.date.fromisoformat
WEATHER = ["晴", "晴", "曇", "晴", "曇時々晴", "雨", "曇", "晴", "晴のち曇", "小雨", "晴", "曇"]
KY = ["高所作業時の安全帯（フルハーネス）使用徹底", "入炉前の酸素・CO 濃度測定", "重量物の玉掛け合図の統一", "火気使用時の消火器配置",
      "グラインダ使用時の保護具", "足場上の工具落下防止", "槽内作業の監視人配置", "停電・隔離操作の施錠（ロックアウト）確認",
      "熱中症・脱水に注意（水分補給）", "通路の資材仮置き禁止"]


def weather(d):
    return WEATHER[(d.toordinal() * 7 + 3) % len(WEATHER)]


def tend(h):
    return {4.0: "12:00", 5.0: "14:00", 6.0: "15:00", 7.0: "16:00", 8.0: "17:00", 9.0: "18:00", 9.5: "18:30"}[h]


def style_of(fy):
    return "m1" if fy <= 2013 else ("m2" if fy <= 2020 else "m3")


def form_of(fy):
    return "班別1枚様式" if fy <= 2005 else ("工程班表様式" if fy <= 2013 else ("H26縦型様式" if fy <= 2020 else "R3縦型拡張様式"))


def lines_by_day(win):
    W = core.get()
    tids = {t["id"] for t in win["tasks"]}
    by = defaultdict(list)
    for ln in W.lines:
        if ln["task"]["id"] in tids and ln["date"] <= core.TODAY:
            by[ln["date"]].append(ln)
    return by


def fname(win, form):
    W = core.get()
    c = W.contracts[win["contract"]]
    short = {"TEISEI": "定期整備業務委託", "ZENRO": "全炉停止時共通設備整備業務委託", "TURB": "蒸気タービン定期事業者検査",
             "KOSHIN": "更新工事", "HOSHU": "補修工事", "KIKAN": "基幹的設備改良工事"}[c["kind"]]
    part = win["label"].replace(" ", "").replace("定期整備", "").replace("更新工事", "").replace("補修工事", "")
    if c["kind"] == "KIKAN":
        part = ""
    return f"【架空】みなと臨海工場_作業日報_{core.wa_fy(win['fy'])}_{short}{part}_{form}.xlsx"


def cname_of(win):
    W = core.get()
    c = W.contracts[win["contract"]]
    return c["name"], win["label"]


# ---------------------------------------------------------------- A：班別 1 枚（H16〜H17）
def build_a(win):
    by = lines_by_day(win)
    if not by:
        return None
    fn = fname(win, form_of(win["fy"]))
    ans.start(fn, False, "作業日報", win["fy"], "班別1枚（m1 H16）", "m1（各班 職長）", win["contract"])
    wb = Workbook()
    wb.remove(wb.active)
    F = lambda **k: font(FONT_G, **k)
    cn, lab = cname_of(win)
    for d in sorted(by):
        teams = defaultdict(list)
        for ln in by[d]:
            teams[ln["team"]["key"]].append(ln)
        for tk, lns in teams.items():
            tm = lns[0]["team"]
            ws = wb.create_sheet(f"{d.month:02d}{d.day:02d}{tk}")
            for c, w in zip(range(1, 5), [10, 44, 8, 10]):
                ws.column_dimensions[CL(c)].width = w
            merge(ws, 1, 1, 1, 4, "作業日報（班）", F(size=14, bold=True), AL_C, None)
            put(ws, 2, 1, "件名", F(size=9, bold=True), AL_C, B_ALL)
            merge(ws, 2, 2, 2, 4, f"{cn}（{lab}）", F(size=9), AL_L, B_ALL)
            put(ws, 3, 1, "日付", F(size=9, bold=True), AL_C, B_ALL)
            put(ws, 3, 2, f"{d_wa(d)}（{WD[d.weekday()]}）　天候 {weather(d)}", F(size=9), AL_L, B_ALL)
            put(ws, 3, 3, "班", F(size=9, bold=True), AL_C, B_ALL)
            put(ws, 3, 4, tm["name"], F(size=9), AL_C, B_ALL)
            put(ws, 5, 1, "No.", F(size=9, bold=True), AL_C, B_ALL)
            put(ws, 5, 2, "作業内容", F(size=9, bold=True), AL_C, B_ALL)
            put(ws, 5, 3, "人員", F(size=9, bold=True), AL_C, B_ALL)
            put(ws, 5, 4, "時間", F(size=9, bold=True), AL_C, B_ALL)
            r = 6
            for k, ln in enumerate(lns, start=1):
                txt = texts.line_text(ln, "m1")
                vals = [k, f"{ln['task']['name'].split('（')[0]}　{txt}", ln["persons"], f"8:00-{tend(ln['hours'])}"]
                for c, v in enumerate(vals, start=1):
                    put(ws, r, c, v, F(size=9), AL_LW if c == 2 else AL_C, B_ALL)
                ws.row_dimensions[r].height = 28
                ans.rec(fn, ws.title, r, "item", dict(date=d.isoformat(), team=tm["name"], foreman=tm["foreman"], text=vals[1],
                                                       persons=ln["persons"], time=vals[3], manDays=round(ln["persons"] * ln["hours"] / 8, 2),
                                                       manDaysText=None), works=[w["id"] for w in ln["works"]] or None, check=[txt[:10]])
                r += 1
            while r < 12:
                for c in range(1, 5):
                    put(ws, r, c, None, None, None, B_ALL)
                r += 1
            put(ws, r + 1, 1, "職長", F(size=9), AL_C, B_ALL)
            put(ws, r + 1, 2, tm["foreman"], F(size=9), AL_L, B_ALL)
            put(ws, r + 1, 3, "代理人", F(size=9), AL_C, B_ALL)
            put(ws, r + 1, 4, "m1", F(size=9), AL_C, B_ALL)
            put(ws, r + 3, 1, "【架空】汐見重工環境株式会社", F(size=8))
            page(ws, landscape=False, paper=9, fit_w=1, fit_h=1, area=f"A1:D{r + 3}", margins=(0.8, 0.8, 0.8, 0.8))
    props(wb, f"【架空】{cn} 作業日報（{lab}）", "m1", "作業日報", dt.datetime(2004, 5, 1, 8, 0), modified_by="m1")
    wb.save(os.path.join(OUT, fn))
    return fn


# ---------------------------------------------------------------- B：工程×班 の表（H18〜H25）
def build_b(win):
    by = lines_by_day(win)
    if not by:
        return None
    fn = fname(win, form_of(win["fy"]))
    ans.start(fn, False, "作業日報", win["fy"], "工程×班 1日1枚（m1 H18）", core.agent(win["fy"]), win["contract"])
    wb = Workbook()
    wb.remove(wb.active)
    F = lambda **k: font(FONT_G, **k)
    cn, lab = cname_of(win)
    allteams = []
    for t in win["tasks"]:
        if t["team"] not in allteams:
            allteams.append(t["team"])
    for k, d in enumerate(sorted(by), start=1):
        ws = wb.create_sheet(f"{d.month}-{d.day}")
        cols = 3 + len(allteams) + 2
        widths = [7, 26, 40] + [6] * len(allteams) + [6, 9]
        for c, w in enumerate(widths, start=1):
            ws.column_dimensions[CL(c)].width = w
        merge(ws, 1, 1, 1, cols, "作　業　日　報", F(size=16, bold=True), AL_C, None)
        merge(ws, 2, 1, 2, 3, f"件名：{cn}（{lab}）", F(size=9), AL_L, None)
        merge(ws, 2, 4, 2, cols, f"{d_wa_long(d)}（{WD[d.weekday()]}）　天候：{weather(d)}　No.{k}", F(size=9), AL_R, None)
        put(ws, 4, 1, "工程No.", F(size=8, bold=True), AL_C, B_ALL, fill("EEEEEE"))
        put(ws, 4, 2, "作業項目", F(size=8, bold=True), AL_C, B_ALL, fill("EEEEEE"))
        put(ws, 4, 3, "本日の実施内容", F(size=8, bold=True), AL_C, B_ALL, fill("EEEEEE"))
        for i, tk in enumerate(allteams):
            put(ws, 4, 4 + i, core.TEAM_BASE[tk][0].replace("機械班", "機").replace("電気計装班", "電計"), F(size=7, bold=True), AL_C, B_ALL, fill("EEEEEE"))
        put(ws, 4, cols - 1, "計", F(size=8, bold=True), AL_C, B_ALL, fill("EEEEEE"))
        put(ws, 4, cols, "時間", F(size=8, bold=True), AL_C, B_ALL, fill("EEEEEE"))
        ws.row_dimensions[4].height = 30
        r = 5
        colsum = defaultdict(int)
        for ln in sorted(by[d], key=lambda x: x["task"]["no"]):
            t = ln["task"]
            txt = texts.line_text(ln, "m1" if win["fy"] <= 2009 else "m2")
            put(ws, r, 1, t["no"], F(size=8), AL_C, B_ALL)
            put(ws, r, 2, t["name"], F(size=8), AL_LW, B_ALL)
            put(ws, r, 3, txt, F(size=8), AL_LW, B_ALL)
            for i, tk in enumerate(allteams):
                v = ln["persons"] if tk == ln["team"]["key"] else None
                put(ws, r, 4 + i, v, F(size=8), AL_C, B_ALL)
                if v:
                    colsum[tk] += v
            put(ws, r, cols - 1, ln["persons"], F(size=8), AL_C, B_ALL)
            put(ws, r, cols, f"{ln['hours']:.1f}", F(size=8), AL_C, B_ALL)
            ws.row_dimensions[r].height = 26
            ans.rec(fn, ws.title, r, "item", dict(date=d.isoformat(), no=t["no"], name=t["name"], text=txt, team=ln["team"]["name"],
                                                   persons=ln["persons"], hours=ln["hours"], manDays=round(ln["persons"] * ln["hours"] / 8, 2),
                                                   manDaysText=None), works=[w["id"] for w in ln["works"]] or None, check=[t["no"]])
            r += 1
        put(ws, r, 1, "", None, None, B_ALL, fill("F2F2F2"))
        put(ws, r, 2, "合　計", F(size=8, bold=True), AL_C, B_ALL, fill("F2F2F2"))
        put(ws, r, 3, "", None, None, B_ALL, fill("F2F2F2"))
        for i, tk in enumerate(allteams):
            put(ws, r, 4 + i, colsum.get(tk) or None, F(size=8, bold=True), AL_C, B_ALL, fill("F2F2F2"))
        tot = sum(colsum.values())
        put(ws, r, cols - 1, tot, F(size=8, bold=True), AL_C, B_ALL, fill("F2F2F2"))
        put(ws, r, cols, "", None, None, B_ALL, fill("F2F2F2"))
        ans.rec(fn, ws.title, r, "total", dict(date=d.isoformat(), persons=tot), check=[str(tot)])
        r += 2
        eps = [ln["episode"][4] for ln in by[d] if ln.get("episode")]
        merge(ws, r, 1, r + 1, 1, "特記", F(size=8, bold=True), AL_C, B_ALL)
        merge(ws, r, 2, r + 1, cols, "　".join(eps), F(size=8), AL_LT, B_ALL)
        r += 3
        put(ws, r, cols - 1, "代理人", F(size=8), AL_C, B_ALL)
        put(ws, r, cols, core.agent(win["fy"]), F(size=8), AL_C, B_ALL)
        put(ws, r, 1, "【架空】汐見重工環境株式会社", F(size=7))
        page(ws, landscape=True, paper=9, fit_w=1, fit_h=1, area=f"A1:{CL(cols)}{r}")
    props(wb, f"【架空】{cn} 作業日報（{lab}）", "m1", "作業日報", dt.datetime(2006, 4, 3, 8, 0), modified_by=core.agent(win["fy"]))
    wb.save(os.path.join(OUT, fn))
    return fn


# ---------------------------------------------------------------- C：H26 縦型（H26〜R2）
def build_c(win):
    by = lines_by_day(win)
    if not by:
        return None
    fy = win["fy"]
    fn = fname(win, form_of(fy))
    ans.start(fn, False, "作業日報", fy, "H26縦型（m2）", core.agent(fy), win["contract"])
    wb = Workbook()
    wb.remove(wb.active)
    F = lambda **k: font(FONT_G, **k)
    cn, lab = cname_of(win)
    for k, d in enumerate(sorted(by), start=1):
        ws = wb.create_sheet(f"{d.month}.{d.day}")
        for c, w in zip(range(1, 8), [12, 40, 7, 7, 7, 14, 2]):
            ws.column_dimensions[CL(c)].width = w
        merge(ws, 1, 1, 1, 6, "作　業　日　報", F(size=18, bold=True), AL_C, None)
        put(ws, 2, 6, f"No.{k}", F(size=9), AL_R)
        put(ws, 3, 1, "件　名", F(size=9, bold=True), AL_C, B_ALL, fill("EEEEEE"))
        merge(ws, 3, 2, 3, 6, f"{cn}（{lab}）", F(size=10), AL_L, B_ALL)
        put(ws, 4, 1, "日　付", F(size=9, bold=True), AL_C, B_ALL, fill("EEEEEE"))
        put(ws, 4, 2, f"{d_wa_long(d)}（{WD[d.weekday()]}）", F(size=10), AL_L, B_ALL)
        put(ws, 4, 3, "天候", F(size=9, bold=True), AL_C, B_ALL, fill("EEEEEE"))
        put(ws, 4, 4, weather(d), F(size=10), AL_C, B_ALL)
        put(ws, 4, 5, "時間", F(size=9, bold=True), AL_C, B_ALL, fill("EEEEEE"))
        put(ws, 4, 6, "8:00〜17:00", F(size=9), AL_C, B_ALL)
        hr = 7
        if fy >= 2020:
            n_all = sum(x["persons"] for x in by[d]) + 1
            merge(ws, 5, 1, 5, 6, f"検温・体調確認　全員異常なし（{n_all}名　朝礼時）", F(size=8, color="0000CC"), AL_L, B_ALL)
        for c, h in enumerate(["班　名", "作　業　内　容", "人員", "時間", "", "備　考"], start=1):
            put(ws, hr, c, h, F(size=9, bold=True), AL_C, B_ALL, fill("EEEEEE"))
        ws.merge_cells(start_row=hr, start_column=4, end_row=hr, end_column=5)
        r = hr + 1
        prev = None
        for ln in sorted(by[d], key=lambda x: (x["task"]["team"], x["task"]["no"])):
            tm = ln["team"]["name"]
            txt = texts.line_text(ln, "m2")
            body = f"{ln['task']['name'].split('（')[0]}　{txt}"
            put(ws, r, 1, tm if tm != prev else "〃", F(size=9), AL_C, B_ALL)
            prev = tm
            put(ws, r, 2, body, F(size=9), AL_LW, B_ALL)
            put(ws, r, 3, f"{ln['persons']}名", F(size=9), AL_C, B_ALL)
            put(ws, r, 4, f"{ln['hours']:.1f}", F(size=9), AL_C, B_ALL)
            put(ws, r, 5, "h", F(size=8), AL_L, B_ALL)
            put(ws, r, 6, "", F(size=8), AL_L, B_ALL)
            ws.row_dimensions[r].height = 26 if len(body) < 34 else 36
            ans.rec(fn, ws.title, r, "item", dict(date=d.isoformat(), team=tm, text=body, persons=ln["persons"], hours=ln["hours"],
                                                   manDays=round(ln["persons"] * ln["hours"] / 8, 2), manDaysText=f"{ln['persons']}名×{ln['hours']:.1f}h"),
                    works=[w["id"] for w in ln["works"]] or None, check=[body[:12]])
            r += 1
        while r < hr + 15:
            for c in range(1, 7):
                put(ws, r, c, None, None, None, B_ALL)
            r += 1
        tot = sum(x["persons"] for x in by[d])
        put(ws, r, 1, "計", F(size=9, bold=True), AL_C, B_ALL)
        put(ws, r, 2, "", None, None, B_ALL)
        put(ws, r, 3, f"{tot}名", F(size=9, bold=True), AL_C, B_ALL)
        merge(ws, r, 4, r, 6, "（代理人 1 名 含まず）", F(size=8), AL_L, B_ALL)
        ans.rec(fn, ws.title, r, "total", dict(date=d.isoformat(), persons=tot), check=[f"{tot}名"])
        r += 2
        merge(ws, r, 1, r, 1, "危険予知", F(size=9, bold=True), AL_C, B_ALL, fill("EEEEEE"))
        merge(ws, r, 2, r, 6, KY[d.toordinal() % len(KY)], F(size=9), AL_L, B_ALL)
        r += 1
        merge(ws, r, 1, r + 2, 1, "特記事項", F(size=9, bold=True), AL_C, B_ALL, fill("EEEEEE"))
        eps = [ln["episode"][4] for ln in by[d] if ln.get("episode")]
        merge(ws, r, 2, r + 2, 6, "　".join(eps), F(size=9), AL_LT, B_ALL)
        r += 4
        put(ws, r, 4, "代理人", F(size=8), AL_C, B_ALL)
        put(ws, r, 5, "", None, None, B_ALL)
        put(ws, r, 6, "監督員確認", F(size=8), AL_C, B_ALL)
        put(ws, r + 1, 4, core.agent(fy), F(size=9), AL_C, B_ALL)
        ws.merge_cells(start_row=r + 1, start_column=4, end_row=r + 1, end_column=5)
        put(ws, r + 1, 6, core.city(fy), F(size=9), AL_C, B_ALL)
        put(ws, r + 3, 1, "【架空】汐見重工環境株式会社　みなと臨海工場 定期整備作業所", F(size=8))
        page(ws, landscape=False, paper=9, fit_w=1, fit_h=1, area=f"A1:F{r + 3}", margins=(0.6, 0.5, 0.6, 0.6))
    props(wb, f"【架空】{cn} 作業日報（{lab}）", "m2", "作業日報", dt.datetime(2014, 4, 7, 8, 0), modified_by=core.agent(fy))
    wb.save(os.path.join(OUT, fn))
    return fn


# ---------------------------------------------------------------- D：R3 縦型拡張（R3〜）
RANK = {"遅れ": 3, "完了(遅れ)": 2, "予定どおり": 1, "完了": 1, "前倒し": 0}


def judge(d, t, prog):
    pe, ps = t["plan_end"], t["plan_start"]
    hi = "遅れ" if d > pe else ("前倒し" if d < ps else "予定どおり")
    if prog == 100:
        hi = "完了(遅れ)" if d > pe else "完了"
    return hi


def build_d(win):
    by = lines_by_day(win)
    if not by:
        return None
    fy = win["fy"]
    fn = fname(win, form_of(fy))
    who = "m4" if win["contract"] == "KIKAN" else core.agent(fy)
    ans.start(fn, False, "作業日報", fy, "R3縦型拡張（m3）", who, win["contract"])
    wb = Workbook()
    wb.remove(wb.active)
    F = lambda **k: font(FONT_P, **k)
    cn, lab = cname_of(win)
    done = defaultdict(int)
    cum_co = defaultdict(int)
    for k, d in enumerate(sorted(by), start=1):
        ws = wb.create_sheet(f"{d.month}.{d.day}")
        for c, w in zip(range(1, 10), [11, 36, 6, 6, 3, 10, 8, 9, 9]):
            ws.column_dimensions[CL(c)].width = w
        merge(ws, 1, 1, 1, 9, "作　業　日　報", F(size=18, bold=True), AL_C, None)
        put(ws, 2, 9, f"No.{k}", F(size=9), AL_R)
        put(ws, 3, 1, "件　名", F(size=9, bold=True), AL_C, B_ALL, fill("EEEEEE"))
        merge(ws, 3, 2, 3, 9, f"{cn}（{lab}）", F(size=10), AL_L, B_ALL)
        put(ws, 4, 1, "日　付", F(size=9, bold=True), AL_C, B_ALL, fill("EEEEEE"))
        put(ws, 4, 2, f"{d_wa_long(d)}（{WD[d.weekday()]}）", F(size=10), AL_L, B_ALL)
        put(ws, 4, 3, "天候", F(size=9, bold=True), AL_C, B_ALL, fill("EEEEEE"))
        merge(ws, 4, 4, 4, 5, weather(d), F(size=10), AL_C, B_ALL)
        put(ws, 4, 6, "時間", F(size=9, bold=True), AL_C, B_ALL, fill("EEEEEE"))
        merge(ws, 4, 7, 4, 9, "8:00〜17:00（残業は各行）", F(size=8), AL_C, B_ALL)
        n_all = sum(x["persons"] for x in by[d]) + 1
        merge(ws, 5, 1, 5, 9, f"検温・体調確認　全員異常なし（{n_all}名　朝礼時）", F(size=8, color="0000CC"), AL_L, B_ALL)
        for c, h in enumerate(["班　名", "作　業　内　容", "人員", "時間", "", "工程No.", "工数\n(人・h)", "予定比", "備　考"], start=1):
            put(ws, 7, c, h, F(size=9, bold=True), AL_C, B_ALL, fill("EEEEEE") if c <= 5 or c == 9 else fill("DDEBF7"))
        ws.merge_cells(start_row=7, start_column=4, end_row=7, end_column=5)
        ws.row_dimensions[7].height = 26
        r = 8
        teams = defaultdict(list)
        for x in by[d]:
            teams[x["task"]["team"]].append(x)
        prog_rows = []
        co_today = defaultdict(lambda: [0, 0.0])
        for tk in sorted(teams, key=lambda t: min(x["task"]["no"] for x in teams[t])):
            xs = sorted(teams[tk], key=lambda x: x["task"]["no"])
            tx = [f"{x['task']['name'].split('（')[0]}：{texts.line_text(x, 'm3')}" for x in xs]
            pers = sum(x["persons"] for x in xs)
            hrs = max(x["hours"] for x in xs)
            kos = sum(x["persons"] * x["hours"] for x in xs)
            hj = []
            for x in xs:
                t = x["task"]
                done[t["no"]] += 1
                prog = round(100 * done[t["no"]] / len(t["days"]))
                hj.append(judge(d, t, prog))
                prog_rows.append((t["no"], t["name"], prog, hj[-1]))
            worst = max(hj, key=lambda h: RANK[h])
            co = xs[0]["team"]["co"].replace("【架空】", "")
            co_today[co][0] += pers
            co_today[co][1] += kos
            nos = "・".join(x["task"]["no"] for x in xs)
            vals = [xs[0]["team"]["name"], "\n".join(tx), f"{pers}名", f"{hrs:.1f}", "h", nos, round(kos, 1), worst, ""]
            for c, v in enumerate(vals, start=1):
                put(ws, r, c, v, F(size=8), AL_LW if c == 2 else AL_C, B_ALL)
            if "遅れ" in worst:
                ws.cell(row=r, column=8).font = F(size=8, color="C00000")
            ws.row_dimensions[r].height = max(24, 14 * len(xs) + 8 * sum(len(t) > 34 for t in tx))
            ans.rec(fn, ws.title, r, "item", dict(date=d.isoformat(), team=vals[0], text=vals[1], persons=pers, hours=hrs, no=nos,
                                                   manDays=round(kos / 8, 2), manDaysText=f"{round(kos, 1)}", 予定比=worst),
                    works=[w["id"] for x in xs for w in x["works"]] or None, check=[vals[0], nos])
            r += 1
        while r < 16:
            for c in range(1, 10):
                put(ws, r, c, None, None, None, B_ALL)
            r += 1
        tot = sum(x["persons"] for x in by[d])
        totk = round(sum(x["persons"] * x["hours"] for x in by[d]), 1)
        put(ws, r, 1, "計", F(size=9, bold=True), AL_C, B_ALL)
        put(ws, r, 2, "", None, None, B_ALL)
        put(ws, r, 3, f"{tot}名", F(size=9, bold=True), AL_C, B_ALL)
        merge(ws, r, 4, r, 6, "（代理人 1 名 含まず）", F(size=7), AL_L, B_ALL)
        put(ws, r, 7, totk, F(size=9, bold=True), AL_C, B_ALL)
        merge(ws, r, 8, r, 9, "", None, None, B_ALL)
        ans.rec(fn, ws.title, r, "total", dict(date=d.isoformat(), persons=tot, manDays=round(totk / 8, 2), manDaysText=str(totk)), check=[f"{tot}名"])
        r += 2
        merge(ws, r, 1, r, 4, "協力会社別 出面", F(size=9, bold=True), AL_L, B_ALL, fill("DDEBF7"))
        merge(ws, r, 5, r, 6, "本日 人", F(size=8, bold=True), AL_C, B_ALL, fill("DDEBF7"))
        put(ws, r, 7, "本日 人・h", F(size=8, bold=True), AL_C, B_ALL, fill("DDEBF7"))
        merge(ws, r, 8, r, 9, "累計 延べ人", F(size=8, bold=True), AL_C, B_ALL, fill("DDEBF7"))
        r += 1
        for co in sorted(co_today, key=lambda c: (c != "汐見重工環境", c)):
            cum_co[co] += co_today[co][0]
            merge(ws, r, 1, r, 4, ("（元請）" if co == "汐見重工環境" else "") + co, F(size=8), AL_L, B_ALL)
            merge(ws, r, 5, r, 6, co_today[co][0], F(size=8), AL_C, B_ALL)
            put(ws, r, 7, round(co_today[co][1], 1), F(size=8), AL_C, B_ALL)
            merge(ws, r, 8, r, 9, cum_co[co], F(size=8), AL_C, B_ALL)
            ans.rec(fn, ws.title, r, "subtotal", dict(date=d.isoformat(), 会社=co, persons=co_today[co][0], manDaysText=str(round(co_today[co][1], 1)),
                                                       manDays=round(co_today[co][1] / 8, 2), 累計=cum_co[co]), check=[co])
            r += 1
        r += 1
        merge(ws, r, 1, r, 7, "工程進捗（工程表 No.）", F(size=9, bold=True), AL_L, B_ALL, fill("DDEBF7"))
        put(ws, r, 8, "進捗 %", F(size=8, bold=True), AL_C, B_ALL, fill("DDEBF7"))
        put(ws, r, 9, "予定比", F(size=8, bold=True), AL_C, B_ALL, fill("DDEBF7"))
        r += 1
        for no, nm, pg, hj in sorted(prog_rows):
            put(ws, r, 1, no, F(size=8), AL_C, B_ALL)
            merge(ws, r, 2, r, 7, nm, F(size=8), AL_L, B_ALL)
            put(ws, r, 8, pg, F(size=8), AL_C, B_ALL)
            put(ws, r, 9, hj, F(size=8, color="C00000" if "遅れ" in hj else "000000"), AL_C, B_ALL)
            ans.rec(fn, ws.title, r, "item", dict(date=d.isoformat(), no=no, name=nm, 進捗=pg, 予定比=hj), check=[no])
            r += 1
        r += 1
        put(ws, r, 1, "危険予知", F(size=9, bold=True), AL_C, B_ALL, fill("EEEEEE"))
        merge(ws, r, 2, r, 9, KY[(d.toordinal() + 3) % len(KY)], F(size=9), AL_L, B_ALL)
        r += 1
        merge(ws, r, 1, r + 2, 1, "特記事項", F(size=9, bold=True), AL_C, B_ALL, fill("EEEEEE"))
        eps = [ln["episode"][4] for ln in by[d] if ln.get("episode")]
        merge(ws, r, 2, r + 2, 9, "　".join(eps), F(size=9), AL_LT, B_ALL)
        r += 4
        put(ws, r, 6, "代理人" if who != "m4" else "監理技術者", F(size=8), AL_C, B_ALL)
        put(ws, r, 7, "", None, None, B_ALL)
        merge(ws, r, 8, r, 9, "監督員確認", F(size=8), AL_C, B_ALL)
        merge(ws, r + 1, 6, r + 1, 7, who, F(size=9), AL_C, B_ALL)
        merge(ws, r + 1, 8, r + 1, 9, core.city(fy), F(size=9), AL_C, B_ALL)
        ws.row_dimensions[r + 1].height = 28
        put(ws, r + 3, 1, "【架空】汐見重工環境株式会社　みなと臨海工場作業所", F(size=8))
        page(ws, landscape=False, paper=9, fit_w=1, fit_h=1, area=f"A1:I{r + 3}", margins=(0.5, 0.5, 0.5, 0.5))
    props(wb, f"【架空】{cn} 作業日報（{lab}）", "m2", "作業日報", dt.datetime(2014, 4, 7, 8, 0), modified_by=who)
    wb.save(os.path.join(OUT, fn))
    return fn


# ---------------------------------------------------------------- E：東雲 作業週報（タービン）
def build_weekly(win):
    by = lines_by_day(win)
    if not by:
        return None
    fy = win["fy"]
    fn = f"【架空】みなと臨海工場_作業日報_{core.wa_fy(fy)}_蒸気タービン定期事業者検査_協力会社週報様式.xlsx"
    ans.start(fn, False, "作業日報", fy, "協力会社 作業週報（東雲）", "t1", win["contract"])
    wb = Workbook()
    wb.remove(wb.active)
    F = lambda **k: font(FONT_G, **k)
    cn = core.get().contracts[win["contract"]]["name"]
    days = sorted(by)
    mon = days[0] - dt.timedelta(days=days[0].weekday())
    wk = 0
    insp = {t["type"]: t["days"] for t in win["tasks"]}
    while mon <= days[-1]:
        wk += 1
        sun = mon + dt.timedelta(days=6)
        ws = wb.create_sheet(f"第{wk}週")
        for c, w in zip(range(1, 10), [6, 4, 44, 5, 5, 5, 8, 14, 18]):
            ws.column_dimensions[CL(c)].width = w
        merge(ws, 1, 1, 1, 9, "作 業 週 報", F(size=16, bold=True), AL_C, None)
        merge(ws, 2, 1, 2, 9, "【架空】東雲タービンサービス株式会社　みなと臨海工場 蒸気タービン定期事業者検査", F(size=9), AL_C, None)
        put(ws, 3, 1, "件名", F(size=8, bold=True), AL_C, B_ALL, fill("E2EFDA"))
        merge(ws, 3, 2, 3, 9, cn + f"（元請 汐見重工環境 現場代理人 {core.agent(fy)}）", F(size=9), AL_L, B_ALL)
        put(ws, 4, 1, "期間", F(size=8, bold=True), AL_C, B_ALL, fill("E2EFDA"))
        merge(ws, 4, 2, 4, 6, f"{mon.year}/{mon.month}/{mon.day}（月）〜 {sun.month}/{sun.day}（日）　第 {wk} 週", F(size=9), AL_L, B_ALL)
        put(ws, 4, 7, "作成", F(size=8, bold=True), AL_C, B_ALL, fill("E2EFDA"))
        merge(ws, 4, 8, 4, 9, "主任技術者 t1", F(size=9), AL_C, B_ALL)
        for c, h in enumerate(["日付", "曜", "作業内容（工程No.）", "東雲", "元請", "保温", "時間", "立会・検査", "記事"], start=1):
            put(ws, 6, c, h, F(size=8, bold=True), AL_C, B_ALL, fill("E2EFDA"))
        r = 7
        wsum = defaultdict(int)
        for i in range(7):
            d = mon + dt.timedelta(days=i)
            xs = sorted(by.get(d, []), key=lambda x: x["task"]["no"])
            cnt = defaultdict(int)
            for x in xs:
                key = {"東雲": "東雲", "保温": "保温"}.get(x["task"]["team"], "元請")
                cnt[key] += x["persons"]
                wsum[key] += x["persons"]
            txt = "\n".join(f"{texts.line_text(x, 'm3')}（{x['task']['no']}）" for x in xs) or ("休工" if d.weekday() == 6 else "—")
            hrs = max((x["hours"] for x in xs), default=None)
            tachi = ""
            if d in insp.get("t_insp", []):
                tachi = f"開放検査（市 {core.city(fy)}）"
            if d in insp.get("t_final", []):
                tachi = f"総合負荷検査（市 {core.city(fy)}）"
            vals = [f"{d.month}/{d.day}", WD[d.weekday()], txt, cnt.get("東雲") or "", cnt.get("元請") or "", cnt.get("保温") or "",
                    f"{hrs:.1f}h" if hrs else "", tachi, ""]
            for c, v in enumerate(vals, start=1):
                put(ws, r, c, v, F(size=8, color="FF0000" if (c <= 2 and d.weekday() == 6) else "000000"), AL_LW if c in (3, 9) else AL_C, B_ALL)
            ws.row_dimensions[r].height = max(30, 13 * max(1, len(xs)) + 10)
            if xs:
                p = sum(x["persons"] for x in xs)
                ans.rec(fn, ws.title, r, "item", dict(date=d.isoformat(), text=txt, 東雲=cnt.get("東雲", 0), 元請=cnt.get("元請", 0),
                                                       保温=cnt.get("保温", 0), time=vals[6], 立会=tachi,
                                                       manDays=round(sum(x["persons"] * x["hours"] for x in xs) / 8, 2), manDaysText=None),
                        works=[w["id"] for x in xs for w in x["works"]] or None, check=[vals[0]])
            r += 1
        put(ws, r, 1, "週計", F(size=8, bold=True), AL_C, B_ALL, fill("F2F2F2"))
        merge(ws, r, 2, r, 3, "延べ人数", F(size=8), AL_R, B_ALL, fill("F2F2F2"))
        for c, kk in ((4, "東雲"), (5, "元請"), (6, "保温")):
            put(ws, r, c, wsum.get(kk) or "", F(size=8, bold=True), AL_C, B_ALL, fill("F2F2F2"))
        merge(ws, r, 7, r, 9, "", None, None, B_ALL, fill("F2F2F2"))
        ans.rec(fn, ws.title, r, "total", dict(week=wk, 東雲=wsum.get("東雲", 0), 元請=wsum.get("元請", 0), 保温=wsum.get("保温", 0)), check=["週計"])
        r += 2
        nxt = [t for t in win["tasks"] if sun < t["plan_start"] <= sun + dt.timedelta(days=7)]
        merge(ws, r, 1, r + 2, 2, "来週\n予定", F(size=8, bold=True), AL_C, B_ALL, fill("E2EFDA"))
        merge(ws, r, 3, r + 2, 9, "\n".join(f"{t['no']} {t['name']}（{t['plan_start'].month}/{t['plan_start'].day}〜）" for t in nxt) or "（継続作業）",
              F(size=8), AL_LT, B_ALL)
        r += 4
        put(ws, r, 7, "元請確認", F(size=8), AL_C, B_ALL)
        merge(ws, r, 8, r, 9, core.agent(fy), F(size=9), AL_C, B_ALL)
        page(ws, landscape=False, paper=9, fit_w=1, fit_h=1, area=f"A1:I{r}", margins=(0.5, 0.5, 0.5, 0.5))
        mon += dt.timedelta(days=7)
    props(wb, f"【架空】{cn} 作業週報", "t1", "作業日報（週報）", dt.datetime(2006, 10, 2, 8, 0), modified_by="t1")
    wb.save(os.path.join(OUT, fn))
    return fn


def main():
    W = core.get()
    out = []
    for win in sorted(W.windows.values(), key=lambda w: (w["fy"], w["id"])):
        c = W.contracts[win["contract"]]
        if c["kind"] == "TURB":
            f = build_weekly(win)
        else:
            fy = win["fy"]
            f = (build_a if fy <= 2005 else build_b if fy <= 2013 else build_c if fy <= 2020 else build_d)(win)
        if f:
            out.append(f)
    return out


if __name__ == "__main__":
    import time
    t0 = time.time()
    print(len(main()), time.time() - t0)
