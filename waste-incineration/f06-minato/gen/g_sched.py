# -*- coding: utf-8 -*-
"""工程表（全年度）
m1 1炉1枚（H16〜H21）／ 3炉1枚（H22〜R3、H22 に m1 が作成、m2・m3 が継続）／ 年間＋詳細（R4〜、m3）
"""
import datetime as dt
import os

from openpyxl import Workbook
from openpyxl.styles import Border
from openpyxl.utils import get_column_letter as CL

import ans
import core
from xlhelp import *

D = dt.date.fromisoformat
FAC = "【架空】みなと臨海工場"


def drange(a, b):
    d = a
    while d <= b:
        yield d
        d += dt.timedelta(days=1)


def short(c):
    return {"TEISEI": "定期整備業務委託", "ZENRO": "全炉停止時共通設備整備業務委託", "TURB": "蒸気タービン定期事業者検査業務委託",
            "KOSHIN": "機器更新工事", "HOSHU": "補修工事", "KIKAN": "基幹的設備改良工事"}[c["kind"]]


def fname(fy, label, form):
    return f"【架空】みなと臨海工場_工程表_{core.wa_fy(fy)}_{label}_{form}.xlsx"


def persons(t):
    lns = [ln for ln in core.get().lines if ln["task"] is t]
    return max([ln["persons"] for ln in lns], default=core.TEAM_BASE[t["team"]][2])


# ---------------------------------------------------------------- m1 1炉1枚
def m1_sheet(wb, win, fn, title):
    W = core.get()
    ws = wb.create_sheet(win["label"].replace(" ", "")[:28])
    F = lambda **k: font(FONT_G, **k)
    a, b, pa, pb = W.window_dates(win)
    days = list(drange(pa, pb + dt.timedelta(days=1)))
    DC = 5
    for c, w in zip(range(1, DC), [4, 28, 11, 5]):
        ws.column_dimensions[CL(c)].width = w
    for i in range(len(days)):
        ws.column_dimensions[CL(DC + i)].width = 3.0
    put(ws, 1, 1, title, F(size=13, bold=True))
    put(ws, 2, 1, f"{win['label']}　予定工程表　　【架空】汐見重工環境株式会社　現場代理人 {core.agent(win['fy'])}", F(size=9))
    put(ws, 4, 1, "No.", F(size=8, bold=True), AL_C, B_ALL)
    put(ws, 4, 2, "作業項目", F(size=8, bold=True), AL_C, B_ALL)
    put(ws, 4, 3, "担当", F(size=8, bold=True), AL_C, B_ALL)
    put(ws, 4, 4, "日数", F(size=8, bold=True), AL_C, B_ALL)
    for i, d in enumerate(days):
        put(ws, 3, DC + i, f"{d.month}/" if (i == 0 or d.day == 1) else None, F(size=7), AL_C)
        put(ws, 4, DC + i, d.day, F(size=7, color="FF0000" if d.weekday() == 6 else "000000"), AL_C, B_ALL)
    r = 5
    for k, t in enumerate(win["tasks"], start=1):
        tm = core.TEAM_BASE[t["team"]][0]
        n = len(t["pdays"])
        vals = [k, t["name"], tm, n]
        for c, v in enumerate(vals, start=1):
            put(ws, r, c, v, F(size=8), AL_L if c == 2 else AL_C, B_ALL)
        ps = set(t["pdays"])
        first, last = t["pdays"][0], t["pdays"][-1]
        for i, d in enumerate(days):
            v = None
            if first <= d <= last:
                v = "◎" if first == last else ("○" if d in (first, last) else "─")
            put(ws, r, DC + i, v, F(size=8), AL_C, Border(left=dot, right=dot, top=thin, bottom=thin))
        ans.rec(fn, ws.title, r, "item", dict(no=str(k), name=t["name"], team=tm, days=n, planStart=first.isoformat(),
                                               planEnd=last.isoformat()), works=t["work_ids"], check=[t["name"]])
        r += 1
    put(ws, r + 1, 2, "記号　○：開始・終了　─：作業期間　◎：1 日", F(size=8))
    page(ws, landscape=True, paper=9, fit_w=1, fit_h=1, area=f"A1:{CL(DC + len(days) - 1)}{r + 1}")


def build_m1(fy, contract_key, wins, label):
    W = core.get()
    c = W.contracts[contract_key]
    fn = fname(fy, label, "1炉1枚様式")
    ans.start(fn, False, "工程表", fy, "1炉1枚（m1 H16）", core.agent(fy), contract_key)
    wb = Workbook()
    wb.remove(wb.active)
    for win in wins:
        m1_sheet(wb, win, fn, f"{core.wa_fy(fy)}　{FAC}　{short(c)}　工程表")
    props(wb, f"【架空】{c['name']} 工程表", "m1", "工程表", dt.datetime(fy, 4, 10, 9, 0), modified_by="m1")
    wb.save(os.path.join(OUT, fn))
    return fn


# ---------------------------------------------------------------- 3炉1枚（ブロックを縦に並べる）
def block(ws, r, win, fn, NDAY, ditto=False, extra_rows=0):
    W = core.get()
    F = lambda **k: font(FONT_G, **k)
    DC = 5
    a, b, pa, pb = W.window_dates(win)
    st = pa
    merge(ws, r, 1, r, DC + NDAY - 1, f"{win['label']}　　{d_wa_long(pa)} 〜 {d_wa_long(pb)}", F(size=10, bold=True), AL_L, B_ALL, fill("D9D9D9"))
    r += 1
    merge(ws, r, 1, r + 2, 1, "No.", F(size=8, bold=True), AL_C)
    merge(ws, r, 2, r + 2, 2, "作　業　項　目", F(size=8, bold=True), AL_C)
    merge(ws, r, 3, r + 2, 3, "担　当", F(size=8, bold=True), AL_C)
    merge(ws, r, 4, r + 2, 4, "日数", F(size=8, bold=True), AL_C)
    days = [st + dt.timedelta(days=i) for i in range(NDAY)]
    m0 = None
    for i, d in enumerate(days):
        c = DC + i
        if d.month != m0:
            m0 = d.month
            j = i
            while j < NDAY and days[j].month == d.month:
                j += 1
            merge(ws, r, c, r, DC + j - 1, f"{d.month}月", F(size=8, bold=True), AL_C)
        put(ws, r + 1, c, d.day, F(size=7), AL_C, B_ALL)
        put(ws, r + 2, c, WD[d.weekday()], F(size=7, color="FF0000" if d.weekday() == 6 else ("0000FF" if d.weekday() == 5 else "000000")), AL_C, B_ALL)
        if d.weekday() == 6:
            ws.cell(row=r + 1, column=c).fill = fill("FDE9D9")
            ws.cell(row=r + 2, column=c).fill = fill("FDE9D9")
    r += 3
    prev = None
    bar = fill("595959")
    for k, t in enumerate(win["tasks"], start=1):
        tm = core.TEAM_BASE[t["team"]][0]
        a1, b1 = t["pdays"][0], t["pdays"][-1]
        put(ws, r, 1, k, F(size=8), AL_C, B_ALL)
        put(ws, r, 2, t["name"], F(size=8), AL_L, B_ALL)
        put(ws, r, 3, "〃" if (ditto and tm == prev) else tm, F(size=7), AL_C, B_ALL)
        prev = tm
        put(ws, r, 4, (b1 - a1).days + 1, F(size=8), AL_C, B_ALL)
        for i, d in enumerate(days):
            cell = put(ws, r, DC + i, None, None, None, Border(left=dot, right=dot, top=thin, bottom=thin))
            if a1 <= d <= b1:
                cell.fill = bar
        ans.rec(fn, ws.title, r, "item", dict(block=win["label"], no=str(k), name=t["name"], team=tm, days=(b1 - a1).days + 1,
                                               planStart=a1.isoformat(), planEnd=b1.isoformat()), works=t["work_ids"], check=[t["name"]])
        r += 1
    for _ in range(extra_rows):
        for c in range(1, DC + NDAY):
            put(ws, r, c, None, None, None, B_ALL)
        r += 1
    return r


def build_3ro(fy, contract_key, wins, label, note_line=None):
    W = core.get()
    c = W.contracts[contract_key]
    fn = fname(fy, label, "3炉1枚様式")
    ans.start(fn, False, "工程表", fy, "3炉1枚（H22 m1）", core.agent(fy), contract_key)
    wb = Workbook()
    ws = wb.active
    ws.title = f"{core.wa_fy(fy)[:-1].replace('平成', 'H').replace('令和', 'R')}工程表"
    F = lambda **k: font(FONT_G, **k)
    NDAY = max((W.window_dates(w)[3] - W.window_dates(w)[2]).days + 3 for w in wins)
    NDAY = max(NDAY, 20)
    DC = 5
    for cc, wdt in zip(range(1, DC), [4, 30, 10, 5]):
        ws.column_dimensions[CL(cc)].width = wdt
    for i in range(NDAY):
        ws.column_dimensions[CL(DC + i)].width = 3.1
    put(ws, 1, 1, f"{core.wa_fy(fy)}　{FAC}　{short(c)}　工程表", F(size=14, bold=True))
    put(ws, 2, 1, f"【架空】汐見重工環境株式会社　現場代理人 {core.agent(fy)}", F(size=9))
    put(ws, 2, 26, "様式 H22（m1）", F(size=8, color="808080"))
    r = 4
    for i, win in enumerate(wins):
        r = block(ws, r, win, fn, NDAY, ditto=i > 0, extra_rows=2 if (i == 1 and len(wins) >= 3) else 0)
        if i == 1 and note_line:
            put(ws, r, 2, note_line, F(size=8, color="0000FF"))
            r += 1
        r += 1
    put(ws, r, 2, "※ 立下げ・立上げは運転管理業務委託（運転会社）による。工程は炉の運転状況により変更することがある。", F(size=8))
    page(ws, landscape=len(wins) == 1, paper=8 if len(wins) > 1 else 9, fit_w=1, fit_h=1, area=f"A1:{CL(DC + NDAY - 1)}{r}")
    props(wb, f"【架空】{c['name']} 工程表", "m1", "工程表", dt.datetime(2010, 4, 12, 8, 30), modified_by=core.agent(fy))
    wb.save(os.path.join(OUT, fn))
    return fn


# ---------------------------------------------------------------- R4〜 年間＋詳細
def week_cols(fy):
    st = D(f"{fy}-04-01")
    st = st - dt.timedelta(days=st.weekday())
    out = []
    d = st
    while d <= D(f"{fy + 1}-03-31"):
        out.append(d)
        d += dt.timedelta(days=7)
    return out


def detail_sheet(wb, win, fn, title, note_rows, asof):
    W = core.get()
    F = lambda **k: font(FONT_P, **k)
    w2 = wb.create_sheet(win["label"].replace(" ", "").replace("蒸気タービン定期事業者検査", "タービン検査")[:28])
    DC = 6
    a, b, pa, pb = W.window_dates(win)
    st, en = min(a, pa), max(b, pb)
    days = list(drange(st, en))
    for c, wdt in zip(range(1, DC), [6, 32, 13, 5, 4.5]):
        w2.column_dimensions[CL(c)].width = wdt
    for i in range(len(days)):
        w2.column_dimensions[CL(DC + i)].width = 3.0 if len(days) <= 60 else 2.2
    put(w2, 1, 1, title, F(size=13, bold=True))
    put(w2, 2, 1, f"【架空】汐見重工環境株式会社　{'監理技術者 m4' if win['contract'] == 'KIKAN' else '現場代理人 ' + core.agent(win['fy'])}　　上段：予定　下段：実績"
                  + (f"（{asof} 現在）" if asof else ""), F(size=9))
    r = 4
    for c, h in ((1, "工程\nNo."), (2, "作　業　項　目"), (3, "担当班"), (4, "人員"), (5, "")):
        merge(w2, r, c, r + 2, c, h, F(size=8, bold=True), AL_C, B_ALL, fill("DDEBF7"))
    m0 = None
    for i, d in enumerate(days):
        c = DC + i
        if d.month != m0:
            m0 = d.month
            j = i
            while j < len(days) and days[j].month == d.month:
                j += 1
            merge(w2, r, c, r, DC + j - 1, f"{d.month}月", F(size=8, bold=True), AL_C, B_ALL, fill("DDEBF7"))
        put(w2, r + 1, c, d.day, F(size=6), AL_C, B_ALL, fill("DDEBF7"))
        put(w2, r + 2, c, WD[d.weekday()], F(size=6, color="FF0000" if d.weekday() == 6 else "000000"), AL_C, B_ALL,
            fill("FDE9D9") if d.weekday() == 6 else fill("DDEBF7"))
    r += 3
    lim = D("2026-11-13")
    for t in win["tasks"]:
        tm = core.TEAM_BASE[t["team"]][0]
        merge(w2, r, 1, r + 1, 1, t["no"], F(size=8), AL_C)
        merge(w2, r, 2, r + 1, 2, t["name"], F(size=8), AL_LW)
        merge(w2, r, 3, r + 1, 3, tm, F(size=7), AL_C)
        merge(w2, r, 4, r + 1, 4, persons(t), F(size=8), AL_C)
        put(w2, r, 5, "予", F(size=6), AL_C, B_ALL)
        put(w2, r + 1, 5, "実", F(size=6), AL_C, B_ALL)
        pdays, adays = set(t["pdays"]), set(d for d in t["days"] if d <= lim)
        for i, d in enumerate(days):
            c1 = put(w2, r, DC + i, None, None, None, Border(left=dot, right=dot, top=thin, bottom=hair))
            c2 = put(w2, r + 1, DC + i, None, None, None, Border(left=dot, right=dot, top=hair, bottom=thin))
            if d in pdays:
                c1.fill = fill("BDD7EE")
            if d in adays:
                c2.fill = fill("1F4E78")
        ad = sorted(adays)
        ans.rec(fn, w2.title, r, "item", dict(no=t["no"], name=t["name"], team=tm, persons=persons(t), planStart=t["pdays"][0].isoformat(),
                                               planEnd=t["pdays"][-1].isoformat(), actStart=ad[0].isoformat() if ad else None,
                                               actEnd=ad[-1].isoformat() if ad else None), works=t["work_ids"], check=[t["no"], t["name"]])
        r += 2
    for (label, a1, b1, color) in note_rows:
        merge(w2, r, 1, r, 5, label, F(size=7, color="C00000"), AL_LW)
        for i, d in enumerate(days):
            cell = put(w2, r, DC + i, None, None, None, Border(left=dot, right=dot, top=thin, bottom=thin))
            if a1 <= d <= b1:
                cell.fill = fill(color)
        r += 1
    page(w2, landscape=True, paper=8, fit_w=1, fit_h=1, area=f"A1:{CL(DC + len(days) - 1)}{r}")
    return w2


def build_r4(fy):
    """年間工程（全契約）＋各停止期間の詳細。基幹改良は別ブック"""
    W = core.get()
    fn = fname(fy, "年間工程", "基幹改良併記様式")
    ans.start(fn, False, "工程表", fy, "年間＋詳細（R4 m3）", core.agent(fy), None)
    wb = Workbook()
    ws = wb.active
    ws.title = "年間工程"
    F = lambda **k: font(FONT_P, **k)
    weeks = week_cols(fy)
    WC = 6
    for c, w in zip(range(1, WC), [9, 34, 4.5, 30, 16]):
        ws.column_dimensions[CL(c)].width = w
    for i in range(len(weeks)):
        ws.column_dimensions[CL(WC + i)].width = 2.6
    asof = "R8.11.13" if fy == 2026 else None
    put(ws, 1, 1, f"{FAC}　{core.wa_fy(fy)}　年間工程表（定期整備・全炉停止・タービン検査・基幹改良）", F(size=14, bold=True))
    put(ws, 2, 1, f"【架空】汐見重工環境株式会社　現場代理人 {core.agent(fy)}　　" + (f"{asof} 現在 実績記入" if asof else f"{d_wa(f'{fy + 1}-03-31')} 実績記入"), F(size=9))
    put(ws, 3, 2, "上段 予定　／　下段 実績", F(size=8))
    put(ws, 3, 3, None, None, None, B_ALL, fill("BDD7EE"))
    put(ws, 3, 4, "予定　　　　　　　　実績", F(size=8))
    put(ws, 3, 5, None, None, None, B_ALL, fill("1F4E78"))
    for c, h in ((1, "区分"), (2, "工事・作業"), (3, ""), (4, "契約件名"), (5, "備考")):
        merge(ws, 5, c, 6, c, h, F(size=8, bold=True), AL_C, B_ALL, fill("DDEBF7"))
    m0, j0 = None, 0
    for i, w in enumerate(weeks + [None]):
        mm = (w + dt.timedelta(days=3)).month if w else None
        if mm != m0:
            if m0 is not None:
                merge(ws, 5, WC + j0, 5, WC + i - 1, f"{m0}月", F(size=8, bold=True), AL_C, B_ALL, fill("DDEBF7"))
            m0, j0 = mm, i
        if w:
            put(ws, 6, WC + i, w.day, F(size=6), AL_C, B_ALL, fill("DDEBF7"))
    lim = D("2026-11-13")
    rows = []
    wins = sorted(W.windows_of(fy), key=lambda w: W.window_dates(w)[2])
    for win in wins:
        a, b, pa, pb = W.window_dates(win)
        c = W.contracts[win["contract"]]
        kub = {"TEISEI": "定期整備", "ZENRO": "全炉停止", "TURB": "タービン", "KIKAN": "基幹改良", "KOSHIN": "更新工事", "HOSHU": "補修"}[c["kind"]]
        if c["kind"] == "TEISEI" and win["unit"] == "Z":
            kub = "全炉停止"
        rows.append((kub, win["label"], (pa, pb), (a, min(b, lim)) if a <= lim else None, c["name"], "", win))
    r = 7
    def wk_hit(w, a, b):
        return a <= w + dt.timedelta(days=6) and w <= b
    grp = None
    for (kub, name, pl, ac, cn, bk, win) in sorted(rows, key=lambda x: (["定期整備", "全炉停止", "タービン", "基幹改良", "更新工事", "補修"].index(x[0]), x[2][0])):
        if kub != grp:
            grp = kub
            merge(ws, r, 1, r, WC + len(weeks) - 1, f"■ {kub}", F(size=8, bold=True), AL_L, B_ALL, fill("F2F2F2"))
            r += 1
        merge(ws, r, 1, r + 1, 1, kub, F(size=7), AL_C)
        merge(ws, r, 2, r + 1, 2, name, F(size=8), AL_LW)
        put(ws, r, 3, "予定", F(size=6), AL_C, B_ALL)
        put(ws, r + 1, 3, "実績", F(size=6), AL_C, B_ALL)
        delay = ""
        if ac and ac[1] > pl[1] and fy < 2026:
            delay = f"{(ac[1] - pl[1]).days} 日延長"
        merge(ws, r, 4, r + 1, 4, cn, F(size=7), AL_LW)
        merge(ws, r, 5, r + 1, 5, delay, F(size=7), AL_LW)
        for i, w in enumerate(weeks):
            c1 = put(ws, r, WC + i, None, None, None, Border(left=dot, right=dot, top=thin, bottom=hair))
            c2 = put(ws, r + 1, WC + i, None, None, None, Border(left=dot, right=dot, top=hair, bottom=thin))
            if wk_hit(w, pl[0], pl[1]):
                c1.fill = fill("BDD7EE")
            if ac and wk_hit(w, ac[0], ac[1]):
                c2.fill = fill("1F4E78")
        ans.rec(fn, "年間工程", r, "item", dict(区分=kub, name=name, contract=cn, planStart=pl[0].isoformat(), planEnd=pl[1].isoformat(),
                                               actStart=ac[0].isoformat() if ac else None, actEnd=ac[1].isoformat() if ac else None, 備考=delay),
                works=[wid for t in win["tasks"] for wid in t["work_ids"]], check=[name])
        r += 2
    put(ws, r + 1, 1, "※ 基幹改良は 1 年に 1 炉ずつ長期停止（R4 1号炉、R5 2号炉、R6 3号炉、R7 共通・タービン制御）。2 炉同時の停止は全炉停止時のみ。", F(size=8))
    put(ws, r + 2, 1, "※ タービン停止中はタービンバイパス（減圧減温）で各炉運転継続。", F(size=8))
    page(ws, landscape=True, paper=8, fit_w=1, fit_h=1, area=f"A1:{CL(WC + len(weeks) - 1)}{r + 2}")
    zw = [w for w in wins if w["unit"] == "Z"]
    for win in wins:
        if win["contract"] == "KIKAN":
            continue
        c = W.contracts[win["contract"]]
        notes = []
        if win["unit"] in "123" and zw:
            a, b, pa, pb = W.window_dates(zw[0])
            notes.append(("参考　全炉停止（別契約）", pa, pb, "F4B084"))
        detail_sheet(wb, win, fn, f"{core.wa_fy(fy)}　{short(c)}　{win['label']} 工程表", notes, asof)
    props(wb, f"【架空】{core.wa_fy(fy)} みなと臨海工場 年間工程表", "m3", "工程表", dt.datetime(2022, 3, 18, 9, 0), modified_by="m3")
    wb.save(os.path.join(OUT, fn))
    return fn


def build_kikan(fy):
    W = core.get()
    win = W.windows[f"K{fy}"]
    fn = fname(fy, "基幹的設備改良工事", "基幹改良併記様式")
    ans.start(fn, False, "工程表", fy, "年間＋詳細（R4 m3 様式を m4 が使用）", "m4", "KIKAN")
    wb = Workbook()
    wb.remove(wb.active)
    detail_sheet(wb, win, fn, f"{W.contracts['KIKAN']['name']}　{win['label']} 工程表", [], None)
    props(wb, f"【架空】みなと臨海工場 基幹的設備改良工事 工程表（{core.wa_fy(fy)}）", "m4", "工程表", dt.datetime(2022, 5, 9, 9, 0), modified_by="m4")
    wb.save(os.path.join(OUT, fn))
    return fn


def main():
    W = core.get()
    out = []
    for fy in core.FY_RANGE:
        if fy >= 2022:
            out.append(build_r4(fy))
            if fy <= 2025:
                out.append(build_kikan(fy))
            continue
        tw = [W.windows[f"{fy}-{u}"] for u in "123"]
        zw = W.windows[f"{fy}-Z"]
        if fy <= 2009:
            out.append(build_m1(fy, f"{fy}-TEISEI", tw + [zw], "定期整備業務委託"))
            if fy in core.TURB_YEARS:
                out.append(build_m1(fy, f"{fy}-TURB", [W.windows[f"{fy}-T"]], "タービン定期事業者検査"))
        else:
            if fy < core.ZENRO_SPLIT_FY:
                out.append(build_3ro(fy, f"{fy}-TEISEI", tw + [zw], "定期整備業務委託"))
            else:
                a, b, pa, pb = W.window_dates(zw)
                out.append(build_3ro(fy, f"{fy}-TEISEI", tw, "定期整備業務委託", note_line=f"参考：全炉停止 {d_wa(pa)}〜{d_wa(pb)}（別契約）"))
                out.append(build_3ro(fy, f"{fy}-ZENRO", [zw], "全炉停止時共通設備整備業務委託"))
            if fy in core.TURB_YEARS:
                out.append(build_3ro(fy, f"{fy}-TURB", [W.windows[f"{fy}-T"]], "タービン定期事業者検査"))
        for c in [c for c in W.contracts.values() if c["fy"] == fy and c["kind"] in ("KOSHIN", "HOSHU")]:
            wins = sorted([w for w in W.windows.values() if w["contract"] == c["id"]], key=lambda w: w["id"])
            lab = c["name"].replace(core.wa_fy(fy) + " みなと臨海工場 ", "")
            out.append((build_m1 if fy <= 2009 else build_3ro)(fy, c["id"], wins, lab))
    return out


if __name__ == "__main__":
    print(len(main()))
