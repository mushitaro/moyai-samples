# -*- coding: utf-8 -*-
"""中長期整備計画（各年度の版）
H20〜H24 版：h2 保全計画（初版 H20）10 か年を 5 年ずつ「前期・後期」シート。H24 版は h3 が H23 版を年だけ送った（追記漏れ）
H25〜H31 版：h3 施設保全計画（長寿命化総合計画）年度ごと 計画・実績・費用 の 3 段（H29 から h4、H29 の実績欄は空のまま）
R2〜R8 版 ：h4 延命化計画（1 シート、旧 3 段の列は非表示で残る。R4 から h5）
"""
import datetime as dt
import os

from openpyxl import Workbook
from openpyxl.styles import Border, Font
from openpyxl.utils import get_column_letter as CL

import ans
import core
import equipment as EQ
import world as W
from s_ledger import SETSUBI_ORDER, UNIT_ORDER, cycle_txt, hozen
from s_plan import SYM_NEW, SYM_OLD, cost_of, health, kinds, legal_rows, symstr
from xlhelp import *

SYM_H2 = {"I": "○", "O": "◇", "R": "●", "U": "●"}
UN = {"1": "1号炉", "2": "2号炉", "3": "3号炉", "C": "共通"}


def known_rows(v):
    """その版の時点で市の台帳に載っていた機器（改良の反映は台帳と同じく遅れる）"""
    known_from = {"1": 2025, "2": 2024, "3": 2025, "C": 2026}
    out = []
    for e in core.get().eq:
        if e["added"]:
            if v >= known_from[e["unit"]] and e["added"] < v:
                out.append(e)
            continue
        if e["removed"] and v >= known_from[e["unit"]] and e["removed"] < v:
            continue
        out.append(e)
    return out


def uv(v):
    return "none" if v <= 2015 else ("2019" if v <= 2019 else "2020")


def cyc(e):
    p = W._params(e)
    return min([x["iv"] for x in p], default=None)


def fn_of(v, form):
    lab = {"h2": "保全計画様式", "h3": "施設保全計画様式", "h4": "延命化計画様式"}[form]
    ver = "平成31年度版" if v == 2019 else core.wa_fy(v).replace("年度", "年度版")
    return f"【架空】みなと臨海工場_中長期整備計画_{ver}_{lab}.xlsx"


# ================================================================ h2（H20〜H24）
def build_h2(v):
    fn = fn_of(v, "h2")
    who = "h2" if v <= 2011 else "h3"
    ans.start(fn, True, "中長期整備計画", v, "保全計画（h2 H20 初版、前期・後期）", who)
    wb = Workbook()
    wb.remove(wb.active)
    F = lambda **k: font(FONT_G, **k)
    cut = v if v <= 2011 else 2011  # H24 版は H23 版の内容のまま年だけ送った
    years = list(range(v, v + 10))
    rows = sorted([e for e in core.get().eq if EQ.alive(e, v)], key=lambda e: (UNIT_ORDER[e["unit"]], e["sysno"], e["loop"], e["tag"]))
    plans = {e["tag"]: W.simulate(e, cutoff=cut, uversion="none", until=v + 10) for e in rows}
    for half, ys in (("前期", years[:5]), ("後期", years[5:])):
        ws = wb.create_sheet(f"{half}（{core.wa_fy(ys[0])[:-1].replace('平成', 'H')}〜{core.wa_fy(ys[-1])[:-1].replace('平成', 'H').replace('令和', 'R')}）")
        for c, w in zip(range(1, 8), [5, 7, 13, 28, 12, 22, 4]):
            ws.column_dimensions[CL(c)].width = w
        put(ws, 1, 1, f"【架空】みなと臨海工場　保全計画（{half}）", F(size=14, bold=True))
        put(ws, 2, 1, f"{core.wa_fy(v)}版　環境局施設課 {who}　（H20 初版 h2）　凡例 ○：点検　◇：整備　●：更新", F(size=9))
        hdr = ["No.", "号炉", "系統", "機器名称", "タグNo.", "整備内容（周期）", ""]
        for c, h in enumerate(hdr, start=1):
            put(ws, 4, c, h, F(size=8, bold=True), AL_C, B_ALL, fill("E7E6E6"))
        for i, y in enumerate(ys):
            put(ws, 4, 8 + i, core.wa_fy(y).replace("年度", "").replace("平成", "H").replace("令和", "R"), F(size=8, bold=True), AL_C, B_ALL, fill("E7E6E6"))
            ws.column_dimensions[CL(8 + i)].width = 5.5
        put(ws, 4, 13, "備考", F(size=8, bold=True), AL_C, B_ALL, fill("E7E6E6"))
        ws.column_dimensions["M"].width = 20
        ws.column_dimensions["G"].hidden = True
        r = 5
        prev_u = None
        for n, e in enumerate(rows, start=1):
            if e["unit"] != prev_u:
                merge(ws, r, 1, r, 13, f"◆ {UN[e['unit']]}", F(size=9, bold=True), AL_L, B_ALL, fill("F2F2F2"))
                ans.rec(fn, ws.title, r, "parent", dict(path=[UN[e["unit"]]]), check=[UN[e["unit"]]])
                r += 1
                prev_u = e["unit"]
            pl = plans[e["tag"]]
            cy = cycle_txt(e)
            vals = [n, e["unit"] if e["unit"] != "C" else "共", e["keito"], e["name"], e["tag"], cy or "（状態監視）"]
            for c, val in enumerate(vals, start=1):
                put(ws, r, c, val, F(size=8), AL_C if c in (1, 2) else AL_L, B_HAIR)
            marks = {}
            for i, y in enumerate(ys):
                ks = [k for k, _, _ in pl.get(y, [])]
                s = "".join(sorted({SYM_H2[k] for k in ks})) if ks else None
                put(ws, r, 8 + i, s, F(size=9), AL_C, B_HAIR)
                if s:
                    marks[str(y)] = s
            put(ws, r, 13, None, None, None, B_HAIR)
            ans.rec(fn, ws.title, r, "item", dict(path=[UN[e["unit"]], e["keito"], e["name"]], code=e["tag"], work=cy, cycle=cyc(e),
                                                   cost=None, unit=None, marks=marks, memo=None, comments={}), check=[e["tag"]])
            r += 1
        page(ws, landscape=False, paper=9, fit_w=1, fit_h=0, area=f"A1:M{r}", titles="4:4", footer="&P / &N")
    props(wb, f"【架空】みなと臨海工場 保全計画（{core.wa_fy(v)}版）", "h2", "中長期整備計画", dt.datetime(2008, 11, 4, 9, 0), modified_by=who)
    wb.save(os.path.join(OUT, fn))
    return fn


# ================================================================ h3（H25〜H31）
Y0, Y1 = 2013, 2027
OLD_YEARS = list(range(Y0, Y1 + 1))
YC = 9


def ycol(y, k=0):
    return YC + (y - Y0) * 3 + k


def build_h3(v):
    fn = fn_of(v, "h3")
    who = "h3" if v <= 2016 else "h4"
    ans.start(fn, True, "中長期整備計画", v, "施設保全計画（h3 H25、計画・実績・費用 3 段）", who)
    wb = Workbook()
    ws = wb.active
    ws.title = "整備計画"
    F = lambda **k: font(FONT_G, **k)
    cB = ycol(Y1, 2) + 1
    ver = "平成31年度版" if v == 2019 else core.wa_fy(v) + "版"
    put(ws, 1, 1, "【架空】みなと臨海工場　施設保全計画（長寿命化総合計画）　整備スケジュール表", F(size=14, bold=True))
    put(ws, 2, 1, f"{ver}　（H25 策定 施設課 h3" + (f" ／ {d_wa(dt.date(v, 4, 1))[:-3]} 更新 {who}）" if v > 2013 else "）") + "　各年度：計画・実績・費用（千円）", F(size=10))
    put(ws, 3, 1, "凡例　●：点検・検査　◎：整備（分解・補修）　★：更新　　実績欄は実施した作業を記入　費用は計画年度＝概算、過年度＝実績（千円）", F(size=9))
    if v >= 2016:
        put(ws, 3, 30, "▲：基幹的設備改良（予定・検討中）", F(size=9, color="FF0000"))
        note(ws.cell(row=3, column=30), "H28 追加。基幹改良の範囲・年度は延命化計画で確定させる", "h3")
    hdr = ["No.", "設備区分", "機器名称", "号炉", "タグNo.", "重要\n度", "健全\n度", "保全\n方式"]
    for i, (h, w) in enumerate(zip(hdr, [5, 12, 26, 4.5, 11, 4, 4, 4.5]), start=1):
        merge(ws, 5, i, 6, i, h, F(size=9, bold=True), AL_C, B_ALL, fill("D9D9D9"))
        ws.column_dimensions[CL(i)].width = w
    for y in OLD_YEARS:
        lab = "H31" if y == 2019 else fy_wa(y)
        merge(ws, 5, ycol(y), 5, ycol(y, 2), lab, F(size=9, bold=True), AL_C, B_ALL, fill("D9D9D9"))
        for k, t in enumerate(("計画", "実績", "費用")):
            put(ws, 6, ycol(y, k), t, F(size=7), AL_C, B_ALL, fill("D9D9D9" if k < 2 else "EDEDED"))
            ws.column_dimensions[CL(ycol(y, k))].width = 3.6 if k < 2 else 5.6
    merge(ws, 5, cB, 6, cB, "備　考", F(size=9, bold=True), AL_C, B_ALL, fill("D9D9D9"))
    ws.column_dimensions[CL(cB)].width = 26
    rows = [e for e in core.get().eq if EQ.alive(e, min(v, 2021))]
    rows.sort(key=lambda e: (SETSUBI_ORDER.index(e["setsubi"]), UNIT_ORDER[e["unit"]], e["sysno"], e["loop"], e["tag"]))
    r = 7
    n = 0
    group = None
    sub = {}
    skip = ("C-P-820A", 2020) if v >= 2017 else None
    first = True

    def flush(g, r):
        merge(ws, r, 1, r, 8, f"小計（千円）　{g}", F(size=8, bold=True), AL_R, B_ALL, fill("F2F2F2"))
        cost = {}
        for y in OLD_YEARS:
            put(ws, r, ycol(y), None, None, None, B_ALL, fill("F2F2F2"))
            put(ws, r, ycol(y, 1), None, None, None, B_ALL, fill("F2F2F2"))
            val = int(round(sub.get(y, 0))) if sub.get(y) else None
            put(ws, r, ycol(y, 2), val, F(size=6), AL_R, B_ALL, fill("F2F2F2"), nf="#,##0")
            if val:
                cost[str(y)] = val
        put(ws, r, cB, None, None, None, B_ALL, fill("F2F2F2"))
        ans.rec(fn, "整備計画", r, "subtotal", dict(path=[g], cost=cost, unit="千円"), check=["小計"])
        return r + 1

    for e in rows:
        if e["setsubi"] != group:
            if group is not None:
                r = flush(group, r)
            group = e["setsubi"]
            sub = {}
            merge(ws, r, 1, r, cB, f"【{group}】", F(size=9, bold=True), AL_L, B_ALL, fill("EDEDED"))
            ans.rec(fn, "整備計画", r, "parent", dict(path=[group]), check=[group])
            r += 1
            first = True
        n += 1
        orig = W.simulate(e, cutoff=2013, uversion="none", until=2027)
        act = W.simulate(e, cutoff=None, until=v - 1)
        plan = W.simulate(e, cutoff=v, uversion=uv(v), until=Y1)
        vals = [n, e["setsubi"] if first else "〃", e["name"], e["unit"] if e["unit"] != "C" else "共", e["tag"], e["imp"], health(e, v, plan), hozen(e)]
        for c, val in enumerate(vals, start=1):
            put(ws, r, c, val, F(size=8), AL_C if c in (1, 2, 4, 6, 7, 8) else AL_L, B_HAIR)
        marks, actual, costs = {}, {}, {}
        for y in OLD_YEARS:
            cp = put(ws, r, ycol(y), None, F(size=8), AL_C, Border(left=thin, right=dot, top=hair, bottom=hair))
            ca = put(ws, r, ycol(y, 1), None, F(size=8), AL_C, Border(left=dot, right=dot, top=hair, bottom=hair))
            cc = put(ws, r, ycol(y, 2), None, F(size=6), AL_R, Border(left=dot, right=thin, top=hair, bottom=hair), nf="#,##0")
            if y < v:
                ok = kinds(orig.get(y, []))
                if ok:
                    cp.value = symstr(ok, SYM_OLD)
                    marks[str(y)] = cp.value
                if y == 2017:
                    continue
                ak = kinds(act.get(y, []))
                if ak:
                    ca.value = symstr(ak, SYM_OLD)
                    actual[str(y)] = ca.value
                    if set(ak) - set(ok):
                        ca.font = F(size=8, color="FF0000")
                    val = cost_of(act.get(y, []))
                    if val:
                        cc.value = int(round(val))
                        costs[str(y)] = cc.value
                    sub[y] = sub.get(y, 0) + val
            else:
                ks = kinds(plan.get(y, []))
                if ks:
                    cp.value = symstr(ks, SYM_OLD)
                    marks[str(y)] = cp.value
                    if "U" in ks:
                        cp.font = F(size=8, color="FF0000")
                    val = cost_of(plan.get(y, []))
                    if val:
                        cc.value = int(round(val))
                        costs[str(y)] = cc.value
                    if skip != (e["tag"], y):
                        sub[y] = sub.get(y, 0) + val
        b = "基幹改良で更新予定（範囲検討中）" if W.upgrade_year(e, uv(v)) else ""
        if e["tag"] == "C-ST-501":
            b = "定期事業者検査は別契約（4 年ごと）"
        put(ws, r, cB, b or None, F(size=7), AL_L, B_HAIR)
        ans.rec(fn, "整備計画", r, "item", dict(path=[e["setsubi"], e["name"]], code=e["tag"], work=cycle_txt(e), cycle=cyc(e), cost=costs,
                                              unit="千円", marks=marks, actual=actual, memo=b or None, comments={},
                                              号炉=e["unit"], 重要度=e["imp"], 健全度=vals[6], 保全方式=vals[7]), check=[e["tag"]])
        first = False
        r += 1
    r = flush(group, r)
    r += 1
    merge(ws, r, 1, r, cB, "【法定点検・検査】", F(size=9, bold=True), AL_L, B_ALL, fill("EDEDED"))
    ans.rec(fn, "整備計画", r, "parent", dict(path=["法定点検・検査"]), check=["法定点検"])
    r += 1
    for (nm, law, cy, ys) in legal_rows("2019"):
        merge(ws, r, 1, r, 2, law, F(size=7), AL_L, B_HAIR)
        merge(ws, r, 3, r, 5, nm, F(size=8), AL_L, B_HAIR)
        merge(ws, r, 6, r, 8, cy, F(size=8), AL_C, B_HAIR)
        marks = {}
        for y in OLD_YEARS:
            put(ws, r, ycol(y), "●" if ys.get(y) else None, F(size=8), AL_C, B_HAIR)
            put(ws, r, ycol(y, 1), "●" if (ys.get(y) and y < v and y != 2017) else None, F(size=8), AL_C, B_HAIR)
            put(ws, r, ycol(y, 2), None, None, None, B_HAIR)
            if ys.get(y):
                marks[str(y)] = "●"
        put(ws, r, cB, None, None, None, B_HAIR)
        ans.rec(fn, "整備計画", r, "item", dict(path=["法定点検・検査", nm], code=None, work=law, cycle=cy, marks=marks), check=[nm[:8]])
        r += 1
    ws.freeze_panes = f"{CL(YC)}7"
    page(ws, landscape=True, paper=8, fit_w=1, fit_h=0, area=f"A1:{CL(cB)}{r - 1}", titles="5:6", footer="&P / &N")
    props(wb, f"【架空】みなと臨海工場 施設保全計画（長寿命化総合計画）{ver}", "h3", "中長期整備計画", dt.datetime(2013, 12, 2, 10, 0), modified_by=who)
    wb.save(os.path.join(OUT, fn))
    return fn


# ================================================================ h4（R2〜R8）
def build_h4(v):
    fn = fn_of(v, "h4")
    who = "h4" if v <= 2021 else "h5"
    ans.start(fn, True, "中長期整備計画", v, "延命化計画（h4 R2、1 シート・旧 3 段列を非表示）", who)
    wb = Workbook()
    wb.remove(wb.active)
    F = lambda **k: font(FONT_P, **k)
    G = lambda **k: font(FONT_G, **k)
    new_years = list(range(2028, 2041))
    years_vis = list(range(2020, 2041))
    col_of = {y: ycol(y) for y in OLD_YEARS}
    for i, y in enumerate(new_years):
        col_of[y] = ycol(Y1, 2) + 1 + i
    cPrev = ycol(Y1, 2) + 1 + len(new_years)
    cB = cPrev + 1
    ws = wb.create_sheet("延命化計画（星取表）")
    put(ws, 1, 1, "【架空】みなと臨海工場　延命化計画　機器別整備計画（星取表）", F(size=14, bold=True))
    put(ws, 2, 1, f"{core.wa_fy(v)}版（施設課 {who}）　計画期間 2020〜2040 年度（延命目標 2040 年度）" + ("　R2.11 策定 h4" if v > 2020 else ""), F(size=10))
    put(ws, 3, 1, "凡例　○：点検・計測　△：補修　◎：分解整備　■：更新　▲：基幹的設備改良工事で施工（R4〜R7）　過年度の列は当初計画（R2 版）のまま", F(size=9))
    hdr = ["No.", "系統", "機器名称", "号炉", "タグNo.", "重要\n度", "健全\n度", "保全\n方式"]
    for i, (h, w) in enumerate(zip(hdr, [5, 11, 28, 4.5, 11, 4, 4, 4.5]), start=1):
        merge(ws, 5, i, 7, i, h, F(size=9, bold=True), AL_C, B_ALL, fill("DDEBF7"))
        ws.column_dimensions[CL(i)].width = w
    for y in OLD_YEARS:
        for k in range(3):
            c = ycol(y, k)
            if y < 2020 or k > 0:
                put(ws, 6, c, (fy_wa(y) if k == 0 else ("実績", "費用")[k - 1]), G(size=7), AL_C, B_ALL, fill("D9D9D9"))
                ws.column_dimensions[CL(c)].hidden = True
    for y in years_vis:
        c = col_of[y]
        put(ws, 6, c, y, F(size=8, bold=True), AL_C, B_ALL, fill("DDEBF7"))
        put(ws, 7, c, fy_wa(y), F(size=8), AL_C, B_ALL, fill("DDEBF7" if y >= v else "E7E6E6"))
        ws.column_dimensions[CL(c)].width = 4.6 if y <= 2027 else 4.2
    put(ws, 5, col_of[2022], "基幹改良（R4〜R7）", F(size=8, bold=True), AL_L, B_ALL, fill("FCE4D6"))
    put(ws, 5, col_of[2026], "延命化後", F(size=8), AL_L, B_ALL, fill("DDEBF7"))
    merge(ws, 5, cPrev, 7, cPrev, "前回実施\n（年度・内容）", F(size=8, bold=True), AL_C, B_ALL, fill("DDEBF7"))
    merge(ws, 5, cB, 7, cB, "備　考", F(size=9, bold=True), AL_C, B_ALL, fill("DDEBF7"))
    ws.column_dimensions[CL(cPrev)].width = 20
    ws.column_dimensions[CL(cB)].width = 24
    note(ws.cell(row=6, column=col_of[2020]), "H31版をコピー。旧 3 段（計画・実績・費用）は非表示で残し、R10 以降を右に追加（h4）", "h4")
    rows = known_rows(v)
    rows.sort(key=lambda e: (SETSUBI_ORDER.index(e["setsubi"]), e["sysno"] if e["unit"] != "C" else 100 + e["sysno"], e["keito"], e["loop"],
                             UNIT_ORDER[e["unit"]], e["tag"]))
    r = 8
    cur = (None, None)
    n = 0
    act_until = 2020 if v == 2022 else v - 1  # R4 版（h5 の引継ぎ）は前回実施を R2 までしか直していない
    left = False
    for e in rows:
        u, s, k = e["unit"], e["setsubi"], e["keito"]
        if s != cur[0]:
            if cur[0] == "受入供給設備" and not left:
                merge(ws, r, 1, r, 8, "小計（千円）　受入供給設備", G(size=8, color="808080"), AL_R, B_ALL)
                for y in range(2020, 2028):
                    put(ws, r, col_of[y], [1260, 3450, 980, 2310, 1260, 4120, 980, 2310][y - 2020], G(size=6, color="808080"), AL_R, B_ALL, nf="#,##0")
                ans.rec(fn, ws.title, r, "subtotal", dict(path=["受入供給設備"], cost={str(y): [1260, 3450, 980, 2310, 1260, 4120, 980, 2310][y - 2020] for y in range(2020, 2028)},
                                                           unit="千円", memo="旧様式の小計行（消し忘れ・値は H31 版のまま）"), check=["小計"])
                ws.row_dimensions[r].outline_level = 1
                r += 1
                left = True
            merge(ws, r, 1, r, cB, f"■ {s}", F(size=9, bold=True), AL_L, B_ALL, fill("BDD7EE"))
            ans.rec(fn, ws.title, r, "parent", dict(path=[s]), check=[s])
            r += 1
            cur = (s, None)
        if k != cur[1]:
            merge(ws, r, 1, r, cB, f"　　{k}（{'共通' if u == 'C' else '1〜3号炉'}）", F(size=8, bold=True), AL_L, B_ALL, fill("F2F2F2"))
            ws.row_dimensions[r].outline_level = 1
            ans.rec(fn, ws.title, r, "parent", dict(path=[s, k]), check=[k])
            r += 1
            cur = (s, k)
        n += 1
        plan_now = W.simulate(e, cutoff=v, uversion="2020", until=2040)
        plan_r2 = W.simulate(e, cutoff=2020, uversion="2020", until=2040)
        act = W.simulate(e, cutoff=None, until=act_until)
        vals = [n, k, "　" + e["name"], u if u != "C" else "共", e["tag"], e["imp"], health(e, v, plan_now), hozen(e)]
        for c, val in enumerate(vals, start=1):
            put(ws, r, c, val, F(size=8), AL_C if c in (1, 4, 6, 7, 8) else AL_L, B_HAIR)
        orig = W.simulate(e, cutoff=2013, uversion="2019", until=2019)
        for y in range(2013, 2020):
            ks = kinds(orig.get(y, []))
            put(ws, r, ycol(y), symstr(ks, SYM_OLD) if ks else None, G(size=8), AL_C, B_HAIR)
        marks = {}
        for y in years_vis:
            src = plan_r2 if y < v else plan_now
            ks = kinds(src.get(y, []), new=True, e=e)
            if e["added"] and y < e["added"]:
                ks = []
            cell = put(ws, r, col_of[y], symstr(ks, SYM_NEW) if ks else None, F(size=9), AL_C, B_HAIR)
            if "U" in ks:
                cell.font = F(size=9, bold=True, color="C00000")
            if cell.value:
                marks[str(y)] = cell.value
        prev = [(y, a) for y, a in act.items() if any(kk in ("O", "R", "U") for kk, _, _ in a)]
        ptxt = ""
        if prev:
            y, a = prev[-1]
            kk = [x for x in a if x[0] in ("O", "R", "U")][0]
            ptxt = f"{fy_wa(y)} " + (kk[2] or {"O": "分解整備", "R": "更新", "U": "基幹改良"}[kk[0]])
        put(ws, r, cPrev, ptxt or None, F(size=7), AL_L, B_HAIR)
        b = "基幹改良（R4〜R7）対象" if W.upgrade_year(e, "2020") else ""
        if e["added"]:
            b = f"{fy_wa(e['added'])} 基幹改良で新設"
        put(ws, r, cB, b or None, F(size=7), AL_L, B_HAIR)
        ws.row_dimensions[r].outline_level = 2
        ans.rec(fn, ws.title, r, "item", dict(path=[s, k, e["name"]], code=e["tag"], work=cycle_txt(e), cycle=cyc(e), cost=None, unit=None,
                                              marks=marks, memo=b or None, comments={}, 前回実施=ptxt or None, 号炉=u, 重要度=e["imp"],
                                              健全度=vals[6], 保全方式=vals[7]), check=[e["tag"]])
        r += 1
    r += 1
    merge(ws, r, 1, r, cB, "■ 法定点検・検査", F(size=9, bold=True), AL_L, B_ALL, fill("BDD7EE"))
    ans.rec(fn, ws.title, r, "parent", dict(path=["法定点検・検査"]), check=["法定点検"])
    r += 1
    for (nm, law, cy, ys) in legal_rows("2020"):
        merge(ws, r, 1, r, 2, cy, F(size=8), AL_C, B_HAIR)
        merge(ws, r, 3, r, 5, nm, F(size=8), AL_L, B_HAIR)
        merge(ws, r, 6, r, 8, "", None, None, B_HAIR)
        marks = {}
        for y in years_vis:
            put(ws, r, col_of[y], "法" if ys.get(y) else None, F(size=8), AL_C, B_HAIR)
            if ys.get(y):
                marks[str(y)] = "法"
        put(ws, r, cPrev, None, None, None, B_HAIR)
        put(ws, r, cB, law, F(size=7), AL_L, B_HAIR)
        ans.rec(fn, ws.title, r, "item", dict(path=["法定点検・検査", nm], code=None, work=law, cycle=cy, marks=marks), check=[nm[:8]])
        r += 1
    ws.sheet_properties.outlinePr.summaryBelow = False
    ws.freeze_panes = f"{CL(col_of[2020])}8"
    page(ws, landscape=True, paper=8, fit_w=1, fit_h=0, area=f"A1:{CL(cB)}{r - 1}", titles="5:7", footer="&P / &N")
    # 費用集計
    w2 = wb.create_sheet("費用集計")
    put(w2, 1, 1, "【架空】みなと臨海工場　延命化計画　年度別整備費（概算）", F(size=13, bold=True))
    put(w2, 2, 1, f"{core.wa_fy(v)}版　単位：百万円（税抜）　※ 機器別整備費を号炉・設備区分で積上げ。基幹的設備改良工事は別行", F(size=9))
    put(w2, 4, 1, "号炉", F(size=9, bold=True), AL_C, B_ALL, fill("DDEBF7"))
    put(w2, 4, 2, "設備区分", F(size=9, bold=True), AL_C, B_ALL, fill("DDEBF7"))
    w2.column_dimensions["A"].width = 7
    w2.column_dimensions["B"].width = 16
    for i, y in enumerate(years_vis):
        put(w2, 4, 3 + i, f"{y}\n{fy_wa(y)}", F(size=8, bold=True), AL_C, B_ALL, fill("DDEBF7"))
        w2.column_dimensions[CL(3 + i)].width = 6.5
    costs = {}
    for e in rows:
        pl = W.simulate(e, cutoff=v, uversion="2020", until=2040)
        for y in years_vis:
            costs.setdefault((e["unit"], e["setsubi"]), {}).setdefault(y, 0)
            costs[(e["unit"], e["setsubi"])][y] += cost_of(pl.get(y, [])) if y >= v else 0
    rr = 5
    for (u, s), d in sorted(costs.items(), key=lambda x: (UNIT_ORDER[x[0][0]], SETSUBI_ORDER.index(x[0][1]))):
        put(w2, rr, 1, UN[u], F(size=8), AL_C, B_HAIR)
        put(w2, rr, 2, s, F(size=8), AL_L, B_HAIR)
        cost = {}
        for i, y in enumerate(years_vis):
            val = round(d.get(y, 0) / 1000, 1)
            put(w2, rr, 3 + i, val if val else None, F(size=8), AL_R, B_HAIR, nf="0.0")
            if val:
                cost[str(y)] = val
        ans.rec(fn, "費用集計", rr, "item", dict(path=[UN[u], s], cost=cost, unit="百万円"), check=[s])
        rr += 1
    kikan = {2022: 1820.0, 2023: 2140.0, 2024: 2210.0, 2025: 1990.0}
    put(w2, rr, 2, "基幹的設備改良工事", F(size=8, bold=True), AL_L, B_ALL, fill("FCE4D6"))
    for i, y in enumerate(years_vis):
        put(w2, rr, 3 + i, kikan.get(y), F(size=8), AL_R, B_ALL, nf="#,##0.0")
    ans.rec(fn, "費用集計", rr, "item", dict(path=["基幹的設備改良工事"], cost={str(k): x for k, x in kikan.items()}, unit="百万円"), check=["基幹的設備改良工事"])
    put(w2, rr + 2, 1, f"※ 過年度（{v} 年度より前）は実績のため本表では空欄。", F(size=8))
    page(w2, landscape=True, paper=8, fit_w=1, fit_h=1, area=f"A1:{CL(2 + len(years_vis))}{rr + 2}")
    props(wb, f"【架空】みなと臨海工場 延命化計画（{core.wa_fy(v)}版）", "h3", "中長期整備計画", dt.datetime(2013, 12, 2, 10, 0), modified_by=who)
    wb.save(os.path.join(OUT, fn))
    return fn


def main():
    out = []
    for v in range(2008, 2027):
        out.append(build_h2(v) if v <= 2012 else build_h3(v) if v <= 2019 else build_h4(v))
    return out


if __name__ == "__main__":
    import time
    t0 = time.time()
    print(len(main()), time.time() - t0)
