# -*- coding: utf-8 -*-
"""中長期整備計画（星取表）
旧：施設保全計画（長寿命化総合計画）平成31年度版　h3 が H25 に作った様式を h4 が更新したもの
新：延命化計画 令和2年度版　h4 が旧様式をコピーして直したもの（旧年度列は非表示で残る）
"""
import datetime as dt
import os

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font
from openpyxl.utils import get_column_letter as CL

import equipment as EQ
import world as W
from s_ledger import SETSUBI_ORDER, UNIT_ORDER, hozen
from xlhelp import *

SYM_OLD = {"I": "●", "O": "◎", "R": "★", "U": "▲"}
SYM_NEW = {"I": "○", "O": "◎", "R": "■", "U": "▲", "H": "△"}
ORDER = {"U": 0, "R": 1, "O": 2, "H": 2, "I": 3}

REPAIR_CLS = {"furnace", "boiler", "scrubber", "stack_in", "hv", "hv_hp", "hv_corr", "static_hot"}


def kinds(acts, new=False, e=None):
    ks = []
    for k, cost, note in acts:
        if new and k == "O" and (e["cls"] in REPAIR_CLS or "補修" in note):
            k = "H"
        ks.append(k)
    return ks


def symstr(ks, table):
    ks = sorted(set(ks), key=lambda k: ORDER[k])
    # 1 マスに 2 つまで（主たる作業を先に）
    return "".join(table[k] for k in ks[:2])


def cost_of(acts):
    return sum(c for k, c, n in acts if k != "U")


def health(e, fy, plan):
    if e["cls"].startswith("inst"):
        yrs = [y for y in range(fy, fy + 4) if any(k == "U" for k, _, _ in plan.get(y, []))]
    else:
        yrs = [y for y in range(fy, fy + 4) if any(k in ("R", "U") for k, _, _ in plan.get(y, []))]
    if yrs:
        return 2
    f = [y for y in W.findings(e) if fy - 3 <= y < fy]
    if f or any("減肉" in n or "漏れ" in n or "振動" in n for y, k, n in W.FIXED.get(e["tag"], []) if fy - 4 <= y < fy):
        return 3
    if fy - e["installed"] >= 12 and e["cls"] in ("grate", "grate_wc", "boiler", "sh", "sh2", "sh3", "rh", "idf", "scrubber",
                                                      "scr", "conveyor", "pump_corr", "crane", "kneader", "furnace"):
        return 3
    return 4


# ---------------------------------------------------------------- 法定点検（設備に属さない行）
def legal_rows(version):
    """(名称, 根拠, 周期, {年度: 記号}) version='2019'（旧）/'2020'（新）"""
    rows = []
    rng = range(2004, 2041)
    def yrs(f):
        return {y: True for y in rng if f(y)}
    rows.append(("ボイラ 定期事業者検査（1号炉）", "電気事業法", "2年", yrs(lambda y: y % 2 == 0)))
    rows.append(("ボイラ 定期事業者検査（2号炉）", "電気事業法", "2年", yrs(lambda y: y % 2 == 1)))
    rows.append(("ボイラ 定期事業者検査（3号炉）", "電気事業法", "2年", yrs(lambda y: y % 2 == 0)))
    rows.append(("蒸気タービン 定期事業者検査", "電気事業法", "4年", yrs(lambda y: (y - 2006) % 4 == 0)))
    rows.append(("第一種圧力容器 性能検査（脱気器・空気予熱器 3基）", "労働安全衛生法", "1年", yrs(lambda y: True)))
    rows.append(("ごみクレーン 性能検査（1〜3号）", "クレーン則", "2年", yrs(lambda y: y % 2 == 0)))
    rows.append(("ごみクレーン 年次自主検査", "クレーン則", "1年", yrs(lambda y: True)))
    if version != "2019":
        rows.append(("灰クレーン 性能検査", "クレーン則", "2年", yrs(lambda y: y % 2 == 1)))
    rows.append(("保守用クレーン 性能検査（炉室 10t・タービン室 30t・IDF 室 3t）", "クレーン則", "2年", yrs(lambda y: y % 2 == 0)))
    rows.append(("特別高圧受変電設備 年次点検（月次は保守点検委託）", "電気事業法（保安規程）", "1年", yrs(lambda y: True)))
    rows.append(("非常用発電機 点検・負荷運転", "消防法・保安規程", "1年", yrs(lambda y: True)))
    rows.append(("排ガス測定（ばいじん・NOx・SOx・HCl　年6回）", "大気汚染防止法", "2月", yrs(lambda y: True)))
    rows.append(("排ガス ダイオキシン類測定", "ダイオキシン類対策特別措置法", "1年", yrs(lambda y: True)))
    rows.append(("ダイオキシン類 作業環境測定（年2回）", "労働安全衛生規則", "6月", yrs(lambda y: True)))
    rows.append(("ごみ計量機 定期検査（2基）", "計量法", "2年", yrs(lambda y: y % 2 == 0)))
    rows.append(("消防用設備等点検（機器 6月・総合 1年）", "消防法", "1年", yrs(lambda y: True)))
    if version == "2019":
        # h3 の入力：H28 の次を H32 にしてしまっている（3 年を超える）
        rows.append(("精密機能検査", "廃棄物処理法施行規則", "3年", {2013: True, 2016: True, 2020: True, 2023: True, 2026: True}))
    else:
        rows.append(("精密機能検査", "廃棄物処理法施行規則", "3年", yrs(lambda y: (y - 2013) % 3 == 0)))
    return rows


# ---------------------------------------------------------------- 旧様式（h3：年度ごとに 計画・実績・費用 の 3 段）
UNAME = {"1": "1号炉", "2": "2号炉", "3": "3号炉", "C": "共通"}
Y0, Y1 = 2013, 2027
OLD_YEARS = list(range(Y0, Y1 + 1))
YC = 9  # I 列 = H25 の「計画」


def ycol(y, k=0):
    """旧様式の列：年度 y の k 段目（0 計画・1 実績・2 費用）"""
    return YC + (y - Y0) * 3 + k


def build_old(eq, path):
    wb = Workbook()
    ws = wb.active
    ws.title = "整備計画"
    F = lambda **k: font(FONT_G, **k)
    cB = ycol(Y1, 2) + 1
    put(ws, 1, 1, "【架空】みなと臨海工場　施設保全計画（長寿命化総合計画）　整備スケジュール表", F(size=14, bold=True))
    put(ws, 2, 1, "平成31年度版　（H25 策定 施設課 h3 ／ H31.4 更新 h4）　各年度：計画・実績・費用（千円）", F(size=10))
    put(ws, 3, 1, "凡例　●：点検・検査　◎：整備（分解・補修）　★：更新　　実績欄は実施した作業を記入　費用は計画年度＝概算、過年度＝実績（千円）", F(size=9))
    put(ws, 3, 30, "▲：基幹的設備改良（予定・検討中）", F(size=9, color="FF0000"))
    note(ws.cell(row=3, column=30), "H31 追加。基幹改良の範囲・年度は R2 延命化計画で確定させる（h4）", "h4")
    hdr = ["No.", "設備区分", "機器名称", "号炉", "タグNo.", "重要\n度", "健全\n度", "保全\n方式"]
    wid = [5, 12, 26, 4.5, 11, 4, 4, 4.5]
    for i, (h, w) in enumerate(zip(hdr, wid), start=1):
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
    note(ws.cell(row=5, column=ycol(2019)), "H31 → 元号が変わったが年度表記は H31 のまま（h4）", "h4")

    rows = [e for e in eq if EQ.alive(e, 2019)]
    rows.sort(key=lambda e: (SETSUBI_ORDER.index(e["setsubi"]), UNIT_ORDER[e["unit"]], e["sysno"], e["loop"], e["tag"]))
    r = 7
    n = 0
    group = None
    sub = {}
    skip_cost = {"排水処理設備": ("C-P-820A", 2020)}  # 小計に入っていない（行を足したとき式を直さなかった）
    first_in_group = True

    def flush_sub(g, r):
        merge(ws, r, 1, r, 8, f"小計（千円）　{g}", F(size=8, bold=True), AL_R, B_ALL, fill("F2F2F2"))
        for y in OLD_YEARS:
            put(ws, r, ycol(y), None, None, None, B_ALL, fill("F2F2F2"))
            put(ws, r, ycol(y, 1), None, None, None, B_ALL, fill("F2F2F2"))
            v = sub.get(y, 0)
            put(ws, r, ycol(y, 2), int(round(v)) if v else None, F(size=6), AL_R, B_ALL, fill("F2F2F2"), nf="#,##0")
        put(ws, r, cB, None, None, None, B_ALL, fill("F2F2F2"))
        return r + 1

    for e in rows:
        if e["setsubi"] != group:
            if group is not None:
                r = flush_sub(group, r)
            group = e["setsubi"]
            sub = {}
            merge(ws, r, 1, r, cB, f"【{group}】", F(size=9, bold=True), AL_L, B_ALL, fill("EDEDED"))
            r += 1
            first_in_group = True
        n += 1
        orig = W.simulate(e, cutoff=2013, uversion="2019", until=2018)
        act = W.simulate(e, cutoff=None, until=2018)
        plan = W.simulate(e, cutoff=2019, uversion="2019", until=Y1)
        vals = [n, e["setsubi"] if first_in_group else "〃", e["name"], e["unit"] if e["unit"] != "C" else "共", e["tag"], e["imp"],
                health(e, 2019, plan), hozen(e)]
        for c, v in enumerate(vals, start=1):
            put(ws, r, c, v, F(size=8), AL_C if c in (1, 2, 4, 6, 7, 8) else AL_L, B_HAIR)
        for y in OLD_YEARS:
            cp = put(ws, r, ycol(y), None, F(size=8), AL_C, Border(left=thin, right=dot, top=hair, bottom=hair))
            ca = put(ws, r, ycol(y, 1), None, F(size=8), AL_C, Border(left=dot, right=dot, top=hair, bottom=hair))
            cc = put(ws, r, ycol(y, 2), None, F(size=6), AL_R, Border(left=dot, right=thin, top=hair, bottom=hair), nf="#,##0")
            if y < 2019:
                ok = kinds(orig.get(y, []))
                ak = kinds(act.get(y, []))
                if ok:
                    cp.value = symstr(ok, SYM_OLD)
                if y == 2017:  # H29（h3→h4 の交代年）は実績・費用の記入がない
                    continue
                if ak:
                    ca.value = symstr(ak, SYM_OLD)
                    if set(ak) - set(ok):
                        ca.font = F(size=8, color="FF0000")
                    v = cost_of(act.get(y, []))
                    cc.value = int(round(v)) if v else None
                    sub[y] = sub.get(y, 0) + v
            else:
                ks = kinds(plan.get(y, []))
                if ks:
                    cp.value = symstr(ks, SYM_OLD)
                    if "U" in ks:
                        cp.font = F(size=8, color="FF0000")
                    v = cost_of(plan.get(y, []))
                    cc.value = int(round(v)) if v else None
                    if not (skip_cost.get(group) == (e["tag"], y)):
                        sub[y] = sub.get(y, 0) + v
        b = ""
        if W.upgrade_year(e, "2019"):
            b = "基幹改良で更新予定（範囲検討中）"
        if e["tag"] == "2-SH-304":
            b = "H29 減肉管 6 本取替。改良まで毎年肉厚測定"
            note(ws.cell(row=r, column=ycol(2017, 1)), "H28 定修 UT で最小 3.9mm（管理値 4.0）。H29 部分取替（h3 メモ）", "h3")
        if e["tag"] == "C-ST-501":
            b = "定期事業者検査は別契約（4 年ごと）"
        if e["tag"] == "3-IDF-607":
            b = "H30 振動上昇（7.1mm/s）→ R2 軸受交換予定"
        put(ws, r, cB, b or None, F(size=7), AL_L, B_HAIR)
        first_in_group = False
        r += 1
    r = flush_sub(group, r)
    r += 1
    merge(ws, r, 1, r, cB, "【法定点検・検査】", F(size=9, bold=True), AL_L, B_ALL, fill("EDEDED"))
    r += 1
    for (nm, law, cyc, ys) in legal_rows("2019"):
        merge(ws, r, 1, r, 2, law, F(size=7), AL_L, B_HAIR)
        merge(ws, r, 3, r, 5, nm, F(size=8), AL_L, B_HAIR)
        merge(ws, r, 6, r, 8, cyc, F(size=8), AL_C, B_HAIR)
        for y in OLD_YEARS:
            put(ws, r, ycol(y), "●" if ys.get(y) else None, F(size=8), AL_C, B_HAIR)
            put(ws, r, ycol(y, 1), "●" if (ys.get(y) and y < 2019 and y != 2017) else None, F(size=8), AL_C, B_HAIR)
            put(ws, r, ycol(y, 2), None, None, None, B_HAIR)
        put(ws, r, cB, None, None, None, B_HAIR)
        r += 1
    last = r - 1
    ws.freeze_panes = f"{CL(YC)}7"
    page(ws, landscape=True, paper=8, fit_w=1, fit_h=0, area=f"A1:{CL(cB)}{last}", titles="5:6",
         footer="&P / &N", header="【架空】みなと臨海工場 施設保全計画 H31")
    props(wb, "【架空】みなと臨海工場 施設保全計画（長寿命化総合計画）平成31年度版", "h3", "中長期整備計画",
          dt.datetime(2013, 12, 2, 10, 0), modified_by="h4")
    wb.save(path)
    return n


# ---------------------------------------------------------------- 新様式（h4：号炉・共通ごとのシート＋法定点検＋費用集計）
def build_new(eq, path):
    wb = Workbook()
    wb.remove(wb.active)
    F = lambda **k: font(FONT_P, **k)
    G = lambda **k: font(FONT_G, **k)
    new_years = list(range(2028, 2041))
    years_vis = list(range(2020, 2041))
    col_of = {y: ycol(y) for y in OLD_YEARS}
    base_new = ycol(Y1, 2) + 1
    for i, y in enumerate(new_years):
        col_of[y] = base_new + i
    cPrev = base_new + len(new_years)
    cB = cPrev + 1
    costs = {}
    n = 0
    rows_all = [e for e in eq if e["installed"] <= 2020 and (e["removed"] is None or e["removed"] > 2020)]
    rows_all.sort(key=lambda e: (UNIT_ORDER[e["unit"]], SETSUBI_ORDER.index(e["setsubi"]), e["sysno"], e["loop"], e["tag"]))

    def header(ws, title):
        put(ws, 1, 1, f"【架空】みなと臨海工場　延命化計画　機器別整備計画（星取表）　{title}", F(size=14, bold=True))
        put(ws, 2, 1, "令和2年度版（R2.11 施設課 h4）　計画期間 2020〜2040 年度（延命目標 2040 年度）", F(size=10))
        put(ws, 3, 1, "凡例　○：点検・計測　△：補修　◎：分解整備　■：更新　▲：基幹的設備改良工事で施工（R4〜R7）　費用は「費用集計」シート", F(size=9))
        hdr = ["No.", "系統", "機器名称", "号炉", "タグNo.", "重要\n度", "健全\n度", "保全\n方式"]
        wid = [5, 11, 28, 4.5, 11, 4, 4, 4.5]
        for i, (h, w) in enumerate(zip(hdr, wid), start=1):
            merge(ws, 5, i, 7, i, h, F(size=9, bold=True), AL_C, B_ALL, fill("DDEBF7"))
            ws.column_dimensions[CL(i)].width = w
        # 旧様式の 3 段の列は残す：H25〜R1 は 3 段とも非表示、R2〜R9 は「計画」だけ表示し 実績・費用 は非表示
        for y in OLD_YEARS:
            for k in range(3):
                c = ycol(y, k)
                if y < 2020 or k > 0:
                    put(ws, 6, c, (fy_wa(y) if k == 0 else ("実績", "費用")[k - 1]), G(size=7), AL_C, B_ALL, fill("D9D9D9"))
                    ws.column_dimensions[CL(c)].hidden = True
        for y in years_vis:
            c = col_of[y]
            put(ws, 6, c, y, F(size=8, bold=True), AL_C, B_ALL, fill("DDEBF7"))
            put(ws, 7, c, fy_wa(y), F(size=8), AL_C, B_ALL, fill("DDEBF7"))
            ws.column_dimensions[CL(c)].width = 4.6 if y <= 2027 else 4.2
        for y in range(2022, 2026):
            ws.cell(row=6, column=col_of[y]).fill = fill("FCE4D6")
            ws.cell(row=7, column=col_of[y]).fill = fill("FCE4D6")
        put(ws, 5, col_of[2020], "", None, None, B_ALL, fill("DDEBF7"))
        put(ws, 5, col_of[2021], "", None, None, B_ALL, fill("DDEBF7"))
        put(ws, 5, col_of[2022], "基幹改良（R4〜R7）", F(size=8, bold=True), AL_L, B_ALL, fill("FCE4D6"))
        put(ws, 5, col_of[2026], "延命化後", F(size=8), AL_L, B_ALL, fill("DDEBF7"))
        merge(ws, 5, cPrev, 7, cPrev, "前回実施\n（年度・内容）", F(size=8, bold=True), AL_C, B_ALL, fill("DDEBF7"))
        merge(ws, 5, cB, 7, cB, "備　考", F(size=9, bold=True), AL_C, B_ALL, fill("DDEBF7"))
        ws.column_dimensions[CL(cPrev)].width = 20
        ws.column_dimensions[CL(cB)].width = 24
        note(ws.cell(row=6, column=col_of[2020]), "H31版をコピー。旧 3 段（計画・実績・費用）は非表示で残し、R10 以降を右に追加（h4）", "h4")

    ws = wb.create_sheet("延命化計画（星取表）")
    header(ws, "")
    rows1 = sorted(rows_all, key=lambda e: (SETSUBI_ORDER.index(e["setsubi"]), e["sysno"] if e["unit"] != "C" else 100 + e["sysno"],
                                            e["keito"], e["loop"], UNIT_ORDER[e["unit"]], e["tag"]))
    r = 8
    cur = (None, None)
    left_sub_done = False
    for e in rows1:
        u, s, k = e["unit"], e["setsubi"], e["keito"]
        if s != cur[0]:
            if cur[0] == "受入供給設備" and not left_sub_done:
                merge(ws, r, 1, r, 8, "小計（千円）　受入供給設備", G(size=8, color="808080"), AL_R, B_ALL)
                for y in range(2020, 2028):
                    put(ws, r, col_of[y], [1260, 3450, 980, 2310, 1260, 4120, 980, 2310][y - 2020], G(size=6, color="808080"), AL_R, B_ALL, nf="#,##0")
                ws.row_dimensions[r].outline_level = 1
                r += 1
                left_sub_done = True
            merge(ws, r, 1, r, cB, f"■ {s}", F(size=9, bold=True), AL_L, B_ALL, fill("BDD7EE"))
            r += 1
            cur = (s, None)
        if k != cur[1]:
            merge(ws, r, 1, r, cB, f"　　{k}（{'共通' if u == 'C' else '1〜3号炉'}）", F(size=8, bold=True), AL_L, B_ALL, fill("F2F2F2"))
            ws.row_dimensions[r].outline_level = 1
            r += 1
            cur = (s, k)
        n += 1
        plan = W.simulate(e, cutoff=2020, uversion="2020", until=2040)
        act = W.simulate(e, cutoff=None, until=2019)
        vals = [n, k, "　" + e["name"], e["unit"] if u != "C" else "共", e["tag"], e["imp"], health(e, 2020, plan), hozen(e)]
        for c, v in enumerate(vals, start=1):
            put(ws, r, c, v, F(size=8), AL_C if c in (1, 4, 6, 7, 8) else AL_L, B_HAIR)
        if u != "C":
            ws.cell(row=r, column=4).fill = fill({"1": "FFFFFF", "2": "F3F7FB", "3": "E9F0F8"}[u])
        orig = W.simulate(e, cutoff=2013, uversion="2019", until=2019)
        for y in range(2013, 2020):
            ks = kinds(orig.get(y, []))
            put(ws, r, ycol(y), symstr(ks, SYM_OLD) if ks else None, G(size=8), AL_C, B_HAIR)
        for y in years_vis:
            ks = kinds(plan.get(y, []), new=True, e=e)
            cell = put(ws, r, col_of[y], symstr(ks, SYM_NEW) if ks else None, F(size=9), AL_C, B_HAIR)
            if 2022 <= y <= 2025:
                cell.fill = fill("FFF5EE")
            if "U" in ks:
                cell.font = F(size=9, bold=True, color="C00000")
            costs.setdefault((u, s), {})
            costs[(u, s)][y] = costs[(u, s)].get(y, 0) + cost_of(plan.get(y, []))
        prev = [(y, a) for y, a in act.items() if any(kk in ("O", "R") for kk, _, _ in a)]
        ptxt = ""
        if prev:
            y, a = prev[-1]
            kk = [x for x in a if x[0] in ("O", "R")][0]
            ptxt = f"{fy_wa(y)} " + (kk[2] or {"O": "分解整備", "R": "更新"}[kk[0]])
        put(ws, r, cPrev, ptxt or None, F(size=7), AL_L, B_HAIR)
        b = "基幹改良（R4〜R7）対象" if W.upgrade_year(e, "2020") else ""
        b = {"2-RH-515": "施栓 累計 3 本。R8 更新（延命化）", "C-ST-501": "定期事業者検査 別契約（R4・R8…）",
             "C-TR-1002": "R元 油中ガス アセチレン微量 → 経過観察", "3-IDF-607": "R2 軸受交換。R6 改良で更新"}.get(e["tag"], b)
        put(ws, r, cB, b or None, F(size=7), AL_L, B_HAIR)
        ws.row_dimensions[r].outline_level = 2
        r += 1
    r += 1
    merge(ws, r, 1, r, cB, "■ 法定点検・検査", F(size=9, bold=True), AL_L, B_ALL, fill("BDD7EE"))
    r += 1
    for (nm, law, cyc, ys) in legal_rows("2020"):
        merge(ws, r, 1, r, 2, cyc, F(size=8), AL_C, B_HAIR)
        merge(ws, r, 3, r, 5, nm, F(size=8), AL_L, B_HAIR)
        merge(ws, r, 6, r, 8, "", None, None, B_HAIR)
        for y in years_vis:
            put(ws, r, col_of[y], "法" if ys.get(y) else None, F(size=8), AL_C, B_HAIR)
        put(ws, r, cPrev, None, None, None, B_HAIR)
        put(ws, r, cB, law, F(size=7), AL_L, B_HAIR)
        r += 1
    ws.sheet_properties.outlinePr.summaryBelow = False
    ws.freeze_panes = f"{CL(col_of[2020])}8"
    page(ws, landscape=True, paper=8, fit_w=1, fit_h=0, area=f"A1:{CL(cB)}{r - 1}", titles="5:7",
         footer="&P / &N", header="【架空】みなと臨海工場 延命化計画 R2")

    # 費用集計（百万円）
    w2 = wb.create_sheet("費用集計")
    put(w2, 1, 1, "【架空】みなと臨海工場　延命化計画　年度別整備費（概算）", F(size=13, bold=True))
    put(w2, 2, 1, "単位：百万円（税抜）　※ 各号炉シート・共通シートの機器別整備費の積上げ。基幹的設備改良工事は別行（R2.10 概算）", F(size=9))
    put(w2, 4, 1, "号炉", F(size=9, bold=True), AL_C, B_ALL, fill("DDEBF7"))
    put(w2, 4, 2, "設備区分", F(size=9, bold=True), AL_C, B_ALL, fill("DDEBF7"))
    w2.column_dimensions["A"].width = 7
    w2.column_dimensions["B"].width = 16
    for i, y in enumerate(years_vis):
        put(w2, 4, 3 + i, f"{y}\n{fy_wa(y)}", F(size=8, bold=True), AL_C, B_ALL, fill("DDEBF7"))
        w2.column_dimensions[CL(3 + i)].width = 6.5
    cT = 3 + len(years_vis)
    put(w2, 4, cT, "計", F(size=9, bold=True), AL_C, B_ALL, fill("DDEBF7"))
    w2.column_dimensions[CL(cT)].width = 8
    w2.row_dimensions[4].height = 26
    rr = 5
    tot = {y: 0 for y in years_vis}
    for (u, s), d in sorted(costs.items(), key=lambda x: (UNIT_ORDER[x[0][0]], SETSUBI_ORDER.index(x[0][1]))):
        put(w2, rr, 1, UNAME[u], F(size=8), AL_C, B_HAIR)
        put(w2, rr, 2, s, F(size=8), AL_L, B_HAIR)
        rs = 0
        for i, y in enumerate(years_vis):
            v = round(d.get(y, 0) / 1000, 1)
            tot[y] += v
            rs += v
            put(w2, rr, 3 + i, v if v else None, F(size=8), AL_R, B_HAIR, nf="0.0")
        put(w2, rr, cT, round(rs, 1), F(size=8), AL_R, B_HAIR, nf="#,##0.0")
        rr += 1
    put(w2, rr, 1, "", None, None, B_ALL, fill("F2F2F2"))
    put(w2, rr, 2, "機器別 計", F(size=8, bold=True), AL_L, B_ALL, fill("F2F2F2"))
    for i, y in enumerate(years_vis):
        put(w2, rr, 3 + i, round(tot[y], 1), F(size=8, bold=True), AL_R, B_ALL, fill("F2F2F2"), nf="#,##0.0")
    put(w2, rr, cT, round(sum(tot.values()), 1), F(size=8, bold=True), AL_R, B_ALL, fill("F2F2F2"), nf="#,##0.0")
    rr += 1
    kikan = {2022: 1820.0, 2023: 2140.0, 2024: 2210.0, 2025: 1990.0}
    put(w2, rr, 1, "", None, None, B_ALL)
    put(w2, rr, 2, "基幹的設備改良工事", F(size=8, bold=True), AL_L, B_ALL, fill("FCE4D6"))
    for i, y in enumerate(years_vis):
        put(w2, rr, 3 + i, kikan.get(y), F(size=8), AL_R, B_ALL, fill("FCE4D6") if y in kikan else None, nf="#,##0.0")
    put(w2, rr, cT, sum(kikan.values()), F(size=8), AL_R, B_ALL, nf="#,##0.0")
    rr += 1
    put(w2, rr, 2, "合　計", F(size=9, bold=True), AL_L, B_ALL, fill("BDD7EE"))
    for i, y in enumerate(years_vis):
        put(w2, rr, 3 + i, round(tot[y] + kikan.get(y, 0), 1), F(size=8, bold=True), AL_R, B_ALL, fill("BDD7EE"), nf="#,##0.0")
    put(w2, rr, cT, round(sum(tot.values()) + sum(kikan.values()), 1), F(size=8, bold=True), AL_R, B_ALL, fill("BDD7EE"), nf="#,##0.0")
    put(w2, rr + 2, 1, "※ 基幹改良の内訳：1号炉 R4、2号炉 R5、3号炉 R6、共通（DCS 共通部・タービン制御）R7。各年 1 炉ずつ長期停止。", F(size=8))
    put(w2, rr + 3, 1, "※ 定期事業者検査（タービン）R4・R8 は共通 余熱利用設備に含む。", F(size=8))
    page(w2, landscape=True, paper=8, fit_w=1, fit_h=1, area=f"A1:{CL(cT)}{rr + 3}")
    props(wb, "【架空】みなと臨海工場 延命化計画（令和2年度版）", "h3", "中長期整備計画", dt.datetime(2013, 12, 2, 10, 0),
          modified_by="h4")
    wb.save(path)
    return n



def main():
    eq = EQ.build()
    p1 = os.path.join(OUT, "【架空】みなと臨海工場_中長期整備計画_平成31年度版_施設保全計画様式.xlsx")
    p2 = os.path.join(OUT, "【架空】みなと臨海工場_中長期整備計画_令和2年度版_延命化計画様式.xlsx")
    print("plan rows", build_old(eq, p1), build_new(eq, p2))
    return [p1, p2]


if __name__ == "__main__":
    main()
