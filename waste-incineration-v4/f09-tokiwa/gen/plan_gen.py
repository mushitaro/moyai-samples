# -*- coding: utf-8 -*-
"""中長期整備計画（SPC 長期修繕計画 sv1 2016 版／sv2 2017 版、組合 中長期整備計画 令和6年度版）"""
import os, datetime as dt
from openpyxl import Workbook
from openpyxl.comments import Comment
import world as W
from xlh import *

GROUP_ORDER = ["受入供給設備", "燃焼溶融設備", "燃焼ガス冷却設備", "排ガス処理設備", "通風設備", "灰出し・スラグ設備",
               "余熱利用設備", "給水設備", "排水処理設備", "電気設備", "計装設備"]

def sort_key_block(e):
    return (e.furn if e.furn else 9, GROUP_ORDER.index(e.group), e.sys, e.idx)

def rk(key):
    return [r for r in W.REP if r["key"] == key][0]

# 代表機器：版ごとの計画（その版を作った時点で知っていたことだけ）
def annual(y0, y1, code="I"):
    return {y: code for y in range(y0, y1 + 1)}

def rep_marks(key, ver, y0, y1):
    m = {}
    if key in ("R1", "R2"):
        m = annual(y0, y1)
        if ver == "spc2016":
            yrs = {"R1": [2012, 2015, 2019, 2023], "R2": [2013, 2017, 2021]}[key]
            for y in yrs: m[y] = "O"
        elif ver == "spc2017":
            yrs = {"R1": [2019, 2023], "R2": [2017, 2021]}[key]
            for y in yrs: m[y] = "P"
            m[{"R1": 2026, "R2": 2028}[key]] = "R"
        elif ver == "u2024":
            m[{"R1": 2025, "R2": 2026}[key]] = "U"
    elif key in ("R3", "R4"):
        odd = [y for y in range(y0, y1 + 1) if y % 2 == 1]
        m = {y: "I" for y in odd}
        if ver == "spc2016":
            pass
        elif ver == "spc2017":
            m[2019 if key == "R3" else 2021] = "P"
            m[2027 if key == "R3" else 2029] = "R"
        elif ver == "u2024":
            m[2025 if key == "R3" else 2026] = "U"
    elif key == "R5":
        if ver == "spc2016":
            m = {2012: "O", 2015: "O", 2019: "O", 2023: "O"}
        elif ver == "spc2017":
            m = {2020: "O", 2024: "R"}
        else:
            m = {2025: "U"}
    elif key == "R6":
        if ver == "spc2016":
            m = {y: "O" for y in range(2009, 2024) if y % 2 == 1 and y > 2009}
            m[2021] = "R"
        elif ver == "spc2017":
            m = {2017: "O", 2019: "R", 2021: "O", 2023: "O", 2025: "O", 2027: "O", 2029: "O", 2031: "R", 2033: "O"}
        else:
            m = {2024: "", 2026: "O", 2028: "O", 2030: "O", 2031: "R", 2033: "O"}
            m.pop(2024)
            m[2025] = "I"
    elif key == "R7":
        if ver == "spc2016":
            m = {y: "O" for y in range(2010, 2024, 2)}
            m[2025 if False else 2022] = "O"
        elif ver == "spc2017":
            m = {2018: "O", 2020: "O", 2022: "O", 2024: "O", 2026: "R"}
        else:
            m = {2025: "O", 2028: "O", 2030: "R"}
    elif key == "R8":
        if ver == "spc2016":
            m = {2014: "O", 2018: "O", 2022: "O"}
            for y in range(y0, y1 + 1): m.setdefault(y, "I")
        elif ver == "spc2017":
            m = {2018: "O", 2022: "O", 2026: "O", 2030: "O"}
            for y in range(y0, y1 + 1): m.setdefault(y, "I")
        else:
            m = {2025: "O", 2031: "O"}
            for y in range(y0, y1 + 1): m.setdefault(y, "I")
    elif key == "R9":
        if ver == "spc2016":
            m = {2013: "O", 2016: "O", 2019: "O", 2022: "O"}
        elif ver == "spc2017":
            m = {2019: "O", 2022: "O", 2025: "O", 2028: "O", 2031: "R"}
        else:
            m = {2026: "O", 2029: "O", 2030: "R"}
    elif key == "R10":
        m = {y: "O" for y in (2012, 2016, 2020, 2024, 2028, 2032)}
    elif key == "R11":
        if ver == "spc2016":
            m = {2019: "R"}
        elif ver == "spc2017":
            m = {2019: "R", 2027: "R"}
        else:
            m = {2028: "R"}
    elif key == "R12":
        if ver == "spc2016":
            m = {2012: "O", 2016: "O", 2020: "O"}
        elif ver == "spc2017":
            m = {2020: "O", 2024: "O", 2028: "O"}
        else:
            m = {2026: "O", 2030: "O"}
    return {y: c for y, c in m.items() if y0 <= y <= y1 and c}

REP_BY_TAG = {r["tag"]: r["key"] for r in W.REP}
KIKAN = {e.tag: (code, fy, cost, txt) for e, code, fy, cost, txt in W.kikan_items()}

def marks_for(e, ver, y0, y1):
    if e.tag in REP_BY_TAG:
        return rep_marks(REP_BY_TAG[e.tag], ver, y0, y1)
    m = W.policy_marks(e, y0, y1)
    if ver == "u2024" and e.tag in KIKAN:
        code, fy, cost, txt = KIKAN[e.tag]
        m = {y: c for y, c in m.items() if not ((fy - 3 <= y <= fy + 14 and c == "R") or (fy - 1 <= y <= fy + 1 and c == "O"))}
        m[fy] = code
    if ver == "spc2016":
        # sv1 の時代は C 重要度の小物は印を付けていない（担当の割り切り）
        if e.crit == "C" and e.cls in ("conv", "mvalve", "fan_s", "tank", "lp"):
            m = {y: c for y, c in m.items() if c == "R"}
    return m

def legal_rows(ver, y0, y1):
    rows = []
    def yrs(f):
        return {y: "L" for y in range(y0, y1 + 1) if f(y)}
    rows.append(("ボイラー定期事業者検査（1号）", "電気事業法", "2年", yrs(lambda y: y % 2 == 1)))
    rows.append(("ボイラー定期事業者検査（2号）", "電気事業法", "2年", yrs(lambda y: y % 2 == 1)))
    rows.append(("蒸気タービン定期事業者検査", "電気事業法", "4年", yrs(lambda y: y % 4 == 0)))
    rows.append(("受変電設備 年次点検（月次点検は毎月）", "保安規程", "1年", yrs(lambda y: True)))
    rows.append(("非常用発電機 点検（年2回）", "消防法・保安規程", "6月", yrs(lambda y: True)))
    rows.append(("ごみクレーン 性能検査（1号・2号）", "クレーン則", "2年", yrs(lambda y: y % 2 == 1)))
    rows.append(("ごみクレーン 年次自主検査", "クレーン則", "1年", yrs(lambda y: y % 2 == 0)))
    rows.append(("タービン建屋天井クレーン 性能検査", "クレーン則", "2年", yrs(lambda y: y % 2 == 0)))
    rows.append(("ばい煙測定（年2回）・DXN排ガス測定（年1回）", "大防法・DXN法", "6月", yrs(lambda y: True)))
    rows.append(("ダイオキシン類 作業環境測定", "安衛則", "6月", yrs(lambda y: True)))
    if ver != "u2024":
        rows.append(("計量機 定期検査", "計量法", "2年", yrs(lambda y: y % 2 == 0)))   # 組合版はこの行が抜けている（わざと）
    rows.append(("消防用設備等点検", "消防法", "6月", yrs(lambda y: True)))
    if ver == "spc2016":
        pz = [2011, 2014, 2018, 2021]       # 2014→2018 の 4 年（周期の超え。わざと）
    elif ver == "spc2017":
        pz = [2018, 2021, 2024, 2027, 2030, 2033]
    else:
        pz = [2024, 2027, 2030, 2033]
    rows.append(("精密機能検査", "廃掃法施行規則", "3年", {y: "L" for y in pz if y0 <= y <= y1}))
    if ver != "spc2016":
        rows.append(("第一種圧力容器 性能検査（脱気器）", "ボイラー則", "1年", yrs(lambda y: True)))
    return rows

# ============================================================== SPC sv1 2016 年度版
def cyc_of(e, ver):
    """保全システムの「周期（年）」。代表機器は版ごとの判断"""
    k = REP_BY_TAG.get(e.tag)
    if k:
        return {"R1": 4 if ver == "spc2016" else 3, "R2": 4, "R3": 2, "R4": 2, "R5": 4, "R6": 2, "R7": 2 if ver == "spc2016" else 3,
                "R8": 4, "R9": 3, "R10": 4, "R11": None, "R12": 4}[k]
    ins, ovh, life = W.POLICY.get(e.cls, W.DEFAULT_POLICY)
    return ovh or ins or None

def nendo_row(ws, r, yc0, Y0, Y1, lastcol):
    """事業年度の n 年目（第1期 2009〜2023、第2期 2024〜）"""
    merge(ws, r, 1, r, lastcol, "事業年度（n 年目）", f(8, True), al("right"), B_ALL, fill("F2F2F2"))
    for y in range(Y0, Y1 + 1):
        n = y - 2008 if y <= 2023 else y - 2023
        put(ws, r, yc0 + y - Y0, n, f(7, color="C00000" if y == 2024 else "000000"), al("center"), B_ALL, fill("F2F2F2"))
    put(ws, r, yc0 + Y1 - Y0 + 1, "2024 から第2期 1 年目" if Y1 > 2023 else "第1期 15 年", f(7), al("left"), B_ALL, fill("F2F2F2"))

def spc_sv1(path):
    wb = Workbook(); ws = wb.active; ws.title = "長期修繕計画"
    Y0, Y1 = 2009, 2023
    yc0 = 8
    ncol = yc0 + (Y1 - Y0 + 1)   # 備考列
    SYM = {"I": "○", "O": "◎", "P": "◎", "R": "●", "U": "●", "L": "○"}
    merge(ws, 1, 1, 1, ncol, f"{W.FAC}　長期修繕計画表（2016年度改訂）", f(16, True), al("center"), None)
    merge(ws, 2, 1, 2, ncol, f"{W.SPC}　保全グループ　作成：sv1　改訂日 2016/4/15　（第1期 長期包括運営委託 2009〜2023年度）", f(9), al("right"), None)
    put(ws, 3, 1, "凡例", f(9, True))
    merge(ws, 3, 2, 3, 6, "○：点検　◎：整備（分解・補修）　●：更新", f(9), al("left"), None)
    # 見出し
    hdr = ["No.", "設備", "系統", "機器名称", "タグNo.", "区分", "周期\n(年)"]
    for i, h in enumerate(hdr):
        merge(ws, 5, i + 1, 6, i + 1, h, f(9, True), al("center", wrap=True), B_ALL, fill("DDEBF7"))
    merge(ws, 5, yc0, 5, ncol - 1, "年　　度（西暦）", f(9, True), al("center"), B_ALL, fill("DDEBF7"))
    for y in range(Y0, Y1 + 1):
        put(ws, 6, yc0 + y - Y0, f"'{str(y)[2:]}", f(9, True), al("center"), B_ALL, fill("DDEBF7"))
    merge(ws, 5, ncol, 6, ncol, "備　考", f(9, True), al("center"), B_ALL, fill("DDEBF7"))
    nendo_row(ws, 7, yc0, Y0, Y1, 7)
    widths(ws, {"A": 5, "B": 11, "C": 11, "D": 26, "E": 10, "F": 5, "G": 5})
    for c in range(yc0, ncol): ws.column_dimensions[L(c)].width = 3.6
    ws.column_dimensions[L(ncol)].width = 30
    r = 8; no = 0
    costsum = {}
    eqs = sorted(W.EQS, key=sort_key_block)
    blocks = [(1, "■ 1号炉"), (2, "■ 2号炉"), (0, "■ 共通設備")]
    for furn, label in blocks:
        merge(ws, r, 1, r, ncol, label, f(10, True), al("left"), B_ALL, fill("FFF2CC")); r += 1
        rows = [e for e in eqs if e.furn == furn]
        gstart = r; sstart = r; prevg = prevs = None
        for e in rows:
            no += 1
            put(ws, r, 1, no, f(8), al("center"), B_HAIR)
            if e.group != prevg:
                if prevg is not None and furn != 2 and gstart < r - 1:
                    ws.merge_cells(start_row=gstart, start_column=2, end_row=r - 1, end_column=2)
                gstart = r
                put(ws, r, 2, e.group, f(8), al("center", "center", True), B_HAIR)
            else:
                put(ws, r, 2, "〃" if furn == 2 else None, f(8), al("center"), B_HAIR)
            if e.sub != prevs or e.group != prevg:
                if prevs is not None and furn == 1 and sstart < r - 1:
                    ws.merge_cells(start_row=sstart, start_column=3, end_row=r - 1, end_column=3)
                sstart = r
                put(ws, r, 3, e.sub, f(8), al("left", "center", True), B_HAIR)
            else:
                put(ws, r, 3, "〃" if furn != 1 else None, f(8), al("center" if furn != 1 else "left"), B_HAIR)
            prevg, prevs = e.group, e.sub
            nm = e.name
            put(ws, r, 4, nm, f(8), al("left"), B_HAIR)
            put(ws, r, 5, e.tag, f(8), al("center"), B_HAIR)
            put(ws, r, 6, e.duty if e.duty != "交互" else "", f(8), al("center"), B_HAIR)
            put(ws, r, 7, cyc_of(e, "spc2016"), f(8), al("center"), B_HAIR)
            m = marks_for(e, "spc2016", Y0, Y1)
            for y in range(Y0, Y1 + 1):
                c = put(ws, r, yc0 + y - Y0, SYM.get(m.get(y), None), f(9), al("center"), B_HAIR)
                if y < 2016:
                    c.fill = fill("F2F2F2")
                if m.get(y) in ("O", "R", "P", "U"):
                    costsum.setdefault(e.group, {}).setdefault(y, 0)
                    costsum[e.group][y] += W.cost(e, m[y])
            note = ""
            if e.tag in REP_BY_TAG:
                note = {"R1": "出滓口・側壁の減りに注意", "R3": "H27 減肉1本 肉盛済（要監視）", "R6": "2年毎 A/B交互",
                        "R8": "ろ布4年毎", "R11": "DCS更新（メーカー提案 2019〜）", "R10": "開放 4年毎"}.get(REP_BY_TAG[e.tag], "")
            put(ws, r, ncol, note, f(8), al("left", shrink=True), B_HAIR)
            r += 1
        if furn != 2 and gstart < r - 1:
            ws.merge_cells(start_row=gstart, start_column=2, end_row=r - 1, end_column=2)
        if furn == 1 and sstart < r - 1:
            ws.merge_cells(start_row=sstart, start_column=3, end_row=r - 1, end_column=3)
    # 法定点検
    merge(ws, r, 1, r, ncol, "■ 法定点検・測定", f(10, True), al("left"), B_ALL, fill("FFF2CC")); r += 1
    for name, basis, cyc, m in legal_rows("spc2016", Y0, Y1):
        merge(ws, r, 2, r, 4, name, f(8), al("left"), B_HAIR)
        put(ws, r, 5, basis, f(7), al("center", shrink=True), B_HAIR)
        put(ws, r, 6, None, None, None, B_HAIR)
        put(ws, r, 7, cyc, f(8), al("center"), B_HAIR)
        put(ws, r, 1, None, None, None, B_HAIR)
        for y in range(Y0, Y1 + 1):
            c = put(ws, r, yc0 + y - Y0, SYM.get(m.get(y)), f(9), al("center"), B_HAIR)
            if y < 2016: c.fill = fill("F2F2F2")
        put(ws, r, ncol, "", f(8), None, B_HAIR)
        r += 1
    # 費用
    r += 1
    merge(ws, r, 1, r, ncol, "■ 年度別概算費用（百万円・税抜、点検費は定期整備費に含む）", f(10, True), al("left"), B_ALL, fill("FFF2CC")); r += 1
    tot = {}
    for g in GROUP_ORDER:
        if g not in costsum: continue
        merge(ws, r, 2, r, 7, g, f(8), al("left"), B_HAIR)
        for y in range(Y0, Y1 + 1):
            v = round(costsum[g].get(y, 0) / 1000, 1)
            tot[y] = tot.get(y, 0) + v
            put(ws, r, yc0 + y - Y0, v if v else None, f(7), al("right", shrink=True), B_HAIR, fmt="0.0")
        r += 1
    merge(ws, r, 2, r, 7, "合　計", f(9, True), al("center"), B_ALL, fill("EDEDED"))
    for y in range(Y0, Y1 + 1):
        put(ws, r, yc0 + y - Y0, round(tot.get(y, 0), 1), f(7, True), al("right", shrink=True), B_ALL, fill("EDEDED"), fmt="0.0")
    last = r
    outline(ws, 5, 1, last, ncol)
    ws.freeze_panes = "H8"
    page(ws, "landscape", 8, 1, 0, f"A1:{L(ncol)}{last}", "5:7")
    ws.oddFooter.center.text = "&P / &N"
    ws.oddFooter.right.text = "ときわ北部エコサービス 保全G"
    props(wb, f"{W.FAC} 長期修繕計画表（2016年度改訂）", "sv1（【架空】ときわ北部エコサービス 保全グループ）", dt.datetime(2016, 4, 15, 10, 0))
    wb.save(path)
    return last

# ============================================================== SPC sv2 2017 年度版（sv1 版をコピーして直した）
def spc_sv2(path):
    wb = Workbook(); ws = wb.active; ws.title = "長期修繕計画"
    Y0, Y1 = 2009, 2033
    # 列：A No / B 設備 / C 系統 / D 号炉(挿入) / E 機器名称 / F タグ / G 区分 / H 重要度(追加) / I 前回整備(追加) / J.. 年度 / 備考
    yc0 = 11
    ncol = yc0 + (Y1 - Y0 + 1)
    SYM = {"I": "○", "O": "◎", "P": "△", "R": "●", "U": "●", "L": "◇"}
    merge(ws, 1, 1, 1, ncol, f"{W.FAC}　長期修繕計画表（2017年度改訂）", f(16, True), al("center"), None)
    merge(ws, 2, 1, 2, ncol, f"{W.SPC}　保全グループ　作成：sv1　改訂：sv2 2017/5/22　（第1期 2009〜2023年度／2024年度以降は第2期の参考）", f(9), al("right"), None)
    # 凡例は旧版の位置のまま（列を 1 つ挿入したので結合が 1 列左にずれて見える）
    put(ws, 3, 1, "凡例", f(9, True))
    merge(ws, 3, 2, 3, 6, "○：点検　◎：整備（分解・補修）　△：部分補修　●：更新　◇：法定", f(9), al("left"), None)
    hdr = ["No.", "設備", "系統", "号炉", "機器名称", "タグNo.", "区分", "周期\n(年)", "重要度", "前回\n整備"]
    for i, h in enumerate(hdr):
        merge(ws, 5, i + 1, 6, i + 1, h, f(9, True), al("center", wrap=True), B_ALL, fill("DDEBF7" if i not in (3, 8, 9) else "E2EFDA"))
    merge(ws, 5, yc0, 5, yc0 + 2023 - Y0, "第1期（西暦年度）", f(9, True), al("center"), B_ALL, fill("DDEBF7"))
    merge(ws, 5, yc0 + 2024 - Y0, 5, ncol - 1, "第2期（参考）", f(9, True), al("center"), B_ALL, fill("FCE4D6"))
    for y in range(Y0, Y1 + 1):
        put(ws, 6, yc0 + y - Y0, f"'{str(y)[2:]}", f(9, True), al("center"), B_ALL, fill("DDEBF7" if y <= 2023 else "FCE4D6"))
    merge(ws, 5, ncol, 6, ncol, "備　考", f(9, True), al("center"), B_ALL, fill("DDEBF7"))
    nendo_row(ws, 7, yc0, Y0, Y1, 10)
    widths(ws, {"A": 5, "B": 11, "C": 11, "D": 5, "E": 26, "F": 10, "G": 5, "H": 5, "I": 5, "J": 6})
    for c in range(yc0, ncol): ws.column_dimensions[L(c)].width = 3.6
    ws.column_dimensions[L(ncol)].width = 30
    # 過去年度（2009〜2016）は非表示にして残した
    for y in range(Y0, 2017):
        ws.column_dimensions[L(yc0 + y - Y0)].hidden = True
    r = 8; no = 0
    # 旧版の「■ 1号炉」見出し行が先頭に残っている（空の余分な行）
    merge(ws, r, 1, r, ncol, "■ 1号炉", f(10, True), al("left"), B_ALL, fill("FFF2CC")); r += 1
    put(ws, r, 1, None); r += 1  # 空行（旧版の行を消した跡）
    eqs = sorted(W.EQS, key=lambda e: (GROUP_ORDER.index(e.group), e.sys, e.base, e.suffix, e.furn if e.furn else 9))
    costsum = {}
    prevg = None
    gstart = r
    for e in eqs:
        if e.group != prevg:
            if prevg is not None and gstart < r - 1:
                ws.merge_cells(start_row=gstart, start_column=2, end_row=r - 1, end_column=2)
            gstart = r
            put(ws, r, 2, e.group, f(8), al("center", "center", True), B_HAIR)
        else:
            put(ws, r, 2, None, None, None, B_HAIR)
        prevg = e.group
        no += 1
        put(ws, r, 1, no, f(8), al("center"), B_HAIR)
        put(ws, r, 3, e.sub, f(8), al("left", shrink=True), B_HAIR)
        put(ws, r, 4, {0: "共", 1: "1", 2: "2"}[e.furn], f(8), al("center"), B_HAIR)
        put(ws, r, 5, e.name, f(8), al("left"), B_HAIR)
        put(ws, r, 6, e.tag, f(8), al("center"), B_HAIR)
        put(ws, r, 7, e.duty if e.duty != "交互" else "", f(8), al("center"), B_HAIR)
        put(ws, r, 8, cyc_of(e, "spc2017"), f(8), al("center"), B_HAIR)
        put(ws, r, 9, e.crit, f(8), al("center"), B_HAIR)
        mall = marks_for(e, "spc2017", 2017, Y1)
        old = marks_for(e, "spc2016", Y0, 2016)
        past = [y for y, c in old.items() if c in ("O", "R")]
        put(ws, r, 10, f"'{str(max(past))[2:]}" if past else None, f(8), al("center"), B_HAIR)
        for y in range(Y0, Y1 + 1):
            code = old.get(y) if y < 2017 else mall.get(y)
            c = put(ws, r, yc0 + y - Y0, SYM.get(code), f(9), al("center"), B_HAIR)
            if code in ("O", "R", "P", "U") and y >= 2017:
                costsum.setdefault(e.group, {}).setdefault(y, 0)
                costsum[e.group][y] += W.cost(e, code)
        note = ""
        if e.tag in REP_BY_TAG:
            note = {"R1": "H27・H31 部分補修。全面張替は第2期で要検討", "R2": "2号炉の2年遅れ", "R3": "H27 減肉→肉盛。材質変更検討",
                    "R4": "2号炉に準じる", "R5": "羽根車摩耗 監視", "R6": "A/B 2年毎交互", "R7": "2号機（予備）併用で刃の持ち改善",
                    "R8": "ろ布4年毎", "R9": "ワイヤ3年毎", "R10": "定期事業者検査 開放", "R11": "DCS更新 2019〜2020（別途工事）", "R12": "チェーン4年毎"}[REP_BY_TAG[e.tag]]
        elif e.tag in KIKAN and any(v == "R" for v in mall.values()):
            note = "第2期 大規模改修の候補"
        c = put(ws, r, ncol, note, f(8), al("left", shrink=True), B_HAIR)
        if e.tag in REP_BY_TAG and REP_BY_TAG[e.tag] == "R3":
            c.comment = Comment("2015.10 定期事業者検査で二次過熱器下部ベンド減肉 4.0→2.6mm。肉盛補修済。次回 2017 で再測定。sv2", "sv2")
        r += 1
    ws.merge_cells(start_row=gstart, start_column=2, end_row=r - 1, end_column=2)
    merge(ws, r, 1, r, ncol, "■ 法定点検・測定", f(10, True), al("left"), B_ALL, fill("FFF2CC")); r += 1
    for name, basis, cyc, m in legal_rows("spc2017", 2017, Y1):
        old = dict((y, "L") for y in [2011, 2014] if "精密" in name)
        merge(ws, r, 2, r, 5, name, f(8), al("left"), B_HAIR)
        put(ws, r, 6, basis, f(7), al("center", shrink=True), B_HAIR)
        put(ws, r, 8, cyc, f(8), al("center"), B_HAIR)
        for c_ in (1, 7, 9, 10): put(ws, r, c_, None, None, None, B_HAIR)
        for y in range(Y0, Y1 + 1):
            put(ws, r, yc0 + y - Y0, SYM.get(m.get(y) if y >= 2017 else old.get(y)), f(9), al("center"), B_HAIR)
        put(ws, r, ncol, "2018 実施（周期 4 年となった）" if "精密" in name else ("2017 から計上" if "圧力容器" in name else ""), f(8), al("left", shrink=True), B_HAIR)
        r += 1
    last = r - 1
    outline(ws, 5, 1, last, ncol)
    ws.freeze_panes = "K8"
    page(ws, "landscape", 8, 1, 0, f"A1:{L(ncol)}{last}", "5:7")
    ws.oddFooter.center.text = "&P / &N"
    ws.oddFooter.right.text = "ときわ北部エコサービス 保全G"
    # 費用は別シートへ（sv2 で移した）
    ws2 = wb.create_sheet("費用集計")
    merge(ws2, 1, 1, 1, 20, f"{W.FAC}　長期修繕計画 年度別費用集計（千円・税抜）2017年度改訂", f(13, True), al("left"), None)
    put(ws2, 2, 1, "※点検費は年度の定期整備費に含むため計上しない。2024年度以降は第2期の参考値。", f(8))
    yrs = list(range(2017, 2034))
    put(ws2, 4, 1, "設備", f(9, True), al("center"), B_ALL, fill("DDEBF7"))
    for i, y in enumerate(yrs):
        put(ws2, 4, 2 + i, f"{y}年度", f(8, True), al("center"), B_ALL, fill("DDEBF7" if y <= 2023 else "FCE4D6"))
    rr = 5; tot = {}
    for g in GROUP_ORDER:
        if g not in costsum: continue
        put(ws2, rr, 1, g, f(9), al("left"), B_ALL)
        for i, y in enumerate(yrs):
            v = costsum[g].get(y, 0)
            tot[y] = tot.get(y, 0) + v
            put(ws2, rr, 2 + i, v if v else None, f(8), al("right"), B_ALL, fmt="#,##0")
        rr += 1
    put(ws2, rr, 1, "合計", f(9, True), al("center"), B_ALL, fill("EDEDED"))
    for i, y in enumerate(yrs):
        v = tot.get(y, 0)
        if y == 2019:
            v = v - W.COST_R["dcsc"] * 4   # DCS 更新分を後から行に足したが合計を打ち直していない（わざと）
        put(ws2, rr, 2 + i, v, f(8, True), al("right"), B_ALL, fill("EDEDED"), fmt="#,##0")
    put(ws2, rr + 2, 1, "2019年度は DCS 更新（別途工事）を含む", f(8))
    widths(ws2, {"A": 16}); [setattr(ws2.column_dimensions[L(c)], "width", 9.5) for c in range(2, 2 + len(yrs))]
    page(ws2, "landscape", 9, 1, 1, f"A1:{L(1 + len(yrs))}{rr + 2}")
    props(wb, f"{W.FAC} 長期修繕計画表（2017年度改訂）", "sv1（【架空】ときわ北部エコサービス 保全グループ）", dt.datetime(2017, 5, 22, 16, 0), "sv2")
    wb.save(path)
    return last

# ============================================================== 組合 中長期整備計画 令和6年度版
U_GROUPS = [("受入供給", ["受入供給設備"]), ("燃焼溶融", ["燃焼溶融設備"]), ("燃焼ガス冷却", ["燃焼ガス冷却設備"]),
            ("排ガス処理", ["排ガス処理設備"]), ("通風", ["通風設備"]), ("灰出し・スラグ", ["灰出し・スラグ設備"]),
            ("余熱利用", ["余熱利用設備"]), ("給水・排水", ["給水設備", "排水処理設備"]), ("電気・計装", ["電気設備", "計装設備"])]
U_ALIAS = {"流動空気押込送風機": "押込送風機（流動空気）", "誘引通風機": "誘引送風機", "バグフィルタ本体": "ろ過式集じん器",
           "旋回溶融炉本体": "溶融炉（旋回式）本体", "ガス化炉本体": "ガス化炉（流動床式）本体", "前処理破砕機": "破砕機（前処理）",
           "蒸気タービン": "蒸気タービン", "スラグ水砕コンベヤ": "水砕スラグコンベヤ"}

def u_name(e):
    b = U_ALIAS.get(e.base, e.base)
    if e.suffix:
        if e.suffix.endswith("号"): return f"{e.suffix}{b}"
        if e.suffix in ("A", "B"): return f"{b}（{e.suffix}）"
        if e.suffix.startswith("No.") or "-" in e.suffix: return f"{b} {e.suffix}"
        return f"{b}（{e.suffix}）"
    return b

def actual_u(e, m):
    """R6・R7 の実績（令和7年度改訂の時点で分かっていること）。多くは計画どおり、C 重要度の一部は繰越"""
    act = {}; carried = False
    for y in (2024, 2025):
        c = m.get(y)
        if c:
            act[y] = c
    if m.get(2024) == "O" and e.crit == "C" and e.idx % 4 == 0:
        act.pop(2024, None); act[2025] = "O"; carried = True
    return act, carried

def union_plan(path):
    wb = Workbook()
    Y0, Y1 = 2024, 2033
    SYM = {"I": "点", "O": "整", "P": "補", "R": "更", "U": "改", "L": "法"}
    summ = {}
    first = True
    for sheet, groups in U_GROUPS:
        ws = wb.active if first else wb.create_sheet()
        if first: first = False
        ws.title = sheet
        NY = Y1 - Y0 + 1
        nc = 6 + NY * 3 + 1
        merge(ws, 1, 1, 1, nc, f"{W.FAC}　中長期整備計画（令和6年度策定・令和7年度改訂）　【{sheet}設備】", f(14, True), al("left"), None)
        merge(ws, 2, 1, 2, nc, f"{W.UNION}　施設課　策定 u1 令和6年11月／改訂 u1 令和8年3月（R6・R7 実績記入、基幹改良の工期変更を反映）　計画期間 令和6〜15年度、千円・税込", f(9), al("right"), None)
        merge(ws, 3, 1, 3, nc, "記号　点：点検　整：整備　補：補修　更：更新　改：改良（基幹的設備改良工事）　費用は計画（概算）額。点検は運営委託料に含む。実績欄は年度末に記入", f(8), al("left"), None)
        heads = ["番号", "中分類", "機器名称", "号炉", "設置", "重要度"]
        for i, h in enumerate(heads):
            merge(ws, 5, i + 1, 6, i + 1, h, f(9, True), al("center", wrap=True), B_ALL, fill("E2EFDA"))
        for k, y in enumerate(range(Y0, Y1 + 1)):
            c = 7 + k * 3
            merge(ws, 5, c, 5, c + 2, wareki_fy(y).replace("年度", ""), f(9, True), al("center"), B_ALL, fill("E2EFDA" if y not in (2025, 2026, 2027) else "FFE699"))
            put(ws, 6, c, "計画", f(6, True), al("center"), B_ALL, fill("E2EFDA"))
            put(ws, 6, c + 1, "実績", f(6, True), al("center"), B_ALL, fill("FCE4D6" if y <= 2025 else "E2EFDA"))
            put(ws, 6, c + 2, "費用", f(6, True), al("center"), B_ALL, fill("E2EFDA"))
        merge(ws, 5, nc, 6, nc, "備考", f(9, True), al("center"), B_ALL, fill("E2EFDA"))
        widths(ws, {"A": 6, "B": 11, "C": 24, "D": 4.5, "E": 5, "F": 4.5})
        for k in range(NY):
            ws.column_dimensions[L(7 + k * 3)].width = 3.4
            ws.column_dimensions[L(8 + k * 3)].width = 3.4
            ws.column_dimensions[L(9 + k * 3)].width = 7.2
        ws.column_dimensions[L(nc)].width = 28
        r = 7
        eqs = [e for e in W.EQS if e.group in groups]
        eqs.sort(key=lambda e: (groups.index(e.group), e.sys, e.furn if e.furn else 9, e.idx))
        prev = None; mstart = r; seqn = 0
        sub_tot = {}
        for e in eqs:
            seqn += 1
            if e.sub != prev:
                if prev is not None and mstart < r - 1:
                    ws.merge_cells(start_row=mstart, start_column=2, end_row=r - 1, end_column=2)
                mstart = r
                put(ws, r, 2, e.sub, f(8), al("center", "center", True), B_HAIR)
            else:
                put(ws, r, 2, None, None, None, B_HAIR)
            prev = e.sub
            put(ws, r, 1, f"{sheet[:2]}-{seqn:03d}", f(7), al("center"), B_HAIR)
            put(ws, r, 3, u_name(e), f(8), al("left", shrink=True), B_HAIR)
            put(ws, r, 4, {0: "共通", 1: "1号", 2: "2号"}[e.furn], f(7), al("center"), B_HAIR)
            put(ws, r, 5, "H20" if e.install.startswith("2009") else wareki_fy(W.start_fy(e) - 1, True), f(7), al("center"), B_HAIR)
            put(ws, r, 6, {"A": "Ⅰ", "B": "Ⅱ", "C": "Ⅲ"}[e.crit], f(8), al("center"), B_HAIR)
            m = marks_for(e, "u2024", Y0, Y1)
            act, carried = actual_u(e, m)
            for k, y in enumerate(range(Y0, Y1 + 1)):
                c = 7 + k * 3
                code = m.get(y)
                put(ws, r, c, SYM.get(code), f(8, code == "U"), al("center"), B_HAIR,
                    fill("FFE699") if code == "U" or (code == "R" and e.tag in KIKAN and KIKAN[e.tag][1] == y) else None)
                a_ = act.get(y)
                put(ws, r, c + 1, SYM.get(a_) if a_ else None, f(8, color="C00000" if a_ and a_ != code else "000000"), al("center"), B_HAIR,
                    fill("F2F2F2") if y > 2025 else None)
                v = None
                if code in ("O", "P", "R", "U"):
                    if e.tag in KIKAN and KIKAN[e.tag][1] == y:
                        v = KIKAN[e.tag][2]
                    else:
                        v = round(W.cost(e, code) * 1.1)
                    if v:
                        sub_tot[y] = sub_tot.get(y, 0) + v
                put(ws, r, c + 2, v if v else None, f(7), al("right", shrink=True), B_HAIR, fmt="#,##0")
            note = ""
            if e.tag in KIKAN:
                note = "基幹改良 " + KIKAN[e.tag][3]
                if KIKAN[e.tag][1] == 2025:
                    note += "。R8.1.16 完了（工期延長）"
            elif e.tag in REP_BY_TAG:
                note = {"R6": "R7 は基幹改良と重なるため点検のみ", "R7": "刃交換は運営委託の整備費", "R8": "ろ布 R7 交換済",
                        "R9": "性能検査 奇数年度", "R10": "定期事業者検査 R6 実施", "R11": "OS サポート R10 まで", "R12": "R8 1号停止中に交換",
                        "R2": "R7 定修 残存厚 118mm"}.get(REP_BY_TAG[e.tag], "")
            if carried:
                note = (note + " " if note else "") + "R6 整備を R7 へ繰越"
            put(ws, r, nc, note, f(7), al("left", shrink=True), B_HAIR)
            r += 1
        if mstart < r - 1:
            ws.merge_cells(start_row=mstart, start_column=2, end_row=r - 1, end_column=2)
        merge(ws, r, 1, r, 6, "小　計（計画額）", f(9, True), al("center"), B_ALL, fill("EDEDED"))
        for k, y in enumerate(range(Y0, Y1 + 1)):
            c = 7 + k * 3
            put(ws, r, c, None, None, None, B_ALL, fill("EDEDED")); put(ws, r, c + 1, None, None, None, B_ALL, fill("EDEDED"))
            put(ws, r, c + 2, sub_tot.get(y, 0), f(7, True), al("right", shrink=True), B_ALL, fill("EDEDED"), fmt="#,##0")
        put(ws, r, nc, None, None, None, B_ALL, fill("EDEDED"))
        summ[sheet] = sub_tot
        outline(ws, 5, 1, r, nc)
        ws.freeze_panes = "G7"
        page(ws, "landscape", 8, 1, 0, f"A1:{L(nc)}{r}", "5:6")
        ws.oddHeader.right.text = "ときわ北部環境組合"
        ws.oddFooter.center.text = f"{sheet} &P/&N"
    # 法定点検シート
    ws = wb.create_sheet("法定点検")
    merge(ws, 1, 1, 1, 14, f"{W.FAC}　中長期整備計画（令和6年度策定・令和7年度改訂）　【法定点検・測定】", f(14, True), al("left"), None)
    for i, h in enumerate(["項目", "根拠", "周期"] + [wareki_fy(y).replace("年度", "") for y in range(Y0, Y1 + 1)]):
        put(ws, 3, 1 + i, h, f(9, True), al("center"), B_ALL, fill("E2EFDA"))
    r = 4
    for name, basis, cyc, m in legal_rows("u2024", Y0, Y1):
        put(ws, r, 1, name, f(8), al("left"), B_ALL); put(ws, r, 2, basis, f(8), al("center"), B_ALL); put(ws, r, 3, cyc, f(8), al("center"), B_ALL)
        for k, y in enumerate(range(Y0, Y1 + 1)):
            put(ws, r, 4 + k, "法" if m.get(y) else None, f(8), al("center"), B_ALL)
        r += 1
    widths(ws, {"A": 40, "B": 14, "C": 6}); [setattr(ws.column_dimensions[L(c)], "width", 7) for c in range(4, 14)]
    page(ws, "landscape", 9, 1, 1, f"A1:M{r - 1}")
    # 総括表（先頭へ）
    ws = wb.create_sheet("総括表", 0)
    merge(ws, 1, 1, 1, 12, f"{W.FAC}　中長期整備計画（令和6年度策定・令和7年度改訂）　総括表", f(15, True), al("left"), None)
    merge(ws, 2, 1, 2, 12, f"{W.UNION}　施設課　u1　（単位：千円・税込）", f(9), al("right"), None)
    put(ws, 4, 1, "設備区分", f(9, True), al("center"), B_ALL, fill("E2EFDA"))
    for k, y in enumerate(range(Y0, Y1 + 1)):
        put(ws, 4, 2 + k, wareki_fy(y), f(8, True), al("center", wrap=True), B_ALL, fill("E2EFDA" if y not in (2025, 2026, 2027) else "FFE699"))
    put(ws, 4, 12, "計", f(9, True), al("center"), B_ALL, fill("E2EFDA"))
    r = 5; colt = {}
    for sheet, _ in U_GROUPS:
        put(ws, r, 1, sheet + "設備", f(9), al("left"), B_ALL)
        rowt = 0
        for k, y in enumerate(range(Y0, Y1 + 1)):
            v = summ[sheet].get(y, 0)
            if sheet == "燃焼溶融" and y == 2026:
                v = round(v / 1.1)       # この 1 つだけ税抜のまま転記した（わざと。シートの小計と合わない）
            colt[y] = colt.get(y, 0) + v; rowt += v
            put(ws, r, 2 + k, v, f(8), al("right"), B_ALL, fmt="#,##0")
        put(ws, r, 12, rowt, f(8, True), al("right"), B_ALL, fmt="#,##0")
        r += 1
    put(ws, r, 1, "合　計", f(9, True), al("center"), B_ALL, fill("EDEDED"))
    for k, y in enumerate(range(Y0, Y1 + 1)):
        put(ws, r, 2 + k, colt[y], f(8, True), al("right"), B_ALL, fill("EDEDED"), fmt="#,##0")
    put(ws, r, 12, sum(colt.values()), f(8, True), al("right"), B_ALL, fill("EDEDED"), fmt="#,##0")
    r += 2
    notes = ["1　令和7〜9年度は基幹的設備改良工事（組合発注、継続費）を含む。2号炉 R7、1号炉 R8、共通 R9。R7 分は過熱器管材の納期遅れで R8.1.16 完了（変更契約 R7.12.5、年度内）。",
             "2　運営委託（第2期）の年度上限内で行う整備は、運営事業者の長期修繕計画と調整のうえ計上した。",
             "3　金額は令和6年度単価による概算（税込）。物価変動は見込んでいない。"]
    for t in notes:
        merge(ws, r, 1, r, 12, t, f(8), al("left", wrap=True), None); ws.row_dimensions[r].height = 24; r += 1
    widths(ws, {"A": 18, "L": 12}); [setattr(ws.column_dimensions[L(c)], "width", 11) for c in range(2, 12)]
    page(ws, "landscape", 9, 1, 1, f"A1:L{r - 1}")
    props(wb, f"{W.FAC} 中長期整備計画（令和6年度策定・令和7年度改訂）", "u1（【架空】ときわ北部環境組合 施設課）", dt.datetime(2024, 11, 28, 14, 0))
    wb.properties.modified = dt.datetime(2026, 3, 24, 16, 0)
    wb.save(path)
