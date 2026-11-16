# -*- coding: utf-8 -*-
"""機器台帳：旧（h3 平成25年度再編様式、令和6年度版）→ 新（h5 令和7年度 基幹改良反映様式）"""
import datetime as dt
import os

from openpyxl import Workbook
from openpyxl.utils import get_column_letter as CL
from openpyxl.styles import Alignment, Border, Font

import equipment as EQ
import world as W
from xlhelp import *

SETSUBI_ORDER = ["受入供給設備", "燃焼設備", "燃焼ガス冷却設備", "排ガス処理設備", "通風設備", "余熱利用設備", "灰出し設備",
                 "給水設備", "排水処理設備", "電気設備", "計装制御設備", "電気・計装設備", "その他設備"]
UNIT_ORDER = {"1": 0, "2": 1, "3": 2, "C": 3}
UNIT_LABEL_OLD = {"1": "1号", "2": "2号", "3": "3号", "C": "共通"}
UNIT_LABEL_NEW = {"1": "1号炉", "2": "2号炉", "3": "3号炉", "C": "共通"}

HOZEN = {"hv": "BM", "static": "BM", "panel_l": "BM", "mech_small": "BM", "tank": "TBM", "silo": "TBM",
         "boiler": "CBM", "sh": "CBM", "sh2": "CBM", "sh3": "CBM", "eco": "CBM", "rh": "CBM", "idf": "CBM",
         "furnace": "CBM", "grate": "CBM", "grate_wc": "CBM", "scrubber": "CBM", "scr": "CBM", "fan_big": "CBM",
         "tr_main": "CBM", "pump_corr": "CBM"}


def hozen(e):
    if not W._params(e):
        return "BM"
    h = HOZEN.get(e["cls"], "TBM")
    return "TBM" if h == "BM" else h


def cycle_txt(e):
    s = W.STRATEGY.get(e["cls"], [])
    prm = W._params(e)
    parts = []
    for p in prm:
        k = {"I": "点検", "O": "整備", "R": "更新"}[p["kind"]]
        parts.append(f"{k}{p['iv']}")
    return "・".join(parts)


def sort_key(e):
    return (SETSUBI_ORDER.index(e["setsubi"]), UNIT_ORDER[e["unit"]], e["sysno"], e["loop"], e["tag"])


def inst_ym(e):
    if e["installed"] == 2004:
        return "H16.3"
    if e["tag"] in ("C-CV-956", "C-HP-957"):
        return "H24.2"
    if e["tag"] == "C-LOG-1118":
        return "H27.3"
    y = e["installed"]
    return {2022: "R4.10", 2023: "R5.12", 2024: "R7.2", 2025: "R8.3"}[y]


import re as _re

SYSNAME_F = {1: "給じん系統図", 2: "燃焼設備系統図", 3: "ボイラ系統図", 4: "減温・集じん系統図", 5: "洗煙・再加熱系統図",
             6: "脱硝・通風系統図", 7: "灰出し系統図", 8: "炉別電気単線結線図・計装ループ図"}
SYSNAME_C = {0: "計量設備系統図", 1: "プラットホーム系統図", 2: "ごみピット・クレーン系統図", 3: "薬品供給系統図",
             4: "給水・主蒸気系統図", 5: "タービン発電設備系統図", 6: "復水・補機冷却水系統図", 7: "用水・圧縮空気系統図",
             8: "洗煙排水処理系統図", 9: "灰処理系統図", 10: "受変電単線結線図", 11: "中央制御システム構成図", 12: "煙突・保守用クレーン"}


def dwg(e):
    """完成図書の系統図番号（P&ID 図番）"""
    if e["unit"] == "C":
        return f"P-C{e['sysno']:02d}", SYSNAME_C[e["sysno"]]
    return f"P-{e['unit']}{e['sysno']:02d}", f"{e['unit']}号炉 " + SYSNAME_F[e["sysno"]]


def motor(e):
    """回転機の付属電動機（台数に数えない付属行）"""
    m = _re.search(r"(\d+(?:\.\d+)?)kW", e["spec"])
    if not m or e["type"] in ("INV", "BAT", "EG", "GEN", "ST"):
        return None
    kw = float(m.group(1))
    v = "6.6kV" if kw >= 200 else "440V"
    pole = "4P" if kw < 100 else ("6P" if e["type"] in ("IDF", "FDF") else "4P")
    enc = "全閉外扇 屋外" if e["type"] in ("ACF", "CT") else "全閉外扇"
    return f"{m.group(1)}kW {pole} {v} {enc}"


def kaimei(e):
    """基幹改良工事の機器番号（メーカーの改良工事図書の番号）"""
    uy = W.upgrade_year(e, "actual") or e.get("added")
    if not uy:
        return ""
    k = {2022: "K4", 2023: "K5", 2024: "K6", 2025: "K7"}[uy]
    return f"{k}-{e['unit']}{e['sysno']:02d}-{e['type']}{e['loop'] % 100:02d}"


def nested_rows(rows):
    """系統図ごとに 見出し → 機器 → 付属（電動機）"""
    out = []
    cur = None
    for e in rows:
        d = dwg(e)
        if d != cur:
            out.append(("H", d, e))
            cur = d
        out.append(("E", None, e))
        mt = motor(e)
        if mt:
            out.append(("M", mt, e))
    return out


def old_view(eq):
    """令和6年度版（R6.4.1）の台帳に載っている行。1号炉の R4 改良分は反映漏れ、2号炉 R5 分は反映済み（末尾に追記）"""
    base, tail = [], []
    for e in eq:
        u = e["unit"]
        asof = 2021 if u == "1" else 2023
        if u == "2" and e["added"] == 2023:
            tail.append(e)
            continue
        if u == "2" and e["removed"] == 2023:
            base.append(dict(e, _struck=True))
            continue
        if EQ.alive(e, asof):
            base.append(e)
    base.sort(key=lambda e: (UNIT_ORDER[e["unit"]], e["sysno"], e["loop"], e["tag"]))
    return base, tail


def build_old(eq, path):
    wb = Workbook()
    ws = wb.active
    ws.title = "機器台帳"
    F = lambda **k: font(FONT_G, **k)
    put(ws, 1, 1, "【架空】みなと臨海工場　機器台帳", F(size=14, bold=True))
    put(ws, 2, 1, "（令和6年度）　系統図（完成図書 P&ID）番号順", F(size=10))
    put(ws, 1, 9, "みなと市環境局施設課", F(size=9))
    put(ws, 2, 9, "H25.10 再編（h3）／最終更新 R6.4.1", F(size=9))
    put(ws, 3, 1, "※ P&ID タグNo. 1 件を 1 台とする（No. を振った行）。└ の付属行（電動機）は台数に含めない。小口径手動弁・配管は除く。", F(size=8))
    hdr = ["No.", "系統図\nNo.", "号炉", "タグNo.", "機器名称", "仕様・能力", "常用\n予備", "メーカー", "設置\n年月", "重要度", "保全\n区分", "備　考"]
    widths = [5.5, 7, 5, 12, 30, 34, 5, 14, 7, 5.5, 5.5, 22]
    for i, (h, w) in enumerate(zip(hdr, widths), start=1):
        put(ws, 5, i, h, F(size=9, bold=True), AL_C, B_ALL, fill("D9D9D9"))
        ws.column_dimensions[CL(i)].width = w
    ws.row_dimensions[5].height = 26
    base, tail = old_view(eq)
    r = 6
    n = 0
    prev_mk = None
    for kind, d, e in nested_rows(base) + [("T", None, None)] + [("E", None, x) for x in tail]:
        if kind == "H":
            merge(ws, r, 1, r, 12, f"{d[0]}　{d[1]}　　［{e['setsubi']}］", F(size=9, bold=True), AL_L, B_ALL, fill("EDEDED"))
            prev_mk = None
            r += 1
            continue
        if kind == "T":
            merge(ws, r, 1, r, 12, "追記（R5 基幹改良 2号炉）", F(size=9, bold=True, color="0000CC"), AL_L, B_ALL, fill("EDEDED"))
            r += 1
            continue
        if kind == "M":
            vals = [None, None, None, None, "　└ 電動機（付属）", d, None, "汐見電機", None, None, None, None]
            for c, v in enumerate(vals, start=1):
                put(ws, r, c, v, F(size=7, color="404040"), AL_L, Border(left=hair, right=hair, top=None, bottom=hair))
            r += 1
            continue
        n += 1
        st = e.get("_struck")
        fz = Font(name=FONT_G, size=8, strike=bool(st), color="0000CC" if e in tail else "000000")
        mk = "〃" if (e["maker"] == prev_mk and e not in tail) else e["maker"]
        prev_mk = e["maker"]
        duty = {"常用": "常用", "予備": "予備", "交互": "交互", "—": ""}[e["duty"]]
        biko = ""
        if st:
            biko = "R5 撤去（基幹改良）"
        elif e in tail:
            biko = "R5 基幹改良で追加（系統図 rev 待ち）"
        elif e["tag"] in ("C-CV-956", "C-HP-957"):
            biko = "主灰セメント原料化 追設"
        elif e["tag"] == "C-LOG-1118":
            biko = "H26 DCS一部更新時 追設"
        vals = [n, dwg(e)[0], UNIT_LABEL_OLD[e["unit"]], e["tag"], e["name"], e["spec"], duty, mk, inst_ym(e), e["imp"], hozen(e), biko]
        for c, v in enumerate(vals, start=1):
            al = AL_C if c in (1, 2, 3, 7, 9, 10, 11) or v == "〃" else AL_L
            put(ws, r, c, v, fz, al, B_HAIR)
        r += 1
    note(ws.cell(row=r - len(tail), column=5), "R5.12 2号炉 改良分を追記。系統図の場所へは後で移す（h5）", "h5")
    note(ws.cell(row=5, column=10), "A：停止で炉停止・法令影響　B：予備あり／一部影響　C：影響小（H25 h3）", "h3")
    note(ws.cell(row=5, column=2), "H25 再編時に竣工図書の P&ID 図番で並べ替え（h3）", "h3")
    last = r - 1
    put(ws, last + 2, 1, f"計　{n} 台（No. を振った行。うち撤去表示 1）", F(size=9))
    ws.freeze_panes = "E6"
    page(ws, landscape=True, paper=9, fit_w=1, fit_h=0, area=f"A1:L{last + 2}", titles="5:5",
         footer="&P / &N", header="【架空】みなと臨海工場 機器台帳")
    props(wb, "【架空】みなと臨海工場 機器台帳（令和6年度）", "h3", "機器台帳", dt.datetime(2013, 10, 21, 9, 0), modified_by="h5")
    wb.save(path)
    return n


def build_new(eq, path):
    """旧台帳をコピーして直したもの。系統図順の入れ子は残し、改良の列を足す。新設タグは別シートに足した"""
    wb = Workbook()
    ws = wb.active
    ws.title = "機器台帳"
    F = lambda **k: font(FONT_P, **k)
    G = lambda **k: font(FONT_G, **k)
    put(ws, 1, 1, "【架空】みなと臨海工場　機器台帳", F(size=14, bold=True))
    put(ws, 2, 1, "令和7年度（基幹的設備改良工事 完了反映）　系統図番号順", F(size=10))
    put(ws, 1, 12, "みなと市環境局施設課", F(size=9))
    put(ws, 2, 12, "R8.3.31 現在　h5", F(size=9))
    put(ws, 3, 1, "※ No. を振った行が管理対象（P&ID タグ 1 件＝1 台）。└ 付属行は数えない。基幹改良で新設したタグは「基幹改良新設」シート、撤去は「撤去機器」シート。", F(size=8))
    put(ws, 4, 1, "※ 改良工事機器番号：基幹的設備改良工事の完成図書の番号（K4〜K7）。市のタグNo. は変えずに併記する。", F(size=8))
    hdr = ["No.", "系統図\nNo.", "号炉", "タグNo.", "機器名称", "仕様・能力", "常用\n予備", "メーカー", "設置\n年月", "設置\n年度",
           "基幹改良", "改良\n年度", "改良工事\n機器番号", "重要度", "保全\n区分", "点検・整備周期\n（年）", "備　考"]
    widths = [5.5, 7, 5, 12, 30, 34, 5, 14, 7, 6, 7, 6, 13, 5.5, 5.5, 12, 30]
    for i, (h, w) in enumerate(zip(hdr, widths), start=1):
        put(ws, 6, i, h, F(size=9, bold=True), AL_C, B_ALL, fill("DDEBF7"))
        ws.column_dimensions[CL(i)].width = w
    ws.row_dimensions[6].height = 30
    ws.column_dimensions["I"].hidden = True  # 旧 設置年月（和暦）は非表示で残す
    note(ws.cell(row=6, column=10), "R7.4 列追加（西暦）。旧 設置年月（和暦）は I 列に非表示で残す（h5）", "h5")
    note(ws.cell(row=6, column=13), "改良工事の完成図書（汐見重工環境）の機器番号を併記（h5）", "h5")
    main = [e for e in eq if not e["added"]]
    main.sort(key=lambda e: (UNIT_ORDER[e["unit"]], e["sysno"], e["loop"], e["tag"]))
    r = 7
    n = 0
    n_alive = 0
    for kind, d, e in nested_rows(main):
        if kind == "H":
            rev = " rev.2（R7 改訂）" if (W.upgrade_year(e, "actual") or e["unit"] != "C" and e["sysno"] in (6, 8)) else ""
            merge(ws, r, 1, r, 17, f"{d[0]}{rev}　{d[1]}　　［{e['setsubi']}］", F(size=9, bold=True), AL_L, B_ALL, fill("F2F2F2"))
            r += 1
            continue
        if kind == "M":
            vals = [None] * 4 + ["　└ 電動機（付属）", d, None, "汐見電機"] + [None] * 9
            if W.upgrade_year(e, "actual") and e["type"] == "IDF":
                vals[5] = d.replace("全閉外扇", "全閉外扇（インバータ駆動）")
            for c, v in enumerate(vals, start=1):
                put(ws, r, c, v, F(size=7, color="404040"), AL_L, Border(left=hair, right=hair, bottom=hair))
            r += 1
            continue
        n += 1
        st = e["removed"] is not None
        if not st:
            n_alive += 1
        uy = W.upgrade_year(e, "actual")
        kk = "撤去" if st else ("更新" if uy else "")
        duty = {"常用": "常用", "予備": "予備", "交互": "交互", "—": ""}[e["duty"]]
        biko = "→「撤去機器」シート" if st else ""
        if e["tag"] in ("C-CV-956", "C-HP-957"):
            biko = "主灰セメント原料化 追設"
        elif e["tag"] == "C-LOG-1118":
            biko = "H26 DCS一部更新時 追設"
        vals = [n, dwg(e)[0], UNIT_LABEL_NEW[e["unit"]], e["tag"], e["name"], e["spec"], duty, e["maker"], inst_ym(e),
                e["installed"], kk, (e["removed"] if st else uy) or "", kaimei(e) if (uy and not st) else "", e["imp"],
                "" if st else hozen(e), "" if st else cycle_txt(e), biko]
        for c, v in enumerate(vals, start=1):
            al = AL_C if c in (1, 2, 3, 7, 9, 10, 11, 12, 14, 15) else AL_L
            put(ws, r, c, v, Font(name=FONT_P, size=8, strike=True, color="808080") if st else F(size=8), al, B_HAIR)
        if kk == "更新":
            ws.cell(row=r, column=11).fill = fill("FFF2CC")
        r += 1
    last = r - 1
    put(ws, last + 2, 1, f"本シート {n_alive} 台（取消線の撤去 {n - n_alive} 行を除く）＋ 基幹改良新設シート ＝ 集計シートのとおり", F(size=9))
    ws.freeze_panes = "E7"
    page(ws, landscape=True, paper=9, fit_w=1, fit_h=0, area=f"A1:Q{last + 2}", titles="6:6",
         footer="&P / &N", header="【架空】みなと臨海工場 機器台帳（令和7年度）")

    # 基幹改良新設（後から足したシート）
    w2 = wb.create_sheet("基幹改良新設（R4〜R7）")
    put(w2, 1, 1, "【架空】みなと臨海工場　基幹的設備改良工事で新設した機器（R4〜R7）", F(size=12, bold=True))
    put(w2, 2, 1, "R7.5 h5 作成。本表の機器は管理対象（台帳本表と合わせて数える）。系統図への反映は rev.2 で実施済み。", F(size=9))
    h2 = ["No.", "系統図\nNo.", "号炉", "タグNo.", "改良工事\n機器番号", "機器名称", "仕様・能力", "メーカー", "設置年度", "重要度", "保全\n区分",
          "点検・整備周期\n（年）", "置き換えた旧機器", "備考"]
    for i, (h, wdt) in enumerate(zip(h2, [5, 7, 6, 12, 13, 30, 30, 12, 7, 5.5, 5.5, 12, 22, 28]), start=1):
        put(w2, 4, i, h, F(size=9, bold=True), AL_C, B_ALL, fill("DDEBF7"))
        w2.column_dimensions[CL(i)].width = wdt
    w2.row_dimensions[4].height = 30
    added = sorted([e for e in eq if e["added"]], key=lambda e: (e["added"], UNIT_ORDER[e["unit"]], e["tag"]))
    rr = 5
    for i, e in enumerate(added, start=1):
        rep = f"{e['unit']}-MD-608 誘引通風機入口ダンパ" if e["type"] == "INV" else ("共通 FCS-C1〜C4・中継端子盤" if e["unit"] == "C" else "—（新規）")
        bk = "R4 分 台帳反映漏れ → R7.5 追記（h5）" if e["unit"] == "1" else ""
        vals = [i, dwg(e)[0], UNIT_LABEL_NEW[e["unit"]], e["tag"], kaimei(e), e["name"], e["spec"], e["maker"], e["installed"], e["imp"],
                hozen(e), cycle_txt(e), rep, bk]
        for c, v in enumerate(vals, start=1):
            put(w2, rr, c, v, F(size=8), AL_C if c in (1, 2, 3, 9, 10, 11) else AL_L, B_HAIR)
        rr += 1
    page(w2, landscape=True, fit_w=1, fit_h=0, area=f"A1:N{rr}", titles="4:4")

    # 撤去機器
    w3 = wb.create_sheet("撤去機器")
    put(w3, 1, 1, "【架空】みなと臨海工場　撤去機器一覧（基幹的設備改良工事 R4〜R7）", F(size=12, bold=True))
    h3 = ["No.", "号炉", "タグNo.", "機器名称", "仕様", "撤去年度", "代替", "備考"]
    for i, (h, wdt) in enumerate(zip(h3, [5, 7, 12, 30, 24, 8, 34, 30]), start=1):
        put(w3, 3, i, h, F(size=9, bold=True), AL_C, B_ALL, fill("DDEBF7"))
        w3.column_dimensions[CL(i)].width = wdt
    rr = 4
    rem = sorted([e for e in eq if e["removed"]], key=lambda e: (e["removed"], e["tag"]))
    for i, e in enumerate(rem, start=1):
        alt = f"{e['unit']}-INV-815 誘引通風機インバータ盤（回転数制御）" if e["type"] == "MD" else "CTL-C1〜C3・リモート I/O 盤（C-IO-1141〜1143）"
        bk = "R6 版台帳では撤去表示なし（反映漏れ）" if e["tag"] == "1-MD-608" else ""
        for c, v in enumerate([i, UNIT_LABEL_NEW[e["unit"]], e["tag"], e["name"], e["spec"], e["removed"], alt, bk], start=1):
            put(w3, rr, c, v, F(size=9), AL_C if c in (1, 2, 6) else AL_L, B_HAIR)
        rr += 1
    page(w3, landscape=True, fit_w=1, fit_h=0, area=f"A1:H{rr}")

    # 集計
    w4 = wb.create_sheet("集計")
    put(w4, 1, 1, "【架空】みなと臨海工場　管理対象機器 集計（R8.3.31）", F(size=12, bold=True))
    cols = ["シート", "1号炉", "2号炉", "3号炉", "共通", "計"]
    for i, h in enumerate(cols, start=1):
        put(w4, 3, i, h, F(size=9, bold=True), AL_C, B_ALL, fill("DDEBF7"))
        w4.column_dimensions[CL(i)].width = 24 if i == 1 else 9
    alive_main = [e for e in main if e["removed"] is None]
    rr = 4
    tot = {u: 0 for u in "123C"}
    for label, rows in (("機器台帳（本表、撤去除く）", alive_main), ("基幹改良新設（R4〜R7）", added)):
        cnt = {u: sum(1 for e in rows if e["unit"] == u) for u in "123C"}
        for u in "123C":
            tot[u] += cnt[u]
        for c, v in enumerate([label, cnt["1"], cnt["2"], cnt["3"], cnt["C"], sum(cnt.values())], start=1):
            put(w4, rr, c, v, F(size=9), AL_L if c == 1 else AL_R, B_ALL)
        rr += 1
    for c, v in enumerate(["合計（管理対象）", tot["1"], tot["2"], tot["3"], tot["C"], sum(tot.values())], start=1):
        put(w4, rr, c, v, F(size=9, bold=True), AL_L if c == 1 else AL_R, B_ALL, fill("F2F2F2"))
    rr += 2
    put(w4, rr, 1, "設備区分別（本表＋新設）", F(size=9, bold=True))
    rr += 1
    allrows = alive_main + added
    for s in SETSUBI_ORDER:
        cnt = {u: sum(1 for e in allrows if e["setsubi"] == s and e["unit"] == u) for u in "123C"}
        if sum(cnt.values()) == 0:
            continue
        for c, v in enumerate([s, cnt["1"], cnt["2"], cnt["3"], cnt["C"], sum(cnt.values())], start=1):
            put(w4, rr, c, v, F(size=9), AL_L if c == 1 else AL_R, B_ALL)
        rr += 1
    page(w4, landscape=False, fit_w=1, fit_h=1, area=f"A1:F{rr}")
    props(wb, "【架空】みなと臨海工場 機器台帳（令和7年度 基幹改良反映）", "h3", "機器台帳",
          dt.datetime(2013, 10, 21, 9, 0), modified_by="h5")
    wb.save(path)
    return n_alive + len(added)



def main():
    eq = EQ.build()
    p1 = os.path.join(OUT, "【架空】みなと臨海工場_機器台帳_令和6年度_平成25年度再編様式.xlsx")
    p2 = os.path.join(OUT, "【架空】みなと臨海工場_機器台帳_令和7年度_基幹改良反映様式.xlsx")
    n1 = build_old(eq, p1)
    n2 = build_new(eq, p2)
    print("ledger rows", n1, n2)
    return [p1, p2]


if __name__ == "__main__":
    main()
