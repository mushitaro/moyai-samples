# -*- coding: utf-8 -*-
"""機器台帳（版）
H16 竣工図書の機器表（メーカー作成、号炉の列が横に並ぶ）／ H25・H29・R3・R6 h3 再編様式（P&ID 図番順の入れ子）／ R7 h5 基幹改良反映様式
"""
import datetime as dt
import os

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter as CL

import ans
import core
import equipment as EQ
import s_ledger as L
from xlhelp import *

UN = {"1": "1号炉", "2": "2号炉", "3": "3号炉", "C": "共通"}


def build_maker(path):
    """竣工図書の機器表：炉別機器は 1 行に 1〜3 号炉のタグを横に並べる"""
    eq = [e for e in core.get().eq if EQ.alive(e, 2004)]
    wb = Workbook()
    ws = wb.active
    ws.title = "炉設備"
    F = lambda **k: font(FONT_G, **k)
    fn = os.path.basename(path)
    ans.start(fn, False, "機器台帳", 2004, "竣工図書 機器表（メーカー）", "メーカー 設計部")
    def head(ws, title, cols, widths):
        put(ws, 1, 1, title, F(size=13, bold=True))
        put(ws, 2, 1, "【架空】汐見重工環境株式会社　完成図書 第 3 分冊 機器表（平成16年3月）　みなと臨海工場 ごみ焼却施設建設工事", F(size=9))
        for i, (h, w) in enumerate(zip(cols, widths), start=1):
            put(ws, 4, i, h, F(size=8, bold=True), AL_C, B_ALL, fill("D9D9D9"))
            ws.column_dimensions[CL(i)].width = w
        ws.row_dimensions[4].height = 26
    head(ws, "【架空】みなと臨海工場　機器表（炉設備）", ["機器\n番号", "機器名称", "仕様", "1号炉", "2号炉", "3号炉", "数量", "製作者", "系統図"],
         [7, 30, 36, 12, 12, 12, 5, 14, 8])
    rows = {}
    for e in eq:
        if e["unit"] == "C":
            continue
        key = (e["sysno"], e["loop"], e["name"])
        rows.setdefault(key, {})[e["unit"]] = e
    r = 5
    cur = None
    for key in sorted(rows):
        d = rows[key]
        e0 = d.get("1") or list(d.values())[0]
        if e0["sysno"] != cur:
            cur = e0["sysno"]
            merge(ws, r, 1, r, 9, f"{cur}　{L.SYSNAME_F[cur].replace('系統図', '').replace('・計装ループ図', '')}（{e0['setsubi']}）", F(size=9, bold=True), AL_L, B_ALL, fill("EDEDED"))
            ans.rec(fn, ws.title, r, "parent", dict(path=[e0["setsubi"], L.SYSNAME_F[cur]]), check=[e0["setsubi"]])
            r += 1
        vals = [f"{e0['type']}-{e0['loop']}", e0["name"], e0["spec"], d.get("1", {}).get("tag"), d.get("2", {}).get("tag"), d.get("3", {}).get("tag"),
                len(d), e0["maker"], f"P-{e0['sysno']:02d}"]
        for c, v in enumerate(vals, start=1):
            put(ws, r, c, v, F(size=8), AL_C if c in (1, 4, 5, 6, 7, 9) else AL_L, B_HAIR)
        ans.rec(fn, ws.title, r, "item", dict(path=[e0["setsubi"], e0["name"]], name=e0["name"], spec=e0["spec"],
                                              codes={u: d[u]["tag"] for u in sorted(d)}, 数量=len(d), maker=e0["maker"]), check=[vals[3] or vals[4]])
        r += 1
    page(ws, landscape=True, fit_w=1, fit_h=0, area=f"A1:I{r}", titles="4:4", footer="&P / &N")
    w2 = wb.create_sheet("共通設備")
    head(w2, "【架空】みなと臨海工場　機器表（共通設備）", ["機器\n番号", "機器名称", "仕様", "機器番号（タグ）", "常用／予備", "数量", "製作者", "系統図"],
         [7, 30, 36, 13, 8, 5, 14, 8])
    r = 5
    cur = None
    for e in sorted([e for e in eq if e["unit"] == "C"], key=lambda e: (e["sysno"], e["loop"], e["tag"])):
        if e["sysno"] != cur:
            cur = e["sysno"]
            merge(w2, r, 1, r, 8, f"C{cur:02d}　{L.SYSNAME_C[cur]}（{e['setsubi']}）", F(size=9, bold=True), AL_L, B_ALL, fill("EDEDED"))
            ans.rec(fn, w2.title, r, "parent", dict(path=[e["setsubi"], L.SYSNAME_C[cur]]), check=[e["setsubi"]])
            r += 1
        vals = [f"{e['type']}-{e['loop']}", e["name"], e["spec"], e["tag"], e["duty"] if e["duty"] != "—" else "", 1, e["maker"], f"P-C{e['sysno']:02d}"]
        for c, v in enumerate(vals, start=1):
            put(w2, r, c, v, F(size=8), AL_C if c in (1, 4, 5, 6, 8) else AL_L, B_HAIR)
        ans.rec(fn, w2.title, r, "item", dict(path=[e["setsubi"], e["name"]], name=e["name"], spec=e["spec"], code=e["tag"], 数量=1, maker=e["maker"]),
                check=[e["tag"]])
        r += 1
    page(w2, landscape=True, fit_w=1, fit_h=0, area=f"A1:H{r}", titles="4:4", footer="&P / &N")
    props(wb, "【架空】みなと臨海工場 完成図書 機器表", "汐見重工環境 設計部", "機器台帳", dt.datetime(2004, 3, 15, 9, 0))
    wb.save(path)


def parse_records(path, fn, sheet_names=None):
    """生成した h3/h5 様式の台帳を読み、行ごとの正解を作る（No. のある行＝機器、└ の行＝付属、図番の見出し＝親）"""
    wb = load_workbook(path)
    for ws in wb.worksheets:
        if sheet_names and ws.title not in sheet_names:
            continue
        hdr_row = None
        for r in range(1, 10):
            if ws.cell(row=r, column=1).value == "No.":
                hdr_row = r
                break
        if not hdr_row:
            continue
        hdr = [ws.cell(row=hdr_row, column=c).value for c in range(1, ws.max_column + 1)]
        hdr = [(h or "").replace("\n", "") for h in hdr]
        parent = None
        for r in range(hdr_row + 1, ws.max_row + 1):
            a = ws.cell(row=r, column=1).value
            vals = {hdr[c - 1]: ws.cell(row=r, column=c).value for c in range(1, len(hdr) + 1) if hdr[c - 1]}
            if isinstance(a, str) and (a.startswith("P-") or a.startswith("追記") or a.startswith("■")):
                parent = a
                ans.rec(fn, ws.title, r, "parent", dict(path=[a]), check=[a[:5]])
                continue
            if isinstance(a, int):
                tag = vals.get("タグNo.")
                f = {k: v for k, v in vals.items() if v not in (None, "") and k != "No."}
                f["path"] = [parent, vals.get("機器名称")] if parent else [vals.get("機器名称")]
                f["code"] = tag
                f["struck"] = bool(ws.cell(row=r, column=4).font.strike)
                ans.rec(fn, ws.title, r, "item", f, check=[tag])
                continue
            e = ws.cell(row=r, column=5).value
            if isinstance(e, str) and e.strip().startswith("└"):
                f = {k: v for k, v in vals.items() if v not in (None, "")}
                f["path"] = [parent, "付属（電動機）"]
                ans.rec(fn, ws.title, r, "item", f, check=["└"])


def retitle(path, label, upd, who):
    wb = load_workbook(path)
    ws = wb["機器台帳"]
    ws["A2"] = f"（{label}）　系統図（完成図書 P&ID）番号順"
    ws["I2"] = f"H25.10 再編（h3）／最終更新 {upd}"
    wb.properties.title = f"【架空】みなと臨海工場 機器台帳（{label}）"
    wb.properties.lastModifiedBy = who
    wb.save(path)


def main():
    out = []
    p = os.path.join(OUT, "【架空】みなと臨海工場_機器台帳_平成16年度_竣工図書機器表様式.xlsx")
    build_maker(p)
    out.append(p)
    eq = core.get().eq
    orig_view = L.old_view
    for v, upd, who in ((2013, "H25.10.21", "h3"), (2017, "H29.6.1", "h4"), (2021, "R3.4.1", "h4")):
        L.old_view = lambda eq_, v=v: (sorted([e for e in eq_ if EQ.alive(e, v)], key=lambda e: (L.UNIT_ORDER[e["unit"]], e["sysno"], e["loop"], e["tag"])), [])
        path = os.path.join(OUT, f"【架空】みなと臨海工場_機器台帳_{core.wa_fy(v)}_平成25年度再編様式.xlsx")
        L.build_old(eq, path)
        retitle(path, core.wa_fy(v), upd, who)
        fn = os.path.basename(path)
        ans.start(fn, False, "機器台帳", v, "h3 再編（P&ID 図番順の入れ子）", who)
        parse_records(path, fn)
        out.append(path)
    L.old_view = orig_view
    path = os.path.join(OUT, "【架空】みなと臨海工場_機器台帳_令和6年度_平成25年度再編様式.xlsx")
    L.build_old(eq, path)
    fn = os.path.basename(path)
    ans.start(fn, False, "機器台帳", 2024, "h3 再編（P&ID 図番順の入れ子）", "h5")
    parse_records(path, fn)
    out.append(path)
    path = os.path.join(OUT, "【架空】みなと臨海工場_機器台帳_令和7年度_基幹改良反映様式.xlsx")
    L.build_new(eq, path)
    fn = os.path.basename(path)
    ans.start(fn, False, "機器台帳", 2025, "h5 基幹改良反映（改良番号併記・新設／撤去シート）", "h5")
    parse_records(path, fn)
    out.append(path)
    return out


if __name__ == "__main__":
    print(len(main()))
