# -*- coding: utf-8 -*-
"""中長期整備計画（星取表）。
旧様式：市 i1 が H21.12 に策定した「長期整備計画（H22–H31）」の H27.5 改訂版（i2 が i1 の様式のまま改訂）。
  シート「星取表」とシート「費用」が別で、行の並びは同じ（費用はメーカー見積の機器ごとの概算、千円）。
  階層は 設備 → 号炉 → 系統 の 3 列を縦に結合し、機器名の付属機器は「└」で字下げ（4〜5 段）。
新様式：i2 が H28.6 に星取表をコピーして作り替えた「停止までの整備計画（H29–R6）」の R4.10 改訂版（最後の更新）。
  各年度の下に「計画・実績」の 2 段。費用は機器ごとには書かず、設備の小計行の年度欄（2 列結合）にだけ書く
  （財政への説明が設備単位だったため）。R3 まで実績額、R4 以降は計画額。
  旧表の「周期」列は隠して残し、旧「費用」シートも隠しシートのまま残っている（更新していない）。"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from openpyxl import Workbook  # noqa: E402
from openpyxl.comments import Comment  # noqa: E402
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side  # noqa: E402
from openpyxl.utils import get_column_letter  # noqa: E402
from common import FAC_K, FONT, rng, set_props  # noqa: E402
from rec import Recs, path as out_path  # noqa: E402
import world  # noqa: E402
from world import SETSUBI, PARENT  # noqa: E402
from works import EQ, LEGAL, STRAT, DEFAULT, actual_works, plan_new, plan_old, policy  # noqa: E402

thin, hair = Side(style="thin"), Side(style="hair")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
ROW = Border(left=thin, right=thin, top=hair, bottom=hair)
F = Font(name=FONT, size=9)
FB = Font(name=FONT, size=9, bold=True)
CE = Alignment(horizontal="center", vertical="center", wrap_text=True)
LE = Alignment(horizontal="left", vertical="center", wrap_text=True)
VT = Alignment(horizontal="center", vertical="center", wrap_text=True, text_rotation=255)
GRAY = PatternFill("solid", fgColor="D9D9D9")
SUBF = PatternFill("solid", fgColor="F2F2F2")
NEWROW = PatternFill("solid", fgColor="FFF2CC")
LATE = PatternFill("solid", fgColor="EDEDED")
LINE_NAME = {0: "共通", 1: "1号炉", 2: "2号炉"}

COST = {"○": {"A": 600, "B": 150, "C": 60}, "◎": {"A": 3800, "B": 900, "C": 300},
        "●": {"A": 42000, "B": 6000, "C": 1500}, "△": {"A": 5200, "B": 800, "C": 250},
        "▲": {"A": 2400, "B": 500, "C": 200}, "◇": {"A": 1800, "B": 400, "C": 150}}
SPECIAL_COST = {("バグフィルタ", "●"): 21600, ("中央監視制御盤", "●"): 186000,
                ("ごみクレーン", "●"): 98000, ("前処理破砕機", "●"): 64000,
                ("散気装置", "◇"): 6800, ("バグフィルタ", "◇"): 1900,
                ("中央監視制御盤", "◇"): 3600}


def cost(e, sym):
    v = SPECIAL_COST.get((e["base"], sym))
    return COST[sym][e["imp"]] if v is None else v


def legal_years(name, fys):
    if "精密機能検査" in name:
        return [y for y in fys if (y - 2008) % 3 == 0]
    if "クレーン性能検査" in name:
        return [y for y in fys if y % 2 == 0]
    if "計量機" in name:
        return [y for y in fys if y % 2 == 1]
    return list(fys)


def setchi(e):
    return "H2" if e["year"] == 1989 else f"H{e['year'] + 1 - 1988}"


def cycle(e):
    ins, oh = STRAT.get(e["base"], DEFAULT[e["imp"]])
    return "" if not (ins or oh) else str(oh or ins)


def eq_name(e):
    return ("　└ " if e["base"] in PARENT else "") + e["name"]


def build_rows(legal_skip=()):
    """行の並び（星取表・費用・新様式で共通）。kind: legal / eq / sub。"""
    rows = [dict(kind="legal", name=n, cyc=c, law=l) for (n, c, l) in LEGAL if not any(k in n for k in legal_skip)]
    cur = None
    for e in world.tree_order(EQ):
        if cur is not None and cur != e["setsubi"]:
            rows.append(dict(kind="sub", setsubi=cur))
        cur = e["setsubi"]
        rows.append(dict(kind="eq", e=e))
    rows.append(dict(kind="sub", setsubi=cur))
    return rows


def merge_runs(ws, rows, r0, col, keyf):
    """keyf(row) が同じで続く行を縦に結合する（None は結合しない）。"""
    start, prev = None, object()
    for k, row in enumerate(rows + [None]):
        key = keyf(row) if row is not None else object()
        if key != prev:
            if start is not None and prev is not None and k - 1 > start:
                ws.merge_cells(start_row=r0 + start, start_column=col, end_row=r0 + k - 1, end_column=col)
            start, prev = k, key


def left_block(ws, rows, r0, ncols, with_pol=None):
    """A〜G（＋新様式は H）の左側：設備・号炉・系統の結合、機器番号・名前・設置・周期。"""
    for k, row in enumerate(rows):
        r = r0 + k
        if row["kind"] == "legal":
            ws.cell(row=r, column=1, value="法定点検等")
            ws.cell(row=r, column=2, value=row["name"])
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
            ws.cell(row=r, column=7, value=f"{row['cyc']}年")
        elif row["kind"] == "eq":
            e = row["e"]
            ws.cell(row=r, column=1, value=SETSUBI[e["setsubi"]])
            ws.cell(row=r, column=2, value=LINE_NAME[e["line"]])
            ws.cell(row=r, column=3, value=e["keitou"])
            ws.cell(row=r, column=4, value=e["id"])
            ws.cell(row=r, column=5, value=eq_name(e))
            ws.cell(row=r, column=6, value=setchi(e))
            ws.cell(row=r, column=7, value=cycle(e))
            if with_pol:
                ws.cell(row=r, column=8, value=with_pol[e["id"]])
        else:
            ws.cell(row=r, column=1, value=SETSUBI[row["setsubi"]])
            ws.cell(row=r, column=2, value="小　計")
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7 + (1 if with_pol else 0))
        for j in range(1, ncols + 1):
            c = ws.cell(row=r, column=j)
            c.font = FB if row["kind"] == "sub" else F
            c.border = ROW if row["kind"] == "eq" else BOX
            c.alignment = LE if j in (2, 3, 5) else CE
            if row["kind"] == "sub":
                c.fill = SUBF
    merge_runs(ws, rows, r0, 1, lambda x: "L" if x["kind"] == "legal" else x.get("setsubi") or x["e"]["setsubi"])
    merge_runs(ws, rows, r0, 2, lambda x: (x["e"]["setsubi"], x["e"]["line"]) if x["kind"] == "eq" else None)
    merge_runs(ws, rows, r0, 3, lambda x: (x["e"]["setsubi"], x["e"]["line"], x["e"]["keitou"]) if x["kind"] == "eq" else None)
    for k in range(len(rows)):
        ws.cell(row=r0 + k, column=1).alignment = VT
        if rows[k]["kind"] == "eq":
            ws.cell(row=r0 + k, column=2).alignment = CE
            ws.cell(row=r0 + k, column=3).alignment = CE


def page_setup(ws, last_col, last_row, title_rows):
    ws.print_title_rows = title_rows
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_area = f"A1:{last_col}{last_row}"
    ws.oddFooter.center.text = "&A　&P / &N"


LEFT = ["設備", "号炉", "系統", "機器番号", "機器名", "設置", "周期"]
LEFTW = [5, 6, 9, 8, 26, 5, 4.5]


def harvest(R, ws, rows, r0, ycol, mode, acol=None):
    """書いたシートから正解の行を拾う（ファイルの中身と必ず一致させるため）。"""
    for k, row in enumerate(rows):
        r = r0 + k
        if row["kind"] == "sub":
            f = {"path": [SETSUBI[row["setsubi"]]]}
            if mode in ("cost", "new"):
                f["cost"] = {str(y): ws.cell(row=r, column=c).value for y, c in ycol.items() if ws.cell(row=r, column=c).value}
                f["unit"] = "千円"
            else:
                f["memo"] = "（費用シート）"
            R.add(ws.title, r, "subtotal", f)
            continue
        if row["kind"] == "legal":
            f = {"path": ["法定点検等", row["name"]], "cycle": row["cyc"]}
        else:
            e = row["e"]
            f = {"path": [SETSUBI[e["setsubi"]], LINE_NAME[e["line"]], e["keitou"], e["name"]], "code": e["id"],
                 "cycle": int(cycle(e)) if cycle(e) else None}
            if mode == "new":
                f["policy"] = ws.cell(row=r, column=8).value
        if mode == "cost":
            f["cost"] = {str(y): ws.cell(row=r, column=c).value for y, c in ycol.items() if ws.cell(row=r, column=c).value}
            f["unit"] = "千円"
            memo = ws.cell(row=r, column=max(ycol.values()) + 1).value
        else:
            f["marks"] = {str(y): ws.cell(row=r, column=c).value for y, c in ycol.items() if ws.cell(row=r, column=c).value}
            if acol:
                f["marksActual"] = {str(y): ws.cell(row=r, column=c).value for y, c in acol.items()
                                    if ws.cell(row=r, column=c).value}
            memo = ws.cell(row=r, column=(max(acol.values()) if acol else max(ycol.values())) + 1).value
            cm = {}
            for c in list(ycol.values()):
                cell = ws.cell(row=r, column=c)
                if cell.comment:
                    cm[cell.coordinate] = cell.comment.text
            if cm:
                f["comments"] = cm
        if memo:
            f["memo"] = memo
        wids = None
        if row["kind"] == "eq" and mode != "cost":
            wids = [w["wid"] for w in _wby().get(row["e"]["id"], []) if w["fy"] in ycol]
        R.add(ws.title, r, "item", f, wids)


WBY = {}


def _wby():
    if not WBY:
        for w in actual_works():
            WBY.setdefault(w["eq"], []).append(w)
    return WBY


# ====================================================================== 旧様式
def old_cost_sheet(ws, rows, P, years, title_note, ver=2015, R=None):
    """費用シート（行の並びは星取表と同じ）。小計は計算済みの数値。燃焼設備の H27 小計だけ改訂前の値。"""
    ws["A1"] = f"{FAC_K}　長期整備計画（平成22年度～平成31年度）　概算費用"
    ws["A1"].font = Font(name=FONT, size=13, bold=True)
    ws["A2"] = title_note
    ws["A2"].font = F
    heads = LEFT + [f"H{y - 1988}" for y in years] + ["計（千円）"]
    for j, h in enumerate(heads, 1):
        c = ws.cell(row=4, column=j, value=h)
        c.font, c.alignment, c.border, c.fill = FB, CE, BOX, GRAY
        ws.column_dimensions[get_column_letter(j)].width = (LEFTW + [7.5] * len(years) + [9])[j - 1]
    left_block(ws, rows, 5, len(heads))
    sub = {}
    grand = {y: 0 for y in years}
    for k, row in enumerate(rows):
        r = 5 + k
        if row["kind"] == "legal":
            ws.cell(row=r, column=len(heads), value="委託料で別計上").font = F
            continue
        if row["kind"] == "eq":
            e = row["e"]
            tot = 0
            for i, y in enumerate(years):
                s = P[e["id"]].get(y)
                if s:
                    v = cost(e, s)
                    tot += v
                    sub.setdefault(e["setsubi"], {}).setdefault(y, 0)
                    sub[e["setsubi"]][y] += v
                    ws.cell(row=r, column=8 + i, value=v).number_format = "#,##0"
            ws.cell(row=r, column=len(heads), value=tot or None).number_format = "#,##0"
        else:
            s = row["setsubi"]
            tot = 0
            for i, y in enumerate(years):
                v = sub.get(s, {}).get(y, 0)
                if s == "B" and y == 2015 and ver == 2015:
                    v -= 600            # 改訂前の小計のまま（わざと入れた食い違い）
                grand[y] += v
                tot += v
                ws.cell(row=r, column=8 + i, value=v or None).number_format = "#,##0"
            ws.cell(row=r, column=len(heads), value=tot).number_format = "#,##0"
    r = 5 + len(rows)
    ws.cell(row=r, column=2, value="合　計")
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
    for i, y in enumerate(years):
        ws.cell(row=r, column=8 + i, value=grand[y]).number_format = "#,##0"
    ws.cell(row=r, column=len(heads), value=sum(grand.values())).number_format = "#,##0"
    for j in range(1, len(heads) + 1):
        c = ws.cell(row=r, column=j)
        c.font, c.border = FB, BOX
    ws.cell(row=r + 1, column=1, value="※ 単位 千円（税込）。H21年度メーカー見積による概算。小計は設備ごと。").font = F
    ws.freeze_panes = "H5"
    page_setup(ws, get_column_letter(len(heads)), r + 1, "4:4")
    if R is not None:
        ycol = {y: 8 + i for i, y in enumerate(years)}
        harvest(R, ws, rows, 5, ycol, "cost")
        R.add(ws.title, r, "total", {"cost": {str(y): grand[y] for y in years}, "unit": "千円"})
    return r


OLDV = {2009: ("平成21年度策定", "平成21年12月策定（環境施設課）"),
        2012: ("平成24年度改訂", "平成21年12月策定（環境施設課）　平成24年5月改訂"),
        2015: ("平成27年度改訂", "平成21年12月策定（環境施設課）　平成27年5月改訂")}


def old_plan(ver=2015):
    P, comments = plan_old(ver)
    years = list(range(2010, 2020))
    rows = build_rows(legal_skip=("作業環境測定", "第二種圧力容器"))
    wb = Workbook()
    ws = wb.active
    ws.title = "星取表"
    ws["A1"] = f"{FAC_K}　長期整備計画（平成22年度～平成31年度）"
    ws["A1"].font = Font(name=FONT, size=14, bold=True)
    ws["A2"] = OLDV[ver][1] + "　　費用は「費用」シート（行は同じ並び）"
    ws["A2"].font = F
    ws["H2"] = "凡例　○：点検・整備　◎：分解整備（オーバーホール）　●：更新　△：補修"
    ws["H2"].font = F
    heads = LEFT + [f"H{y - 1988}" for y in years] + ["備考"]
    for j, h in enumerate(heads, 1):
        c = ws.cell(row=4, column=j, value=h)
        c.font, c.alignment, c.border, c.fill = FB, CE, BOX, GRAY
        ws.column_dimensions[get_column_letter(j)].width = (LEFTW + [5] * len(years) + [26])[j - 1]
    left_block(ws, rows, 5, len(heads))
    for k, row in enumerate(rows):
        r = 5 + k
        if row["kind"] == "legal":
            for i, y in enumerate(years):
                if y in legal_years(row["name"], years):
                    ws.cell(row=r, column=8 + i, value="○")
            ws.cell(row=r, column=len(heads), value=row["law"])
        elif row["kind"] == "eq":
            e = row["e"]
            for i, y in enumerate(years):
                s = P[e["id"]].get(y)
                if s:
                    c = ws.cell(row=r, column=8 + i, value=s)
                    if (e["id"], y) in comments:
                        c.comment = Comment(comments[(e["id"], y)], "i1" if ver < 2014 else "i2")
            note = e["duty"]
            if any((e["id"], y) in comments for y in years):
                note = (note + " ※コメント参照").strip()
            ws.cell(row=r, column=len(heads), value=note)
        else:
            ws.cell(row=r, column=8, value="（費用シート）")
            ws.merge_cells(start_row=r, start_column=8, end_row=r, end_column=7 + len(years))
        ws.cell(row=r, column=len(heads)).alignment = LE
    r = 5 + len(rows)
    ws.cell(row=r + 1, column=1, value="※ 機器番号は機器台帳（H17.4）による。└ は付属機器。").font = F
    ws.freeze_panes = "F5"
    page_setup(ws, get_column_letter(len(heads)), r + 1, "4:4")
    fn = f"{FAC_K}_中長期整備計画_{OLDV[ver][0]}_旧様式（市i1・星取表と費用表の2シート）.xlsx"
    R = Recs(fn, plan=True)
    harvest(R, ws, rows, 5, {y: 8 + i for i, y in enumerate(years)}, "star")
    old_cost_sheet(wb.create_sheet("費用"), rows, P, years, OLDV[ver][1], ver, R)
    set_props(wb, f"{FAC_K} 長期整備計画（H22～H31）{OLDV[ver][0]}")
    wb.save(out_path(fn))
    R.save()
    return fn


# ====================================================================== 新様式
def actual_sym(w):
    if w["kind"] == "REPLACE":
        return "●"
    if w["kind"] == "INSPECT":
        return "○"
    if "予備" in w["label"]:
        return "◇"
    if w["fy"] >= 2019:
        return "▲"
    return "△" if "補修" in w["label"] else "◎"


NOTES = {
    "2B-002": "R3点検 閉塞・摩耗 18本。R4予備品120本購入予定（納期5か月）。R5ノズル取替",
    "1E-007": "R3点検 ろ布破損8本（閉止）。R4部分交換（予備ろ布）",
    "0M-006": "更新中止（H28）。予備基板で延命。R2 CPU基板交換",
    "0A-019": "予備機。R3 減速機をNo.1へ転用",
    "1C-005": "R2 チェーン延び測定 余裕あり→見送り",
    "0M-001": "R2 故障・更新（計画外）",
}


NEWV = {2016: ("平成28年度作成", "平成28年度作成（H28.11 方針決定：令和6年度末停止→H36年度末）　H28年12月作成"),
        2019: ("令和元年度改訂", "平成28年度作成（H28.11 方針決定：令和6年度末停止）　令和元年10月改訂（実績はH30年度まで）"),
        2022: ("令和4年度改訂", "平成28年度作成（H28.11 方針決定：令和6年度末停止）　令和4年10月改訂（実績はR3年度まで）")}


def new_plan(ver=2022):
    P = plan_new(ver)
    pol = dict(policy())
    if ver < 2021:
        pol["0A-019"] = "停止まで使用"
    if ver < 2020:
        pol["0M-001"] = "停止まで使用"
    years = list(range(2017, 2025))
    known = {2016: 2016, 2019: 2018, 2022: 2021}[ver]
    rr_ = rng("newcost")
    act, acost = {}, {}
    for w in actual_works():
        if 2017 <= w["fy"] <= known:
            e = next(x for x in EQ if x["id"] == w["eq"])
            s = actual_sym(w)
            act.setdefault(w["eq"], {})[w["fy"]] = s
            acost.setdefault(w["eq"], {})[w["fy"]] = int(round(cost(e, s) * rr_.uniform(0.85, 1.15) / 10.0)) * 10
    rows = build_rows()
    wb = Workbook()
    ws = wb.active
    ws.title = "停止までの整備計画"
    ws["A1"] = f"{FAC_K}　停止までの整備計画（平成29年度～" + ("平成36年度）" if ver == 2016 else "令和6年度）")
    ws["A1"].font = Font(name=FONT, size=14, bold=True)
    ws["A1"].comment = Comment("H28.6 長期整備計画（H22-H31）の星取表をコピーして作成。費用シートは旧版のまま非表示 i2", "i2")
    ws["A2"] = NEWV[ver][1].replace("令和6年度末停止→H36年度末", "平成36年度末停止") + (
        "　小計行の年度欄は費用（千円）：計画額" if ver == 2016 else
        f"　小計行の年度欄は費用（千円）：H{known - 1988}まで実績額、以降は計画額" if ver == 2019 else
        "　小計行の年度欄は費用（千円）：R3まで実績額、R4以降は計画額")
    ws["A2"].font = F
    ws["I3"] = ("凡例　○：点検　◎：分解整備　△：補修　▲：停止までの最小限の補修　"
                "◇：予備品で交換　●：更新（やむを得ないもの）")
    ws["I3"].font = F
    left = LEFT + ["停止までの扱い"]
    for j, h in enumerate(left, 1):
        ws.merge_cells(start_row=4, start_column=j, end_row=5, end_column=j)
        c = ws.cell(row=4, column=j, value=h)
        c.font, c.alignment, c.fill = FB, CE, GRAY
        ws.column_dimensions[get_column_letter(j)].width = (LEFTW + [11])[j - 1]
    ws.column_dimensions["G"].hidden = True       # 旧表の「周期」を消さずに隠した
    ycol = {}
    col = len(left) + 1
    for y in years:
        ycol[y] = col
        lab = (f"H{y - 1988}" if ver == 2016 else {2019: "H31/R1"}.get(y, f"H{y - 1988}" if y < 2019 else f"R{y - 2018}"))
        ws.merge_cells(start_row=4, start_column=col, end_row=4, end_column=col + 1)
        ws.cell(row=4, column=col, value=lab)
        for k, h in enumerate(("計画", "実績")):
            c = ws.cell(row=5, column=col + k, value=h)
            ws.column_dimensions[get_column_letter(col + k)].width = 4.6
            if y > known and k == 1:
                c.fill = LATE
        col += 2
    ncol = col
    ws.cell(row=4, column=ncol, value="備考")
    ws.merge_cells(start_row=4, start_column=ncol, end_row=5, end_column=ncol)
    ws.column_dimensions[get_column_letter(ncol)].width = 30
    for rr in (4, 5):
        for j in range(1, ncol + 1):
            c = ws.cell(row=rr, column=j)
            c.font, c.alignment, c.border = FB, CE, BOX
            if c.fill.fgColor.rgb in (None, "00000000"):
                c.fill = GRAY
    left_block(ws, rows, 6, ncol, with_pol=pol)
    sub = {}
    for k, row in enumerate(rows):
        r = 6 + k
        if row["kind"] == "legal":
            plan = legal_years(row["name"], years)
            for y in plan:
                if not ("クレーン性能検査" in row["name"] and y == 2024 and ver >= 2022):   # 抜け：R6 の性能検査
                    ws.cell(row=r, column=ycol[y], value="○")
                if y <= known:
                    ws.cell(row=r, column=ycol[y] + 1, value="○")
            add = "作業環境測定" in row["name"] or "第二種圧力容器" in row["name"]
            ws.cell(row=r, column=ncol, value=row["law"] + ("　※H28追加" if add else ""))
            if add:
                for j in range(2, 9):
                    ws.cell(row=r, column=j).fill = NEWROW
        elif row["kind"] == "eq":
            e = row["e"]
            for y in years:
                s = P[e["id"]].get(y)
                a = act.get(e["id"], {}).get(y)
                if s:
                    ws.cell(row=r, column=ycol[y], value=s)
                if a:
                    ws.cell(row=r, column=ycol[y] + 1, value=a)
                v = acost.get(e["id"], {}).get(y) if y <= known else (cost(e, s) if s else None)
                if v:            # 機器ごとの費用は書かない（設備小計にだけ）
                    sub.setdefault(e["setsubi"], {}).setdefault(y, 0)
                    sub[e["setsubi"]][y] += v
            note = NOTES.get(e["id"], e["duty"]) if ver >= 2022 else (NOTES["0M-006"] .split("。R2")[0] if e["id"] == "0M-006" else e["duty"])
            ws.cell(row=r, column=ncol, value=note)
        else:
            for y in years:
                v = sub.get(row["setsubi"], {}).get(y)
                c = ws.cell(row=r, column=ycol[y], value=v)
                c.number_format = "#,##0"
                ws.merge_cells(start_row=r, start_column=ycol[y], end_row=r, end_column=ycol[y] + 1)
        for y in years:
            if y > known and row["kind"] != "sub":
                ws.cell(row=r, column=ycol[y] + 1).fill = LATE
            for k2 in range(2):
                ws.cell(row=r, column=ycol[y] + k2).border = ROW if row["kind"] == "eq" else BOX
        ws.cell(row=r, column=ncol).alignment = LE
    r = 6 + len(rows)
    ws.cell(row=r, column=2, value="合　計（千円）")
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
    for y in years:
        tot = sum(sub.get(s, {}).get(y, 0) for s in SETSUBI)
        ws.cell(row=r, column=ycol[y], value=tot).number_format = "#,##0"
        ws.merge_cells(start_row=r, start_column=ycol[y], end_row=r, end_column=ycol[y] + 1)
    for j in range(1, ncol + 1):
        ws.cell(row=r, column=j).font = FB
        ws.cell(row=r, column=j).border = BOX
    ws.cell(row=r + 1, column=1, value=("※ 費用の実績は契約額の按分による。└ は付属機器。" if ver < 2022 else
                                        "※ R4改訂以降は未更新（R4〜R6の実績欄は空き）。費用の実績は契約額の按分による。└ は付属機器。")).font = F
    ws.freeze_panes = "I6"
    page_setup(ws, get_column_letter(ncol), r + 1, "4:5")
    fn = f"{FAC_K}_中長期整備計画_{NEWV[ver][0]}_新様式（市i2・停止までの整備計画）.xlsx"
    R = Recs(fn, plan=True)
    harvest(R, ws, rows, 6, ycol, "new", {y: c + 1 for y, c in ycol.items()})
    tot_r = 6 + len(rows)
    R.add(ws.title, tot_r, "total", {"cost": {str(y): ws.cell(row=tot_r, column=c).value for y, c in ycol.items()}, "unit": "千円"})
    # 旧版の費用シートが隠しシートのまま残っている
    hid = wb.create_sheet("費用")
    P_old, _ = plan_old(2015)
    old_cost_sheet(hid, build_rows(legal_skip=("作業環境測定", "第二種圧力容器")), P_old,
                   list(range(2010, 2020)), "平成21年12月策定　平成27年5月改訂（H28 以降 未使用）")
    hid.sheet_state = "hidden"
    set_props(wb, f"{FAC_K} 停止までの整備計画 {NEWV[ver][0]}")
    wb.save(out_path(fn))
    R.save()
    return fn


if __name__ == "__main__":
    print(old_plan())
    print(new_plan())
