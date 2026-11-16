# -*- coding: utf-8 -*-
"""中長期整備計画（成立した版だけ）。
H13（2001）k1「整備年次計画」A3 1 枚（系統単位）／H23（2011）k2「長寿命化計画」（設備別 1 シート・タグ単位）／
H30（2018）k3 が k2 様式で年度送り／R1（2019）〜R8（2026）k3 様式（炉別シート・計画／実績 2 段・総括・法定点検）。
構造の選択（structure-variety）：k3 様式は「系統（炉）ごとにシートが分かれ、最後に集計（総括）」＋「改訂のたびに列が足され旧版の列が隠しで残る」。
k1 様式は「設備の階層を結合と字下げで表す」（号炉→設備→系統の 3 段）。"""
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment
from world import SYS, CLS, wareki, wa_y, KIKAN_FY, DXN_FY, AM_FY
from planner import plan_all, actual
from rec import Rec
from model import WORKS
import xlhelp as X

F = X.font
CACHE = {}


def version_plan(eqs, v):
    if v not in CACHE:
        if v == 2011:
            CACHE[v] = (plan_all(eqs, 2011, list(range(2011, 2021)), "k2"), list(range(2011, 2021)))
        elif v == 2018:
            CACHE[v] = (plan_all(eqs, 2018, list(range(2018, 2028)), "k2"), list(range(2018, 2028)))
        else:
            CACHE[v] = (plan_all(eqs, v, list(range(v, v + 10)), "k3"), list(range(v, v + 10)))
    return CACHE[v][0]


def plan_in_effect(eqs, y):
    v = max(x for x in [2011, 2018] + list(range(2019, 2027)) if x <= y) if y >= 2011 else None
    return version_plan(eqs, v) if v else {}


def _known_last(e, as_of):
    ks = [x for x in e.events if x["fy"] < as_of and x["kind"] in ("OVERHAUL", "REPLACE", "UPGRADE", "REPAIR")]
    if not ks:
        return "（記録なし）" if e.install < 2001 else "%s 設置" % wareki(e.install, True)
    x = ks[-1]
    k = {"OVERHAUL": "整備", "REPLACE": "更新", "UPGRADE": "追設", "REPAIR": "補修"}[x["kind"]]
    return "%s %s" % (wareki(x["fy"], True), k)


def _cycle(e):
    c = CLS[e.cls]
    p = []
    if c.get("oh"):
        p.append("整%d" % c["oh"])
    if c.get("rep"):
        p.append("更%d" % c["rep"])
    if c.get("ins"):
        p.append("点%d" % c["ins"])
    return ("/".join(p) if p else "－"), (c.get("oh") or c.get("rep") or c.get("ins"))


def _remark_k2(e, as_of):
    if e.idle and as_of > 2013:
        return "休止中（H25.10〜）"
    if e.origin == "kikan" and KIKAN_FY[e.unit] < as_of:
        return "基幹改良（%s）で更新" % wareki(KIKAN_FY[e.unit], True)
    if e.origin == "kikan" and as_of == 2011:
        return "基幹改良（%s）予定" % wareki(KIKAN_FY[e.unit], True)
    if e.origin == "dxn":
        return "DXN対策で追設"
    return ""


# ------------------------------------------------------------------ k1 H13 整備年次計画（A3 1 枚）
K1_GROUPS = [("燃焼設備", "ストーカ・給じん", ("IN",), ("stoker", "feeder", "hyd")), ("燃焼設備", "炉本体（耐火物）", ("IN",), ("furnace",)),
             ("燃焼ガス冷却設備", "ボイラ（水管・過熱器）", ("BL",), ("boiler", "sh", "eco")), ("燃焼ガス冷却設備", "スートブロワ・安全弁", ("BL",), ("sb", "sv")),
             ("燃焼ガス冷却設備", "給水ポンプ・弁類", ("BL",), ("pump_m", "pump_s", "mov", "cv")),
             ("排ガス処理設備", "集じん（EP／BF）", ("GT",), ("bf",)), ("排ガス処理設備", "洗煙塔・循環ポンプ", ("GT",), ("scrub", "pump_m")),
             ("排ガス処理設備", "減温塔・触媒（DXN対策）", ("GT",), ("tower", "cat", "hx")),
             ("通風設備", "誘引通風機・押込送風機", ("FD",), ("fan_big", "fan")), ("灰出し設備", "灰押出・コンベヤ", ("AS",), ("conv",)),
             ("電気計装設備", "動力盤・計装", ("EL", "IC"), ("panel", "dcs", "inst", "inst_a", "anal"))]
K1_COMMON = [("受入供給設備", "ごみクレーン", ("RF",), ("crane",)), ("受入供給設備", "計量機・投入扉", ("RF",), ("scale", "door", "hyd")),
             ("余熱利用設備", "蒸気タービン・発電機", ("ST",), ("turbine", "gen")), ("余熱利用設備", "復水器・脱気器・純水", ("ST",), ("acc", "tank", "hx", "fan")),
             ("排水処理設備", "洗煙排水処理", ("WW",), ("tank", "misc", "pump_s")), ("用水・空気設備", "冷却塔・圧縮機", ("UT",), ("hx", "comp", "pump_m")),
             ("電気設備", "受変電・非常用発電機", ("EL",), ("tr", "panel", "egen")), ("計装制御設備", "中央制御盤（DCS 化前）", ("IC",), ("dcs", "panel")),
             ("雑設備", "煙突・クレーン", ("MS",), ("struct", "crane", "hoist"))]


def k1_2001(eqs, path, fname):
    rec = Rec(fname, plan=True)
    years = list(range(2001, 2011))
    wb = Workbook()
    ws = wb.active
    ws.title = "H13整備年次計画"
    X.props(wb, "【架空】くぬぎ台環境センター 整備年次計画（平成13年度）", "くぬぎ台市 清掃事業所 k1", "中長期整備計画", "A3 1 枚・系統単位")
    X.widths(ws, [8, 16, 24, 26, 7] + [6.5] * 10 + [9, 26])
    X.put(ws, 1, 1, "【架空】くぬぎ台環境センター　整備年次計画（平成13年度～平成22年度）", F(size=15, bold=True), X.NOWRAP)
    X.put(ws, 2, 1, "平成13年4月　くぬぎ台市 清掃事業所 環境センター（k1）　○点検　◎整備　●更新・改造　費用：百万円", F(size=9), X.NOWRAP)
    hdr = ["号炉", "設備", "系統・機器", "主な作業", "周期"] + ["H%d" % (y - 1988) for y in years] + ["概算\n(百万円)", "備考"]
    for i, h in enumerate(hdr, 1):
        X.put(ws, 4, i, h, F(bold=True, size=9), X.CENTER, X.fill("E7E6E6"))
    ws.row_dimensions[4].height = 26
    r = 5
    for unit in (1, 2, 3, 0):
        groups = K1_GROUPS if unit else K1_COMMON
        u0 = r
        prev = None
        for setsu, keito, syss, clss in groups:
            mem = [e for e in eqs if e.unit == unit and e.sys in syss and e.cls in clss and e.install <= 2001]
            marks = {}
            cost = 0
            works = set()
            for e in mem:
                p = _k1_plan(e, years)
                for y, (s, c, w) in p.items():
                    if "●" in s or marks.get(y, "") != "●":
                        marks[y] = s if not (marks.get(y) == "●") else "●"
                    cost += c
                    works.add(w)
            if "DXN" in keito or "EP" in keito:
                if unit in (2, 3):
                    marks[DXN_FY[unit]] = "●"
                    cost += 1450000 // 2
                    works.add("DXN 恒久対策（EP→BF、減温塔・触媒追設）")
            if not mem and not marks:
                continue
            X.put(ws, r, 2, setsu if setsu != prev else "", F(size=9), X.LEFT)
            X.put(ws, r, 3, "　" + keito, F(size=9), X.LEFT)
            wk = "・".join(sorted(works))[:30] if works else "点検"
            X.put(ws, r, 4, wk, F(size=8), X.LEFT)
            cyc = {"stoker": 3, "furnace": 2, "boiler": 1, "sb": 3, "bf": 7, "scrub": 12, "fan_big": 4, "crane": 3, "turbine": 4, "egen": 8}.get(clss[0])
            X.put(ws, r, 5, ("%d年" % cyc) if cyc else "", F(size=8), X.CENTER)
            for j, y in enumerate(years):
                if y in marks:
                    X.put(ws, r, 6 + j, marks[y], F(size=10), X.CENTER)
            cm = round(cost / 1000.0)
            X.put(ws, r, 16, cm if cm else None, F(size=8), X.RIGHT)
            rem = ""
            if unit == 1 and "EP" in keito:
                rem = "H12 BF化済"
            if unit in (2, 3) and "EP" in keito:
                rem = "%s DXN対策工事" % ("H13" if unit == 2 else "H14")
            X.put(ws, r, 17, rem, F(size=8), X.LEFT)
            rec.add(ws.title, r, "item", {"path": [("%d号炉" % unit) if unit else "共通", setsu, keito], "work": wk, "cycle": cyc, "cost": cm or None,
                                          "unit": ("%d号炉" % unit) if unit else "共通", "marks": {str(y): s for y, s in marks.items()}, "memo": rem})
            prev = setsu
            r += 1
        X.put(ws, u0, 1, ("%d号炉" % unit) if unit else "共通", F(bold=True, size=10), Alignment(horizontal="center", vertical="center", text_rotation=255 if False else 0))
        ws.merge_cells(start_row=u0, start_column=1, end_row=r - 1, end_column=1)
        X.box(ws, u0, 1, r - 1, 17, inner=X.hair)
    X.put(ws, r, 2, "灰溶融設備", F(size=9), X.LEFT)
    X.put(ws, r, 3, "　電気抵抗式灰溶融炉（追設）", F(size=9), X.LEFT)
    X.put(ws, r, 4, "設置工事", F(size=8), X.LEFT)
    X.put(ws, r, 6 + 2, "●", F(size=10), X.CENTER)
    X.put(ws, r, 16, 1980, F(size=8), X.RIGHT)
    X.put(ws, r, 17, "H15 予定（基本設計中）", F(size=8), X.LEFT)
    rec.add(ws.title, r, "item", {"path": ["共通", "灰溶融設備", "電気抵抗式灰溶融炉（追設）"], "work": "設置工事", "cost": 1980, "unit": "共通",
                                  "marks": {"2003": "●"}, "memo": "H15 予定（基本設計中）"})
    X.box(ws, r, 1, r, 17)
    r += 2
    X.put(ws, r, 1, "※ 定期整備（年 1 回 各炉・全炉停止時）の一般点検は記載しない。法定点検は別に管理（ボイラ・クレーン・受変電）。", F(size=9), X.NOWRAP)
    X.page(ws, "landscape", "A3", 1, 1)
    wb.save(path)
    rec.save()


def _k1_plan(e, years):
    c = CLS[e.cls]
    out = {}
    base = e.install
    if c.get("oh"):
        y = base + c["oh"]
        while y < years[0]:
            y += c["oh"]
        while y <= years[-1]:
            if c["coh"] >= 1000:
                out[y] = ("◎", c["coh"], c.get("oh_name", "分解整備"))
            y += c["oh"]
    if c.get("rep") and c["crep"] >= 5000:
        y = base + c["rep"]
        if years[0] <= y <= years[-1]:
            out[y] = ("●", c["crep"], "更新")
    return out


# ------------------------------------------------------------------ k2 様式（H23 原案・H30 改訂）
def LEGAL_K2(years):
    return [("ボイラー定期事業者検査", "1・2号炉", {y: "★" for y in years if y % 2 == 0}),
            ("ボイラー定期事業者検査", "3号炉", {y: "★" for y in years if y % 2 == 1}),
            ("蒸気タービン定期事業者検査", "共通", {y: "★" for y in years if y % 4 == 0}),
            ("第一種圧力容器性能検査（脱気器）", "共通", {y: "★" for y in years}),
            ("受変電設備年次点検", "共通", {y: "★" for y in years}),
            ("クレーン性能検査（ごみクレーン・タービン室）", "共通", {y: "★" for y in years if y % 2 == 1}),
            ("排ガス測定（ばい煙）", "各炉", {y: "年6回" for y in years}),
            ("ダイオキシン類測定（排ガス）", "各炉", {y: "年1回" for y in years}),
            ("精密機能検査", "全体", {y: "★" for y in years if y % 3 == 1})]


def k2(eqs, v, path, fname):
    rec = Rec(fname, plan=True)
    years = list(range(v, v + 10))
    plan = version_plan(eqs, v)
    wb = Workbook()
    ws = wb.active
    ws.title = "H%d整備計画" % (v - 1988)
    rev = "平成23年度 原案" if v == 2011 else "平成30年度改訂"
    X.props(wb, "【架空】くぬぎ台環境センター 長寿命化計画 整備スケジュール（%s）" % rev, "くぬぎ台市 環境センター %s" % ("k2" if v == 2011 else "k3"),
            "中長期整備計画", "H23 長寿命化計画（k2）の様式")
    ny = len(years)
    cy0, ccost = 10, 10 + ny
    crem = ccost + 1
    X.widths(ws, [5, 11, 12, 5, 9, 24, 6, 9, 9] + [5.2] * ny + [9, 22])
    X.put(ws, 1, 1, "【架空】くぬぎ台環境センター　長寿命化計画　整備スケジュール表", F(size=14, bold=True), X.NOWRAP)
    X.put(ws, 2, 1, "（%s）　くぬぎ台市環境部 環境センター 施設係" % rev, F(size=10), X.NOWRAP)
    X.put(ws, 2, cy0 + 3, "凡例　○点検・測定　△補修　◎整備・部分取替　●更新　★法定　　費用：千円（税抜概算）", F(size=9), X.NOWRAP)
    X.put(ws, 3, crem, "H23.3 原案" + (" ／ H30.5 改訂" if v == 2018 else ""), F(size=8), X.RIGHT)
    for i, h in enumerate(["No", "設備", "系統", "号炉", "タグNo", "機器名称", "設置", "前回整備", "周期(年)"], 1):
        X.put(ws, 4, i, h, F(bold=True), X.CENTER, X.fill("DDEBF7"))
        ws.merge_cells(start_row=4, start_column=i, end_row=5, end_column=i)
    X.put(ws, 4, cy0, "年　度　別　計　画", F(bold=True), X.CENTER, X.fill("DDEBF7"))
    ws.merge_cells(start_row=4, start_column=cy0, end_row=4, end_column=cy0 + ny - 1)
    for j, y in enumerate(years):
        X.put(ws, 5, cy0 + j, "H%d" % (y - 1988), F(bold=True), X.CENTER, X.fill("DDEBF7"))
    if v == 2018:
        ws.cell(5, cy0 + ny - 1).fill = X.fill("FCE4D6")
        ws.cell(5, cy0 + ny - 1).comment = Comment("H30改訂で追加（k3）。H29版の H29 列は削除", "k3")
    X.put(ws, 4, ccost, "概算費用\n(10年計)", F(bold=True), X.CENTER, X.fill("DDEBF7"))
    ws.merge_cells(start_row=4, start_column=ccost, end_row=5, end_column=ccost)
    X.put(ws, 4, crem, "備考", F(bold=True), X.CENTER, X.fill("DDEBF7"))
    ws.merge_cells(start_row=4, start_column=crem, end_row=5, end_column=crem)
    r = 6
    no = 0
    for si, sysc in enumerate(SYS):
        group = [e for e in eqs if e.sys == sysc and e.install <= v]
        if not group:
            continue
        group.sort(key=lambda e: (e.unit if e.unit else 9, e.tag))
        setsu, keito = SYS[sysc]
        if v < 2013 and sysc == "AM":
            setsu = "灰溶融設備"
        X.put(ws, r, 2, "【%s】%s" % (setsu, keito), F(bold=True), X.LEFT, X.fill("F2F2F2"))
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
        for c in range(1, crem + 1):
            ws.cell(r, c).fill = X.fill("F2F2F2")
        rec.add(ws.title, r, "parent", {"path": [setsu]})
        r += 1
        first = r
        ysum = {y: 0 for y in years}
        prev_unit = None
        for e in group:
            no += 1
            if v == 2018:
                e.old_no = no
            X.put(ws, r, 1, no, F(size=8), X.CENTER)
            if si >= 2:
                X.put(ws, r, 2, setsu if r == first else "", F(size=8), X.LEFT)
            X.put(ws, r, 3, keito if r == first else "〃", F(size=8), X.CENTER if r != first else X.LEFT)
            un = {0: "共通", 1: "1", 2: "2", 3: "3"}[e.unit]
            X.put(ws, r, 4, un if un != prev_unit else "〃", F(size=8), X.CENTER)
            prev_unit = un
            X.put(ws, r, 5, e.tag, F(size=8), X.CENTER)
            X.put(ws, r, 6, e.name, F(size=8), X.LEFT)
            X.put(ws, r, 7, wareki(e.install, True), F(size=8), X.CENTER)
            last = _known_last(e, v)
            X.put(ws, r, 8, last, F(size=8), X.CENTER)
            ct, cn = _cycle(e)
            X.put(ws, r, 9, ct, F(size=7), X.CENTER)
            tot = 0
            marks = {}
            for j, y in enumerate(years):
                items = plan.get(e.tag, {}).get(y, [])
                if items:
                    sym = "".join(s.replace("◇", "△").replace("(●)", "●") for s, _, _ in items)
                    X.put(ws, r, cy0 + j, sym, F(size=9), X.CENTER)
                    marks[str(y)] = sym
                    cost = sum(cst for _, cst, _ in items)
                    tot += cost
                    ysum[y] += cost
            if tot:
                X.put(ws, r, ccost, tot, F(size=8), X.RIGHT, nf="#,##0")
            rem = _remark_k2(e, v)
            if e.idle and v > 2013:
                for c in range(1, crem + 1):
                    ws.cell(r, c).fill = X.fill("D9D9D9")
            X.put(ws, r, crem, rem, F(size=7), X.LEFT)
            cm = {}
            if v == 2018 and e.name == "過熱器" and e.unit == 2:
                ws.cell(r, cy0 + 5).comment = Comment("H35 取替は H30 測定次第で見直す（k3）", "k3")
                cm[ws.cell(r, cy0 + 5).coordinate] = "H35 取替は H30 測定次第で見直す（k3）"
            rec.add(ws.title, r, "item", {"path": [setsu, keito, e.name], "code": e.tag, "cycle": cn, "cycleText": ct, "cost": tot or None,
                                          "unit": un if un != "共通" else "共通", "marks": marks, "memo": rem, "comments": cm,
                                          "設置": wareki(e.install, True), "前回整備": last})
            r += 1
        if si < 2:
            ws.merge_cells(start_row=first, start_column=2, end_row=r - 1, end_column=2)
            X.put(ws, first, 2, setsu, F(size=9), Alignment(vertical="top", wrap_text=True, horizontal="center"))
        X.put(ws, r, 6, "小計（%s）" % setsu, F(bold=True, size=8), X.RIGHT)
        vals = {}
        for j, y in enumerate(years):
            val = ysum[y]
            if v == 2018 and sysc == "GT" and y == 2021:
                val -= 1200
            if v == 2011 and sysc == "BL" and y == 2016:
                val += 3000
            if val:
                X.put(ws, r, cy0 + j, val, F(size=7, bold=True), X.RIGHT, nf="#,##0")
                vals[str(y)] = val
        X.put(ws, r, ccost, sum(ysum.values()), F(size=7, bold=True), X.RIGHT, nf="#,##0")
        for c in range(1, crem + 1):
            ws.cell(r, c).fill = X.fill("FFF2CC")
        rec.add(ws.title, r, "subtotal", {"path": [setsu], "values": vals, "cost": sum(ysum.values())})
        X.box(ws, first - 1, 1, r, crem, side=X.thin, inner=X.hair)
        r += 1
        if sysc == "AM" and v > 2013:
            r += 2
    r += 1
    X.put(ws, r, 2, "【法定点検等】", F(bold=True), X.LEFT, X.fill("E2EFDA"))
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
    r += 1
    f0 = r
    for name, tgt, marks in LEGAL_K2(years):
        X.put(ws, r, 3, tgt, F(size=8), X.CENTER)
        X.put(ws, r, 6, name, F(size=8), X.LEFT)
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
        for j, y in enumerate(years):
            if y in marks:
                X.put(ws, r, cy0 + j, marks[y], F(size=7), X.CENTER)
        rec.add(ws.title, r, "item", {"path": ["法定点検等", name], "unit": tgt, "marks": {str(y): s for y, s in marks.items()}})
        r += 1
    X.box(ws, f0 - 1, 1, r - 1, crem, side=X.thin, inner=X.hair)
    X.box(ws, 4, 1, 5, crem, side=X.med, inner=X.thin)
    ws.freeze_panes = "J6"
    X.page(ws, "landscape", "A3", 1, 0, area="A1:%s%d" % (X.get_column_letter(crem), r), titles="4:5")
    ws.oddFooter.center.text = "&P / &N"
    wb.save(path)
    rec.save()


# ------------------------------------------------------------------ k3 様式（R1〜）
def health4(e, as_of):
    if e.idle:
        return "－"
    if "health4" in e.aliases and as_of >= 2025:
        return str(e.aliases["health4"])
    c = CLS[e.cls]
    base = e.install
    for x in e.events:
        if x["fy"] < as_of and x["kind"] in ("REPLACE", "UPGRADE"):
            base = max(base, x["fy"])
    r = (as_of - base) / float(c.get("rep") or 30)
    v = 4 if r < 0.5 else 3 if r < 0.85 else 2
    v = min(4, max(1, v + (1 if e.health > 0 else -1 if e.health < 0 and v > 2 else 0)))
    return str(v)


LEGAL_K3 = [
    ("ボイラー定期事業者検査", "電気事業法 第55条", "2年", "1号炉・2号炉", lambda y: "★" if y % 2 == 0 else ""),
    ("ボイラー定期事業者検査", "電気事業法 第55条", "2年", "3号炉", lambda y: "★" if y % 2 == 1 else ""),
    ("蒸気タービン定期事業者検査", "電気事業法 第55条", "4年", "0ST-001", lambda y: "★" if y % 4 == 0 else ""),
    ("第一種圧力容器 性能検査", "ボイラー及び圧力容器安全規則", "1年", "脱気器 0ST-031", lambda y: "★"),
    ("受変電設備 年次点検（停電）", "保安規程", "1年", "全炉停止時", lambda y: "★"),
    ("受変電設備 月次点検", "保安規程", "毎月", "電気主任技術者", lambda y: "12回"),
    ("非常用発電機 点検", "消防法・保安規程", "6月／1年", "0EL-023", lambda y: "2回"),
    ("ごみクレーン 性能検査", "クレーン等安全規則", "2年", "0RF-012・013", lambda y: "★" if y % 2 == 1 else ""),
    ("クレーン 年次自主検査", "クレーン等安全規則 第34条", "1年", "ごみ・灰・タービン室・ホイスト", lambda y: "★"),
    ("タービン室天井クレーン 性能検査", "クレーン等安全規則", "2年", "0MS-002", lambda y: "★" if y % 2 == 0 else ""),
    ("ばい煙量等測定（排ガス）", "大気汚染防止法", "2月", "各炉", lambda y: "6回"),
    ("ダイオキシン類測定（排ガス・飛灰）", "DXN 特措法", "1年", "各炉", lambda y: "1回"),
    ("作業環境測定（ダイオキシン類）", "安衛則 第592条の2", "6月", "炉室・灰出し", lambda y: "2回"),
    ("計量機 定期検査", "計量法", "2年", "0RF-001・002", lambda y: "★" if y % 2 == 1 else ""),
    ("消防用設備等点検", "消防法 第17条の3の3", "6月", "全館", lambda y: "2回"),
    ("精密機能検査", "廃掃法施行規則 第4条の5", "3年", "施設全体", lambda y: "★" if y % 3 == 1 else ""),
]

REM_K3 = {
    (2, "誘引通風機"): "R6 羽根車肉盛・軸受交換で振動 1.9mm/s。肉盛は 4〜5 年もつ見込み→R11 羽根車更新",
    (2, "バグフィルタ"): "R6 ろ布全数交換（市支給 288 本）。強度試験を 2 年ごと",
    (2, "過熱器"): "1次 最小 3.0mm（R7）、減肉 0.1〜0.2mm/年、管理値 2.6mm。管材の手配に 10 か月→R9 全数取替",
    (1, "燃焼ストーカ"): "火格子 摩耗 最大 9mm（R6）。R8 44 枚交換済。フレームは竣工時のまま→更新は方針待ち",
    (0, "ごみクレーン 1号"): "常用。ワイヤは 3 年。制御盤（0RF-014）と同時に更新を検討",
    (0, "ごみクレーン 2号"): "予備。1号とワイヤ交換年をずらす",
    (0, "蒸気タービン"): "R6 開放：動翼先端エロージョン軽微→R10 開放で再確認",
    (3, "湿式洗煙塔"): "R7 ライニング全面更新済。2 年ごと内面点検（スパーク試験）",
    (2, "洗煙循環ポンプ A"): "R6 羽根車・ライナ交換。腐食摩耗のため 4 年で分解",
    (0, "非常用発電機"): "部品供給終了（R7.9 通知）。大規模改修の方針に関わらず R9 更新",
    (1, "ボイラ給水ポンプ A"): "R6 メカシール漏れ（直営交換）、R7 分解整備",
    (1, "主灰押出装置"): "シュート側板 穴あき（R6 直営 当て板）。R7 予算は見送り→R8 シュート更新済",
    (0, "灰溶融炉本体"): "H25.10 休止・残置。直営で 3 年ごと目視（R6 実施、さび進行なし）。撤去は方針待ち",
}

HIST = [("H23.3", "長寿命化計画（原案 k2）。設備別 1 シート、H23〜H32", 2011),
        ("H30.5", "年度送り（k3 引継ぎ）。H39 列を追加", 2018),
        ("R1.5", "様式改訂（k3）：炉別シート化（定修契約の単位に合わせる）、計画／実績 2 段、重要度・健全度欄、法定点検を別シート、総括表", 2019),
        ("R4.4", "台帳の 1 シート統合（R4）に合わせてタグ No を台帳と一致させた", 2022),
        ("R6.4", "(●) 区分を追加（大規模改修の方針検討開始）", 2024)]


def _yh(y):
    return "%s\n%d" % (wareki(y, True), y)


def k3(eqs, v, path, fname):
    rec = Rec(fname, plan=True)
    past = [v - 3, v - 2, v - 1]
    years = list(range(v, v + 10))
    allyears = past + years
    plan = version_plan(eqs, v)
    pplan = {y: plan_in_effect(eqs, y) for y in past}
    wb = Workbook()
    made = "R1.5.20" if v == 2019 else "R%d.4.20" % (v - 2018)
    X.props(wb, "【架空】くぬぎ台環境センター 中長期整備計画（%s版）" % wareki(v).replace("年度", "年度"), "くぬぎ台市 環境センター k3",
            "中長期整備計画", "R1 様式（炉別シート・計画/実績2段）。%s 作成" % made)
    ws0 = wb.active
    ws0.title = "総括"
    sheets = [("1号炉", lambda e: e.unit == 1), ("2号炉", lambda e: e.unit == 2), ("3号炉", lambda e: e.unit == 3),
              ("共通", lambda e: e.unit == 0 and not e.idle), ("灰溶融（休止）", lambda e: e.idle)]
    totals, pend = {}, {}
    for sname, pred in sheets:
        ws = wb.create_sheet(sname)
        t, p = _k3_sheet(ws, sname, [e for e in eqs if pred(e)], plan, pplan, past, years, v, rec)
        totals[sname], pend[sname] = t, p
    _k3_summary(ws0, totals, pend, past, years, v, rec)
    _k3_legal(wb.create_sheet("法定点検"), allyears, v, rec)
    _k3_legend(wb.create_sheet("凡例・改訂履歴"), v)
    wb.save(path)
    rec.save()


def _k3_sheet(ws, sname, group, plan, pplan, past, years, v, rec):
    allyears = past + years
    c0 = 10
    cr = c0 + len(allyears)
    X.widths(ws, [10, 12, 5, 9, 24, 5, 5, 5, 4] + [5.4] * len(allyears) + [30])
    ws.column_dimensions["C"].hidden = True
    X.put(ws, 1, 1, "【架空】くぬぎ台環境センター　中長期整備計画（%s版）　%s" % (wareki(v), sname), F(size=14, bold=True), X.NOWRAP)
    upd = "／R8.10.30 上期実績反映" if v == 2026 else ""
    X.put(ws, 2, 1, "くぬぎ台市環境部 環境センター 施設係　　%s 作成%s（k3）" % ("R1.5.20" if v == 2019 else "R%d.4.20" % (v - 2018), upd), F(size=9), X.NOWRAP)
    leg = "○点検・測定　◇補修　◎分解整備・部分取替　●更新" + ("　(●)大規模改修の方針決定後に判断" if v >= 2024 else "") + "　実績行は網掛け"
    X.put(ws, 2, c0 + 4, leg, F(size=8), X.NOWRAP)
    for i, h in enumerate(["設備", "系統", "旧No", "タグNo", "機器名称", "区分", "重要度", "健全度", "計/実"], 1):
        X.put(ws, 4, i, h, F(bold=True, size=8), X.CENTER, X.fill("D9E1F2"))
        ws.merge_cells(start_row=4, start_column=i, end_row=5, end_column=i)
    X.put(ws, 4, c0, "実　績（当初計画）", F(bold=True, size=8), X.CENTER, X.fill("EDEDED"))
    ws.merge_cells(start_row=4, start_column=c0, end_row=4, end_column=c0 + len(past) - 1)
    X.put(ws, 4, c0 + len(past), "計　画", F(bold=True, size=8), X.CENTER, X.fill("D9E1F2"))
    ws.merge_cells(start_row=4, start_column=c0 + len(past), end_row=4, end_column=cr - 1)
    for j, y in enumerate(allyears):
        X.put(ws, 5, c0 + j, _yh(y), F(bold=True, size=8), X.CENTER, X.fill("EDEDED" if y in past else ("FFE699" if y == v else "D9E1F2")))
    X.put(ws, 4, cr, "備考（状態・根拠）", F(bold=True, size=8), X.CENTER, X.fill("D9E1F2"))
    ws.merge_cells(start_row=4, start_column=cr, end_row=5, end_column=cr)
    ws.row_dimensions[5].height = 24
    r = 6
    tot = {y: 0 for y in allyears}
    pen = {y: 0 for y in allyears}
    for sysc in SYS:
        g = [e for e in group if e.sys == sysc]
        if not g:
            continue
        setsu, keito = SYS[sysc]
        X.put(ws, r, 1, "%s　（%s）" % (setsu, keito), F(bold=True, size=9), X.LEFT, X.fill("F2F2F2"))
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=cr)
        rec.add(sname, r, "parent", {"path": [sname, setsu]})
        r += 1
        g0 = r
        sub = {y: 0 for y in allyears}
        for e in g:
            X.put(ws, r, 2, keito, F(size=7), X.CENTER)
            X.put(ws, r, 3, e.old_no, F(size=7), X.CENTER)
            X.put(ws, r, 4, e.tag, F(size=8), X.CENTER)
            X.put(ws, r, 5, e.name, F(size=8), X.LEFT)
            kb = "休止" if e.idle else (e.duty or "－")
            X.put(ws, r, 6, kb, F(size=7), X.CENTER)
            imp = "－" if e.idle else e.imp
            X.put(ws, r, 7, imp, F(size=8), X.CENTER)
            h4 = health4(e, v)
            X.put(ws, r, 8, h4, F(size=8), X.CENTER)
            X.put(ws, r, 9, "計", F(size=7), X.CENTER)
            X.put(ws, r + 1, 9, "実", F(size=7), X.CENTER, X.fill("EDEDED"))
            pm, am = {}, {}
            wl = []
            for j, y in enumerate(allyears):
                src = pplan[y].get(e.tag, {}) if y in past else plan.get(e.tag, {})
                items = src.get(y, [])
                if items:
                    sym = "".join(s for s, _, _ in items)
                    X.put(ws, r, c0 + j, sym, F(size=8), X.CENTER)
                    pm[str(y)] = sym
                    for s, cost, _ in items:
                        if s == "(●)":
                            pen[y] += cost
                        else:
                            sub[y] += cost
                ws.cell(r + 1, c0 + j).fill = X.fill("EDEDED")
                if y in past or y == v:
                    a = [s for s, _ in actual(e, y) if not (y == v and v == 2026 and e.unit != 1)]
                    if y == v and v != 2026:
                        a = []
                    if a:
                        sym = "".join(dict.fromkeys(a))
                        X.put(ws, r + 1, c0 + j, sym, F(size=8, bold=True), X.CENTER, X.fill("BFBFBF"))
                        am[str(y)] = sym
                        wl += [w["id"] for w in WORKS.values() if w["tag"] == e.tag and w["fy"] == y]
            for c in range(1, 9):
                ws.merge_cells(start_row=r, start_column=c, end_row=r + 1, end_column=c)
            rem = REM_K3.get((e.unit, e.name)) if v == 2026 else None
            if rem is None:
                rem = "休止（H25.10〜）・残置" if e.idle else _remark_k2(e, v)
            X.put(ws, r, cr, rem, F(size=7), X.LEFT)
            ws.merge_cells(start_row=r, start_column=cr, end_row=r + 1, end_column=cr)
            ws.row_dimensions[r].height = 11.5
            ws.row_dimensions[r + 1].height = 11.5
            base = {"path": [sname, setsu, keito, e.name], "code": e.tag, "unit": sname if sname != "灰溶融（休止）" else "共通",
                    "区分": kb, "重要度": imp, "健全度": h4, "memo": rem, "旧No": e.old_no}
            rec.add(sname, r, "item", dict(base, **{"計実": "計", "marks": pm}))
            rec.add(sname, r + 1, "item", dict(base, **{"計実": "実", "marks": am}), works=wl, r_start=r)
            r += 2
        X.put(ws, g0, 1, setsu, F(size=8), Alignment(vertical="top", horizontal="center", wrap_text=True))
        ws.merge_cells(start_row=g0, start_column=1, end_row=r - 1, end_column=1)
        X.box(ws, g0, 1, r - 1, cr, side=X.thin, inner=X.hair)
        X.put(ws, r, 5, "小計（百万円）", F(size=8, bold=True), X.RIGHT, X.fill("FFF2CC"))
        vals = {}
        for j, y in enumerate(allyears):
            if sub[y]:
                X.put(ws, r, c0 + j, round(sub[y] / 1000, 1), F(size=7, bold=True), X.RIGHT, X.fill("FFF2CC"), nf="0.0")
                vals[str(y)] = round(sub[y] / 1000, 1)
            else:
                ws.cell(r, c0 + j).fill = X.fill("FFF2CC")
            tot[y] += sub[y]
        rec.add(sname, r, "subtotal", {"path": [sname, setsu], "values": vals, "unit": "百万円"})
        X.box(ws, r, 1, r, cr)
        r += 1
    if sname.startswith("灰溶融"):
        X.put(ws, r + 1, 1, "※ 灰溶融設備は H25.10 から休止（電力費）。再稼働の予定なし。撤去は 2 回目の大規模改修（又は建替え）の方針と併せて判断。"
              "休止中は直営で 3 年ごとに目視点検。台帳上は残置。", F(size=9), X.LEFT)
        ws.merge_cells(start_row=r + 1, start_column=1, end_row=r + 2, end_column=cr)
        r += 3
    else:
        X.put(ws, r, 5, "合計（百万円）", F(size=9, bold=True), X.RIGHT, X.fill("F8CBAD"))
        vals = {}
        for j, y in enumerate(allyears):
            X.put(ws, r, c0 + j, round(tot[y] / 1000, 1) if tot[y] else None, F(size=7, bold=True), X.RIGHT, X.fill("F8CBAD"), nf="0.0")
            if tot[y]:
                vals[str(y)] = round(tot[y] / 1000, 1)
        rec.add(sname, r, "total", {"path": [sname], "values": vals, "unit": "百万円"})
        if v >= 2024:
            X.put(ws, r + 1, 5, "(●) 参考（百万円）", F(size=8), X.RIGHT)
            pv = {}
            for j, y in enumerate(allyears):
                if pen[y]:
                    X.put(ws, r + 1, c0 + j, "(%.1f)" % (pen[y] / 1000), F(size=7), X.RIGHT)
                    pv[str(y)] = "(%.1f)" % (pen[y] / 1000)
            if pv:
                rec.add(sname, r + 1, "subtotal", {"path": [sname, "(●) 参考"], "values": pv, "unit": "百万円"})
        X.box(ws, r, 1, r + 1, cr)
        r += 2
    X.box(ws, 4, 1, 5, cr, side=X.med)
    ws.freeze_panes = "%s6" % X.get_column_letter(c0)
    X.page(ws, "landscape", "A3", 1, 0, area="A1:%s%d" % (X.get_column_letter(cr), r), titles="4:5", margins=(0.3, 0.3, 0.4, 0.4))
    ws.oddFooter.center.text = "%s　&P / &N" % sname
    return tot, pen


def _k3_summary(ws, totals, pend, past, years, v, rec):
    allyears = past + years
    X.widths(ws, [16] + [7.5] * len(allyears) + [10])
    X.put(ws, 1, 1, "【架空】くぬぎ台環境センター　中長期整備計画（%s版）　総括表" % wareki(v), F(size=14, bold=True), X.NOWRAP)
    X.put(ws, 2, 1, "単位：百万円（税抜概算）。%s〜%s は当時の計画額（当初）。" % (wareki(past[0], True), wareki(past[-1], True)) +
          ("(●) は 2 回目の大規模改修（又は建替え）の方針決定後に判断するもので、合計に含めない。" if v >= 2024 else ""), F(size=9), X.NOWRAP)
    X.put(ws, 4, 1, "区分", F(bold=True), X.CENTER, X.fill("D9E1F2"))
    for j, y in enumerate(allyears):
        X.put(ws, 4, 2 + j, _yh(y), F(bold=True, size=9), X.CENTER, X.fill("EDEDED" if y in past else "D9E1F2"))
    lab = "%s〜%s計" % (wareki(years[0], True), wareki(years[-1], True))
    X.put(ws, 4, 2 + len(allyears), lab, F(bold=True, size=9), X.CENTER, X.fill("D9E1F2"))
    ws.row_dimensions[4].height = 28
    r = 5
    grand = {y: 0 for y in allyears}
    for s in ("1号炉", "2号炉", "3号炉", "共通"):
        X.put(ws, r, 1, s, F(), X.CENTER)
        t = 0
        vals = {}
        for j, y in enumerate(allyears):
            val = totals[s][y]
            grand[y] += val
            if y in years:
                t += val
            X.put(ws, r, 2 + j, round(val / 1000, 1), F(size=9), X.RIGHT, nf="0.0")
            vals[str(y)] = round(val / 1000, 1)
        X.put(ws, r, 2 + len(allyears), round(t / 1000, 1), F(size=9, bold=True), X.RIGHT, nf="#,##0.0")
        rec.add("総括", r, "subtotal", {"path": [s], "values": vals, "cost": round(t / 1000, 1), "unit": "百万円"})
        r += 1
    X.put(ws, r, 1, "合計", F(bold=True), X.CENTER, X.fill("F8CBAD"))
    vals = {}
    for j, y in enumerate(allyears):
        val = round(grand[y] / 1000, 1)
        if v == 2022 and y == 2027:
            val = round(val + 12.0, 1)    # わざと：総括の手直し漏れ（R4 版、R9 列が炉別シートの合計より 12.0 多い）
        X.put(ws, r, 2 + j, val, F(size=9, bold=True), X.RIGHT, X.fill("F8CBAD"), nf="0.0")
        vals[str(y)] = val
    gt = round(sum(grand[y] for y in years) / 1000, 1)
    X.put(ws, r, 2 + len(allyears), gt, F(size=9, bold=True), X.RIGHT, X.fill("F8CBAD"), nf="#,##0.0")
    rec.add("総括", r, "total", {"path": ["合計"], "values": vals, "cost": gt, "unit": "百万円"})
    r += 1
    if v >= 2024:
        X.put(ws, r, 1, "(●) 参考", F(size=9), X.CENTER)
        pv = {}
        for j, y in enumerate(allyears):
            val = sum(pend[s][y] for s in pend)
            if val:
                X.put(ws, r, 2 + j, "(%.1f)" % (val / 1000), F(size=8), X.RIGHT)
                pv[str(y)] = "(%.1f)" % (val / 1000)
        rec.add("総括", r, "subtotal", {"path": ["(●) 参考"], "values": pv, "unit": "百万円"})
    X.box(ws, 4, 1, r, 2 + len(allyears))
    r += 2
    notes = ["1. 定期整備（炉別 3 件・全炉停止時共通 1 件）の一般点検は各仕様書による。本表には特別な点検・測定、分解整備、更新を載せる。",
             "2. 2 回目の大規模改修か建替えかは検討中（%s 時点）。" % wareki(v, True) + ("R11 以降の大きな更新は (●) とし、方針決定後に計画を見直す。" if v >= 2024 else ""),
             "3. 灰溶融設備（41 台）は休止のため計画なし。別シート参照。", "4. 法定点検は「法定点検」シートにまとめた（R1 改訂で本表から分離）。"]
    if v == 2026:
        notes.insert(2, "3'. 非常用発電機（R9）、2号炉 1次過熱器（R9）は方針に関わらず実施する。")
    for t in notes:
        X.put(ws, r, 1, t, F(size=9), X.NOWRAP)
        r += 1
    X.page(ws, "landscape", "A3", 1, 1)


def _k3_legal(ws, allyears, v, rec):
    X.widths(ws, [30, 24, 8, 22] + [6] * len(allyears) + [12])
    X.put(ws, 1, 1, "【架空】くぬぎ台環境センター　法定点検等 年度計画（%s版）" % wareki(v), F(size=13, bold=True), X.NOWRAP)
    X.put(ws, 2, 1, "★：実施年度　回数：年間の実施回数　（R1 改訂で計画表から分離。H30 版に無かった作業環境測定・計量機を追加）", F(size=9), X.NOWRAP)
    for i, h in enumerate(["項目", "根拠", "周期", "対象"], 1):
        X.put(ws, 4, i, h, F(bold=True), X.CENTER, X.fill("E2EFDA"))
    for j, y in enumerate(allyears):
        X.put(ws, 4, 5 + j, _yh(y), F(bold=True, size=8), X.CENTER, X.fill("E2EFDA"))
    X.put(ws, 4, 5 + len(allyears), "委託先等", F(bold=True), X.CENTER, X.fill("E2EFDA"))
    ws.row_dimensions[4].height = 26
    r = 5
    for name, law, cyc, tgt, fn in LEGAL_K3:
        if v < 2020 and name.startswith("計量機"):
            continue     # わざと：R1 版は計量機の行を足し忘れ（R2 版で追加）
        X.put(ws, r, 1, name, F(size=9), X.LEFT)
        X.put(ws, r, 2, law, F(size=8), X.LEFT)
        X.put(ws, r, 3, cyc, F(size=8), X.CENTER)
        X.put(ws, r, 4, tgt, F(size=8), X.LEFT)
        mk = {}
        for j, y in enumerate(allyears):
            val = fn(y)
            if val:
                X.put(ws, r, 5 + j, val, F(size=8), X.CENTER, X.fill("EDEDED") if j < 3 else None)
                mk[str(y)] = val
        rec.add("法定点検", r, "item", {"path": ["法定点検", name], "work": name, "cycleText": cyc, "unit": tgt, "marks": mk, "根拠": law})
        r += 1
    X.box(ws, 4, 1, r - 1, 5 + len(allyears))
    X.put(ws, r + 1, 1, "※ 電気事業法の定期事業者検査の周期（ボイラー 2 年・タービン 4 年）は保安規程の定めによる。要確認（電気保安担当）。", F(size=8), X.NOWRAP)
    X.page(ws, "landscape", "A3", 1, 1)


def _k3_legend(ws, v):
    X.widths(ws, [14, 90])
    X.put(ws, 1, 1, "【架空】くぬぎ台環境センター　中長期整備計画　凡例・改訂履歴", F(size=13, bold=True), X.NOWRAP)
    rows = [("○", "点検・測定（定修の一般点検は書かない）"), ("◇", "補修（H30 版までの △）"), ("◎", "分解整備・部分取替"), ("●", "更新")]
    if v >= 2024:
        rows.append(("(●)", "2 回目大規模改修（又は建替え）の方針決定後に判断する更新。費用は参考"))
    rows += [("計／実", "上段：計画（過去年度は当初計画のまま残す）、下段：実績（網掛け）"),
             ("重要度", "A：炉停止・環境・安全に直結、予備なし　B：予備あり又は短期停止で対応可　C：その他"),
             ("健全度", "4：問題なし　3：軽微な劣化　2：劣化進行（計画的対応）　1：要早期対応 又は 休止"), ("", ""), ("改訂履歴", "")]
    n0 = len(rows)
    rows += [(a, b) for a, b, y in HIST if y <= v]
    rows.append(("R1.5.20" if v == 2019 else "R%d.4.20" % (v - 2018), "%s版 作成" % wareki(v)))
    if v == 2026:
        rows.append(("R8.10.30", "上期実績を反映（1号炉 定修 9 月）"))
    for i, (a, b) in enumerate(rows, 3):
        X.put(ws, i, 1, a, F(size=10, bold=(a == "改訂履歴")), X.CENTER)
        X.put(ws, i, 2, b, F(size=10), X.LEFT)
    X.box(ws, 3, 1, 3 + n0 - 3, 2)
    X.box(ws, 3 + n0, 1, 3 + len(rows) - 1, 2)
    X.page(ws, "portrait", "A4", 1, 1)
