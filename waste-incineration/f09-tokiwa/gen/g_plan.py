# -*- coding: utf-8 -*-
"""中長期整備計画（全版）
 SPC 長期修繕計画：2009–2012 初版様式（sv1）、2013–2016 sv1 2013 様式、2017–2026 sv2 様式（sv1 版のコピー）
 組合 中長期整備計画：令和6年度策定（2024.11）、令和7年度改訂（2026.3）。令和8年度改訂は未作成（2027.3 予定）
"""
import datetime as dt
from openpyxl import Workbook
from openpyxl.comments import Comment
import world as W, contracts as C
import plan_gen as PG
from plan_gen import GROUP_ORDER, sort_key_block, REP_BY_TAG, KIKAN, U_GROUPS, u_name, cyc_of, nendo_row, actual_u
from xlh import *
from rec import Rec

def ver_key(V):
    return "spc2016" if V <= 2016 else ("spc2017" if V <= 2023 else "spc2024")

def marks(e, V, y0, y1):
    k = ver_key(V)
    if k == "spc2024":
        m = PG.marks_for(e, "u2024", y0, y1)
        if e.tag in KIKAN:
            fy = KIKAN[e.tag][1]
            if m.get(fy) in ("R", "U"): m[fy] = "K"
        return m
    return PG.marks_for(e, k, y0, y1)

def legal(V, y0, y1, union=False):
    rows = PG.legal_rows("u2024" if union else ("spc2016" if V <= 2016 else "spc2017"), y0, y1)
    out = []
    for name, basis, cyc, m in rows:
        if "精密" in name:
            if V <= 2014: pz = [2011, 2014, 2017, 2020, 2023]
            elif V <= 2016: pz = [2011, 2014, 2018, 2021]           # 2015 版から 2017→2018（周期の超え）
            elif V <= 2023: pz = [2018, 2021, 2024, 2027, 2030, 2033]
            else: pz = [2024, 2027, 2030, 2033]
            m = {y: "L" for y in pz if y0 <= y <= y1}
        out.append((name, basis, cyc, m))
    return out

def note_sv1(e, V):
    k = REP_BY_TAG.get(e.tag)
    if not k: return ""
    d = {"R1": "出滓口・側壁の減りに注意" if V >= 2013 else "", "R3": "H27 減肉1本 肉盛済（要監視）" if V >= 2016 else "",
         "R6": "2年毎 A/B交互", "R8": "ろ布4年毎", "R11": "DCS更新（メーカー提案 2019〜）" if V >= 2015 else "", "R10": "開放 4年毎"}
    return d.get(k, "")

def note_sv2(e, V, mall):
    k = REP_BY_TAG.get(e.tag)
    if k:
        d = {"R1": "H27 部分補修。全面張替は第2期で要検討" if V <= 2019 else "H27・R1 部分補修。全面張替は第2期で要検討",
             "R2": "2号炉の2年遅れ", "R3": "H27 減肉→肉盛。材質変更検討" if V <= 2019 else "R1 部分取替。材質変更検討",
             "R4": "2号炉に準じる", "R5": "羽根車摩耗 監視", "R6": "A/B 2年毎交互", "R7": "2号機（予備）併用で刃の持ち改善",
             "R8": "ろ布4年毎", "R9": "ワイヤ3年毎", "R10": "定期事業者検査 開放", "R11": "DCS更新 2019〜2020（別途工事）" if V <= 2019 else "2020.3 DCS 更新済",
             "R12": "チェーン4年毎"}
        n = d[k]
        if V >= 2024 and e.tag in KIKAN: n = "基幹改良（組合発注）" + (" R7 完了" if V >= 2026 and KIKAN[e.tag][1] == 2025 else "")
        return n
    if V >= 2024 and e.tag in KIKAN:
        return "基幹改良（組合発注）" + (" R7 完了" if V >= 2026 and KIKAN[e.tag][1] == 2025 else "")
    if e.tag in KIKAN and any(v == "R" for v in mall.values()):
        return "第2期 大規模改修の候補"
    return ""

# ============================================================== sv1（初版 2009–2012／2013 様式 2013–2016）
def sv1(V, path, fname):
    rec = Rec(fname, plan=True)
    first = V <= 2012
    wb = Workbook(); ws = wb.active; ws.title = "長期修繕計画"
    Y0, Y1 = 2009, 2023
    yc0 = 7 if first else 8
    ncol = yc0 + (Y1 - Y0 + 1)
    SYM = {"I": "○", "O": "◎", "P": "◎", "R": "●", "U": "●", "L": "○"}
    md = dt.date(V, 4, 15) if V > 2009 else dt.date(2009, 6, 30)
    ttl = f"{W.FAC}　長期修繕計画表（{V}年度{'版' if first else '改訂'}）"
    merge(ws, 1, 1, 1, ncol, ttl, f(16, True), al("center"), None)
    merge(ws, 2, 1, 2, ncol, f"{W.SPC}　保全グループ　作成：sv1　{'作成日' if V == 2009 else '改訂日'} {md.year}/{md.month}/{md.day}　（第1期 長期包括運営委託 2009〜2023年度）", f(9), al("right"), None)
    put(ws, 3, 1, "凡例", f(9, True))
    merge(ws, 3, 2, 3, 6, "○：点検　◎：整備（分解・補修）　●：更新", f(9), al("left"), None)
    hdr = ["No.", "設備", "系統", "機器名称", "タグNo.", "区分"] + ([] if first else ["周期\n(年)"])
    for i, h in enumerate(hdr):
        merge(ws, 5, i + 1, 6, i + 1, h, f(9, True), al("center", wrap=True), B_ALL, fill("DDEBF7"))
    merge(ws, 5, yc0, 5, ncol - 1, "年　　度（西暦）", f(9, True), al("center"), B_ALL, fill("DDEBF7"))
    for y in range(Y0, Y1 + 1):
        put(ws, 6, yc0 + y - Y0, f"'{str(y)[2:]}", f(9, True), al("center"), B_ALL, fill("DDEBF7"))
    merge(ws, 5, ncol, 6, ncol, "備　考", f(9, True), al("center"), B_ALL, fill("DDEBF7"))
    if not first:
        nendo_row(ws, 7, yc0, Y0, Y1, 7)
    widths(ws, {"A": 5, "B": 11, "C": 11, "D": 26, "E": 10, "F": 5, "G": 5})
    for c in range(yc0, ncol): ws.column_dimensions[L(c)].width = 3.6
    ws.column_dimensions[L(ncol)].width = 30
    r = 7 if first else 8
    no = 0; costsum = {}
    eqs = sorted(W.EQS, key=sort_key_block)
    for furn, label in [(1, "■ 1号炉"), (2, "■ 2号炉"), (0, "■ 共通設備")]:
        merge(ws, r, 1, r, ncol, label, f(10, True), al("left"), B_ALL, fill("FFF2CC"))
        rec.add("parent", {"path": [label.replace("■ ", "")]}, ws.title, r); r += 1
        rows = [e for e in eqs if e.furn == furn and W.start_fy(e) <= V + (1 if e.base == "排ガス水銀連続分析計" else 0) and not (e.base == "排ガス水銀連続分析計" and V < 2018)]
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
            put(ws, r, 4, e.name, f(8), al("left"), B_HAIR)
            put(ws, r, 5, e.tag, f(8), al("center"), B_HAIR)
            put(ws, r, 6, e.duty if e.duty != "交互" else "", f(8), al("center"), B_HAIR)
            cyc = None
            if not first:
                cyc = cyc_of(e, "spc2016"); put(ws, r, 7, cyc, f(8), al("center"), B_HAIR)
            m = marks(e, V, Y0, Y1)
            mk = {}
            for y in range(Y0, Y1 + 1):
                s = SYM.get(m.get(y))
                cc = put(ws, r, yc0 + y - Y0, s, f(9), al("center"), B_HAIR)
                if y < V: cc.fill = fill("F2F2F2")
                if s: mk[str(y)] = s
                if m.get(y) in ("O", "R", "P", "U"):
                    costsum.setdefault(e.group, {}).setdefault(y, 0); costsum[e.group][y] += W.cost(e, m[y])
            note = note_sv1(e, V)
            put(ws, r, ncol, note, f(8), al("left", shrink=True), B_HAIR)
            rec.add("item", {"path": [label.replace("■ ", ""), e.group, e.sub, e.name], "code": e.tag, "cycle": cyc, "marks": mk, "memo": note,
                             "duty": e.duty if e.duty != "交互" else ""}, ws.title, r)
            r += 1
        if furn != 2 and gstart < r - 1: ws.merge_cells(start_row=gstart, start_column=2, end_row=r - 1, end_column=2)
        if furn == 1 and sstart < r - 1: ws.merge_cells(start_row=sstart, start_column=3, end_row=r - 1, end_column=3)
    merge(ws, r, 1, r, ncol, "■ 法定点検・測定", f(10, True), al("left"), B_ALL, fill("FFF2CC"))
    rec.add("parent", {"path": ["法定点検・測定"]}, ws.title, r); r += 1
    for name, basis, cyc, m in legal(V, Y0, Y1):
        if "圧力容器" in name: continue
        merge(ws, r, 2, r, 4, name, f(8), al("left"), B_HAIR)
        put(ws, r, 5, basis, f(7), al("center", shrink=True), B_HAIR)
        put(ws, r, 6, None if not first else cyc, f(8), al("center"), B_HAIR)
        if not first: put(ws, r, 7, cyc, f(8), al("center"), B_HAIR)
        put(ws, r, 1, None, None, None, B_HAIR)
        mk = {}
        for y in range(Y0, Y1 + 1):
            s = SYM.get(m.get(y)); cc = put(ws, r, yc0 + y - Y0, s, f(9), al("center"), B_HAIR)
            if s: mk[str(y)] = s
            if y < V: cc.fill = fill("F2F2F2")
        put(ws, r, ncol, "", f(8), None, B_HAIR)
        rec.add("item", {"path": ["法定点検・測定", name], "work": basis, "cycle": int(cyc[0]) if cyc.endswith("年") else 0.5, "marks": mk}, ws.title, r)
        r += 1
    last = r - 1
    if not first:
        r += 1
        merge(ws, r, 1, r, ncol, "■ 年度別概算費用（百万円・税抜、点検費は定期整備費に含む）", f(10, True), al("left"), B_ALL, fill("FFF2CC"))
        rec.add("parent", {"path": ["年度別概算費用"]}, ws.title, r); r += 1
        tot = {}
        for g in GROUP_ORDER:
            if g not in costsum: continue
            merge(ws, r, 2, r, 7, g, f(8), al("left"), B_HAIR)
            vals = {}
            for y in range(Y0, Y1 + 1):
                v = round(costsum[g].get(y, 0) / 1000, 1); tot[y] = tot.get(y, 0) + v
                put(ws, r, yc0 + y - Y0, v if v else None, f(7), al("right", shrink=True), B_HAIR, fmt="0.0")
                if v: vals[str(y)] = v
            rec.add("subtotal", {"path": ["年度別概算費用", g], "unit": "百万円", "costs": vals}, ws.title, r)
            r += 1
        merge(ws, r, 2, r, 7, "合　計", f(9, True), al("center"), B_ALL, fill("EDEDED"))
        vals = {}
        for y in range(Y0, Y1 + 1):
            v = round(tot.get(y, 0), 1); vals[str(y)] = v
            put(ws, r, yc0 + y - Y0, v, f(7, True), al("right", shrink=True), B_ALL, fill("EDEDED"), fmt="0.0")
        rec.add("total", {"path": ["年度別概算費用", "合計"], "unit": "百万円", "costs": vals}, ws.title, r)
        last = r
    outline(ws, 5, 1, last, ncol)
    ws.freeze_panes = "G7" if first else "H8"
    page(ws, "landscape", 8, 1, 0, f"A1:{L(ncol)}{last}", "5:6" if first else "5:7")
    ws.oddFooter.center.text = "&P / &N"
    ws.oddFooter.right.text = "ときわ北部エコサービス 保全G"
    props(wb, ttl.replace("　", " "), "sv1（【架空】ときわ北部エコサービス 保全グループ）", dt.datetime(md.year, md.month, md.day, 10))
    wb.save(path); rec.save()

# ============================================================== sv2（2017–2026）
def sv2(V, path, fname):
    rec = Rec(fname, plan=True)
    wb = Workbook(); ws = wb.active; ws.title = "長期修繕計画"
    Y0, Y1 = 2009, 2033
    yc0 = 11; ncol = yc0 + (Y1 - Y0 + 1)
    SYM = {"I": "○", "O": "◎", "P": "△", "R": "●", "U": "●", "L": "◇", "K": "◆"}
    md = dt.date(2017, 5, 22) if V == 2017 else dt.date(V, 4, 20 + V % 5)
    ttl = f"{W.FAC}　長期修繕計画表（{V}年度改訂）"
    merge(ws, 1, 1, 1, ncol, ttl, f(16, True), al("center"), None)
    ph = "（第1期 2009〜2023年度／2024年度以降は第2期の参考）" if V <= 2023 else "（第2期 運営委託 2024〜2033年度、整備は年度上限内）"
    merge(ws, 2, 1, 2, ncol, f"{W.SPC}　保全グループ　作成：sv1　改訂：sv2 {md.year}/{md.month}/{md.day}　{ph}", f(9), al("right"), None)
    put(ws, 3, 1, "凡例", f(9, True))
    leg = "○：点検　◎：整備（分解・補修）　△：部分補修　●：更新　◇：法定" + ("　◆：基幹改良（組合発注）" if V >= 2024 else "")
    merge(ws, 3, 2, 3, 6, leg, f(9), al("left"), None)
    hdr = ["No.", "設備", "系統", "号炉", "機器名称", "タグNo.", "区分", "周期\n(年)", "重要度", "前回\n整備"]
    for i, h in enumerate(hdr):
        merge(ws, 5, i + 1, 6, i + 1, h, f(9, True), al("center", wrap=True), B_ALL, fill("DDEBF7" if i not in (3, 8, 9) else "E2EFDA"))
    merge(ws, 5, yc0, 5, yc0 + 2023 - Y0, "第1期（西暦年度）", f(9, True), al("center"), B_ALL, fill("DDEBF7"))
    merge(ws, 5, yc0 + 2024 - Y0, 5, ncol - 1, "第2期（参考）" if V <= 2023 else "第2期", f(9, True), al("center"), B_ALL, fill("FCE4D6"))
    for y in range(Y0, Y1 + 1):
        put(ws, 6, yc0 + y - Y0, f"'{str(y)[2:]}", f(9, True), al("center"), B_ALL, fill("DDEBF7" if y <= 2023 else "FCE4D6"))
    merge(ws, 5, ncol, 6, ncol, "備　考", f(9, True), al("center"), B_ALL, fill("DDEBF7"))
    nendo_row(ws, 7, yc0, Y0, Y1, 10)
    widths(ws, {"A": 5, "B": 11, "C": 11, "D": 5, "E": 26, "F": 10, "G": 5, "H": 5, "I": 5, "J": 6})
    for c in range(yc0, ncol): ws.column_dimensions[L(c)].width = 3.6
    ws.column_dimensions[L(ncol)].width = 30
    for y in range(Y0, V):
        ws.column_dimensions[L(yc0 + y - Y0)].hidden = True
    r = 8; no = 0
    merge(ws, r, 1, r, ncol, "■ 1号炉", f(10, True), al("left"), B_ALL, fill("FFF2CC")); r += 1   # 旧版の見出しの残り
    put(ws, r, 1, None); r += 1
    if V >= 2021:
        r += 1          # 2021 版で 1 行挿入した跡（空行が 2 行に）
    eqs = sorted(W.EQS, key=lambda e: (GROUP_ORDER.index(e.group), e.sys, e.base, e.suffix, e.furn if e.furn else 9))
    eqs = [e for e in eqs if not (e.base == "排ガス水銀連続分析計" and V < 2018)]
    costsum = {}; prevg = None; gstart = r
    for e in eqs:
        if e.group != prevg:
            if prevg is not None and gstart < r - 1: ws.merge_cells(start_row=gstart, start_column=2, end_row=r - 1, end_column=2)
            gstart = r; put(ws, r, 2, e.group, f(8), al("center", "center", True), B_HAIR)
        else:
            put(ws, r, 2, None, None, None, B_HAIR)
        prevg = e.group; no += 1
        put(ws, r, 1, no, f(8), al("center"), B_HAIR)
        put(ws, r, 3, e.sub, f(8), al("left", shrink=True), B_HAIR)
        put(ws, r, 4, {0: "共", 1: "1", 2: "2"}[e.furn], f(8), al("center"), B_HAIR)
        put(ws, r, 5, e.name, f(8), al("left"), B_HAIR)
        put(ws, r, 6, e.tag, f(8), al("center"), B_HAIR)
        put(ws, r, 7, e.duty if e.duty != "交互" else "", f(8), al("center"), B_HAIR)
        cyc = cyc_of(e, "spc2017"); put(ws, r, 8, cyc, f(8), al("center"), B_HAIR)
        put(ws, r, 9, e.crit, f(8), al("center"), B_HAIR)
        mall = marks(e, V, V, Y1)
        old = PG.marks_for(e, "spc2016", Y0, 2016)
        if V >= 2018:
            old.update({y: c for y, c in PG.marks_for(e, "spc2017", 2017, min(V - 1, 2023)).items()})
        if V >= 2025:
            old.update({y: c for y, c in PG.marks_for(e, "u2024", 2024, V - 1).items()})
        past = [y for y, c in old.items() if c in ("O", "R") and y < V]
        prev_ = f"'{str(max(past))[2:]}" if past else None
        put(ws, r, 10, prev_, f(8), al("center"), B_HAIR)
        mk = {}
        for y in range(Y0, Y1 + 1):
            code = old.get(y) if y < V else mall.get(y)
            s = SYM.get(code); put(ws, r, yc0 + y - Y0, s, f(9), al("center"), B_HAIR)
            if s: mk[str(y)] = s
            if code in ("O", "R", "P", "U") and y >= V:
                costsum.setdefault(e.group, {}).setdefault(y, 0); costsum[e.group][y] += W.cost(e, code)
        note = note_sv2(e, V, mall)
        cc = put(ws, r, ncol, note, f(8), al("left", shrink=True), B_HAIR)
        comments = {}
        if REP_BY_TAG.get(e.tag) == "R3":
            txt = "2015.10 定期事業者検査で二次過熱器下部ベンド減肉 4.0→2.6mm。肉盛補修済。次回 2017 で再測定。sv2"
            cc.comment = Comment(txt, "sv2"); comments[f"{L(ncol)}{r}"] = txt
        rec.add("item", {"path": [e.group, e.sub, e.name], "furnace": {0: "共", 1: "1", 2: "2"}[e.furn], "code": e.tag, "cycle": cyc, "crit": e.crit,
                         "last": prev_, "marks": mk, "memo": note, "comments": comments,
                         "duty": e.duty if e.duty != "交互" else ""}, ws.title, r)
        r += 1
    ws.merge_cells(start_row=gstart, start_column=2, end_row=r - 1, end_column=2)
    merge(ws, r, 1, r, ncol, "■ 法定点検・測定", f(10, True), al("left"), B_ALL, fill("FFF2CC"))
    rec.add("parent", {"path": ["法定点検・測定"]}, ws.title, r); r += 1
    for name, basis, cyc, m in legal(V, V, Y1):
        old = {y: "L" for y in [2011, 2014] if "精密" in name}
        merge(ws, r, 2, r, 5, name, f(8), al("left"), B_HAIR)
        put(ws, r, 6, basis, f(7), al("center", shrink=True), B_HAIR)
        put(ws, r, 8, cyc, f(8), al("center"), B_HAIR)
        for c_ in (1, 7, 9, 10): put(ws, r, c_, None, None, None, B_HAIR)
        mk = {}
        for y in range(Y0, Y1 + 1):
            s = SYM.get(m.get(y) if y >= V else old.get(y)); put(ws, r, yc0 + y - Y0, s, f(9), al("center"), B_HAIR)
            if s: mk[str(y)] = s
        nt = ("2018 実施（周期 4 年となった）" if V <= 2018 else "") if "精密" in name else ("2017 から計上" if "圧力容器" in name else "")
        put(ws, r, ncol, nt, f(8), al("left", shrink=True), B_HAIR)
        rec.add("item", {"path": ["法定点検・測定", name], "work": basis, "cycle": int(cyc[0]) if cyc.endswith("年") else 0.5, "marks": mk, "memo": nt}, ws.title, r)
        r += 1
    last = r - 1
    outline(ws, 5, 1, last, ncol)
    ws.freeze_panes = "K8"
    page(ws, "landscape", 8, 1, 0, f"A1:{L(ncol)}{last}", "5:7")
    ws.oddFooter.center.text = "&P / &N"; ws.oddFooter.right.text = "ときわ北部エコサービス 保全G"
    ws2 = wb.create_sheet("費用集計")
    merge(ws2, 1, 1, 1, 20, f"{W.FAC}　長期修繕計画 年度別費用集計（千円・税抜）{V}年度改訂", f(13, True), al("left"), None)
    put(ws2, 2, 1, "※点検費は年度の定期整備費に含むため計上しない。" + ("2024年度以降は第2期の参考値。" if V <= 2023 else "基幹改良（◆）は組合発注のため含まない。"), f(8))
    yrs = list(range(V, 2034))
    put(ws2, 4, 1, "設備", f(9, True), al("center"), B_ALL, fill("DDEBF7"))
    for i, y in enumerate(yrs):
        put(ws2, 4, 2 + i, f"{y}年度", f(8, True), al("center"), B_ALL, fill("DDEBF7" if y <= 2023 else "FCE4D6"))
    rr = 5; tot = {}
    for g in GROUP_ORDER:
        if g not in costsum: continue
        put(ws2, rr, 1, g, f(9), al("left"), B_ALL)
        vals = {}
        for i, y in enumerate(yrs):
            v = costsum[g].get(y, 0); tot[y] = tot.get(y, 0) + v
            put(ws2, rr, 2 + i, v if v else None, f(8), al("right"), B_ALL, fmt="#,##0")
            if v: vals[str(y)] = v
        rec.add("subtotal", {"path": ["費用集計", g], "unit": "千円", "costs": vals}, ws2.title, rr)
        rr += 1
    put(ws2, rr, 1, "合計", f(9, True), al("center"), B_ALL, fill("EDEDED"))
    vals = {}
    for i, y in enumerate(yrs):
        v = tot.get(y, 0)
        if y == 2019 and V <= 2019:
            v = v - W.COST_R["dcsc"] * 4   # DCS 更新分を後から足し、合計を打ち直していない（わざと）
        put(ws2, rr, 2 + i, v, f(8, True), al("right"), B_ALL, fill("EDEDED"), fmt="#,##0"); vals[str(y)] = v
    rec.add("total", {"path": ["費用集計", "合計"], "unit": "千円", "costs": vals}, ws2.title, rr)
    if V <= 2019: put(ws2, rr + 2, 1, "2019年度は DCS 更新（別途工事）を含む", f(8))
    widths(ws2, {"A": 16}); [setattr(ws2.column_dimensions[L(c)], "width", 9.5) for c in range(2, 2 + len(yrs))]
    page(ws2, "landscape", 9, 1, 1, f"A1:{L(1 + len(yrs))}{rr + 2}")
    props(wb, ttl.replace("　", " "), "sv1（【架空】ときわ北部エコサービス 保全グループ）", dt.datetime(md.year, md.month, md.day, 16), "sv2")
    wb.save(path); rec.save()

# ============================================================== 組合（R6 策定／R7 改訂）
def union(V, path, fname):
    rec = Rec(fname, plan=True)
    rev = V >= 2025
    wb = Workbook()
    Y0, Y1 = 2024, 2033
    SYM = {"I": "点", "O": "整", "P": "補", "R": "更", "U": "改", "L": "法"}
    T = "中長期整備計画（令和6年度策定・令和7年度改訂）" if rev else "中長期整備計画（令和6年度策定）"
    summ = {}; first = True
    for sheet, groups in U_GROUPS:
        ws = wb.active if first else wb.create_sheet(); first = False
        ws.title = sheet
        NY = Y1 - Y0 + 1; nc = 6 + NY * 3 + 1
        merge(ws, 1, 1, 1, nc, f"{W.FAC}　{T}　【{sheet}設備】", f(14, True), al("left"), None)
        sub = ("策定 u1 令和6年11月／改訂 u1 令和8年3月（R6・R7 実績記入、基幹改良の工期変更を反映）" if rev else "策定 u1 令和6年11月") + "　計画期間 令和6〜15年度、千円・税込"
        merge(ws, 2, 1, 2, nc, f"{W.UNION}　施設課　{sub}", f(9), al("right"), None)
        merge(ws, 3, 1, 3, nc, "記号　点：点検　整：整備　補：補修　更：更新　改：改良（基幹的設備改良工事）　費用は計画（概算）額。点検は運営委託料に含む。実績欄は年度末に記入", f(8), al("left"), None)
        for i, h in enumerate(["番号", "中分類", "機器名称", "号炉", "設置", "重要度"]):
            merge(ws, 5, i + 1, 6, i + 1, h, f(9, True), al("center", wrap=True), B_ALL, fill("E2EFDA"))
        for k, y in enumerate(range(Y0, Y1 + 1)):
            c = 7 + k * 3
            merge(ws, 5, c, 5, c + 2, wareki_fy(y).replace("年度", ""), f(9, True), al("center"), B_ALL, fill("E2EFDA" if y not in (2025, 2026, 2027) else "FFE699"))
            put(ws, 6, c, "計画", f(6, True), al("center"), B_ALL, fill("E2EFDA"))
            put(ws, 6, c + 1, "実績", f(6, True), al("center"), B_ALL, fill("FCE4D6" if (rev and y <= 2025) else "E2EFDA"))
            put(ws, 6, c + 2, "費用", f(6, True), al("center"), B_ALL, fill("E2EFDA"))
        merge(ws, 5, nc, 6, nc, "備考", f(9, True), al("center"), B_ALL, fill("E2EFDA"))
        widths(ws, {"A": 6, "B": 11, "C": 24, "D": 4.5, "E": 5, "F": 4.5})
        for k in range(NY):
            ws.column_dimensions[L(7 + k * 3)].width = 3.4; ws.column_dimensions[L(8 + k * 3)].width = 3.4; ws.column_dimensions[L(9 + k * 3)].width = 7.2
        ws.column_dimensions[L(nc)].width = 28
        r = 7
        eqs = [e for e in W.EQS if e.group in groups]
        eqs.sort(key=lambda e: (groups.index(e.group), e.sys, e.furn if e.furn else 9, e.idx))
        prev = None; mstart = r; seqn = 0; sub_tot = {}
        for e in eqs:
            seqn += 1
            if e.sub != prev:
                if prev is not None and mstart < r - 1: ws.merge_cells(start_row=mstart, start_column=2, end_row=r - 1, end_column=2)
                mstart = r; put(ws, r, 2, e.sub, f(8), al("center", "center", True), B_HAIR)
            else:
                put(ws, r, 2, None, None, None, B_HAIR)
            prev = e.sub
            no = f"{sheet[:2]}-{seqn:03d}"
            put(ws, r, 1, no, f(7), al("center"), B_HAIR)
            put(ws, r, 3, u_name(e), f(8), al("left", shrink=True), B_HAIR)
            put(ws, r, 4, {0: "共通", 1: "1号", 2: "2号"}[e.furn], f(7), al("center"), B_HAIR)
            setti = "H20" if e.install.startswith("2009") else wareki_fy(W.start_fy(e) - 1, True)
            put(ws, r, 5, setti, f(7), al("center"), B_HAIR)
            put(ws, r, 6, {"A": "Ⅰ", "B": "Ⅱ", "C": "Ⅲ"}[e.crit], f(8), al("center"), B_HAIR)
            m = PG.marks_for(e, "u2024", Y0, Y1)
            act, carried = actual_u(e, m) if rev else ({}, False)
            mk = {}; ak = {}; costs = {}
            for k, y in enumerate(range(Y0, Y1 + 1)):
                c = 7 + k * 3; code = m.get(y)
                put(ws, r, c, SYM.get(code), f(8, code == "U"), al("center"), B_HAIR,
                    fill("FFE699") if code == "U" or (code == "R" and e.tag in KIKAN and KIKAN[e.tag][1] == y) else None)
                if code: mk[str(y)] = SYM[code]
                a_ = act.get(y)
                put(ws, r, c + 1, SYM.get(a_) if a_ else None, f(8, color="C00000" if a_ and a_ != code else "000000"), al("center"), B_HAIR,
                    fill("F2F2F2") if (not rev or y > 2025) else None)
                if a_: ak[str(y)] = SYM[a_]
                v = None
                if code in ("O", "P", "R", "U"):
                    v = KIKAN[e.tag][2] if (e.tag in KIKAN and KIKAN[e.tag][1] == y) else round(W.cost(e, code) * 1.1)
                    if v: sub_tot[y] = sub_tot.get(y, 0) + v; costs[str(y)] = v
                put(ws, r, c + 2, v if v else None, f(7), al("right", shrink=True), B_HAIR, fmt="#,##0")
            note = ""
            if e.tag in KIKAN:
                note = "基幹改良 " + KIKAN[e.tag][3]
                if rev and KIKAN[e.tag][1] == 2025: note += "。R8.1.16 完了（工期延長）"
            elif e.tag in REP_BY_TAG:
                note = {"R6": "R7 は基幹改良と重なるため点検のみ", "R7": "刃交換は運営委託の整備費", "R8": "ろ布 R7 交換" + ("済" if rev else ""),
                        "R9": "性能検査 奇数年度", "R10": "定期事業者検査 R6 実施", "R11": "OS サポート R10 まで", "R12": "R8 1号停止中に交換",
                        "R2": "R7 定修 残存厚 118mm" if rev else ""}.get(REP_BY_TAG[e.tag], "")
            if carried: note = (note + " " if note else "") + "R6 整備を R7 へ繰越"
            put(ws, r, nc, note, f(7), al("left", shrink=True), B_HAIR)
            rec.add("item", {"no": no, "path": [sheet + "設備", e.sub, u_name(e)], "furnace": {0: "共通", 1: "1号", 2: "2号"}[e.furn], "installed": setti,
                             "crit": {"A": "Ⅰ", "B": "Ⅱ", "C": "Ⅲ"}[e.crit], "marks": mk, "actual": ak, "costs": costs, "unit": "千円", "memo": note}, ws.title, r)
            r += 1
        if mstart < r - 1: ws.merge_cells(start_row=mstart, start_column=2, end_row=r - 1, end_column=2)
        merge(ws, r, 1, r, 6, "小　計（計画額）", f(9, True), al("center"), B_ALL, fill("EDEDED"))
        for k, y in enumerate(range(Y0, Y1 + 1)):
            c = 7 + k * 3
            put(ws, r, c, None, None, None, B_ALL, fill("EDEDED")); put(ws, r, c + 1, None, None, None, B_ALL, fill("EDEDED"))
            put(ws, r, c + 2, sub_tot.get(y, 0), f(7, True), al("right", shrink=True), B_ALL, fill("EDEDED"), fmt="#,##0")
        put(ws, r, nc, None, None, None, B_ALL, fill("EDEDED"))
        rec.add("subtotal", {"path": [sheet + "設備", "小計"], "unit": "千円", "costs": {str(y): sub_tot.get(y, 0) for y in range(Y0, Y1 + 1)}}, ws.title, r)
        summ[sheet] = sub_tot
        outline(ws, 5, 1, r, nc)
        ws.freeze_panes = "G7"
        page(ws, "landscape", 8, 1, 0, f"A1:{L(nc)}{r}", "5:6")
        ws.oddHeader.right.text = "ときわ北部環境組合"; ws.oddFooter.center.text = f"{sheet} &P/&N"
    ws = wb.create_sheet("法定点検")
    merge(ws, 1, 1, 1, 14, f"{W.FAC}　{T}　【法定点検・測定】", f(14, True), al("left"), None)
    for i, h in enumerate(["項目", "根拠", "周期"] + [wareki_fy(y).replace("年度", "") for y in range(Y0, Y1 + 1)]):
        put(ws, 3, 1 + i, h, f(9, True), al("center"), B_ALL, fill("E2EFDA"))
    r = 4
    for name, basis, cyc, m in legal(V, Y0, Y1, union=True):
        put(ws, r, 1, name, f(8), al("left"), B_ALL); put(ws, r, 2, basis, f(8), al("center"), B_ALL); put(ws, r, 3, cyc, f(8), al("center"), B_ALL)
        mk = {}
        for k, y in enumerate(range(Y0, Y1 + 1)):
            put(ws, r, 4 + k, "法" if m.get(y) else None, f(8), al("center"), B_ALL)
            if m.get(y): mk[str(y)] = "法"
        rec.add("item", {"path": ["法定点検・測定", name], "work": basis, "cycle": int(cyc[0]) if cyc.endswith("年") else 0.5, "marks": mk}, ws.title, r)
        r += 1
    widths(ws, {"A": 40, "B": 14, "C": 6}); [setattr(ws.column_dimensions[L(c)], "width", 7) for c in range(4, 14)]
    page(ws, "landscape", 9, 1, 1, f"A1:M{r - 1}")
    ws = wb.create_sheet("総括表", 0)
    merge(ws, 1, 1, 1, 12, f"{W.FAC}　{T}　総括表", f(15, True), al("left"), None)
    merge(ws, 2, 1, 2, 12, f"{W.UNION}　施設課　u1　（単位：千円・税込）", f(9), al("right"), None)
    put(ws, 4, 1, "設備区分", f(9, True), al("center"), B_ALL, fill("E2EFDA"))
    for k, y in enumerate(range(Y0, Y1 + 1)):
        put(ws, 4, 2 + k, wareki_fy(y), f(8, True), al("center", wrap=True), B_ALL, fill("E2EFDA" if y not in (2025, 2026, 2027) else "FFE699"))
    put(ws, 4, 12, "計", f(9, True), al("center"), B_ALL, fill("E2EFDA"))
    r = 5; colt = {}
    for sheet, _ in U_GROUPS:
        put(ws, r, 1, sheet + "設備", f(9), al("left"), B_ALL)
        rowt = 0; vals = {}
        for k, y in enumerate(range(Y0, Y1 + 1)):
            v = summ[sheet].get(y, 0)
            if sheet == "燃焼溶融" and y == 2026: v = round(v / 1.1)      # 税抜のまま転記（わざと）
            colt[y] = colt.get(y, 0) + v; rowt += v; vals[str(y)] = v
            put(ws, r, 2 + k, v, f(8), al("right"), B_ALL, fmt="#,##0")
        put(ws, r, 12, rowt, f(8, True), al("right"), B_ALL, fmt="#,##0")
        rec.add("subtotal", {"path": ["総括表", sheet + "設備"], "unit": "千円", "costs": vals, "cost": rowt}, ws.title, r)
        r += 1
    put(ws, r, 1, "合　計", f(9, True), al("center"), B_ALL, fill("EDEDED"))
    for k, y in enumerate(range(Y0, Y1 + 1)):
        put(ws, r, 2 + k, colt[y], f(8, True), al("right"), B_ALL, fill("EDEDED"), fmt="#,##0")
    put(ws, r, 12, sum(colt.values()), f(8, True), al("right"), B_ALL, fill("EDEDED"), fmt="#,##0")
    rec.add("total", {"path": ["総括表", "合計"], "unit": "千円", "costs": {str(y): colt[y] for y in colt}, "cost": sum(colt.values())}, ws.title, r)
    r += 2
    n1 = "1　令和7〜9年度は基幹的設備改良工事（組合発注、継続費）を含む。2号炉 R7、1号炉 R8、共通 R9。" + ("R7 分は過熱器管材の納期遅れで R8.1.16 完了（変更契約 R7.12.5、年度内）。" if rev else "割付は契約時点の工程による。")
    for t in [n1, "2　運営委託（第2期）の年度上限内で行う整備は、運営事業者の長期修繕計画と調整のうえ計上した。", "3　金額は令和6年度単価による概算（税込）。物価変動は見込んでいない。"]:
        merge(ws, r, 1, r, 12, t, f(8), al("left", wrap=True), None); ws.row_dimensions[r].height = 24; r += 1
    widths(ws, {"A": 18, "L": 12}); [setattr(ws.column_dimensions[L(c)], "width", 11) for c in range(2, 12)]
    page(ws, "landscape", 9, 1, 1, f"A1:L{r - 1}")
    props(wb, f"{W.FAC} {T}", "u1（【架空】ときわ北部環境組合 施設課）", dt.datetime(2024, 11, 28, 14))
    if rev: wb.properties.modified = dt.datetime(2026, 3, 24, 16)
    wb.save(path); rec.save()
