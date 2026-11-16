# -*- coding: utf-8 -*-
"""中長期整備計画（星取表）全版：2016（第1版）〜2026（第11版）。
s1 様式：2016〜2022（2018 改訂で「廃止まで」列を挿入、2027〜2030 の旧計画を隠し列で残す）。
s2 様式：2023〜2026（廃止までの年度を 1号・2号の小列で並べる）。"""
import datetime as dt
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.pagebreak import Break
import world as W
import works as WK
from xlhelp import *
from rec import Rec, path

SYM1 = {"I": "○", "O": "◎", "R": "△", "X": "●", "U": "●"}
SYM2 = {"I": "点", "O": "整", "R": "補", "X": "換", "U": "改"}
FILL2 = {"整": "DDEBF7", "補": "FCE4D6", "換": "F8CBAD", "改": "F8CBAD"}
COST_O = {"idf": 6800, "fdf": 4200, "fan_m": 1800, "hydraulic": 3500, "burner": 900, "sootb": 2600, "pump_r": 900,
          "pump_s": 900, "bfp_r": 2400, "bfp_s": 2400, "feeder_s": 700, "blower": 600, "inv": 2200,
          "crane_r": 5200, "crane_s": 5200, "crane_a": 3800, "accfan": 1500, "comp": 1100, "turbine": 38000}
COST_R = {"furnace": 14000, "stoker_d": 9000, "stoker_c": 16000, "stoker_a": 8000, "feeder": 3000, "spray": 600,
          "conv": 2200, "boiler": 6500, "sh": 9000, "eco": 7000, "tower": 4800, "duct": 3200, "ashpush": 3600,
          "gate": 700, "acc": 5500, "ct": 2800, "mixer": 4200, "cat": 20000, "bf": 6000, "dcs": 15000,
          "crane_r": 6000, "panel": 3000, "tr": 4000, "hv": 6000, "dc": 3000, "ups": 2500}
COST_X = {"bf": 26000, "cat": 72000, "dcs": 90000, "inv": 12000, "bfp_r": 9500, "scale": 18000, "panel": 15000,
          "tr": 22000, "hv": 40000, "crane_r": 160000, "mixer": 30000, "acc": 120000, "sh": 45000, "dc": 9000, "ups": 12000}


def cost(e, w):
    c = 0
    for k in w:
        if k == "I":
            c += {"A": 250, "B": 90, "C": 40}[e.imp]
        elif k == "O":
            c += COST_O.get(e.strat, 500)
        elif k == "R":
            c += COST_R.get(e.strat, 800)
        elif k in "XU":
            c += COST_X.get(e.strat, 6000)
    return c


def ordered():
    out = []
    for cat in W.CATS:
        for x in [x for x in W.PER_FURNACE if x[0] == cat]:
            out.append((cat, x[1], [W.eq(x[1], 1), W.eq(x[1], 2)]))
        for x in W.COMMON:
            if x[0] == cat:
                out.append((cat, x[1], [W.eq(x[1], 0)]))
    return out


def fname(v):
    sty = "s1様式" if v <= 2022 else "s2様式"
    return f"【架空】ひばり野クリーンセンター_中長期整備計画_{v}年度改訂_{sty}.xlsx"


def made(v):
    return dt.datetime(v, 4, 15 if v <= 2022 else 20, 10, 0)


def edition(v):
    return v - 2015


def notes_for(e, v):
    """その版の時点で書ける備考。"""
    n = []
    k = (e.name, e.f)
    if k == ("ボイラ本体", 2):
        if v >= 2018:
            n.append("前壁水管 減肉傾向（UT 年0.15mm程度）")
        if v >= 2022:
            n.append("22定修 UT 前壁水管減肉 最小3.2mm → 市へ報告（R5追加工事 要望）")
    if k == ("押込送風機", 2) and v >= 2018:
        n.append("分解整備 17→25（周期8年 廃止までのため延伸）")
    if k == ("誘引通風機", 1) and v >= 2021:
        n.append("振動 傾向管理中（" + ("21年度末 3.8mm/s" if v == 2022 else "20年度 3.2mm/s") + "）")
    if k == ("煙突内筒", 2) and v >= 2020:
        n.append("20点検 下部腐食 → 市追加工事（R3）予定" if v == 2020 else "R3 市追加工事で下部補修")
    if k == ("バグフィルタ", 1) and v >= 2019:
        n.append("19 ろ布全数交換")
    if k == ("触媒反応塔", 1) and v >= 2018:
        n.append("H30 全量更新中止。活性試験で判断")
    return n


# =============================================================== s1 様式
def build_s1(v):
    fn = fname(v)
    R = Rec(fn, plan=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "整備計画"
    author = "SPC保全課 s1" if v <= 2021 else "SPC保全課 s2"
    props(wb, f"【架空】ひばり野クリーンセンター 中長期整備計画 {v}年度改訂", author, made(v))
    Y = W.YEARS
    has_pol = v >= 2018
    C_NO, C_CAT, C_NAME, C_F, C_ASSET = 1, 2, 3, 4, 5
    C_POL = 6 if has_pol else None
    C_PA = 7 if has_pol else 6
    C_Y0 = C_PA + 1
    C_NOTE = C_Y0 + len(Y)
    C_HID = C_NOTE + 1          # 2018 改訂で残した旧計画（H29版 2027〜2030）の隠し列
    for c, w in {C_NO: 4.5, C_CAT: 3.5, C_NAME: 24, C_F: 4, C_ASSET: 9.5, C_PA: 3.5, C_NOTE: 30}.items():
        ws.column_dimensions[L(c)].width = w
    if has_pol:
        ws.column_dimensions[L(C_POL)].width = 13
    for i in range(len(Y)):
        ws.column_dimensions[L(C_Y0 + i)].width = 4.6
    put(ws, 1, 1, "【架空】ひばり野クリーンセンター　中長期整備計画（2016〜2030年度）", font(16, True))
    who = "s1作成" if v <= 2021 else "s2改訂（初版 2016/06 s1作成）"
    put(ws, 2, 1, f"{v}年度改訂（第{edition(v)}版）　{made(v):%Y/%m/%d}　ひばり野環境サービス株式会社 保全課　{who}", font(9))
    put(ws, 3, 1, "凡例　○：点検・測定　◎：分解整備　△：補修　●：更新　　上段＝計画（各年度当初の計画）　下段＝実績（網掛け）", font(9))
    if has_pol:
        put(ws, 3, C_Y0 + 6, "2018年度改訂：2030年度末廃止決定により 2027年度以降の更新を取りやめ、補修で延命（赤字△）", font(9, True, color="C00000"))
    R1, R2 = 5, 6
    ws.merge_cells(start_row=R1, start_column=C_NO, end_row=R2, end_column=C_NO)
    put(ws, R1, C_NO, "No", font(9, True), CENTER, B_ALL, fill("D9E1F2"))
    ws.merge_cells(start_row=R1, start_column=C_CAT, end_row=R1, end_column=C_ASSET)
    put(ws, R1, C_CAT, "機　　器", font(9, True), CENTER, B_ALL, fill("D9E1F2"))
    for c, h in ((C_CAT, "区分"), (C_NAME, "機器名称"), (C_F, "炉"), (C_ASSET, "資産番号")):
        put(ws, R2, c, h, font(8, True), CENTER, B_ALL, fill("D9E1F2"))
    if has_pol:
        put(ws, R1, C_POL, "廃止まで", font(9, True), CENTER, B_ALL, fill("FCE4D6"))
        put(ws, R2, C_POL, "（H30追加）", font(7), CENTER, B_ALL, fill("FCE4D6"))
    ws.merge_cells(start_row=R1, start_column=C_PA, end_row=R2, end_column=C_PA)
    put(ws, R1, C_PA, "計実", font(8, True), CENTER, B_ALL, fill("D9E1F2"))
    ws.merge_cells(start_row=R1, start_column=C_Y0, end_row=R1, end_column=C_Y0 + len(Y) - 1)
    put(ws, R1, C_Y0, "年　　度（西暦）", font(9, True), CENTER, B_ALL, fill("D9E1F2"))
    for i, y in enumerate(Y):
        put(ws, R2, C_Y0 + i, str(y)[2:] if y != 2016 else "2016", font(8, True), CENTER, B_ALL,
            fill("BFBFBF" if y < v else ("FFE699" if y == v else "D9E1F2")))
    ws.merge_cells(start_row=R1, start_column=C_NOTE, end_row=R2, end_column=C_NOTE)
    put(ws, R1, C_NOTE, "備　考", font(9, True), CENTER, B_ALL, fill("D9E1F2"))
    if has_pol:
        ws.merge_cells(start_row=R1, start_column=C_HID, end_row=R1, end_column=C_HID + 3)
        put(ws, R1, C_HID, "（H29版 27〜30）", font(7), CENTER)
        for i, y in enumerate(range(2027, 2031)):
            put(ws, R2, C_HID + i, str(y)[2:], font(7), CENTER)
            ws.column_dimensions[L(C_HID + i)].hidden = True
            ws.column_dimensions[L(C_HID + i)].width = 4
    box(ws, R1, 1, R2, C_NOTE)
    r = R2 + 1
    no = 0
    curcat = None
    pastfill = fill("E7E6E6")
    red = font(9, color="C00000")
    cost_by = {}
    since = r
    for cat, name, eqs in ordered():
        if r - since >= 50:
            ws.row_breaks.append(Break(id=r - 1))
            since = r
        if cat != curcat:
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=C_NOTE)
            put(ws, r, 1, f"■ {cat}　{W.CATS[cat]}", font(10, True), LEFT, None, fill("EDEDED"))
            box(ws, r, 1, r, C_NOTE)
            R.add("整備計画", r, "parent", {"path": [W.CATS[cat]], "code": cat})
            r += 1
            curcat = cat
        for e in eqs:
            no += 1
            fl = {0: "共通", 1: "1", 2: "2"}[e.f]
            for c, val, f_, a_ in ((C_NO, no, font(8), CENTER), (C_NAME, e.name, font(9), LEFT), (C_F, fl, font(8), CENTER),
                                   (C_ASSET, e.no, font(8), CENTER), (C_CAT, cat, font(7), CENTER)):
                ws.merge_cells(start_row=r, start_column=c, end_row=r + 1, end_column=c)
                put(ws, r, c, val, f_, a_)
            if has_pol:
                ws.merge_cells(start_row=r, start_column=C_POL, end_row=r + 1, end_column=C_POL)
                put(ws, r, C_POL, e.policy, font(7), LEFT)
            put(ws, r, C_PA, "計", font(7), CENTER)
            put(ws, r + 1, C_PA, "実", font(7), CENTER)
            notes = list(notes_for(e, v))
            pm, am, cm, ws_ = {}, {}, {}, []
            for i, y in enumerate(Y):
                c = C_Y0 + i
                pv = y if y <= v else v
                pw = W.plan_work(e, y, pv)
                txt = next((SYM1[k] for k in "XUORI" if k in pw), "")
                cell = put(ws, r, c, txt or None, font(9), CENTER)
                if txt:
                    pm[str(y)] = txt
                if has_pol and y >= 2027 and y in W.PROPOSAL_REPLACE.get((e.name, e.f), []):
                    cell.font = red
                    cell.comment = Comment("H30改訂　●更新 → △補修（廃止まで延命）", "s1")
                    cm[f"{L(c)}{r}"] = "H30改訂　●更新 → △補修（廃止まで延命）"
                if (e.name, e.f, y) in W.CITY_WORKS and txt:
                    cell.comment = Comment("市発注（追加工事）", "s1" if v <= 2021 else "s2")
                    cm[f"{L(c)}{r}"] = "市発注（追加工事）"
                if y >= v:
                    cost_by[(cat, y)] = cost_by.get((cat, y), 0) + (0 if (e.name, e.f, y) in W.CITY_WORKS else cost(e, pw))
                if y < v:
                    aw = W.actual_work(e, y) or set()
                    at = next((SYM1[k] for k in "XUORI" if k in aw), "")
                    put(ws, r + 1, c, at or None, font(9), CENTER, None, pastfill)
                    if at:
                        am[str(y)] = at
                        ws_.append(WK.wid(e, y))
                    k = (e.name, e.f, y)
                    if k in W.ACTUAL_DEV and "計画どおり" not in W.ACTUAL_DEV[k][1]:
                        notes.append(f"{str(y)[2:]}：{W.ACTUAL_DEV[k][1]}")
            hid = {}
            if has_pol:
                for i, y in enumerate(range(2027, 2031)):
                    ow = W.plan_work(e, y, 2017)
                    t_ = next((SYM1[k] for k in "XUORI" if k in ow), None)
                    put(ws, r, C_HID + i, t_, font(7), CENTER)
                    if t_:
                        hid[str(y)] = t_
            ws.merge_cells(start_row=r, start_column=C_NOTE, end_row=r + 1, end_column=C_NOTE)
            memo = "／".join(notes) or None
            put(ws, r, C_NOTE, memo, font(7), LEFT)
            box(ws, r, 1, r + 1, C_NOTE, B_HAIR)
            for c in range(1, C_NOTE + 1):
                ws.cell(row=r + 1, column=c).border = Border(left=hair, right=hair, top=dot, bottom=thin)
            ws.row_dimensions[r].height = 12
            ws.row_dimensions[r + 1].height = 12
            pathv = [W.CATS[cat], e.name if e.f == 0 else f"{e.name}（{e.f}号）"]
            f1 = {"path": pathv, "code": e.no, "炉": fl, "計実": "計", "marks": pm, "memo": memo, "comments": cm}
            if has_pol:
                f1["廃止まで"] = e.policy
                if hid:
                    f1["hiddenMarks"] = hid
            R.add("整備計画", r, "item", f1)
            if v > 2016:
                R.add("整備計画", r + 1, "item", {"path": pathv, "code": e.no, "炉": fl, "計実": "実", "marks": am}, ws_)
            r += 2
    ws.row_breaks.append(Break(id=r))
    r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=C_NOTE)
    put(ws, r, 1, "■ 法定点検・定期測定（施設全体）", font(10, True), LEFT, None, fill("EDEDED"))
    box(ws, r, 1, r, C_NOTE)
    R.add("整備計画", r, "parent", {"path": ["法定点検・定期測定"]})
    r += 1
    for (nm, tgt, cyc, rule, conf, basis) in W.LEGAL:
        no += 1
        ws.merge_cells(start_row=r, start_column=C_NAME, end_row=r + 1, end_column=C_NAME)
        put(ws, r, C_NO, no, font(8), CENTER)
        put(ws, r, C_NAME, nm, font(8), LEFT)
        ws.merge_cells(start_row=r, start_column=C_F, end_row=r + 1, end_column=C_PA - 1)
        put(ws, r, C_F, f"{tgt}（{cyc}）", font(7), LEFT)
        put(ws, r, C_PA, "計", font(7), CENTER)
        put(ws, r + 1, C_PA, "実", font(7), CENTER)
        pm, am = {}, {}
        for i, y in enumerate(Y):
            on = rule(y)
            if nm == "計量機 定期検査" and y == 2024 and v in (2021, 2022):
                on = False
            if on:
                pm[str(y)] = "○"
            put(ws, r, C_Y0 + i, "○" if on else None, font(9), CENTER)
            if y < v:
                put(ws, r + 1, C_Y0 + i, "○" if rule(y) else None, font(9), CENTER, None, pastfill)
                if rule(y):
                    am[str(y)] = "○"
        put(ws, r, C_NOTE, basis[:28], font(7), LEFT)
        ws.merge_cells(start_row=r, start_column=C_NOTE, end_row=r + 1, end_column=C_NOTE)
        box(ws, r, 1, r + 1, C_NOTE, B_HAIR)
        cy = int(cyc[0]) if cyc[0].isdigit() else None
        R.add("整備計画", r, "item", {"path": ["法定点検・定期測定", nm], "work": nm, "memo": f"{tgt}（{cyc}）", "cycle": cy, "計実": "計", "marks": pm})
        if v > 2016:
            R.add("整備計画", r + 1, "item", {"path": ["法定点検・定期測定", nm], "work": nm, "計実": "実", "marks": am})
        r += 2
    end = r - 1
    ws.freeze_panes = f"{L(C_PA + 1)}{R2 + 1}"
    page(ws, landscape=True, paper=8, fit_w=1, fit_h=0, area=f"A1:{L(C_NOTE)}{end}", titles=f"{R1}:{R2}",
         footer=f"&L中長期整備計画 {v}年度改訂&R&P / &N")
    # 費用
    wc = wb.create_sheet("年度別概算費用")
    put(wc, 1, 1, "【架空】ひばり野クリーンセンター　中長期整備計画　年度別概算費用（計画）", font(13, True))
    put(wc, 2, 1, f"単位：千円（税抜）　{v}年度改訂　※SPC 年度補修上限額内の分。市追加工事は含まない", font(9))
    YY = [y for y in Y if y >= v]
    put(wc, 4, 1, "設備区分", font(9, True), CENTER, B_ALL, fill("D9E1F2"))
    wc.column_dimensions["A"].width = 26
    for i, y in enumerate(YY):
        put(wc, 4, 2 + i, f"{y}年度", font(9, True), CENTER, B_ALL, fill("D9E1F2"))
        wc.column_dimensions[L(2 + i)].width = 10
    rr = 5
    tot = {y: 0 for y in YY}
    for cat in W.CATS:
        put(wc, rr, 1, f"{cat} {W.CATS[cat]}", font(9), LEFT, B_ALL)
        cs = {}
        for i, y in enumerate(YY):
            val = int(round(cost_by.get((cat, y), 0), -1))
            tot[y] += val
            put(wc, rr, 2 + i, val, font(9), RIGHT, B_ALL, None, "#,##0")
            cs[str(y)] = val
        R.add("年度別概算費用", rr, "item", {"path": [W.CATS[cat]], "code": cat, "cost": cs, "unit": "千円"})
        rr += 1
    if has_pol:
        put(wc, rr, 1, "廃止前 撤去準備調査（H30追加）", font(9, color="C00000"), LEFT, B_ALL)
        extra = {2028: 3200, 2029: 4800, 2030: 6500}
        cs = {}
        for i, y in enumerate(YY):
            put(wc, rr, 2 + i, extra.get(y), font(9, color="C00000"), RIGHT, B_ALL, None, "#,##0")
            if y in extra:
                cs[str(y)] = extra[y]
        R.add("年度別概算費用", rr, "item", {"path": ["廃止前 撤去準備調査（H30追加）"], "cost": cs, "unit": "千円",
                                       "memo": "合計に含まれていない（計画の欠け）"})
        rr += 1
    put(wc, rr, 1, "合　計", font(9, True), CENTER, B_ALL, fill("FFF2CC"))
    for i, y in enumerate(YY):
        put(wc, rr, 2 + i, tot[y], font(9, True), RIGHT, B_ALL, fill("FFF2CC"), "#,##0")
    R.add("年度別概算費用", rr, "total", {"path": ["合計"], "cost": {str(y): tot[y] for y in YY}, "unit": "千円"})
    page(wc, landscape=True, paper=9, fit_w=1, fit_h=1)
    wb.save(path(fn))
    return fn


# =============================================================== s2 様式
NOTE2 = {
    "ボイラ本体": "2号：22UT 前壁減肉 最小3.2mm（必要厚2.9）。R5 市追加工事で6本取替・肉盛。25以降は隣接管を監視",
    "誘引通風機": "1号：振動 22年度末 4.1mm/s（上昇傾向）→ 24 軸受交換・ランナ点検",
    "燃焼ストーカ": "1号：22 測定で摩耗速い → 25予定を24へ前倒し",
    "バグフィルタ": "ろ布 1号19・2号20 全数交換。寿命6年として1号25・2号26に再交換（廃止まで1回）",
    "触媒反応塔": "全量更新は取りやめ（H30）。2年ごと活性試験で段ごとに判断",
    "押込送風機": "2号 分解 17→25（8年、廃止まで延伸）",
    "灰押出装置": "1号 底板腐食 22測定 残厚5.1mm → 24 当て板",
    "混練機": "パドル摩耗 24 交換予定",
    "蒸気タービン": "定期事業者検査 開放 4年（25・29）",
    "ごみクレーン1号": "21 ワイヤ前倒し交換。年次・性能検査は法定欄",
}
NOTE2_LATER = {
    2024: {"燃焼ストーカ": "1号：24 交換（前倒し）。2号：23秋 測定で 24→25 へ延伸", "誘引通風機": "1号：24 軸受交換予定（振動 4.4mm/s）",
           "混練機": "24 パドル交換（予備品使用）", "ごみクレーン1号": "走行車輪 摩耗進行 → 26 補修"},
    2025: {"過熱器": "1号：24UT 2次入口 外面減肉 最小3.6mm → 26 補修", "バグフィルタ": "1号：24 抜取試験 強度残存率72% → 26 へ延伸",
           "誘引通風機": "1号：24 軸受交換・ランナ肉盛済（2.1mm/s）", "燃焼ストーカ": "1号：24 交換済。2号：25 交換",
           "灰押出装置": "1号：24 底板・側板 当て板済", "ボイラ本体": "2号：R5 取替済。25 隣接管監視・肉盛。1号：26 部分補修"},
    2026: {"過熱器": "1号：26 2次入口 部分取替（秋定修）", "バグフィルタ": "1号・2号とも 26 全数交換（廃止まで最後）",
           "誘引通風機": "1号：24 整備後 良好。廃止まで傾向監視", "燃焼ストーカ": "1号：28 残り列", "ボイラ本体": "2号：25 隣接2本肉盛済。1号：26 補修",
           "蒸気タービン": "25 開放済。29 は廃止前年だが法定のため実施"},
}


def build_s2(v):
    fn = fname(v)
    R = Rec(fn, plan=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "整備計画"
    props(wb, f"【架空】ひばり野クリーンセンター 中長期整備計画 {v}年度改訂", "SPC保全課 s2", made(v))
    Y = list(range(v, W.CLOSE_FY + 1))
    C_NO, C_NAME, C_ASSET, C_IMP, C_POL, C_LAST1, C_LAST2 = 1, 2, 3, 4, 5, 6, 7
    C_Y0 = 8
    C_NOTE = C_Y0 + 2 * len(Y)
    for c, w in {C_NO: 4.5, C_NAME: 29, C_ASSET: 15, C_IMP: 4, C_POL: 13, C_LAST1: 9, C_LAST2: 9, C_NOTE: 34}.items():
        ws.column_dimensions[L(c)].width = w
    for i in range(2 * len(Y)):
        ws.column_dimensions[L(C_Y0 + i)].width = 3.8
    put(ws, 1, 1, f"【架空】ひばり野クリーンセンター　中長期整備計画（{v}〜2030年度　廃止まで）", font(16, True))
    put(ws, 2, 1, f"{v}年度改訂（第{edition(v)}版）　{made(v):%Y/%m/%d}　ひばり野環境サービス株式会社 保全課　s2　　※2022年度までの計画・実績は第7版（旧様式）を参照",
        font(9))
    put(ws, 3, 1, "凡例　点：点検・測定　整：分解整備　補：補修　換：更新・取替　（市）：市発注の追加工事　灰色列＝定修のない炉", font(9))
    R1, R2, R3 = 5, 6, 7
    hd = fill("DDEBF7")
    for c, h in ((C_NO, "旧No"), (C_NAME, "機器名称"), (C_ASSET, "資産番号"), (C_IMP, "重要度"), (C_POL, "廃止まで")):
        ws.merge_cells(start_row=R1, start_column=c, end_row=R3, end_column=c)
        put(ws, R1, c, h, font(9, True), CENTER, B_ALL, hd)
    ws.merge_cells(start_row=R1, start_column=C_LAST1, end_row=R2, end_column=C_LAST2)
    put(ws, R1, C_LAST1, f"前回実施（〜{v - 1}）", font(8, True), CENTER, B_ALL, hd)
    put(ws, R3, C_LAST1, "1号/共通", font(7, True), CENTER, B_ALL, hd)
    put(ws, R3, C_LAST2, "2号", font(7, True), CENTER, B_ALL, hd)
    for i, y in enumerate(Y):
        c = C_Y0 + 2 * i
        ws.merge_cells(start_row=R1, start_column=c, end_row=R1, end_column=c + 1)
        put(ws, R1, c, str(y), font(9, True), CENTER, B_ALL, fill("FFE699") if y == v else hd)
        ws.merge_cells(start_row=R2, start_column=c, end_row=R2, end_column=c + 1)
        put(ws, R2, c, f"R{y - 2018}", font(7), CENTER, B_ALL, hd)
        put(ws, R3, c, "1号", font(7, True), CENTER, B_ALL, hd)
        put(ws, R3, c + 1, "2号", font(7, True), CENTER, B_ALL, hd)
    ws.merge_cells(start_row=R1, start_column=C_NOTE, end_row=R3, end_column=C_NOTE)
    put(ws, R1, C_NOTE, "備考（状態・根拠）", font(9, True), CENTER, B_ALL, hd)
    box(ws, R1, 1, R3, C_NOTE)
    r = R3 + 1
    oldno = 0
    curcat = None
    money = {}
    notes_v = dict(NOTE2)
    for vv in range(2024, v + 1):
        notes_v.update(NOTE2_LATER.get(vv, {}))
    for cat, name, eqs in ordered():
        if cat != curcat:
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=C_NOTE)
            put(ws, r, 1, f"{W.CATS[cat]}", font(10, True), LEFT, None, fill("E2EFDA"))
            box(ws, r, 1, r, C_NOTE)
            R.add("整備計画", r, "parent", {"path": [W.CATS[cat]]})
            r += 1
            curcat = cat
            gstart = r
        e0 = eqs[0]
        oldno += 1
        thisno = oldno
        oldno += len(eqs) - 1
        put(ws, r, C_NO, thisno, font(7, color="808080"), CENTER)
        put(ws, r, C_NAME, name, font(9), LEFT)
        put(ws, r, C_ASSET, " / ".join(W.asset_no(e, v - 1) for e in eqs), font(7), CENTER)
        put(ws, r, C_IMP, e0.imp, font(8), CENTER)
        pol = e0.policy.replace("（補修で継続）", "")
        put(ws, r, C_POL, pol, font(7), LEFT)
        lasts = []
        for j, e in enumerate(eqs):
            last = None
            for y in range(v - 1, 2015, -1):
                a = W.actual_work(e, y)
                if a and (a - {"I"}):
                    k = [x for x in "XUORI" if x in a][0]
                    last = f"{str(y)[2:]}{SYM2[k]}"
                    break
            lasts.append(last)
            put(ws, r, C_LAST1 + j, last, font(7), CENTER)
        if len(eqs) == 1:
            ws.merge_cells(start_row=r, start_column=C_LAST1, end_row=r, end_column=C_LAST2)
        marks = [{} for _ in eqs]
        for i, y in enumerate(Y):
            c = C_Y0 + 2 * i
            if len(eqs) == 1:
                ws.merge_cells(start_row=r, start_column=c, end_row=r, end_column=c + 1)
            for j, e in enumerate(eqs):
                pw = W.plan_work(e, y, v)
                key = next((k for k in "XUOR" if k in pw), "I" if "I" in pw else None)
                txt = SYM2[key] if key else None
                if (e.name, e.f, y) in W.CITY_WORKS and txt:
                    txt = txt + "市"
                cell = put(ws, r, c + j, txt, font(8, True if key in ("X", "O", "R") else False), CENTER)
                if txt and txt[0] in FILL2:
                    cell.fill = fill(FILL2[txt[0]])
                if txt:
                    marks[j][str(y)] = txt
                money[(y, e.f)] = money.get((y, e.f), 0) + (0 if (txt or "").endswith("市") else cost(e, pw))
        note = notes_v.get(name)
        put(ws, r, C_NOTE, note, font(7), LEFT)
        box(ws, r, 1, r, C_NOTE, B_HAIR)
        for i, y in enumerate(Y):
            ws.cell(row=r, column=C_Y0 + 2 * i).border = Border(left=thin, right=hair, top=hair, bottom=hair)
        ws.row_dimensions[r].height = 13 if not note else 22
        for j, e in enumerate(eqs):
            R.add("整備計画", r, "item", {"path": [W.CATS[cat], name], "code": W.asset_no(e, v - 1),
                                      "号炉": {0: "共通", 1: "1号", 2: "2号"}[e.f], "重要度": e.imp, "廃止まで": pol,
                                      "前回実施": lasts[j], "marks": marks[j], "memo": note})
        r += 1
        ws.row_dimensions.group(gstart, r - 1, outline_level=1, hidden=False)
    r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=C_NOTE)
    put(ws, r, 1, "法定点検・定期測定（施設全体）　※旧様式から移記", font(10, True), LEFT, None, fill("E2EFDA"))
    box(ws, r, 1, r, C_NOTE)
    R.add("整備計画", r, "parent", {"path": ["法定点検・定期測定"]})
    r += 1
    for (nm, tgt, cyc, rule, conf, basis) in W.LEGAL:
        oldno += 1
        put(ws, r, C_NO, oldno, font(7, color="808080"), CENTER)
        put(ws, r, C_NAME, nm, font(8), LEFT)
        ws.merge_cells(start_row=r, start_column=C_ASSET, end_row=r, end_column=C_LAST2)
        put(ws, r, C_ASSET, f"{tgt}（{cyc}）", font(7), LEFT)
        pm = {}
        for i, y in enumerate(Y):
            c = C_Y0 + 2 * i
            on = rule(y)
            if nm == "精密機能検査" and y == 2029:
                on = False
            if on:
                pm[str(y)] = "点"
            if nm.startswith("ボイラ定期"):
                f_ = 1 if "1号" in tgt else 2
                put(ws, r, c + (f_ - 1), "点" if on else None, font(8), CENTER)
            else:
                ws.merge_cells(start_row=r, start_column=c, end_row=r, end_column=c + 1)
                put(ws, r, c, "点" if on else None, font(8), CENTER)
        memo = "29：廃止前年のため要協議" if nm == "精密機能検査" else (basis[:30] if conf != "高" else None)
        put(ws, r, C_NOTE, memo, font(7), LEFT)
        box(ws, r, 1, r, C_NOTE, B_HAIR)
        cy = int(cyc[0]) if cyc[0].isdigit() else None
        R.add("整備計画", r, "item", {"path": ["法定点検・定期測定", nm], "work": nm, "cycle": cy, "marks": pm, "memo": memo})
        r += 1
    r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=C_NOTE)
    put(ws, r, 1, "年度計画額（百万円・税抜、SPC 補修上限額内）", font(10, True), LEFT, None, fill("E2EFDA"))
    r += 1
    sums = {}
    for lab, f_ in (("1号炉", 1), ("2号炉", 2), ("共通", 0)):
        ws.merge_cells(start_row=r, start_column=C_NAME, end_row=r, end_column=C_LAST2)
        put(ws, r, C_NAME, lab, font(9), LEFT)
        cs = {}
        for i, y in enumerate(Y):
            c = C_Y0 + 2 * i
            ws.merge_cells(start_row=r, start_column=c, end_row=r, end_column=c + 1)
            val = round(money.get((y, f_), 0) / 1000, 1)
            sums[y] = round(sums.get(y, 0) + val, 1)
            put(ws, r, c, val, font(8), RIGHT, None, None, "0.0")
            cs[str(y)] = val
        box(ws, r, 1, r, C_NOTE, B_HAIR)
        R.add("整備計画", r, "item", {"path": ["年度計画額", lab], "cost": cs, "unit": "百万円"})
        r += 1
    ws.merge_cells(start_row=r, start_column=C_NAME, end_row=r, end_column=C_LAST2)
    put(ws, r, C_NAME, "計", font(9, True), CENTER, None, fill("FFF2CC"))
    for i, y in enumerate(Y):
        c = C_Y0 + 2 * i
        ws.merge_cells(start_row=r, start_column=c, end_row=r, end_column=c + 1)
        put(ws, r, c, sums[y], font(8, True), RIGHT, None, fill("FFF2CC"), "0.0")
    box(ws, r, 1, r, C_NOTE, B_HAIR)
    R.add("整備計画", r, "total", {"path": ["年度計画額", "計"], "cost": {str(y): sums[y] for y in Y}, "unit": "百万円"})
    r += 1
    ws.merge_cells(start_row=r, start_column=C_NAME, end_row=r, end_column=C_LAST2)
    put(ws, r, C_NAME, "（参考）市追加工事", font(8), LEFT)
    if v == 2023:
        ws.merge_cells(start_row=r, start_column=C_Y0, end_row=r, end_column=C_Y0 + 1)
        put(ws, r, C_Y0, "市予算", font(8), CENTER)
    box(ws, r, 1, r, C_NOTE, B_HAIR)
    end = r
    ws.freeze_panes = f"{L(C_Y0)}{R3 + 1}"
    ws.sheet_properties.outlinePr.summaryBelow = False
    page(ws, landscape=True, paper=8, fit_w=1, fit_h=0, area=f"A1:{L(C_NOTE)}{end}", titles=f"{R1}:{R3}",
         footer=f"&L中長期整備計画 {v}年度改訂（第{edition(v)}版）&R&P / &N")
    wc = wb.create_sheet("年度別概算費用")
    put(wc, 1, 1, "【架空】ひばり野クリーンセンター　中長期整備計画　年度別概算費用（計画）", font(13, True))
    put(wc, 2, 1, f"単位：千円（税抜）　{v}年度改訂　※SPC 年度補修上限額内の分。市追加工事は含まない", font(9))
    put(wc, 4, 1, "炉", font(9, True), CENTER, B_ALL, fill("D9E1F2"))
    wc.column_dimensions["A"].width = 22
    for i, y in enumerate(Y):
        put(wc, 4, 2 + i, f"{y}年度", font(9, True), CENTER, B_ALL, fill("D9E1F2"))
        wc.column_dimensions[L(2 + i)].width = 11
    tot = {}
    for k, (lab, f_) in enumerate((("1号炉", 1), ("2号炉", 2), ("共通", 0))):
        put(wc, 5 + k, 1, lab, font(9), LEFT, B_ALL)
        cs = {}
        for i, y in enumerate(Y):
            val = int(round(money.get((y, f_), 0), -1))
            tot[y] = tot.get(y, 0) + val
            put(wc, 5 + k, 2 + i, val, font(9), RIGHT, B_ALL, None, "#,##0")
            cs[str(y)] = val
        R.add("年度別概算費用", 5 + k, "item", {"path": [lab], "cost": cs, "unit": "千円"})
    put(wc, 8, 1, "合　計", font(9, True), CENTER, B_ALL, fill("FFF2CC"))
    for i, y in enumerate(Y):
        put(wc, 8, 2 + i, tot[y], font(9, True), RIGHT, B_ALL, fill("FFF2CC"), "#,##0")
    R.add("年度別概算費用", 8, "total", {"path": ["合計"], "cost": {str(y): tot[y] for y in Y}, "unit": "千円"})
    page(wc, landscape=True, paper=9, fit_w=1, fit_h=1)
    wb.save(path(fn))
    return fn


def build_all():
    out = []
    for v in range(2016, 2027):
        out.append(build_s1(v) if v <= 2022 else build_s2(v))
    return out
