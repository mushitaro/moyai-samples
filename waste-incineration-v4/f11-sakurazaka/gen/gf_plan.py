# -*- coding: utf-8 -*-
"""中長期整備計画：l1「廃止までの整備計画」R1.12 策定（R2〜R9 星取）→ l2 R5.4 改訂（各年度 計画・実績・費用の 3 段、当初の列は隠して残す）"""
import os
from openpyxl import Workbook
from openpyxl.comments import Comment
from common import *
import world as W
import plan_data as P

FN_OLD = "【架空】さくら坂焼却場_中長期整備計画_廃止までの整備計画R2-R9_l1様式.xlsx"
FN_NEW = "【架空】さくら坂焼却場_中長期整備計画_廃止までの整備計画R5改訂_l2様式.xlsx"
FN_R8 = "【架空】さくら坂焼却場_中長期整備計画_廃止までの整備計画R5改訂R8実績追記_l2様式.xlsx"
MEMO_R8 = {"2F-07": "R7.10 羽根車取替済（1.6mm/s）", "2B-03": "R7 補修工事で 30 枚取替済", "1F-07": "R6 軸受交換済",
           "0G-08": "R6 パドル取替済（補修工事）", "2C-12": "R6 当て板済", "1B-03": "R5 補修工事 72 枚取替済、予備 18 は R7 2号へ",
           "0A-07": "R5 走行車輪取替済。R7 性能検査 合格", "1C-01": "R8 測定 最小 2.50mm、廃止まで測定のみ"}
YEARS = P.YEARS
MARK_COLOR = {"●": "C00000", "◎": "1F4E79", "△": "C55A11", "※": "7030A0", "☆": "375623", "○": "000000"}
LEGEND = "凡例　○点検　◎分解整備　△補修　●取替・更新　※部品購入　☆法定検査　　（定期整備業務委託の範囲は○◎、補修工事・購入は別契約）"

def last_record(e):
    """l1 が R1 に紙の記録から書いた「前回整備」"""
    tbl = {"1B-03": "H22 全数更新", "2B-03": "H23 全数更新", "1D-05": "H26 ろ布", "2D-05": "H26 ろ布",
           "1C-01": "R1 肉厚測定", "2C-01": "R1 肉厚測定", "0A-07": "R1 年次・性能", "0L-01": "H22 更新",
           "2F-07": "H23 羽根車", "1F-07": "H22 羽根車", "0G-08": "H29 パドル", "0E-01": "H27 開放"}
    if e.no in tbl:
        return tbl[e.no]
    if e.name in P.YEARLY:
        return "R1 定修"
    if e.renew.startswith("H2"):
        return e.renew.split()[0]
    return ""

def thinking(e):
    n = e.name
    if e.no in P.SPECIFIC_L1 or n in P.YEARLY:
        return "毎年点検"
    if e.unit == "共通":
        d = P.COMMON_PLAN.get(n)
        if d is None:
            return "状態監視"
        return "個別" if d else "状態監視"
    if n in P.ALT:
        return "隔年点検"
    if n in P.OH_ONCE:
        return "分解 1 回"
    return "状態監視"

def rows_in_order():
    out = []
    for cat in W.CAT_ORDER:
        names = []
        for e in W.EQUIP:
            if e.cat == cat and e.unit in ("1", "2") and e.name not in names:
                names.append(e.name)
        grp = []
        for n in names:
            grp.append([W.find("1", n), W.find("2", n)])
        for e in W.EQUIP:
            if e.cat == cat and e.unit == "共通":
                grp.append([e])
        out.append((cat, grp))
    return out


def put_mark(ws, r, c, m, size=10):
    if not m:
        return
    cell(ws, "%s%d" % (get_column_letter(c), r), m, F(size, True, MARK_COLOR.get(m, "000000")), CENTER)


def build_old():
    wb = Workbook()
    ws = wb.active
    ws.title = "整備計画"
    cell(ws, "A1", "【架空】さくら坂焼却場　廃止までの整備計画（令和2年度～令和9年度）", F(14, True))
    cell(ws, "A2", "令和元年12月　環境部 施設管理課（l1）　／　廃止予定：令和9年度末（2号炉は令和9年9月末停止）", F(9))
    cell(ws, "A3", LEGEND, F(9))
    heads = ["項番", "機器名称", "号炉", "台数", "台帳No", "整備の\n考え方", "前回整備\n（記録）"] + ["R%d" % (y - 2018) for y in YEARS] + ["備考"]
    widths = [5, 24, 5, 5, 7, 8, 11] + [5] * 8 + [36]
    R0 = 5
    for i, h in enumerate(heads):
        col = get_column_letter(1 + i)
        cell(ws, "%s%d" % (col, R0), h, F(9, True), CENTER, B_ALL, fill("DDEBF7"))
        ws.column_dimensions[col].width = widths[i]
    ws.row_dimensions[R0].height = 28
    r = R0 + 1
    YC = 8  # R2 の列
    for ci, (cat, grp) in enumerate(rows_in_order(), 1):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=16)
        cell(ws, "A%d" % r, "%d　%s" % (ci, W.CAT[cat]), F(10, True), LEFT, B_ALL, fill("E7E6E6"))
        r += 1
        for gi, g in enumerate(grp, 1):
            r_start = r
            for e in g:
                cell(ws, "A%d" % r, "%d-%d" % (ci, gi) if e is g[0] else None, F(8), CENTER, B_HAIR)
                cell(ws, "B%d" % r, e.name if e is g[0] else None, F(9), LEFT, B_HAIR)
                cell(ws, "C%d" % r, e.unit if e.unit != "共通" else "共通", F(8), CENTER, B_HAIR)
                cell(ws, "D%d" % r, e.units, F(8), CENTER, B_HAIR)
                cell(ws, "E%d" % r, e.no, F(7), CENTER, B_HAIR)
                cell(ws, "F%d" % r, thinking(e), F(7), CENTER, B_HAIR)
                cell(ws, "G%d" % r, last_record(e), F(7), CENTER, B_HAIR)
                pl = P.l1_plan(e)
                for j, y in enumerate(YEARS):
                    c = cell(ws, "%s%d" % (get_column_letter(YC + j), r), None, F(9), CENTER, B_HAIR)
                    put_mark(ws, r, YC + j, pl.get(y))
                cell(ws, "P%d" % r, P.MEMO_L1.get(e.no, ""), F(8), LEFT, B_HAIR)
                r += 1
            if len(g) == 2:
                ws.merge_cells(start_row=r_start, start_column=1, end_row=r - 1, end_column=1)
                ws.merge_cells(start_row=r_start, start_column=2, end_row=r - 1, end_column=2)
    # 法定
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=16)
    cell(ws, "A%d" % r, "13　法定点検・検査（運転委託・測定委託の分を含む）", F(10, True), LEFT, B_ALL, fill("E2EFDA"))
    r += 1
    li = 0
    for key, name, cyc, basis, y1, ya, y2 in P.LEGAL:
        if y1 is None:
            continue   # l1 は第一種圧力容器（温水熱交換器）を入れていない
        li += 1
        cell(ws, "A%d" % r, "13-%d" % li, F(8), CENTER, B_HAIR)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
        cell(ws, "B%d" % r, name, F(9), LEFT, B_HAIR)
        cell(ws, "F%d" % r, cyc, F(7), CENTER, B_HAIR)
        cell(ws, "G%d" % r, "", F(7), CENTER, B_HAIR)
        for j, y in enumerate(YEARS):
            cell(ws, "%s%d" % (get_column_letter(YC + j), r), None, F(9), CENTER, B_HAIR)
            if y in y1:
                put_mark(ws, r, YC + j, "☆")
        cell(ws, "P%d" % r, basis, F(8), LEFT, B_HAIR)
        r += 1
    # 費用
    r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=16)
    cell(ws, "A%d" % r, "年度別 概算費用（千円・税込）　※当初見込み。補修工事は上表 △● の合計", F(10, True), LEFT, B_ALL, fill("FCE4D6"))
    r += 1
    tot = {y: 0 for y in YEARS}
    for name in P.COST_ROWS:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
        cell(ws, "A%d" % r, name, F(9), LEFT, B_ALL)
        for j, y in enumerate(YEARS):
            v = P.COST_L1[name].get(y, 0)
            tot[y] += v
            cell(ws, "%s%d" % (get_column_letter(YC + j), r), v if v else "—", F(8), RIGHT, B_ALL, fmt="#,##0")
        cell(ws, "P%d" % r, "", F(8), LEFT, B_ALL)
        r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
    cell(ws, "A%d" % r, "計", F(9, True), CENTER, B_ALL, fill("F2F2F2"))
    for j, y in enumerate(YEARS):
        cell(ws, "%s%d" % (get_column_letter(YC + j), r), tot[y], F(8, True), RIGHT, B_ALL, fill("F2F2F2"), fmt="#,##0")
    cell(ws, "P%d" % r, "R9 は廃止年度（1号炉のみ縮小定修）", F(8), LEFT, B_ALL)
    ws.freeze_panes = "H6"
    page(ws, "portrait", 8, 1, 0, "5:5")
    props(wb, "【架空】さくら坂焼却場 廃止までの整備計画（R2〜R9）", "【架空】さくら坂市 施設管理課 l1")
    fn = os.path.join(OUT, FN_OLD)
    wb.save(fn)
    return fn


SUB_ROWS = {
    "1B-03": ("└ 火格子片 最終ロット購入（90 枚）", {2022: "※"}, {2022: 13068}, "朝凪炉工 R4.9 通知：R4.12 受注分で生産終了"),
    "2F-07": ("└ 羽根車 購入（受注生産 10 か月）", {2024: "※"}, {2024: 9900}, "R5.12 見積依頼予定。図面は朝凪炉工のみ保有"),
}

def build_new(asof=2023, actual_fn=None, done_fn=None, cost_act=None, legal_done26=()):
    """l2 R5.4 改訂：l1 のファイルをそのまま使い、年度 1 列の星取に実績を「済」と色で書き足し、
    改訂した年度の記号はセルの色とコメントで当初を残す。費用は下の年度計を予算・決算の 2 行に。"""
    from plan_data import l1_plan, actual, l2_plan
    if actual_fn:
        actual = actual_fn
    R8 = asof == 2026
    PAST_BASE = set(range(2020, 2023)) if not R8 else set(range(2020, 2026))
    from gen_ledger import supply
    wb = Workbook()
    ws = wb.active
    ws.title = "整備計画"
    cell(ws, "A1", "【架空】さくら坂焼却場　廃止までの整備計画（令和2年度～令和9年度）　R5.4 改訂" + ("（R8.11 実績追記）" if R8 else ""), F(14, True))
    cell(ws, "A2", "令和元年12月　環境部 施設管理課（l1）　／　令和5年4月改訂（l2）　廃止予定：令和9年度末（2号炉は令和9年9月末停止）", F(9))
    cell(ws, "A3", LEGEND, F(9))
    cell(ws, "A4", ("R2〜R4" if not R8 else "R2〜R8（R8 は 11 月までに済んだもの）") + "：実施済は記号に「済」＋緑、計画外の実施は黄、未実施は橙（コメントに理由）。" + ("R5〜R9" if not R8 else "まだの欄") + "：改訂した欄は黄（コメントに当初の記号）。", F(9, color="0000CC"))
    heads = ["項番", "機器名称", "号炉", "台数", "台帳No", "整備の\n考え方", "前回整備\n（記録）"] + ["R%d" % (y - 2018) for y in YEARS] + ["備考", "部品供給", "改訂メモ（R5.4 l2）" if not R8 else "改訂メモ（R5.4／R8.11 l2）"]
    widths = [5, 24, 5, 5, 7, 8, 11] + [5.5] * 8 + [30, 9, 36]
    R0 = 6
    for i, h in enumerate(heads):
        col = get_column_letter(1 + i)
        cell(ws, "%s%d" % (col, R0), h, F(9, True), CENTER, B_ALL, fill("FFF2CC") if i >= 16 else fill("DDEBF7"))
        ws.column_dimensions[col].width = widths[i]
    ws.row_dimensions[R0].height = 28
    YC = 8
    NC = len(heads)
    PAST = PAST_BASE
    G_DONE, G_EXTRA, G_MISS, G_REV = fill("E2EFDA"), fill("FFF2CC"), fill("FCE4D6"), fill("FFF2CC")
    def yearcells(r, old, act, pn, legal=False, l2d=None, past=None):
        past = PAST if past is None else past
        for j, y in enumerate(YEARS):
            col = get_column_letter(YC + j)
            c = cell(ws, "%s%d" % (col, r), None, F(8), CENTER, B_HAIR)
            if y in past:
                o = old.get(y) if (y <= 2022 or l2d is None) else l2d.get(y)
                o1 = old.get(y)
                a = act.get(y)
                if a:
                    c.value = a + "済"
                    c.font = F(8, True, MARK_COLOR.get(a, "000000"))
                    c.fill = G_DONE if o == a else G_EXTRA
                    if o1 != a:
                        c.comment = Comment("当初：%s" % (o1 or "なし") + ("／R5改訂：%s" % (o or "なし") if y >= 2023 and o != o1 else ""), "l2")
                elif o:
                    c.value = o
                    c.font = F(8, color="808080")
                    c.fill = G_MISS
                    c.comment = Comment("未実施（翌年度以降へ）", "l2")
            else:
                o, n = old.get(y), pn.get(y)
                if n:
                    c.value = n
                    c.font = F(9, True, MARK_COLOR.get(n, "000000"))
                if n != o:
                    c.fill = G_REV
                    c.comment = Comment("当初：%s" % (o or "なし"), "l2")
    r = R0 + 1
    for ci, (cat, grp) in enumerate(rows_in_order(), 1):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=NC)
        cell(ws, "A%d" % r, "%d　%s" % (ci, W.CAT[cat]), F(10, True), LEFT, B_ALL, fill("E7E6E6"))
        r += 1
        for gi, g in enumerate(grp, 1):
            r_start = r
            for e in g:
                cell(ws, "A%d" % r, "%d-%d" % (ci, gi) if e is g[0] else None, F(8), CENTER, B_HAIR)
                cell(ws, "B%d" % r, e.name if e is g[0] else None, F(9), LEFT, B_HAIR)
                cell(ws, "C%d" % r, e.unit, F(8), CENTER, B_HAIR)
                cell(ws, "D%d" % r, e.units, F(8), CENTER, B_HAIR)
                cell(ws, "E%d" % r, e.no, F(7), CENTER, B_HAIR)
                cell(ws, "F%d" % r, thinking(e), F(7), CENTER, B_HAIR)
                cell(ws, "G%d" % r, last_record(e), F(7), CENTER, B_HAIR)
                act = dict(actual(e)); pn = l2_plan(e)
                if e.no == "1B-03": act[2022] = "○"
                if e.no == "2F-07":
                    pn[2024] = "○"
                    if act.get(2024) == "※": act[2024] = "○"
                past = set(PAST)
                if R8 and done_fn and done_fn(e, 2026):
                    past.add(2026)
                yearcells(r, l1_plan(e), act, pn, l2d=l2_plan(e) if R8 else None, past=past)
                cell(ws, "P%d" % r, P.MEMO_L1.get(e.no, ""), F(7), LEFT, B_HAIR)
                sp = supply(e)
                spv = "中止" if "生産中止" in sp else ("長納期" if "長納期" in sp else "")
                cell(ws, "Q%d" % r, spv, F(7, color="C00000" if spv == "中止" else "000000"), CENTER, B_HAIR)
                memo = P.MEMO_L2.get(e.no, "")
                if R8 and e.no in MEMO_R8:
                    memo += "／R8.11：" + MEMO_R8[e.no]
                cst = P.COST_L2.get(e.no, {})
                if e.no in ("1B-03", "2F-07"):
                    cst = {k: v for k, v in cst.items() if not (e.no == "1B-03" and k == 2022) and not (e.no == "2F-07" and k == 2024)}
                if cst:
                    memo += "（" + "、".join("R%d %s %s千円" % (y - 2018, "決算" if y in PAST else "予算", format(v, ",")) for y, v in sorted(cst.items())) + "）"
                cell(ws, "R%d" % r, memo, F(7, color="0000CC"), LEFT, B_HAIR)
                r += 1
                if e.no in SUB_ROWS:
                    name, marks, cst2, cm = SUB_ROWS[e.no]
                    cell(ws, "A%d" % r, None, F(8), CENTER, B_HAIR)
                    cell(ws, "B%d" % r, "　" + name, F(8, color="0000CC"), LEFT, B_HAIR, fill("FFF2CC"))
                    for c in "CDEFG":
                        cell(ws, "%s%d" % (c, r), None, F(8), CENTER, B_HAIR, fill("FFF2CC"))
                    ws["C%d" % r].value = e.unit
                    yearcells(r, ({} if not R8 else {}), {y: m for y, m in marks.items() if y in PAST}, {y: m for y, m in marks.items() if y not in PAST}, l2d=(marks if R8 else None))
                    cell(ws, "P%d" % r, "", F(7), LEFT, B_HAIR)
                    cell(ws, "Q%d" % r, "", F(7), LEFT, B_HAIR)
                    y0, v0 = list(cst2.items())[0]
                    cell(ws, "R%d" % r, "R5.4 行追加。%s（R%d %s %s千円）" % (cm, y0 - 2018, "決算" if y0 in PAST else "予算", format(v0, ",")) + ("／R8.11：購入済・納入済" if R8 else ""), F(7, color="0000CC"), LEFT, B_HAIR)
                    r += 1
            if len(g) == 2:
                ws.merge_cells(start_row=r_start, start_column=1, end_row=r - 1, end_column=1)
                if not any(e.no in SUB_ROWS for e in g):
                    ws.merge_cells(start_row=r_start, start_column=2, end_row=r - 1, end_column=2)
    # 法定
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=NC)
    cell(ws, "A%d" % r, "13　法定点検・検査（運転委託・測定委託の分を含む）", F(10, True), LEFT, B_ALL, fill("E2EFDA"))
    r += 1
    li = 0
    for key, name, cyc, basis, y1, ya, y2 in P.LEGAL:
        added = y1 is None
        li += 1
        cell(ws, "A%d" % r, "13-%d" % li if not added else "13-追", F(8), CENTER, B_HAIR)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
        cell(ws, "B%d" % r, name, F(9, color="0000CC" if added else "000000"), LEFT, B_HAIR, fill("FFF2CC") if added else None)
        cell(ws, "F%d" % r, cyc, F(7), CENTER, B_HAIR)
        cell(ws, "G%d" % r, "", F(7), CENTER, B_HAIR)
        actl = {y: "☆" for y in ya if y in PAST}
        pastl = set(PAST)
        if R8 and key in legal_done26:
            actl[2026] = "☆"; pastl.add(2026)
        yearcells(r, {y: "☆" for y in (y1 or [])}, actl, {y: "☆" for y in y2 if y not in pastl}, l2d=({y: "☆" for y in y2} if R8 else None), past=pastl)
        cell(ws, "P%d" % r, basis, F(7), LEFT, B_HAIR)
        cell(ws, "Q%d" % r, "", F(7), LEFT, B_HAIR)
        memo = ""
        if added:
            memo = "R5.4 行追加。当初計画に無かったが運転委託が毎年受検していた"
            ws["B%d" % r].comment = Comment("R4 の監査で指摘。温水プール用 2 基とも第一種圧力容器（検査証あり）。l2", "l2")
        if key == "scale":
            memo = "当初は R3・R5…としていたが、前回 H30.11 のため R2・R4 に受検済。以後 R6・R8"
        if key == "boiler2":
            memo = "2号は R9.9 末停止のため R9 なし"
        cell(ws, "R%d" % r, memo, F(7, color="0000CC"), LEFT, B_HAIR)
        r += 1
    # 費用：予算・決算の 2 行
    r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=NC)
    cell(ws, "A%d" % r, "年度別 費用（千円・税込）　当初は概算 1 行 → R5.4 から「予算」「決算」の 2 行（R2〜R4 の予算は当初概算）", F(10, True), LEFT, B_ALL, fill("FCE4D6"))
    r += 1
    totb = {y: 0 for y in YEARS}; tota = {y: 0 for y in YEARS}
    for name in P.COST_ROWS:
        for k, lab in enumerate(("予算", "決算")):
            if k == 0:
                ws.merge_cells(start_row=r, start_column=1, end_row=r + 1, end_column=6)
                cell(ws, "A%d" % r, name, F(9), LEFT, B_ALL)
            cell(ws, "G%d" % r, lab, F(8), CENTER, B_ALL, fill("F2F2F2") if k else None)
            for j, y in enumerate(YEARS):
                if k == 0:
                    v = P.COST_L1[name].get(y, 0) if y <= 2022 else P.COST_L2Y[name].get(y, 0)
                    totb[y] += v
                else:
                    v = (cost_act or P.COST_ACT)[name].get(y) if y in PAST else None
                    if v: tota[y] += v
                cell(ws, "%s%d" % (get_column_letter(YC + j), r), v if v else ("—" if (k == 0 or y in PAST) else None), F(7), RIGHT, B_ALL, fill("F2F2F2") if k else None, fmt="#,##0")
            for c in "PQR":
                cell(ws, "%s%d" % (c, r), None, F(7), LEFT, B_ALL)
            r += 1
    for k, (lab, tot) in enumerate((("予算 計", totb), ("決算 計", tota))):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
        cell(ws, "A%d" % r, lab, F(9, True), CENTER, B_ALL, fill("F2F2F2"))
        for j, y in enumerate(YEARS):
            v = tot[y] + (P.TOTAL_TYPO.get(y, 0) if k == 1 else 0)
            cell(ws, "%s%d" % (get_column_letter(YC + j), r), v if v else None, F(7, True), RIGHT, B_ALL, fill("F2F2F2"), fmt="#,##0")
        cell(ws, "R%d" % r, ("R5 以降は予算要求額" if not R8 else "R5 以降は予算要求額（R8 は当初予算）") if k == 0 else ("決算は R4 まで" if not R8 else "決算は R7 まで（R8.11 追記）"), F(7, color="0000CC"), LEFT, B_ALL)
        r += 1
    ws.freeze_panes = "H7"
    page(ws, "portrait", 8, 1, 0, "6:6")
    props(wb, "【架空】さくら坂焼却場 廃止までの整備計画（R5 改訂%s）" % ("・R8.11 実績追記" if R8 else ""), "【架空】さくら坂市 施設管理課 l2")
    fn = os.path.join(OUT, FN_NEW if not R8 else FN_R8)
    wb.save(fn)
    return fn


def main():
    return [build_old(), build_new()]

if __name__ == "__main__":
    main()
