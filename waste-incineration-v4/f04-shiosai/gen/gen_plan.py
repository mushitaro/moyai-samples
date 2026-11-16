# -*- coding: utf-8 -*-
"""中長期整備計画（星取表）
 旧：H27.3 コンサル策定の「整備スケジュール表」。t1（H26〜H29）が灰色塗り＋「済」、t2（H30〜）が赤字「実」で実績を書き足した。
     R2・R3 は追記なし。R4 は t2 が一部だけ書いた。当初計画の記号はそのまま残す（運営者の決定）。
 新：R5 t3。旧表をコピーして、過去年度の列を消し「前回の主な整備」「状態・所見」に要約、年度を 計画/実績 の 2 列に分けた。
     法定点検の欄を上に足した。R6 に管理番号の列を足した。"""
import datetime as dt
from collections import OrderedDict
import world as W, cw
from xl import *

OLD_YEARS = list(range(2014, 2034))  # H26〜H45
NEW_YEARS = list(range(2023, 2034))  # R5〜R15

GRAY = fill("BFBFBF")
def old_head(fy):
    if fy == 2019: return "R1\n(H31)"
    return "H%d" % (fy - 1988)

def build_old(original=False):
    fname = "【架空】しおさい町ごみ処理場_中長期整備計画_H27策定長寿命化計画_%s.xlsx" % ("コンサル策定版" if original else "コンサル様式t1t2追記")
    rec_open(fname, plan=True)
    wb = new_wb("【架空】しおさい町ごみ処理施設 長寿命化計画 整備スケジュール表", W.CONSUL.replace("【架空】", "") + "（【架空】）",
                dt.datetime(2015, 3, 20, 10, 0), None if original else "t2")
    ws = wb.active; ws.title = "整備スケジュール"
    ncol = 8 + len(OLD_YEARS) + 3
    wl = [5, 11, 22, 5, 6, 5, 5, 7] + [4.6] * len(OLD_YEARS) + [7, 7, 22]
    widths(ws, wl)
    put(ws, 1, 1, "【架空】しおさい町ごみ処理施設　長寿命化計画　整備スケジュール表", F(14, True))
    put(ws, 2, 1, "平成27年3月　%s" % W.CONSUL, F(9))
    put(ws, 3, 1, "凡例　●：更新　◎：整備（オーバーホール）　○：補修　△：点検・調査　　※定期点検整備業務委託による毎年の点検は表示を省略", F(9))
    if not original: put(ws, 3, 20, "灰色＋済＝実施済（H26〜H29 記入 t1）　赤字「実」＝実施（H30〜 記入 t2）", F(9, color="C00000"))
    # 見出し
    hdr = ["No", "設備区分", "機器名称", "号炉", "設置\n年度", "重要\n度", "健全\n度", "対策\n区分"]
    for i, h in enumerate(hdr, 1):
        merge(ws, 5, i, 6, i, h, F(9, True), AC, B_ALL, fill("D9E1F2"))
    merge(ws, 5, 9, 5, 8 + len(OLD_YEARS), "計　画　年　度", F(9, True), AC, B_ALL, fill("D9E1F2"))
    for j, fy in enumerate(OLD_YEARS):
        put(ws, 6, 9 + j, old_head(fy) if not original else "H%d" % (fy - 1988), F(8, True, "C00000" if (fy == 2019 and not original) else None), AC, B_ALL, fill("D9E1F2"))
    c_up, c_oh, c_nt = 9 + len(OLD_YEARS), 10 + len(OLD_YEARS), 11 + len(OLD_YEARS)
    merge(ws, 5, c_up, 5, c_oh, "概算費用(千円)", F(8, True), AC, B_ALL, fill("D9E1F2"))
    put(ws, 6, c_up, "更新", F(8, True), AC, B_ALL, fill("D9E1F2"))
    put(ws, 6, c_oh, "整備1回", F(8, True), AC, B_ALL, fill("D9E1F2"))
    merge(ws, 5, c_nt, 6, c_nt, "備考", F(9, True), AC, B_ALL, fill("D9E1F2"))
    if not original: comment(ws, 6, 9 + OLD_YEARS.index(2019), "元号が変わったので R1 と書き直した。H32以降は直していない。", "t2")
    ws.row_dimensions[6].height = 24

    r = 7; no = 0
    works = {}
    for w in W.WORKS:
        works.setdefault((w.mid, w.fy), []).append(w)
    grand = {fy: 0 for fy in OLD_YEARS}
    for sysc in W.SYS_ORDER:
        es = [e for e in W.EQ if e.sys == sysc]
        if not es: continue
        r0 = r
        sub = {fy: 0 for fy in OLD_YEARS}
        for e in es:
            if e.added_2015 and original: continue
            if e.added_2015:
                ntxt = "%d-2" % no  # t1 が H28 に行を挿入
            else:
                no += 1; ntxt = no
            ad = e.added_2015
            col = "0070C0" if ad else None
            put(ws, r, 1, ntxt, F(8, color=col), AC, B_HAIR)
            put(ws, r, 3, e.name if e.name != "灰出しコンベヤNo.1" else "灰出しコンベヤ No.1", F(9, color=col), AL, B_HAIR)
            put(ws, r, 4, "共通" if e.furnace == 0 else "%d号" % e.furnace, F(8, color=col), AC, B_HAIR)
            put(ws, r, 5, W.wareki(e.inst), F(8, color=col), AC, B_HAIR)
            put(ws, r, 6, e.imp, F(8, color=col), AC, B_HAIR)
            put(ws, r, 7, getattr(e, "health2014", "") if not ad else "", F(8), AC, B_HAIR)
            p = W.CPLAN[e.mid]
            rowworks = []
            kubun = "更新" if any(s == "●" for s, _ in p.values()) else ("延命" if p else "維持")
            put(ws, r, 8, "" if ad else kubun, F(8), AC, B_HAIR)
            up = W.plan_cost(e, "●") if any(s == "●" for s, _ in p.values()) else None
            oh = W.plan_cost(e, "◎") if any(s == "◎" for s, _ in p.values()) else (W.plan_cost(e, "△") if p else None)
            for j, fy in enumerate(OLD_YEARS):
                c = 9 + j
                cell = put(ws, r, c, None, F(9), AC, B_HAIR)
                sym = p.get(fy, (None, None))[0]
                if sym:
                    cell.value = sym
                    cost = W.plan_cost(e, sym)
                    sub[fy] += cost; grand[fy] += cost
                done = [w for w in works.get((e.mid, fy), []) if w.kind != "INSPECT" or sym == "△"]
                if not done or original: continue
                rowworks.extend(done)
                w0 = done[0]
                dsym = {"REPLACE": "●", "UPGRADE": "●", "OVERHAUL": "○", "INSPECT": "△"}[w0.kind]
                if fy <= 2017:  # t1
                    cell.value = (sym or dsym) + "済"
                    cell.fill = GRAY
                    if not sym: cell.font = F(9, color="0070C0")
                elif fy in (2018, 2019):  # t2
                    cell.value = (sym or dsym) + "実"
                    cell.font = F(8, True, "C00000")
                elif fy == 2022 and w0.contract in ("R4-MK", "R4-CH") and w0.kind != "INSPECT":
                    cell.value = (sym or dsym) + "実"
                    cell.font = F(8, True, "C00000")
            if ad:
                # t1 が追記：インバータ盤（費用は入れたが小計は直していない）
                fy = 2015 if e.furnace == 1 else 2016
                cell = ws.cell(row=r, column=9 + OLD_YEARS.index(fy))
                cell.value = "●済"; cell.fill = GRAY; cell.font = F(9, color="0070C0")
                up = 4800
            put(ws, r, c_up, up, F(8, color=col), AR, B_HAIR, fmt="#,##0")
            put(ws, r, c_oh, oh, F(8), AR, B_HAIR, fmt="#,##0")
            nt = ""
            if e.name in W.EN_SCOPE or e.name in ("中央監視操作盤", "データ処理装置"): nt = "長寿命化工事（H27・H28）"
            if e.name == "バグフィルタ": nt = "ろ布取替。H27/H28は長寿命化工事"
            if e.name == "ごみクレーン": nt = "予備機なし。H38 更新（要検討）"
            if e.name == "灰出しコンベヤNo.1": nt = "チェーン・スプロケット取替（H21取替）"
            if ad: nt = "H28 行追加（t1）長寿命化で新設"
            if e.name == "灰出しコンベヤNo.1": pass
            put(ws, r, c_nt, nt, F(8, color=col), AL, B_HAIR)
            marks = {str(fy): ws.cell(row=r, column=9 + j).value for j, fy in enumerate(OLD_YEARS) if ws.cell(row=r, column=9 + j).value}
            rec(ws.title, r, "item", {"path": [W.SYS[sysc], e.name], "code": str(ntxt), "号炉": "共通" if e.furnace == 0 else "%d号" % e.furnace,
                                      "設置年度": e.inst, "重要度": e.imp, "健全度": getattr(e, "health2014", None) if not ad else None, "work": "" if ad else kubun,
                                      "cost": {"更新": up, "整備1回": oh}, "unit": "千円", "marks": marks, "memo": nt}, rowworks)
            r += 1
        merge(ws, r0, 2, r - 1, 2, W.SYS[sysc], F(9, True), AC, B_ALL)
        # 小計（コンサル計算。t1 の追記行は入っていない）
        put(ws, r, 3, "小計（千円）", F(8, True), AR, B_ALL, fill("F2F2F2"))
        for c in list(range(1, 9)) + [c_up, c_oh, c_nt]:
            ws.cell(row=r, column=c).border = B_ALL; ws.cell(row=r, column=c).fill = fill("F2F2F2")
        for j, fy in enumerate(OLD_YEARS):
            put(ws, r, 9 + j, sub[fy] if sub[fy] else None, F(7), AR, B_ALL, fill("F2F2F2"), "#,##0")
        rec(ws.title, r, "subtotal", {"path": [W.SYS[sysc]], "unit": "千円", "marks": {str(fy): sub[fy] for fy in OLD_YEARS if sub[fy]}})
        r += 1
    put(ws, r, 3, "合計（千円）", F(9, True), AR, B_ALL, fill("D9E1F2"))
    for c in list(range(1, 9)) + [c_up, c_oh, c_nt]:
        ws.cell(row=r, column=c).border = B_ALL; ws.cell(row=r, column=c).fill = fill("D9E1F2")
    for j, fy in enumerate(OLD_YEARS):
        put(ws, r, 9 + j, grand[fy], F(7, True), AR, B_ALL, fill("D9E1F2"), "#,##0")
    rec(ws.title, r, "total", {"unit": "千円", "marks": {str(fy): grand[fy] for fy in OLD_YEARS if grand[fy]}})
    r += 2
    put(ws, r, 1, "※概算費用は平成26年度単価（税抜）。長寿命化工事分（H27・H28）は工事費内訳による。", F(8)); r += 1
    if not original: put(ws, r, 1, "※R2・R3 の実績は未記入（修繕執行管理表を参照）。R4 は委託・チェーン取替のみ記入。　t2", F(8, color="C00000"))
    ws.freeze_panes = "I7"
    page(ws, landscape=True, paper=8, fit_w=1, fit_h=0, titles="5:6", area="A1:%s%d" % (CL(c_nt), r))
    ws.oddFooter.center.text = "&P / &N"
    ws.oddHeader.right.text = "長寿命化計画 資料-3"
    return save(wb, fname)

# ---------------- 新（t3） ----------------
SYMFILL = {"更": "F4B084", "整": "FFD966", "補": "FFF2CC", "点": "BDD7EE"}

def prev_summary(e):
    ws_ = [w for w in W.works_of(e.mid) if w.fy <= 2022 and w.kind != "INSPECT"]
    if not ws_: return ""
    w = ws_[-1]
    t = "%s %s" % (W.wareki(w.fy), w.desc)
    if w.fy in (2020, 2021): t += "（管理表より）"
    return t

STATE = {
 ("ごみクレーン", 0): "R4年次：横行レール摩耗 軽微。制御部品 供給終了",
 ("ごみクレーン操作盤", 0): "PLC・インバータ 部品供給終了（R3通知）",
 ("ごみクレーン動力盤", 0): "同上",
 ("灰出しコンベヤNo.1", 0): "R4 チェーン新品。テール軸受も交換済",
 ("燃焼火格子", 1): "R4 片12枚交換。摩耗進行 1号＞2号",
 ("燃焼火格子", 2): "R4 摩耗測定 限度内",
 ("ガス冷却室", 2): "下部ホッパ 最小3.4mm（原厚6.0）",
 ("ガス冷却室", 1): "ノズルチップ2本交換。ホッパ 4.6mm",
 ("バグフィルタ", 1): "R4 ろ布全数取替",
 ("バグフィルタ", 2): "差圧上昇・破損2本（R4）",
 ("誘引通風機", 1): "振動 3.1mm/s 上昇傾向",
 ("誘引通風機", 2): "振動 1.8mm/s",
 ("噴射水ポンプA", 1): "メカシール 再漏れ（R1・R4）",
 ("噴射水ポンプB", 1): "H24整備以降なし。30年経過",
 ("空気圧縮機No.2", 0): "H7製 レシプロ。予備機",
 ("計量データ処理装置", 0): "OSサポート終了",
 ("高圧受電盤", 0): "H29 VCB取替。保護継電器 H7製",
 ("火格子駆動用油圧ユニット", 2): "R2 ポンプ取替。油分析 良",
}

def build_new():
    fname = "【架空】しおさい町ごみ処理場_中長期整備計画_R5〜R15_t3計画実績2列様式.xlsx"
    rec_open(fname, plan=True)
    wb = new_wb("【架空】しおさい町ごみ処理場 整備計画表（R5〜）", W.CONSUL.replace("【架空】", "") + "（【架空】）",
                dt.datetime(2015, 3, 20, 10, 0), "しおさい町 環境衛生課 t3")
    ws = wb.active; ws.title = "整備計画R5～"
    # 列：A No B 設備区分 C 機器名称 D 号炉 E 管理番号(R6追加) F 設置 G 重要度 H 健全度(非表示) I 対策区分(非表示)
    #     J 前回の主な整備 K 状態・所見  L.. 年度×2  概算 備考
    fixed = ["No", "設備区分", "機器名称", "号炉", "管理番号", "設置\n年度", "重要\n度", "健全\n度", "対策\n区分", "前回の主な整備", "状態・所見"]
    nf = len(fixed)
    yc0 = nf + 1
    c_cost = yc0 + 2 * len(NEW_YEARS); c_nt = c_cost + 1
    wl = [5, 11, 22, 5, 8, 5, 5, 5, 6, 24, 22] + [3.6, 3.6] * len(NEW_YEARS) + [7, 20]
    widths(ws, wl)
    ws.column_dimensions["H"].hidden = True
    ws.column_dimensions["I"].hidden = True
    put(ws, 1, 1, "【架空】しおさい町ごみ処理場　整備計画表（令和5年度〜令和15年度）", F(14, True))
    put(ws, 2, 1, "（元：長寿命化計画 整備スケジュール表 H27.3 %s をコピーして作成）　R5.3 作成／R8.4 見直し　t3" % W.CONSUL, F(8, color="808080"))
    put(ws, 3, 1, "記号　更：更新　整：整備（分解・オーバーホール）　補：補修（部分）　点：点検・調査（毎年の委託点検は書かない）　実績欄　済：実施　→：繰延べ（右の年度へ）　重要度C は原則 事後保全", F(9))
    for i, h in enumerate(fixed, 1):
        fl = fill("FCE4D6") if h == "管理番号" else fill("E2EFDA")
        merge(ws, 5, i, 6, i, h, F(9, True), AC, B_ALL, fl)
    comment(ws, 5, 5, "R6.4 台帳改訂に合わせて列を追加", "t3")
    for j, fy in enumerate(NEW_YEARS):
        c = yc0 + 2 * j
        lab = "R%d\n(%d)" % (fy - 2018, fy)
        if fy == 2026: lab = "R8\n(2026)上期"
        merge(ws, 5, c, 5, c + 1, lab, F(8, True), AC, B_ALL, fill("E2EFDA") if fy < 2033 else fill("D9D9D9"))
        put(ws, 6, c, "計", F(7, True), AC, B_ALL, fill("E2EFDA"))
        put(ws, 6, c + 1, "実", F(7, True), AC, B_ALL, fill("E2EFDA"))
    ws.row_dimensions[5].height = 30
    merge(ws, 5, c_cost, 6, c_cost, "概算\n(千円)", F(8, True), AC, B_ALL, fill("E2EFDA"))
    merge(ws, 5, c_nt, 6, c_nt, "備考", F(9, True), AC, B_ALL, fill("E2EFDA"))
    put(ws, 4, yc0 + 2 * len(NEW_YEARS) - 2, "R15末 廃止予定（広域化）", F(8, color="C00000"))

    r = 7
    # 法定点検（t3 追加）
    merge(ws, r, 1, r, c_nt, "■ 法定点検・測定（R5 追加。委託・役務で毎年度予算化するもの）", F(9, True), AL, B_ALL, fill("FFF2CC"))
    rec(ws.title, r, "parent", {"path": ["法定点検・測定"]})
    r += 1
    for (nm, cyc, law, f) in W.LEGAL:
        merge(ws, r, 2, r, 4, nm, F(8), AL, B_HAIR)
        put(ws, r, 5, cyc, F(7), AC, B_HAIR)
        merge(ws, r, 10, r, 11, law, F(7), AL, B_HAIR)
        for j, fy in enumerate(NEW_YEARS):
            c = yc0 + 2 * j
            put(ws, r, c, None, F(8), AC, B_HAIR); put(ws, r, c + 1, None, F(8), AC, B_HAIR)
            on = f(fy)
            if nm.startswith("精密機能検査") and fy == 2029: on = True  # t3 は次回を R11 と書いた
            if on:
                ws.cell(row=r, column=c).value = "○"
                if fy <= 2025 or (fy == 2026 and ("月次" in nm or "6か月" in cyc or "年2回" in cyc)):
                    ws.cell(row=r, column=c + 1).value = "済" if fy <= 2025 else "1回"
        if nm.startswith("精密機能検査"):
            put(ws, r, c_nt, "前回R4。次回R7→R11", F(7), AL, B_HAIR)
        mk_ = {}
        for j, fy in enumerate(NEW_YEARS):
            a_ = ws.cell(row=r, column=yc0 + 2 * j).value; b_ = ws.cell(row=r, column=yc0 + 2 * j + 1).value
            if a_ or b_: mk_[str(fy)] = "%s/%s" % (a_ or "", b_ or "")
        cyc_n = {"毎年": 1, "2年": 2, "3年": 3, "年1回": 1, "年2回": 0.5, "6か月": 0.5, "毎月": 1 / 12.0}.get(cyc)
        rec(ws.title, r, "item", {"path": ["法定点検・測定", nm], "cycle": cyc_n, "根拠": law, "memo": "前回R4。次回R7→R11" if nm.startswith("精密") else "", "周期": cyc, "marks": mk_})
        r += 1
    r += 1
    works = {}
    for w in W.WORKS:
        works.setdefault((w.mid, w.fy), []).append(w)
    no = 0
    totals = {fy: 0 for fy in NEW_YEARS}
    for sysc in W.SYS_ORDER:
        es = [e for e in W.EQ if e.sys == sysc]
        r0 = r
        for e in es:
            if e.added_2015: ntxt = "%d-2" % no
            else: no += 1; ntxt = no
            put(ws, r, 1, ntxt, F(8), AC, B_HAIR)
            put(ws, r, 3, e.name, F(9), AL, B_HAIR)
            put(ws, r, 4, "共通" if e.furnace == 0 else "%d号" % e.furnace, F(8), AC, B_HAIR)
            put(ws, r, 5, e.mid, F(7, color="C65911"), AC, B_HAIR)
            put(ws, r, 6, W.wareki(e.inst), F(8), AC, B_HAIR)
            put(ws, r, 7, e.imp, F(8, True), AC, B_HAIR)
            put(ws, r, 8, getattr(e, "health2014", ""), F(8), AC, B_HAIR)
            put(ws, r, 9, "", F(8), AC, B_HAIR)
            put(ws, r, 10, prev_summary(e), F(7), AL, B_HAIR)
            put(ws, r, 11, STATE.get((e.name, e.furnace), ""), F(7), AL, B_HAIR)
            p = W.T3PLAN[e.mid]
            # 実績
            done = {}
            for fy in NEW_YEARS:
                if fy > 2026: break
                ds = [w for w in works.get((e.mid, fy), []) if w.kind != "INSPECT" or p.get(fy) == "点"]
                if fy == 2026: ds = [w for w in ds if w.start and w.start <= dt.date(2026, 9, 30)]
                if ds: done[fy] = ds[0]
            first_cost = None
            for j, fy in enumerate(NEW_YEARS):
                c = yc0 + 2 * j
                s = p.get(fy)
                c1 = put(ws, r, c, s, F(8, True), AC, Border(left=thin, right=hair, top=hair, bottom=hair), fill(SYMFILL[s]) if s else None)
                c2 = put(ws, r, c + 1, None, F(7, color="C00000"), AC, Border(left=hair, right=thin, top=hair, bottom=hair))
                if s:
                    k = {"更": "●", "整": "◎", "補": "◎", "点": "△"}[s]
                    cost = W.plan_cost(e, k) or {"更": 800, "整": 300, "補": 250, "点": 200}[s]
                    if e.name in ("ごみクレーン操作盤", "ごみクレーン動力盤"): cost = 12500
                    if e.name == "ガス冷却室" and s == "補": cost = 1800
                    if e.name == "ガス冷却室" and s == "更": cost = 6800
                    if e.name == "燃焼火格子" and s == "補": cost = 2600
                    if fy >= 2026 and first_cost is None: first_cost = cost
                    totals[fy] += cost
                if fy in done:
                    c2.value = "済"
                elif s and fy <= 2025:
                    nxt = [y for y in done if y > fy]
                    c2.value = ("→" + W.wareki(nxt[0])) if nxt else "→"
            put(ws, r, c_cost, first_cost, F(7), AR, B_HAIR, fmt="#,##0")
            nt = ""
            if e.imp == "C" and not p: nt = "事後保全"
            if e.name == "ごみクレーン": nt = "予備なし。更新はしない（廃止まで）"
            if e.name in ("ごみクレーン操作盤",): nt = "R6 制御装置更新（動力盤と一括）"
            if e.name == "噴射水ポンプB" and e.furnace == 1: nt = "R6予算でA のみ。BはR7"
            if e.name == "空気圧縮機No.2": nt = "R7→R8 7月実施"
            if e.name == "ガス冷却室" and e.furnace == 2: nt = "R8 定修でホッパ部更新"
            if e.name == "誘引通風機" and e.furnace == 1: nt = "R8 定修で軸受交換"
            if e.added_2015: nt = "旧表 H28 追加行"
            put(ws, r, c_nt, nt, F(7), AL, B_HAIR)
            mk_ = {}
            for j, fy in enumerate(NEW_YEARS):
                a_ = ws.cell(row=r, column=yc0 + 2 * j).value; b_ = ws.cell(row=r, column=yc0 + 2 * j + 1).value
                if a_ or b_: mk_[str(fy)] = "%s/%s" % (a_ or "", b_ or "")
            rec(ws.title, r, "item", {"path": [W.SYS[sysc], e.name], "code": str(ntxt), "管理番号": e.mid, "号炉": "共通" if e.furnace == 0 else "%d号" % e.furnace,
                                      "設置年度": e.inst, "重要度": e.imp, "前回の主な整備": prev_summary(e), "状態・所見": STATE.get((e.name, e.furnace), ""),
                                      "cost": first_cost, "unit": "千円", "marks": mk_, "memo": nt,
                                      "comments": {}}, list(done.values()))
            r += 1
        merge(ws, r0, 2, r - 1, 2, W.SYS[sysc], F(9, True), AC, B_ALL)
    rec(ws.title, r, "total", {"unit": "千円", "marks": {str(fy): totals[fy] for fy in NEW_YEARS if totals[fy]}})
    put(ws, r, 3, "年度計（計画・千円）", F(8, True), AR, B_ALL, fill("E2EFDA"))
    for j, fy in enumerate(NEW_YEARS):
        c = yc0 + 2 * j
        merge(ws, r, c, r, c + 1, totals[fy], F(7, True), AR, B_ALL, fill("E2EFDA"))
        ws.cell(row=r, column=c).number_format = "#,##0"
    r += 2
    put(ws, r, 1, "※R2・R3 は旧表に実績の記入なし。前回の主な整備は修繕執行管理表から拾った（「管理表より」）。", F(8)); r += 1
    put(ws, r, 1, "※概算はR5単価（税抜）。法定点検は委託料で別計上のため年度計に入れない。", F(8)); r += 1
    put(ws, r, 1, "※概算費用は平成26年度単価（税抜）。長寿命化工事分（H27・H28）は工事費内訳による。", F(8, color="808080"))  # 旧表の注記が残ったまま
    ws.freeze_panes = "%s7" % CL(yc0)
    page(ws, landscape=True, paper=8, fit_w=1, fit_h=0, titles="5:6", area="A1:%s%d" % (CL(c_nt), r))
    ws.oddFooter.center.text = "&P / &N"
    ws.oddHeader.right.text = "長寿命化計画 資料-3"  # 旧表のヘッダが残っている
    return save(wb, fname)

if __name__ == "__main__":
    print(build_old()); print(build_new())
