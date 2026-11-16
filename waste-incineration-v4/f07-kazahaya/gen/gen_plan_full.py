# -*- coding: utf-8 -*-
"""中長期整備計画（事業契約の長期修繕計画 20 年の毎年度見直し版）
 旧：sp1 2019 年度版（2015–2019 年度に sp1 が使った様式。事業契約の様式をほぼそのまま）
 新：sp2 2026 年度版（2020 年度の 5 年目見直しで sp2 が旧様式をコピーして直した様式を、毎年度コピーして更新）
"""
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font
from world import *
from common import *
PLANLINK = []
ACTLINK = []

CATS = ["受入供給設備", "燃焼設備", "通風設備", "ボイラ設備", "排ガス処理設備", "灰出し設備", "余熱利用設備",
        "給水設備", "排水処理設備", "電気計装設備", "電気設備", "計装設備", "雑設備"]
UN = {"1": "1号炉", "2": "2号炉", "0": "共通"}
SYM1 = {"I": "○", "O": "◎", "R": "△", "X": "●", "L": "◇", "U": "★"}   # sp1
SYM2 = {"I": "○", "O": "◎", "R": "▲", "X": "●", "L": "☆", "U": "★"}   # sp2
PRIO = "LXROI"  # 1 セルに複数あるときの表示の強さ（左ほど強い…ただし X が最優先）
PRIO = "XROLI"

LEGAL_ROWS = [
    # (名称, 根拠・周期, 年度→記号 の関数)
    ("ばい煙測定（ばいじん・SOx・NOx・HCl）", "大気汚染防止法 2月に1回以上（炉別）", lambda y: True),
    ("ダイオキシン類測定（排ガス・飛灰処理物・主灰）", "ダイオキシン類対策特別措置法 年1回以上（炉別）", lambda y: True),
    ("ダイオキシン類作業環境測定", "安衛則 第592条の2 6月以内ごと", lambda y: True),
    ("消防用設備等点検", "消防法 機器点検6月・総合点検1年", lambda y: True),
    ("精密機能検査", "廃棄物処理法施行規則 3年に1回以上", lambda y: y in (2017, 2020, 2023, 2026, 2029, 2032)),
    ("受変電設備 月次点検", "保安規程（外部委託）毎月 ※年次は電気設備の各行", lambda y: True),
]


def cell_marks(ws_list, sym):
    ts = [t for (t, w, *_) in ws_list]
    for p in PRIO:
        if p in ts:
            return sym[p]
    return ""


def year_cost(e, works):
    return sum(cost_of(w, e) for (t, w, *_) in works if t in "XRO" or w.startswith("定期事業者検査（開放"))


def ordered():
    """大区分 → 号炉 → 機器（電動機は直後）の順に並べる。"""
    for cat in CATS:
        for u in ("1", "2", "0"):
            es = [e for e in ALL if e["cat"] == cat and e["unit"] == u]
            if es:
                yield cat, u, es


def title_block(ws, title, sub, ncol):
    ws["A1"] = title
    ws["A1"].font = font(16, True)
    ws["A2"] = sub
    ws["A2"].font = font(9)


# ================================================================ 旧 sp1 2019 年度版
def make_old(ver=2019):
    fname = f"{FN}_中長期整備計画_{ver}年度版_sp1長期修繕計画表.xlsx"
    set_plan(fname)
    wb = Workbook()
    setp(wb, f"【架空】かざはやエコプラザ 長期修繕計画表（{ver}年度見直し）", "【架空】かざはや環境パートナーズ sp1")
    SH = "長期修繕計画"
    ws = wb.active
    ws.title = "長期修繕計画"
    Y0 = 7  # 年度列の開始（G 列）
    ncol = Y0 + len(YEARS)  # 備考列
    ws["A1"] = "【架空】かざはやエコプラザ　長期修繕計画表（事業期間20年）"
    ws["A1"].font = font(16, True)
    ws["A2"] = (f"事業契約書 別紙12「長期修繕計画」の見直し版　{ver}年度見直し（第{ver - 2015}回）　作成：SPC保全 sp1　{ver}年4月"
                if ver > 2015 else "事業契約書 別紙12「長期修繕計画」　2015年度版（事業契約時の計画を転記）　作成：SPC保全 sp1　2015年4月")
    ws["A2"].font = font(9)
    ws["A3"] = "凡例　○：点検・測定　◎：分解整備　△：部分補修　●：更新　◇：法定検査　（費用は千円・税抜・概算、◎△●のみ計上）"
    ws["A3"].font = font(9)
    ws.cell(1, ncol - 3, "承認").font = font(9)
    for i, t in enumerate(["所長", "副所長", "保全"]):
        c = ws.cell(2, ncol - 3 + i, t)
        c.font = font(8)
        c.border = ALL_THIN
        c.alignment = C
        ws.cell(3, ncol - 3 + i).border = ALL_THIN
    ws.cell(3, ncol - 1, "sp1").font = font(8, color="C00000")
    hr = 5
    heads = ["大区分", "号炉", "中区分", "機器名", "ｺｰﾄﾞ", "周期"]
    for i, t in enumerate(heads, 1):
        ws.merge_cells(start_row=hr, start_column=i, end_row=hr + 1, end_column=i)
        c = ws.cell(hr, i, t)
    for k, y in enumerate(YEARS):
        ws.cell(hr, Y0 + k, f"{y}")
        ws.cell(hr + 1, Y0 + k, f"{k + 1}年目")
    ws.merge_cells(start_row=hr, start_column=ncol, end_row=hr + 1, end_column=ncol)
    ws.cell(hr, ncol, "備考")
    for r in (hr, hr + 1):
        for cidx in range(1, ncol + 1):
            c = ws.cell(r, cidx)
            c.font = font(8, True)
            c.alignment = C
            c.border = box(t=med if r == hr else thin, b=med if r == hr + 1 else thin)
            c.fill = fill("FFF2CC")
    for r in (hr, hr + 1):
        ws.cell(r, Y0 + YEARS.index(ver)).fill = fill("F4B084")
    r = hr + 2
    totals = {y: 0 for y in YEARS}
    for cat in CATS:
        groups = [(u, es) for (c_, u, es) in ordered() if c_ == cat]
        if not groups:
            continue
        # 大区分 行
        c = ws.cell(r, 1, cat)
        c.font = font(10, True)
        for cidx in range(1, ncol + 1):
            ws.cell(r, cidx).fill = fill("DDEBF7")
            ws.cell(r, cidx).border = box(t=thin, b=thin)
        rec(fname, "parent", {"path": [cat]}, r=r, sheet=SH)
        r += 1
        sub = {y: 0 for y in YEARS}
        for u, es in groups:
            ws.cell(r, 2, UN[u]).font = font(9, True)
            for cidx in range(2, ncol + 1):
                ws.cell(r, cidx).border = box(t=hair, b=hair)
            rec(fname, "parent", {"path": [cat, UN[u]]}, r=r, sheet=SH)
            r += 1
            prev_sub = None
            for e in es:
                plan = rule(e, False, ver)
                ws.cell(r, 3, e["sub"] if e["sub"] != prev_sub else "〃")
                prev_sub = e["sub"]
                nm = e["name"] if e["kind"] != "電動機" else "　電動機"
                ws.cell(r, 4, nm)
                ws.cell(r, 5, e["code"])
                # 周期（年）：最も短い ◎△● の間隔
                ys = sorted(y for y, w in plan.items() if any(t in "ORX" for (t, _) in w))
                cyc = min((b - a for a, b in zip(ys, ys[1:])), default="")
                if all(any(t == "L" for (t, _) in plan.get(y, [])) for y in range(2016, 2020)) and plan:
                    cyc = "毎年"
                ws.cell(r, 6, cyc)
                for k, y in enumerate(YEARS):
                    m = cell_marks(plan.get(y, []), SYM1)
                    if m:
                        ws.cell(r, Y0 + k, m)
                        sub[y] += year_cost(e, plan.get(y, []))
                for cidx in range(1, ncol + 1):
                    c = ws.cell(r, cidx)
                    c.font = font(8)
                    c.border = box(l=hair if cidx > 1 else thin, r=hair, t=hair, b=hair)
                    c.alignment = C if cidx >= Y0 or cidx in (2, 6) else CL
                if e["kind"] == "電動機":
                    ws.cell(r, 4).font = font(8, color="595959")
                # sp1 のメモ（その年の 4 月時点で分かっていたことだけ）
                comments = {}
                if e["code"] == "1NS-030" and ver >= 2018:
                    ws.cell(r, ncol, "2017摩耗測定 可動火格子 最大6.1mm 要注意" if ver == 2018 else
                            "2018摩耗測定 可動火格子 最大8.6mm（想定の約1.6倍）→次回見直しで要検討")
                    if ver == 2019:
                        cc = get_column_letter(Y0 + YEARS.index(2027)) + str(r)
                        ws[cc].comment = Comment("sp1: 事業契約時の更新年(12年目)のまま。2019定修の測定結果を見て判断。", "sp1")
                        comments[cc] = ws[cc].comment.text
                if e["code"] == "2NS-030" and ver == 2019:
                    ws.cell(r, ncol, "1号に同じ傾向か 2019測定で確認")
                if e["code"] == "0UK-010":
                    ws.cell(r, ncol, "計量法 2年毎")
                if e["code"] == "0YN-010":
                    ws.cell(r, ncol, "電気事業法 定期事業者検査 4年毎" + ("（2018実施済）" if ver >= 2019 else ""))
                ws.cell(r, ncol).font = font(8, color="C00000")
                ws.cell(r, ncol).alignment = CL
                marks = {str(y): ws.cell(r, Y0 + k).value for k, y in enumerate(YEARS) if ws.cell(r, Y0 + k).value}
                f = {"path": [cat, UN[u], e["sub"], e["name"]], "code": e["code"], "cycle": cyc if cyc != "" else None,
                     "marks": marks, "memo": ws.cell(r, ncol).value or ""}
                if comments:
                    f["comments"] = comments
                rec(fname, "item", f, r=r, sheet=SH)
                PLANLINK.append((fname, SH, r, e["code"], marks, "plan"))
                r += 1
        # 小計
        ws.cell(r, 4, f"{cat}　小計（千円）")
        for k, y in enumerate(YEARS):
            ws.cell(r, Y0 + k, sub[y] if sub[y] else None)
            ws.cell(r, Y0 + k).number_format = "#,##0"
            totals[y] += sub[y]
        rec(fname, "subtotal", {"path": [cat], "cost": {str(y): sub[y] for y in YEARS if sub[y]}, "unit": "千円"}, r=r, sheet=SH)
        for cidx in range(1, ncol + 1):
            c = ws.cell(r, cidx)
            c.font = font(8, True)
            c.border = box(t=thin, b=thin)
            c.fill = fill("F2F2F2")
            if cidx >= Y0:
                c.alignment = CR
        r += 1
    # 法定（台帳外）
    ws.cell(r, 1, "法定点検・測定（機器台帳外）").font = font(10, True)
    for cidx in range(1, ncol + 1):
        ws.cell(r, cidx).fill = fill("DDEBF7")
    r += 1
    for nm, basis, fn in LEGAL_ROWS:
        if nm.startswith("精密機能検査"):
            pass
        ws.cell(r, 4, nm)
        ws.cell(r, ncol, basis)
        for k, y in enumerate(YEARS):
            if fn(y):
                ws.cell(r, Y0 + k, "◇")
        rec(fname, "item", {"path": ["法定点検・測定（機器台帳外）", nm], "marks": {str(y): "◇" for y in YEARS if fn(y)}, "memo": basis},
            r=r, sheet=SH)
        for cidx in range(1, ncol + 1):
            c = ws.cell(r, cidx)
            c.font = font(8)
            c.border = box(t=hair, b=hair, l=hair, r=hair)
            c.alignment = C if cidx >= Y0 else CL
        r += 1
    ws.cell(r, 4, "合計（千円）")
    for k, y in enumerate(YEARS):
        ws.cell(r, Y0 + k, totals[y])
        ws.cell(r, Y0 + k).number_format = "#,##0"
    rec(fname, "total", {"path": ["合計"], "cost": {str(y): totals[y] for y in YEARS}, "unit": "千円"}, r=r, sheet=SH)
    for cidx in range(1, ncol + 1):
        c = ws.cell(r, cidx)
        c.font = font(8, True)
        c.border = box(t=dbl, b=med)
        if cidx >= Y0:
            c.alignment = CR
    last = r
    widths = [12, 6, 10, 26, 9, 5] + [5.2] * len(YEARS) + [34]
    for i, wdt in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = wdt
    ws.freeze_panes = f"{get_column_letter(Y0)}{hr + 2}"
    ws.print_title_rows = f"{hr}:{hr + 1}"
    ws.print_area = f"A1:{get_column_letter(ncol)}{last}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = 8
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.oddFooter.center.text = "&P / &N"
    ws.oddHeader.right.text = f"【架空】かざはやエコプラザ 長期修繕計画 {ver}年度見直し"
    wb.save(outpath(fname))
    return fname


# ================================================================ 新 sp2 2026 年度版
# 構造（2020 年度に sp2 が旧表をコピーして作り替えた）：
#  ・事業期間 20 年を 5 年ごとのシート（第1期〜第4期）に分割。第1期（2015–2019）は旧表の写しを隠しシートで残す。
#  ・各年度の下に「計画・実績・費用」の 3 段の見出し。計画＝2020 見直し後の計画（過去年度も当初計画を残す）、
#    実績＝PM-NAVI の実績記号、費用＝行ごとの概算（千円、◎▲●のみ）。
#  ・行の並び（大区分→号炉→中区分→機器→電動機）と左端の列は旧表のまま、重要度・保全方式を挿入。
#  ・費用集計シートは各期シートの小計を転記。
PERIODS = [("第1期(2015-2019)", range(2015, 2020)), ("第2期(2020-2024)", range(2020, 2025)),
           ("第3期(2025-2029)", range(2025, 2030)), ("第4期(2030-2034)", range(2030, 2035))]
NOTE2_V = {  # 機器コード: [(この版から, 文)]（後の版ほど後ろ。版の作成時点で分かっていたことだけ）
    "1NS-030": [(2020, "2020見直しで●2027→2023へ前倒し（2018測定 摩耗速度 約1.6倍）"), (2024, "2020見直しで●2027→2023へ前倒し。2023.10更新済（192枚）")],
    "2NS-030": [(2020, "2020見直しで●2028→2024"), (2025, "2020見直しで●2028→2024。2024.6更新済")],
    "1NS-020": [(2020, "2020見直し ●2029→2027"), (2025, "2020見直し ●2029→2027（2024測定 最大11.8mm）")],
    "1KT-100": [(2022, "2021 振動上昇（4.6mm/s）臨時点検"), (2024, "2023 軸受交換・インペラ肉盛。●インペラ2029（肉盛2回目不可）")],
    "1HG-380": [(2026, "予備層充填 2025実施。次回性能試験2026（施工中）")],
    "2HG-380": [(2026, "予備層充填 2026.6実施")],
    "1BO-070": [(2022, "下段管 最小6.3mm(2021) 限界4.5mm"), (2024, "下段管 最小5.6mm(2023) 限界4.5mm 2027部分更新"),
                (2026, "下段管 最小5.1mm(2025) 限界4.5mm 2027部分更新")],
    "0YN-010": [(2020, "電気事業法 定期事業者検査 4年毎（2018実施済）"), (2023, "2022 定期事業者検査 異常なし。2026 初の分解精密点検"),
                (2026, "2026 分解精密点検（初）施工中 10/19～11/20 定期事業者検査併施")],
    "0YN-300": [(2023, "2022 異音 臨時点検 → 2023 減速機OH"), (2024, "2023 異音→減速機OH（2022臨時点検）")],
    "0UK-010": [(2020, "計量法 2年毎（2020は実施、旧表の抜けを訂正）")],
    "0KS-010": [(2025, "部品供給2029終了予告→DCS更新2030（●2032から前倒し）")],
    "0DE-270": [(2024, "機関OH 2026（6年毎、2027から前倒し）"), (2026, "2026 機関OH 全停で施工中")],
    "2BO-180": [(2023, "2022予定→2023へ延期（部品納期）")],
}


def note2(code, ver):
    t = [x for (v, x) in NOTE2_V.get(code, []) if v <= ver]
    return t[-1] if t else ""


def hoshiki2(e, ver):
    r = rule(e, True, ver)
    ts = {t for ws in r.values() for (t, w) in ws}
    if not r:
        return "BDM"
    if "L" in ts:
        return "法定+TBM" if ("O" in ts or "X" in ts) else "法定"
    names = {w for ws in r.values() for (t, w) in ws}
    if any(("測定" in n or "試験" in n or "診断" in n) for n in names):
        return "CBM"
    return "TBM"


def make_new(ver=2026):
    fname = f"{FN}_中長期整備計画_{ver}年度版_sp2星取表.xlsx"
    set_plan(fname)
    wb = Workbook()
    setp(wb, f"【架空】かざはやエコプラザ 長期修繕計画表（{ver}年度見直し）", "【架空】かざはや環境パートナーズ sp2")
    hozen_hoshiki = lambda e: hoshiki2(e, ver)
    shown = lambda y: y < ver or (ver == 2026 and y == 2026)
    wb.remove(wb.active)
    GRAY = fill("D9D9D9")
    YEL = fill("FFFF99")
    subtot_all = {}   # (cat, year) -> 小計（シートに書いた値）
    grand = {}
    for pname, yrs in PERIODS:
        yrs = list(yrs)
        ws = wb.create_sheet(pname)
        if pname.startswith("第1期"):
            ws.sheet_state = "hidden"
        Y0 = 9
        ncol = Y0 + 2 * len(yrs)  # 備考
        ws["A1"] = f"【架空】かざはやエコプラザ　長期修繕計画表（事業期間20年）　{pname}"
        ws["A1"].font = font(16, True)
        ws["A2"] = (f"事業契約書 別紙12「長期修繕計画」の見直し版　{ver}年度見直し（第{ver - 2015}回）　作成：SPC保全 sp2　{ver}年4月"
                    + ("（2026.11.13 上期実績反映）" if ver == 2026 else ""))
        ws["A2"].font = font(9)
        ws["A3"] = ("凡例　○：点検・測定　◎：分解整備　▲：部分補修　●：更新　☆：法定検査　計画＝2020見直し後の計画（当初）"
                    + ("　実績＝PM-NAVI 実績（灰色＝年度確定、黄＝2026下期 施工中・未反映）" if ver == 2026 else "　実績＝PM-NAVI 実績（灰色）")
                    + "　※費用は「費用集計」シート（千円）")
        ws["A3"].font = font(9)
        if pname.startswith("第1期"):
            ws["A4"] = "※旧表（sp1 2019年度版）の写し。2020見直しで本シートは非表示。計画列は旧表の記号のまま。"
        else:
            ws["A4"] = "※2020年度 5年目見直しで火格子・誘引通風機・触媒・分析計の更新年を見直し（実績摩耗に合わせ前倒し）"
        ws["A4"].font = font(8, color="C00000")
        # 承認欄は備考列の上へ移した（2020 改訂で列が増え、旧位置では凡例と重なるため）
        ws.cell(1, ncol, "承認　　所長 ＿＿＿＿　　保全 sp2").font = font(9)
        ws.cell(1, ncol).border = ALL_THIN
        hr = 5
        heads = ["大区分", "号炉", "重要度", "保全方式", "中区分", "機器名", "ｺｰﾄﾞ", "周期"]
        for i, t in enumerate(heads, 1):
            ws.merge_cells(start_row=hr, start_column=i, end_row=hr + 2, end_column=i)
            ws.cell(hr, i, t)
        for k, y in enumerate(yrs):
            c0 = Y0 + 2 * k
            ws.merge_cells(start_row=hr, start_column=c0, end_row=hr, end_column=c0 + 1)
            ws.cell(hr, c0, f"{y}年度（R{y - 2018}）" if y >= 2019 else f"{y}年度（H{y - 1988}）")
            for j, t in enumerate(["計画", "実績"]):
                ws.merge_cells(start_row=hr + 1, start_column=c0 + j, end_row=hr + 2, end_column=c0 + j)
                ws.cell(hr + 1, c0 + j, t)
        ws.merge_cells(start_row=hr, start_column=ncol, end_row=hr + 2, end_column=ncol)
        ws.cell(hr, ncol, "備考")
        for r in range(hr, hr + 3):
            for cidx in range(1, ncol + 1):
                c = ws.cell(r, cidx)
                c.font = font(8, True); c.alignment = C
                c.border = box(t=med if r == hr else thin, b=med if r == hr + 2 else thin)
                c.fill = fill("FFF2CC")
        for k, y in enumerate(yrs):
            if y == ver:
                for j in range(2):
                    ws.cell(hr, Y0 + 2 * k + j).fill = fill("F4B084")
        r = hr + 3
        for cat in CATS:
            groups = [(u, es) for (c_, u, es) in ordered() if c_ == cat]
            ws.cell(r, 1, cat).font = font(10, True)
            for cidx in range(1, ncol + 1):
                ws.cell(r, cidx).fill = fill("DDEBF7")
                ws.cell(r, cidx).border = box(t=thin, b=thin)
            rec(fname, "parent", {"path": [cat]}, r=r, sheet=pname)
            r += 1
            sub = {y: 0 for y in yrs}
            for u, es in groups:
                ws.cell(r, 2, UN[u]).font = font(9, True)
                for cidx in range(2, ncol + 1):
                    ws.cell(r, cidx).border = box(t=hair, b=hair)
                rec(fname, "parent", {"path": [cat, UN[u]]}, r=r, sheet=pname)
                r += 1
                prev_sub = None
                for e in es:
                    old_rule = pname.startswith("第1期")
                    plan = rule(e, not old_rule, 2019 if old_rule else ver)
                    acts = actual(e)
                    ws.cell(r, 3, e["imp"]); ws.cell(r, 4, hozen_hoshiki(e))
                    ws.cell(r, 5, e["sub"] if e["sub"] != prev_sub else "〃")
                    prev_sub = e["sub"]
                    ws.cell(r, 6, e["name"] if e["kind"] != "電動機" else "　└電動機")
                    ws.cell(r, 7, e["code"])
                    ys = sorted(y for y, w in plan.items() if any(t in "ORX" for (t, _) in w))
                    cyc = min((b - a for a, b in zip(ys, ys[1:])), default="")
                    if plan and all(any(t == "L" for (t, _) in plan.get(y, [])) for y in range(2020, 2026)):
                        cyc = "毎年"
                    ws.cell(r, 8, cyc)
                    pend = e["unit"] == "1" or (e["unit"] == "0" and e["sys"] in ("YN", "DE"))
                    cmts = {}
                    for k, y in enumerate(yrs):
                        c0 = Y0 + 2 * k
                        pm = cell_marks(plan.get(y, []), SYM1 if old_rule else SYM2)
                        if pm:
                            ws.cell(r, c0, pm)
                        if shown(y):
                            a = acts.get(y, [])
                            am = cell_marks(a, SYM2)
                            ca = ws.cell(r, c0 + 1)
                            if am:
                                ca.value = am
                            if y < 2026 or not pend:
                                ca.fill = GRAY
                                if am:
                                    ACTLINK.append((fname, pname, r, e["code"], y, am))
                            elif a:
                                ca.fill = YEL
                            for (t, w, note) in a:
                                if note and note != "施工中":
                                    ca.comment = Comment(f"sp2: {w}（{note}）", "sp2")
                                    cmts[ca.coordinate] = ca.comment.text
                            cost = year_cost(e, a if (y < 2026 or not pend) else plan.get(y, []))
                        else:
                            cost = year_cost(e, plan.get(y, []))
                        sub[y] += cost
                    if not old_rule:
                        ws.cell(r, ncol, note2(e["code"], ver))
                        if e["code"] == "0YN-400" and 2030 in yrs:
                            ws.cell(r, Y0 + 2 * yrs.index(2030), "◇")  # 旧記号が 1 か所残る（コピーの跡）
                    ws.cell(r, ncol).font = font(8, color="C00000")
                    for cidx in range(1, ncol + 1):
                        c = ws.cell(r, cidx)
                        if cidx != ncol:
                            c.font = font(8)
                        lft = thin if cidx >= Y0 and (cidx - Y0) % 2 == 0 else hair
                        c.border = box(l=lft if cidx > 1 else thin, r=hair, t=hair, b=hair)
                        if cidx >= Y0 and cidx < ncol:
                            c.alignment = C
                        elif cidx in (2, 3, 4, 8):
                            c.alignment = C
                        else:
                            c.alignment = CL
                    if e["kind"] == "電動機":
                        ws.cell(r, 6).font = font(8, color="595959")
                    pmk = {str(y): ws.cell(r, Y0 + 2 * k).value for k, y in enumerate(yrs) if ws.cell(r, Y0 + 2 * k).value}
                    amk = {str(y): ws.cell(r, Y0 + 2 * k + 1).value for k, y in enumerate(yrs) if ws.cell(r, Y0 + 2 * k + 1).value}
                    f = {"path": [cat, UN[u], e["sub"], e["name"]], "code": e["code"], "importance": e["imp"],
                         "method": hozen_hoshiki(e), "cycle": cyc if cyc != "" else None, "marks": pmk, "marksActual": amk,
                         "memo": ws.cell(r, ncol).value or ""}
                    if cmts:
                        f["comments"] = cmts
                    rec(fname, "item", f, r=r, sheet=pname)
                    PLANLINK.append((fname, pname, r, e["code"], pmk, "plan"))
                    r += 1
            for y in yrs:
                v = sub[y]
                if cat == "燃焼設備" and y == 2027:  # 計画の欠け：乾燥火格子の前倒し分を費用集計へ入れ忘れ
                    v = sub[y] - COST["乾燥火格子更新"]
                subtot_all[(cat, y)] = v
                grand[y] = grand.get(y, 0) + sub[y]
        ws.cell(r, 1, "法定点検・測定（機器台帳外）").font = font(10, True)
        for cidx in range(1, ncol + 1):
            ws.cell(r, cidx).fill = fill("DDEBF7")
        r += 1
        for nm, basis, fn in LEGAL_ROWS:
            ws.cell(r, 6, nm)
            ws.cell(r, ncol, basis)
            for k, y in enumerate(yrs):
                if fn(y):
                    ws.cell(r, Y0 + 2 * k, "☆" if not pname.startswith("第1期") else "◇")
                    if y < min(ver, 2026):
                        ws.cell(r, Y0 + 2 * k + 1, "☆").fill = GRAY
            rec(fname, "item", {"path": ["法定点検・測定（機器台帳外）", nm],
                                "marks": {str(y): ws.cell(r, Y0 + 2 * k).value for k, y in enumerate(yrs) if ws.cell(r, Y0 + 2 * k).value},
                                "marksActual": {str(y): "☆" for k, y in enumerate(yrs) if ws.cell(r, Y0 + 2 * k + 1).value},
                                "memo": basis}, r=r, sheet=pname)
            for cidx in range(1, ncol + 1):
                c = ws.cell(r, cidx)
                c.font = font(8); c.border = box(t=hair, b=hair, l=hair, r=hair)
                c.alignment = C if Y0 <= cidx < ncol else CL
            r += 1
        if 2030 in yrs:
            ws.cell(r, 6, "DCS更新（一式・共通計上）")
            ws.cell(r, Y0 + 2 * yrs.index(2030), "●")
            ws.cell(r, ncol, f"一式 {DCS_TOTAL:,}千円（費用集計に計上）")
            rec(fname, "item", {"path": ["DCS更新（一式・共通計上）"], "marks": {"2030": "●"}, "memo": ws.cell(r, ncol).value}, r=r, sheet=pname)
            for cidx in range(1, ncol + 1):
                ws.cell(r, cidx).font = font(8); ws.cell(r, cidx).border = box(t=hair, b=hair)
            r += 1
        for y in yrs:
            grand[("T", y)] = grand.get(y, 0) + (DCS_TOTAL if y == 2030 else 0)
        r += 1
        last = r - 1
        widths = [12, 6, 5, 7, 10, 26, 9, 5] + [4.6, 4.6] * len(yrs) + [44]
        for i, wdt in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = wdt
        ws.freeze_panes = f"{get_column_letter(Y0)}{hr + 3}"
        ws.print_title_rows = f"{hr}:{hr + 2}"
        ws.print_area = f"A1:{get_column_letter(ncol)}{last}"
        ws.page_setup.orientation = "landscape"
        ws.page_setup.paperSize = 8
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.oddFooter.center.text = "&A  &P / &N"
        ws.oddHeader.right.text = f"【架空】かざはやエコプラザ 長期修繕計画 {ver}年度見直し"
    # 費用集計シート
    ws2 = wb.create_sheet("費用集計")
    ws2["A1"] = "【架空】かざはやエコプラザ　長期修繕計画　年度別費用（千円・税抜）"
    ws2["A1"].font = font(13, True)
    ws2["A2"] = f"各期シートの星取（計画・実績）を概算単価で積算（2020–{ver - 1}は実績、{ver}以降は計画。契約額ではない）"
    ws2["A2"].font = font(8)
    yrs = list(range(2020, 2035))
    ws2.cell(4, 1, "大区分")
    for k, y in enumerate(yrs):
        ws2.cell(4, 2 + k, y)
    for cidx in range(1, 2 + len(yrs)):
        c = ws2.cell(4, cidx); c.font = font(9, True); c.alignment = C; c.border = ALL_THIN; c.fill = fill("FFF2CC")
    rr = 5
    for cat in CATS + ["DCS更新"]:
        ws2.cell(rr, 1, cat).font = font(9)
        ws2.cell(rr, 1).border = ALL_THIN
        for k, y in enumerate(yrs):
            v = subtot_all.get((cat, y)) if cat != "DCS更新" else (DCS_TOTAL if y == 2030 else None)
            c = ws2.cell(rr, 2 + k, v or None)
            c.number_format = "#,##0"; c.font = font(9); c.border = ALL_THIN
            if y < ver:
                c.fill = GRAY
        rec(fname, "subtotal", {"path": [cat], "cost": {str(y): ws2.cell(rr, 2 + k).value for k, y in enumerate(yrs) if ws2.cell(rr, 2 + k).value},
                                "unit": "千円"}, r=rr, sheet="費用集計")
        rr += 1
    ws2.cell(rr, 1, "合計").font = font(9, True)
    ws2.cell(rr, 1).border = box(t=dbl)
    for k, y in enumerate(yrs):
        c = ws2.cell(rr, 2 + k, grand[("T", y)])
        c.number_format = "#,##0"; c.font = font(9, True); c.border = box(t=dbl)
    rec(fname, "total", {"path": ["合計"], "cost": {str(y): grand[("T", y)] for y in yrs}, "unit": "千円"}, r=rr, sheet="費用集計")
    ws2.column_dimensions["A"].width = 16
    for k in range(len(yrs)):
        ws2.column_dimensions[get_column_letter(2 + k)].width = 9.5
    ws2.page_setup.orientation = "landscape"; ws2.page_setup.paperSize = 9
    ws2.page_setup.fitToWidth = 1; ws2.sheet_properties.pageSetUpPr.fitToPage = True
    wb.active = 1
    wb.save(outpath(fname))
    return fname


def run():
    out = []
    for v in range(2015, 2020):
        out.append(make_old(v))
    for v in range(2020, 2027):
        out.append(make_new(v))
    return out


if __name__ == "__main__":
    import time
    t = time.time()
    for f in run():
        print(f)
    print(time.time() - t)
