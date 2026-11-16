# -*- coding: utf-8 -*-
"""機器台帳：SPC 資産システム（SPC-AM）抽出＋Excel 加工。各年度末（2026 年度は 10 月末）の版。
2016・2017 年度末：s1 の旧様式。2018 年度末〜：s1 が「廃止まで」列・色分け・「更新・付替え」シートを足した様式。
2022 年度末〜：s2 が右端に「最終整備(FY)」列を足す（前回整備欄は s1 の記入のまま止まる）。"""
import datetime as dt
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter as L
import world as W
import works as WK
from xlhelp import *
from rec import Rec, path

KIND = {"I": "点検", "O": "分解整備", "R": "補修", "X": "更新", "U": "改良"}


def last_work(e, upto):
    for y in range(upto, 2015, -1):
        a = W.actual_work(e, y)
        if a and (a - {"I"}):
            k = [x for x in "XUORI" if x in a][0]
            return y, W.oh_month(e.f, y, e.seq), KIND[k]
    return None


def note_for(e, fy):
    n = ""
    if e.name.startswith("活性炭") and e.install.startswith("2010"):
        n = "H22 活性炭供給装置更新（市直営時代）"
    if e.name == "ボイラ給水ポンプA":
        if e.f == 1 and fy >= 2017:
            n = "2017.10 更新（隼ポンプ）旧番号 HB08-101"
        if e.f == 2 and fy >= 2018:
            n = "2018.06 更新（隼ポンプ）旧番号 HB08-201"
    if e.name == "DCS中央オペレータステーション":
        n = "HMI 更新 2019予定（提案書）" if fy < 2019 else "2019.10 HMI更新済。コントローラは更新中止（H30）"
    if e.name == "バグフィルタ":
        n = "ろ布 前回交換 不明（市直営時代）"
        if e.f == 1 and fy >= 2019:
            n = "ろ布全数交換 2019.05"
        if e.f == 2 and fy >= 2020:
            n = "ろ布全数交換 2020.05"
        if fy >= 2026 and e.f == 2:
            n = "ろ布全数交換 2020.05、2026.05"
        if fy >= 2026 and e.f == 1:
            n = "ろ布全数交換 2019.05、2026.10"
    if e.name.startswith("計量機") and fy >= 2020:
        n = f"2020.08 更新（ロードセル・指示計）旧番号 {e.no}"
    if e.name == "計量棟操作盤" and fy >= 2020:
        n = f"2020.08 更新 旧番号 {e.no}"
    if e.name == "触媒反応塔" and fy >= 2018:
        n = "全量更新中止（H30）。活性試験で判断" if e.f == 1 else ("全量更新中止（H30）" if fy < 2021 else "2021.10 1段目のみ交換")
    if e.name == "蒸気タービン":
        n = "開放点検 " + "・".join(str(y) for y in (2017, 2021, 2025) if y <= fy) + "（定期事業者検査）" if fy >= 2017 else ""
    if e.name == "煙突内筒" and e.f == 2 and fy >= 2021:
        n = "R3 市追加工事 下部腐食部 当て板補修"
    if e.name == "ボイラ本体" and e.f == 2 and fy >= 2023:
        n = "R5 市追加工事 前壁水管6本取替・肉盛12箇所" + ("。FY2025 隣接2本肉盛" if fy >= 2025 else "")
    if e.name == "過熱器" and e.f == 1 and fy >= 2024:
        n = "FY2024 UT 2次入口 外面減肉 最小3.6mm → FY2026 補修" + ("済" if fy >= 2026 else "")
    if e.name == "誘引通風機" and e.f == 1 and fy >= 2024:
        n = "FY2024 軸受交換・ランナ肉盛（振動対策）"
    return n


def fname(fy):
    lab = "2026年10月" if fy == 2026 else f"{fy}年度末"
    sty = "s1様式" if fy <= 2017 else ("2018年度改訂様式" if fy <= 2021 else "2018年度改訂様式_s2")
    return f"【架空】ひばり野クリーンセンター_機器台帳_{lab}_{sty}.xlsx"


def build(fy):
    fn = fname(fy)
    R = Rec(fn)
    new = fy >= 2018
    s2 = fy >= 2022
    out_d = dt.datetime(2026, 10, 30, 16, 5) if fy == 2026 else dt.datetime(fy + 1, 3, 30, 17, 20)
    wb = Workbook()
    ws = wb.active
    ws.title = "資産一覧"
    props(wb, f"【架空】ひばり野クリーンセンター 設備資産台帳（保全対象）{'2026年10月' if fy == 2026 else f'{fy}年度末'}",
          "SPC保全課 s2" if s2 else "SPC保全課 s1", out_d)
    put(ws, 1, 1, "【架空】ひばり野クリーンセンター　設備資産台帳（保全対象機器）", font(14, True))
    put(ws, 2, 1, f"ひばり野環境サービス㈱ 保全課　　資産システム SPC-AM 抽出（{out_d:%Y/%m/%d}）　抽出条件：資産区分=保全対象、状態=稼働中", font(9))
    put(ws, 3, 1, "※ 系列 1＝1号炉、2＝2号炉、0＝共通。 前回整備・備考は保全課で記入（システム外）", font(8, color="666666"))
    if new:
        put(ws, 3, 9, "H30.4 廃止決定（2030年度末）により「廃止まで」列を追加。灰色＝事後保全（故障時対応）", font(8, True, color="C00000"))
    heads = ["資産番号", "資産名称", "設備区分CD", "設備区分名", "系列", "型式・仕様", "メーカー", "設置年月", "法定区分", "重要度"]
    widths = [11, 26, 6, 14, 5, 34, 12, 9, 10, 6]
    if new:
        heads += ["廃止まで"]
        widths += [22]
    heads += ["予備機", "前回整備", "備考"]
    widths += [6, 16, 34 if new else 26]
    if s2:
        heads += ["最終整備(FY)"]
        widths += [9]
    NC = len(heads)
    R0 = 5
    for i, (h, w) in enumerate(zip(heads, widths), 1):
        col = "DDEBF7" if i <= 10 else ("FCE4D6" if h == "廃止まで" else ("E2EFDA" if h.startswith("最終") else "FFF2CC"))
        put(ws, R0, i, h, font(9, True), CENTER, B_ALL, fill(col))
        ws.column_dimensions[L(i)].width = w
    ws.row_dimensions[R0].height = 24
    if new:
        ws.cell(row=R0, column=11).comment = Comment("H30改訂で追加（廃止決定）。区分：法定点検継続／更新中止・補修で延命／更新なし（補修で継続）／事後保全", "s1")
    if s2:
        ws.cell(row=R0, column=NC).comment = Comment("2022～ 資産システムの作業実績から転記。前回整備欄は旧担当の記入のまま残している。", "s2")
    r = R0 + 1
    cur, cnt = None, 0
    gray = fill("E7E6E6")
    eqs = sorted(W.EQ, key=lambda e: e.no)

    def subtotal(rr, cat, n):
        put(ws, rr, 2, f"　{W.CATS[cat]}　計 {n} 台", font(9, True), LEFT)
        for c in range(1, NC + 1):
            ws.cell(row=rr, column=c).fill = fill("F2F2F2")
            ws.cell(row=rr, column=c).border = B_HAIR
        R.add("資産一覧", rr, "subtotal", {"設備区分名": W.CATS[cat], "台数": n})

    for e in eqs:
        if cur is not None and e.cat != cur:
            subtotal(r, cur, cnt)
            r += 1
            cnt = 0
        cur = e.cat
        cnt += 1
        upto = min(fy, 2021) if new else fy
        lw = last_work(e, upto)
        prev = f"{lw[0]}.{lw[1]:02d} {lw[2]}" if lw else ""
        an = W.asset_no(e, fy)
        nm = e.name if e.f == 0 else f"{e.f}号{e.name}"
        inst = e.install
        if an != e.no:
            inst = {"ボイラ給水ポンプA": "2017-10" if e.f == 1 else "2018-06"}.get(e.name, inst)
        vals = [an, nm, e.cat, W.CATS[e.cat], e.f, e.spec, e.maker, inst.replace("-", "/"), W.LEGAL_TAG.get(e.strat, ""), e.imp]
        if new:
            vals.append(e.policy)
        vals += [e.role, prev, note_for(e, fy) or None]
        fyv = None
        if s2:
            lw2 = last_work(e, fy)
            fyv = lw2[0] if lw2 else None
            vals.append(fyv)
        for i, val in enumerate(vals, 1):
            put(ws, r, i, val, font(9), CENTER if heads[i - 1] in ("設備区分CD", "系列", "設置年月", "法定区分", "重要度", "予備機", "最終整備(FY)") else LEFT, B_HAIR)
        if new and e.policy == "事後保全":
            for c in range(1, NC + 1):
                ws.cell(row=r, column=c).fill = gray
        flds = {h: v for h, v in zip(heads, vals) if v not in (None, "")}
        R.add("資産一覧", r, "item", flds)
        r += 1
    subtotal(r, cur, cnt)
    r += 2
    put(ws, r, 2, f"合計 {len(eqs)} 台（1号炉 72・2号炉 72・共通 118）", font(10, True))
    R.add("資産一覧", r, "total", {"台数": len(eqs), "1号炉": 72, "2号炉": 72, "共通": 118})
    if new:
        ws.auto_filter.ref = f"A{R0}:{L(NC)}{r - 2}"
    ws.freeze_panes = ws.cell(row=R0 + 1, column=3)
    page(ws, landscape=True, paper=9, fit_w=1, fit_h=0, area=f"A1:{L(NC)}{r}", titles=f"{R0}:{R0}", footer="&P / &N")
    if new:
        # 更新・付替えシート（2018 年度に s1 が追加。更新で資産番号が変わったものを旧新併記）
        w2 = wb.create_sheet("更新・付替え")
        put(w2, 1, 1, "【架空】ひばり野クリーンセンター　更新・付替え資産（旧番号と新番号）", font(12, True))
        put(w2, 2, 1, "SPC-AM は更新時に旧資産を除却し、末尾 R の新番号を起こす。保全課で旧新を対応付け（H30 追加）", font(8))
        hh = ["旧資産番号", "新資産番号", "資産名称", "更新年月", "内容", "記入"]
        for i, (h, w) in enumerate(zip(hh, [12, 12, 26, 10, 34, 6]), 1):
            put(w2, 4, i, h, font(9, True), CENTER, B_ALL, fill("FFF2CC"))
            w2.column_dimensions[L(i)].width = w
        rr = 5
        for (nm_, f_), (yy, suf) in W.RENUM.items():
            if yy > fy:
                continue
            e = W.eq(nm_, f_)
            row = [e.no, e.no + suf, e.full, {"ボイラ給水ポンプA": f"{yy}.{'10' if f_ == 1 else '06'}"}.get(nm_, f"{yy}.08"),
                   "同形式で更新" if "ポンプ" in nm_ else "ロードセル・指示計・データ処理装置 更新", "s1"]
            for i, val in enumerate(row, 1):
                put(w2, rr, i, val, font(9), LEFT, B_HAIR)
            R.add("更新・付替え", rr, "item", dict(zip(hh, row)))
            rr += 1
        page(w2, landscape=False, paper=9, fit_w=1, fit_h=1)
    wb.save(path(fn))
    return fn


def build_all():
    return [build(fy) for fy in range(2016, 2027)]
