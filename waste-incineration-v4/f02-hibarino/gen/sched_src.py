# -*- coding: utf-8 -*-
"""工程表。
旧 = SPC s1 様式（〜2018）：1工事1枚、日単位、業者ごとの色で塗る。
新 = 協力会社 w1 様式（2019〜）：日単位、予定／実績の 2 段を線（下罫線）で引く。2024 から w2 がコピーして使用。
市追加工事 = メーカー（鷹羽）様式：半日単位、和暦。"""
import datetime as dt
import os
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter as L
import world as W
import contracts as K
from xlhelp import *

HOL = {dt.date(2018, 10, 8), dt.date(2018, 11, 3), dt.date(2024, 10, 14), dt.date(2024, 11, 3), dt.date(2024, 11, 4),
       dt.date(2025, 10, 13), dt.date(2025, 11, 3), dt.date(2023, 10, 9), dt.date(2023, 11, 3)}

# ---------------------------------------------------------------- 2018 1号炉（s1 様式の見本）
T2018 = [
    ("共通", "炉立下げ・冷却", "SPC運転", "10/9", "10/12"),
    ("共通", "縁切り・足場組立", "鳶・電気", "10/11", "10/13"),
    ("燃焼設備", "炉内清掃", "野分工業", "10/13", "10/15"),
    ("燃焼設備", "乾燥・燃焼・後燃焼ストーカ 火格子一部交換", "野分工業", "10/16", "10/25"),
    ("燃焼設備", "炉内耐火物補修", "白砂耐火", "10/18", "10/27"),
    ("燃焼ガス冷却設備", "ボイラ清掃・肉厚測定", "野分工業", "10/15", "10/19"),
    ("燃焼ガス冷却設備", "ボイラ定期事業者検査", "SPC・鷹羽", "10/23", "10/23"),
    ("燃焼ガス冷却設備", "スートブロワ分解整備", "野分工業", "10/22", "10/26"),
    ("排ガス処理設備", "減温塔・バグフィルタ点検", "野分工業", "10/15", "10/18"),
    ("排ガス処理設備", "消石灰・活性炭 定量供給機 分解整備", "野分工業", "10/19", "10/23"),
    ("排ガス処理設備", "触媒反応塔 点検", "野分工業", "10/24", "10/24"),
    ("通風設備", "二次送風機 分解整備", "野分工業", "10/25", "10/29"),
    ("通風設備", "誘引通風機・押込送風機 点検", "野分工業", "10/30", "10/31"),
    ("灰出し設備", "灰押出装置 補修", "野分工業", "10/16", "10/20"),
    ("電気設備", "インバータ盤（IDF・FDF）整備", "鷹羽電機", "10/29", "10/31"),
    ("共通", "足場解体・復旧", "鳶", "11/1", "11/2"),
    ("共通", "ボイラ水圧試験", "SPC・鷹羽", "11/2", "11/2"),
    ("共通", "試運転・立上げ", "SPC運転", "11/5", "11/8"),
]
COL18 = {"SPC運転": "BDD7EE", "鳶・電気": "D9D9D9", "野分工業": "A9D08E", "白砂耐火": "F4B084", "SPC・鷹羽": "FFD966",
         "鷹羽電機": "C9C9FF", "鳶": "D9D9D9"}


def build_2018(path):
    wb = Workbook()
    ws = wb.active
    ws.title = "工程表"
    props(wb, "【架空】ひばり野クリーンセンター 2018年度 1号炉定期整備工事 工程表", "SPC保全課 s1", dt.datetime(2018, 8, 20, 15, 0))
    days = K.dates(2018, "10/8", "11/9")
    C0 = 5
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 34
    ws.column_dimensions["D"].width = 11
    for i in range(len(days)):
        ws.column_dimensions[L(C0 + i)].width = 3.2
    put(ws, 1, 1, "【架空】ひばり野クリーンセンター　2018年度 1号炉定期整備工事　工程表", font(15, True))
    put(ws, 2, 1, "契約工期：2018/9/3〜2018/11/30　　炉停止期間：2018/10/9〜2018/11/5（予定）　　発注：ひばり野環境サービス㈱　受注：鷹羽エンジニアリング㈱", font(9))
    put(ws, 3, 1, "2018/08/20 作成　保全課", font(9))
    # 押印欄
    for j, h in enumerate(("承認", "確認", "作成")):
        c = C0 + len(days) - 9 + 3 * j
        ws.merge_cells(start_row=1, start_column=c, end_row=1, end_column=c + 2)
        put(ws, 1, c, h, font(8), CENTER)
        ws.merge_cells(start_row=2, start_column=c, end_row=3, end_column=c + 2)
        put(ws, 2, c, "早瀬" if h == "作成" else None, font(9), CENTER)
        box(ws, 1, c, 3, c + 2)
    R = 5
    for c, h in ((1, "No"), (2, "設備"), (3, "作業項目"), (4, "施工")):
        ws.merge_cells(start_row=R, start_column=c, end_row=R + 2, end_column=c)
        put(ws, R, c, h, font(9, True), CENTER, B_ALL, fill("D9E1F2"))
    # 月の見出し
    m0 = None
    for i, x in enumerate(days):
        c = C0 + i
        if x.month != m0:
            start = c
            m0 = x.month
            put(ws, R, c, f"{x.month}月", font(9, True), Alignment(horizontal="left", vertical="center", wrap_text=False), None, fill("D9E1F2"))
        put(ws, R + 1, c, x.day, font(8), CENTER, B_ALL, fill("D9E1F2"))
        wd = K.WD[x.weekday()]
        put(ws, R + 2, c, wd, font(8, color="C00000" if wd == "日" or x in HOL else None), CENTER, B_ALL,
            fill("F8CBAD") if wd == "日" or x in HOL else fill("D9E1F2"))
    for c in range(C0, C0 + len(days)):
        ws.cell(row=R, column=c).fill = fill("D9E1F2")
        ws.cell(row=R, column=c).border = Border(top=thin, bottom=thin)
    r = R + 3
    prev = None
    for n, (area, item, who, s, e) in enumerate(T2018, 1):
        put(ws, r, 1, n, font(9), CENTER, B_ALL)
        put(ws, r, 2, area if area != prev else "〃", font(9), CENTER if area == prev else LEFT, B_ALL)
        prev = area
        put(ws, r, 3, item, font(9), LEFT, B_ALL)
        put(ws, r, 4, who, font(8), CENTER, B_ALL)
        ds = set(K.dates(2018, s, e))
        for i, x in enumerate(days):
            c = C0 + i
            cell = ws.cell(row=r, column=c)
            cell.border = B_HAIR
            if x.weekday() == 6 or x in HOL:
                cell.fill = fill("FCE4D6")
            if x in ds:
                cell.fill = fill(COL18[who])
                if "検査" in item or "水圧" in item:
                    cell.value = "★"
                    cell.font = font(9, True)
                    cell.alignment = CENTER
        ws.row_dimensions[r].height = 18
        r += 1
    r += 1
    put(ws, r, 2, "色分け：", font(8))
    c = C0
    for who, col in COL18.items():
        if who == "鳶":
            continue
        ws.merge_cells(start_row=r, start_column=c, end_row=r, end_column=c + 2)
        put(ws, r, c, who, font(7), CENTER, B_ALL, fill(col))
        c += 3
    put(ws, r + 1, 3, "★：立会検査（SPC・検査員）　網掛け：日曜・祝日（休工）", font(8))
    ws.cell(row=R + 3, column=C0).comment = Comment("炉停止は10/9 6:00 の予定。運転課と調整済み", "s1")
    page(ws, landscape=True, paper=8, fit_w=1, fit_h=1, area=f"A1:{L(C0 + len(days) - 1)}{r + 1}")
    wb.save(path)


# ---------------------------------------------------------------- 協力会社 w1 様式（2024, 2025 は w2）
def build_w(path, tasks, year, cid, crew, sheetname, days_rng, made, author, title_note=None, leftover_note=None):
    con = K.CONTRACTS[cid]
    wb = Workbook()
    ws = wb.active
    ws.title = sheetname
    props(wb, f"【架空】ひばり野クリーンセンター {con['title']} 工程表", author, made)
    days = K.dates(year, *days_rng)
    C0 = 7
    for c, w in {1: 4, 2: 13, 3: 34, 4: 11, 5: 4, 6: 5}.items():
        ws.column_dimensions[L(c)].width = w
    for i in range(len(days)):
        ws.column_dimensions[L(C0 + i)].width = 3.0
    CN = C0 + len(days)
    ws.column_dimensions[L(CN)].width = 22
    put(ws, 1, 1, f"【架空】{con['title']}　工　程　表", font(15, True))
    put(ws, 2, 1, f"施設：【架空】ひばり野クリーンセンター　　発注者：{con['orderer']}　元請：{con['contractor']}　施工：{con['sub']}", font(9))
    put(ws, 3, 1, f"工期：{con['period'][0]:%Y/%m/%d}〜{con['period'][1]:%Y/%m/%d}　　{title_note or ''}", font(9))
    put(ws, 3, CN - 12, f"野分工業㈱ 職長 森野　作成 {made:%Y/%m/%d}　実績記入 {con['period'][1] - dt.timedelta(days=23):%Y/%m/%d}", font(9))
    R = 5
    for c, h in ((1, "No"), (2, "設備"), (3, "作業項目"), (4, "班"), (5, "予\n実"), (6, "人員")):
        ws.merge_cells(start_row=R, start_column=c, end_row=R + 1, end_column=c)
        put(ws, R, c, h, font(9, True), CENTER, B_ALL, fill("E2EFDA"))
    for i, x in enumerate(days):
        c = C0 + i
        put(ws, R, c, f"{x.month}/{x.day}" if (x.day == 1 or i == 0 or x.weekday() == 0) else x.day, font(6), CENTER, B_ALL,
            fill("E2EFDA"))
        wd = K.WD[x.weekday()]
        put(ws, R + 1, c, wd, font(7, color="C00000" if wd == "日" else None), CENTER, B_ALL, fill("E2EFDA"))
    ws.merge_cells(start_row=R, start_column=CN, end_row=R + 1, end_column=CN)
    put(ws, R, CN, "備考", font(9, True), CENTER, B_ALL, fill("E2EFDA"))
    blue = Side(style="thick", color="2F5597")
    redl = Side(style="thick", color="C00000")
    r = R + 2
    groups = []
    for t in tasks:
        if t[1] not in groups:
            groups.append(t[1])
    order = [g for g in groups if g != "共通"]
    seq = [t for t in tasks if t[1] == "共通" and t[0] in ("T01", "T02", "T03", "T04")]
    for g in order:
        seq += [t for t in tasks if t[1] == g]
    seq += [t for t in tasks if t[1] == "共通" and t not in seq]
    n = 0
    prev = None
    blank_done = False
    for t in seq:
        if t[1] == "共通" and t[0] not in ("T01", "T02", "T03", "T04") and not blank_done:
            # 2024 年に w2 が行を足して消し忘れた 2 行（以後のコピーにも残る）
            for _ in range(2):
                box(ws, r, 1, r + 1, CN, B_HAIR)
                r += 2
            blank_done = True
        n += 1
        tid, area, item, cw, men, (ps, pe), act = t[:7]
        ws.merge_cells(start_row=r, start_column=1, end_row=r + 1, end_column=1)
        put(ws, r, 1, n, font(9), CENTER)
        ws.merge_cells(start_row=r, start_column=2, end_row=r + 1, end_column=2)
        put(ws, r, 2, area if area != prev else "", font(8), LEFT)
        prev = area
        ws.merge_cells(start_row=r, start_column=3, end_row=r + 1, end_column=3)
        put(ws, r, 3, item, font(9), LEFT)
        ws.merge_cells(start_row=r, start_column=4, end_row=r + 1, end_column=4)
        put(ws, r, 4, cw.replace("SPC運転", "SPC"), font(8), CENTER)
        put(ws, r, 5, "予", font(7), CENTER)
        put(ws, r + 1, 5, "実", font(7, color="C00000"), CENTER)
        ws.merge_cells(start_row=r, start_column=6, end_row=r + 1, end_column=6)
        put(ws, r, 6, men or "－", font(9), CENTER)
        pd = set(K.dates(year, ps, pe))
        ad = set(K.task_days(t, year)) if year < 2026 else set()
        for i, x in enumerate(days):
            c = C0 + i
            for rr, ds, side in ((r, pd, blue), (r + 1, ad, redl)):
                cell = ws.cell(row=rr, column=c)
                bl = cell.border
                if x in ds:
                    cell.border = Border(left=hair, right=hair, top=hair, bottom=side)
                else:
                    cell.border = Border(left=hair, right=hair, top=hair, bottom=hair)
                if x.weekday() == 6:
                    cell.fill = fill("FCE4D6")
            if x in pd and ("検査" in item or "水圧" in item) and x == max(pd):
                put(ws, r, c, "◇", font(8, True, color="2F5597"), CENTER)
        note = ""
        if act and act != (ps, pe):
            note = f"実績 {act[0]}〜{act[1]}"
            if tid == "T07" and year == 2024:
                note += "（支給火格子 2回目 2日遅れ）"
            if tid == "T20" and year == 2024:
                note += "（ランナ肉盛 追加）"
            if tid == "T21" and year == 2024:
                note += "（IDF 追加作業のため 1日後ろへ）"
            if tid == "T24" and year == 2024:
                note += "（側板 当て板 追加）"
            if tid == "T10" and year == 2025:
                note += "（付着灰 多く 1日延長）"
            if tid == "T19" and year == 2025:
                note += "（潤滑油フラッシング 延長）"
        ws.merge_cells(start_row=r, start_column=CN, end_row=r + 1, end_column=CN)
        put(ws, r, CN, note or None, font(7), LEFT)
        for c in range(1, 7):
            ws.cell(row=r, column=c).border = B_ALL
            ws.cell(row=r + 1, column=c).border = B_ALL
        ws.cell(row=r, column=CN).border = B_ALL
        ws.cell(row=r + 1, column=CN).border = B_ALL
        ws.row_dimensions[r].height = 12
        ws.row_dimensions[r + 1].height = 12
        r += 2
    r += 1
    put(ws, r, 2, "凡例", font(8, True))
    ws.cell(row=r, column=3).border = Border(bottom=blue)
    put(ws, r, 3, "予定（青線）", font(8))
    ws.cell(row=r + 1, column=3).border = Border(bottom=redl)
    put(ws, r + 1, 3, "実績（赤線）　◇ 立会検査", font(8))
    if leftover_note:
        put(ws, r, C0 + 2, leftover_note, font(8, color="808080"))
    page(ws, landscape=True, paper=8, fit_w=1, fit_h=1, area=f"A1:{L(CN)}{r + 1}")
    wb.save(path)


# ---------------------------------------------------------------- メーカー（鷹羽）様式 市追加工事 R5
def build_city(path):
    con = K.CONTRACTS["CITY-R5"]
    wb = Workbook()
    ws = wb.active
    ws.title = "実施工程表"
    props(wb, "【架空】令和5年度 ひばり野クリーンセンター 2号ボイラ水管補修工事 実施工程表", "鷹羽エンジニアリング 現場代理人 m1",
          dt.datetime(2023, 9, 28, 11, 0))
    days = K.dates(2023, "10/2", "11/2")
    work = K.dates(2023, "10/9", "10/28")
    C0 = 5
    for c, w in {1: 5, 2: 8, 3: 36, 4: 6}.items():
        ws.column_dimensions[L(c)].width = w
    for i in range(2 * len(work)):
        ws.column_dimensions[L(C0 + i)].width = 2.4
    put(ws, 1, 1, "【架空】令和5年度 ひばり野クリーンセンター 2号ボイラ水管補修工事　実施工程表", font(14, True))
    put(ws, 2, 1, "【架空】施設：ひばり野クリーンセンター（ひばり野市）　発注者：ひばり野市　受注者：鷹羽エンジニアリング株式会社", font(9))
    put(ws, 3, 1, "工期：令和5年9月25日〜令和5年12月15日　　現地施工：令和5年10月10日〜10月27日（2号炉定期整備（SPC）の炉停止期間内）", font(9))
    # 承認欄
    cE = C0 + 2 * len(work) - 12
    for j, h in enumerate(("監督員", "現場代理人")):
        c = cE + 6 * j
        ws.merge_cells(start_row=1, start_column=c, end_row=1, end_column=c + 5)
        put(ws, 1, c, h, font(8), CENTER)
        ws.merge_cells(start_row=2, start_column=c, end_row=3, end_column=c + 5)
        put(ws, 2, c, "宇佐見" if h == "現場代理人" else None, font(9), CENTER)
        box(ws, 1, c, 3, c + 5)
    R = 5
    for c, h in ((1, "番号"), (2, "工種"), (3, "細　別"), (4, "人員")):
        ws.merge_cells(start_row=R, start_column=c, end_row=R + 2, end_column=c)
        put(ws, R, c, h, font(9, True), CENTER, B_ALL, fill("EDEDED"))
    for i, x in enumerate(work):
        c = C0 + 2 * i
        ws.merge_cells(start_row=R, start_column=c, end_row=R, end_column=c + 1)
        put(ws, R, c, f"{x.month}/{x.day}", font(7), CENTER, B_ALL, fill("EDEDED"))
        ws.merge_cells(start_row=R + 1, start_column=c, end_row=R + 1, end_column=c + 1)
        wd = K.WD[x.weekday()]
        put(ws, R + 1, c, wd, font(7, color="C00000" if wd == "日" or x in HOL else None), CENTER, B_ALL, fill("EDEDED"))
        put(ws, R + 2, c, "前", font(6), CENTER, B_ALL, fill("EDEDED"))
        put(ws, R + 2, c + 1, "後", font(6), CENTER, B_ALL, fill("EDEDED"))
    r = R + 3
    # 参考：SPC の定修（別契約）
    put(ws, r, 1, "参考", font(8), CENTER, B_ALL)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
    put(ws, r, 2, "2号炉定期整備工事（SPC 発注・別契約）炉停止 10/2〜11/2", font(8, color="595959"), LEFT, B_ALL)
    for i in range(2 * len(work)):
        ws.cell(row=r, column=C0 + i).border = Border(bottom=Side(style="dashed", color="7F7F7F"), top=hair)
    r += 1
    slot = {}
    for i, x in enumerate(work):
        slot[(x, "AM")] = C0 + 2 * i
        slot[(x, "PM")] = C0 + 2 * i + 1
    prevk = None
    for n, (cid, kind, item, crew_, men, (s, e), flag) in enumerate(K.TCITY, 1):
        put(ws, r, 1, n, font(9), CENTER, B_ALL)
        put(ws, r, 2, kind if kind != prevk else "", font(9), CENTER, B_ALL)
        prevk = kind
        put(ws, r, 3, item + ("　※追加" if flag else ""), font(9, color="C00000" if flag else None), LEFT, B_ALL)
        put(ws, r, 4, men, font(9), CENTER, B_ALL)
        sx = dt.date(2023, *K.d(s[0]))
        ex = dt.date(2023, *K.d(e[0]))
        on = False
        for i, x in enumerate(work):
            for h in ("AM", "PM"):
                c = slot[(x, h)]
                cell = ws.cell(row=r, column=c)
                cell.border = Border(left=hair if h == "AM" else None, right=hair if h == "PM" else None, top=hair, bottom=hair)
                if (x, h) == (sx, s[1]):
                    on = True
                if on and x.weekday() != 6:
                    cell.fill = fill("7F7F7F" if not flag else "C00000")
                if (x, h) == (ex, e[1]):
                    on = False
                if x.weekday() == 6 and not cell.fill.fgColor.rgb.endswith("7F7F7F"):
                    cell.fill = fill("F2F2F2")
        ws.row_dimensions[r].height = 17
        r += 1
    r += 1
    put(ws, r, 3, "■ 施工（灰）　■ 追加（赤：RT 不合格継手の補修・再撮影）　日曜は休工。RT は 10/20 17:00 以降 立入制限", font(8))
    put(ws, r + 1, 3, "※ 水圧試験は SPC 2号炉定修のボイラ水圧試験と兼ねず単独で実施（市・SPC 立会）", font(8))
    page(ws, landscape=True, paper=9, fit_w=1, fit_h=1, area=f"A1:{L(C0 + 2 * len(work) - 1)}{r + 1}")
    wb.save(path)


FN = {
    "2018": "【架空】ひばり野クリーンセンター_工程表_2018年度1号炉定期整備工事_SPC様式.xlsx",
    "2024": "【架空】ひばり野クリーンセンター_工程表_2024年度1号炉定期整備工事_協力会社様式.xlsx",
    "2025": "【架空】ひばり野クリーンセンター_工程表_2025年度2号炉定期整備工事_協力会社様式.xlsx",
    "R5": "【架空】ひばり野クリーンセンター_工程表_令和5年度2号ボイラ水管補修工事_メーカー様式.xlsx",
}

if __name__ == "__main__":
    p = {k: os.path.join(OUT, v) for k, v in FN.items()}
    build_2018(p["2018"])
    build_w(p["2024"], K.T2024, 2024, "OH2024-1", K.CREW_2024, "2023_2号炉", ("9/30", "11/6"), dt.datetime(2024, 8, 26, 18, 10),
            "野分工業 職長 w2", "炉停止：2024/10/1〜10/31（予定）", "※ 2023年度 2号炉の表をコピーして作成")
    build_w(p["2025"], K.T2025, 2025, "OH2025-2", K.CREW_2025, "工程表(2025)", ("9/28", "10/30"), dt.datetime(2025, 8, 22, 17, 40),
            "野分工業 職長 w2", "炉停止：2025/9/29〜10/25（予定）　1号炉は運転継続（タービン開放中は復水器へバイパス）")
    build_city(p["R5"])
    for k in p:
        print(render(p[k], [1]))
