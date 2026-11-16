# -*- coding: utf-8 -*-
"""工程表（定期整備工事）。2016〜2018：SPC s1 様式（予定を業者色で塗る）。2019〜：協力会社 w1 様式（予定／実績 2 段の線）。
2024〜は w2 が写して使う（空き 2 行・シート名の残り）。"""
import datetime as dt
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter as L
import world as W
import works as WK
import contracts as K
from xlhelp import *
from rec import Rec, path

COL18 = {"SPC運転": "BDD7EE", "鳶班": "D9D9D9", "電気班": "C9C9FF", "築炉班": "F4B084", "検査班": "FFD966", "タービン班": "FFC7CE"}
WHO18 = {"SPC運転": "SPC運転", "鳶班": "鳶", "電気班": "鷹羽電機", "築炉班": "白砂耐火", "検査班": "東栄検査", "タービン班": "鷹羽"}


def fname(c):
    sty = "SPC様式" if c["y"] <= 2018 else "協力会社様式"
    return f"【架空】ひばり野クリーンセンター_工程表_{c['y']}年度{c['f']}号炉定期整備工事_{sty}.xlsx"


def works_of(t, y):
    out = []
    for nm, f in t[7]:
        e = W.eq(nm, f)
        w = WK.wid(e, y)
        if w:
            out.append(w)
    return out


def iso(y, s):
    m, d = s.split("/")
    return f"{y}-{int(m):02d}-{int(d):02d}"


def build_s1(c):
    fn = fname(c)
    R = Rec(fn)
    y, f, tasks = c["y"], c["f"], c["tasks"]
    wb = Workbook()
    ws = wb.active
    ws.title = "工程表"
    mk = dt.datetime(y, c["stop"].month - 1, 20, 15, 0)
    props(wb, f"【架空】ひばり野クリーンセンター {c['title']} 工程表", "SPC保全課 s1", mk)
    alld = [x for t in tasks for x in K.dates(y, *t[5])]
    days = K.dates(y, "%d/%d" % ((min(alld) - dt.timedelta(days=1)).month, (min(alld) - dt.timedelta(days=1)).day),
                   "%d/%d" % (max(alld).month, max(alld).day))
    C0 = 5
    for col, w in {1: 4, 2: 16, 3: 40, 4: 11}.items():
        ws.column_dimensions[L(col)].width = w
    for i in range(len(days)):
        ws.column_dimensions[L(C0 + i)].width = 3.2
    put(ws, 1, 1, f"【架空】ひばり野クリーンセンター　{c['title']}　工程表", font(15, True))
    st = c["stop"]
    put(ws, 2, 1, f"契約工期：{c['period'][0]:%Y/%-m/%-d}〜{c['period'][1]:%Y/%-m/%-d}　　炉停止：{st:%Y/%-m/%-d}〜（予定）　　発注：ひばり野環境サービス㈱　受注：鷹羽エンジニアリング㈱", font(9))
    put(ws, 3, 1, f"{mk:%Y/%m/%d} 作成　保全課", font(9))
    for j, h in enumerate(("承認", "確認", "作成")):
        cc = C0 + len(days) - 9 + 3 * j
        ws.merge_cells(start_row=1, start_column=cc, end_row=1, end_column=cc + 2)
        put(ws, 1, cc, h, font(8), CENTER)
        ws.merge_cells(start_row=2, start_column=cc, end_row=3, end_column=cc + 2)
        put(ws, 2, cc, "早瀬" if h == "作成" else None, font(9), CENTER)
        box(ws, 1, cc, 3, cc + 2)
    Rr = 5
    for col, h in ((1, "No"), (2, "設備"), (3, "作業項目"), (4, "施工")):
        ws.merge_cells(start_row=Rr, start_column=col, end_row=Rr + 2, end_column=col)
        put(ws, Rr, col, h, font(9, True), CENTER, B_ALL, fill("D9E1F2"))
    m0 = None
    for i, x in enumerate(days):
        col = C0 + i
        if x.month != m0:
            m0 = x.month
            put(ws, Rr, col, f"{x.month}月", font(9, True), Alignment(horizontal="left", vertical="center", wrap_text=False), None, fill("D9E1F2"))
        put(ws, Rr + 1, col, x.day, font(8), CENTER, B_ALL, fill("D9E1F2"))
        wd = K.WD[x.weekday()]
        put(ws, Rr + 2, col, wd, font(8, color="C00000" if wd == "日" else None), CENTER, B_ALL,
            fill("F8CBAD") if wd == "日" else fill("D9E1F2"))
    for col in range(C0, C0 + len(days)):
        ws.cell(row=Rr, column=col).fill = fill("D9E1F2")
    r = Rr + 3
    prev = None
    for n, t in enumerate(tasks, 1):
        area = t[1]
        who = WHO18.get(t[3].replace("（応援）", ""), "野分工業")
        col_ = COL18.get(t[3].replace("（応援）", ""), "A9D08E")
        put(ws, r, 1, n, font(9), CENTER, B_ALL)
        av = area if area != prev else "〃"
        put(ws, r, 2, av, font(9), CENTER if area == prev else LEFT, B_ALL)
        prev = area
        put(ws, r, 3, t[2], font(9), LEFT, B_ALL)
        put(ws, r, 4, who, font(8), CENTER, B_ALL)
        ds = set(K.dates(y, *t[5]))
        for i, x in enumerate(days):
            cell = ws.cell(row=r, column=C0 + i)
            cell.border = B_HAIR
            if x.weekday() == 6:
                cell.fill = fill("FCE4D6")
            if x in ds:
                cell.fill = fill(col_)
                if ("検査" in t[2] or "水圧" in t[2]) and x == max(ds):
                    cell.value = "★"
                    cell.font = font(9, True)
                    cell.alignment = CENTER
        ws.row_dimensions[r].height = 16
        R.add("工程表", r, "item", {"No": n, "設備": av, "作業項目": t[2], "施工": who,
                                 "予定開始": iso(y, t[5][0]), "予定終了": iso(y, t[5][1])}, works_of(t, y))
        r += 1
    r += 1
    put(ws, r, 2, "色分け：", font(8))
    cc = C0
    for k_, col_ in list(COL18.items()) + [("野分工業", "A9D08E")]:
        ws.merge_cells(start_row=r, start_column=cc, end_row=r, end_column=cc + 2)
        put(ws, r, cc, WHO18.get(k_, k_), font(7), CENTER, B_ALL, fill(col_))
        cc += 3
    put(ws, r + 1, 3, "★：立会検査（SPC・検査員）　網掛け：日曜（休工）", font(8))
    page(ws, landscape=True, paper=8, fit_w=1, fit_h=1, area=f"A1:{L(C0 + len(days) - 1)}{r + 1}")
    wb.save(path(fn))
    return fn


def build_w(c):
    fn = fname(c)
    R = Rec(fn)
    y, f, tasks, cid = c["y"], c["f"], c["tasks"], c["cid"]
    con = K.CONTRACTS[cid]
    w2 = y >= 2024
    name = "森野" if w2 else "梶本"
    author = "野分工業 職長 w2" if w2 else "野分工業 職長 w1"
    st = c["stop"]
    made = dt.datetime(y, st.month - 1, 22, 17, 40)
    sheetname = "2023_2号炉" if y == 2024 else (f"工程表({y})" if w2 else "工程表")
    wb = Workbook()
    ws = wb.active
    ws.title = sheetname
    props(wb, f"【架空】ひばり野クリーンセンター {con['title']} 工程表", author, made)
    alld = [x for t in tasks for x in K.dates(y, *t[5])] + [x for t in tasks for x in K.task_days(t, y)]
    d0 = min(alld) - dt.timedelta(days=1)
    d1 = max(alld) + dt.timedelta(days=1)
    days = K.dates(y, f"{d0.month}/{d0.day}", f"{d1.month}/{d1.day}")
    C0 = 7
    for col, w in {1: 4, 2: 13, 3: 34, 4: 11, 5: 4, 6: 5}.items():
        ws.column_dimensions[L(col)].width = w
    for i in range(len(days)):
        ws.column_dimensions[L(C0 + i)].width = 3.0
    CN = C0 + len(days)
    ws.column_dimensions[L(CN)].width = 22
    put(ws, 1, 1, f"【架空】{con['title']}　工　程　表", font(15, True))
    put(ws, 2, 1, f"施設：【架空】ひばり野クリーンセンター　　発注者：{con['orderer']}　元請：{con['contractor']}　施工：{con['sub']}", font(9))
    stops = [x for t in tasks if t[2].startswith("炉立下げ") for x in K.dates(y, *t[5])]
    ends = [x for t in tasks if t[2].startswith("炉立上げ") for x in K.dates(y, *t[5])]
    note = f"炉停止：{min(stops):%Y/%-m/%-d}〜{max(ends):%-m/%-d}（予定）"
    if (y, f) == (2025, 2):
        note += "　1号炉は運転継続（タービン開放中は復水器へバイパス）"
    put(ws, 3, 1, f"工期：{con['period'][0]:%Y/%m/%d}〜{con['period'][1]:%Y/%m/%d}　　{note}", font(9))
    put(ws, 3, max(CN - 12, 8), f"野分工業㈱ 職長 {name}　作成 {made:%Y/%m/%d}　実績記入 {max(alld) + dt.timedelta(days=1):%Y/%m/%d}", font(9))
    Rr = 5
    for col, h in ((1, "No"), (2, "設備"), (3, "作業項目"), (4, "班"), (5, "予\n実"), (6, "人員")):
        ws.merge_cells(start_row=Rr, start_column=col, end_row=Rr + 1, end_column=col)
        put(ws, Rr, col, h, font(9, True), CENTER, B_ALL, fill("E2EFDA"))
    for i, x in enumerate(days):
        col = C0 + i
        put(ws, Rr, col, f"{x.month}/{x.day}" if (x.day == 1 or i == 0 or x.weekday() == 0) else x.day, font(6), CENTER, B_ALL, fill("E2EFDA"))
        wd = K.WD[x.weekday()]
        put(ws, Rr + 1, col, wd, font(7, color="C00000" if wd == "日" else None), CENTER, B_ALL, fill("E2EFDA"))
    ws.merge_cells(start_row=Rr, start_column=CN, end_row=Rr + 1, end_column=CN)
    put(ws, Rr, CN, "備考", font(9, True), CENTER, B_ALL, fill("E2EFDA"))
    blue = Side(style="thick", color="2F5597")
    redl = Side(style="thick", color="C00000")
    r = Rr + 2
    groups = []
    for t in tasks:
        if t[1] not in groups:
            groups.append(t[1])
    head = [t for t in tasks if t[1] == "共通" and t[0] in ("T01", "T02", "T03", "T04")]
    seq = list(head)
    for g in [g for g in groups if g != "共通"]:
        seq += [t for t in tasks if t[1] == g]
    seq += [t for t in tasks if t[1] == "共通" and t not in seq]
    n = 0
    prev = None
    blank_done = not w2
    for t in seq:
        if t[1] == "共通" and t[0] not in ("T01", "T02", "T03", "T04") and not blank_done:
            for _ in range(2):
                box(ws, r, 1, r + 1, CN, B_HAIR)
                r += 2
            blank_done = True
        n += 1
        tid, area, item, cw, men, (ps, pe), act = t[:7]
        ws.merge_cells(start_row=r, start_column=1, end_row=r + 1, end_column=1)
        put(ws, r, 1, n, font(9), CENTER)
        ws.merge_cells(start_row=r, start_column=2, end_row=r + 1, end_column=2)
        av = area if area != prev else None
        put(ws, r, 2, av, font(8), LEFT)
        prev = area
        ws.merge_cells(start_row=r, start_column=3, end_row=r + 1, end_column=3)
        put(ws, r, 3, item, font(9), LEFT)
        ws.merge_cells(start_row=r, start_column=4, end_row=r + 1, end_column=4)
        put(ws, r, 4, cw.replace("SPC運転", "SPC"), font(8), CENTER)
        put(ws, r, 5, "予", font(7), CENTER)
        put(ws, r + 1, 5, "実", font(7, color="C00000"), CENTER)
        ws.merge_cells(start_row=r, start_column=6, end_row=r + 1, end_column=6)
        put(ws, r, 6, men or "－", font(9), CENTER)
        pd = set(K.dates(y, ps, pe))
        ad = set(K.task_days(t, y))
        for i, x in enumerate(days):
            col = C0 + i
            for rr, ds, side in ((r, pd, blue), (r + 1, ad, redl)):
                cell = ws.cell(row=rr, column=col)
                cell.border = Border(left=hair, right=hair, top=hair, bottom=side if x in ds else hair)
                if x.weekday() == 6:
                    cell.fill = fill("FCE4D6")
            if x in pd and ("検査" in item or "水圧" in item) and x == max(pd):
                put(ws, r, col, "◇", font(8, True, color="2F5597"), CENTER)
        note = ""
        if act and act != (ps, pe):
            note = f"実績 {act[0]}〜{act[1]}"
            reason = t[10] if len(t) > 10 and t[10] else None
            special = {("T07", 2024): "（支給火格子 2回目 2日遅れ）", ("T20", 2024): "（ランナ肉盛 追加）",
                       ("T21", 2024): "（IDF 追加作業のため 1日後ろへ）", ("T24", 2024): "（側板 当て板 追加）",
                       ("T10", 2025): "（付着灰 多く 1日延長）", ("T19", 2025): "（潤滑油フラッシング 延長）"}
            if (y, f) in ((2024, 1), (2025, 2)):
                note += special.get((tid, y), "")
            elif reason:
                note += f"（{reason}）"
            else:
                note += "（前工程の遅れ）"
        ws.merge_cells(start_row=r, start_column=CN, end_row=r + 1, end_column=CN)
        put(ws, r, CN, note or None, font(7), LEFT)
        for col in list(range(1, 7)) + [CN]:
            ws.cell(row=r, column=col).border = B_ALL
            ws.cell(row=r + 1, column=col).border = B_ALL
        ws.row_dimensions[r].height = 12
        ws.row_dimensions[r + 1].height = 12
        a = act or (ps, pe)
        flds = {"No": n, "設備": av, "作業項目": item, "班": cw.replace("SPC運転", "SPC"), "人員": men or None,
                "予定開始": iso(y, ps), "予定終了": iso(y, pe), "実績開始": iso(y, a[0]), "実績終了": iso(y, a[1])}
        if note:
            flds["備考"] = note
        R.add(sheetname, r, "item", flds, works_of(t, y))
        r += 2
    r += 1
    put(ws, r, 2, "凡例", font(8, True))
    ws.cell(row=r, column=3).border = Border(bottom=blue)
    put(ws, r, 3, "予定（青線）", font(8))
    ws.cell(row=r + 1, column=3).border = Border(bottom=redl)
    put(ws, r + 1, 3, "実績（赤線）　◇ 立会検査", font(8))
    if y == 2024 and f == 2:
        put(ws, r, C0 + 2, "※ 2023年度 2号炉の表をコピーして作成", font(8, color="808080"))
    if y == 2024 and f == 1:
        put(ws, r, C0 + 2, "※ 2023年度 2号炉の表をコピーして作成", font(8, color="808080"))
    page(ws, landscape=True, paper=8, fit_w=1, fit_h=1, area=f"A1:{L(CN)}{r + 1}")
    wb.save(path(fn))
    return fn


def build(c):
    return build_s1(c) if c["y"] <= 2018 else build_w(c)
