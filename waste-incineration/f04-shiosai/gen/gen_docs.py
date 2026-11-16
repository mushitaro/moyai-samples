# -*- coding: utf-8 -*-
"""工程表と作業日報（全年度）。様式は見本で承認されたもの：
 工程表  メーカー v1：H24〜H26 旧様式（粗い帯）→ H27〜 実施工程表（班の色）／浜風機工：g1 日単位（〜H30）→ g2 時間単位（R1〜）
 日報    メーカー v1：1 日 1 枚（全年度）／浜風機工：g1 1 日 1 枚（〜R1、R1 は g2 が使用）→ g2 工事別 1 枚（R2〜）"""
import datetime as dt
import world as W, cw, sg
from xl import *
from cw import D

YOUBI = "月火水木金土日"
def jd(d): return d.strftime("%Y-%m-%d")
def span_txt(spans): return ["%s/%s" % (jd(a), jd(b)) for (a, b) in spans]

def vname(c): return c.vendor.replace("【架空】", "").replace("有限会社 ", "").replace("株式会社", "").strip()

def label(c, fy=None):
    fy = fy or c.fy
    if c.kind == "MK": return "%s定期点検整備業務委託" % W.wareki(fy)
    if c.kind == "EN": return "%s長寿命化改良工事%d号炉" % (W.wareki(fy), 1 if fy == 2015 else 2)
    if c.kind == "CRUP": return "R6ごみクレーン制御装置更新工事"
    t = c.short.replace("（", "").replace("）", "").replace(" ", "")
    return W.wareki(fy) + ("年度" if t[:1].isdigit() else "") + t

def lname(c):
    """地元業者の担当（g1 は H24〜H30、g2 は R1〜）"""
    return "g1" if c.fy <= 2018 else "g2"

# ===================== 工程表 =====================
def koutei_v1_old(c, s):
    fy = c.fy
    fname = "【架空】しおさい町ごみ処理場_工程表_%s_水樹プラント旧様式.xlsx" % label(c)
    rec_open(fname)
    wb = new_wb("【架空】" + c.name + " 工程表", "水樹プラント v1", dt.datetime(fy, 9, 25, 15, 0))
    ws = wb.active; ws.title = "工程表"
    d0 = s.site[0] - dt.timedelta(days=1)
    days = [d0 + dt.timedelta(days=i) for i in range((s.site[1] - d0).days + 1)]
    widths(ws, [10, 30] + [3.6] * len(days) + [16])
    merge(ws, 1, 1, 1, 3 + len(days), "【架空】定期点検整備　工程表", F(14, True), AC, None)
    put(ws, 2, 1, "件名　【架空】" + c.name, F(9))
    put(ws, 3, 1, "%s　現場責任者 v1" % W.MAKER, F(9))
    r = 5
    put(ws, r, 1, "号炉", F(9, True), AC, B_ALL, fill("EDEDED")); put(ws, r, 2, "項目", F(9, True), AC, B_ALL, fill("EDEDED"))
    for j, d in enumerate(days):
        put(ws, r, 3 + j, "%d/%d" % (d.month, d.day) if j == 0 or d.day == 1 else d.day, F(7), AC, B_ALL, fill("EDEDED"))
        put(ws, r + 1, 3 + j, YOUBI[d.weekday()], F(7, color="C00000" if d.weekday() == 6 else None), AC, B_ALL)
    put(ws, r + 1, 1, None, None, AC, B_ALL); put(ws, r + 1, 2, None, None, AC, B_ALL)
    put(ws, r, 3 + len(days), "備考", F(9, True), AC, B_ALL, fill("EDEDED"))
    r += 2
    rows = []
    for g in ("1号炉", "共通設備", "2号炉"):
        ts = [t for t in s.tasks if t.group == g]
        if not ts: continue
        std = [t for t in ts if not t.works or t.name.startswith(("火格子", "送風機", "ガス冷却室", "パルス", "燃焼制御盤", "ごみクレーン", "空気槽", "中央監視"))]
        spec = [t for t in ts if t not in std and not t.ref]
        allw = [w for t in std for w in t.works]
        a = min(t.plan[0][0] for t in std); b = max(t.plan[0][1] for t in std)
        rows.append((g, "点検整備" if g != "共通設備" else "共通設備 点検（全炉停止）", [(a, b)], allw, ""))
        for t in spec:
            rows.append((g, "　" + t.name, t.plan, t.works, "特記"))
        for t in ts:
            if t.ref: rows.append((g, t.name, t.plan, [], "町発注"))
    for (g, nm, spans, wk, nt) in rows:
        put(ws, r, 1, g, F(8), AC, B_ALL)
        put(ws, r, 2, nm, F(8), AL, B_ALL)
        for j, d in enumerate(days):
            cell = put(ws, r, 3 + j, None, F(8), AC, B_ALL)
            for (a, b) in spans:
                if a <= d <= b: cell.fill = fill("A6A6A6") if nt != "町発注" else PatternFill("lightGray", fgColor="808080")
            for f, sd in s.stop.items():
                if g == "%d号炉" % f and d == sd and nm == "点検整備": cell.value = "▼"
            for f, su in s.startup.items():
                if g == "%d号炉" % f and d == su and nm == "点検整備": cell.value = "★"
        put(ws, r, 3 + len(days), nt, F(8), AL, B_ALL)
        rec("工程表", r, "item", {"号炉": g, "項目": nm.strip(), "予定": span_txt(spans), "備考": nt}, wk)
        r += 1
    r += 1
    put(ws, r, 2, "▼ 停止（前日夜 立下げ）　★ 立上げ　日曜は休工", F(8))
    page(ws, landscape=True, paper=9, area="A1:%s%d" % (CL(3 + len(days)), r))
    return save(wb, fname)

TEAM_FILL = {"機械A": "8EA9DB", "機械B": "A9D08E", "電計": "F4B084", "SV": "FFD966", "町": "D9D9D9", "全員": "B4C6E7", "別途": "FFFFFF",
             "足場": "C9C9C9", "耐火": "D5A6BD", "クレーン": "9BC2E6"}

def koutei_v1(c, s, fy=None):
    fy = fy or c.fy
    fname = "【架空】しおさい町ごみ処理場_工程表_%s_水樹プラント実施工程表様式.xlsx" % label(c, fy)
    rec_open(fname)
    wb = new_wb("【架空】" + c.name + " 実施工程表", "水樹プラント v1", dt.datetime(fy, 8 if c.kind == "EN" else 9, 20, 15, 0))
    ws = wb.active; ws.title = "実施工程表"
    d0 = min(min(t.plan[0][0] for t in s.tasks), s.site[0] - dt.timedelta(days=1))
    d1 = max(max(t.plan[-1][1] for t in s.tasks), s.site[1])
    days = [d0 + dt.timedelta(days=i) for i in range((d1 - d0).days + 1)]
    wide = len(days) > 30
    c0 = 5
    widths(ws, [8, 38, 7, 5] + [(2.3 if wide else 4.4)] * len(days) + [18])
    merge(ws, 1, 1, 1, c0 + len(days), "【架空】" + c.name + ("（%s）" % W.wareki_long(fy) if c.kind == "EN" else "") + "　実施工程表", F(14, True), AC, None)
    put(ws, 2, 1, "発注者：しおさい町　　受注者：%s　現場責任者（SV）v1　　現地作業：%d年%d月%d日（%s）〜%d月%d日（%s）" % (
        W.MAKER, s.site[0].year, s.site[0].month, s.site[0].day, YOUBI[s.site[0].weekday()], s.site[1].month, s.site[1].day, YOUBI[s.site[1].weekday()]), F(9))
    if c.subcontract: put(ws, 2, 30 if wide else 20, "協力会社：" + c.subcontract, F(9))
    put(ws, 3, 1, "▼停止　★立上げ　■作業（班の色）　▨別途工事（町発注・参考）　作業時間 8:00〜17:00　日曜休工", F(8))
    r = 5
    for i, h in enumerate(["区分", "作業項目", "班", "人員"], 1):
        merge(ws, r, i, r + 1, i, h, F(9, True), AC, B_ALL, fill("D9E1F2"))
    for j, d in enumerate(days):
        lab = "%d/%d" % (d.month, d.day) if (not wide or d.day == 1 or j == 0 or d.weekday() == 0) else d.day
        put(ws, r, c0 + j, lab, F(6 if wide else 8, True), AC, B_ALL, fill("D9E1F2"))
        wd = YOUBI[d.weekday()]
        put(ws, r + 1, c0 + j, wd, F(6 if wide else 8, color="C00000" if wd == "日" else None), AC, B_ALL, fill("FCE4D6") if wd == "日" else fill("D9E1F2"))
    merge(ws, r, c0 + len(days), r + 1, c0 + len(days), "備考", F(9, True), AC, B_ALL, fill("D9E1F2"))
    r += 2
    if c.kind == "MK" and len(s.stop) == 2:
        put(ws, r, 1, "運転", F(8, True), AC, B_ALL); put(ws, r, 2, "1号炉／2号炉 運転状況", F(8), AL, B_ALL)
        put(ws, r, 3, "", None, AC, B_ALL); put(ws, r, 4, "", None, AC, B_ALL)
        for j, d in enumerate(days):
            a = "停" if s.stop[1] < d <= s.startup[1] else "運"
            b = "停" if s.stop[2] < d <= s.startup[2] else "運"
            if d == s.stop[1]: a = "▼"
            if d == s.stop[2]: b = "▼"
            if d == s.startup[1]: a = "★"
            if d == s.startup[2]: b = "★"
            put(ws, r, c0 + j, "%s/%s" % (a, b), F(6), AC, B_ALL)
        put(ws, r, c0 + len(days), "1号/2号", F(7), AL, B_ALL)
        r += 1
    order = ["共通", "1号炉", "共通設備", "2号炉", "ごみクレーン"]
    groups = [g for g in order if any(t.group == g for t in s.tasks)]
    for gname in groups:
        r0 = r
        for t in [t for t in s.tasks if t.group == gname]:
            put(ws, r, 2, t.name, F(8, color="7F7F7F" if t.team in ("別途", "町") else None), AL, B_HAIR)
            put(ws, r, 3, t.team, F(8), AC, B_HAIR)
            put(ws, r, 4, t.crew, F(8), AC, B_HAIR)
            for j, d in enumerate(days):
                cell = put(ws, r, c0 + j, None, F(9, True), AC, Border(left=hair, right=hair, top=hair, bottom=hair))
                if d.weekday() == 6: cell.fill = fill("FCE4D6")
                for (a, b) in t.plan:
                    if a <= d <= b:
                        if t.team == "別途": cell.fill = PatternFill("darkUp", fgColor="808080", bgColor="FFFFFF")
                        else: cell.fill = fill(TEAM_FILL.get(t.team, "BDD7EE"))
                        if "停止" in t.name and d == a: cell.value = "▼"
                        if "立上げ" in t.name: cell.value = "★"
            nt = ""
            if "受変電" in t.name: nt = "みさき電気管理事務所"
            put(ws, r, c0 + len(days), nt, F(7), AL, B_HAIR)
            rec("実施工程表", r, "item", {"区分": gname, "作業項目": t.name, "班": t.team, "人員": t.crew, "予定": span_txt(t.plan)}, t.works)
            r += 1
        merge(ws, r0, 1, r - 1, 1, gname, F(9, True), AC, B_ALL)
    r += 1
    teams = sorted({t.team for t in s.tasks if t.team not in ("町", "別途", "全員")})
    put(ws, r, 2, "班編成：" + "、".join("%s %s名" % (tm, max((t.crew for t in s.tasks if t.team == tm and t.crew != ""), default="")) for tm in teams), F(8))
    ws.freeze_panes = "E7"
    page(ws, landscape=True, paper=8, area="A1:%s%d" % (CL(c0 + len(days)), r))
    return save(wb, fname)

def g1_tasks(c, s):
    if c.key == "H30-ASH": return [(t.name, t.plan, t.works, t.ref) for t in s.tasks]
    out = [("契約・施工計画書提出", [(c.cdate, wd_after2(c.cdate, 2))], [], False),
           ("材料手配", [(wd_after2(c.cdate, 1), s.site[0] - dt.timedelta(days=1))], [], False)]
    if c.in_mk_ref:
        out.append(c.in_mk_ref)
    # 作業を日ごとにまとめる
    for h in s.htasks:
        ds = sorted({x[0] for x in h["plan"] or h["act"]})
        out.append((h["name"], [(ds[0], ds[-1])], h["works"], False))
    out.append(("完了届・写真整理", [(wd_after2(s.site[1], 1), c.kanryo)], [], False))
    out.append(("完了検査（町）", [(c.kensa, c.kensa)], [], False))
    return out

def wd_after2(d, n): return cw.wd_after(d, n)

def koutei_g1(c, s):
    fname = "【架空】しおさい町ごみ処理場_工程表_%s_浜風機工g1日単位様式.xlsx" % label(c)
    rec_open(fname)
    wb = new_wb("【架空】" + c.name + " 工程表", "浜風機工 g1", dt.datetime(c.cdate.year, c.cdate.month, c.cdate.day, 17, 30))
    ws = wb.active; ws.title = "工程表"
    tasks = g1_tasks(c, s)
    a = min(sp[0][0] for _, sp, _, _ in tasks); b = max(sp[-1][1] for _, sp, _, _ in tasks)
    a = c.cdate if c.cdate < a else a
    n = max(21, (b - a).days + 3)
    days = [a + dt.timedelta(days=i) for i in range(n)]
    widths(ws, [4, 34] + [3.2] * len(days) + [10])
    merge(ws, 1, 1, 1, 2 + len(days) + 1, "工　程　表", F(16, True), AC, None)
    put(ws, 3, 1, "工事名：　【架空】" + c.name, F(11))
    put(ws, 4, 1, "工　期：　%s 〜 %s" % (jdate_short(c.start), jdate_short(c.end)), F(11))
    put(ws, 2, max(14, len(days) - 6), "【架空】有限会社 浜風機工", F(10))
    put(ws, 4, max(14, len(days) - 6), "現場代理人　g1　㊞", F(10))
    r = 6
    merge(ws, r, 1, r + 1, 1, "No", F(9, True)); merge(ws, r, 2, r + 1, 2, "作　業　項　目", F(10, True))
    for j, d in enumerate(days):
        lab = ("%d/%d" % (d.month, d.day)) if (j == 0 or d.day == 1) else str(d.day)
        put(ws, r, 3 + j, lab, F(7), AC, B_ALL)
        wdj = YOUBI[d.weekday()]
        put(ws, r + 1, 3 + j, wdj, F(7, color="C00000" if wdj in "土日" else None), AC, B_ALL)
    merge(ws, r, 3 + len(days), r + 1, 3 + len(days), "備考", F(9, True))
    r += 2
    for i, (task, spans, wk, ref) in enumerate(tasks, 1):
        put(ws, r, 1, i, F(9), AC, B_ALL)
        put(ws, r, 2, task, F(9, color="808080" if ref else None), AL, B_ALL)
        for j, d in enumerate(days):
            cell = put(ws, r, 3 + j, None, None, AC, B_ALL)
            if d.weekday() == 6: cell.fill = fill("F2F2F2")
            for (x, y) in spans:
                if x <= d <= y: cell.fill = fill("BFBFBF") if ref else fill("595959")
        nt = "町 立会" if "完了検査" in task else ("参考" if ref else "")
        put(ws, r, 3 + len(days), nt, F(8), AL, B_ALL)
        ws.row_dimensions[r].height = 20
        rec("工程表", r, "item", {"No": i, "作業項目": task, "予定": span_txt(spans), "備考": nt}, wk)
        r += 1
    for _ in range(2):
        put(ws, r, 1, None, None, AC, B_ALL)
        for cc in range(2, 4 + len(days)): ws.cell(row=r, column=cc).border = B_ALL
        ws.row_dimensions[r].height = 20; r += 1
    r += 1
    put(ws, r, 2, "※日曜日は休み。" + ("現地作業はメーカー定期点検の停止期間内に行う。" if c.in_mk_ref else ""), F(9))
    page(ws, landscape=True, paper=9, area="A1:%s%d" % (CL(3 + len(days)), r))
    return save(wb, fname)

def jdate_short(d):
    if d >= D(2019, 5, 1): return "令和%d年%d月%d日" % (d.year - 2018, d.month, d.day)
    return "平成%d年%d月%d日" % (d.year - 1988, d.month, d.day)

def koutei_g2(c, s):
    fname = "【架空】しおさい町ごみ処理場_工程表_%s_浜風機工g2時間単位様式.xlsx" % label(c)
    rec_open(fname)
    wb = new_wb("【架空】" + c.name + " 工程表", "浜風機工 g2", dt.datetime(c.cdate.year, c.cdate.month, c.cdate.day, 18, 0), "g2")
    ws = wb.active; ws.title = "工程表（時間）"
    sdays = sorted({x[0] for h in s.htasks for x in (h["plan"] + h["act"])})
    if c.key == "R4-CH":
        slots = [(D(2022, 10, d), h) for (d, h0, h1) in sg.S.CH_SLOTS for h in range(h0, h1)]
    else:
        slots = []
        for d in sdays:
            last = max([x[2] for h in s.htasks for x in (h["plan"] + h["act"]) if x[0] == d] + [17])
            slots += [(d, h) for h in range(8, last)]
    c0 = 4
    widths(ws, [4, 32, 5] + [2.9] * len(slots) + [16])
    put(ws, 1, 1, "【架空】" + c.name + "　工程表（時間単位）", F(14, True))
    put(ws, 2, 1, "工期 %s〜%s　現地 %d月%d日〜%d月%d日" % (jdate_short(c.start), jdate_short(c.end), s.site[0].month, s.site[0].day, s.site[1].month, s.site[1].day)
        + ("　（メーカー定期点検の停止中に行う）" if c.in_mk_ref else ""), F(9))
    put(ws, 4, 2, "【架空】有限会社 浜風機工　所長 g2", F(9))
    put(ws, 3, 1, "上段＝予定（網掛け）　下段＝実績（塗り）　実績は日々記入　A班：職長＋作業員2" + ("　B班：溶接工1" if any(h["team"] == "B" for h in s.htasks) else ""), F(8))
    r = 5
    merge(ws, r, 1, r + 2, 1, "No", F(9, True)); merge(ws, r, 2, r + 2, 2, "作業", F(9, True)); merge(ws, r, 3, r + 2, 3, "班", F(9, True))
    j = 0
    for d in sorted({x[0] for x in slots}):
        hs = [h for (dd, h) in slots if dd == d]
        merge(ws, r, c0 + j, r, c0 + j + len(hs) - 1, "%d/%d（%s）" % (d.month, d.day, YOUBI[d.weekday()]), F(8, True), AC, B_ALL, fill("DDEBF7"))
        for k, h in enumerate(hs):
            put(ws, r + 1, c0 + j + k, h, F(7), AC, B_ALL); put(ws, r + 2, c0 + j + k, None, None, AC, B_ALL)
        j += len(hs)
    merge(ws, r, c0 + len(slots), r + 2, c0 + len(slots), "記事", F(9, True))
    r += 3
    def col_of(d, h): return c0 + slots.index((d, h)) if (d, h) in slots else None
    for i, h in enumerate(s.htasks, 1):
        task = h["name"]; ext = task.startswith("〔停電") or "メーカー" in task
        merge(ws, r, 1, r + 1, 1, i, F(9))
        merge(ws, r, 2, r + 1, 2, task, F(9, color="C00000" if h["extra"] else ("7F7F7F" if ext else None)), AL)
        merge(ws, r, 3, r + 1, 3, h["team"], F(9))
        for k in range(len(slots)):
            put(ws, r, c0 + k, None, None, AC, Border(left=hair, right=hair, top=thin, bottom=hair))
            put(ws, r + 1, c0 + k, None, None, AC, Border(left=hair, right=hair, top=hair, bottom=thin))
        for (d, a, b) in h["plan"]:
            for hh in range(a, b):
                cc = col_of(d, hh)
                if cc: ws.cell(row=r, column=cc).fill = PatternFill("lightGray", fgColor="808080")
        for (d, a, b) in h["act"]:
            for hh in range(a, b):
                cc = col_of(d, hh)
                if cc: ws.cell(row=r + 1, column=cc).fill = fill("C00000" if h["extra"] else ("A6A6A6" if ext else "2F5597"))
        nt = ""
        if c.key == "R4-CH":
            nt = {"〔追加〕テール軸受取替": "10/20 14時 軸受損傷発見。町 t2 承認", "新チェーン据付・スクレーパ取付": "追加作業のため後ろへ"}.get(task, "")
            if "1号炉" in task: nt = "着火 予定14時→16時"
        elif h["extra"]:
            nt = "現地で判明。町 %s 承認" % ("t2" if c.fy <= 2022 else "t3")
        merge(ws, r, c0 + len(slots), r + 1, c0 + len(slots), nt, F(7), AL)
        pl = ["%s %d-%d" % (jd(d), a, b) for (d, a, b) in h["plan"]]
        ac = ["%s %d-%d" % (jd(d), a, b) for (d, a, b) in h["act"]]
        rec("工程表（時間）", r, "item", {"No": i, "作業": task, "班": h["team"], "予定": pl, "実績": ac, "記事": nt}, h["works"])
        r += 2
    r += 1
    put(ws, r, 2, "※作業時間 8〜17時（12〜13時 休憩）。", F(8))
    ws.freeze_panes = "D8"
    page(ws, landscape=True, paper=8 if len(slots) > 20 else 9, area="A1:%s%d" % (CL(c0 + len(slots)), r))
    return save(wb, fname)

# ===================== 日報 =====================
def nippo_v1(c, s, fy=None):
    fy = fy or c.fy
    fname = "【架空】しおさい町ごみ処理場_作業日報_%s_水樹プラント一日一枚様式.xlsx" % label(c, fy)
    rec_open(fname)
    wb = new_wb("【架空】" + c.name + " 作業日報", "水樹プラント v1", dt.datetime(s.site[0].year, s.site[0].month, s.site[0].day, 18, 0))
    first = True
    town = "t1" if fy <= 2017 else ("t2" if fy <= 2022 else "t3")
    for n, day in enumerate(s.days, 1):
        d = day.date
        ws = wb.active if first else wb.create_sheet()
        first = False
        ws.title = "%02d%02d" % (d.month, d.day)
        widths(ws, [8, 7, 52, 6, 6, 8])
        merge(ws, 1, 1, 1, 6, "作 業 日 報", F(16, True), AC, None)
        put(ws, 2, 1, "【架空】" + c.name, F(9))
        put(ws, 3, 1, "日付", F(9, True), AC, B_ALL, fill("D9E1F2"))
        merge(ws, 3, 2, 3, 3, "%d年%d月%d日（%s）　No.%d" % (d.year, d.month, d.day, YOUBI[d.weekday()], n), F(10), AL)
        put(ws, 3, 4, "天候", F(9, True), AC, B_ALL, fill("D9E1F2"))
        merge(ws, 3, 5, 3, 6, day.wx, F(10), AC)
        put(ws, 4, 1, "作業時間", F(9, True), AC, B_ALL, fill("D9E1F2"))
        ot = any(l.hours > 8 for l in day.lines)
        merge(ws, 4, 2, 4, 6, "8:00〜17:00（休憩 12:00〜13:00）" + ("　SV 18:00 まで" if ot else ""), F(9), AL)
        r = 6
        for i, h in enumerate(["班", "号炉", "作業内容", "人員", "時間", "工数\n(人・h)"], 1):
            put(ws, r, i, h, F(9, True), AC, B_ALL, fill("D9E1F2"))
        ws.row_dimensions[r].height = 26
        r += 1
        tot = 0; men = {}
        for l in day.lines:
            g = l.gou
            if n >= 11 and fy >= 2018 and g in ("1号", "2号"): g = "#" + g[0]
            vals = [l.team, g, l.text, l.crew, l.hours, l.crew * l.hours]
            for i, v in enumerate(vals, 1):
                put(ws, r, i, v, F(9), AL if i == 3 else (AR if i == 6 else AC), B_HAIR, fmt="0.0" if i == 6 else None)
            ws.row_dimensions[r].height = 30
            mh = l.crew * l.hours
            rec(ws.title, r, "item", {"日付": jd(d), "No": n, "天候": day.wx, "班": l.team, "号炉": g, "作業内容": l.text, "人員": l.crew, "時間": l.hours,
                                      "manDays": round(mh / 8.0, 2), "manDaysText": "%.1f" % mh}, l.works)
            tot += mh; men[l.team] = max(men.get(l.team, 0), l.crew)
            r += 1
        for _ in range(max(0, 7 - len(day.lines))):
            for cc in range(1, 7): put(ws, r, cc, None, None, AC, B_HAIR)
            ws.row_dimensions[r].height = 20; r += 1
        box(ws, 7, 1, r - 1, 6, B_HAIR)
        put(ws, r, 3, "日　計", F(9, True), AR, B_ALL)
        put(ws, r, 4, sum(men.values()), F(9, True), AC, B_ALL)
        put(ws, r, 5, "", None, AC, B_ALL)
        put(ws, r, 6, tot, F(9, True), AR, B_ALL, fmt="0.0")
        for cc in (1, 2): put(ws, r, cc, None, None, AC, B_ALL)
        rec(ws.title, r, "total", {"日付": jd(d), "人員": sum(men.values()), "manDays": round(tot / 8.0, 2), "manDaysText": "%.1f" % tot})
        r += 2
        merge(ws, r, 2, r + 2, 6, day.memo, F(9), ALT)
        merge(ws, r, 1, r + 2, 1, "記事", F(9, True), AC, B_ALL, fill("D9E1F2"))
        r += 4
        put(ws, r, 4, "SV", F(8), AC, B_ALL); put(ws, r, 5, "町", F(8), AC, B_ALL)
        put(ws, r + 1, 4, "v1", F(9), AC, B_ALL); put(ws, r + 1, 5, town, F(9), AC, B_ALL)
        ws.row_dimensions[r + 1].height = 30
        put(ws, r + 2, 1, W.MAKER + "　しおさい町ごみ処理場 作業所", F(8))
        page(ws, landscape=False, paper=9, area="A1:F%d" % (r + 2))
    return save(wb, fname)

def nippo_g1(c, s):
    who = lname(c)
    fname = "【架空】しおさい町ごみ処理場_作業日報_%s_浜風機工g1一日一枚様式.xlsx" % label(c)
    rec_open(fname)
    wb = new_wb("【架空】" + c.name + " 作業日報", "浜風機工 g1", dt.datetime(s.site[0].year, s.site[0].month, s.site[0].day, 18, 0), who)
    town = "t1" if c.fy <= 2017 else "t2"
    first = True
    for i, day in enumerate(s.days, 1):
        d = day.date
        ws = wb.active if first else wb.create_sheet()
        first = False
        ws.title = "%d.%d" % (d.month, d.day)
        widths(ws, [12, 14, 14, 14, 14, 14])
        merge(ws, 1, 1, 1, 6, "作　業　日　報", F(18, True), AC, None)
        put(ws, 2, 5, "No. %d" % i, F(10), AR)
        if hasattr(day, "items"):
            items, crew, mach, ky, tm = day.items, day.crew, day.mach, day.ky, day.tm
        else:
            items = [t for t in day.lines[0].text.split("、") if t]
            h = day.lines[0].hours
            crew = [("%s（職長）" % who, h), ("作業員", h), ("作業員", h)] + ([("溶接工", h)] if getattr(day, "weld", 0) else [])
            mach = ["チェーンブロック1t", "ディスクグラインダ"] + (["エンジン溶接機"] if getattr(day, "weld", 0) else []) + (["2tトラック"] if i == 1 else [])
            ky = ["KY：重量物の吊り込み、挟まれ", "KY：電源ロック・表示確認", "KY：足元すべり、転倒", "KY：回転体 巻込まれ（試運転時 立入禁止）"][(d.toordinal()) % 4]
            tm = "%d:00〜%d:00" % (getattr(day, "t0", 8), min(getattr(day, "t1", 17), 18))
        rows = [("工 事 名", "【架空】" + c.name), ("日　　付", jdate_short(d) + "（%s）" % YOUBI[d.weekday()]), ("天　　候", day.wx), ("作業時間", tm)]
        r = 3
        for k, v in rows:
            put(ws, r, 1, k, F(10, True), AC, B_ALL, fill("F2F2F2")); merge(ws, r, 2, r, 6, v, F(10), AL)
            ws.row_dimensions[r].height = 22; r += 1
        r += 1
        merge(ws, r, 1, r, 6, "作　業　内　容", F(11, True), AC, B_ALL, fill("F2F2F2")); r += 1
        for it in items + [""] * max(0, 8 - len(items)):
            merge(ws, r, 1, r, 6, ("・" + it) if it else "", F(10), AL, Border(left=thin, right=thin, top=dot, bottom=dot))
            if it: rec(ws.title, r, "item", {"日付": jd(d), "No": i, "天候": day.wx, "作業時間": tm, "作業内容": it}, day.lines[0].works)
            ws.row_dimensions[r].height = 22; r += 1
        if day.memo:
            merge(ws, r - 1, 1, r - 1, 6, "※" + day.memo, F(9, color="C00000"), AL, Border(left=thin, right=thin, top=dot, bottom=dot))
            rec(ws.title, r - 1, "item", {"日付": jd(d), "記事": day.memo})
        r += 1
        put(ws, r, 1, "作 業 員", F(10, True), AC, B_ALL, fill("F2F2F2"))
        for cc, h in zip((2, 3, 4, 5), ("氏名", "時間", "氏名", "時間")):
            put(ws, r, cc, h, F(9, True), AC, B_ALL, fill("F2F2F2"))
        put(ws, r, 6, "", None, AC, B_ALL, fill("F2F2F2")); r += 1
        for k in range(0, 4, 2):
            put(ws, r, 1, "", None, AC, B_ALL)
            for jj in range(2):
                if k + jj < len(crew):
                    nm, h = crew[k + jj]
                    ht = "%sh" % (("%g" % h))
                    put(ws, r, 2 + jj * 2, nm, F(10), AC, B_ALL); put(ws, r, 3 + jj * 2, ht, F(10), AC, B_ALL)
                    rec(ws.title, r, "item", {"日付": jd(d), "氏名": nm, "時間": ht, "manDays": round(h / 8.0, 3), "manDaysText": ht}, day.lines[0].works)
                else:
                    put(ws, r, 2 + jj * 2, "", None, AC, B_ALL); put(ws, r, 3 + jj * 2, "", None, AC, B_ALL)
            put(ws, r, 6, "", None, AC, B_ALL)
            ws.row_dimensions[r].height = 22; r += 1
        r += 1
        put(ws, r, 1, "使用機械", F(10, True), AC, B_ALL, fill("F2F2F2")); merge(ws, r, 2, r, 6, "、".join(mach), F(10), AL); r += 1
        put(ws, r, 1, "安全（KY）", F(10, True), AC, B_ALL, fill("F2F2F2")); merge(ws, r, 2, r, 6, ky, F(10), AL)
        ws.row_dimensions[r].height = 30; r += 2
        put(ws, r, 4, "現場代理人", F(9), AC, B_ALL); put(ws, r, 5, "町確認", F(9), AC, B_ALL); r += 1
        put(ws, r, 4, who, F(10), AC, B_ALL); put(ws, r, 5, town if not (c.key == "H30-ASH" and d.day == 23) else "", F(10), AC, B_ALL)
        ws.row_dimensions[r].height = 36
        put(ws, r + 2, 1, "【架空】有限会社 浜風機工", F(9))
        page(ws, landscape=False, paper=9, area="A1:F%d" % (r + 2))
    return save(wb, fname)

def nippo_g2(c, s):
    fname = "【架空】しおさい町ごみ処理場_作業日報_%s_浜風機工g2工事別一枚様式.xlsx" % label(c)
    rec_open(fname)
    wb = new_wb("【架空】" + c.name + " 作業日報", "浜風機工 g1", dt.datetime(2018, 10, 23, 18, 0), "g2")
    ws = wb.active; ws.title = "日報"
    widths(ws, [11, 4, 5, 11, 50, 5, 5, 5, 7, 6, 30, 7])
    town = "t2" if c.fy <= 2022 else "t3"
    put(ws, 1, 1, "【架空】作業日報（工事別）", F(14, True))
    put(ws, 2, 1, "工事名　【架空】" + c.name, F(10))
    loc = "、".join(sorted({(("%d号炉 " % W.BYMID[w.mid].furnace) if W.BYMID[w.mid].furnace else "") + W.BYMID[w.mid].name for w in c.works}))
    put(ws, 3, 1, "工期　%s〜%s　　施工場所　%s" % (jdate_short(c.start), jdate_short(c.end), "灰出し室（灰出しコンベヤNo.1）" if c.key == "R4-CH" else loc), F(9))
    put(ws, 2, 11, "【架空】有限会社 浜風機工　所長 g2", F(9))
    r = 5
    hdr1 = ["月日", "曜", "天候", "作業時間", "作業内容", "人数", None, None, "延べ\n時間", "進捗\n％", "記事（発見・変更）", "町確認"]
    for i, h in enumerate(hdr1, 1):
        if i in (6, 7, 8): continue
        merge(ws, r, i, r + 1, i, h, F(9, True), AC, B_ALL, fill("DDEBF7"))
    merge(ws, r, 6, r, 8, "人数", F(9, True), AC, B_ALL, fill("DDEBF7"))
    for i, h in zip((6, 7, 8), ("職長", "作業員", "溶接")):
        put(ws, r + 1, i, h, F(8, True), AC, B_ALL, fill("DDEBF7"))
    r += 2
    tot = 0
    n = len(s.days)
    for k, day in enumerate(s.days, 1):
        d = day.date
        if hasattr(day, "tm"):
            tm, a, b, wl, hrs, prog = day.tm, day.a, day.b, day.w, day.hrs, day.prog
            txt = day.lines[0].text
        else:
            l = day.lines[0]
            t0, t1 = getattr(day, "t0", 8), getattr(day, "t1", 17)
            tm = "%d:00-%d:00" % (t0, t1)
            a, b, wl = 1, 2, (1 if getattr(day, "weld", 0) else 0)
            hrs = float(l.hours * (a + b + wl))
            prog = int(round(100.0 * k / n))
            txt = l.text
        vals = ["%d/%d" % (d.month, d.day), YOUBI[d.weekday()], day.wx, tm, txt, a, b, wl or "", hrs, prog, day.memo, town]
        for i, v in enumerate(vals, 1):
            put(ws, r, i, v, F(9 if i in (3, 4, 5, 11, 12) else 10, color="C00000" if (i == 11 and "追加" in str(v)) else None),
                AL if i in (5, 11) else AC, B_ALL, fmt="0.0" if i == 9 else None)
        ws.row_dimensions[r].height = 48
        rec("日報", r, "item", {"月日": jd(d), "曜": YOUBI[d.weekday()], "天候": day.wx, "作業時間": tm, "作業内容": txt, "職長": a, "作業員": b, "溶接": wl,
                                "延べ時間": hrs, "進捗": prog, "記事": day.memo, "町確認": town, "manDays": round(hrs / 8.0, 2), "manDaysText": "%.1f" % hrs}, day.lines[0].works)
        tot += hrs; r += 1
    for _ in range(2):
        for cc in range(1, 13): put(ws, r, cc, None, None, AC, B_ALL)
        ws.row_dimensions[r].height = 30; r += 1
    put(ws, r, 5, "計", F(10, True), AR, B_ALL)
    put(ws, r, 9, tot, F(10, True), AC, B_ALL, fmt="0.0")
    rec("日報", r, "total", {"延べ時間": tot, "manDays": round(tot / 8.0, 2), "manDaysText": "%.1f" % tot})
    r += 2
    mach = "エンジン溶接機（10/20）、チェーンブロック1t×2、レバーブロック、ディスクグラインダ" if c.key == "R4-CH" else (
        "チェーンブロック1t、ディスクグラインダ" + ("、エンジン溶接機" if any(getattr(d, "weld", 0) for d in s.days) else ""))
    put(ws, r, 1, "使用機械：" + mach, F(9)); r += 1
    put(ws, r, 1, "延べ時間は休憩を除く実働。", F(8))
    page(ws, landscape=True, paper=9, area="A1:L%d" % r)
    return save(wb, fname)
