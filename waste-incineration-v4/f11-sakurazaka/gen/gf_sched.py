# -*- coding: utf-8 -*-
"""工程表（本作り）：e1 旧様式 R1〜R2（年間 1 枚・月の下に第1〜5週・━ の線。補修工事は日の欄に ━ の線・予定のみ）
→ R3〜 新様式（1 工事 1 枚・日単位・予定/実績 2 段の塗り）"""
import os, datetime as dt
from openpyxl import Workbook
from openpyxl.comments import Comment
from common import *
import world as W
import fullworld as FW

WD = FW.WD
MONTHS = [4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3]

def nendo(y):
    return W.nendo_name(y)

def fname(job, style):
    return "【架空】さくら坂焼却場_工程表_%s%s_%s.xlsx" % (nendo(job.fy), job.c["short"], style)

def iso(d):
    return d.isoformat()

# ---------------- 旧：年間・週 ----------------
def build_old_T(job):
    y = job.fy
    recs = []
    wb = Workbook(); ws = wb.active; ws.title = "R%d年間工程" % (y - 2018) if y > 2019 else "R1年間工程"
    cell(ws, "A1", "【架空】さくら坂焼却場　%s　定期整備業務委託　年間工程表" % nendo(y), F(14, True))
    cell(ws, "A2", "【架空】さくら坂機工株式会社　所長 e1　　%s4月1日作成" % W.wareki_k(y).replace("年", "年") , F(9))
    cell(ws, "A3", "▼停止　▲立上げ　━作業　┅別途工事　☆検査　■全停電　○月例　　欄は各月の第1〜5週（1〜7日が第1週）", F(9))
    setw(ws, {"A": 7, "B": 40})
    for i, m in enumerate(MONTHS):
        c0 = 3 + i * 5
        ws.merge_cells(start_row=5, start_column=c0, end_row=5, end_column=c0 + 4)
        cell(ws, "%s5" % get_column_letter(c0), "%d月" % m, F(9, True), CENTER, B_ALL, fill("E2EFDA"))
        for k in range(5):
            col = get_column_letter(c0 + k)
            cell(ws, "%s6" % col, str(k + 1), F(7), CENTER, B_ALL)
            ws.column_dimensions[col].width = 2.3
            if k: cell(ws, "%s5" % col, None, F(9), CENTER, B_ALL)
    ws.merge_cells("A5:A6"); ws.merge_cells("B5:B6")
    cell(ws, "A5", "炉", F(9, True), CENTER, B_ALL, fill("E2EFDA")); cell(ws, "B5", "作業項目", F(9, True), CENTER, B_ALL, fill("E2EFDA"))
    cell(ws, "A6", None, F(9), CENTER, B_ALL); cell(ws, "B6", None, F(9), CENTER, B_ALL)
    def colof(d):
        return 3 + MONTHS.index(d.month) * 5 + (d.day - 1) // 7
    rows = []
    for sname, (a, b), tasks in job.sheets:
        for t in tasks:
            unit = "共通" if t.unit == "共通" else sname
            mk = {"stop": "▼", "start": "▲", "binsp": "☆", "c_elec": "■", "crane": "☆" if y in FW.CRANE_PERF else "━"}.get(t.kind, "━")
            rows.append((unit, t, mk))
    # 別途工事（補修）の行
    for c in W.CONTRACTS:
        if c["fy"] == y and c["id"].startswith("R") and c["id"] != "R2020b":
            rj = FW.build_repair(c["id"])
            ts = rj.sheets[0][2]
            fake = FW.Task("—", "（別途）" + c["short"], "B班", ts[0].unit, [], "other", sorted({d for t in ts for d in t.plan}), [])
            ul = {"1": "1号炉", "2": "2号炉"}.get(ts[0].unit, "共通")
            rows.append((ul, fake, "┅"))
    rows.append(("共通", FW.Task("—", "月例点検（各設備、運転委託と共同）", "A班", "共通", [], "monthly", [dt.date(y if m >= 4 else y + 1, m, 15) for m in (4, 5, 7, 8, 9, 12, 1, 2, 3)], []), "○"))
    rows.sort(key=lambda x: {"1号炉": 0, "2号炉": 1, "共通": 2}[x[0]])
    r = 7; blk = None; bs = r
    for unit, t, mk in rows:
        if unit != blk:
            if blk is not None and r - 1 > bs:
                ws.merge_cells(start_row=bs, start_column=1, end_row=r - 1, end_column=1)
            blk, bs = unit, r
        cell(ws, "A%d" % r, unit if r == bs else None, F(9, True), CENTER, B_ALL)
        cell(ws, "B%d" % r, t.name, F(9, color="7F7F7F" if mk == "┅" else "000000"), LEFT, B_ALL)
        for c in range(3, 63):
            cell(ws, "%s%d" % (get_column_letter(c), r), None, F(9), CENTER,
                 Border(left=dot if (c - 3) % 5 else thin, right=dot if (c - 3) % 5 != 4 else thin, top=thin, bottom=thin))
        cols = sorted({colof(d) for d in t.plan})
        for c in cols:
            bar = mk in ("━", "┅")
            cell(ws, "%s%d" % (get_column_letter(c), r), mk, F(9 if bar else 10, True, {"☆": "375623", "■": "C00000", "▼": "1F4E79", "▲": "1F4E79", "┅": "7F7F7F"}.get(mk, "000000")),
                 Alignment(horizontal="fill", vertical="center") if bar else CENTER)
        weeks = ["%d月第%d週" % (d.month, (d.day - 1) // 7 + 1) for d in sorted(t.plan)]
        weeks = list(dict.fromkeys(weeks))
        recs.append(dict(sheet=ws.title, r=r, kind="item", fields={"炉": unit, "作業": t.name, "記号": mk, "予定": weeks,
                                                                     "予定日": [iso(d) for d in sorted(t.plan)], "task": t.no, "contract": job.cid if mk != "┅" else "別途"}))
        r += 1
    ws.merge_cells(start_row=bs, start_column=1, end_row=r - 1, end_column=1)
    r += 2
    s = FW.SHUT[y]
    cell(ws, "B%d" % r, "日程（予定）", F(9, True))
    notes = ["1号炉：%d/%d（%s）停止 → %d/%d（%s）立上げ" % (s["u1"][0].month, s["u1"][0].day, WD[s["u1"][0].weekday()], s["u1"][1].month, s["u1"][1].day, WD[s["u1"][1].weekday()]),
             "2号炉：%d/%d（%s）停止 → %d/%d（%s）立上げ　全停電 %d/%d〜%d/%d　1号炉も停止" % (s["u2"][0].month, s["u2"][0].day, WD[s["u2"][0].weekday()], s["u2"][1].month, s["u2"][1].day, WD[s["u2"][1].weekday()], s["zen"][0].month, s["zen"][0].day, s["zen"][1].month, s["zen"][1].day),
             "クレーン年次点検 %d/%d〜%d/%d%s" % (s["crane"][0].month, s["crane"][0].day, s["crane"][1].month, s["crane"][1].day, "、性能検査あり" if y in FW.CRANE_PERF else "、性能検査なし")]
    for i, t in enumerate(notes):
        cell(ws, "B%d" % (r + 1 + i), t, F(9))
    page(ws, "landscape", 9, 1, 1)
    props(wb, "【架空】さくら坂焼却場 %s 定期整備業務委託 年間工程表" % nendo(y), "【架空】さくら坂機工 e1")
    fn = os.path.join(OUT, fname(job, "年間週単位様式"))
    wb.save(fn)
    return fn, recs


def build_old_R(job):
    """旧様式の補修工事：日の欄に ━ の線（予定だけ）"""
    recs = []
    sname, (a, b), tasks = job.sheets[0]
    wb = Workbook(); ws = wb.active; ws.title = "工程"
    cell(ws, "A1", "【架空】さくら坂焼却場　%s　工程表" % job.c["short"], F(14, True))
    cell(ws, "A2", "工事件名：%s　　工期：%s〜%s　【架空】さくら坂機工株式会社 所長 e1" % (job.title, job.c["start"].strftime("%Y/%m/%d"), job.c["end"].strftime("%Y/%m/%d")), F(9))
    cell(ws, "A3", "━作業（予定）　☆検査", F(9))
    d0 = min(min(t.plan) for t in tasks) - dt.timedelta(3); d1 = max(max(t.plan) for t in tasks) + dt.timedelta(3)
    days = [d0 + dt.timedelta(i) for i in range((d1 - d0).days + 1)]
    setw(ws, {"A": 4, "B": 40})
    for j, d in enumerate(days):
        col = get_column_letter(3 + j)
        cell(ws, "%s5" % col, "%d/%d" % (d.month, d.day) if j == 0 else d.day, F(7, True), CENTER, B_ALL, fill("E2EFDA") if d.weekday() < 5 else fill("D9D9D9"))
        cell(ws, "%s6" % col, WD[d.weekday()], F(7), CENTER, B_ALL)
        ws.column_dimensions[col].width = 3.0
    cell(ws, "A5", "No", F(9, True), CENTER, B_ALL, fill("E2EFDA")); cell(ws, "B5", "作業項目", F(9, True), CENTER, B_ALL, fill("E2EFDA"))
    cell(ws, "A6", None, F(9), CENTER, B_ALL); cell(ws, "B6", None, F(9), CENTER, B_ALL)
    r = 7
    for t in tasks:
        cell(ws, "A%d" % r, t.no, F(9), CENTER, B_ALL)
        cell(ws, "B%d" % r, t.name, F(9), LEFT, B_ALL)
        for j, d in enumerate(days):
            c = cell(ws, "%s%d" % (get_column_letter(3 + j), r), None, F(9), CENTER, Border(left=dot, right=dot, top=thin, bottom=thin))
            if d in t.plan:
                c.value = "☆" if "検査" in t.name and "立会" in t.name else "━"
                c.alignment = Alignment(horizontal="fill", vertical="center") if c.value == "━" else CENTER
        recs.append(dict(sheet="工程", r=r, kind="item", fields={"No": t.no, "作業": t.name, "予定": [iso(d) for d in t.plan], "task": t.no, "contract": job.cid}))
        r += 1
    cell(ws, "B%d" % (r + 1), "※ 同じ停止の定期整備（A班）と炉内の作業時間を分ける。", F(8))
    page(ws, "landscape", 9, 1, 1)
    props(wb, "【架空】さくら坂焼却場 %s 工程表" % job.title, "【架空】さくら坂機工 e1")
    fn = os.path.join(OUT, fname(job, "日の欄に線様式"))
    wb.save(fn)
    return fn, recs


# ---------------- 新：日単位・予定/実績 ----------------
def daily_sheet(ws, title, job, d0, d1, tasks, crew_note, made, extra_rows=(), recs=None):
    cell(ws, "A1", "【架空】" + title, F(14, True))
    cell(ws, "A2", "工事件名：%s" % job.title, F(10))
    cell(ws, "A3", "工期：%s〜%s　　受注者：【架空】さくら坂機工株式会社（所長 e1）　%s" % (job.c["start"].strftime("%Y/%m/%d"), job.c["end"].strftime("%Y/%m/%d"), made), F(9))
    cell(ws, "A4", "上段＝予定（薄い塗り）　下段＝実績（濃い塗り）　灰色の列＝休工日", F(8))
    days = [d0 + dt.timedelta(i) for i in range((d1 - d0).days + 1)]
    for i, h in enumerate(["No", "作業項目", "班", "段"]):
        ws.merge_cells(start_row=6, start_column=1 + i, end_row=7, end_column=1 + i)
        cell(ws, "%s6" % get_column_letter(1 + i), h, F(9, True), CENTER, B_ALL, fill("DDEBF7"))
        cell(ws, "%s7" % get_column_letter(1 + i), None, F(9), CENTER, B_ALL)
    setw(ws, {"A": 4, "B": 38, "C": 5, "D": 4})
    for j, dd in enumerate(days):
        col = get_column_letter(5 + j)
        off = dd.weekday() >= 5
        cell(ws, "%s6" % col, "%d/%d" % (dd.month, dd.day) if (j == 0 or dd.day == 1) else dd.day, F(7, True), CENTER, B_ALL, fill("D9D9D9") if off else fill("DDEBF7"))
        cell(ws, "%s7" % col, WD[dd.weekday()], F(7, color="C00000" if dd.weekday() == 6 else "000000"), CENTER, B_ALL, fill("D9D9D9") if off else None)
        ws.column_dimensions[col].width = 3.3
    ci = 5 + len(days)
    cr = get_column_letter(ci)
    cell(ws, "%s6" % cr, "備考", F(9, True), CENTER, B_ALL, fill("DDEBF7"))
    ws.merge_cells("%s6:%s7" % (cr, cr)); ws.column_dimensions[cr].width = 32
    r = 8
    for t, color in tasks:
        act = [d for d in t.act if d <= FW.TODAY]
        for k, (lab, ds, shade) in enumerate((("予", t.plan, color[0]), ("実", act, color[1]))):
            cell(ws, "A%d" % r, t.no if k == 0 else None, F(8), CENTER, B_ALL)
            cell(ws, "B%d" % r, t.name if k == 0 else None, F(9), LEFT, B_ALL)
            cell(ws, "C%d" % r, t.team if k == 0 else None, F(8), CENTER, B_ALL)
            cell(ws, "D%d" % r, lab, F(7), CENTER, B_ALL)
            for j, dd in enumerate(days):
                c = cell(ws, "%s%d" % (get_column_letter(5 + j), r), None, F(7), CENTER,
                         Border(left=hair, right=hair, top=thin if k == 0 else hair, bottom=thin if k == 1 else hair), fill("EDEDED") if dd.weekday() >= 5 else None)
                if dd in ds:
                    c.fill = fill(shade)
            cell(ws, "%s%d" % (cr, r), t.memo if k == 0 else None, F(7), LEFT, B_ALL)
            r += 1
        for c in (1, 2, 3):
            ws.merge_cells(start_row=r - 2, start_column=c, end_row=r - 1, end_column=c)
        ws.merge_cells(start_row=r - 2, start_column=ci, end_row=r - 1, end_column=ci)
        if recs is not None:
            recs.append(dict(sheet=ws.title, r=r - 2, kind="item", fields={"No": t.no, "作業": t.name, "班": t.team, "予定": [iso(d) for d in sorted(t.plan)],
                                                                             "実績": [iso(d) for d in sorted(act)], "備考": t.memo, "task": t.no, "contract": job.cid}))
    for lab, vals in extra_rows:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
        cell(ws, "A%d" % r, lab, F(8, True), RIGHT, B_ALL, fill("F2F2F2"))
        for j, dd in enumerate(days):
            v = vals.get(dd)
            cell(ws, "%s%d" % (get_column_letter(5 + j), r), v if v else None, F(7), CENTER, B_ALL, fill("F2F2F2"))
        cell(ws, "%s%d" % (cr, r), None, F(7), LEFT, B_ALL, fill("F2F2F2"))
        if recs is not None:
            recs.append(dict(sheet=ws.title, r=r, kind="subtotal", fields={"項目": lab, "人数": {iso(k): v for k, v in vals.items()}}))
        r += 1
    if crew_note:
        cell(ws, "A%d" % (r + 1), crew_note, F(8))
    page(ws, "landscape", 9 if len(days) <= 32 else 8, 1, 1)


def build_new(job, rng):
    recs = []
    wb = Workbook(); wb.remove(wb.active)
    A = ("C6E0B4", "548235"); B = ("BDD7EE", "2F5597"); X = ("E7E6E6", "A6A6A6"); E = ("F8CBAD", "C00000")
    for sname, (a, b), tasks in job.sheets:
        ws = wb.create_sheet(sname[:28])
        lst = []
        for t in tasks:
            col = {"A班": A, "B班": B, "運転": X, "電気": E}.get(t.team, A)
            lst.append((t, col))
        # 定期整備のシートには同じ停止の別契約を灰色の行で
        if job.cid.startswith("T"):
            for c in W.CONTRACTS:
                if c["fy"] == job.fy and c["id"].startswith("R") and c["id"] != "R2020b":
                    rj = FW.build_repair(c["id"])
                    ts = rj.sheets[0][2]
                    if ("%s号炉" % ts[0].unit) == sname or (ts[0].unit == "共通" and sname == "2号炉"):
                        fake = FW.Task("別", "（別契約）" + c["short"], "B班", ts[0].unit, [], "other",
                                       sorted({d for t in ts for d in t.plan}), sorted({d for t in ts for d in t.act}), "別紙工程表")
                        lst.insert(1, (fake, X))
        d0 = min(min(t.plan + t.act) for t in tasks); d1 = max(max(t.plan + t.act) for t in tasks)
        extra = []
        if job.cid.startswith("R"):
            pn, an = {}, {}
            for t in tasks:
                for d in t.plan: pn[d] = pn.get(d, 0) + 1
            pn = {d: 4 + 2 * min(v, 2) - 1 for d, v in pn.items()}
            an = {dd["d"]: sum(k for _, k in dd["crew"]) for dd in job.days}
            extra = [("予定人数（B班）", pn), ("実績人数（B班）", an)]
        made = "%s 作成／%s 実績記入" % ((d0 - dt.timedelta(10)).strftime("%Y/%m/%d"), min(d1 + dt.timedelta(1), FW.TODAY).strftime("%Y/%m/%d"))
        title = "さくら坂焼却場　%s　工程表（%s）" % (job.c["short"], sname) if job.cid.startswith("T") else "さくら坂焼却場　%s　工程表" % job.c["short"]
        note = "A班：職長 f1、機械工 2〜5、電工・溶接工は日により、普通作業員 1〜3。" if job.cid.startswith("T") else "B班：職長 f2、機械工 2〜3、溶接工は日により、普通作業員 1〜2。"
        daily_sheet(ws, title, job, d0 - dt.timedelta(0), d1, lst, note, made, extra, recs)
    props(wb, "【架空】さくら坂焼却場 %s 工程表" % job.title, "【架空】さくら坂機工 e1")
    fn = os.path.join(OUT, fname(job, "日単位様式"))
    wb.save(fn)
    return fn, recs


def build(job, rng):
    if job.fy <= 2020:
        return build_old_T(job) if job.cid.startswith("T") else build_old_R(job)
    return build_new(job, rng)
