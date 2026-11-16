# -*- coding: utf-8 -*-
"""市の追加工事（市→鷹羽エンジ、請負）。工程表・工事日報・完成報告書はメーカー様式、契約一覧は市 環境課 b1 の Excel。
R3（2021）2号煙突内筒補修工事、R5（2023）2号ボイラ水管補修工事。どちらも SPC の 2号炉定修の炉停止期間内に施工。"""
import datetime as dt
import random
from openpyxl import Workbook
from openpyxl.utils import get_column_letter as L
import world as W
import works as WK
import contracts as K
from daily_src import CITY_DAY as R5_DAY, WEATHER, KY, TRADES, MACH
from xlhelp import *
from rec import Rec, path

R3_TASKS = [
    ("C01", "準備", "資機材搬入・点検口開放", "鷹羽", 3, (("10/12", "AM"), ("10/12", "AM")), None),
    ("C02", "仮設", "内筒内 ゴンドラ設置・安全設備", "鳶", 4, (("10/12", "PM"), ("10/13", "PM")), None),
    ("C03", "調査", "板厚測定・腐食範囲けがき（下部 GL+0〜8m）", "検査", 2, (("10/14", "AM"), ("10/14", "AM")), None),
    ("C04", "撤去", "腐食部 切断・撤去", "溶接", 4, (("10/14", "PM"), ("10/15", "PM")), None),
    ("C05", "補修", "当て板 SUS316L t6 取付・溶接", "溶接", 4, (("10/16", "AM"), ("10/19", "PM")), None),
    ("C06", "補修", "追加腐食部 当て板（0.8m²）", "溶接", 2, (("10/20", "AM"), ("10/20", "AM")), "追加"),
    ("C07", "検査", "溶接部 浸透探傷試験（PT）", "検査", 2, (("10/20", "PM"), ("10/20", "PM")), None),
    ("C08", "復旧", "外面保温 復旧", "保温", 3, (("10/21", "AM"), ("10/21", "PM")), None),
    ("C09", "復旧", "ゴンドラ撤去・片付け・引渡し", "鳶", 4, (("10/22", "AM"), ("10/23", "AM")), None),
]
R3_DAY = {
    "10/12": ("資機材搬入、2号煙突内筒 点検口開放・換気。午後 ゴンドラ 設置 着手", {"世話役": 1, "とび工": 4, "普通作業員": 2}, "", "新規入場者教育 7名"),
    "10/13": ("ゴンドラ 設置 完了、親綱・安全ネット 設置、内筒内 清掃", {"世話役": 1, "とび工": 4}, "", "ゴンドラ 使用前点検 OK"),
    "10/14": ("午前 内筒 板厚測定（GL+0〜8m 48点）、腐食範囲けがき。午後 腐食部 切断 着手", {"世話役": 1, "検査員": 2, "溶接工": 4},
              "板厚 最小 2.1mm（公称 6.0mm）", "腐食範囲 市 日下部様 確認"),
    "10/15": ("腐食部 切断・撤去 完了（GL+0.5〜6.0m 周方向 約 1/3）", {"世話役": 1, "溶接工": 4}, "", "撤去材 サンプル保管"),
    "10/16": ("当て板 SUS316L t6 取付・溶接（下段 1〜4枚目）", {"世話役": 1, "溶接工": 4}, "溶接条件 記録", ""),
    "10/18": ("当て板 取付・溶接（5〜9枚目）", {"世話役": 1, "溶接工": 4}, "", ""),
    "10/19": ("当て板 取付・溶接（10〜12枚目、計 3.6m²）。上部 GL+6.5m 付近に 追加腐食 確認", {"世話役": 1, "溶接工": 4}, "",
              "追加腐食部 市・SPC へ報告、当て板 0.8m² 追加 了承"),
    "10/20": ("午前 追加腐食部 当て板（0.8m²）。午後 溶接部 PT", {"世話役": 1, "溶接工": 2, "検査員": 2}, "PT 指示模様なし", ""),
    "10/21": ("外面保温 復旧（ラッキング 含む）", {"世話役": 1, "保温工": 3}, "", ""),
    "10/22": ("ゴンドラ 撤去・搬出", {"世話役": 1, "とび工": 4}, "", ""),
    "10/23": ("片付け・清掃、点検口 閉鎖、SPC へ作業場所 返却", {"世話役": 1, "とび工": 2, "普通作業員": 1}, "", "現地作業 完了"),
}
R3_THK = [(f"GL+{h:.1f}m", d, t) for h, d, t in
          [(0.5, "北", 2.4), (1.0, "北", 2.1), (1.5, "北東", 2.6), (2.0, "東", 3.0), (3.0, "北", 2.8), (4.0, "北西", 3.3),
           (5.0, "北", 3.5), (6.0, "西", 4.1), (6.5, "北", 3.4), (7.0, "北", 4.6), (8.0, "北", 5.2)]]

CITY = {
    "CITY-R3": dict(fy=2021, era="令和3年度", title="令和3年度 ひばり野クリーンセンター 2号煙突内筒補修工事",
                    period=(dt.date(2021, 9, 21), dt.date(2021, 12, 10)), site=("10/12", "10/23"), win=("10/11", "10/23"),
                    ref="2号炉定期整備工事（SPC 発注・別契約）", tasks=R3_TASKS, days=R3_DAY, eq=("煙突内筒", 2),
                    place="2号煙突内筒 下部（GL+0〜8m）（2号炉 定期整備 炉停止中）",
                    note1="■ 施工（灰）　■ 追加（赤：上部の追加腐食部）　日曜は休工", note2="※ 内筒内作業は 2号炉 停止・縁切り中のみ（1号炉は運転継続、外筒内 別系統）",
                    amount=12980000, cdate=dt.date(2021, 9, 21), done=dt.date(2021, 11, 30), insp=dt.date(2021, 12, 3)),
    "CITY-R5": dict(fy=2023, era="令和5年度", title="令和5年度 ひばり野クリーンセンター 2号ボイラ水管補修工事",
                    period=(dt.date(2023, 9, 25), dt.date(2023, 12, 15)), site=("10/10", "10/27"), win=("10/9", "10/28"),
                    ref="2号炉定期整備工事（SPC 発注・別契約）", tasks=K.TCITY, days=R5_DAY, eq=("ボイラ本体", 2),
                    place="2号ボイラ 第1放射室 前壁（2号炉 定期整備 炉停止中）",
                    note1="■ 施工（灰）　■ 追加（赤：RT 不合格継手の補修・再撮影）　日曜は休工。RT は 10/20 17:00 以降 立入制限",
                    note2="※ 水圧試験は SPC 2号炉定修のボイラ水圧試験と兼ねず単独で実施（市・SPC 立会）",
                    amount=28160000, cdate=dt.date(2023, 9, 25), done=dt.date(2023, 11, 30), insp=dt.date(2023, 12, 6)),
}


def city_work(cid):
    c = CITY[cid]
    e = W.eq(*c["eq"])
    return f"W{c['fy']}-{e.no}-CITY"


def ref_span(cid, contracts):
    c = CITY[cid]
    oh = [x for x in contracts if x["y"] == c["fy"] and x["f"] == 2][0]
    st = min(x for t in oh["tasks"] if t[2].startswith("炉立下げ") for x in K.task_days(t, c["fy"]))
    up = max(x for t in oh["tasks"] if t[2].startswith("炉立上げ") for x in K.task_days(t, c["fy"]))
    return st, up


def fname(cid, kind, ext="xlsx"):
    c = CITY[cid]
    nm = c["title"].split(" ", 2)[-1].replace("ひばり野クリーンセンター ", "")
    return f"【架空】ひばり野クリーンセンター_{kind}_{c['era']}{nm}_メーカー様式.{ext}"


def sched(cid, contracts):
    c = CITY[cid]
    y = c["fy"]
    fn = fname(cid, "工程表")
    R = Rec(fn)
    wb = Workbook()
    ws = wb.active
    ws.title = "実施工程表"
    props(wb, f"【架空】{c['title']} 実施工程表", "鷹羽エンジニアリング 現場代理人 m1", dt.datetime(y, c["period"][0].month, c["period"][0].day + 3, 11, 0))
    work = K.dates(y, *c["win"])
    C0 = 5
    for col, w in {1: 5, 2: 8, 3: 38, 4: 6}.items():
        ws.column_dimensions[L(col)].width = w
    for i in range(2 * len(work)):
        ws.column_dimensions[L(C0 + i)].width = 2.4
    p0, p1 = c["period"]
    put(ws, 1, 1, f"【架空】{c['title']}　実施工程表", font(14, True))
    put(ws, 2, 1, "【架空】施設：ひばり野クリーンセンター（ひばり野市）　発注者：ひばり野市　受注者：鷹羽エンジニアリング株式会社", font(9))
    s0, s1 = [dt.date(y, *K.d(x)) for x in c["site"]]
    put(ws, 3, 1, f"工期：令和{y - 2018}年{p0.month}月{p0.day}日〜令和{y - 2018}年{p1.month}月{p1.day}日　　現地施工：令和{y - 2018}年{s0.month}月{s0.day}日〜{s1.month}月{s1.day}日（2号炉定期整備（SPC）の炉停止期間内）", font(9))
    cE = C0 + 2 * len(work) - 12
    for j, h in enumerate(("監督員", "現場代理人")):
        col = cE + 6 * j
        ws.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col + 5)
        put(ws, 1, col, h, font(8), CENTER)
        ws.merge_cells(start_row=2, start_column=col, end_row=3, end_column=col + 5)
        put(ws, 2, col, "宇佐見" if h == "現場代理人" else None, font(9), CENTER)
        box(ws, 1, col, 3, col + 5)
    Rr = 5
    for col, h in ((1, "番号"), (2, "工種"), (3, "細　別"), (4, "人員")):
        ws.merge_cells(start_row=Rr, start_column=col, end_row=Rr + 2, end_column=col)
        put(ws, Rr, col, h, font(9, True), CENTER, B_ALL, fill("EDEDED"))
    for i, x in enumerate(work):
        col = C0 + 2 * i
        ws.merge_cells(start_row=Rr, start_column=col, end_row=Rr, end_column=col + 1)
        put(ws, Rr, col, f"{x.month}/{x.day}", font(7), CENTER, B_ALL, fill("EDEDED"))
        ws.merge_cells(start_row=Rr + 1, start_column=col, end_row=Rr + 1, end_column=col + 1)
        wd = K.WD[x.weekday()]
        hol = x in (dt.date(2023, 10, 9),)
        put(ws, Rr + 1, col, wd, font(7, color="C00000" if wd == "日" or hol else None), CENTER, B_ALL, fill("EDEDED"))
        put(ws, Rr + 2, col, "前", font(6), CENTER, B_ALL, fill("EDEDED"))
        put(ws, Rr + 2, col + 1, "後", font(6), CENTER, B_ALL, fill("EDEDED"))
    r = Rr + 3
    st, up = ref_span(cid, contracts)
    put(ws, r, 1, "参考", font(8), CENTER, B_ALL)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
    put(ws, r, 2, f"{c['ref']} 炉停止 {st.month}/{st.day}〜{up.month}/{up.day}", font(8, color="595959"), LEFT, B_ALL)
    for i in range(2 * len(work)):
        ws.cell(row=r, column=C0 + i).border = Border(bottom=Side(style="dashed", color="7F7F7F"), top=hair)
    R.add("実施工程表", r, "item", {"番号": "参考", "細別": f"{c['ref']} 炉停止 {st.month}/{st.day}〜{up.month}/{up.day}"})
    r += 1
    slot = {}
    for i, x in enumerate(work):
        slot[(x, "AM")] = C0 + 2 * i
        slot[(x, "PM")] = C0 + 2 * i + 1
    prevk = None
    for n, (tid, kind, item, crew_, men, (s, e), flag) in enumerate(c["tasks"], 1):
        put(ws, r, 1, n, font(9), CENTER, B_ALL)
        kv = kind if kind != prevk else None
        put(ws, r, 2, kv, font(9), CENTER, B_ALL)
        prevk = kind
        put(ws, r, 3, item + ("　※追加" if flag else ""), font(9, color="C00000" if flag else None), LEFT, B_ALL)
        put(ws, r, 4, men, font(9), CENTER, B_ALL)
        sx = dt.date(y, *K.d(s[0]))
        ex = dt.date(y, *K.d(e[0]))
        on = False
        for x in work:
            for h in ("AM", "PM"):
                cell = ws.cell(row=r, column=slot[(x, h)])
                cell.border = Border(left=hair if h == "AM" else None, right=hair if h == "PM" else None, top=hair, bottom=hair)
                if (x, h) == (sx, s[1]):
                    on = True
                if on and x.weekday() != 6:
                    cell.fill = fill("7F7F7F" if not flag else "C00000")
                elif x.weekday() == 6:
                    cell.fill = fill("F2F2F2")
                if (x, h) == (ex, e[1]):
                    on = False
        ws.row_dimensions[r].height = 17
        R.add("実施工程表", r, "item", {"番号": n, "工種": kv, "細別": item + ("　※追加" if flag else ""), "人員": men,
                                     "開始": f"{sx.isoformat()} {'午前' if s[1] == 'AM' else '午後'}",
                                     "終了": f"{ex.isoformat()} {'午前' if e[1] == 'AM' else '午後'}"}, [city_work(cid)])
        r += 1
    r += 1
    put(ws, r, 3, c["note1"], font(8))
    put(ws, r + 1, 3, c["note2"], font(8))
    page(ws, landscape=True, paper=9, fit_w=1, fit_h=1, area=f"A1:{L(C0 + 2 * len(work) - 1)}{r + 1}")
    wb.save(path(fn))
    return fn


def wareki(x):
    return f"令和{x.year - 2018}年{x.month}月{x.day}日（{K.WD[x.weekday()]}）"


def daily(cid):
    c = CITY[cid]
    y = c["fy"]
    fn = fname(cid, "作業日報")
    R = Rec(fn)
    r_ = random.Random(W.SEED + y)
    wb = Workbook()
    wb.remove(wb.active)
    props(wb, f"【架空】{c['title']} 工事日報", "鷹羽エンジニアリング 現場代理人 m1", dt.datetime(y, 10, 28, 9, 0))
    trades = TRADES if cid == "CITY-R5" else ["世話役", "溶接工", "とび工", "保温工", "検査員", "普通作業員"]
    mach = dict(MACH, とび工="ゴンドラ（仮設）", 保温工="保温材 切断機")
    for k, (txt, men, q, note) in c["days"].items():
        x = dt.date(y, *K.d(k))
        sh = f"{x.month}{x.day:02d}"
        ws = wb.create_sheet(sh)
        for col in range(1, 8):
            ws.column_dimensions[L(col)].width = 12
        ws.merge_cells("A1:G1")
        put(ws, 1, 1, "【架空】工 事 日 報", font(15, True), CENTER)
        ws.merge_cells("F2:G2")
        put(ws, 2, 6, "鷹羽エンジニアリング株式会社", font(8), RIGHT)
        put(ws, 3, 1, "工事名", font(9, True), CENTER, B_ALL, fill("EDEDED"))
        ws.merge_cells("B3:G3")
        put(ws, 3, 2, c["title"], font(9), LEFT, B_ALL)
        box(ws, 3, 1, 3, 7)
        put(ws, 4, 1, "年月日", font(9, True), CENTER, B_ALL, fill("EDEDED"))
        ws.merge_cells("B4:D4")
        put(ws, 4, 2, wareki(x), font(9), CENTER, B_ALL)
        wea = r_.choice(WEATHER)
        tmp = f"{r_.randint(14, 22)}℃"
        put(ws, 4, 5, "天候", font(9, True), CENTER, B_ALL, fill("EDEDED"))
        put(ws, 4, 6, wea, font(9), CENTER, B_ALL)
        put(ws, 4, 7, tmp, font(9), CENTER, B_ALL)
        box(ws, 4, 1, 4, 7)
        put(ws, 5, 1, "作業場所", font(9, True), CENTER, B_ALL, fill("EDEDED"))
        ws.merge_cells("B5:G5")
        extra = "（RT 17:00〜20:30）" if (cid, k) == ("CITY-R5", "10/20") else ""
        put(ws, 5, 2, f"{c['place']}　作業時間 8:00〜17:00{extra}", font(9), LEFT, B_ALL)
        box(ws, 5, 1, 5, 7)
        ws.merge_cells("A7:G7")
        put(ws, 7, 1, "作業内容", font(9, True), CENTER, B_ALL, fill("EDEDED"))
        ws.merge_cells("A8:G12")
        put(ws, 8, 1, txt, font(10), LEFT_TOP)
        box(ws, 8, 1, 12, 7)
        put(ws, 14, 1, "職種別人員", font(9, True), LEFT)
        for i, tname in enumerate(trades + ["計"], 1):
            put(ws, 15, i, tname, font(9, True), CENTER, B_ALL, fill("EDEDED"))
            v = sum(men.values()) if tname == "計" else men.get(tname)
            put(ws, 16, i, v, font(10), CENTER, B_ALL)
        mtxt = "、".join(mach[t] for t in trades if t in men and t in mach) or "－"
        put(ws, 18, 1, "使用機械", font(9, True), CENTER, B_ALL, fill("EDEDED"))
        ws.merge_cells("B18:G18")
        put(ws, 18, 2, mtxt, font(9), LEFT, B_ALL)
        put(ws, 19, 1, "品質記録", font(9, True), CENTER, B_ALL, fill("EDEDED"))
        ws.merge_cells("B19:G19")
        put(ws, 19, 2, q or "－", font(9), LEFT, B_ALL)
        put(ws, 20, 1, "安全・連絡", font(9, True), CENTER, B_ALL, fill("EDEDED"))
        ws.merge_cells("B20:G21")
        sv = (note + "　" if note else "") + f"KY：{KY[x.toordinal() % len(KY)]}"
        put(ws, 20, 2, sv, font(9), LEFT_TOP, B_ALL)
        ws.merge_cells("A20:A21")
        box(ws, 18, 1, 21, 7)
        for j, h in enumerate(("監督員", "現場代理人")):
            put(ws, 23, 6 + j, h, font(8), CENTER, B_ALL)
            put(ws, 24, 6 + j, None if h == "監督員" else "宇佐見", font(9, color="C00000"), CENTER, B_ALL)
        ws.row_dimensions[24].height = 30
        tot = sum(men.values())
        R.add(sh, 8, "item", {"年月日": x.isoformat(), "天候": wea, "気温": tmp, "作業場所": c["place"], "作業内容": txt,
                              "職種別人員": men, "人員計": tot, "使用機械": mtxt, "品質記録": q or "－", "安全・連絡": sv,
                              "manDays": tot, "manDaysText": f"{tot}人"}, [city_work(cid)])
        page(ws, landscape=False, paper=9, fit_w=1, fit_h=1, area="A1:G24")
    wb.save(path(fn))
    return fn


# ---------------------------------------------------------------- 完成報告書（reportlab）
def report(cid):
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
    try:
        pdfmetrics.getFont("IPAG")
    except KeyError:
        pdfmetrics.registerFont(TTFont("IPAG", "/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf"))
    c = CITY[cid]
    y = c["fy"]
    fn = fname(cid, "工事報告書", "pdf")
    R = Rec(fn)
    wk = [city_work(cid)]
    st = ParagraphStyle("n", fontName="IPAG", fontSize=9.5, leading=15)
    h1 = ParagraphStyle("h1", fontName="IPAG", fontSize=13, leading=20, spaceBefore=6, spaceAfter=4)
    h2 = ParagraphStyle("h2", fontName="IPAG", fontSize=11, leading=17, spaceBefore=4, spaceAfter=2)
    ct = ParagraphStyle("c", fontName="IPAG", fontSize=18, leading=28, alignment=1)
    cs = ParagraphStyle("cs", fontName="IPAG", fontSize=12, leading=20, alignment=1)
    small = ParagraphStyle("t", fontName="IPAG", fontSize=8, leading=11)

    def tbl(data, widths, fs=8.5, sec=None):
        t = Table(data, colWidths=[w * mm for w in widths], repeatRows=1)
        t.setStyle(TableStyle([("FONT", (0, 0), (-1, -1), "IPAG", fs), ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
                               ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E7E6E6")), ("VALIGN", (0, 0), (-1, -1), "MIDDLE")]))
        for row in data[1:]:
            fl = {h: (v.text if hasattr(v, "text") else v) for h, v in zip(data[0], row) if v not in ("", None)}
            fl["表"] = sec
            R.add_seq("item", fl, wk)
        return t

    def on_page(cv, d):
        cv.setFont("IPAG", 7.5)
        if d.page > 1:
            cv.drawString(20 * mm, 287 * mm, f"{c['title']} 完成報告書")
            cv.drawRightString(190 * mm, 287 * mm, "鷹羽エンジニアリング株式会社")
            cv.drawCentredString(105 * mm, 10 * mm, f"{d.page - 1}")
        cv.drawString(20 * mm, 6 * mm, "TKW-QF-0712 Rev." + ("3" if y < 2023 else "4"))

    doc = SimpleDocTemplate(path(fn), pagesize=A4, leftMargin=20 * mm, rightMargin=20 * mm, topMargin=20 * mm, bottomMargin=18 * mm,
                            title=f"【架空】{c['title']} 完成報告書", author="鷹羽エンジニアリング株式会社 m1")
    p0, p1 = c["period"]
    done = c["done"]
    s = [Spacer(1, 50 * mm), Paragraph("【架空】", cs), Paragraph(c["title"], ct), Spacer(1, 8 * mm), Paragraph("完 成 報 告 書", ct),
         Spacer(1, 60 * mm), Paragraph(f"令和{done.year - 2018}年{done.month}月{done.day}日", cs), Spacer(1, 6 * mm), Paragraph("ひばり野市長　様", cs),
         Spacer(1, 4 * mm), Paragraph("鷹羽エンジニアリング株式会社　環境プラント事業部", cs), Paragraph("現場代理人　宇佐見", cs), PageBreak()]
    s.append(Paragraph("目　次", h1))
    for x in ("1. 工事概要", "2. 工事経過", "3. 施工内容", "4. 検査記録", "5. 使用材料", "6. 写真"):
        s.append(Paragraph(x, st))
    s.append(Spacer(1, 6 * mm))
    s.append(Paragraph("1. 工事概要", h1))
    e = W.eq(*c["eq"])
    s0, s1 = [dt.date(y, *K.d(x)) for x in c["site"]]
    if cid == "CITY-R5":
        tgt = "2号ボイラ本体（資産番号 HB03-201）第1放射室 前壁 水管 " + K.TUBE_SPEC
        cont = "減肉水管 6本 部分取替（各1.8m）、減肉部 肉盛溶接 12箇所、非破壊検査、水圧試験、耐火物復旧"
    else:
        tgt = f"2号煙突内筒（資産番号 {e.no}）SUS 内径1,000 下部 GL+0〜8m"
        cont = "腐食部 切断撤去、当て板 SUS316L t6 計 4.4m²（うち追加 0.8m²）、浸透探傷試験、外面保温復旧"
    ov = [["項目", "内容"], ["工事名", c["title"]], ["工事場所", "ひばり野市（【架空】ひばり野クリーンセンター）"],
          ["発注者", "ひばり野市（環境課）　監督員 日下部"], ["受注者", "鷹羽エンジニアリング株式会社　現場代理人 宇佐見"],
          ["工期", f"令和{y - 2018}年{p0.month}月{p0.day}日 〜 令和{y - 2018}年{p1.month}月{p1.day}日"],
          ["現地施工", f"令和{y - 2018}年{s0.month}月{s0.day}日 〜 令和{y - 2018}年{s1.month}月{s1.day}日（2号炉定期整備の炉停止期間内）"],
          ["対象設備", Paragraph(tgt, small)], ["工事内容", Paragraph(cont, small)]]
    s.append(tbl(ov, [32, 138], sec="工事概要"))
    s.append(Paragraph("2. 工事経過", h1))
    rows = [["月日", "作業内容", "人員"]]
    for k, (txt, men, q, note) in c["days"].items():
        x = dt.date(y, *K.d(k))
        rows.append([f"{x.month}/{x.day}（{K.WD[x.weekday()]}）", Paragraph(txt, small), str(sum(men.values()))])
    s.append(tbl(rows, [24, 132, 14], 8, sec="工事経過"))
    s.append(Paragraph("3. 施工内容", h1))
    if cid == "CITY-R5":
        paras = ["令和4年度の2号炉定期整備（ひばり野環境サービス㈱）の肉厚測定において、第1放射室前壁の水管に減肉（最小3.2mm、設計必要厚2.9mm）が"
                 "確認されたことを受け、本工事を実施した。着工時に前壁63点の肉厚を再測定し、最小肉厚3.5mm以下の6本（F-14〜F-16、F-21〜F-23）を"
                 "部分取替、3.8〜4.0mmの12箇所を肉盛溶接の対象とした。取替は既設管を各1.8m切断撤去し、新管を突合せ溶接（TIG初層＋被覆アーク）で接続した（計12継手）。",
                 "放射線透過試験の結果、F-22下側継手にブローホール（3類）が認められたため、当該部をはつり取り補修溶接の上、再撮影し合格を確認した。"
                 "この補修のため10月21日（土）に作業を追加したが、全体工程への影響はなかった。肉盛部は浸透探傷試験で指示模様のないことを確認し、"
                 "肉盛後の肉厚は4.6〜5.2mmであった。水圧試験（4.5MPa、30分保持）は市監督員およびSPCの立会のもと実施し、漏れのないことを確認した。"
                 "最後に、はつり取った範囲の耐火物をキャスタブル（4.2m²）で復旧した。"]
    else:
        paras = ["令和2年度の2号炉定期整備（ひばり野環境サービス㈱）において、2号煙突内筒下部に腐食が確認され、SPC の年度補修上限を超える規模であるため"
                 "市発注の追加工事として本工事を実施した。着工時に GL+0〜8m の範囲で板厚を測定し（48点）、最小 2.1mm（公称 6.0mm）の範囲を"
                 "周方向約 1/3 にわたり切断撤去し、SUS316L t6 の当て板 12 枚（計 3.6m²）を全周溶接で取り付けた。",
                 "施工中、GL+6.5m 付近に新たな腐食（残厚 3.4mm）を確認したため、市監督員およびSPCと協議のうえ当て板 0.8m² を追加した（10月20日）。"
                 "溶接部は浸透探傷試験で指示模様のないことを確認し、外面保温を復旧した。内筒内の作業は 2号炉の停止・縁切り中に限り、ゴンドラを用いて行った。"]
    for p_ in paras:
        s.append(Paragraph(p_, st))
        R.add_seq("item", {"節": "施工内容", "本文": p_}, wk)
    s.append(Paragraph("4. 検査記録", h1))
    if cid == "CITY-R5":
        s.append(Paragraph("4.1 肉厚測定（取替対象）", h2))
        s.append(tbl([["管番号", "令和4年度 最小(mm)", "着工時 最小(mm)", "措置"]] + [[a, f"{b:.1f}", f"{c_:.1f}", d_] for a, b, c_, d_ in K.TUBES],
                     [25, 40, 40, 65], sec="肉厚測定"))
        s.append(Paragraph("4.2 肉盛溶接 箇所", h2))
        data = [["管番号", "着工時 最小(mm)", "PT", "肉盛後(mm)"]]
        for i, (a, b) in enumerate(K.OVERLAY):
            data.append([a, f"{b:.1f}", "合格", f"{4.6 + (i * 7 % 6) / 10:.1f}"])
        s.append(tbl(data, [30, 45, 30, 45], sec="肉盛溶接"))
        s.append(Paragraph("4.3 放射線透過試験（RT）", h2))
        data = [["継手", "撮影日", "判定", "再撮影", "最終"]]
        for a, *_ in K.TUBES:
            for pos in ("上", "下"):
                bad = (a == "F-22" and pos == "下")
                data.append([f"{a} {pos}側", "10/20", "3類 不合格" if bad else "1類 合格", "10/21 1類 合格" if bad else "－", "合格"])
        s.append(tbl(data, [30, 25, 35, 45, 35], sec="RT"))
        s.append(Paragraph("4.4 水圧試験：令和5年10月24日 試験圧力 4.5MPa 保持30分 漏れなし（立会：市 日下部、SPC 小鳥遊）", st))
        mat = [["品名", "規格", "数量", "備考"], ["ボイラ水管", "STB340 φ60.5×t4.5 L=2.0m", "7本", "6本使用、1本は市へ予備品として引渡し"],
               ["溶接材料", "TIG棒 YGT50・被覆棒 E4316", "1式", ""], ["キャスタブル耐火物", "高アルミナ質", "1.9t", "4.2m²"],
               ["Yアンカー", "SUS310S", "120本", ""]]
        caps = ["写真1 着工前 前壁水管（はつり後）", "写真2 既設管 切断撤去", "写真3 新管 溶接完了", "写真4 肉盛溶接 完了",
                "写真5 RT 撮影（不合格部 補修後）", "写真6 耐火物復旧 完了"]
    else:
        s.append(Paragraph("4.1 板厚測定（着工時）", h2))
        s.append(tbl([["位置", "方位", "最小板厚(mm)", "措置"]] + [[a, b, f"{t:.1f}", "当て板（追加）" if a == "GL+6.5m" else ("当て板" if t < 4.0 else "－")]
                                                              for a, b, t in R3_THK], [35, 25, 35, 75], sec="板厚測定"))
        s.append(Paragraph("4.2 浸透探傷試験（PT）：令和3年10月20日 当て板溶接部 全線 指示模様なし（立会：市 日下部）", st))
        mat = [["品名", "規格", "数量", "備考"], ["当て板", "SUS316L t6", "4.4m²", "うち追加 0.8m²"], ["溶接材料", "被覆棒 E316L", "1式", ""],
               ["保温材", "ロックウール t50・ラッキング", "1式", "外面復旧"]]
        caps = ["写真1 内筒下部 腐食状況", "写真2 腐食部 切断撤去", "写真3 当て板 溶接完了", "写真4 追加腐食部（GL+6.5m）",
                "写真5 PT 実施状況", "写真6 保温復旧 完了"]
    s.append(Paragraph("5. 使用材料", h1))
    s.append(tbl(mat, [40, 60, 20, 50], sec="使用材料"))
    s.append(PageBreak())
    s.append(Paragraph("6. 写真", h1))
    ph = [["", ""], caps[0:2], ["", ""], caps[2:4], ["", ""], caps[4:6]]
    t = Table(ph, colWidths=[85 * mm, 85 * mm], rowHeights=[62 * mm, 8 * mm] * 3)
    t.setStyle(TableStyle([("FONT", (0, 0), (-1, -1), "IPAG", 8.5), ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                           ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#F2F2F2")), ("BACKGROUND", (0, 2), (-1, 2), colors.HexColor("#F2F2F2")),
                           ("BACKGROUND", (0, 4), (-1, 4), colors.HexColor("#F2F2F2")), ("ALIGN", (0, 0), (-1, -1), "CENTER")]))
    s.append(t)
    doc.build(s, onFirstPage=on_page, onLaterPages=on_page)
    return fn


# ---------------------------------------------------------------- 市の契約一覧（市 環境課 b1）
def ledger():
    fn = "【架空】ひばり野クリーンセンター_発注台帳_市追加工事契約一覧_令和3〜8年度_市様式.xlsx"
    R = Rec(fn)
    wb = Workbook()
    ws = wb.active
    ws.title = "追加工事"
    props(wb, "【架空】ひばり野クリーンセンター 追加工事 契約一覧（市 環境課）", "ひばり野市 環境課 b1", dt.datetime(2026, 4, 10, 9, 0))
    put(ws, 1, 1, "【架空】ひばり野クリーンセンター　長期包括運営委託に係る追加工事（市発注）契約一覧", font(13, True))
    put(ws, 2, 1, "環境課 施設係（監視担当 日下部）　※SPC の年度補修上限額を超えるもの。金額は税込", font(9))
    heads = ["年度", "契約番号", "件名", "契約方法", "契約日", "工期", "契約者", "契約金額（円）", "監督員", "完成検査", "備考"]
    widths = [8, 11, 40, 14, 11, 22, 18, 13, 8, 11, 34]
    for i, (h, w) in enumerate(zip(heads, widths), 1):
        put(ws, 4, i, h, font(9, True), CENTER, B_ALL, fill("DDEBF7"))
        ws.column_dimensions[L(i)].width = w
    rows = []
    for k, cid in enumerate(("CITY-R3", "CITY-R5")):
        c = CITY[cid]
        p0, p1 = c["period"]
        y = c["fy"]
        rows.append(([c["era"].replace("年度", ""), f"環施第{y - 2018}-{17 + 4 * k}号", c["title"].split(" ", 1)[1], "随意契約（特命）",
                      f"R{y - 2018}.{c['cdate'].month}.{c['cdate'].day}", f"R{y - 2018}.{p0.month}.{p0.day}〜R{y - 2018}.{p1.month}.{p1.day}",
                      "鷹羽エンジニアリング㈱", c["amount"], "日下部", f"R{y - 2018}.{c['insp'].month}.{c['insp'].day}",
                      "竣工メーカーのため特命。2号炉定修（SPC）の停止期間内" + ("。追加当て板 0.8m² は変更なし（数量内）" if cid == "CITY-R3" else "。RT 不合格1継手は受注者負担で補修")],
                     [city_work(cid)]))
    rows.append((["R7", "", "（1号過熱器 補修）", "", "", "", "", None, "", "", "SPC 年度上限内で実施のため市発注なし（R7.3 協議）"], []))
    rows.append((["R8", "", "（予定なし）", "", "", "", "", None, "", "", "廃止（R12末）まで追加工事は原則行わない方針（R8.4 確認）"], []))
    r = 5
    for row, wk in rows:
        for i, v in enumerate(row, 1):
            put(ws, r, i, v, font(9), LEFT if i in (3, 7, 11) else CENTER, B_ALL, None, "#,##0" if i == 8 else None)
        ws.row_dimensions[r].height = 30
        R.add("追加工事", r, "item", {h: v for h, v in zip(heads, row) if v not in ("", None)}, wk)
        r += 1
    put(ws, r, 7, "計", font(9, True), CENTER, B_ALL, fill("FFF2CC"))
    tot = sum(CITY[c]["amount"] for c in CITY)
    put(ws, r, 8, tot, font(9, True), CENTER, B_ALL, fill("FFF2CC"), "#,##0")
    R.add("追加工事", r, "total", {"契約金額（円）": tot})
    ws.cell(row=5, column=3).comment = None
    page(ws, landscape=True, paper=9, fit_w=1, fit_h=1, area=f"A1:K{r}")
    wb.save(path(fn))
    return fn


def build_all(contracts):
    out = []
    for cid in CITY:
        out += [sched(cid, contracts), daily(cid), report(cid)]
    out.append(ledger())
    return out
