# -*- coding: utf-8 -*-
"""工事報告書（本作り、reportlab の文字のある PDF。ページをまたぐ表は見出しをくり返す）
 R1：文章型の完了報告書（e1 の最初の年）／R2〜：1 枚目にサマリー表、作業内容・工程実績・測定・材料・写真帳"""
import os, random
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import BaseDocTemplate, PageTemplate, Frame, Paragraph, Spacer, Table, TableStyle, PageBreak, Flowable
from reportlab.lib.styles import ParagraphStyle
from common import OUT
import world as W
import fullworld as FW
import r5grate as G

pdfmetrics.registerFont(TTFont("IPAG", "/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf"))
pdfmetrics.registerFont(TTFont("IPAPG", "/usr/share/fonts/opentype/ipafont-gothic/ipagp.ttf"))

def S(size=10, lead=None, font="IPAPG", **kw):
    return ParagraphStyle("s%d%s" % (size, font), fontName=font, fontSize=size, leading=lead or size * 1.6, wordWrap="CJK", **kw)
H1 = S(13, font="IPAG"); H2 = S(11, font="IPAG"); BODY = S(10); SMALL = S(8.5)

def P(t, st=None):
    return Paragraph(str(t), st or SMALL)

def tstyle(size=8.5):
    return TableStyle([("FONTNAME", (0, 0), (-1, -1), "IPAPG"), ("FONTSIZE", (0, 0), (-1, -1), size),
                       ("GRID", (0, 0), (-1, -1), 0.5, colors.black), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                       ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E7E6E6"))])

def fname(job, style):
    return "【架空】さくら坂焼却場_工事報告書_%s%s_%s.pdf" % (W.nendo_name(job.fy), job.c["short"], style)

def mkdoc(fn, title, header):
    d = BaseDocTemplate(fn, pagesize=A4, leftMargin=20 * mm, rightMargin=20 * mm, topMargin=22 * mm, bottomMargin=18 * mm,
                        title=title, author="【架空】さくら坂機工株式会社 所長 e1", subject=title, creator="reportlab")
    fr = Frame(d.leftMargin, d.bottomMargin, d.width, d.height, id="f")
    def onpage(c, dd):
        c.saveState(); c.setFont("IPAPG", 8)
        if dd.page > 1:
            c.drawString(20 * mm, A4[1] - 12 * mm, header)
        c.drawCentredString(A4[0] / 2, 10 * mm, "- %d -" % dd.page)
        c.restoreState()
    d.addPageTemplates([PageTemplate(id="p", frames=[fr], onPage=onpage)])
    return d

class Cover(Flowable):
    def __init__(self, lines):
        super().__init__(); self.lines = lines
    def wrap(self, w, h):
        self.w = w; return w, h - 2
    def draw(self):
        c = self.canv; y = 640
        c.setFont("IPAPG", 10); c.drawString(0, 700, "【架空】")
        for size, text in self.lines:
            if size == 0:
                y -= 30; continue
            c.setFont("IPAG" if size >= 14 else "IPAPG", size); c.drawCentredString(self.w / 2, y, text); y -= size * 2.2

class Photo(Flowable):
    def __init__(self, w, h, kind, cap):
        super().__init__(); self.w, self.h, self.kind, self.cap = w, h, kind, cap
    def wrap(self, aw, ah):
        return self.w, self.h + 16
    def draw(self):
        c = self.canv; c.setLineWidth(0.6); c.rect(0, 16, self.w, self.h)
        c.setStrokeColor(colors.HexColor("#555555"))
        h = hash(self.cap) % 5
        if self.kind in ("grate", "bf"):
            for i in range(8):
                x = 12 + i * (self.w - 24) / 8
                c.rect(x, 30 + (i % 2) * 6, (self.w - 24) / 8 - 3, self.h * (0.45 if self.kind == "grate" else 0.7))
        elif self.kind in ("boiler",):
            for i in range(10):
                x = 15 + i * (self.w - 30) / 10
                c.line(x, 25, x, self.h)
        else:
            c.circle(self.w / 2, self.h / 2 + 10, self.h * 0.3)
            c.line(self.w / 2 - self.h * 0.3, self.h / 2 + 10, self.w / 2 + self.h * 0.3, self.h / 2 + 10 + h)
        c.setFont("IPAPG", 7); c.drawString(4, self.h + 6, "【架空】見本用の写真代替図")
        c.setFont("IPAPG", 8.5); c.drawString(0, 3, self.cap)

# 定修で見つけたこと（年度・機器 → 文）
FIND = {
    (2019, "1B-03"): "摩耗は最大 8.6mm（取替の目安 12mm）。年 1mm 前後の進み",
    (2019, "1D-05"): "差圧 1.6〜1.9kPa、抜取ろ布の強度 新品の約 55%。R2 取替を提案",
    (2019, "0A-07"): "予備機 巻上ワイヤ 素線切れ 2 か所",
    (2020, "0G-08"): "常用機パドル 摩耗。肉盛補修",
    (2021, "0E-01"): "常用機 管漏れ 3 本、施栓し水圧試験 良（計画外）",
    (2021, "1G-01"): "ケーシング腐食、分解整備をやめ当て板補修に変更",
    (2021, "2F-07"): "羽根車に付着物、除去・現地バランス修正（3.6→2.1mm/s）",
    (2021, "1C-01"): "第 1 パス前面 最小 2.4mm。R4 補修工事を提案",
    (2022, "2F-07"): "振動 4.5mm/s、羽根先端 残厚約 45%。羽根車の手配を提案",
    (2022, "1C-12"): "下部コーン腐食部 当て板",
    (2022, "0A-07"): "常用機 走行車輪 摩耗、取替は R5 へ繰越（予算）",
    (2023, "0A-07"): "常用機 走行車輪 取替（繰越分）",
    (2023, "2F-07"): "振動 4.8mm/s",
    (2024, "2F-07"): "振動 5.0mm/s（羽根車 納入待ち）",
    (2024, "2C-12"): "下部コーン当て板",
    (2024, "1F-07"): "軸受交換、振動 1.8→1.2mm/s",
    (2025, "2F-07"): "羽根車取替（R6 購入品）・軸受交換、試運転 1.6mm/s",
}

def result_of(job, t):
    y = job.fy
    out = []
    for e, m in t.eqs:
        if (y, e.no) in FIND:
            out.append(FIND[(y, e.no)])
        else:
            ml = FW.meas_line(job, e, y, None)
            if ml:
                out.append(ml.split(" ", 1)[1] if " " in ml else ml)
    if t.kind == "binsp":
        return "合格"
    if t.kind == "crane":
        return "良" + ("、性能検査 合格" if y in FW.CRANE_PERF else "") + ("（%s）" % t.memo if t.memo else "")
    if t.kind == "c_elec":
        return "異常なし（受変電 年次点検は保安法人の報告による）"
    if not out:
        hv = [e for e, m in t.eqs if m in ("◎", "△", "●")]
        out.append("実施、良" if hv else "異常なし")
    if t.memo:
        out.append(t.memo)
    return "。".join(out)

def parts_of(job):
    y = job.fy
    rows = [("減温塔噴霧ノズルチップ", "2 流体", "6 個", "1号・2号 減温塔", "市支給（在庫品）")]
    for sname, ab, tasks in job.sheets:
        for t in tasks:
            for e, m in t.eqs:
                if m == "◎":
                    rows.append(("軸受・パッキン一式", "%s 用" % e.name, "1 式", "%s %s" % (sname if e.unit != "共通" else "共通", e.name), "当社手配"))
                if m == "△":
                    rows.append(("鋼板（当て板）", "SS400 t6", "1 式", "%s %s" % (sname if e.unit != "共通" else "共通", e.name), "当社手配"))
    rows += [("Vベルト", "B-62", "8 本", "1・2号 押込送風機", "当社手配"),
             ("潤滑油", "タービン油 VG68", "40 L", "誘引通風機 1・2号", "当社手配"),
             ("グリス", "リチウム 2 号", "%d kg" % (14 + y % 5), "各機器", "当社手配"),
             ("耐火キャスタブル", "SK34 相当", "%d kg" % (150 + (y % 4) * 50), "1・2号 側壁上部", "当社手配")]
    if y == 2025:
        rows.append(("誘引通風機 羽根車", "AN-F07-IMP2", "1 個", "2号 誘引通風機", "市支給（R6 購入品）"))
        rows.append(("誘引通風機 軸受", "球面ころ軸受 22220", "2 個", "2号 誘引通風機", "市支給（R6 購入品）"))
    if y == 2020:
        rows.append(("クレーン巻上ワイヤ", "6×Fi(29) φ20 120m", "2 本", "ごみクレーン 予備機", "市支給（R2 購入品）"))
    if y == 2021:
        rows.append(("管栓", "SUS304 φ19", "3 個", "温水熱交換器（常用機）", "市手配（17 個は予備）"))
    if y == 2023:
        rows.append(("走行車輪", "常用機用", "2 個", "ごみクレーン 常用機", "当社手配"))
    return rows

REPAIR_PARTS = {
    "R2020a": [("ろ布", "PTFE φ150×6m", "540 本", "1号 BF", "当社手配（工事に含む）"), ("ガスケット", "点検口用", "12 枚", "1号 BF", "当社手配")],
    "R2021a": [("ろ布", "PTFE φ150×6m", "540 本", "2号 BF", "当社手配（工事に含む）"), ("ガスケット", "点検口用", "12 枚", "2号 BF", "当社手配")],
    "R2022a": [("ボイラ水管", "STB340 φ45.0×3.2", "8 本分", "1号 廃熱ボイラ 第1パス前面", "当社手配（工事に含む、2 本は変更で追加）"), ("溶接棒", "低水素系 φ2.6・φ3.2", "10 kg", "同上", "当社手配"), ("保温材", "ロックウール", "1 式", "同上", "当社手配")],
    "R2024a": [("混練機パドル", "SS400＋肉盛 地元製作", "24 枚", "飛灰混練機（常用機）", "当社製作（工事に含む）"), ("ボルト", "M16 SUS", "48 組", "同上", "当社手配")],
    "R2025a": [("火格子片（予備品）", "AN-G60-203", "18 枚", "2号 燃焼火格子", "市支給（R4 購入品の残り）"), ("火格子片（再使用）", "1号取外し品の良品", "12 枚", "2号 燃焼火格子", "市支給（場内保管品）"), ("六角ボルト・ナット", "M20×70", "64 組", "同上", "当社手配")],
}
PHOTO_KIND = {"R2020a": "bf", "R2021a": "bf", "R2022a": "boiler", "R2024a": "x", "R2025a": "grate"}


def build_old(job):
    """R1 完了報告書（文章型）"""
    recs = []
    seq = [0]
    def rec(fields, kind="item"):
        recs.append(dict(seq=seq[0], kind=kind, fields=fields)); seq[0] += 1
    s = FW.SHUT[2019]
    fn = os.path.join(OUT, fname(job, "文章型"))
    title = "【架空】令和元年度 さくら坂焼却場 定期整備業務委託 完了報告書"
    d = mkdoc(fn, title, "令和元年度 さくら坂焼却場 定期整備業務委託 完了報告書")
    st = [Cover([(18, "令和元年度"), (18, "さくら坂焼却場 定期整備業務委託"), (20, "完 了 報 告 書"), (0, ""), (0, ""),
                 (11, "【架空】さくら坂市 環境部 施設管理課 様"), (0, ""), (0, ""), (11, "令和2年3月27日"), (12, "【架空】さくら坂機工株式会社"), (11, "所長　e1")]), PageBreak()]
    st.append(Paragraph("1. 業務の概要", H1))
    paras = ["本業務は、令和元年度にさくら坂焼却場（60t/24h×2 炉、階段式ストーカ、廃熱ボイラ 1.5MPa 飽和）の定期整備を行ったものである。本年度から当社が受託した。竣工メーカー（【架空】朝凪炉工）が整備から撤退したため、前年度までの点検要領書と完成図書（市保管）を借用して作業項目を定めた。",
             "業務期間：令和元年4月1日〜令和2年3月31日。1号炉 6月%d日停止〜6月%d日立上げ、2号炉 10月%d日停止〜11月%d日立上げ。全停電 10月%d日〜%d日。" % (s["u1"][0].day, s["u1"][1].day, s["u2"][0].day, s["u2"][1].day, s["zen"][0].day, s["zen"][1].day),
             "作業班：職長 f1 ほか 5〜7 名、延べ %.1f 人日。電気・計装の一部は【架空】みずなら電設（再委託）。" % sum(x["md"] for x in job.days)]
    for t in paras:
        st.append(Paragraph(t, BODY)); st.append(Spacer(1, 3))
    st.append(Paragraph("2. 実施した作業", H1))
    for sname, (a, b), tasks in job.sheets:
        st.append(Paragraph("2.%s %s" % (sname[0], sname), H2))
        txt = []
        for t in tasks:
            if t.team == "運転":
                continue
            ds = sorted(t.act)
            txt.append("%s（%d月%d日%s）" % (t.name, ds[0].month, ds[0].day, "〜%d月%d日" % (ds[-1].month, ds[-1].day) if len(ds) > 1 else ""))
        st.append(Paragraph("、".join(txt) + "を行った。", BODY)); st.append(Spacer(1, 4))
    st.append(Paragraph("3. 点検の結果と所見", H1))
    obs = [("1B-03", "1号炉 燃焼火格子", "摩耗量は最大 8.6mm（取替の目安 12mm）。平成22年の全数更新から 9 年で、年 1mm 前後の進みである。令和5〜6年度に目安に達する片が出る見込み。"),
           ("1C-01", "1号炉 廃熱ボイラ", "第1パス前面の水管で最小 2.8mm（原厚 3.2mm、計算上の必要厚さ 2.0mm）。毎年の測定を続けること。"),
           ("1D-05", "1号炉 バグフィルタ", "差圧 1.6〜1.9kPa。抜取ろ布の引張強さは新品の約 55%。平成26年の取替から 5 年が経ち、令和2年度の取替を提案する。"),
           ("2F-07", "2号炉 誘引通風機", "振動 2.3mm/s、羽根先端にダストの付着と軽い摩耗あり。"),
           ("0A-07", "ごみクレーン", "常用機は良好。予備機（竣工時）は巻上ワイヤの素線切れ 2 か所、ブレーキライニングの摩耗が進んでおり、令和2年度の整備を提案する。性能検査は合格。")]
    for i, (no, nm, t) in enumerate(obs, 1):
        st.append(Paragraph("（%d）%s：%s" % (i, nm, t), BODY)); st.append(Spacer(1, 2))
        rec({"区分": "所見", "対象機器": nm, "_equip": no, "内容": t})
    st.append(Paragraph("4. 使用した部品・材料", H1))
    data = [["品名", "仕様", "数量", "使用箇所", "備考"]]
    for r in parts_of(job):
        data.append([P(x) for x in r])
        rec({"区分": "使用部品", "品名": r[0], "仕様": r[1], "数量": r[2], "使用箇所": r[3], "支給/手配": r[4]})
    t = Table(data, colWidths=[38 * mm, 32 * mm, 16 * mm, 42 * mm, 36 * mm], repeatRows=1); t.setStyle(tstyle()); st.append(t)
    st.append(Spacer(1, 6))
    st.append(Paragraph("5. 次年度以降への申し送り", H1))
    for t in ["・1号炉 BF ろ布の全数取替（令和2年度）、2号炉は令和3年度。", "・ごみクレーン予備機の巻上ワイヤ・ブレーキ整備（令和2年度 定期整備）。",
              "・中央監視装置の監視端末は OS のサポートが終わるため、早めの更新が必要（当社の範囲外）。", "・朝凪炉工の部品は納期が長いものがあり、火格子片・羽根車などは早めに見積を取ること。"]:
        st.append(Paragraph(t, BODY))
    st.append(Spacer(1, 10)); st.append(Paragraph("以上", S(10, alignment=2)))
    d.build(st)
    return fn, recs


def build_new(job, rnd):
    recs = []
    seq = [0]
    def rec(fields, kind="item", image=None):
        x = dict(seq=seq[0], kind=kind, fields=fields)
        if image: x["image"] = image
        recs.append(x); seq[0] += 1
    fn = os.path.join(OUT, fname(job, "サマリー付表型"))
    title = "【架空】%s 工事報告書" % job.title if job.cid.startswith("R") else "【架空】%s 完了報告書" % job.title
    d = mkdoc(fn, title, job.title + (" 工事報告書" if job.cid.startswith("R") else " 完了報告書"))
    last = max(x["d"] for x in job.days)
    import datetime as dt
    rep_date = job.c["end"] - dt.timedelta(4) if job.cid.startswith("T") else last + dt.timedelta(20)
    if job.cid == "R2023a":
        rep_date = dt.date(2023, 7, 14)
    while rep_date.weekday() >= 5:
        rep_date -= dt.timedelta(1)
    wd = W.wareki_k(rep_date.year)
    st = [Cover([(18, W.nendo_name(job.fy)), (18, "さくら坂焼却場 " + job.c["short"]), (20, "工 事 報 告 書" if job.cid.startswith("R") else "完 了 報 告 書"), (0, ""), (0, ""),
                 (11, "【架空】さくら坂市 環境部 施設管理課 様"), (0, ""), (0, ""), (11, "%s%d月%d日" % (wd, rep_date.month, rep_date.day)),
                 (12, "【架空】さくら坂機工株式会社"), (11, "所長　e1" + ("　　現場責任者　f2" if job.cid.startswith("R") else "　　職長　f1"))]), PageBreak()]
    st.append(Paragraph("報告サマリー", H1))
    a0 = min(x["d"] for x in job.days); a1 = last
    info = [["業務・工事件名", job.title], ["期間", "%s〜%s（現場作業 %s〜%s、実働 %d 日）" % (job.c["start"].strftime("%Y/%m/%d"), job.c["end"].strftime("%Y/%m/%d"), a0.strftime("%m/%d"), a1.strftime("%m/%d"), len({x["d"] for x in job.days}))],
            ["施工場所", "さくら坂焼却場 " + "・".join(s[0] for s in job.sheets)], ["受注者", "【架空】さくら坂機工株式会社"]]
    t = Table([[P(a, S(9, font="IPAG")), P(b, S(9))] for a, b in info], colWidths=[30 * mm, 134 * mm])
    t.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.5, colors.black), ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#E7E6E6")), ("VALIGN", (0, 0), (-1, -1), "MIDDLE")]))
    st.append(t); st.append(Spacer(1, 6))
    data = [["対象機器（台帳No）", "作業", "実施日", "結果・判定"]]
    if job.cid == "R2023a":
        rows = [("1号炉 燃焼火格子（1B-03）", "摩耗測定・取替片の選定", "6/9", "240 枚測定。摩耗 12mm 以上・割れ 72 枚を取替対象"),
                ("1号炉 燃焼火格子（1B-03）", "火格子片 取外し（72 枚）", "6/12〜6/15", "固定ボルト固着、ガス切断で除去（1 日延長）"),
                ("1号炉 燃焼火格子（1B-03）", "可動フレーム支持ローラ取替（2 個）", "6/16", "R3・R4 偏摩耗。市支給の予備品で取替（追加）"),
                ("1号炉 燃焼火格子（1B-03）", "火格子片 取付・回り止め（72 枚）", "6/16〜6/20", "新品ボルト M20 で締付、回り止め溶接"),
                ("1号炉 燃焼火格子（1B-03）", "クリアランス調整・測定", "6/21", "片間 2.0〜4.0mm、側壁 5〜8mm　良"),
                ("1号炉 燃焼火格子（1B-03）", "空作動試験・立会検査", "6/22", "異音・引っ掛かりなし　合格")]
        for r in rows:
            data.append([P(x) for x in r]); rec({"区分": "サマリー", "対象機器": r[0], "作業": r[1], "実施日": r[2], "結果": r[3]})
    else:
        for sname, ab, tasks in job.sheets:
            for tk in tasks:
                if tk.team == "運転":
                    continue
                act = sorted(d_ for d_ in tk.act if d_ <= FW.TODAY)
                if not act:
                    continue
                eqs = tk.eqs
                tg = "、".join("%s（%s）" % (("%s号炉 " % e.unit if e.unit in ("1", "2") else "") + e.name, e.no) for e, m in eqs[:3]) + (" ほか %d" % (len(eqs) - 3) if len(eqs) > 3 else "")
                ds = "%d/%d" % (act[0].month, act[0].day) + ("〜%d/%d" % (act[-1].month, act[-1].day) if len(act) > 1 else "")
                res = result_of(job, tk)
                data.append([P(tg or sname), P(tk.name), P(ds), P(res)])
                rec({"区分": "サマリー", "対象機器": [e.no for e, m in eqs], "作業": tk.name, "実施日": [x.isoformat() for x in act], "結果": res, "task": tk.no, "sheet": sname})
    t = Table(data, colWidths=[48 * mm, 40 * mm, 20 * mm, 62 * mm], repeatRows=1); t.setStyle(tstyle(8)); st.append(t)
    st.append(PageBreak())
    st.append(Paragraph("1. 作業の内容", H1))
    if job.cid == "R2023a":
        for tx in ["1号炉の定期整備（別契約）の停止期間を使い、燃焼火格子の摩耗した火格子片を取り替えた。炉内清掃（定期整備 A班）が終わった 6月8日午後から B班が炉内に入った。",
                   "6月9日に燃焼段 240 枚の摩耗を測定し、摩耗 12mm 以上の片と割れのある片、計 72 枚（うち割れ 5 枚）を取替対象とした。",
                   "取外しは固定ボルトの固着が多く、ガス切断で除去したため 1 日延びて 6月15日に終わった。支持ローラ R3・R4 の偏摩耗を認め、倉庫の予備ローラ 2 個で 6月16日に取り替えた。",
                   "取付は 6月16日・19日・20日、クリアランス調整 6月21日、空作動試験と市の立会検査 6月22日 合格。予備の火格子片 18 枚は倉庫 B-3 棚に保管した。"]:
            st.append(Paragraph(tx, BODY)); st.append(Spacer(1, 3))
    else:
        for sname, (a, b), tasks in job.sheets:
            ts = [x for x in tasks if x.team != "運転" and x.act and min(x.act) <= FW.TODAY]
            st.append(Paragraph("1.%d %s" % (job.sheets.index((sname, (a, b), tasks)) + 1, sname), H2))
            parts = []
            for tk in ts:
                act = sorted(tk.act)
                parts.append("%s（%d/%d%s）" % (tk.name, act[0].month, act[0].day, "〜%d/%d" % (act[-1].month, act[-1].day) if len(act) > 1 else ""))
            st.append(Paragraph("停止 %d月%d日〜立上げ %d月%d日。" % (a.month, a.day, b.month, b.day) + "、".join(parts) + "を行った。" if job.cid.startswith("T") else "、".join(parts) + "を行った。", BODY))
            fnd = [FIND[(job.fy, e.no)] for tk in ts for e, m in tk.eqs if (job.fy, e.no) in FIND]
            if fnd:
                st.append(Paragraph("主な所見：" + "。".join(dict.fromkeys(fnd)) + "。", BODY))
            st.append(Spacer(1, 4))
    st.append(Paragraph("2. 工程の実績", H1))
    data = [["日付", "班", "主な作業", "人員", "工数（人日）"]]
    for x in job.days:
        n = sum(k for _, k in x["crew"])
        data.append([P("%d/%d" % (x["d"].month, x["d"].day)), P(x["team"]), P(x["lines"][0] if x["lines"] else ""), P("%d 名" % n), P("%.2f" % x["md"])])
        rec({"区分": "工程実績", "日付": x["d"].isoformat(), "班": x["team"], "主な作業": x["lines"][0] if x["lines"] else "", "人数": n, "manDays": round(x["md"], 2), "manDaysText": "%.2f" % x["md"]})
    tot = sum(x["md"] for x in job.days)
    data.append([P("計"), P(""), P(""), P(""), P("%.2f" % tot)])
    rec({"区分": "工程実績", "項目": "計", "manDays": round(tot, 2), "manDaysText": "%.2f" % tot}, kind="total")
    t = Table(data, colWidths=[16 * mm, 14 * mm, 98 * mm, 16 * mm, 20 * mm], repeatRows=1); t.setStyle(tstyle(8)); st.append(t)
    st.append(PageBreak())
    st.append(Paragraph("3. 測定の記録", H1))
    if job.cid == "R2023a":
        data = [["列", "測定枚数", "最大摩耗(mm)", "取替枚数", "取替後クリアランス(mm)"]]
        for i in range(12):
            data.append([str(i + 1), "20", "%.1f" % G.MAX_WEAR[i], str(G.REPLACE_BY_ROW[i]), "%.1f〜%.1f" % G.CLEAR_AFTER[i]])
            rec({"区分": "測定", "列": i + 1, "最大摩耗": G.MAX_WEAR[i], "取替枚数": G.REPLACE_BY_ROW[i]})
    elif job.cid.startswith("T"):
        data = [["号炉", "測定", "今回", "前回", "判定"]]
        for sname, ab, tasks in job.sheets:
            u = sname[0]
            for key, lab, unit, lim in (("grate", "燃焼火格子 最大摩耗", "mm", 12.0), ("boiler", "ボイラ水管 最小肉厚", "mm", 2.0), ("idf", "誘引通風機 振動", "mm/s", 4.5), ("bf", "BF 差圧", "kPa", 1.8)):
                nmk = {"grate": "燃焼火格子", "boiler": "廃熱ボイラ本体", "idf": "誘引通風機", "bf": "バグフィルタ本体"}[key]
                done = any(e.name == nmk and max(tk.act) <= FW.TODAY for tk in tasks for e, m in tk.eqs)
                if not done:
                    continue
                v = FW.MEAS[(key, u)][job.fy]; pv = FW.MEAS[(key, u)].get(job.fy - 1)
                ok = (v < lim) if key != "boiler" else (v > lim + 0.3)
                if key == "grate" and job.fy == 2023 and u == "1":
                    ok = False
                jd = "良" if ok else "要対策"
                data.append([sname, lab, "%s %s" % (v, unit), "%s %s" % (pv, unit) if pv else "—", jd])
                rec({"区分": "測定", "号炉": sname, "測定": lab, "今回": v, "前回": pv, "単位": unit, "判定": jd})
    else:
        data = [["項目", "結果"]]
        mrows = {"R2020a": [("通風試験 差圧（取替後）", "0.9kPa"), ("ろ布 取替本数", "540 本")], "R2021a": [("通風試験 差圧（取替後）", "0.9kPa"), ("ろ布 取替本数", "540 本")],
                 "R2022a": [("取替前 最小肉厚", "2.3mm"), ("放射線透過試験", "16 継手 合格"), ("水圧試験", "2.25MPa 10 分 漏れなし")],
                 "R2024a": [("パドル取替枚数", "24 枚"), ("試運転 電流", "定格の 72%")], "R2025a": [("取替前 最大摩耗", "14.8mm"), ("取替枚数", "30 枚（予備 18・再使用 12）"), ("取替後クリアランス", "2.0〜4.0mm")]}[job.cid]
        for a, b in mrows:
            data.append([P(a), P(b)]); rec({"区分": "測定", "項目": a, "結果": b})
    t = Table(data, repeatRows=1); t.setStyle(tstyle()); st.append(t)
    st.append(Spacer(1, 8))
    st.append(Paragraph("4. 使用した部品・材料", H1))
    data = [["品名", "仕様・図番", "数量", "使用箇所", "支給/手配"]]
    rows = REPAIR_PARTS.get(job.cid) or (None if job.cid == "R2023a" else parts_of(job))
    if job.cid == "R2023a":
        rows = [("火格子片（燃焼段）", "AN-G60-203", "72 枚（受入 90、残 18 は予備品として倉庫 B-3 へ）", "1号 燃焼火格子", "市支給（R4 購入品）"),
                ("可動フレーム支持ローラ", "AN-G60-311", "2 個", "1号 燃焼火格子", "市支給（倉庫の予備品）"),
                ("六角ボルト・ナット", "M20×70", "144 組（受入 150）", "同上", "当社手配"), ("溶接棒", "耐熱用 φ3.2", "3 kg", "同上", "当社手配")]
    for r in rows:
        data.append([P(x) for x in r]); rec({"区分": "使用部品", "品名": r[0], "仕様": r[1], "数量": r[2], "使用箇所": r[3], "支給/手配": r[4]})
    t = Table(data, colWidths=[34 * mm, 34 * mm, 26 * mm, 40 * mm, 36 * mm], repeatRows=1); t.setStyle(tstyle()); st.append(t)
    st.append(PageBreak())
    st.append(Paragraph("5. 写真帳", H1))
    ph = []
    if job.cid.startswith("T"):
        for sname, ab, tasks in job.sheets:
            for tk in tasks:
                if tk.kind in ("grate", "boiler", "bf") and max(tk.act) <= FW.TODAY:
                    ph.append((tk.kind, "%s %s %d/%d" % (sname, tk.name.split("・")[0], tk.act[0].month, tk.act[0].day)))
    else:
        k = PHOTO_KIND.get(job.cid, "grate")
        for tk in job.sheets[0][2][:6]:
            if tk.act:
                ph.append((k, "%s %d/%d" % (tk.name.split("（")[0][:20], tk.act[0].month, tk.act[0].day)))
    ph = ph[:6] or [("x", "作業状況")]
    if len(ph) % 2:
        ph.append(("x", "作業状況（全景）"))
    tbl = []
    for i in range(0, len(ph), 2):
        tbl.append([Photo(78 * mm, 55 * mm, *ph[i]), Photo(78 * mm, 55 * mm, *ph[i + 1])])
        for j in (i, i + 1):
            rec({"区分": "写真", "説明": "写真%d　%s" % (j + 1, ph[j][1])}, kind="item", image="写真%d" % (j + 1))
    for row in tbl:
        for f in row:
            f.cap = "写真%d　%s" % (tbl.index(row) * 2 + row.index(f) + 1, f.cap)
    t = Table(tbl, colWidths=[85 * mm, 85 * mm]); t.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("BOTTOMPADDING", (0, 0), (-1, -1), 10)]))
    st.append(t)
    d.build(st)
    return fn, recs


def build(job, rnd):
    if job.cid == "T2019":
        return build_old(job)
    if job.cid == "T2026":
        return None, None   # 年度末提出（記録の時点では未作成）
    return build_new(job, rnd)
