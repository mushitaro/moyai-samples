# -*- coding: utf-8 -*-
"""工事報告書（文字のある PDF）全年度
 メーカー：H24〜H27 文章型（総括表なし）→ H28〜 点検結果総括表付き（Word→PDF）
 浜風機工：完了報告書（Excel の写真帳→PDF）
 原価は書かない。"""
import os, datetime as dt
from docx import Document
from docx.shared import Pt, Mm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from PIL import Image, ImageDraw, ImageFont
from openpyxl.drawing.image import Image as XLImage
import world as W, cw, sg, sched as S
from xl import *
import render
from gen_docs import label, jdate_short, jd

FONT = "IPAPGothic"
FONT_PATH = "/usr/share/fonts/opentype/ipafont-gothic/ipagp.ttf"
IMG_DIR = os.path.join(BUILD, "img")

def photo(name, caption, kind="box"):
    os.makedirs(IMG_DIR, exist_ok=True)
    p = os.path.join(IMG_DIR, name + ".png")
    if os.path.exists(p): return p
    im = Image.new("RGB", (640, 440), (205, 205, 200))
    d = ImageDraw.Draw(im)
    f = ImageFont.truetype(FONT_PATH, 22); fs = ImageFont.truetype(FONT_PATH, 16)
    if kind == "chain":
        for x in range(40, 600, 46): d.ellipse((x, 200, x + 40, 240), outline=(60, 60, 60), width=5)
        d.rectangle((30, 260, 610, 300), fill=(120, 110, 100))
    elif kind == "grate":
        for i in range(6):
            d.polygon([(60 + i * 30, 120 + i * 40), (560 + i * 10, 120 + i * 40), (560 + i * 10, 150 + i * 40), (60 + i * 30, 150 + i * 40)], outline=(50, 50, 50), fill=(150, 140, 130))
    elif kind == "bag":
        for x in range(60, 600, 40): d.rectangle((x, 80, x + 24, 380), outline=(80, 80, 80), fill=(235, 235, 225))
    elif kind == "sprocket":
        d.ellipse((200, 100, 440, 340), outline=(50, 50, 50), width=8); d.ellipse((290, 190, 350, 250), fill=(80, 80, 80))
    elif kind == "pump":
        d.rectangle((150, 200, 330, 320), outline=(60, 60, 60), width=5); d.ellipse((330, 190, 470, 330), outline=(60, 60, 60), width=5)
        d.rectangle((100, 320, 520, 340), fill=(110, 110, 110))
    elif kind == "fan":
        d.ellipse((180, 90, 460, 370), outline=(60, 60, 60), width=6)
        for a in range(0, 360, 45):
            import math
            x = 320 + 120 * math.cos(math.radians(a)); y = 230 + 120 * math.sin(math.radians(a))
            d.line((320, 230, x, y), fill=(70, 70, 70), width=6)
    else:
        d.rectangle((120, 100, 520, 340), outline=(70, 70, 70), width=4)
    d.rectangle((0, 0, 640, 34), fill=(60, 60, 60))
    d.text((10, 5), "【架空】見本用の写真代替図", font=fs, fill=(255, 255, 255))
    d.text((20, 395), caption[:26], font=f, fill=(20, 20, 20))
    im.save(p)
    return p

def kind_of(e):
    return {"pump": "pump", "fan": "fan", "idf": "fan", "blower": "fan", "comp": "pump", "grate": "grate", "bf": "bag", "ashconv": "chain", "conv": "chain"}.get(e.cls, "box")

# ---------- docx ----------
def base_doc(title, author):
    doc = Document()
    sec = doc.sections[0]
    sec.page_height = Mm(297); sec.page_width = Mm(210)
    sec.left_margin = sec.right_margin = Mm(20); sec.top_margin = Mm(20); sec.bottom_margin = Mm(18)
    st = doc.styles["Normal"]
    st.font.name = FONT; st.font.size = Pt(10.5)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    for s_ in ("Heading 1", "Heading 2", "Heading 3", "Title"):
        hs = doc.styles[s_]; hs.font.name = FONT; hs.font.color.rgb = RGBColor(0, 0, 0)
        rpr = hs.element.get_or_add_rPr(); rf = rpr.find(qn("w:rFonts"))
        if rf is None:
            rf = OxmlElement("w:rFonts"); rpr.append(rf)
        rf.set(qn("w:eastAsia"), FONT); rf.set(qn("w:ascii"), FONT); rf.set(qn("w:hAnsi"), FONT)
    doc.styles["Heading 1"].font.size = Pt(14)
    doc.styles["Heading 2"].font.size = Pt(12)
    cp = doc.core_properties
    cp.title = title; cp.author = author; cp.subject = "【架空】試験用サンプル"
    fp = sec.footer.paragraphs[0]; fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for tag, txt in (("begin", None), (None, "PAGE"), ("end", None)):
        run = fp.add_run()
        if tag:
            el = OxmlElement("w:fldChar"); el.set(qn("w:fldCharType"), tag); run._r.append(el)
        else:
            el = OxmlElement("w:instrText"); el.set(qn("xml:space"), "preserve"); el.text = txt; run._r.append(el)
    return doc

def para(doc, text, size=None, bold=False, align=None, space_after=4, indent=None):
    p = doc.add_paragraph()
    r = p.add_run(text); r.bold = bold
    if size: r.font.size = Pt(size)
    if align: p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    if indent: p.paragraph_format.left_indent = Mm(indent)
    return p

def table(doc, header, rows, widths_mm, size=8.5, header_fill="D9E1F2"):
    t = doc.add_table(rows=1, cols=len(header))
    t.style = "Table Grid"; t.alignment = WD_TABLE_ALIGNMENT.CENTER
    hr = t.rows[0]
    trPr = hr._tr.get_or_add_trPr(); th = OxmlElement("w:tblHeader"); th.set(qn("w:val"), "true"); trPr.append(th)
    for i, h in enumerate(header):
        c = hr.cells[i]; c.text = ""
        run = c.paragraphs[0].add_run(h); run.bold = True; run.font.size = Pt(size)
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        sh = OxmlElement("w:shd"); sh.set(qn("w:val"), "clear"); sh.set(qn("w:fill"), header_fill)
        c._tc.get_or_add_tcPr().append(sh)
    for row in rows:
        rc = t.add_row().cells
        for i, v in enumerate(row):
            rc[i].text = ""
            run = rc[i].paragraphs[0].add_run(str(v)); run.font.size = Pt(size)
    t.autofit = False
    tblPr = t._tbl.tblPr
    lay = OxmlElement("w:tblLayout"); lay.set(qn("w:type"), "fixed"); tblPr.append(lay)
    for i, gc in enumerate(t._tbl.tblGrid.findall(qn("w:gridCol"))):
        gc.set(qn("w:w"), str(int(widths_mm[i] * 56.7)))
    for row in t.rows:
        for i, w in enumerate(widths_mm): row.cells[i].width = Mm(w)
    return t

class Seq:
    def __init__(s): s.n = 0
    def __call__(s, kind, fields, works=(), image=None):
        rec(None, None, kind, fields, works, seq=s.n, image=image); s.n += 1

def trows(doc, sq, header, rows, widths_mm, size=8.5, works=None, extra=None):
    table(doc, header, [r for r in rows], widths_mm, size)
    for i, r in enumerate(rows):
        f = {h: v for h, v in zip(header, r)}
        if extra: f.update(extra)
        sq("item", f, works[i] if works else ())

def docx_to_pdf(doc, fname):
    os.makedirs(BUILD, exist_ok=True)
    dpath = os.path.join(BUILD, fname.replace(".pdf", ".docx"))
    doc.save(dpath)
    return render.to_pdf(dpath, OUT)

def cover(doc, c, title2, when, fy):
    for _ in range(5): para(doc, "")
    para(doc, "【架空】", 12, align=WD_ALIGN_PARAGRAPH.CENTER)
    para(doc, c.name + ("（%s施工分）" % W.wareki_long(fy) if c.kind == "EN" else ""), 16, True, WD_ALIGN_PARAGRAPH.CENTER)
    para(doc, title2, 22, True, WD_ALIGN_PARAGRAPH.CENTER, 30)
    for _ in range(9): para(doc, "")
    para(doc, when, 12, align=WD_ALIGN_PARAGRAPH.CENTER)
    para(doc, "しおさい町長 様", 12, align=WD_ALIGN_PARAGRAPH.CENTER)
    para(doc, W.MAKER, 13, True, WD_ALIGN_PARAGRAPH.CENTER)
    doc.add_page_break()

def jmonth(d):
    if d >= dt.date(2019, 5, 1): return "令和%d年%d月" % (d.year - 2018, d.month)
    return "平成%d年%d月" % (d.year - 1988, d.month)

def jday(d):
    return jdate_short(d)

# ---------- メーカー 旧様式（文章型） ----------
def mk_old(c, s, fy=None):
    fy = fy or c.fy
    fname = "【架空】しおさい町ごみ処理場_工事報告書_%s_水樹プラント文章型様式.pdf" % label(c, fy)
    rec_open(fname)
    sq = Seq()
    doc = base_doc("【架空】" + c.name + " 報告書", "水樹プラント v1")
    sub = s.site[1] + dt.timedelta(days=45)
    cover(doc, c, "報　告　書" if c.kind == "MK" else "施工報告書", jmonth(sub), fy)
    mh = sum(l.crew * l.hours for d in s.days for l in d.lines)
    doc.add_heading("1．業務概要" if c.kind == "MK" else "1．工事概要", 1)
    info = [("業務名称" if c.kind == "MK" else "工事名称", c.name), ("履行場所", "しおさい町ごみ処理場（【架空】）"),
            ("履行期間", "%s 〜 %s" % (jday(c.start), jday(c.end))),
            ("現地作業", "%s 〜 %s（日曜休工）" % (jday(s.site[0]), jday(s.site[1]))),
            ("作業責任者", "v1（作業員 延べ %d 人日）" % round(mh / 8.0))]
    for i, (k, v) in enumerate(info, 1):
        para(doc, "（%d）%s　　%s" % (i, k, v)); sq("item", {"項目": k, "内容": v})
    doc.add_heading("2．作業の概要", 1)
    if c.kind == "MK":
        if len(s.stop) == 2:
            t = ("本業務は、1号炉および2号炉の焼却設備、燃焼ガス冷却設備、排ガス処理設備、通風設備ならびに共通設備について、仕様書に定める点検整備を行ったものである。"
                 "1号炉は%d月%d日から%d月%d日、2号炉は%d月%d日から%d月%d日に停止して作業を行い、全炉停止日の%sに共通設備（ごみクレーン、空気槽等）の点検を行った。"
                 "各炉とも点検後に立上げを行い、運転状態に異常のないことを確認した。") % (
                (s.stop[1] + dt.timedelta(1)).month, (s.stop[1] + dt.timedelta(1)).day, s.startup[1].month, s.startup[1].day,
                (s.stop[2] + dt.timedelta(1)).month, (s.stop[2] + dt.timedelta(1)).day, s.startup[2].month, s.startup[2].day,
                "・".join("%d月%d日" % (d.month, d.day) for d in s.common))
        else:
            f = list(s.stop)[0]
            t = ("本年度は1号炉が長寿命化改良工事のため、本業務は%d号炉および共通設備を対象とした。%d号炉は%d月%d日から%d月%d日に停止して作業を行った。" %
                 (f, f, (s.stop[f] + dt.timedelta(1)).month, (s.stop[f] + dt.timedelta(1)).day, s.startup[f].month, s.startup[f].day))
    else:
        t = "本工事は、長寿命化計画に基づき%d号炉の火格子、ガス冷却室耐火物、誘引通風機（インバータ化）、燃焼制御盤等を更新したものである。" % list(s.stop)[0]
    para(doc, t); sq("item", {"節": "作業の概要", "本文": t})
    doc.add_heading("3．点検整備の内容" if c.kind == "MK" else "3．施工内容", 1)
    bysys = {}
    for r in s.findings:
        bysys.setdefault(r["sys"] or "施工", []).append(r)
    k = 0
    for sysn, rows in bysys.items():
        k += 1
        if c.kind == "MK":
            sent = []
            for r in rows:
                if r["act"] not in ("－", ""):
                    sent.append("%s%sは%sであったため、%s。" % (r["gou"] + "炉" if r["gou"] != "共通" else "", r["eq"], r["obs"], r["act"].rstrip("。") + "を行った" if not r["act"].endswith("取替") else r["act"] + "した"))
                else:
                    sent.append("%s%sは%s。" % (r["gou"] + "炉" if r["gou"] != "共通" else "", r["eq"], r["obs"]))
            body = "".join(sent)
        else:
            body = "".join("%s（%s）を施工した。" % (r["eq"], r["obs"]) for r in rows)
        para(doc, "（%d）%s" % (k, sysn), bold=True, space_after=2)
        para(doc, body, indent=5, space_after=8)
        sq("item", {"節": sysn, "本文": body}, [w for r in rows for w in r["works"]])
    doc.add_heading("4．所見", 1)
    bad = [r for r in s.findings if r["judge"] in ("C", "D") and r["rec"] not in ("－", "")]
    if c.kind == "MK":
        t = "竣工から%d年が経過し、" % (fy - 1994) + ("、".join(sorted({r["eq"] for r in bad})) or "各設備") + "の劣化が進んでいる。" + \
            ("町で策定中の長寿命化計画に合わせて計画的に更新されることを提案する。" if fy == 2014 else ("これらは定期点検整備の範囲で延命を続けることが難しい段階にあり、計画的な更新を提案する。" if fy < 2015 else "次年度以降も測定を継続されたい。"))
    else:
        t = "施工はすべて完了し、試運転で所定の性能を確認した。"
    para(doc, t); sq("item", {"節": "所見", "本文": t})
    doc.add_heading("添付　測定記録", 1)
    rows = []; wk = []
    if c.kind == "MK":
        for (f, gm) in s.meas["grate"]: rows.append(("%d号炉" % f, "燃焼火格子 火格子片 残厚", "最小 %smm（限度 22mm）" % gm, "－")); wk.append([])
        for (f, fdf, pv, v) in s.meas["vib"]:
            rows.append(("%d号炉" % f, "誘引送風機 振動（反負荷側 水平）", "%.1fmm/s" % v, "更新検討" if v >= 4.5 else ("経過観察" if v >= 2.8 else "－"))); wk.append([])
            rows.append(("%d号炉" % f, "押込送風機 振動", "%.1fmm/s" % fdf, "－")); wk.append([])
    else:
        for (nm, sp, q, kb, place, ws) in s.materials:
            rows.append((place, nm, sp, q)); wk.append(ws)
    hdr = ["号炉", "測定項目", "測定値", "処置"] if c.kind == "MK" else ["場所", "品名", "仕様", "数量"]
    trows(doc, sq, hdr, rows, [22, 66, 50, 30], 9, wk)
    return docx_to_pdf(doc, fname)

# ---------- メーカー 新様式（総括表付き） ----------
def mk_new(c, s, fy=None):
    fy = fy or c.fy
    title_kind = "点検整備報告書" if c.kind == "MK" else ("竣工報告書" if c.kind == "EN" else "工事報告書")
    fname = "【架空】しおさい町ごみ処理場_工事報告書_%s_水樹プラント総括表付様式.pdf" % label(c, fy)
    rec_open(fname)
    sq = Seq()
    doc = base_doc("【架空】" + c.name + " " + title_kind, "水樹プラント v1")
    sub = D2 = (dt.date(2022, 12, 15) if (c.kind == "MK" and fy == 2022) else cw.wd_after(s.site[1], 30))
    for _ in range(5): para(doc, "")
    para(doc, "【架空】", 12, align=WD_ALIGN_PARAGRAPH.CENTER)
    para(doc, c.name + ("（%s施工分）" % W.wareki_long(fy) if c.kind == "EN" else ""), 17, True, WD_ALIGN_PARAGRAPH.CENTER)
    para(doc, title_kind, 24, True, WD_ALIGN_PARAGRAPH.CENTER, 30)
    table(doc, ["提出", "作成", "確認"], [(jday(sub), "v1", "")], [50, 30, 30], 10, "FFFFFF")
    for _ in range(8): para(doc, "")
    para(doc, "しおさい町長 様", 12, align=WD_ALIGN_PARAGRAPH.CENTER)
    para(doc, W.MAKER, 13, True, WD_ALIGN_PARAGRAPH.CENTER)
    para(doc, "環境プラント事業部 サービス課", 10, align=WD_ALIGN_PARAGRAPH.CENTER)
    doc.add_page_break()
    doc.add_heading("1．業務概要" if c.kind == "MK" else "1．工事概要", 1)
    mh = sum(l.crew * l.hours for d in s.days for l in d.lines)
    furn = sorted(s.stop) if s.stop else []
    info = [("業務名" if c.kind == "MK" else "工事名", c.name), ("履行場所", "しおさい町ごみ処理場（【架空】）"),
            ("履行期間", "%s 〜 %s" % (jday(c.start), jday(c.end))),
            ("現地作業期間", "%s 〜 %s（実働%d日、日曜休工）" % (jday(s.site[0]), jday(s.site[1]), len(s.days))),
            ("対象", "・".join("%d号炉" % f for f in furn) + ("（各20t/16h）および共通設備" if c.kind == "MK" else "") if furn else "ごみクレーン（共通）"),
            ("体制", "SV v1、" + "、".join(sorted({"%s班" % l.team for d in s.days for l in d.lines if l.team != "SV"})) + (("、協力会社 " + c.subcontract.replace("【架空】", "")) if c.subcontract else "")),
            ("延べ作業時間", "%.0f 人・時間（作業日報による）" % mh)]
    if c.kind == "MK":
        bet = [lc for lc in cw.contracts_of(fy, ("LOCAL", "CRUP")) if getattr(lc, "in_common", False) or (lc.site and s.site[0] <= lc.site[0] <= s.site[1] and lc.docs if lc.kind == "LOCAL" else False)]
        if bet:
            info.append(("別途工事", "、".join("%s（%s、%d/%d〜%d/%d）" % (lc.short, lc.vendor.replace("【架空】", "").replace("有限会社 ", "").replace("株式会社", ""), lc.site[0].month, lc.site[0].day, lc.site[1].month, lc.site[1].day) for lc in bet)))
    trows(doc, sq, ["項目", "内容"], info, [35, 135], 9)
    doc.add_heading("2．点検結果総括表" if c.kind == "MK" else "2．施工結果総括表", 1)
    para(doc, "判定　A：良好　B：経過観察　C：要整備（本業務で処置したものを含む）　D：早急な対応が必要" if c.kind == "MK" else "判定　完了：施工完了　A：良好", 8.5)
    if c.kind == "MK" and fy == 2022:
        import gen_report_fixed as GR
        rows = [(i,) + r for i, r in enumerate(GR.SUMMARY, 1)]
        wk = [GR.works_for(c, r) for r in GR.SUMMARY]
    else:
        rows = [(i, r["gou"], r["sys"], r["eq"], r["judge"], r["obs"], r["act"], r["rec"]) for i, r in enumerate(s.findings, 1)]
        wk = [r["works"] for r in s.findings]
    trows(doc, sq, ["No", "号炉", "設備", "機器", "判定", "所見", "処置（本業務）", "次回推奨"], rows, [8, 10, 22, 28, 9, 38, 30, 30], 7.5, wk)
    doc.add_heading("3．点検整備内容" if c.kind == "MK" else "3．施工内容", 1)
    if c.kind == "MK" and fy == 2022:
        for h, t in GR.SECTIONS:
            doc.add_heading(h, 2); para(doc, t, space_after=6); sq("item", {"節": h, "本文": t})
    else:
        groups = {}
        for r in s.findings: groups.setdefault(r["gou"], []).append(r)
        for i, (g, rs) in enumerate(groups.items(), 1):
            if g in ("1号", "2号"):
                f = int(g[0])
                h = "3.%d　%s炉（%d月%d日〜%d月%d日）" % (i, g, (s.stop[f] + dt.timedelta(1)).month, (s.stop[f] + dt.timedelta(1)).day, s.startup[f].month, s.startup[f].day) if f in s.stop else "3.%d　%s炉" % (i, g)
            else:
                h = "3.%d　共通設備" % i + ("（%s）" % "・".join("%d月%d日" % (d.month, d.day) for d in s.common) if s.common else "")
            sent = []
            for r in rs:
                if r["act"] not in ("－", ""):
                    sent.append("%sは%s。%sを行った。" % (r["eq"], r["obs"], r["act"]))
                elif r["judge"] not in ("A", "完了"):
                    sent.append("%sは%s。" % (r["eq"], r["obs"]))
            if not sent: sent = ["各機器に異常は認められなかった。"]
            t = "".join(sent)
            doc.add_heading(h, 2); para(doc, t, space_after=6)
            sq("item", {"節": h, "本文": t}, [w for r in rs for w in r["works"]])
    doc.add_heading("4．測定記録" if c.kind == "MK" else "4．試運転記録", 1)
    if c.kind == "MK" and fy == 2022:
        for (cap, hdr, rws, wd) in GR.MEAS:
            para(doc, cap, bold=True); trows(doc, sq, hdr, rws, wd, 8.5, extra={"表": cap}); para(doc, "")
    elif c.kind == "MK":
        para(doc, "（1）燃焼火格子 火格子片 最小残厚（mm、限度 22）", bold=True)
        trows(doc, sq, ["号炉", "最小", "判定"], [("%d号" % f, gm, "限度以下あり" if gm < 22 else "限度内") for f, gm in s.meas["grate"]], [20, 30, 40], 8.5, extra={"表": "火格子"})
        para(doc, "")
        para(doc, "（2）回転機 振動（mm/s、反負荷側 水平）", bold=True)
        rws = []
        for (f, fdf, pv, v) in s.meas["vib"]:
            rws.append(("%d号" % f, "押込送風機", fdf)); rws.append(("%d号" % f, "誘引通風機", "%.1f（前年 %s）" % (v, pv if pv else "－")))
        trows(doc, sq, ["号炉", "機器", "今回"], rws, [20, 40, 50], 8.5, extra={"表": "振動"})
        para(doc, "")
        para(doc, "（3）バグフィルタ 差圧（kPa、運転時）", bold=True)
        trows(doc, sq, ["号炉", "点検前"], [("%d号" % f, dp) for f, dp in s.meas["dp"]], [20, 30], 8.5, extra={"表": "差圧"})
        if s.meas.get("hopper"):
            para(doc, ""); para(doc, "（4）ガス冷却室 下部ホッパ 板厚（mm、原厚 6.0）", bold=True)
            trows(doc, sq, ["号炉", "最小"], [("%d号" % f, hp) for f, hp in s.meas["hopper"]], [20, 30], 8.5, extra={"表": "板厚"})
    else:
        rws = [("排ガス（煙突出口）", "ばいじん", "0.01 g/m³N 未満", "適合"), ("排ガス", "CO（O2 12%換算）", "28 ppm", "適合"),
               ("誘引通風機", "振動", "1.5 mm/s", "良好"), ("燃焼室", "温度（定格時）", "870〜920 ℃", "良好")] if c.kind == "EN" else \
              [("ごみクレーン", "定格荷重試験", "3.2t", "合格"), ("ごみクレーン", "走行・横行・巻上速度", "定格どおり", "良好")]
        trows(doc, sq, ["箇所", "項目", "結果", "判定"], rws, [40, 45, 45, 25], 8.5, extra={"表": "試運転"})
    doc.add_heading("5．交換部品・使用材料", 1)
    if c.kind == "MK" and fy == 2022:
        trows(doc, sq, ["品名", "仕様", "数量", "区分", "使用箇所"], GR.MATS, [30, 38, 18, 32, 52], 8.5, GR.mats_works(c))
        para(doc, GR.MATS_NOTE, 8.5)
    else:
        trows(doc, sq, ["品名", "仕様", "数量", "区分", "使用箇所"], [m[:5] for m in s.materials], [30, 38, 18, 32, 52], 8.5, [m[5] for m in s.materials])
    doc.add_heading("6．写真", 1)
    if c.kind == "MK" and fy == 2022:
        phs = [("p_grate", "1号炉 燃焼火格子 火格子片取替後", "grate"), ("p_bag", "1号炉 バグフィルタ 新ろ布取付状況", "bag"),
               ("p_hop", "2号炉 ガス冷却室 下部ホッパ 板厚測定", "box"), ("p_crane", "ごみクレーン 年次自主検査", "box")]
    else:
        phs = []
        for r in s.findings:
            if r["works"] and r["judge"] in ("C", "D", "完了") and len(phs) < 4:
                e = W.BYMID[r["works"][0].mid]
                phs.append(("p%d_%s" % (fy, e.mid), ("%s %s" % (r["gou"] + "炉" if r["gou"] in ("1号", "2号") else "", r["eq"].split("（")[0].split("・")[0])).strip() + " 状況", kind_of(e)))
        if not phs: phs = [("p%d_crane" % fy, "ごみクレーン 年次自主検査", "box")]
    for (nm, cap, kd) in phs:
        p = photo(nm, cap, kd)
        doc.add_picture(p, width=Mm(110)); doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
        para(doc, "写真　" + cap, 9, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=10)
        sq("item", {"写真": cap}, image=os.path.basename(p))
    doc.add_heading("7．所見と提言", 1)
    if c.kind == "MK" and fy == 2022:
        tips = GR.TIPS
    else:
        tips = []
        for r in s.findings:
            if r["judge"] in ("B", "C", "D") and r["rec"] not in ("－", ""):
                tips.append("（%d）%s%sは%s。%sを推奨します。" % (len(tips) + 1, r["gou"] + "炉" if r["gou"] in ("1号", "2号") else "", r["eq"], r["obs"], r["rec"]))
        if not tips: tips = ["（1）特記事項はありません。次年度も同じ周期で点検を継続してください。"]
    for t in tips:
        para(doc, t, indent=3, space_after=4); sq("item", {"節": "所見と提言", "本文": t})
    return docx_to_pdf(doc, fname)

# ---------- 浜風機工 完了報告（写真帳） ----------
def local_report(c, s):
    who = "g1" if c.fy <= 2018 else "g2"
    fname = "【架空】しおさい町ごみ処理場_工事報告書_%s_浜風機工%s写真帳様式.pdf" % (label(c), who)
    rec_open(fname)
    sq = Seq()
    wb = new_wb("【架空】" + c.name + " 完了報告書", "浜風機工 g1", dt.datetime(c.kanryo.year, c.kanryo.month, c.kanryo.day, 17, 0), who)
    ws = wb.active; ws.title = "表紙"
    widths(ws, [12, 20, 20, 20, 12])
    for r in range(1, 6): ws.row_dimensions[r].height = 30
    merge(ws, 6, 1, 6, 5, "完　了　報　告　書", F(24, True), AC, None); ws.row_dimensions[6].height = 50
    merge(ws, 9, 1, 9, 5, "【架空】" + c.name, F(13, True), AC, None)
    rows = [("工　期", "%s 〜 %s" % (jday(c.start), jday(c.end))), ("施工日", "%s 〜 %s" % (jday(s.site[0]), jday(s.site[1]).split("年")[-1]) if s.site[0] != s.site[1] else jday(s.site[0])),
            ("完了日", jday(s.site[1])), ("施工場所", "しおさい町ごみ処理場")]
    for i, (k, v) in enumerate(rows):
        put(ws, 12 + i, 2, k, F(11), AC, B_ALL); merge(ws, 12 + i, 3, 12 + i, 4, v, F(11), AL); ws.row_dimensions[12 + i].height = 26
        sq("item", {"項目": k.replace("　", ""), "内容": v})
    put(ws, 19, 2, "しおさい町長 様", F(12))
    put(ws, 24, 3, jday(c.kanryo), F(11)); put(ws, 25, 3, "【架空】有限会社 浜風機工", F(12, True)); put(ws, 26, 3, "%s　%s" % ("所長" if who == "g2" else "現場代理人", who), F(11))
    page(ws, landscape=False, area="A1:E28")
    ws = wb.create_sheet("工事概要")
    widths(ws, [5, 30, 22, 10, 26])
    put(ws, 1, 1, "工事概要", F(14, True))
    put(ws, 3, 1, "1．施工内容", F(11, True))
    if c.key == "R4-CH":
        items = ["旧チェーン（2条）撤去、伸び測定 2.4%", "ドライブスプロケット 2枚 取替", "テールスプロケット 2枚 取替",
                 "【追加】テール軸受 UCP212 2個 取替（10/20 損傷発見、町と協議のうえ追加）", "新チェーン 2条 据付、スクレーパ 再取付（既設 良品を流用）",
                 "張り調整、試運転（無負荷 30分、1号炉立上げ後 負荷確認）"]
        iw = [c.works] * len(items)
    elif c.key == "H30-ASH":
        items = ["灰抜き・コンベヤ内清掃（町職員協力）", "不良リンク 8リンク交換（ピン・ブシュ共）", "ドライブスプロケット 歯面肉盛溶接・仕上げ", "スクレーパ 6枚取替", "テイクアップ張り調整、試運転（無負荷 30分）"]
        iw = [c.works] * len(items)
    else:
        items = []; iw = []
        for h in s.htasks:
            if h["name"].startswith(("材料搬入", "KY")): continue
            items.append(("【追加】" + h["name"].replace("（追加）", "") if h["extra"] else h["name"])); iw.append(h["works"])
    r = 3
    for i, t in enumerate(items, 1):
        put(ws, r + i, 1, i, F(10), AC); merge(ws, r + i, 2, r + i, 5, t, F(10, color="C00000" if "追加" in t else None), AL, None)
        sq("item", {"No": i, "施工内容": t}, iw[i - 1])
    r = r + len(items) + 2
    put(ws, r, 1, "2．使用材料", F(11, True)); r += 1
    for i, h in enumerate(["No", "品名", "規格", "数量", "備考"], 1):
        put(ws, r, i, h, F(9, True), AC, B_ALL, fill("DDEBF7"))
    for i, m in enumerate(s.materials, 1):
        r += 1
        row = (i, m[0], m[1], m[2], m[3])
        for j, v in enumerate(row, 1): put(ws, r, j, v, F(9), AC if j in (1, 4) else AL, B_ALL)
        sq("item", {"No": i, "品名": m[0], "規格": m[1], "数量": m[2], "備考": m[3]}, m[4])
    r += 2
    put(ws, r, 1, "3．撤去品", F(11, True)); r += 1
    merge(ws, r, 2, r, 5, "撤去品は町の指示により場内金属くずヤードへ搬出", F(10), AL, None); r += 2
    put(ws, r, 1, "4．作業人員", F(11, True)); r += 1
    mh = sum(l.hours * (l.crew if not hasattr(d, "hrs") else 1) if not hasattr(d, "hrs") else d.hrs for d in s.days for l in d.lines[:1]) if c.key != "H30-ASH" else 72.5
    if c.key == "R4-CH": mh = 92.0
    merge(ws, r, 2, r, 5, "延べ %.1f 時間　作業日報による" % mh, F(10), AL, None)
    sq("item", {"項目": "作業人員", "内容": "延べ %.1f 時間" % mh, "manDays": round(mh / 8.0, 2), "manDaysText": "%.1f 時間" % mh})
    page(ws, landscape=False, area="A1:E%d" % (r + 1))
    # 写真帳
    if c.key == "R4-CH":
        phs = [("c1", "施工前 チェーン伸び測定", "chain", "旧チェーン 伸び 2.4%", "10/19"), ("c2", "旧チェーン撤去", "chain", "2条 撤去", "10/19"),
               ("c3", "ドライブスプロケット 取替前", "sprocket", "歯先摩耗", "10/20"), ("c4", "ドライブスプロケット 取替後", "sprocket", "", "10/20"),
               ("c5", "テール軸受 損傷状況", "box", "ガタ・異音（追加）", "10/20"), ("c6", "テール軸受 取替後", "box", "UCP212", "10/21"),
               ("c7", "新チェーン 据付後", "chain", "張り調整済", "10/21"), ("c8", "試運転状況", "chain", "無負荷 30分 異常なし", "10/21"), ("c9", "完了 全景", "box", "", "10/21")]
        koushu = "灰出しコンベヤ"
    else:
        phs = []
        es = [W.BYMID[w.mid] for w in c.works]
        for e in es:
            kd = kind_of(e)
            nm = (("%d号 " % e.furnace) if e.furnace else "") + e.name
            dd = "%d/%d" % (s.site[0].month, s.site[0].day); de = "%d/%d" % (s.site[1].month, s.site[1].day)
            phs += [("%s_%s_1" % (c.cid, e.mid), nm + " 施工前", kd, "", dd), ("%s_%s_2" % (c.cid, e.mid), nm + " 施工中", kd, "", dd),
                    ("%s_%s_3" % (c.cid, e.mid), nm + " 施工後", kd, "試運転 異常なし", de)]
        if s.extra:
            phs.insert(2, ("%s_x" % c.cid, s.extra[0].replace("（追加）", "")[:14] + " 状況", "box", "追加作業", "%d/%d" % (s.site[1].month, s.site[1].day)))
        koushu = W.SYS[es[0].sys]
    npg = (len(phs) + 2) // 3
    for pg in range(npg):
        ws = wb.create_sheet("写真帳%d" % (pg + 1))
        widths(ws, [44, 12, 28])
        merge(ws, 1, 1, 1, 3, "写　真　帳　（%d／%d）" % (pg + 1, npg), F(13, True), AC, None)
        put(ws, 2, 1, "【架空】" + c.name, F(8))
        for k in range(3):
            i = pg * 3 + k
            if i >= len(phs): break
            nm, cap, kd, memo, dd = phs[i]
            top = 4 + k * 14
            for rr in range(top, top + 13): ws.row_dimensions[rr].height = 17
            p = photo(nm, cap, kd)
            img = XLImage(p); img.width = 300; img.height = 206
            ws.add_image(img, "A%d" % (top + 1))
            for rr in range(top, top + 13):
                ws.cell(row=rr, column=1).border = Border(left=thin, right=thin, top=thin if rr == top else None, bottom=thin if rr == top + 12 else None)
            for j, (k1, v1) in enumerate([("No.", i + 1), ("工種", koushu), ("状況", cap), ("記事", memo), ("日付", dd)]):
                merge(ws, top + j * 2, 2, top + j * 2 + 1, 2, k1, F(9, True), AC, B_ALL, fill("F2F2F2"))
                merge(ws, top + j * 2, 3, top + j * 2 + 1, 3, v1, F(9), AL)
            merge(ws, top + 10, 2, top + 12, 3, None, None)
            sq("item", {"No.": i + 1, "工種": koushu, "状況": cap, "記事": memo, "日付": dd}, c.works, image=os.path.basename(p))
        page(ws, landscape=False, area="A1:C45")
    tmp = os.path.join(BUILD, "tmp_" + c.cid + ".xlsx")
    wb.save(tmp)
    pdf = render.to_pdf(tmp, BUILD)
    final = os.path.join(OUT, fname)
    os.replace(pdf, final)
    os.remove(tmp)
    return final
