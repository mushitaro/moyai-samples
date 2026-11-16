# -*- coding: utf-8 -*-
"""工事報告書（文字のある PDF）
H16〜H28 r1：Word の文章型 → PDF ／ H29〜 r2：Excel の表型（工程No. 順サマリー＋記録表）→ PDF
"""
import datetime as dt
import os
from collections import defaultdict

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt
from openpyxl import Workbook
from openpyxl.utils import get_column_letter as CL

import ans
import core
import render as RD
import texts
from xlhelp import *

D = dt.date.fromisoformat
UNAME = {"1": "1号炉", "2": "2号炉", "3": "3号炉", "C": "共通"}

MAT = {  # cls → (品名, 仕様, 数量) 分解整備・更新で使う主な材料
    "pump_dup": ("メカニカルシール", "ポンプ用", "1 組"), "pump_corr": ("メカニカルシール", "耐食ポンプ用", "1 組"),
    "pump_bfp": ("メカニカルシール・軸受", "給水ポンプ用", "1 式"), "pump_single": ("グランドパッキン", "ポンプ用", "1 組"),
    "pump_dosing": ("ダイヤフラム", "注入ポンプ用", "1 枚"), "pump_slurry": ("インペラ", "スラリーポンプ用", "1 個"),
    "mv": ("グランドパッキン", "弁用", "1 組"), "mv_hp": ("グランドパッキン・ガスケット", "高圧弁用", "1 組"), "hv": ("ガスケット", "", "1 枚"),
    "hv_hp": ("ガスケット", "高圧用", "1 枚"), "hv_corr": ("弁座ゴム", "ゴムライニング弁用", "1 個"), "cv": ("グランドパッキン", "調節弁用", "1 組"),
    "cv_hot": ("弁体・弁座", "調節弁用", "1 式"), "cv_corr": ("ダイヤフラム", "調節弁用", "1 枚"), "grate": ("火格子片", "火格子用", "24 枚"),
    "grate_wc": ("火格子片", "水冷火格子用", "36 枚"), "furnace": ("キャスタブル", "高アルミナ質", "0.6 t"), "conveyor": ("チェーン・スプロケット", "", "1 式"),
    "rv": ("グランドパッキン", "ロータリバルブ用", "1 組"), "damper": ("シール材", "ダンパ用", "1 式"), "damper_bf": ("ポペット弁シート", "", "1 枚"),
    "damper_idf": ("軸受", "ダンパ用", "2 個"), "fan_big": ("軸受", "送風機用", "2 個"), "fan_mid": ("軸受", "送風機用", "2 個"), "idf": ("軸受", "通風機用", "2 個"),
    "blower": ("Vベルト", "", "1 組"), "sootblower": ("ポペット弁・ランスパッキン", "", "1 式"), "sootblower_r": ("パッキン", "", "1 式"),
    "bf": ("ろ布", "PTFE φ150×6,000", "672 本"), "scr": ("脱硝触媒", "ハニカム 1 層分", "1 層"), "me": ("洗浄ノズル", "", "2 個"),
    "scrubber": ("ライニング補修材", "ビニルエステル FRP", "1 式"), "hyd": ("作動油・フィルタ", "", "1 式"), "hyd_drive": ("シリンダパッキン", "", "1 式"),
    "inst_hot": ("熱電対", "R 熱電対", "1 本"), "crane": ("ワイヤロープ", "φ28", "2 本"), "crane_sp": ("ワイヤロープ", "φ28", "2 本"),
    "crane_ash": ("ワイヤロープ", "φ20", "2 本"), "comp": ("オーバーホールキット", "", "1 式"), "kneader": ("パドル", "混練機用", "1 式"),
    "dewater": ("ろ布", "フィルタプレス用", "1 式"), "feeder": ("摺動板", "", "1 式"), "feeder_s": ("スクリュー", "", "1 本"),
    "ash_ext": ("摺動板", "灰押出装置用", "1 式"), "acc_fan": ("軸受", "ファン用", "2 個"), "sah": ("ガスケット", "", "1 式"), "dea": ("ガスケット", "", "1 式"),
    "pure": ("イオン交換樹脂", "", "1 式"), "battery": ("据置蓄電池", "MSE 形 54 セル", "1 組"), "ups": ("無停電電源装置", "", "1 台"),
    "eco": ("エコノマイザ管", "", "1 式"), "rh": ("再加熱器管束", "", "1 式"), "sh": ("過熱器管", "", "1 式"),
}


def material_rows(works, fy):
    agg = defaultdict(lambda: [0, set()])
    for w in works:
        if w["kind"] not in ("O", "R"):
            continue
        e = core.get().by_tag[w["tag"]]
        m = MAT.get(e["cls"])
        if not m:
            continue
        key = (m[0], m[1])
        agg[key][0] += 1
        agg[key][1].add(w["tag"])
    rows = []
    for (nm, sp), (n, tags) in sorted(agg.items()):
        tg = sorted(tags)
        place = "、".join(tg[:3]) + (f" ほか {len(tg) - 3}" if len(tg) > 3 else "")
        q = MAT[[k for k, v in MAT.items() if v[0] == nm and v[1] == sp][0]][2]
        qty = q if n == 1 else f"{q}×{n}"
        sup = "当社手配"
        yobi = "—"
        if nm == "ろ布" and fy >= 2018:
            sup = "市支給品（680 本）"
            yobi = "8 本 返納"
        rows.append((nm, sp, qty, place, sup, yobi))
    return rows


def window_rows(c):
    W = core.get()
    rows = []
    for win in sorted([w for w in W.windows.values() if w["contract"] == c["id"]], key=lambda w: W.window_dates(w)[0]):
        a, b, pa, pb = W.window_dates(win)
        n = len({d for t in win["tasks"] for d in t["days"]})
        note = "；".join(ep[6] for t in win["tasks"] for ep in t.get("episodes", []))[:60]
        rows.append((win, a, b, pa, pb, n, note))
    return rows


def works_of(c):
    W = core.get()
    return [w for w in W.works if w["contract"] == c["id"]]


def fname(c, form):
    lab = c["name"].replace(core.wa_fy(c["fy"]) + " ", "").replace("みなと臨海工場 ", "").replace(" ", "")
    if c["kind"] == "KIKAN":
        lab = f"基幹的設備改良工事{core.wa_fy(c['_rfy'])}出来高"
    return f"【架空】みなと臨海工場_工事報告書_{core.wa_fy(c['fy'] if c['kind'] != 'KIKAN' else c['_rfy'])}_{lab}_{form}"


# ================================================================ r1 docx
def _font(run, size=10.5, bold=False, name="IPAGothic"):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = name
    rpr = run._element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rpr.append(rf)
    for k in ("w:eastAsia", "w:ascii", "w:hAnsi"):
        rf.set(qn(k), name)


def para(doc, text, size=10.5, bold=False, align=None, space_after=4, indent=None):
    p = doc.add_paragraph()
    if align == "c":
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == "r":
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.space_after = Pt(space_after)
    if indent:
        p.paragraph_format.first_line_indent = Pt(indent)
    _font(p.add_run(text), size, bold)
    return p


def table(doc, header, rows, widths, size=8):
    t = doc.add_table(rows=1, cols=len(header))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(header):
        c = t.rows[0].cells[i]
        c.text = ""
        _font(c.paragraphs[0].add_run(h), size, True)
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        sh = OxmlElement("w:shd")
        sh.set(qn("w:val"), "clear")
        sh.set(qn("w:fill"), "D9D9D9")
        c._tc.get_or_add_tcPr().append(sh)
    trpr = t.rows[0]._tr.get_or_add_trPr()
    th = OxmlElement("w:tblHeader")
    th.set(qn("w:val"), "true")
    trpr.append(th)
    for row in rows:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = ""
            _font(cells[i].paragraphs[0].add_run(str(v)), size)
    t.autofit = False
    for i, gc in enumerate(t._tbl.tblGrid.findall(qn("w:gridCol"))):
        gc.set(qn("w:w"), str(int(widths[i] * 56.7)))
    for row in t.rows:
        for i, w in enumerate(widths):
            row.cells[i].width = Mm(w)
    return t


class Seq:
    def __init__(self, fn):
        self.fn, self.n = fn, 0

    def add(self, sheet, kind, fields, works=None, check=None):
        ans.rec(self.fn, sheet, None, kind, fields, works=works, check=check, seq=self.n)
        self.n += 1


def build_r1(c):
    W = core.get()
    fy = c["fy"]
    stem = fname(c, "文章型様式")
    fn = stem + ".pdf"
    ans.start(fn, False, "工事報告書", fy, "文章型（r1 Word）", "r1", c["id"])
    S = Seq(fn)
    doc = Document()
    sec = doc.sections[0]
    sec.page_height, sec.page_width = Mm(297), Mm(210)
    sec.left_margin = sec.right_margin = Mm(20)
    sec.top_margin, sec.bottom_margin = Mm(20), Mm(18)
    cp = doc.core_properties
    cp.title = f"【架空】{c['name']} 報告書"
    cp.author = "r1"
    cp.subject = "工事報告書"
    cp.keywords = "【架空】"
    cp.created = dt.datetime(fy + 1, 3, 20, 10, 0)
    for _ in range(6):
        para(doc, "")
    para(doc, "【架空】", 12, False, "c")
    para(doc, core.wa_fy(fy), 16, True, "c")
    para(doc, c["name"].replace(core.wa_fy(fy) + " ", ""), 18, True, "c", 10)
    para(doc, "報　告　書", 26, True, "c", 60)
    for _ in range(8):
        para(doc, "")
    para(doc, f"{d_wa_long(dt.date(fy + 1, 3, 1))[:-2]}", 12, False, "c")
    para(doc, "【架空】汐見重工環境株式会社", 14, True, "c")
    para(doc, "環境プラント本部　保守技術部", 11, False, "c")
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    para(doc, "1．業務概要", 12, True)
    for k, v in [("件　名", c["name"]), ("場　所", "【架空】みなと市臨海町 みなと臨海工場"),
                 ("履行期間", f"{d_wa_long(dt.date(fy, 4, 1))} 〜 {d_wa_long(dt.date(fy + 1, 3, 31))}"),
                 ("委託者", "【架空】みなと市 環境局 施設課"), ("受託者", f"【架空】汐見重工環境株式会社（現場代理人 {core.agent(fy)}）")]:
        para(doc, f"　{k}：{v}", 10.5)
    para(doc, "")
    para(doc, "2．作業実施期間", 12, True)
    rows = []
    wr = window_rows(c)
    for (win, a, b, pa, pb, n, note) in wr:
        rows.append((win["label"], f"{d_wa_long(a)} 〜 {d_wa_long(b)}", f"{n} 日", note or ""))
        S.add("2", "item", dict(区分=win["label"], 開始=a.isoformat(), 終了=b.isoformat(), 日数=n, 備考=note),
              works=[w for t in win["tasks"] for w in t["work_ids"]], check=[win["label"]])
    table(doc, ["区分", "現地作業期間", "日数", "備考"], rows, [28, 64, 16, 62], size=9)
    para(doc, "")
    para(doc, "3．実施内容", 12, True)
    for i, (win, a, b, pa, pb, n, note) in enumerate(wr, start=1):
        para(doc, f"3.{i}　{win['label']}", 11, True)
        tnames = [t["name"].split("（")[0] for t in win["tasks"] if t["works"]]
        ws = [w for t in win["tasks"] for w in t["works"]]
        nR = sum(1 for w in ws if w["kind"] == "R")
        nO = sum(1 for w in ws if w["kind"] == "O")
        txt = (f"{win['label']}は{d_wa_long(a)}に着手し、{'、'.join(tnames[:6])}{'ほか' if len(tnames) > 6 else ''}を実施した。"
               f"対象機器は延べ {len(ws)} 台で、うち分解整備・補修 {nO} 台、更新・交換 {nR} 台である。")
        for t in win["tasks"]:
            for ep in t.get("episodes", []):
                txt += ep[6]
        para(doc, txt, 10.5, indent=10.5)
    para(doc, "")
    para(doc, "4．点検整備記録", 12, True)
    rows = []
    for w in sorted(works_of(c), key=lambda w: (w["start"], w["tag"])):
        e = W.by_tag[w["tag"]]
        res, act = texts.result(w)
        rows.append((UNAME[e["unit"]], w["tag"], e["name"], texts.KIND_JA[w["kind"]], d_wa(w["start"]), res, act))
        S.add("4", "item", dict(号炉=UNAME[e["unit"]], code=w["tag"], name=e["name"], work=texts.KIND_JA[w["kind"]], date=w["start"],
                                結果=res, 処置=act), works=[w["id"]], check=[w["tag"]])
    table(doc, ["号炉", "タグNo.", "機器名称", "作業", "実施日", "結果", "処置"], rows, [12, 22, 44, 18, 18, 34, 24], size=7.5)
    para(doc, "")
    para(doc, "5．所見及び今後の整備について", 12, True)
    obs = observations(c)
    for i, o in enumerate(obs, start=1):
        para(doc, f"({i}) {o}", 10.5)
    para(doc, "")
    para(doc, "6．使用材料（主なもの）", 12, True)
    mr = material_rows(works_of(c), fy)
    table(doc, ["品名", "仕様", "数量", "使用箇所", "手配"], [r[:5] for r in mr], [42, 30, 20, 50, 28], size=8)
    for m in mr:
        S.add("6", "item", dict(品名=m[0], 仕様=m[1], 数量=m[2], 使用箇所=m[3], 手配=m[4]), check=[m[0]])
    para(doc, "")
    para(doc, "以上", 10.5, False, "r")
    for s in doc.sections:
        fp = s.footer.paragraphs[0]
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _font(fp.add_run("【架空】汐見重工環境株式会社"), 8)
    src = os.path.join(WORK, stem + ".docx")
    doc.save(src)
    pdf = RD.to_pdf(src, WORK)
    os.replace(pdf, os.path.join(OUT, fn))
    return fn


def observations(c):
    """所見：エピソード、翌年度補修の提案（world の発見）、更新の見込み"""
    W = core.get()
    out = []
    for win in [w for w in W.windows.values() if w["contract"] == c["id"]]:
        for t in win["tasks"]:
            for ep in t.get("episodes", []):
                out.append(ep[6])
    for w in works_of(c):
        res, act = texts.result(w)
        if "提案" in act:
            e = W.by_tag[w["tag"]]
            out.append(f"{UNAME[e['unit']]} {e['name']}（{w['tag']}）は劣化が進んでおり、次年度の補修を提案する。")
    if not out:
        out.append("各機器とも大きな異常は認められなかった。引き続き計画に沿った点検整備を行う。")
    return out[:12]


# ================================================================ r2 xlsx
def form_table(wb, name, title, header, widths, rows, S, recf, landscape=True, top=None, row_h=None, foot=None):
    ws = wb.create_sheet(name)
    F = lambda **k: font(FONT_P, **k)
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[CL(i)].width = w
    put(ws, 1, 1, title, F(size=13, bold=True))
    put(ws, 2, 1, "【架空】みなと臨海工場", F(size=9))
    r = 3
    for line in top or []:
        merge(ws, r, 1, r, len(header), line, F(size=9), AL_LW, None)
        r += 1
    r += 1
    hr = r
    for c, h in enumerate(header, start=1):
        put(ws, r, c, h, F(size=8, bold=True), AL_C, B_ALL, fill("DDEBF7"))
    ws.row_dimensions[r].height = 26
    r += 1
    for row in rows:
        vals, fields, works, chk = recf(row)
        for c, v in enumerate(vals, start=1):
            al = AL_C if (isinstance(v, (int, float)) or (isinstance(v, str) and len(v) <= 6)) else AL_LW
            put(ws, r, c, v, F(size=8), al, B_ALL)
        if row_h:
            ws.row_dimensions[r].height = row_h
        S.add(name, "item", fields, works=works, check=chk)
        r += 1
    for line in foot or []:
        r += 1
        merge(ws, r, 1, r, len(header), line, F(size=9), AL_LW, None)
    page(ws, landscape=landscape, fit_w=1, fit_h=0, area=f"A1:{CL(len(header))}{r}", titles=f"{hr}:{hr}", footer="&P / &N（" + name + "）")


def cover(wb, l1, l2, period, date_txt):
    ws = wb.active
    ws.title = "表紙"
    F = lambda **k: font(FONT_P, **k)
    for c in range(1, 9):
        ws.column_dimensions[CL(c)].width = 11
    put(ws, 6, 1, "【架空】", F(size=12))
    merge(ws, 9, 1, 9, 8, l1, F(size=16, bold=True), AL_C, None)
    merge(ws, 11, 1, 11, 8, l2, F(size=16, bold=True), AL_C, None)
    merge(ws, 14, 1, 14, 8, "業 務 完 了 報 告 書" if "工事" not in l2 else "工 事 報 告 書", F(size=24, bold=True), AL_C, None)
    merge(ws, 20, 2, 20, 7, f"履行期間　{period}", F(size=11), AL_C, None)
    merge(ws, 32, 1, 32, 8, date_txt, F(size=12), AL_C, None)
    merge(ws, 34, 1, 34, 8, "【架空】汐見重工環境株式会社", F(size=14, bold=True), AL_C, None)
    merge(ws, 35, 1, 35, 8, "環境プラント本部 保守技術部（報告書作成 r2）", F(size=10), AL_C, None)
    for r in range(1, 41):
        ws.row_dimensions[r].height = 18
    page(ws, landscape=False, fit_w=1, fit_h=1, area="A1:H40", margins=(0.8, 0.8, 0.8, 0.8))


def build_r2(c, kikan_fy=None):
    W = core.get()
    fy = c["fy"] if not kikan_fy else kikan_fy
    stem = fname(c, "表型サマリー様式")
    fn = stem + ".pdf"
    ans.start(fn, False, "工事報告書", fy, "表型サマリー（r2 Excel）", "r2", c["id"])
    S = Seq(fn)
    S.add("表紙", "item", dict(件名=c["name"], 年度=fy), check=[c["name"].split(" ")[-1][:8]])
    wb = Workbook()
    l2 = c["name"].replace(core.wa_fy(c["fy"]) + " ", "")
    cover(wb, core.wa_fy(fy) + ("（出来高）" if kikan_fy else ""), l2,
          f"{d_wa_long(dt.date(fy, 4, 1))} 〜 {d_wa_long(dt.date(fy + 1, 3, 31))}" if not kikan_fy else "令和4年4月28日 〜 令和8年3月31日",
          d_wa_long(dt.date(fy + 1, 3, 25))[:-3] + "月")
    wins = [w for w in W.windows.values() if w["contract"] == c["id"] and w["fy"] == fy]
    wins.sort(key=lambda w: W.window_dates(w)[0])
    tasks = [t for w in wins for t in w["tasks"]]
    works = [w for t in tasks for w in t["works"]]
    top = [f"件名：{c['name']}　　委託者：【架空】みなと市 環境局 施設課　　受託者：【架空】汐見重工環境株式会社（{'監理技術者 m4' if c['kind'] == 'KIKAN' else '現場代理人 ' + core.agent(fy)}）",
           "実施期間：" + " ／ ".join(f"{w['label']} {d_wa(W.window_dates(w)[0])}〜{d_wa(W.window_dates(w)[1])}" for w in wins),
           "工程No. は各停止期間の工程表による。"]
    def rf_sum(t):
        tg = sorted({w["tag"] for w in t["works"]})
        main = "、".join(tg[:3]) + (f" ほか{len(tg) - 3}" if len(tg) > 3 else "")
        res = "良" if not t.get("episodes") else t["episodes"][0][4][:30]
        nr = sum(1 for w in t["works"] if w["kind"] in ("O", "R", "U"))
        act = f"整備・交換 {nr} 台" if nr else "—"
        dd = f"{t['days'][0].month}/{t['days'][0].day}〜{t['days'][-1].month}/{t['days'][-1].day}"
        vals = [t["no"], dd, ({"1": "1号炉", "2": "2号炉", "3": "3号炉", "Z": "共通", "T": "タービン", "C": "共通"}.get(W.windows[t["window"]]["unit"], "共通")), main or "—", t["name"], len(t["works"]), res, act]
        return vals, dict(no=t["no"], date=dd, codes=tg, name=t["name"], 台数=len(t["works"]), 結果=res, 処置=act), t["work_ids"], [t["no"]]
    form_table(wb, "報告サマリー", "1．報告サマリー（工程表 No. 順）", ["工程No.", "実施日", "区分", "タグNo.", "作業項目", "台数", "結果・判定", "処置"],
               [8, 11, 7, 26, 34, 6, 30, 16], tasks, S, rf_sum, top=top, row_h=22)
    def rf_rec(w):
        e = W.by_tag[w["tag"]]
        res, act = texts.result(w)
        t = next(t for t in tasks if w in t["works"])
        vals = [t["no"], UNAME[e["unit"]], w["tag"], e["name"], texts.KIND_JA[w["kind"]], f"{D(w['start']).month}/{D(w['start']).day}", res, act]
        return vals, dict(no=t["no"], 号炉=UNAME[e["unit"]], code=w["tag"], name=e["name"], work=texts.KIND_JA[w["kind"]], date=w["start"],
                          結果=res, 処置=act), [w["id"]], [w["tag"]]
    form_table(wb, "点検整備記録", "2．点検整備記録（機器別）", ["工程No.", "号炉", "タグNo.", "機器名称", "作業", "実施日", "結果", "処置"],
               [8, 7, 13, 30, 10, 8, 26, 18], sorted(works, key=lambda w: (w["start"], w["tag"])), S, rf_rec)
    mr = material_rows(works, fy)
    def rf_mat(m):
        return list(m), dict(品名=m[0], 仕様=m[1], 数量=m[2], 使用箇所=m[3], 手配=m[4], 予備=m[5]), None, [m[0]]
    form_table(wb, "使用材料", "3．使用材料・支給品の使用状況", ["品名", "仕様", "数量", "使用箇所", "手配", "予備（引渡し）"],
               [20, 24, 12, 38, 18, 14], mr, S, rf_mat, foot=["※ 金額は記載しない（契約図書による）。"])
    obs = observations(dict(c, id=c["id"]))
    def rf_obs(o):
        i = obs.index(o) + 1
        return [i, o], dict(no=i, 所見=o), None, [o[:10]]
    form_table(wb, "所見", "4．所見及び次年度以降への提言", ["No.", "所見・提言"], [5, 120], obs, S, rf_obs, row_h=30, foot=["以上"])
    props(wb, f"【架空】{c['name']} 報告書" + (f"（{core.wa_fy(fy)} 出来高）" if kikan_fy else ""), "r2", "工事報告書",
          dt.datetime(2017, 6, 1, 9, 0), modified_by="r2")
    src = os.path.join(WORK, stem + ".xlsx")
    wb.save(src)
    pdf = RD.to_pdf(src, WORK)
    os.replace(pdf, os.path.join(OUT, fn))
    return fn


def main():
    W = core.get()
    out = []
    for c in sorted(W.contracts.values(), key=lambda c: (c["fy"], c["id"])):
        if c["kind"] == "LOCAL":
            continue
        if c["kind"] == "KIKAN":
            for fy in (2022, 2023, 2024, 2025):
                cc = dict(c, _rfy=fy)
                out.append(build_r2(cc, kikan_fy=fy))
            continue
        if c["fy"] >= 2026:
            continue  # 契約期間中（報告書はまだない）
        if not [w for w in W.windows.values() if w["contract"] == c["id"]]:
            continue
        out.append(build_r1(c) if c["fy"] <= 2016 else build_r2(c))
    return out


if __name__ == "__main__":
    import time
    t0 = time.time()
    print(len(main()), time.time() - t0)
