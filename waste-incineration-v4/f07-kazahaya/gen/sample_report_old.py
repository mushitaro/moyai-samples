# -*- coding: utf-8 -*-
"""工事報告書（旧）：メーカー様式（水樹プラント MP-R-03）で水樹テクノサービスが作り、sp1 が受け取った完成報告書。
Word（python-docx）で作り、LibreOffice で文字のある PDF にする（2015–2020 年度）。"""
import os
from docx import Document
from docx.shared import Pt, Mm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from contracts import *
from common import *

GOTHIC = "IPAGothic"


def set_font(run, sz=10, b=False, color=None):
    run.font.name = GOTHIC
    run._element.rPr.rFonts.set(qn("w:eastAsia"), GOTHIC)
    run.font.size = Pt(sz)
    run.bold = b
    if color:
        run.font.color.rgb = RGBColor.from_string(color)


def para(doc_or_cell, text, sz=10, b=False, align=None, space_after=2, color=None):
    p = doc_or_cell.add_paragraph()
    r = p.add_run(text)
    set_font(r, sz, b, color)
    if align == "c":
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == "r":
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(0)
    return p


def cell_text(cell, text, sz=9, b=False, align=None):
    cell.text = ""
    p = cell.paragraphs[0]
    r = p.add_run(str(text))
    set_font(r, sz, b)
    if align == "c":
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(0)


def shade(cell, hexrgb):
    tcPr = cell._element.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hexrgb)
    tcPr.append(shd)


def repeat_header(row):
    trPr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:tblHeader")
    el.set(qn("w:val"), "true")
    trPr.append(el)


def table(doc, head, rows, widths, sz=8.5, head_fill="D9D9D9"):
    t = doc.add_table(rows=1, cols=len(head))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, hd in enumerate(head):
        cell_text(t.rows[0].cells[i], hd, sz, True, "c")
        shade(t.rows[0].cells[i], head_fill)
    repeat_header(t.rows[0])
    for row in rows:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cell_text(cells[i], v, sz, False, "c" if (isinstance(v, (int, float)) or len(str(v)) <= 4) else None)
    set_widths(t, widths)
    return t


def set_widths(t, widths):
    t.autofit = False
    grid = t._tbl.tblGrid
    for i, gc in enumerate(grid.findall(qn("w:gridCol"))):
        if i < len(widths):
            gc.set(qn("w:w"), str(int(widths[i] * 56.7)))
    for row in t.rows:
        for i, w in enumerate(widths):
            row.cells[i].width = Mm(w)


def page_break(doc):
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


# 火格子摩耗測定記録（2018-09-06〜08、正解の世界の測定値）
def wear_rows():
    import zlib
    rows = []
    stages = [("乾燥段", 6, 3.2), ("燃焼段", 8, 7.1), ("後燃焼段", 6, 4.0)]
    for st, ncol, base in stages:
        for side in ("可動", "固定"):
            for col in range(1, ncol + 1):
                vals = []
                for pnt in range(1, 6 if st == "燃焼段" else 4):
                    k = zlib.crc32(f"{st}{side}{col}{pnt}".encode()) % 100 / 100
                    v = base * (1.0 if side == "可動" else 0.55) * (0.8 + 0.35 * k)
                    if st == "燃焼段" and side == "可動" and col in (1, 3):
                        v += 1.0
                    vals.append(round(v, 1))
                if st == "燃焼段" and side == "可動" and col == 3:
                    vals[2] = 8.6
                mx = max(vals)
                pts = vals + [""] * (5 - len(vals))
                judge = "要注意" if mx >= 7.5 else "良"
                rows.append([st if (side == "可動" and col == 1) else "", side if col == 1 else "", f"{col}列"] + pts + [mx, judge])
    return rows


def make():
    c = C18
    stem = f"{FN}_工事報告書_2018年度1号炉定期整備工事_メーカー様式"
    docx_path = os.path.join(WORK, stem + ".docx")
    doc = Document()
    cp = doc.core_properties
    cp.title = "【架空】2018年度 かざはやエコプラザ 1号炉定期整備工事 完成報告書"
    cp.author = "【架空】水樹テクノサービス株式会社 q1"
    cp.subject = "【架空】完成報告書（水樹プラント様式 MP-R-03）"
    sec = doc.sections[0]
    sec.page_height, sec.page_width = Mm(297), Mm(210)
    sec.left_margin = sec.right_margin = Mm(20)
    sec.top_margin = Mm(18)
    sec.bottom_margin = Mm(15)
    hp = sec.header.paragraphs[0]
    set_font(hp.add_run("【架空】水樹プラント様式 MP-R-03　　　　　　　　　　　　　　　【架空】かざはやエコプラザ 1号炉定期整備工事"), 7.5)
    fp = sec.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = fp.add_run()
    set_font(run, 8)
    for kind, txt in (("begin", None), (None, "PAGE"), ("end", None)):
        if kind:
            el = OxmlElement("w:fldChar")
            el.set(qn("w:fldCharType"), kind)
            run._r.append(el)
        else:
            el = OxmlElement("w:instrText")
            el.text = txt
            run._r.append(el)
    # ---- 表紙（1 ページ目）
    for _ in range(5):
        para(doc, "")
    para(doc, "【架空】", 12, False, "c")
    para(doc, "2018年度", 16, True, "c")
    para(doc, "かざはやエコプラザ　1号炉定期整備工事", 20, True, "c", 8)
    para(doc, "完　成　報　告　書", 26, True, "c", 30)
    for _ in range(6):
        para(doc, "")
    para(doc, "工　期　　2018年9月3日 ～ 2018年9月28日", 12, False, "c")
    para(doc, "提出日　　2018年10月12日", 12, False, "c", 40)
    para(doc, "提出先　【架空】かざはや環境パートナーズ株式会社　御中", 12, False, "c", 30)
    para(doc, "【架空】水樹テクノサービス株式会社", 14, True, "c")
    para(doc, "（監修　【架空】水樹プラント株式会社 環境プラント部）", 10, False, "c")
    page_break(doc)
    # ---- 目次・概要
    para(doc, "目　次", 13, True, "c", 6)
    for t in ["1. 工事概要", "2. 工程実績", "3. 作業内容および結果", "4. 火格子摩耗測定記録", "5. 使用材料", "6. 所見および提言",
              "7. 写真（写真代替図）"]:
        para(doc, "　　" + t, 10)
    para(doc, "")
    para(doc, "1. 工事概要", 12, True)
    ov = [("工事件名", "【架空】" + c["name"]), ("工事場所", "【架空】かざはやエコプラザ（かざはや市）1号焼却炉 他"),
          ("工期", "2018年9月3日 ～ 2018年9月28日（炉停止 9月1日 6:00、ごみ投入 9月26日 10:20）"),
          ("発注者", "【架空】かざはや環境パートナーズ株式会社（担当 sp1）"),
          ("受注者", "【架空】水樹テクノサービス株式会社（職長 q1）"),
          ("協力会社", "【架空】東雲築炉工業（築炉）、【架空】北斗足場（足場）、【架空】かざはや分析センター（ろ布試験）"),
          ("工事範囲", "1号炉 ストーカ・給じん装置・灰押出装置・耐火物・スートブロワ・減温塔・バグフィルタ・炉別MCC の点検整備（仕様書 KZ-18-01 による）")]
    t = doc.add_table(rows=0, cols=2)
    t.style = "Table Grid"
    for k, v in ov:
        cells = t.add_row().cells
        cell_text(cells[0], k, 9, True, "c")
        shade(cells[0], "EDEDED")
        cell_text(cells[1], v, 9)
    set_widths(t, [30, 140])
    para(doc, "")
    para(doc, "2. 工程実績", 12, True)
    rows = []
    for (no, nm, code, crew, ps, pe, as_, ae) in c["tasks"]:
        rows.append([no, nm, f"{ps.month}/{ps.day}～{pe.month}/{pe.day}", f"{as_.month}/{as_.day}～{ae.month}/{ae.day}",
                     "測定点追加のため1日延長" if no == 5 else ""])
    table(doc, ["No", "作業項目", "予定", "実績", "備考"], rows, [10, 80, 24, 24, 32], 8)
    page_break(doc)
    # ---- 作業内容および結果（メーカーの呼び名）
    para(doc, "3. 作業内容および結果", 12, True)
    work = [
        ["1", "No.1 ストーカ（乾燥段・燃焼段・後燃焼段）", "火格子摩耗測定（可動・固定 全列）", "燃焼段 可動側の摩耗が大きい（最大8.6mm）。4章・6章参照"],
        ["2", "No.1 炉本体 耐火物", "打診・目視点検、部分補修（燃焼段右側壁 0.9m²、炉出口天井れんが 2枚）", "補修後 異常なし"],
        ["3", "No.1 給じんプッシャ", "摺動部ライナ（摺動板）取替、ストローク確認", "良"],
        ["4", "No.1 灰押出機", "プッシャ底板ライナ 2枚 溶接補修", "良"],
        ["5", "No.1 ボイラ 長抜差スートブロワ No.4", "ランス・グランドパッキン交換、手回し・作動確認", "良"],
        ["6", "No.1 減温塔 噴霧ノズル", "ノズルチップ 6個交換", "噴霧パターン確認 良"],
        ["7", "No.1 バグフィルタ", "ろ布抜取 6室×2本 → 強度・通気度試験（分析センター）", "試験結果 2018/10/26 別紙報告（引張強度保持率 88～92%）"],
        ["8", "No.1 動力制御盤 MCC-3", "盤内清掃・端子増締め・絶縁抵抗測定", "良"],
        ["9", "No.1 ボイラ給水ポンプ（B）電動機", "絶縁診断（PI・tanδ）", "良（PI 4.1）"],
        ["10", "No.1 炉 全体", "復旧・単体／連動試運転、立上げ立会", "9/26 ごみ投入、9/27 定格負荷到達"],
    ]
    table(doc, ["No", "機器（呼称）", "作業内容", "結果"], work, [8, 46, 62, 54], 8.5)
    para(doc, "")
    para(doc, "※機器の呼称は当社の完成図書（機器リスト）による。", 8)
    page_break(doc)
    para(doc, "4. 火格子摩耗測定記録（2018年9月6日～8日、デプスゲージ、単位 mm、新品基準面からの摩耗量）", 11, True)
    para(doc, "管理値：摩耗量 15mm で取替（メーカー基準）。測定点：乾燥段・後燃焼段は各列3点、燃焼段は各列5点（9/8 追加）。", 8.5)
    table(doc, ["段", "側", "列", "点1", "点2", "点3", "点4", "点5", "最大", "判定"], wear_rows(),
          [18, 12, 12, 14, 14, 14, 14, 14, 16, 18], 8)
    page_break(doc)
    para(doc, "5. 使用材料", 12, True)
    mats = [
        ["給じん装置 摺動板（ライナ）", "SS400＋硬化肉盛 t25", "4枚", "当社手配", "使用"],
        ["灰押出機 底板ライナ", "SS400 t19", "2枚", "当社手配", "使用"],
        ["キャスタブル（高アルミナ質）", "東雲築炉 標準品", "1.8t", "協力会社手配", "使用 1.6t／残 0.2t は持帰り"],
        ["耐火れんが", "SK34 並形", "2枚", "協力会社手配", "使用"],
        ["スートブロワ グランドパッキン", "黒鉛系", "1式", "当社手配", "使用"],
        ["減温塔 ノズルチップ", "二流体 SUS316", "6個", "SPC支給（予備品）", "使用"],
    ]
    table(doc, ["品名", "仕様", "数量", "手配", "使用区分"], mats, [46, 40, 16, 30, 38], 8.5)
    para(doc, "")
    para(doc, "6. 所見および提言", 12, True)
    for t in [
        "(1) 燃焼段の可動火格子は、運転3年半で最大8.6mm（1列目・3列目）の摩耗があった。乾燥段（最大3.9mm）・後燃焼段（最大5.1mm）に比べ大きい。"
        "摩耗速度は約2.5mm/年で、管理値15mmに達するのは2021～2022年度と推定する。長期修繕計画の更新年（2027年度）の見直しを提言する。",
        "(2) 当面の対策として、次回定修（2019年度）で摩耗の大きい列（1・3列目）の部分交換を行うことを推奨する。",
        "(3) 燃焼段右側壁の耐火物損傷は、火格子上0～300mmの範囲でクリンカ付着によるもの。次回も同部位を重点点検とする。",
        "(4) その他の機器に異常はなかった。",
    ]:
        para(doc, t, 9.5, space_after=4)
    page_break(doc)
    para(doc, "7. 写真（写真代替図）", 12, True)
    para(doc, "※本見本では写真の代わりに【架空】見本用の写真代替図（枠と説明）を置く。", 8)
    tt = doc.add_table(rows=3, cols=2)
    tt.style = "Table Grid"
    caps = ["写真1 燃焼段 可動火格子 3列目（摩耗部）", "写真2 摩耗測定状況（デプスゲージ）", "写真3 燃焼段右側壁 損傷部（補修前）",
            "写真4 同 補修後", "写真5 給じん装置 摺動板 取替後", "写真6 スートブロワ No.4 ランス交換"]
    for i, cap in enumerate(caps):
        cell = tt.rows[i // 2].cells[i % 2]
        cell_text(cell, "【架空】見本用の写真代替図\n\n\n\n\n\n\n" + cap, 9, False, "c")
    set_widths(tt, [85, 85])
    doc.save(docx_path)
    pdf = to_pdf(docx_path, WORK)
    final = outpath(stem + ".pdf")
    os.replace(pdf, final)
    return stem + ".pdf"


if __name__ == "__main__":
    print(make())
