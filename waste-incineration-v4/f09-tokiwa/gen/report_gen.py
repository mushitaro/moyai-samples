# -*- coding: utf-8 -*-
"""工事報告書（文字のある PDF）
 sv1 様式（2015）：Word で書いた文章型 → LibreOffice で PDF
 sv2 様式（2018～、見本は 2025）：1 枚目サマリー表＋詳細、reportlab
 メーカー様式（基幹改良 R7）：出来形＋写真帳、reportlab
"""
import os, subprocess, datetime as dt
import world as W, works as K
from xlh import wareki_date

# ============================================================== sv1 2015（docx → PDF）
def sv1_2015(pdf_path, workdir):
    from docx import Document
    from docx.shared import Pt, Mm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Mm(210), Mm(297)
    sec.left_margin = sec.right_margin = Mm(22); sec.top_margin = sec.bottom_margin = Mm(20)
    st = doc.styles["Normal"]; st.font.name = "IPAGothic"; st.font.size = Pt(10.5)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), "IPAGothic")
    for s in ("Heading 1", "Heading 2", "Title"):
        h = doc.styles[s]; h.font.name = "IPAGothic"; h.element.rPr.rFonts.set(qn("w:eastAsia"), "IPAGothic")
        h.font.color.rgb = None
    doc.core_properties.title = "【架空】ときわ北部クリーンセンター 2号炉定期整備工事 完了報告書（平成27年度）"
    doc.core_properties.author = "sv1（【架空】ときわ北部エコサービス 保全グループ）"
    doc.core_properties.created = dt.datetime(2015, 11, 20, 10, 0)
    def para(t, size=10.5, bold=False, align=None, space=4):
        p = doc.add_paragraph(); r = p.add_run(t); r.font.size = Pt(size); r.bold = bold
        if align: p.alignment = align
        p.paragraph_format.space_after = Pt(space)
        return p
    def fixw(t, mms):
        t.autofit = False
        tblPr = t._tbl.tblPr
        lay = OxmlElement("w:tblLayout"); lay.set(qn("w:type"), "fixed"); tblPr.append(lay)
        grid = t._tbl.tblGrid
        for gc, w in zip(grid.findall(qn("w:gridCol")), mms):
            gc.set(qn("w:w"), str(int(w * 56.7)))
        for row in t.rows:
            for c, w in zip(row.cells, mms):
                c.width = Mm(w)
    def h1(t):
        p = para(t, 12, True, space=6)
        return p
    # 表紙
    for _ in range(5): para("")
    para("【架空】", 12, False, WD_ALIGN_PARAGRAPH.CENTER)
    para("ときわ北部クリーンセンター", 18, True, WD_ALIGN_PARAGRAPH.CENTER)
    para("2号炉定期整備工事", 22, True, WD_ALIGN_PARAGRAPH.CENTER, 12)
    para("完　了　報　告　書", 22, True, WD_ALIGN_PARAGRAPH.CENTER, 60)
    for _ in range(6): para("")
    para("平成27年11月", 12, False, WD_ALIGN_PARAGRAPH.CENTER)
    para("【架空】ときわ北部エコサービス株式会社", 12, False, WD_ALIGN_PARAGRAPH.CENTER)
    para("保全グループ", 12, False, WD_ALIGN_PARAGRAPH.CENTER)
    doc.add_page_break()
    para("【架空】ときわ北部環境組合　施設課　御中", 10.5, False, space=10)
    h1("1．工事概要")
    rows = [("工事件名", "平成27年度 ときわ北部クリーンセンター 2号炉定期整備工事"),
            ("工　　期", "平成27年10月5日（炉停止）〜平成27年10月30日（定格運転確認）"),
            ("施　　工", "【架空】北嶺プラント工業株式会社（職長 x1）／耐火物：【架空】丸越築炉株式会社"),
            ("監　　督", "【架空】ときわ北部エコサービス株式会社 保全グループ sv1"),
            ("対象範囲", "2号炉 ガス化炉、旋回溶融炉、ボイラ、排ガス処理設備、通風設備、不燃物・スラグ設備")]
    t = doc.add_table(rows=len(rows), cols=2); t.style = "Table Grid"
    fixw(t, [28, 138])
    for i, (a, b) in enumerate(rows):
        t.cell(i, 0).text = a; t.cell(i, 1).text = b
    para("")
    h1("2．工程")
    para("工程は当初計画のとおり 10 月 5 日に炉を停止し、3 日間の冷却ののち 10 月 8 日から開放・点検整備に着手した。"
         "ボイラの肉厚測定で過熱器に減肉が見つかり、肉盛補修を追加したため、ボイラ閉止と水圧試験が当初の 10 月 19 日から 10 月 21 日にずれた。"
         "復旧作業と並行して行い、10 月 26 日の点火、10 月 30 日の定格運転確認は予定どおりとなった。")
    h1("3．実施内容")
    texts = [
        ("(1) ガス化炉", "流動砂約 6m³ を抜き出して炉内を清掃し、散気ノズル 96 本の摩耗量を測定した。交換基準を超えた 24 本を新品に交換し、取付部を溶接した。"
                       "炉内耐火物に有害な損傷はなかった。閉止後、珪砂 6 号を約 6m³ 投入した。"),
        ("(2) 旋回溶融炉", "炉内点検の結果、出滓口まわり（約 1.4m²）と炉頂部（約 0.8m²）の耐火物に溶損があり、はつりのうえ出滓口をれんが積み、炉頂部をキャスタブルで補修した。"
                         "炉頂部のスタッド 18 本を植え替えた。鉄皮に異常はなかった。側壁は今回補修範囲外としたが、下段の減りが進んでおり次回以降の課題とする。"),
        ("(3) ボイラ", "水管・過熱器・エコノマイザのダストを除去し、肉厚を測定した（水管 120 点、過熱器 64 点、エコノマイザ 48 点）。"
                      "二次過熱器下部ベンド No.14 で 2.6mm（新製時 4.0mm）を確認したため、組合と協議のうえ肉盛溶接を追加で行い、PT・UT で欠陥のないことを確認した。"
                      "10 月 14 日に定期事業者検査（内部検査）を行い、10 月 21 日に水圧試験（4.6MPa、30 分保持）で漏れのないことを確認した。"),
        ("(4) スートブロワ", "長抜差式 LR-3、LR-4 を取り外して分解し、ランス管とグランドパッキンを交換した。"),
        ("(5) 不燃物抜出装置A", "スクリューを引き抜いて摩耗を測定し（外径 φ350 → φ331）、工場で羽根の硬化肉盛を行った。軸受・軸封を交換して組み込み、芯出しを行った。"),
        ("(6) 誘引通風機", "軸受（負荷側・反負荷側）を交換し、羽根車のダストを除去した。芯出し後の振動は 2.1mm/s であった。"),
        ("(7) 流動空気押込送風機", "軸受・羽根車を点検し、グリスを交換した。異常なし。"),
        ("(8) バグフィルタ", "6 室のろ布を目視点検し、試験用に 3 本を抜き取って予備ろ布を取り付けた。強度試験は外部機関に依頼した（結果は別紙）。"),
        ("(9) 減温塔", "噴霧ノズル 6 本を交換した。"),
        ("(10) 触媒反応塔", "第 1 層から触媒サンプルを採取した（性能試験は外部機関に依頼）。"),
        ("(11) スラグ水砕コンベヤ", "チェーン伸びは 1.1% で交換基準（2%）未満。摩耗したスクレーパ 8 枚を交換した。"),
    ]
    for a, b in texts:
        para(a, 10.5, True, space=2); para(b, 10.5, space=6)
    h1("4．点検結果（肉厚測定 抜粋）")
    rows = [("測定部位", "点数", "新製時(mm)", "最小値(mm)", "判定"),
            ("蒸発管（前壁）", "40", "4.5", "4.1", "良"), ("蒸発管（側壁）", "80", "4.5", "4.0", "良"),
            ("二次過熱器 下部ベンド", "32", "4.0", "2.6（No.14）", "肉盛補修"), ("二次過熱器 直管部", "32", "4.0", "3.4", "良"),
            ("エコノマイザ", "48", "4.0", "3.7", "良")]
    t = doc.add_table(rows=len(rows), cols=5); t.style = "Table Grid"
    fixw(t, [52, 20, 28, 34, 32])
    for i, row in enumerate(rows):
        for j, v in enumerate(row): t.cell(i, j).text = v
    para("")
    h1("5．使用部品・材料")
    rows = [("品名", "仕様", "数量", "使用", "予備品へ"),
            ("ガス化炉 散気ノズル", "SCS13", "28 本", "24 本", "4 本"), ("減温塔 噴霧ノズル", "二流体", "8 個", "6 個", "2 個"),
            ("スートブロワ ランス管", "STBA24 φ60.5", "2 本", "2 本", "－"), ("グランドパッキン", "", "4 組", "2 組", "2 組"),
            ("スクレーパ", "W600 耐摩耗鋼", "10 枚", "8 枚", "2 枚"), ("不燃物抜出装置 軸受", "22220", "2 個", "2 個", "－"),
            ("誘引通風機 軸受", "22324", "2 個", "2 個", "－"), ("ろ布（抜取部）", "PTFE φ160×6000", "3 本", "3 本", "在庫から使用・補充発注済"),
            ("珪砂 6 号", "", "9 t", "約 9 t", "－"), ("溶接棒", "STBA24 用", "5 kg", "約 2 kg", "残り保管"),
            ("耐火れんが・キャスタブル", "丸越築炉 持込", "一式", "一式", "－")]
    t = doc.add_table(rows=len(rows), cols=5); t.style = "Table Grid"
    fixw(t, [46, 36, 20, 20, 44])
    for i, row in enumerate(rows):
        for j, v in enumerate(row): t.cell(i, j).text = v
    tr = t.rows[0]._tr; trPr = tr.get_or_add_trPr(); el = OxmlElement("w:tblHeader"); el.set(qn("w:val"), "true"); trPr.append(el)
    para("")
    h1("6．今後の課題")
    para("・二次過熱器の減肉は今回 1 本であったが、高温側ベンド部で今後も発生が見込まれる。次回の定期事業者検査で再測定し、材質変更を含めて検討する。")
    para("・旋回溶融炉の側壁下段の減りが進んでいる。次回定期整備で残存厚を測定し、部分補修の範囲を決める。")
    para("・不燃物抜出装置は 2 年ごとの分解整備を継続する（B は平成28年度）。")
    para("以上", align=WD_ALIGN_PARAGRAPH.RIGHT)
    # フッタにページ番号
    fp = sec.footer.paragraphs[0]; fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = fp.add_run()
    for tag, txt in (("begin", None), (None, "PAGE"), ("end", None)):
        if tag:
            e = OxmlElement("w:fldChar"); e.set(qn("w:fldCharType"), tag); run._r.append(e)
        else:
            e = OxmlElement("w:instrText"); e.text = txt; run._r.append(e)
    docx_path = os.path.join(workdir, os.path.splitext(os.path.basename(pdf_path))[0] + ".docx")
    doc.save(docx_path)
    subprocess.run(["soffice", "-env:UserInstallation=file:///tmp/lo_F09", "--headless", "--convert-to", "pdf", "--outdir", workdir, docx_path],
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=300)
    os.replace(os.path.join(workdir, os.path.splitext(os.path.basename(docx_path))[0] + ".pdf"), pdf_path)

# ============================================================== reportlab 共通
def _rl():
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    try:
        pdfmetrics.getFont("IPAG")
    except KeyError:
        pdfmetrics.registerFont(TTFont("IPAG", "/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf"))
        pdfmetrics.registerFont(TTFont("IPAPG", "/usr/share/fonts/opentype/ipafont-gothic/ipagp.ttf"))

def sv2_2025(pdf_path):
    _rl()
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.lib import colors
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer, PageBreak
    sN = ParagraphStyle("n", fontName="IPAG", fontSize=9, leading=13)
    sS = ParagraphStyle("s", fontName="IPAG", fontSize=8, leading=11)
    sH = ParagraphStyle("h", fontName="IPAG", fontSize=11, leading=15, spaceBefore=6, spaceAfter=4)
    sT = ParagraphStyle("t", fontName="IPAG", fontSize=15, leading=20, alignment=1)
    def P(t, s=sS): return Paragraph(t, s)
    story = []
    story.append(P("【架空】ときわ北部クリーンセンター", sN))
    story.append(P("2025年度　1号炉定期整備工事　工事報告書", sT))
    story.append(Spacer(1, 4 * mm))
    head = [[P("工事件名"), P("2025年度 ときわ北部クリーンセンター 1号炉定期整備工事"), P("報告日"), P("2025/6/20")],
            [P("工期"), P("2025/5/12（停止）〜 6/6（定格確認）"), P("作成"), P("sv2（保全グループ）")],
            [P("施工"), P("【架空】北嶺プラント工業（職長 x1）、耐火物 【架空】丸越築炉"), P("提出先"), P("【架空】ときわ北部環境組合 施設課")],
            [P("契約"), P("第2期運営委託 年度整備（上限内）"), P("確認"), P("所長")]]
    t = Table(head, colWidths=[18 * mm, 92 * mm, 16 * mm, 50 * mm])
    t.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.5, colors.black), ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#DDEBF7")),
                           ("BACKGROUND", (2, 0), (2, -1), colors.HexColor("#DDEBF7")), ("VALIGN", (0, 0), (-1, -1), "MIDDLE")]))
    story += [t, Spacer(1, 10 * mm)]
    story.append(P("【架空】ときわ北部環境組合　施設課　御中", sN))
    story.append(Spacer(1, 4 * mm))
    story.append(P("標記工事について、下記のとおり完了しましたので報告します。", sN))
    story.append(Spacer(1, 4 * mm))
    story.append(P("記", ParagraphStyle("c", fontName="IPAG", fontSize=10, alignment=1)))
    kv = [["1", "工事件名", "2025年度 ときわ北部クリーンセンター 1号炉定期整備工事"],
          ["2", "工　　期", "2025年5月12日（1号炉停止）〜 2025年6月6日（定格運転確認）"],
          ["3", "実施範囲", "1号炉 ガス化炉・旋回溶融炉・ボイラ・バグフィルタ・不燃物・給じん・脱硝・スラグ設備、共通 前処理破砕機1号"],
          ["4", "結　　果", "全項目完了。工程変更 1 件（BF ケージ交換により 1 日延長、全体工程に影響なし）"],
          ["5", "添　　付", "(1) 実施結果サマリー　(2) 詳細　(3) 使用部品・材料　(4) 肉厚測定記録"]]
    t = Table([[P(a), P(b), P(c)] for a, b, c in kv], colWidths=[8 * mm, 22 * mm, 140 * mm])
    t.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LINEBELOW", (0, 0), (-1, -1), 0.3, colors.grey)]))
    story += [t, Spacer(1, 16 * mm)]
    st = Table([[P("所長"), P("保全G"), P("作成")], ["", "", ""], [P("（印）"), P("（印）"), P("sv2")]],
               colWidths=[22 * mm] * 3, rowHeights=[6 * mm, 14 * mm, 6 * mm], hAlign="RIGHT")
    st.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.5, colors.black), ("ALIGN", (0, 0), (-1, -1), "CENTER")]))
    story += [st, PageBreak(), P("■ 添付(1) 実施結果サマリー", sH)]
    summ = [["No.", "機器（タグ）", "区分", "実施内容", "結果・判定", "次回"],
            ["1", "1号 ガス化炉本体（1-210-01）", "点検", "開放、散気ノズル摩耗測定、耐火物点検", "異常なし", "R8 定修"],
            ["2", "1号 旋回溶融炉本体（1-310-01）", "点検・補修", "残存厚測定、出滓口れんが部分積替え 0.4m²", "側壁下段 最小 118mm", "R8 基幹改良で全面更新"],
            ["3", "1号 ボイラ本体／過熱器／エコノマイザ", "点検", "肉厚測定、定期事業者検査（5/21）、水圧試験", "過熱器 No.9 最小 2.5mm（必要厚 2.1mm）", "R8 基幹改良で過熱器更新"],
            ["4", "1号 バグフィルタ本体（1-520-01）", "整備", "ろ布 252 本 全数交換、ケージ 3 本交換", "差圧 1.9→0.9kPa", "ろ布 2031 年度目安"],
            ["5", "1号 不燃物抜出装置B（1-220-02）", "整備", "分解、軸受・軸封交換、羽根先端肉盛", "外径 φ336（肉盛後 φ349）", "2027 年度"],
            ["6", "1号 定量供給機A（1-110-04）", "整備", "減速機点検、駆動チェーン交換", "良", "2028 年度"],
            ["7", "1号 アンモニア水注入ポンプA（1-540-04）", "整備", "ダイヤフラム・弁座・弁体交換", "良", "2028 年度"],
            ["8", "1号 スラグ水砕コンベヤ（1-320-02）", "点検", "チェーン伸び計測", "1.8%（基準 2%）", "R8 停止中に交換"],
            ["9", "1号 誘引通風機（1-620-01）", "点検", "軸受点検、振動測定", "4.2mm/s（管理値 7.1）", "R8 基幹改良で更新"],
            ["10", "前処理破砕機 1号（0-102-01）", "整備", "カッタ刃 40 枚・軸受 4 個交換", "良", "2028 年度"]]
    summ = [[P(c) for c in row] for row in summ]
    t = Table(summ, colWidths=[10 * mm, 42 * mm, 13 * mm, 49 * mm, 35 * mm, 27 * mm], repeatRows=1)
    t.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.5, colors.black), ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E79")),
                           ("TEXTCOLOR", (0, 0), (-1, 0), colors.white), ("VALIGN", (0, 0), (-1, -1), "MIDDLE")]))
    story += [t, Spacer(1, 3 * mm)]
    story.append(P("工程：BF 第4室でケージ曲がり 3 本が見つかり交換したため、BF ろ布交換の完了が 5/26 → 5/27 となった。全体工程（6/2 点火・6/6 定格）への影響はなし。", sN))
    story.append(PageBreak())
    story.append(P("■ 添付(2) 詳細", sH))
    det = [
        ("1. ガス化炉", "流動砂を抜き出して炉内を清掃し、散気ノズルの摩耗を測定した。交換基準を超えるものはなく、耐火物にも異常はなかった。閉止後に珪砂 6 号 9t を投入した。"),
        ("2. 旋回溶融炉", "側壁下段の残存厚は最小 118mm で、2023 年度の 2号炉（95mm）より約 2 年分厚い。令和8年度の基幹改良で全面更新されるため、今回は出滓口の損傷れんがを 0.4m² 積み替える最小補修とした。"),
        ("3. ボイラ", "定期事業者検査（内部検査）を 5/21 に行い、5/23 の水圧試験で漏れのないことを確認した。過熱器 No.9 ベンドで 2.5mm を測定。計算必要厚 2.1mm を上回るため継続使用とし、基幹改良での更新まで毎年測定する。"),
        ("4. バグフィルタ", "2018 年度交換のろ布 252 本を全数交換した。第4室で旧ろ布の抜取り時にケージの曲がり 3 本を確認し、予備品のケージに交換した（予備品は補充発注済）。交換前の差圧 1.9kPa は立上げ後 0.9kPa となった。"),
        ("5. 不燃物抜出装置B", "スクリュー外径は φ336 で、羽根先端に硬化肉盛を行い φ349 に戻した。軸受・軸封を交換した。"),
        ("6. 前処理破砕機 1号", "2号機（予備）に切り替えて受入を続けながら、カッタ刃 40 枚と軸受 4 個を交換した。"),
    ]
    for a, b in det:
        story += [P(a, sN), P(b, sN), Spacer(1, 2 * mm)]
    story.append(P("■ 添付(3) 使用部品・材料", sH))
    parts = [["品名", "仕様", "手配数", "使用", "予備・在庫", "発注"],
             ["ろ布", "PTFE φ160×6000", "264 本", "252 本", "12 本（予備品）", "PO25-0018"],
             ["リテーナ（ケージ）", "SUS304 φ155×6000", "12 本", "3 本", "9 本＋補充 3 本", "PO25-0018 / 0041"],
             ["不燃物抜出装置 軸受", "22220 CAKE4", "2 個", "2 個", "－", "PO25-0021"],
             ["軸封パッキン", "グランド 一式", "1 式", "1 式", "－", "PO25-0021"],
             ["駆動チェーン", "RS100 142L", "1 本", "1 本", "－", "PO25-0021"],
             ["ダイヤフラム・弁座キット", "30L/h 用", "2 組", "1 組", "1 組", "PO25-0022"],
             ["破砕機カッタ刃", "t=50", "40 枚", "40 枚", "－", "PO25-0024"],
             ["破砕機 軸受", "23228", "4 個", "4 個", "－", "PO25-0024"],
             ["硬化肉盛溶接棒", "HF-600 φ4.0", "10 kg", "6 kg", "4 kg", "PO25-0025"],
             ["珪砂 6 号", "フレコン 1.5t", "6 袋", "6 袋", "－", "PO25-0028"],
             ["耐火れんが・キャスタブル", "丸越築炉 持込", "一式", "一式", "－", "PO25-0027"]]
    parts = [[P(c) for c in row] for row in parts]
    t = Table(parts, colWidths=[36 * mm, 34 * mm, 16 * mm, 16 * mm, 32 * mm, 30 * mm], repeatRows=1)
    t.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.5, colors.black), ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#DDEBF7"))]))
    story += [t, Spacer(1, 3 * mm)]
    story.append(P("■ 添付(4) 肉厚測定記録（過熱器 抜粋、単位 mm）", sH))
    rows = [["管No.", "部位", "新製時", "2017", "2021", "2025", "判定"]]
    import random
    rr = random.Random(2025)
    for n in range(1, 25):
        v25 = 2.5 if n == 9 else round(rr.uniform(2.8, 3.6), 1)
        v21 = round(v25 + rr.uniform(0.1, 0.3), 1); v17 = round(v21 + rr.uniform(0.2, 0.5), 1)
        rows.append([str(n), "下部ベンド" if n % 2 else "直管部", "4.0", f"{min(v17, 4.0):.1f}", f"{v21:.1f}", f"{v25:.1f}", "要監視" if v25 < 2.8 else "良"])
    t = Table([[P(c) for c in r] for r in rows], colWidths=[14 * mm, 26 * mm, 18 * mm, 18 * mm, 18 * mm, 18 * mm, 20 * mm], repeatRows=1)
    t.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.5, colors.black), ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#DDEBF7"))]))
    story.append(t)
    def onpage(c, d):
        c.setFont("IPAG", 7.5)
        c.drawString(18 * mm, 10 * mm, "【架空】ときわ北部エコサービス株式会社 保全グループ　様式 MR-02（2018改）")
        c.drawRightString(192 * mm, 10 * mm, f"{d.page}")
    doc = SimpleDocTemplate(pdf_path, pagesize=A4, leftMargin=16 * mm, rightMargin=16 * mm, topMargin=14 * mm, bottomMargin=16 * mm,
                            title="【架空】ときわ北部クリーンセンター 2025年度 1号炉定期整備工事 工事報告書", author="sv2（【架空】ときわ北部エコサービス 保全グループ）")
    doc.build(story, onFirstPage=onpage, onLaterPages=onpage)

# ============================================================== メーカー（基幹改良 R7 施工分）
PHOTOS = [
    ("R7.9.10", "溶融炉", "2号 旋回溶融炉 上段側壁", "耐火物解体 状況", "炉"),
    ("R7.9.26", "溶融炉", "2号 旋回溶融炉 炉内", "解体完了 鉄皮露出", "炉"),
    ("R7.10.1", "溶融炉", "2号 旋回溶融炉 鉄皮 下段", "鉄皮 減肉 範囲確認（想定より広い）", "炉"),
    ("R7.10.7", "溶融炉", "2号 旋回溶融炉 鉄皮 下段", "当て板 溶接 完了", "炉"),
    ("R7.11.19", "溶融炉", "2号 旋回溶融炉 炉内", "築炉 完了（高クロミア質）", "炉"),
    ("R7.10.2", "ボイラ", "2号 過熱器", "既設過熱器 撤去 完了", "管"),
    ("R7.10.21", "ボイラ", "過熱器管材", "工場 製作状況（納期遅れ 3 週）", "管"),
    ("R7.12.12", "ボイラ", "2号 過熱器", "現地溶接・RT 状況", "管"),
    ("R7.12.22", "ボイラ", "2号 ボイラ", "水圧試験（6.0MPa）組合立会", "管"),
    ("R7.11.14", "通風", "2号 誘引通風機", "新設 誘引通風機 据付・芯出し", "機"),
    ("R7.11.18", "電気計装", "2号 誘引通風機インバータ盤", "新盤 据付 完了", "盤"),
    ("R8.1.15", "試運転", "2号炉 中央制御室", "定格負荷 確認", "盤"),
]

def maker_k3(pdf_path):
    _rl()
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.pdfgen import canvas
    c = canvas.Canvas(pdf_path, pagesize=A4)
    c.setTitle("【架空】ときわ北部クリーンセンター基幹的設備改良工事 令和7年度 施工報告書（出来形・写真帳）")
    c.setAuthor("m1（【架空】鷹羽エンジニアリング 工事部）")
    Wp, Hp = A4
    def frame(page, total):
        c.setLineWidth(0.8); c.rect(12 * mm, 12 * mm, Wp - 24 * mm, Hp - 24 * mm)
        c.setFont("IPAG", 7.5)
        c.drawString(14 * mm, 7 * mm, "【架空】鷹羽エンジニアリング株式会社　ときわ北部基幹改良 現場事務所")
        c.drawRightString(Wp - 14 * mm, 7 * mm, f"{page} / {total}")
    total = 2 + (len(PHOTOS) + 2) // 3
    # p1 表紙＋概要
    c.setFont("IPAG", 10); c.drawString(20 * mm, Hp - 30 * mm, "【架空】")
    c.setFont("IPAG", 16); c.drawCentredString(Wp / 2, Hp - 45 * mm, "ときわ北部クリーンセンター基幹的設備改良工事（令和7〜9年度）")
    c.setFont("IPAG", 20); c.drawCentredString(Wp / 2, Hp - 60 * mm, "令和7年度　施工報告書")
    c.setFont("IPAG", 12); c.drawCentredString(Wp / 2, Hp - 70 * mm, "（2号炉施工分　出来形・写真帳）")
    y = Hp - 95 * mm
    rows = [("発注者", "【架空】ときわ北部環境組合"), ("受注者", "【架空】鷹羽エンジニアリング株式会社　現場代理人 m1"),
            ("契約工期", "令和7年4月18日〜令和10年3月31日（継続費 3 か年）"),
            ("令和7年度施工", "2号炉　令和7年9月1日（停止）〜令和8年1月16日（引渡し）"),
            ("変更", "第1回変更（R7.12.5）：2号炉施工 12/19→1/16 に延長（過熱器管材の納期遅れ）"),
            ("提出日", "令和8年2月6日")]
    c.setFont("IPAG", 9.5)
    for a, b in rows:
        c.rect(25 * mm, y - 2.5 * mm, 30 * mm, 8 * mm); c.rect(55 * mm, y - 2.5 * mm, 130 * mm, 8 * mm)
        c.drawString(27 * mm, y, a); c.drawString(57 * mm, y, b[:60])
        y -= 8 * mm
    y -= 6 * mm
    c.setFont("IPAG", 11); c.drawString(25 * mm, y, "1．出来形総括（令和7年度分）"); y -= 8 * mm
    hdr = ("工種", "内容", "対象", "数量", "出来形")
    data = [("溶融炉", "旋回溶融炉 耐火物全面更新（高クロミア質）", "2-310-01", "一式（約 42m²）", "100%"),
            ("溶融炉", "鉄皮補修（当て板 追加 2.6m²）", "2-310-01", "一式", "100%"),
            ("溶融炉", "二次燃焼室 耐火物更新", "2-310-02", "一式", "100%"),
            ("溶融炉", "溶融炉バーナ 更新", "2-310-03", "1 基", "100%"),
            ("溶融炉", "炉体冷却水ポンプ A・B 更新", "2-310-06/07", "2 台", "100%"),
            ("ボイラ", "過熱器 一式更新（SUS310S 系）", "2-410-02", "一式", "100%"),
            ("ボイラ", "エコノマイザ 管群更新", "2-410-03", "一式", "100%"),
            ("通風", "誘引通風機 更新（高効率型）", "2-620-01", "1 台", "100%"),
            ("通風", "押込送風機・燃焼用送風機 更新", "2-610-01/02", "2 台", "100%"),
            ("電気計装", "インバータ盤 更新（誘引・押込）", "2-810-17/18", "2 面", "100%"),
            ("電気計装", "炉コントローラ盤 制御改造", "2-850-01/02", "一式", "100%")]
    xs = [25, 45, 120, 145, 170, 185]
    c.setFont("IPAG", 8)
    def row(vals, yy, bold=False):
        for i, v in enumerate(vals):
            c.rect(xs[i] * mm, yy - 2 * mm, (xs[i + 1] - xs[i]) * mm, 6.5 * mm)
            c.drawString(xs[i] * mm + 1.5 * mm, yy, v)
    row(hdr, y); y -= 6.5 * mm
    for v in data:
        row(v, y); y -= 6.5 * mm
    c.setFont("IPAG", 8)
    c.drawString(25 * mm, y - 4 * mm, "※ 出来高・金額は別途 出来高報告書による。本書には記載しない。")
    frame(1, total); c.showPage()
    # p2 経過と変更
    c.setFont("IPAG", 11); c.drawString(20 * mm, Hp - 25 * mm, "2．施工の経過")
    lines = [
        "令和7年9月1日に2号炉を停止し、9月8日から溶融炉の耐火物解体と誘引通風機の撤去に着手した。",
        "溶融炉は解体後の鉄皮肉厚測定で、下段の減肉範囲が事前調査より広いことが分かり（10月1日）、組合と協議のうえ",
        "当て板を 2.6m² 追加した。このため鉄皮補修が 4 日延び、築炉の着手は 10月9日、完了は 11月19日となった。",
        "ボイラは 10月4日までに既設の過熱器・エコノマイザを撤去した。10月21日、過熱器管材（SUS310S 系）の製作工場から",
        "約 3 週間の納期遅れの連絡があり、組合 u1 へ報告した。管材は 11月7日に搬入、11月10日から据付・溶接を始め、",
        "12月20日に完了した。水圧試験は 12月22日（6.0MPa、組合立会）、使用前自主検査は 12月23日に行った。",
        "この遅れにより、炉乾燥を 12月24日〜1月6日（年末年始の休工を含む）、総合試運転を 1月7日〜1月16日に変更した",
        "（第1回変更契約、令和7年12月5日）。誘引通風機・送風機・インバータ盤・制御改造は予定どおり完了した。",
        "",
        "3．試運転結果（抜粋）",
        "　定格負荷（95t/日相当）で 72 時間の連続運転を確認した。溶融炉出口温度 1,310〜1,345℃、スラグ流下は良好。",
        "　誘引通風機の電流は更新前と同じ風量で約 18% 低下した。煙突出口の排ガス濃度は計画値以下であった。",
        "",
        "4．令和8年度への申し送り",
        "　・1号炉の施工（令和8年9月〜）では、鉄皮の減肉範囲を停止前に超音波で再調査する。",
        "　・過熱器管材は令和8年度分を 2 か月前倒しで手配済み。",
    ]
    c.setFont("IPAG", 9)
    y = Hp - 35 * mm
    for ln in lines:
        if ln and ln[0].isdigit():
            c.setFont("IPAG", 11); c.drawString(20 * mm, y, ln); c.setFont("IPAG", 9)
        else:
            c.drawString(22 * mm, y, ln)
        y -= 6 * mm
    frame(2, total); c.showPage()
    # 写真帳
    page = 3
    for i in range(0, len(PHOTOS), 3):
        c.setFont("IPAG", 11); c.drawString(20 * mm, Hp - 22 * mm, "5．写真帳")
        for k, ph in enumerate(PHOTOS[i:i + 3]):
            n = i + k + 1
            top = Hp - 30 * mm - k * 82 * mm
            x0, w, h = 20 * mm, 100 * mm, 72 * mm
            c.setFillGray(0.85); c.rect(x0, top - h, w, h, fill=1, stroke=1); c.setFillGray(0)
            # 写真の代わりの簡単な形（現場写真の位置）
            c.setStrokeGray(0.5)
            kind = ph[4]
            if kind == "炉":
                c.circle(x0 + w / 2, top - h / 2, 25 * mm)
                c.circle(x0 + w / 2, top - h / 2, 18 * mm)
            elif kind == "管":
                for j in range(9): c.line(x0 + 15 * mm + j * 8 * mm, top - 10 * mm, x0 + 15 * mm + j * 8 * mm, top - h + 10 * mm)
            elif kind == "機":
                c.circle(x0 + w / 2 - 10 * mm, top - h / 2, 20 * mm); c.rect(x0 + w / 2 + 12 * mm, top - h / 2 - 10 * mm, 25 * mm, 20 * mm)
            else:
                c.rect(x0 + 25 * mm, top - h + 10 * mm, 50 * mm, h - 20 * mm)
            c.setStrokeGray(0)
            c.setFont("IPAG", 8); c.drawString(x0 + 2 * mm, top - 5 * mm, f"写真 {n}")
            # 説明欄
            bx = 124 * mm
            items = [("写真No.", str(n)), ("撮影日", ph[0]), ("工種", ph[1]), ("箇所", ph[2]), ("内容", ph[3])]
            yy = top - 6 * mm
            for a, b in items:
                c.rect(bx, yy - 3 * mm, 16 * mm, 11 * mm); c.rect(bx + 16 * mm, yy - 3 * mm, 50 * mm, 11 * mm)
                c.setFont("IPAG", 8); c.drawString(bx + 1.5 * mm, yy + 1 * mm, a)
                txt = b
                if len(txt) > 15:
                    c.drawString(bx + 17.5 * mm, yy + 3 * mm, txt[:15]); c.drawString(bx + 17.5 * mm, yy - 1 * mm, txt[15:])
                else:
                    c.drawString(bx + 17.5 * mm, yy + 1 * mm, txt)
                yy -= 11 * mm
        frame(page, total); c.showPage(); page += 1
    c.save()
