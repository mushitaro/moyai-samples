# -*- coding: utf-8 -*-
"""工事報告書（文字のある PDF）。
旧 = SPC s1 様式（文章型、Word → PDF）。2022 年度は s2 が引継ぎ後もこの様式で作成。
新 = SPC s2 様式（2023〜、1 ページ目にサマリー表、表中心、写真帳）。Word → PDF。旧様式の表紙と第 1 章をコピーして使っている。
市追加工事 = 鷹羽エンジの完成報告書（社内文書システム出力、reportlab で PDF）。
客先提出用なので金額は載せない。"""
import datetime as dt
import os
import subprocess
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt, RGBColor
import contracts as K
from xlhelp import OUT, WORK, LO_PROFILE, render

JP = "IPAGothic"


# ---------------------------------------------------------------- docx 補助
def set_font(run, size=10.5, bold=False, color=None):
    run.font.name = JP
    run.font.size = Pt(size)
    run.font.bold = bold
    rpr = run._element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rpr.append(rf)
    for k in ("w:eastAsia", "w:ascii", "w:hAnsi"):
        rf.set(qn(k), JP)
    if color:
        run.font.color.rgb = RGBColor.from_string(color)


def para(doc, text="", size=10.5, bold=False, align=None, space_after=4, indent=None):
    p = doc.add_paragraph()
    if text:
        set_font(p.add_run(text), size, bold)
    if align == "c":
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == "r":
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.space_after = Pt(space_after)
    if bold and size >= 11:
        p.paragraph_format.keep_with_next = True
    if indent:
        p.paragraph_format.left_indent = Mm(indent)
    return p


def shade(cell, hex_):
    tcPr = cell._element.get_or_add_tcPr()
    sh = OxmlElement("w:shd")
    sh.set(qn("w:val"), "clear")
    sh.set(qn("w:color"), "auto")
    sh.set(qn("w:fill"), hex_)
    tcPr.append(sh)


def repeat_header(row):
    trPr = row._tr.get_or_add_trPr()
    h = OxmlElement("w:tblHeader")
    h.set(qn("w:val"), "true")
    trPr.append(h)


def table(doc, head, rows, widths, size=8.5, head_fill="D9E1F2"):
    t = doc.add_table(rows=1, cols=len(head))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(head):
        c = t.rows[0].cells[i]
        c.text = ""
        set_font(c.paragraphs[0].add_run(h), size, True)
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        shade(c, head_fill)
    repeat_header(t.rows[0])
    for r in rows:
        cells = t.add_row().cells
        for i, v in enumerate(r):
            cells[i].text = ""
            set_font(cells[i].paragraphs[0].add_run("" if v is None else str(v)), size)
    fix_widths(t, widths)
    return t


def fix_widths(t, widths):
    t.autofit = False
    grid = t._tbl.tblGrid
    for i, gc in enumerate(grid.findall(qn("w:gridCol"))):
        gc.set(qn("w:w"), str(int(widths[i] * 56.7)))
    for row in t.rows:
        for i, w in enumerate(widths):
            row.cells[i].width = Mm(w)


def photo_box(doc, caps, h=55):
    t = doc.add_table(rows=2, cols=2)
    t.style = "Table Grid"
    for j in range(2):
        for i in range(2):
            k = j * 2 + i
            c = t.rows[j].cells[i]
            c.width = Mm(82)
            c.text = ""
            p = c.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if k < len(caps):
                set_font(p.add_run("\n\n（写真）\n\n\n"), 9, color="808080")
                p2 = c.add_paragraph()
                p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
                set_font(p2.add_run(caps[k]), 8.5)
        fix_widths(t, [82, 82])
        tr = t.rows[j]._tr.get_or_add_trPr()
        hh = OxmlElement("w:trHeight")
        hh.set(qn("w:val"), str(int(h * 56.7)))
        tr.append(hh)
    return t


def new_doc(title, author):
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Mm(210), Mm(297)
    sec.left_margin = sec.right_margin = Mm(20)
    sec.top_margin = Mm(18)
    sec.bottom_margin = Mm(18)
    doc.core_properties.title = title
    doc.core_properties.author = author
    st = doc.styles["Normal"]
    st.font.name = JP
    st.element.rPr.rFonts.set(qn("w:eastAsia"), JP)
    return doc


def footer(doc, text):
    p = doc.sections[0].footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_font(p.add_run(text + "　－ "), 8)
    r = p.add_run()
    for tag, txt in (("begin", None), (None, "PAGE"), ("end", None)):
        if tag:
            e = OxmlElement("w:fldChar")
            e.set(qn("w:fldCharType"), tag)
        else:
            e = OxmlElement("w:instrText")
            e.text = txt
        r._element.append(e)
    set_font(r, 8)
    set_font(p.add_run(" －"), 8)


def page_break(doc):
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


def to_pdf(docx_path, pdf_path):
    subprocess.run(["soffice", f"-env:UserInstallation={LO_PROFILE}", "--headless", "--convert-to", "pdf", "--outdir", WORK, docx_path],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    os.replace(os.path.join(WORK, os.path.splitext(os.path.basename(docx_path))[0] + ".pdf"), pdf_path)


def cover(doc, title, sub, date_text, to, frm, note=None):
    for _ in range(5):
        para(doc)
    para(doc, "【架空】", 12, False, "c")
    para(doc, title, 20, True, "c", 10)
    para(doc, sub, 14, False, "c", 30)
    for _ in range(6):
        para(doc)
    para(doc, date_text, 12, False, "c", 30)
    para(doc, to, 12, False, "c", 4)
    para(doc, frm, 12, False, "c", 4)
    if note:
        para(doc, note, 9, False, "c")
    page_break(doc)


# ================================================================ 旧様式 2022 年度（s1 様式）
def build_2022(pdf_path):
    con = K.CONTRACTS["OH2022-1"]
    doc = new_doc("【架空】ひばり野クリーンセンター 2022年度 1号炉定期整備工事 報告書", "SPC保全課 s2")
    cover(doc, "2022年度 ひばり野CC 1号炉定期整備工事", "工 事 報 告 書", "令和4年12月", "ひばり野市長　様",
          "ひばり野環境サービス株式会社", "（施設：【架空】ひばり野クリーンセンター）")
    para(doc, "1. 工事概要", 13, True)
    for k, v in (("(1) 工事件名", con["title"]), ("(2) 施工場所", "【架空】ひばり野クリーンセンター　1号炉系統"),
                 ("(3) 契約工期", "令和4年9月1日 〜 令和4年11月30日"), ("(4) 炉停止期間", "令和4年10月3日 〜 令和4年11月2日"),
                 ("(5) 発注者", "ひばり野環境サービス株式会社（長期包括運営委託 年度補修計画による）"),
                 ("(6) 施工者", "鷹羽エンジニアリング株式会社（協力会社 野分工業株式会社）")):
        para(doc, f"{k}　{v}", 10.5, False, None, 2, 5)
    para(doc)
    para(doc, "2. 工事内容", 13, True)
    body = [
        ("2.1 燃焼設備",
         "焼却炉本体は、側壁および落じん部の耐火物に損耗が見られたため、不良部をはつり取りキャスタブルにより計2.8m²を補修した。"
         "乾燥ストーカは可動火格子24枚、燃焼ストーカは可動火格子38枚を交換した。燃焼ストーカの火格子は前回（2019年度）交換分の摩耗が想定より早く、"
         "残りの列についても次々回の定期整備までもたない見込みであるため、中長期整備計画の次回交換時期を見直す必要がある。"
         "落じんコンベヤはチェーン伸びが管理値を超えたため、チェーンを交換した。"),
        ("2.2 燃焼ガス冷却設備",
         "1号ボイラは電気事業法に基づく定期事業者検査の年にあたるため、ボイラ水洗の後、水管および過熱器の肉厚測定（UT）を行い、"
         "開放検査および水圧試験を受検した。水管の最小肉厚は4.1mmで、前回測定から有意な減少はなかった。"
         "エコノマイザは管外面に腐食による減肉が4箇所見られたため、肉盛溶接により補修した。"),
        ("2.3 排ガス処理設備",
         "消石灰定量供給機および活性炭定量供給機は分解整備を行い、テーブル・スクリューの摩耗部品を交換した。"
         "減温塔下ダストコンベヤおよびBF下飛灰コンベヤはスクレーパの摩耗部を交換した。バグフィルタは全室を点検し、ろ布の破損はなかった。"),
        ("2.4 通風設備",
         "誘引通風機は軸受の振動・温度を点検した。振動値は4.1mm/sで、前年度（3.8mm/s）から上昇しているため、"
         "運転中の傾向監視を続け、2024年度の定期整備で軸受交換およびランナ点検を行う計画とする。"),
        ("2.5 灰出し設備",
         "灰押出装置は底板の腐食部に当て板を行った。あわせて残厚を測定した結果、当て板範囲外の最小残厚は5.1mmであり、"
         "次回（2024年度）に底板の補修範囲を広げる必要がある。炉下灰コンベヤおよび飛灰搬送コンベヤはチェーン・スクレーパの部分交換を行った。"),
    ]
    for h, t in body:
        para(doc, h, 11, True, None, 2)
        para(doc, t, 10, False, None, 6, 5)
    para(doc, "3. 点検結果および所見", 13, True)
    para(doc, "点検の結果、上記の補修を行った箇所を除き、運転に支障のある異常は認められなかった。主な測定結果を以下に示す。", 10, False, None, 4, 5)
    table(doc, ["機器", "測定項目", "前回（2020）", "今回（2022）", "所見"], [
        ("1号ボイラ 水管", "最小肉厚", "4.2mm", "4.1mm", "良"),
        ("1号過熱器", "最小肉厚", "4.3mm", "4.3mm", "良"),
        ("1号誘引通風機", "軸受振動", "3.2mm/s", "4.1mm/s", "上昇傾向 要監視"),
        ("1号燃焼ストーカ", "火格子すき間", "3〜5mm", "4〜9mm", "摩耗進行"),
        ("1号灰押出装置", "底板残厚", "－", "5.1mm（範囲外）", "次回補修"),
    ], [30, 25, 25, 30, 40])
    para(doc)
    para(doc, "4. 使用材料", 13, True)
    table(doc, ["品名", "仕様", "数量", "備考"], [
        ("乾燥ストーカ 可動火格子", "TKW-RS105D-G", "24枚", "SPC支給"),
        ("燃焼ストーカ 可動火格子", "TKW-RS105C-G", "38枚", "SPC支給"),
        ("落じんコンベヤ チェーン", "スクレーパ 幅600用", "1式", "SPC支給"),
        ("キャスタブル耐火物", "高アルミナ質", "1.4t", "施工者手配"),
        ("当て板", "SS400 t9", "1.8m²", "施工者手配"),
        ("定量供給機 補修部品", "消石灰・活性炭用", "各1式", "SPC支給"),
    ], [45, 45, 20, 50])
    para(doc)
    para(doc, "5. 今後の対応", 13, True)
    para(doc, "(1) 燃焼ストーカ火格子の残り列の交換時期を、2025年度から2024年度へ前倒しすることを検討する。", 10, False, None, 2, 5)
    para(doc, "(2) 誘引通風機は振動の傾向監視を継続し、2024年度に軸受交換を行う。", 10, False, None, 2, 5)
    para(doc, "(3) 灰押出装置は2024年度に底板補修を行う。", 10, False, None, 2, 5)
    para(doc, "以上", 10.5, False, "r")
    para(doc, "添付：作業日報（写し）、工程表、写真帳（別冊）", 9)
    footer(doc, "2022年度 ひばり野CC 1号炉定期整備工事 報告書")
    d = os.path.join(WORK, "rep2022.docx")
    doc.save(d)
    to_pdf(d, pdf_path)


# ================================================================ 新様式 2024 年度（s2 様式）
SUMMARY_2024 = [
    ("HB02-101", "1号焼却炉本体", "補修", "側壁・落じん部 耐火物補修 計3.4m²", "完了", "毎年"),
    ("HB02-102", "1号乾燥ストーカ", "点検", "火格子すき間・摩耗測定（全列）", "管理値内", "－"),
    ("HB02-103", "1号燃焼ストーカ", "補修", "可動火格子 120枚交換（支給品2回目 2日遅れ）", "完了", "2028"),
    ("HB02-104", "1号後燃焼ストーカ", "補修", "可動火格子 18枚交換", "完了", "－"),
    ("HB02-105", "1号ストーカ油圧装置", "点検", "作動油分析・ストレーナ清掃", "NAS 8級 良", "－"),
    ("HB01-101〜103", "1号ホッパ・ホッパゲート・給じん装置", "点検", "摺動面・ライナ・シリンダ点検", "良", "－"),
    ("HB03-101", "1号ボイラ本体", "点検", "水洗、UT 216点、定期事業者検査、水圧試験", "最小4.0mm 良", "2026"),
    ("HB03-102", "1号蒸気ドラム", "点検", "内部確認（定期事業者検査）", "良", "2026"),
    ("HB03-103", "1号過熱器", "点検", "UT 96点", "2次入口 外面減肉 最小3.6mm", "2026 補修"),
    ("HB03-106/107", "1号スートブロワ（長抜差・定置）", "分解整備", "パッキン・チェーン交換、弁座すり合わせ", "完了", "－"),
    ("HB08-102", "1号ボイラ給水ポンプB", "分解整備", "メカシール・軸受交換", "完了", "－"),
    ("HB04-101", "1号減温塔", "点検", "内部清掃、噴霧ノズル6本交換", "良", "－"),
    ("HB04-108/109", "1号バグフィルタ・払落し装置", "点検", "全室点検、ろ布3本抜取（強度試験）", "強度残存率72%", "2026 全数交換"),
    ("HB04-117", "1号触媒反応塔", "点検", "外観・閉塞確認、試料採取（活性試験）", "相対活性 0.78〜0.91", "2026 試験"),
    ("HB04-105/107", "1号消石灰・活性炭 気送ブロワ", "分解整備", "ロータ・軸受点検交換", "完了", "－"),
    ("HB06-104", "1号誘引通風機", "分解整備", "軸受交換、ランナ肉盛・バランス修正（追加）", "振動 4.6→2.1mm/s", "－"),
    ("HB06-102", "1号二次送風機", "分解整備", "軸受・Vベルト交換", "完了", "－"),
    ("HB06-101", "1号押込送風機", "点検", "軸受振動・温度、ケーシング内清掃", "良", "－"),
    ("HB04-111/112 他", "1号ダンパ類（BF入口・出口、IDF入口、FDF出口）", "点検", "開閉・軸封点検", "良", "－"),
    ("HB07-101", "1号灰押出装置", "補修", "底板当て板2.4m²、側板当て板1.5m²（追加）", "完了", "－"),
    ("HB02-110/HB07-103/104", "1号落じん・炉下灰・飛灰搬送コンベヤ", "点検", "チェーン伸び測定・テンション調整", "管理値内", "－"),
    ("HB10-103/104", "1号IDF・FDF インバータ盤", "分解整備", "冷却ファン・平滑コンデンサ交換、絶縁抵抗測定", "完了", "－"),
]

MAT_2024 = [
    ("燃焼ストーカ 可動火格子", "TKW-RS105C-G", "120", "120", "0", "2回に分けて納入（2回目 10/17）"),
    ("後燃焼ストーカ 可動火格子", "TKW-RS105A-G", "20", "18", "2", "予備品として保管"),
    ("誘引通風機 軸受", "22228CCK/W33", "3", "2", "1", "予備品として保管"),
    ("二次送風機 軸受", "6316ZZ", "2", "2", "0", ""),
    ("二次送風機 Vベルト", "C-150", "4", "4", "0", ""),
    ("給水ポンプ メカニカルシール", "SPF-55用", "1", "1", "0", ""),
    ("給水ポンプ 軸受", "6312C3", "2", "2", "0", ""),
    ("減温塔 噴霧ノズル", "二流体 1/2B", "8", "6", "2", "予備品として保管"),
    ("BF ろ布", "PTFE φ165×6000", "3", "3", "0", "抜取り分の補充"),
    ("スートブロワ 補修部品", "長抜差4台・定置10台分", "1式", "1式", "－", ""),
    ("インバータ 補修部品", "IDF・FDF 盤用", "1式", "1式", "－", ""),
    ("気送ブロワ 補修部品", "消石灰・活性炭用", "2式", "2式", "－", ""),
]


def man_hours_2024():
    rows = K.daily_rows(K.T2024, 2024, seed_add=24)
    return sum(m * h for v in rows.values() for (t, d_, m, h) in v if t[3] != "SPC運転")


def build_2024(pdf_path):
    con = K.CONTRACTS["OH2024-1"]
    doc = new_doc("【架空】ひばり野クリーンセンター 2024年度 1号炉定期整備工事 報告書", "SPC保全課 s2")
    # 表紙は旧様式のまま（日付だけ西暦に）
    cover(doc, "2024年度 ひばり野CC 1号炉定期整備工事", "工 事 報 告 書", "2024年12月10日", "ひばり野市長　様",
          "ひばり野環境サービス株式会社", "（施設：【架空】ひばり野クリーンセンター）")
    # 1 ページ目：サマリー
    para(doc, "工事報告 サマリー", 14, True, "c", 6)
    table(doc, ["項目", "内容"], [
        ("工事件名", con["title"]), ("施設・系統", "【架空】ひばり野クリーンセンター　1号炉系統"),
        ("契約工期", "2024/09/02 〜 2024/11/29"), ("炉停止", "2024/10/01 〜 2024/10/31（ごみ投入再開 11/02）"),
        ("施工", "鷹羽エンジニアリング㈱（協力会社 野分工業㈱、築炉 白砂耐火工業、足場 白鷺鳶工業、UT 東栄非破壊検査）"),
        ("延べ工数", f"協力会社日報集計 {man_hours_2024():,.0f} 人・時（SPC 運転・立会を除く）"),
        ("追加・変更", "IDF ランナ肉盛・バランス修正、灰押出装置 側板当て板 1.5m²（変更契約 10/18）"),
        ("遅延", "燃焼ストーカ火格子（支給品）2回目納入が 2日遅れ → 10/19（土）作業で吸収、炉立上げは予定どおり"),
    ], [35, 135], 9)
    para(doc, "")
    table(doc, ["資産番号", "機器", "区分", "実施内容", "結果", "次回"], SUMMARY_2024, [22, 36, 15, 50, 27, 18], 7.5)
    page_break(doc)
    para(doc, "1. 工事概要", 13, True)
    for k, v in (("(1) 工事件名", con["title"]), ("(2) 施工場所", "【架空】ひばり野クリーンセンター　1号炉系統"),
                 ("(3) 契約工期", "2024年9月2日 〜 2024年11月29日"), ("(4) 炉停止期間", "2024年10月1日 〜 2024年10月31日"),
                 ("(5) 発注者", "ひばり野環境サービス株式会社（長期包括運営委託 年度補修計画による）"),
                 ("(6) 施工者", "鷹羽エンジニアリング株式会社（協力会社 野分工業株式会社）")):
        para(doc, f"{k}　{v}", 10.5, False, None, 2, 5)
    para(doc, "2. 実施経過（予定と実績）", 13, True)
    rows = []
    for t in K.T2024:
        if t[3] == "SPC運転" and t[0] not in ("T01", "T30"):
            continue
        ps, pe = t[5]
        a = t[6] or t[5]
        rows.append((t[2], t[3], f"{ps}〜{pe}", f"{a[0]}〜{a[1]}", "変更" if t[6] and t[6] != t[5] else ""))
    table(doc, ["作業項目", "班", "予定", "実績", "差"], rows, [70, 22, 25, 25, 12], 8)
    para(doc)
    para(doc, "3. 機器別の記録", 13, True)
    para(doc, "3.1 ボイラ・過熱器 肉厚測定（UT）", 11, True)
    table(doc, ["部位", "測定点数", "公称", "最小", "前回(2022)", "判定"], [
        ("水管 前壁", "84", "4.5", "4.0", "4.1", "良"), ("水管 側壁", "132", "4.5", "4.2", "4.2", "良"),
        ("過熱器 1次", "48", "4.5", "4.2", "4.3", "良"), ("過熱器 2次 入口側", "48", "4.5", "3.6", "4.3", "要補修（2026）"),
    ], [40, 22, 18, 18, 25, 37], 8.5)
    para(doc, "過熱器2次の入口側に外面減肉が見られた。前回からの減少が大きいため、2025年度の定期整備で再測定し、2026年度に補修（部分取替または肉盛）を計画に入れる。", 9.5, False, None, 6, 3)
    para(doc, "3.2 誘引通風機", 11, True)
    table(doc, ["項目", "整備前", "整備後", "備考"], [
        ("軸受振動（反負荷側）", "4.6mm/s", "2.1mm/s", "管理値 4.5mm/s"), ("軸受温度", "68℃", "52℃", "負荷運転時"),
        ("ランナ羽根 入口側", "摩耗あり（最大 2.5mm）", "肉盛・仕上げ", "追加作業"),
    ], [45, 35, 35, 55], 8.5)
    para(doc, "3.3 バグフィルタ・触媒", 11, True)
    table(doc, ["項目", "結果", "判断"], [
        ("ろ布 引張強度 残存率（3本平均）", "72%（2024/11/12 報告）", "全数交換は 2026 年度へ延伸可"),
        ("触媒 相対活性 1段／2段／3段", "0.78／0.85／0.91（2024/11/27 報告）", "継続使用。2026 年度に再試験"),
    ], [55, 60, 55], 8.5)
    para(doc)
    para(doc, "4. 支給品・材料の使用", 13, True)
    table(doc, ["品名", "型式", "支給数", "使用", "残（予備）", "備考"], MAT_2024, [42, 34, 15, 15, 18, 46], 8)
    para(doc, "施工者手配材料：キャスタブル耐火物 1.6t、当て板 SS400 t9 3.9m²、溶接材料 一式", 9, False, None, 6)
    para(doc, "5. 所見と次回計画への反映", 13, True)
    for s in ("(1) 1号過熱器 2次入口側：2026年度 補修を中長期整備計画に追加する（2025年度改訂）。",
              "(2) 1号バグフィルタ ろ布：試験結果により全数交換を 2025年度から 2026年度へ延伸する。",
              "(3) 1号誘引通風機：整備後の振動は良好。廃止までの分解整備は行わず、傾向監視とする。",
              "(4) 燃焼ストーカ火格子の支給品は、次回から 1回納入（定修 2週間前）とするよう購買に依頼した。"):
        para(doc, s, 10, False, None, 2, 3)
    page_break(doc)
    para(doc, "写 真 帳", 14, True, "c")
    photo_box(doc, ["写真1 燃焼ストーカ 交換前", "写真2 燃焼ストーカ 交換後", "写真3 IDF ランナ 摩耗部", "写真4 IDF ランナ 肉盛後"])
    para(doc)
    photo_box(doc, ["写真5 灰押出装置 底板 当て板", "写真6 灰押出装置 側板（追加）", "写真7 過熱器 2次入口 減肉部", "写真8 水圧試験"])
    footer(doc, "2024年度 ひばり野CC 1号炉定期整備工事 報告書")
    d = os.path.join(WORK, "rep2024.docx")
    doc.save(d)
    to_pdf(d, pdf_path)


# ================================================================ 市追加工事 R5（メーカー、reportlab）
def build_city(pdf_path):
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
    pdfmetrics.registerFont(TTFont("IPAG", "/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf"))
    st = ParagraphStyle("n", fontName="IPAG", fontSize=9.5, leading=15)
    h1 = ParagraphStyle("h1", fontName="IPAG", fontSize=13, leading=20, spaceBefore=6, spaceAfter=4)
    h2 = ParagraphStyle("h2", fontName="IPAG", fontSize=11, leading=17, spaceBefore=4, spaceAfter=2)
    ct = ParagraphStyle("c", fontName="IPAG", fontSize=18, leading=28, alignment=1)
    cs = ParagraphStyle("cs", fontName="IPAG", fontSize=12, leading=20, alignment=1)
    con = K.CONTRACTS["CITY-R5"]

    def tbl(data, widths, fs=8.5):
        t = Table(data, colWidths=[w * mm for w in widths], repeatRows=1)
        t.setStyle(TableStyle([("FONT", (0, 0), (-1, -1), "IPAG", fs), ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
                               ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E7E6E6")), ("VALIGN", (0, 0), (-1, -1), "MIDDLE")]))
        return t

    def on_page(c, d):
        c.setFont("IPAG", 7.5)
        if d.page > 1:
            c.drawString(20 * mm, 287 * mm, "令和5年度 ひばり野クリーンセンター 2号ボイラ水管補修工事 完成報告書")
            c.drawRightString(190 * mm, 287 * mm, "鷹羽エンジニアリング株式会社")
            c.drawCentredString(105 * mm, 10 * mm, f"{d.page - 1}")
        c.drawString(20 * mm, 6 * mm, "TKW-QF-0712 Rev.4")

    doc = SimpleDocTemplate(pdf_path, pagesize=A4, leftMargin=20 * mm, rightMargin=20 * mm, topMargin=20 * mm, bottomMargin=18 * mm,
                            title="【架空】令和5年度 ひばり野クリーンセンター 2号ボイラ水管補修工事 完成報告書",
                            author="鷹羽エンジニアリング株式会社 m1")
    s = [Spacer(1, 50 * mm), Paragraph("【架空】", cs), Paragraph(con["title"], ct), Spacer(1, 8 * mm), Paragraph("完 成 報 告 書", ct),
         Spacer(1, 60 * mm), Paragraph("令和5年11月30日", cs), Spacer(1, 6 * mm), Paragraph("ひばり野市長　様", cs), Spacer(1, 4 * mm),
         Paragraph("鷹羽エンジニアリング株式会社　環境プラント事業部", cs), Paragraph("現場代理人　宇佐見", cs), PageBreak()]
    s.append(Paragraph("目　次", h1))
    for x in ("1. 工事概要", "2. 工事経過", "3. 施工内容", "4. 検査記録", "5. 使用材料", "6. 写真"):
        s.append(Paragraph(x, st))
    s.append(Spacer(1, 6 * mm))
    s.append(Paragraph("1. 工事概要", h1))
    s.append(tbl([["項目", "内容"], ["工事名", con["title"]], ["工事場所", "ひばり野市（【架空】ひばり野クリーンセンター 2号ボイラ）"],
                  ["発注者", "ひばり野市（環境課）　監督員 日下部"], ["受注者", "鷹羽エンジニアリング株式会社　現場代理人 宇佐見"],
                  ["工期", "令和5年9月25日 〜 令和5年12月15日"], ["現地施工", "令和5年10月10日 〜 令和5年10月27日（2号炉定期整備の炉停止期間内）"],
                  ["対象設備", "2号ボイラ本体（資産番号 HB03-201）第1放射室 前壁 水管 " + K.TUBE_SPEC],
                  ["工事内容", "減肉水管 6本 部分取替（各1.8m）、減肉部 肉盛溶接 12箇所、非破壊検査、水圧試験、耐火物復旧"]], [32, 138]))
    s.append(Paragraph("2. 工事経過", h1))
    rows = [["月日", "作業内容", "人員"]]
    from mk_daily import CITY_DAY
    for k, (txt, men, q, note) in CITY_DAY.items():
        x = dt.date(2023, *K.d(k))
        rows.append([f"10/{x.day}（{K.WD[x.weekday()]}）", Paragraph(txt, ParagraphStyle("t", fontName="IPAG", fontSize=8, leading=11)),
                     str(sum(men.values()))])
    s.append(tbl(rows, [24, 132, 14], 8))
    s.append(Paragraph("3. 施工内容", h1))
    s.append(Paragraph(
        "令和4年度の2号炉定期整備（ひばり野環境サービス㈱）の肉厚測定において、第1放射室前壁の水管に減肉（最小3.2mm、設計必要厚2.9mm）が"
        "確認されたことを受け、本工事を実施した。着工時に前壁63点の肉厚を再測定し、最小肉厚3.5mm以下の6本（F-14〜F-16、F-21〜F-23）を"
        "部分取替、3.8〜4.0mmの12箇所を肉盛溶接の対象とした。取替は既設管を各1.8m切断撤去し、新管を突合せ溶接（TIG初層＋被覆アーク）で"
        "接続した（計12継手）。", st))
    s.append(Paragraph(
        "放射線透過試験の結果、F-22下側継手にブローホール（3類）が認められたため、当該部をはつり取り補修溶接の上、再撮影し合格を確認した。"
        "この補修のため10月21日（土）に作業を追加したが、全体工程への影響はなかった。肉盛部は浸透探傷試験で指示模様のないことを確認し、"
        "肉盛後の肉厚は4.6〜5.2mmであった。水圧試験（4.5MPa、30分保持）は市監督員およびSPCの立会のもと実施し、漏れのないことを確認した。"
        "最後に、はつり取った範囲の耐火物をキャスタブル（4.2m²）で復旧した。", st))
    s.append(Paragraph("4. 検査記録", h1))
    s.append(Paragraph("4.1 肉厚測定（取替対象）", h2))
    data = [["管番号", "令和4年度 最小(mm)", "着工時 最小(mm)", "措置"]] + [[a, f"{b:.1f}", f"{c:.1f}", d_] for a, b, c, d_ in K.TUBES]
    s.append(tbl(data, [25, 40, 40, 65]))
    s.append(Paragraph("4.2 肉盛溶接 箇所", h2))
    data = [["管番号", "着工時 最小(mm)", "PT", "肉盛後(mm)"]]
    for i, (a, b) in enumerate(K.OVERLAY):
        data.append([a, f"{b:.1f}", "合格", f"{4.6 + (i * 7 % 6) / 10:.1f}"])
    s.append(tbl(data, [30, 45, 30, 45]))
    s.append(Paragraph("4.3 放射線透過試験（RT）", h2))
    data = [["継手", "撮影日", "判定", "再撮影", "最終"]]
    for a, *_ in K.TUBES:
        for pos in ("上", "下"):
            bad = (a == "F-22" and pos == "下")
            data.append([f"{a} {pos}側", "10/20", "3類 不合格" if bad else "1類 合格", "10/21 1類 合格" if bad else "－", "合格"])
    s.append(tbl(data, [30, 25, 35, 45, 35]))
    s.append(Paragraph("4.4 水圧試験：令和5年10月24日 試験圧力 4.5MPa 保持30分 漏れなし（立会：市 日下部、SPC 小鳥遊）", st))
    s.append(Paragraph("5. 使用材料", h1))
    s.append(tbl([["品名", "規格", "数量", "備考"],
                  ["ボイラ水管", "STB340 φ60.5×t4.5 L=2.0m", "7本", "6本使用、1本は市へ予備品として引渡し"],
                  ["溶接材料", "TIG棒 YGT50・被覆棒 E4316", "1式", ""],
                  ["キャスタブル耐火物", "高アルミナ質", "1.9t", "4.2m²"],
                  ["Yアンカー", "SUS310S", "120本", ""]], [40, 60, 20, 50]))
    s.append(PageBreak())
    s.append(Paragraph("6. 写真", h1))
    ph = [["", ""], ["写真1 着工前 前壁水管（はつり後）", "写真2 既設管 切断撤去"], ["", ""], ["写真3 新管 溶接完了", "写真4 肉盛溶接 完了"],
          ["", ""], ["写真5 RT 撮影（不合格部 補修後）", "写真6 耐火物復旧 完了"]]
    t = Table(ph, colWidths=[85 * mm, 85 * mm], rowHeights=[62 * mm, 8 * mm] * 3)
    t.setStyle(TableStyle([("FONT", (0, 0), (-1, -1), "IPAG", 8.5), ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                           ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#F2F2F2")), ("BACKGROUND", (0, 2), (-1, 2), colors.HexColor("#F2F2F2")),
                           ("BACKGROUND", (0, 4), (-1, 4), colors.HexColor("#F2F2F2")), ("ALIGN", (0, 0), (-1, -1), "CENTER")]))
    s.append(t)
    doc.build(s, onFirstPage=on_page, onLaterPages=on_page)


FN = {"2022": "【架空】ひばり野クリーンセンター_工事報告書_2022年度1号炉定期整備工事_s1様式.pdf",
      "2024": "【架空】ひばり野クリーンセンター_工事報告書_2024年度1号炉定期整備工事_s2様式.pdf",
      "R5": "【架空】ひばり野クリーンセンター_工事報告書_令和5年度2号ボイラ水管補修工事_メーカー様式.pdf"}

if __name__ == "__main__":
    p = {k: os.path.join(OUT, v) for k, v in FN.items()}
    build_2022(p["2022"])
    build_2024(p["2024"])
    build_city(p["R5"])
    print(render(p["2022"], [1, 2]))
    print(render(p["2024"], [2, 3]))
    print(render(p["R5"], [3, 4]))
