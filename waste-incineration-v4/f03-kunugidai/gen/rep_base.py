# -*- coding: utf-8 -*-
"""工事報告書（メーカー本社 r2。Word で作って PDF で提出）。
旧：r2 が H20（2008）に作った文章型の完了報告書（写真は別冊）。H25 1号炉。
新：r2 が R2（2020）に改めた様式。冒頭にサマリー表、機器別の表＋写真、部品の区分表。R6 2号炉・R6 全炉停止時共通。
原価・金額は載せない。
"""
import os
from docx import Document
from docx.shared import Pt, Mm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from PIL import Image, ImageDraw, ImageFont
import xlhelp as X
from world import CONTRACTS
import datetime as dt

REC = None
REC_ON = [True]
FONT = "IPAPGothic"
FONT_M = "IPAGothic"


def _doc(title, author, subject):
    d = Document()
    sec = d.sections[0]
    sec.page_width, sec.page_height = Mm(210), Mm(297)
    sec.left_margin = sec.right_margin = Mm(20)
    sec.top_margin, sec.bottom_margin = Mm(18), Mm(18)
    st = d.styles["Normal"]
    st.font.name = FONT
    st.font.size = Pt(10)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    for hn, sz in (("Heading 1", 13), ("Heading 2", 11), ("Heading 3", 10)):
        h = d.styles[hn]
        h.font.name = FONT
        h.font.size = Pt(sz)
        h.font.bold = True
        h.font.color.rgb = RGBColor(0, 0, 0)
        h.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    cp = d.core_properties
    cp.title = title
    cp.author = author
    cp.subject = subject
    cp.created = X.FIXED_DT
    cp.modified = X.FIXED_DT
    cp.last_modified_by = author
    return d


def _p(d, text="", size=10, bold=False, align=None, space_after=4):
    if REC is not None and REC_ON[0] and text and text.strip() not in ("以上",) and not text.startswith("（別冊）"):
        REC.addp("item", {"text": text.strip()})
    p = d.add_paragraph()
    r = p.add_run(text)
    r.font.size = Pt(size)
    r.bold = bold
    if align == "c":
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == "r":
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.space_after = Pt(space_after)
    return p


def _hdr_repeat(row):
    trPr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:tblHeader")
    el.set(qn("w:val"), "true")
    trPr.append(el)


def _shade(cell, hexrgb):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hexrgb)
    tcPr.append(shd)


def _table(d, header, rows, widths, size=8.5, head_fill="D9D9D9"):
    t = d.add_table(rows=1, cols=len(header))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(header):
        c = t.rows[0].cells[i]
        c.text = ""
        r = c.paragraphs[0].add_run(h)
        r.bold = True
        r.font.size = Pt(size)
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        _shade(c, head_fill)
    _hdr_repeat(t.rows[0])
    if REC is not None and REC_ON[0]:
        for row in rows:
            REC.addp("item", {h: str(v) for h, v in zip(header, row)})
    for row in rows:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = ""
            r = cells[i].paragraphs[0].add_run(str(v))
            r.font.size = Pt(size)
    _widths(t, widths)
    d.add_paragraph().paragraph_format.space_after = Pt(2)
    return t


def _widths(t, widths):
    t.autofit = False
    grid = t._tbl.tblGrid
    for i, gc in enumerate(grid.findall(qn("w:gridCol"))):
        gc.set(qn("w:w"), str(int(widths[i] * 56.7)))
    for row in t.rows:
        for i, w in enumerate(widths):
            row.cells[i].width = Mm(w)


def _footer(d, text):
    sec = d.sections[0]
    p = sec.footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(text + "　－ ")
    r.font.size = Pt(8)
    for code in ("PAGE",):
        f1 = OxmlElement("w:fldSimple")
        f1.set(qn("w:instr"), code)
        rr = OxmlElement("w:r")
        tt = OxmlElement("w:t")
        tt.text = "1"
        rr.append(tt)
        f1.append(rr)
        p._p.append(f1)
    r2 = p.add_run(" －")
    r2.font.size = Pt(8)


def _header(d, text):
    p = d.sections[0].header.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = p.add_run(text)
    r.font.size = Pt(8)


def _photo(path, caption, seed):
    import random
    rng = random.Random(seed)
    im = Image.new("RGB", (640, 420), (150, 150, 150))
    dr = ImageDraw.Draw(im)
    for _ in range(40):
        x, y = rng.randrange(640), rng.randrange(420)
        g = rng.randint(90, 200)
        dr.rectangle((x, y, x + rng.randint(20, 140), y + rng.randint(10, 80)), fill=(g, g, g - 5))
    f = ImageFont.truetype("/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf", 24)
    dr.rectangle((0, 380, 640, 420), fill=(40, 40, 40))
    dr.text((12, 386), "（写真）" + caption, font=f, fill=(255, 255, 255))
    im.save(path)


def _photos(d, items, tmp, seed):
    t = d.add_table(rows=2, cols=len(items))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, cap in enumerate(items):
        pth = os.path.join(tmp, "ph_%d_%d.png" % (seed, i))
        _photo(pth, cap, seed * 10 + i)
        c = t.rows[0].cells[i]
        c.paragraphs[0].add_run().add_picture(pth, width=Mm(72))
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        if REC is not None and REC_ON[0]:
            REC.addp("item", {"caption": "写真 %s" % cap}, image=os.path.basename(pth))
        r = t.rows[1].cells[i].paragraphs[0].add_run("写真 %s" % cap)
        r.font.size = Pt(8)
        t.rows[1].cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    d.add_paragraph().paragraph_format.space_after = Pt(2)


def body():
    REC_ON[0] = True


def _cover(d, lines, no=None):
    REC_ON[0] = False
    for _ in range(6):
        _p(d, "")
    if no:
        _p(d, "報告書番号　" + no, 10, align="r")
    _p(d, "【架空】", 14, True, "c")
    for ln, sz in lines:
        _p(d, ln, sz, True, "c", 10)
    for _ in range(10):
        _p(d, "")


# ------------------------------------------------------------------ 旧（H20 様式）H25 1号炉
def make_old(path_pdf, tmp):
    ct = CONTRACTS["H25-1"]
    d = _doc("【架空】平成25年度 くぬぎ台環境センター 1号炉定期整備業務委託 完了報告書", "汐見重工環境 r2", "工事報告書")
    _cover(d, [("平成25年度", 16), ("くぬぎ台環境センター", 16), ("1号炉定期整備業務委託", 18), ("完 了 報 告 書", 22)])
    _p(d, "平成25年6月", 12, align="c")
    _p(d, "汐見重工環境株式会社", 13, True, "c")
    _p(d, "環境プラント本部 保守サービス部", 11, align="c")
    d.add_page_break()
    body()
    _header(d, "【架空】H25 くぬぎ台環境センター 1号炉定期整備業務委託 完了報告書")
    _footer(d, "汐見重工環境株式会社")
    d.add_heading("目　次", 1)
    REC_ON[0] = False
    for t in ["1. 工事概要", "2. 工事工程", "3. 整備内容", "4. 測定結果", "5. 交換部品・材料", "6. 所見", "（別冊）写真帳"]:
        _p(d, "　" + t, 10)
    REC_ON[0] = True
    d.add_heading("1. 工事概要", 1)
    _table(d, ["項目", "内容"], [
        ("工事名", ct["name"]), ("工事場所", "くぬぎ台市 くぬぎ台環境センター 1号炉"),
        ("工期", "平成25年5月13日 ～ 平成25年6月14日"), ("作業期間", "平成25年5月13日 ～ 平成25年6月7日（炉停止 5/13、立上げ 6/5）"),
        ("発注者", "くぬぎ台市 環境部 環境センター"), ("受注者", "汐見重工環境株式会社"),
        ("現場代理人", "m2"), ("報告者", "r2（保守サービス部）")], [35, 130], 9)
    d.add_heading("2. 工事工程", 1)
    _p(d, "工事は作業工程表（平成25年4月26日付）のとおり実施した。5月13日に市の運転により炉を立ち下げ、冷却後 5月16日から足場を組み、"
          "燃焼設備・ボイラ・排ガス処理設備の点検整備を行った。6月3日にボイラ水圧試験（市立会）を行い、6月5日から耐火物乾燥・立上げ、"
          "6月7日の試運転をもって作業を終えた。工程の遅れはなかった。")
    d.add_heading("3. 整備内容", 1)
    secs = [
        ("3.1 燃焼設備", "燃焼ストーカは燃焼段 2〜3 列目の火格子の摩耗が大きく、前回（H22）交換以後の摩耗量は最大 11mm であった。"
         "摩耗の大きい 24 枚を新品に交換し、左右のサイドシールを交換した。交換後の空打ちで異音・引っかかりはない。乾燥ストーカ・後燃焼ストーカは"
         "目視・打診点検で異常なし。給じん装置はプッシャ摺動部の当たりが良好で、調整のみとした。炉内耐火物は右側壁 下部の損傷部（約 1.8m²）をはつり、"
         "キャスタブルで補修した。"),
        ("3.2 ボイラ", "水管・過熱器・エコノマイザの肉厚を超音波厚さ計で測定した（4 章）。スートブロワ SB-1・SB-2 を分解し、ランスチューブ・パッキン・"
         "軸受を交換した。ドラム安全弁 2 台・過熱器安全弁 1 台を分解し、弁座を擦り合わせ、吹出し圧力を調整した（2.20／2.25／2.10MPa）。"
         "6月3日に水圧試験（2.6MPa、30 分保持）を行い、漏れのないことを市立会のもと確認した。"),
        ("3.3 排ガス処理設備", "減温塔の噴霧ノズル 6 本を交換した（摩耗による噴霧角の乱れ）。バグフィルタは全 6 室のろ布を点検し、4 室で破損ろ布 2 本を認めたため、"
         "市の予備品により交換した。触媒反応塔は入口面の閉塞が軽微で、エアブローで清掃した。洗煙塔は内部のライニングにふくれ・はく離なし、"
         "充填材の閉塞なし。"),
        ("3.4 通風設備", "IDF（誘引通風機）は昨年度の基幹改良で更新した機器で、軸受の状態・振動（1.6mm/s）とも良好。押込送風機は Vベルトの張りを調整した。"),
        ("3.5 電気計装設備", "1号炉 動力制御盤、現場盤の点検、電動機の絶縁抵抗測定、温度計 8 点・圧力計 4 点の校正を行い、インターロック試験を実施した。"),
    ]
    for h, t in secs:
        d.add_heading(h, 2)
        _p(d, t)
    _p(d, "なお、1号炉 灰押出装置のシュート当て板補修（5/21〜23）は市の直営で行われ、本業務の範囲外である。", 9)
    d.add_heading("4. 測定結果", 1)
    _p(d, "ボイラ肉厚測定結果（最小値、単位 mm）")
    _table(d, ["部位", "材質・寸法", "設計肉厚", "H22 測定", "H25 測定", "管理値", "判定"], [
        ("水冷壁 前壁", "STB340 φ50.8", "3.5", "3.3", "3.2", "2.4", "良"),
        ("1次過熱器 下段", "STB410 φ38.1", "4.5", "3.9", "3.6", "2.6", "良（減肉傾向）"),
        ("2次過熱器", "SUS310S φ38.1", "4.0", "3.9", "3.9", "2.6", "良"),
        ("エコノマイザ", "STB340 φ38.1", "3.5", "3.2", "3.1", "2.2", "良")], [30, 30, 18, 18, 18, 16, 25], 8.5)
    d.add_heading("5. 交換部品・材料", 1)
    _table(d, ["品名", "仕様", "数量", "備考"], [
        ("火格子", "燃焼段用 SCH13", "24 枚", "当社手配"), ("サイドシール", "燃焼段 左右", "2 組", "当社手配"),
        ("キャスタブル", "高アルミナ質", "0.8 t", "当社手配"), ("噴霧ノズル", "二流体 SUS316", "6 本", "当社手配"),
        ("ろ布", "PPS φ150×6,000", "2 本", "市 予備品を使用"), ("SB ランスチューブ", "SB-1・2 用", "2 本", "当社手配"),
        ("安全弁 パッキン類", "", "1 式", "当社手配")], [35, 55, 20, 55], 8.5)
    d.add_heading("6. 所見", 1)
    _p(d, "1次過熱器 下段は減肉が進んでいる（3 年で 0.3mm）。管理値までの余裕は約 1mm あるが、次回以降も毎年の測定を続け、"
          "取替の時期を判断されたい。燃焼ストーカの火格子は 3 年周期での部分交換を継続することを提案する。")
    _p(d, "以上", 10, align="r")
    _p(d, "（別冊）写真帳 1 冊", 9)
    dx = os.path.join(tmp, os.path.splitext(os.path.basename(path_pdf))[0] + ".docx")
    d.save(dx)
    return _pdf(dx, path_pdf)


def _pdf(dx, path_pdf):
    PENDING.append((dx, path_pdf))
    return path_pdf


PENDING = []


# ------------------------------------------------------------------ 新（R2 様式）R6 2号炉
def make_new_2(path_pdf, tmp):
    ct = CONTRACTS["R6-2"]
    no = "SH-KN-2024-02"
    d = _doc("【架空】令和6年度 くぬぎ台環境センター 2号炉定期整備業務委託 定期整備報告書", "汐見重工環境 r2", "工事報告書")
    _cover(d, [("令和6年度　くぬぎ台環境センター", 15), ("2号炉定期整備業務委託", 18), ("定期整備報告書", 22)], no)
    _p(d, "2024年11月", 12, align="c")
    _p(d, "汐見重工環境株式会社", 13, True, "c")
    _p(d, "環境プラント本部 保守サービス部", 11, align="c")
    d.add_page_break()
    body()
    _header(d, "【架空】%s　%s" % (ct["name"], no))
    _footer(d, "汐見重工環境株式会社")
    d.add_heading("1. 報告サマリー", 1)
    _p(d, "工期 2024/10/7〜11/15、作業 10/7〜11/8。主な結果を下表に示す（判定 ◎良 ○軽微 △要監視 ×要対策）。"
          "追加作業 1 件（1次過熱器管 3 本取替、変更契約）を実施した。", 9)
    _table(d, ["No", "対象機器（タグ）", "実施作業", "結果", "判定", "次回推奨"], [
        ("1", "誘引通風機（2FD-006）", "羽根車 肉盛補修・現地バランス、軸受交換", "振動 4.2→1.1mm/s（据付後 1.9 に安定）", "○", "羽根車更新 2029 頃"),
        ("2", "バグフィルタ（2GT-014）", "ろ布全数交換 288 本、ケージ 5 本交換", "漏れ試験 合格、差圧 0.9kPa（立上げ後）", "◎", "強度試験 2 年ごと"),
        ("3", "1次過熱器（2BL-002）", "肉厚測定、【追加】管 3 本取替", "最小 3.1mm（取替後）、RT 合格", "△", "全数取替 2027 を推奨"),
        ("4", "ボイラ本体（2BL-001）", "水管肉厚測定、水圧試験（定期事業者検査）", "最小 3.0mm（前壁）、水圧試験 合格", "○", "毎年測定"),
        ("5", "燃焼ストーカ（2IN-005）", "火格子部分交換 30 枚", "空打ち 良", "◎", "3 年後"),
        ("6", "焼却炉本体（2IN-011）", "耐火物部分補修 2.5m²", "乾燥 良", "○", "2 年後"),
        ("7", "洗煙循環ポンプ A（2GT-036）", "分解整備、羽根車・ライナ・メカシール交換", "吐出圧 0.31MPa 回復", "◎", "4 年後"),
        ("8", "湿式洗煙塔（2GT-035）", "内部点検", "ミストエリミネータ 汚れ 中、清掃", "○", "2 年後"),
        ("9", "触媒反応塔（2GT-027）", "点検・触媒抜取り（活性試験は別途報告）", "閉塞 軽微、清掃", "○", "活性試験結果による"),
        ("10", "スートブロワ SB-3・4（2BL-006・007）", "分解整備", "作動 良", "◎", "3 年後"),
        ("11", "給じん装置（2IN-003）", "点検", "摺動板 摩耗 軽微", "○", "－"),
        ("12", "押込・二次送風機（2FD-001・002）", "点検", "良", "◎", "－"),
        ("13", "電気・計装", "MCC 点検、絶縁抵抗、計器校正、インターロック試験", "良", "◎", "－"),
    ], [8, 38, 42, 40, 9, 26], 7.5)
    d.add_heading("2. 工事概要", 1)
    _table(d, ["項目", "内容"], [
        ("工事名", ct["name"]), ("工期", "2024年10月7日 ～ 2024年11月15日"),
        ("変更契約", "2024年10月16日 1次過熱器管 3本取替の追加（市 施設係 k3 協議 10/15）"),
        ("発注者", "くぬぎ台市 環境部 環境センター"), ("受注者", "汐見重工環境株式会社"), ("現場代理人", "m2"), ("報告者", "r2")], [35, 130], 9)
    d.add_heading("3. 工程実績", 1)
    _table(d, ["作業", "予定", "実績", "差異・理由"], [
        ("炉立下げ・冷却（市）", "10/7〜10/9", "10/7〜10/9", "－"),
        ("ボイラ 肉厚測定", "10/10〜10/12", "10/10〜10/12", "減肉 3 本を発見"),
        ("BF ろ布全数交換・漏れ試験", "10/10〜10/16", "10/10〜10/17", "+1 日（ケージ曲がり 5 本交換）"),
        ("【追加】1次過熱器管 取替", "－", "10/17〜10/26", "追加作業"),
        ("誘引通風機 羽根車肉盛・軸受", "10/21〜10/26", "10/21〜10/28", "+2 日（肉盛範囲 9 枚に拡大）"),
        ("ボイラ水圧試験", "10/28", "10/29", "+1 日（追加作業のため）"),
        ("耐火物乾燥・立上げ（市）", "11/1〜11/2", "11/2〜11/3", "+1 日"),
        ("試運転・調整", "11/5〜11/8", "11/5〜11/8", "予定どおり")], [50, 30, 30, 55], 8.5)
    d.add_heading("4. 機器別 実施内容", 1)
    items = [
        ("4.1 誘引通風機（IDF、2FD-006）",
         [("作業前", "振動 4.2mm/s（R5 診断）。羽根入口側の摩耗、ダスト付着によるアンバランス。"),
          ("実施", "ケーシング開放、羽根車摩耗測定（最大 3.2mm、12 枚中 9 枚）。9 枚の入口側を肉盛溶接・グラインダ仕上げ。軸受（両側 22324）交換。現地バランス。"),
          ("作業後", "残留振動 1.1mm/s（試運転時）。"), ("判定", "○　肉盛はおおむね 4〜5 年の寿命。2029 年頃の羽根車更新を推奨。")],
         ["羽根車 入口側 摩耗（開放時）", "肉盛溶接 仕上げ後"]),
        ("4.2 バグフィルタ（2GT-014）",
         [("作業前", "差圧 1.65kPa（R5）、R5 に破損 3 本。R4 抜取り試験で引張強度 残存 62%。"),
          ("実施", "全 6 室 288 本のろ布を取外し、市支給の新品ろ布に交換。ケージ曲がり 5 本は市予備品で交換。蛍光粉による漏れ試験。"),
          ("作業後", "漏れなし。立上げ後の差圧 0.9kPa。"), ("判定", "◎")],
         ["ろ布 取付け（3 室）", "蛍光粉 漏れ試験"]),
        ("4.3 1次過熱器（2BL-002）【追加作業】",
         [("作業前", "10/12 の肉厚測定で下段 3 本が 2.4〜2.5mm（管理値 2.6mm 以下）。"),
          ("実施", "10/15 市と協議し取替を追加（変更契約）。弊社在庫の STB410 φ38.1×4.5t を用い、3 本（各 2.4m）を切断・取替、TIG 溶接 6 箇所、放射線透過試験 6 箇所 合格。"),
          ("作業後", "取替部以外の最小値 3.1mm。"), ("判定", "△　他の管も減肉 0.1〜0.2mm/年で進んでおり、2027 年度の全数取替を推奨（管材の手配に約 10 か月）。")],
         ["減肉部 取替前", "取替後 溶接部"]),
        ("4.4 洗煙循環ポンプ A（2GT-036）",
         [("作業前", "吐出圧 0.26MPa に低下（定格 0.30MPa）。"),
          ("実施", "分解、羽根車 ベーン先端 腐食摩耗 最大 4mm、ライナ摩耗。羽根車・ライナ・メカシールを新品に交換、組立・芯出し。"),
          ("作業後", "吐出圧 0.31MPa。"), ("判定", "◎　予備機 B（R4 整備）は異常なし。")],
         ["羽根車 摩耗状況", "組立後"]),
    ]
    for i, (h, rows, ph) in enumerate(items):
        d.add_heading(h, 2)
        _table(d, ["区分", "内容"], rows, [22, 143], 8.5, "EDEDED")
        _photos(d, ph, tmp, 200 + i)
    d.add_heading("4.5 その他の機器", 2)
    _p(d, "燃焼ストーカ（燃焼段 30 枚交換）、炉内耐火物（左側壁 1.6m²・右側壁 0.9m²）、スートブロワ SB-3・4、洗煙塔内部、触媒反応塔、"
          "給じん装置、押込・二次送風機、電気計装は 1 章の表のとおり。測定記録は付属資料（肉厚測定表・校正記録）に綴じた。", 9)
    d.add_heading("5. 交換部品・支給品", 1)
    _table(d, ["品名", "仕様", "数量", "区分", "使用先・扱い"], [
        ("ろ布", "PPS/PTFE φ150×6,000", "300 本受入、288 本使用", "市支給品", "2GT-014。残 12 本は予備品として市倉庫へ返納（11/6）"),
        ("ケージ", "SUS304 φ145×6,000", "5 本", "市予備品", "2GT-014。曲がり品と交換"),
        ("過熱器管", "STB410 φ38.1×4.5t×2.4m", "3 本", "当社手配（追加）", "2BL-002"),
        ("火格子", "燃焼段用 SCH13", "30 枚", "当社手配", "2IN-005"),
        ("キャスタブル", "高アルミナ質", "1.2 t", "当社手配", "2IN-011"),
        ("軸受", "22324 CAME4", "2 個", "当社手配", "2FD-006"),
        ("羽根車・ライナ", "洗煙循環ポンプ用 SCS14", "各 1", "当社手配", "2GT-036"),
        ("メカニカルシール", "同上", "1 組", "当社手配", "2GT-036"),
        ("SB パッキン・軸受", "SB-3・4 用", "1 式", "当社手配", "2BL-006・007")], [26, 38, 28, 24, 49], 7.5)
    d.add_heading("6. 所見", 1)
    _p(d, "(1) 1次過熱器管は全体に減肉が進んでいる。今回の 3 本は管理値を下回ったため取り替えたが、他の管も 2〜3 年で管理値に達する見込みである。"
          "2027 年度の定期整備で全数取替とすることを推奨する。", 9.5)
    _p(d, "(2) 誘引通風機の羽根車は肉盛で回復したが、母材の減肉もあり、2029 年頃の羽根車更新が望ましい。年 1 回の振動診断を継続されたい。", 9.5)
    _p(d, "(3) 取り外したろ布 288 本は産業廃棄物として 11/5 に搬出した（マニフェスト写し 添付）。", 9.5)
    _p(d, "以上", 10, align="r")
    dx = os.path.join(tmp, os.path.splitext(os.path.basename(path_pdf))[0] + ".docx")
    d.save(dx)
    return _pdf(dx, path_pdf)


def make_new_c(path_pdf, tmp):
    ct = CONTRACTS["R6-C"]
    no = "SH-KN-2024-04"
    d = _doc("【架空】令和6年度 くぬぎ台環境センター 全炉停止時共通設備整備業務委託 定期整備報告書", "汐見重工環境 r2", "工事報告書")
    _cover(d, [("令和6年度　くぬぎ台環境センター", 15), ("全炉停止時共通設備整備業務委託", 17), ("定期整備報告書", 22)], no)
    _p(d, "2024年11月", 12, align="c")
    _p(d, "汐見重工環境株式会社", 13, True, "c")
    d.add_page_break()
    body()
    _header(d, "【架空】%s　%s" % (ct["name"], no))
    _footer(d, "汐見重工環境株式会社")
    d.add_heading("1. 報告サマリー", 1)
    _p(d, "工期 2024/10/9〜11/1。全炉停止 10/16〜10/23、受電停止 10/19。2号炉定期整備（別契約）と同じ期間のため、構内のクレーン・足場を調整して実施した。", 9)
    _table(d, ["No", "対象機器（タグ）", "実施作業", "結果", "判定", "次回推奨"], [
        ("1", "蒸気タービン（0ST-001）・減速機・発電機", "開放点検（定期事業者検査）", "動翼 第8・9段 先端エロージョン 軽微、MT 指示なし", "○", "4 年後 開放"),
        ("2", "脱気器（0ST-031）", "開放点検、第一種圧力容器 性能検査", "10/22 合格", "◎", "1 年後"),
        ("3", "空冷復水器ファン No.1〜6（0ST-013〜018）", "減速機 油交換、羽根点検", "No.5 付着物 清掃後 振動 良", "○", "－"),
        ("4", "ごみクレーン 2号（0RF-013）", "巻上ワイヤロープ交換、ブレーキ整備", "荷重試験 良", "◎", "3 年後"),
        ("5", "受電盤・連系盤・高圧配電盤 No.1（0EL-001・002・008）", "精密点検（受電停止時）", "保護継電器 動作 良", "◎", "－"),
        ("6", "煙突外筒・内筒（0MS-001 他）", "点検、内筒 肉厚測定 18 点", "最小 4.1mm（設計 6.0）", "○", "5 年後"),
        ("7", "冷却塔 A・B（0UT-005・006）", "清掃・点検", "散水板 一部割れ（交換は次回）", "△", "次回 散水板交換"),
    ], [8, 44, 38, 40, 9, 24], 7.5)
    d.add_heading("2. 主な実施内容", 1)
    _table(d, ["区分", "内容"], [
        ("タービン", "上半ケーシング開放（10/16）、動翼・静翼の外観・MT、軸受・ジャーナル点検、組立（10/22〜）、10/29 無負荷試運転・10/31 負荷試験。"),
        ("脱気器", "マンホール開放・内部清掃、スプレーノズル点検、10/22 登録性能検査機関による性能検査 合格、ガスケット交換。"),
        ("受変電", "10/19 8:30〜16:40 受電停止。VCB 開閉特性、保護継電器試験、絶縁抵抗測定。非常用発電機で保安電源を確保（市 電気保安担当 立会）。")],
        [22, 143], 8.5, "EDEDED")
    _photos(d, ["タービン 動翼（開放時）", "脱気器 内部"], tmp, 300)
    d.add_heading("3. 交換部品", 1)
    _table(d, ["品名", "仕様", "数量", "区分", "使用先"], [
        ("ワイヤロープ", "IWRC 6×Fi(29) φ22 ×2 本", "1 式", "当社手配", "0RF-013"),
        ("脱気器 マンホールガスケット", "うず巻形", "2 枚", "当社手配", "0ST-031"),
        ("減速機油", "ISO VG220", "120 L", "当社手配", "0ST-013〜018"),
        ("タービン 軸受パッド", "－", "使用せず（点検の結果 良）", "予備品 未使用", "市倉庫に保管のまま")], [36, 44, 28, 24, 33], 8)
    d.add_heading("4. 所見", 1)
    _p(d, "タービン動翼の先端エロージョンは軽微で、次回開放（2028 年度）で進行を確認されたい。冷却塔 A の散水板割れは次回の共通整備で交換を推奨する。", 9.5)
    _p(d, "以上", 10, align="r")
    dx = os.path.join(tmp, os.path.splitext(os.path.basename(path_pdf))[0] + ".docx")
    d.save(dx)
    return _pdf(dx, path_pdf)
