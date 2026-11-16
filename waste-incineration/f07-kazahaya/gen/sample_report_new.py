# -*- coding: utf-8 -*-
"""工事報告書（新）：sp2 が 2021 年度に作った SPC 独自様式（KZ-H-01）。受注者が記入し SPC へ提出する。
1 枚目＝工事報告書（サマリー帳票）、2 枚目以降＝続紙。reportlab で文字のある PDF を直接作る。"""
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (BaseDocTemplate, PageTemplate, Frame, Table, TableStyle, Paragraph, Spacer, PageBreak,
                                KeepTogether)
from reportlab.lib.styles import ParagraphStyle
from contracts import *
from common import *
from world import BY_CODE

pdfmetrics.registerFont(TTFont("IPAG", "/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf"))
pdfmetrics.registerFont(TTFont("IPAPG", "/usr/share/fonts/opentype/ipafont-gothic/ipagp.ttf"))

ST = ParagraphStyle("b", fontName="IPAPG", fontSize=8.5, leading=11)
STS = ParagraphStyle("s", fontName="IPAPG", fontSize=7.5, leading=9.5)
STH = ParagraphStyle("h", fontName="IPAG", fontSize=11, leading=14, spaceBefore=4, spaceAfter=3)
STT = ParagraphStyle("t", fontName="IPAG", fontSize=16, leading=20, alignment=1)

GRID = [("FONTNAME", (0, 0), (-1, -1), "IPAPG"), ("FONTSIZE", (0, 0), (-1, -1), 8), ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("TOPPADDING", (0, 0), (-1, -1), 1.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5)]
HEAD = colors.HexColor("#DDEBF7")


def P(t, st=ST):
    return Paragraph(str(t).replace("\n", "<br/>"), st)


class Doc(BaseDocTemplate):
    def __init__(self, path, meta, **kw):
        super().__init__(path, pagesize=A4, leftMargin=14 * mm, rightMargin=14 * mm, topMargin=20 * mm, bottomMargin=14 * mm,
                         title=meta["title"], author=meta["author"], subject=meta["subject"], creator="【架空】SPC 報告書様式 KZ-H-01")
        self.meta = meta
        fr = Frame(self.leftMargin, self.bottomMargin, self.width, self.height, id="f")
        self.addPageTemplates([PageTemplate(id="p", frames=[fr], onPage=self._deco)])

    def _deco(self, cv, doc):
        cv.saveState()
        cv.setFont("IPAPG", 7.5)
        pg = cv.getPageNumber()
        label = "様式 KZ-H-01（工事報告書）" if pg == 1 else f"様式 KZ-H-01（続紙 {pg - 1}）"
        cv.drawString(14 * mm, A4[1] - 11 * mm, f"【架空】かざはや環境パートナーズ株式会社　{label}")
        cv.drawRightString(A4[0] - 14 * mm, A4[1] - 11 * mm, f"報告書番号 {self.meta['no']}　{pg} / {self.meta['pages']}")
        cv.line(14 * mm, A4[1] - 12.5 * mm, A4[0] - 14 * mm, A4[1] - 12.5 * mm)
        cv.drawCentredString(A4[0] / 2, 8 * mm, self.meta["footer"])
        cv.restoreState()


def summary_table(rows, first_page_rows):
    head = ["保全ｺｰﾄﾞ", "機器名（PM-NAVI）", "区分", "実施内容", "結果", "PM-NAVI"]
    data = [[P(h, STS) for h in head]] + [[P(c, STS) for c in r] for r in rows]
    t = Table(data, colWidths=[19 * mm, 37 * mm, 11 * mm, 71 * mm, 32 * mm, 12 * mm], repeatRows=1)
    t.setStyle(TableStyle(GRID + [("BACKGROUND", (0, 0), (-1, 0), HEAD)]))
    return t


def kv_table(items, w1=28, total=182):
    data = [[P(k, STS), P(v)] for k, v in items]
    t = Table(data, colWidths=[w1 * mm, (total - w1) * mm])
    t.setStyle(TableStyle(GRID + [("BACKGROUND", (0, 0), (0, -1), HEAD)]))
    return t


def stamp():
    t = Table([["提出", "SPC受付", "SPC保全", "所長"], ["", "", "", ""]], colWidths=[18 * mm] * 4, rowHeights=[5 * mm, 14 * mm])
    t.setStyle(TableStyle(GRID + [("ALIGN", (0, 0), (-1, -1), "CENTER"), ("FONTSIZE", (0, 0), (-1, -1), 7)]))
    return t


# ================================================================ 2023 1 号炉定期整備工事
SUM_C23 = [
    ("1NS-030", "燃焼火格子", "更新", "燃焼段 可動96・固定96 計192枚を全数更新（SPC支給品）。取付後クリアランス3列手直し", "良（クリアランス2～5mm内）", "済"),
    ("1NS-020", "乾燥火格子", "点検", "摩耗測定（各列3点）", "最大 9.6mm／管理15mm 要観察", "済"),
    ("1NS-040", "後燃焼火格子", "点検", "摩耗測定（各列3点）", "最大 6.2mm 良", "済"),
    ("1NS-010", "焼却炉本体", "補修", "炉内耐火物点検、乾燥段側壁1.2m²＋燃焼段側壁下部2.4m²（追加）補修", "良（追加補修 変更契約）", "済"),
    ("1KT-100", "誘引通風機", "整備", "軸受交換（DE・NDE）、インペラ前縁肉盛補修、動バランス修正 G2.5", "振動 5.8→2.1mm/s 良", "済"),
    ("1KT-100M", "誘引通風機 電動機", "整備", "軸受交換、精密点検、絶縁診断（PI 3.8、tanδ 1.2%）", "良", "済"),
    ("1BO-010", "ボイラ本体", "法定", "ボイラ水圧試験（10/25）、定期事業者検査 受検（10/26）", "合格", "済"),
    ("1BO-020", "蒸気ドラム", "点検", "肉厚測定", "良", "済"),
    ("1BO-030", "水管壁（炉頂・放射室）", "点検", "肉厚測定（放射室下部 肉盛部含む）", "最小 4.9mm／限界4.0 要観察", "済"),
    ("1BO-040", "蒸発管群", "点検", "肉厚測定", "良", "済"),
    ("1BO-050", "1次過熱器", "点検", "肉厚測定", "良", "済"),
    ("1BO-060", "2次過熱器", "点検", "肉厚測定", "良", "済"),
    ("1BO-070", "3次過熱器", "点検", "肉厚測定（下段管 重点）", "最小 5.6mm／限界4.5 要観察", "済"),
    ("1BO-090", "エコノマイザ", "点検", "肉厚測定", "良", "済"),
    ("1BO-100", "ボイラ給水ポンプA", "整備", "分解整備：メカニカルシール・軸受・ライナリング交換（SPC支給）", "良（シール漏れ解消）", "済"),
    ("1BO-160", "長抜差型スートブロワ1", "整備", "ランス・グランドパッキン交換", "良", "済"),
    ("1BO-320", "ボイラダスト排出装置1", "整備", "分解整備（ロータ・シール）", "良", "済"),
    ("1BO-340", "ボイラダスト排出装置3", "整備", "分解整備（ロータ・シール）", "良", "済"),
    ("1HG-010", "減温塔", "点検", "内面点検・肉厚測定", "良", "済"),
    ("1KT-160", "煙突内筒", "点検", "内面点検・肉厚測定", "良", "済"),
    ("1HG-130", "バグフィルタ本体", "点検", "ろ布抜取 6本 → 強度・通気度試験", "保持率 81～86%（11/17 結果）", "済"),
    ("1HG-140", "バグフィルタ逆洗装置", "整備", "電磁弁ダイヤフラム 120個交換", "良", "済"),
    ("1HG-170", "BF室別入口ダンパ駆動装置3", "整備", "分解整備（シリンダパッキン）", "良", "済"),
    ("1HG-200", "BF室別入口ダンパ駆動装置6", "整備", "分解整備（シリンダパッキン）", "良", "済"),
    ("1HG-240", "BF飛灰排出装置2", "整備", "分解整備", "良", "済"),
    ("1HG-380", "触媒反応塔", "点検", "触媒抜取 2本 → 性能試験", "脱硝率 1層 92%/2層 89%（11/17）", "済"),
    ("1HG-400", "アンモニア水注入ポンプA", "整備", "ダイヤフラム・逆止弁交換", "良", "済"),
    ("1HG-030M", "減温塔噴霧水ポンプB 電動機", "整備", "軸受交換", "良", "済"),
    ("1KT-020", "押込送風機ダンパ駆動装置", "整備", "分解整備", "良", "済"),
    ("1KT-060", "区分空気ダンパ駆動装置（後燃焼）", "整備", "分解整備", "良", "済"),
    ("1NS-130", "燃焼装置油圧ポンプB", "整備", "分解整備（シール・軸受）", "良", "済"),
    ("1NS-130M", "燃焼装置油圧ポンプB 電動機", "整備", "軸受交換", "良", "済"),
    ("1NS-210", "灰押出油圧ポンプB", "整備", "分解整備", "良", "済"),
    ("1NS-210M", "灰押出油圧ポンプB 電動機", "整備", "軸受交換", "良", "済"),
    ("1NS-330", "炉出口温度計1", "更新", "R熱電対交換", "良", "済"),
    ("1NS-340", "炉出口温度計2", "更新", "R熱電対交換", "良", "済"),
    ("1NS-360", "二次燃焼室ガス温度計", "更新", "R熱電対交換", "良", "済"),
    ("1HG-470", "水銀連続分析計", "更新", "新機（SPC支給）据付・配管接続・校正（校正ガス不良で10/25再校正）", "良", "済"),
]


def c23_pages():
    c = C23
    el = []
    el.append(P("【架空】工 事 報 告 書", STT))
    el.append(Spacer(1, 3 * mm))
    top = Table([[kv_table([
        ("工事件名", "【架空】" + c["name"]),
        ("契約区分", "定期整備（請負）　SPC発注 契約番号 KZ-C-2023-021（変更1回）"),
        ("工期（契約）", "2023年10月2日 ～ 2023年11月2日"),
        ("工期（実績）", "2023年10月1日 炉立下げ ～ 2023年11月1日 ごみ投入・定格到達、11月2日 完了確認"),
        ("受注者", "【架空】水樹テクノサービス株式会社　職長 q2"),
        ("SPC担当", "sp2（保全）　提出日 2023年11月24日"),
    ], 24, 110), stamp()]], colWidths=[112 * mm, 72 * mm])
    top.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0)]))
    el.append(top)
    el.append(Spacer(1, 2 * mm))
    el.append(P(f"1. 対象機器と実施内容（サマリー）　※全{len(SUM_C23)}件。1枚目に載らない分は続紙へ続く", STH))
    el.append(summary_table(SUM_C23, 20))
    el.append(Spacer(1, 2 * mm))
    el.append(P("2. 変更・追加", STH))
    el.append(kv_table([
        ("変更1", "10/11 燃焼段 左右側壁下部（火格子上0～400mm）に耐火物の剥離・欠損 計2.4m²（最大深さ85mm）を発見。"
                  "10/12 協議のうえ追加補修を承認（変更契約 2023/10/20、工期変更なし、耐火 2名増員）。"),
        ("手直し", "10/19 燃焼火格子 取付後のクリアランス測定で 6・7・8列目が基準（2～5mm）を超過（最大7.5mm）。"
                   "固定火格子受け金具の締付順による位置ずれ。10/19～20 再調整し全列 2.5～4.5mm。"),
        ("工程", "火格子取付 2日、耐火補修 4日、水銀計 1日 予定より延びたが、足場解体以降は予定どおり。"),
    ], 18))
    el.append(Spacer(1, 2 * mm))
    el.append(P("3. 次回への提言", STH))
    el.append(kv_table([
        ("乾燥火格子", "最大9.6mm。現在の摩耗速度（約1.1mm/年）では長期修繕計画の更新年 2027 年度で妥当。2025 年度に再測定。"),
        ("3次過熱器", "下段管 最小5.6mm（2021：6.3mm）。減肉速度 約0.35mm/年。2027 年度の下段管部分更新を推奨（計画どおり）。"),
        ("誘引通風機", "インペラ前縁の肉盛は今回で2回目。次回は肉盛でなくインペラ更新とする（計画 2029 年度）。"),
        ("耐火物", "燃焼段側壁下部は今回初めて損傷。2024 年度定修で同部位を重点点検。"),
    ], 18))
    return el


def c23_cont():
    el = [PageBreak()]
    el.append(P("4. 測定結果", STH))
    el.append(P("4.1 燃焼火格子 取付後クリアランス（可動／固定間、単位 mm、基準 2～5）", ST))
    rows = [["列", "左端", "1/4", "中央", "3/4", "右端", "判定", "備考"]]
    import zlib
    for col in range(1, 9):
        vals = [round(2.6 + (zlib.crc32(f"cl{col}{k}".encode()) % 17) / 10, 1) for k in range(5)]
        note = ""
        if col in (6, 7, 8):
            note = "手直し前 最大" + {6: "7.5", 7: "6.8", 8: "6.1"}[col]
        rows.append([f"{col}列"] + vals + ["良", note])
    t = Table(rows, colWidths=[12 * mm] + [16 * mm] * 5 + [14 * mm, 52 * mm], repeatRows=1)
    t.setStyle(TableStyle(GRID + [("BACKGROUND", (0, 0), (-1, 0), HEAD), ("ALIGN", (1, 1), (6, -1), "CENTER")]))
    el.append(t)
    el.append(Spacer(1, 3 * mm))
    el.append(P("4.2 ボイラ肉厚測定（超音波厚さ計、【架空】かざはや非破壊検査、単位 mm）", ST))
    rows = [["保全ｺｰﾄﾞ", "部位", "公称", "限界", "測定点数", "最小", "前回(2021)最小", "判定"],
            ["1BO-030", "水管壁 放射室下部（肉盛部）", "5.0+肉盛", "4.0", "120", "4.9", "5.2", "要観察"],
            ["1BO-030", "水管壁 炉頂", "5.0", "4.0", "60", "4.8", "4.9", "良"],
            ["1BO-040", "蒸発管群", "4.5", "3.5", "80", "4.3", "4.3", "良"],
            ["1BO-050", "1次過熱器", "5.0", "4.5", "60", "4.9", "4.9", "良"],
            ["1BO-060", "2次過熱器", "6.0", "4.5", "60", "5.8", "5.9", "良"],
            ["1BO-070", "3次過熱器 下段管", "7.1", "4.5", "72", "5.6", "6.3", "要観察"],
            ["1BO-070", "3次過熱器 上段管", "7.1", "4.5", "48", "6.8", "6.9", "良"],
            ["1BO-090", "エコノマイザ", "4.0", "3.0", "60", "3.9", "3.9", "良"],
            ["1BO-020", "蒸気ドラム 胴板", "32.0", "28.0", "12", "31.8", "31.8", "良"]]
    t = Table(rows, colWidths=[17 * mm, 48 * mm, 18 * mm, 14 * mm, 18 * mm, 14 * mm, 26 * mm, 18 * mm], repeatRows=1)
    t.setStyle(TableStyle(GRID + [("BACKGROUND", (0, 0), (-1, 0), HEAD), ("ALIGN", (2, 1), (-1, -1), "CENTER")]))
    el.append(t)
    el.append(Spacer(1, 3 * mm))
    el.append(P("4.3 回転機の振動・絶縁（据付後・立上げ時）", ST))
    rows = [["保全ｺｰﾄﾞ", "機器", "項目", "整備前", "整備後", "判定基準", "判定"],
            ["1KT-100", "誘引通風機", "軸受振動（速度 rms）", "5.8 mm/s", "2.1 mm/s（10/31 負荷運転）", "4.5 以下", "良"],
            ["1KT-100M", "誘引通風機 電動機", "成極指数 PI／tanδ", "—", "3.8／1.2%", "2.0以上／5%以下", "良"],
            ["1BO-100", "ボイラ給水ポンプA", "軸受振動", "3.9 mm/s", "1.6 mm/s", "4.5 以下", "良"],
            ["1NS-130", "燃焼装置油圧ポンプB", "吐出圧力", "13.2 MPa", "14.0 MPa", "14.0±0.5", "良"]]
    t = Table(rows, colWidths=[17 * mm, 32 * mm, 36 * mm, 20 * mm, 40 * mm, 22 * mm, 14 * mm], repeatRows=1)
    t.setStyle(TableStyle(GRID + [("BACKGROUND", (0, 0), (-1, 0), HEAD)]))
    el.append(t)
    el.append(Spacer(1, 3 * mm))
    el.append(P("4.4 試験（委託）　ろ布：引張強度保持率 81～86%（交換目安 70%）、通気度 良。触媒：脱硝率 1層 92%・2層 89%（初期 95%）。"
                "予備層充填は計画どおり 2025 年度でよい。（【架空】かざはや分析センター 2023/11/17 報告）", ST))
    el.append(PageBreak())
    el.append(P("5. 使用部品・材料", STH))
    rows = [["品名", "仕様・型式", "手配", "数量", "使用", "予備・残", "保管・扱い"],
            ["燃焼火格子（可動・固定）", "高クロム鋳鋼 MZ-G150-C", "SPC支給", "200枚", "192枚", "8枚", "SPC予備品倉庫（No.3棚）"],
            ["火格子取付ボルト・ナット", "SUS310S M20", "受注者", "420組", "384組", "36組", "受注者持帰り"],
            ["水銀連続分析計", "【架空】青嵐計器 SR-HG7", "SPC支給", "1台", "1台", "—", "旧機はSPCで廃棄"],
            ["誘引通風機 軸受", "22228 CCK／W33 ほか", "受注者", "1組", "1組", "—", "※SPC予備 1組は別途購入・保管"],
            ["インペラ肉盛溶接材", "硬化肉盛 HF-600 φ3.2", "受注者", "35kg", "31kg", "4kg", "受注者持帰り"],
            ["ボイラ給水ポンプ メカニカルシール", "【架空】千鳥ポンプ CP 純正", "SPC支給", "2組", "1組", "1組", "SPC予備品倉庫"],
            ["同 軸受・ライナリング", "純正", "SPC支給", "1式", "1式", "—", ""],
            ["キャスタブル（高アルミナ質）", "東雲築炉 標準品", "受注者（協力会社）", "当初1.5t＋追加1.2t", "2.6t", "0.1t", "持帰り"],
            ["BF電磁弁ダイヤフラム", "1-1/2\" 用", "受注者", "120個", "120個", "—", ""],
            ["R熱電対", "φ8 L=1,500 保護管付", "受注者", "3本", "3本", "—", ""],
            ["油圧ポンプ シールキット", "純正", "受注者", "2式", "2式", "—", ""],
            ["電動機軸受（油圧ポンプB×2・噴霧水ポンプB）", "6312ZZ ほか", "受注者", "3台分", "3台分", "—", ""]]
    t = Table([[P(c, STS) for c in r] for r in rows], colWidths=[40 * mm, 34 * mm, 20 * mm, 22 * mm, 14 * mm, 16 * mm, 36 * mm],
              repeatRows=1)
    t.setStyle(TableStyle(GRID + [("BACKGROUND", (0, 0), (-1, 0), HEAD)]))
    el.append(t)
    el.append(P("※本様式には金額を記載しない（契約・精算は別途）。", STS))
    el.append(Spacer(1, 3 * mm))
    el.append(P("6. 写真（写真代替図）", STH))
    cap = ["燃焼段 撤去後（受台）", "新品火格子 取付状況", "側壁下部 損傷部（追加補修前）", "同 補修後", "IDF インペラ 肉盛後", "水銀計 新機"]
    cells = [[P("【架空】見本用の写真代替図<br/><br/><br/><br/><br/>" + f"写真{i + 1} {cap[i]}", STS) for i in range(k, k + 3)] for k in (0, 3)]
    t = Table(cells, colWidths=[60 * mm] * 3, rowHeights=[38 * mm] * 2)
    t.setStyle(TableStyle(GRID + [("ALIGN", (0, 0), (-1, -1), "CENTER")]))
    el.append(t)
    return el


def make_c23():
    stem = f"{FN}_工事報告書_2023年度1号炉定期整備工事_SPC様式"
    meta = dict(title="【架空】2023年度 かざはやエコプラザ 1号炉定期整備工事 工事報告書", author="【架空】水樹テクノサービス q2（SPC様式 sp2）",
                subject="【架空】工事報告書 様式KZ-H-01", no="KZ-2023-H-014", pages=4,
                footer="【架空】かざはやエコプラザ　2023年度 1号炉定期整備工事")
    path = outpath(stem + ".pdf")
    for _ in range(2):  # 2 回組んで総ページ数を入れる
        Doc(path, meta).build(c23_pages() + c23_cont())
        meta["pages"] = pdf_pages(path)
    return stem + ".pdf"


# ================================================================ 2023 全炉停止時共通設備点検整備工事
SUM_C23C = [
    ("0DE-010", "受電盤（特別高圧）", "法定", "年次点検（GIS ガス圧・開閉試験・保護継電器試験）", "良", "済"),
    ("0DE-020", "主変圧器", "法定", "年次点検、油中ガス分析（採油）", "良（ガス分析 正常）", "済"),
    ("0DE-030", "系統連系保護盤", "法定", "年次点検・継電器特性試験", "良", "済"),
    ("0DE-040～110", "高圧配電盤（8面）", "法定", "年次点検・VCB 開閉試験・絶縁抵抗", "良", "済"),
    ("0DE-130～160", "所内変圧器（4台）", "法定", "年次点検", "良", "済"),
    ("0YN-180", "発電機盤", "法定", "年次点検", "良", "済"),
    ("0DE-270", "非常用発電機", "法定", "年次点検・負荷試験（実負荷30分）", "良", "済"),
    ("0YN-030", "タービン主油ポンプ", "整備", "分解整備（歯車・軸受点検、ブッシュ交換）", "良", "済"),
    ("0YN-210", "空冷式復水器", "点検", "フィン洗浄（高圧水）・目視点検", "良", "済"),
    ("0YN-250", "空冷復水器ファン4", "整備", "分解整備（ハブ・羽根ピッチ確認）", "良", "済"),
    ("0YN-300", "空冷復水器ファン減速機3", "整備", "分解点検、出力側軸受交換、芯出し再調整（10/21）", "要観察（歯面 軽微ピッチング）", "済"),
    ("0YN-400", "脱気器", "法定", "開放点検、性能検査 受検（10/19）", "合格", "済"),
    ("0YN-190", "高圧蒸気だめ", "法定", "定期自主検査", "良", "済"),
    ("0YN-200", "低圧蒸気だめ", "法定", "定期自主検査", "良", "済"),
    ("0YN-350M", "復水ポンプA 電動機", "整備", "軸受交換", "良", "済"),
    ("0YN-420", "脱気器給水ポンプB", "整備", "分解整備", "良", "済"),
    ("0KY-060", "機器冷却水ポンプB", "整備", "分解整備", "良", "済"),
]


def make_c23c():
    c = C23C
    stem = f"{FN}_工事報告書_2023年度全炉停止時共通設備点検整備工事_SPC様式"
    meta = dict(title="【架空】2023年度 かざはやエコプラザ 全炉停止時共通設備点検整備工事 工事報告書",
                author="【架空】水樹テクノサービス q2（SPC様式 sp2）", subject="【架空】工事報告書 様式KZ-H-01", no="KZ-2023-H-015",
                pages=2, footer="【架空】かざはやエコプラザ　2023年度 全炉停止時共通設備点検整備工事")
    path = outpath(stem + ".pdf")
    el = [P("【架空】工 事 報 告 書", STT), Spacer(1, 3 * mm)]
    top = Table([[kv_table([
        ("工事件名", "【架空】" + c["name"]),
        ("契約区分", "定期整備（請負）　SPC発注 契約番号 KZ-C-2023-022"),
        ("工期（契約）", "2023年10月14日 ～ 2023年10月22日"),
        ("工期（実績）", "10月13日 21:00 2号炉立下げ（全炉停止）～ 10月22日 2号炉立上げ、10月23日 タービン併入"),
        ("受注者", "【架空】水樹テクノサービス株式会社　職長 q2（日報 q3）"),
        ("SPC担当", "sp2（保全）　提出日 2023年11月10日"),
    ], 24, 110), stamp()]], colWidths=[112 * mm, 72 * mm])
    top.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0)]))
    el.append(top)
    el.append(Spacer(1, 2 * mm))
    el.append(P("1. 対象機器と実施内容（サマリー）　※1号炉定期整備工事（別契約 KZ-C-2023-021）の範囲は含まない", STH))
    el.append(summary_table(SUM_C23C, 17))
    el.append(Spacer(1, 2 * mm))
    el.append(P("2. 変更・追加", STH))
    el.append(kv_table([("変更", "なし"),
                        ("手直し", "10/20 減速機3 組立後の軸端振れ 0.08mm（基準0.05mm以下）。10/21 芯出し再調整し 0.03mm。工期内。")], 18))
    el.append(Spacer(1, 2 * mm))
    el.append(P("3. 次回への提言", STH))
    el.append(kv_table([
        ("ACC減速機3", "低速段歯面に軽微なピッチング（面積率 約3%）。異音の主因は出力側軸受の外輪剥離で、軸受交換で解消。"
                       "歯面は2025年度全停で再点検し、進行していれば歯車交換（予備歯車の手配 納期約8か月）。"),
        ("ACC減速機（他5台）", "同型のため、2025・2027年度の分解整備を前倒しする場合は振動傾向で判断。"),
    ], 22))
    el.append(PageBreak())
    el.append(P("4. 測定結果", STH))
    rows = [["保全ｺｰﾄﾞ", "機器", "項目", "結果", "基準", "判定"],
            ["0YN-300", "ACCファン減速機3", "出力軸 振れ（再調整後）", "0.03 mm", "0.05 以下", "良"],
            ["0YN-300", "同", "振動（速度 rms、10/23 運転）", "1.9 mm/s（整備前 6.4）", "4.5 以下", "良"],
            ["0DE-020", "主変圧器", "油中ガス（可燃性ガス総量）", "185 ppm", "要注意 1,000", "良"],
            ["0DE-270", "非常用発電機", "負荷試験 電圧・周波数", "6.6kV／50.0Hz 安定", "—", "良"],
            ["0YN-400", "脱気器", "胴板 肉厚 最小", "11.7 mm（公称12）", "—", "合格"]]
    t = Table([[P(c, STS) for c in r] for r in rows], colWidths=[18 * mm, 36 * mm, 46 * mm, 38 * mm, 26 * mm, 16 * mm], repeatRows=1)
    t.setStyle(TableStyle(GRID + [("BACKGROUND", (0, 0), (-1, 0), HEAD)]))
    el.append(t)
    el.append(Spacer(1, 3 * mm))
    el.append(P("5. 使用部品・材料", STH))
    rows = [["品名", "仕様", "手配", "数量", "使用", "予備・残", "扱い"],
            ["ACC減速機 軸受一式（1台分）", "【架空】水樹プラント 純正", "SPC支給", "2台分", "1台分", "1台分", "SPC予備品倉庫"],
            ["減速機 潤滑油", "ISO VG220", "受注者", "60L", "48L", "12L", "持帰り"],
            ["主油ポンプ ブッシュ", "純正", "受注者", "2個", "2個", "—", ""],
            ["ポンプ用メカシール・軸受", "純正（脱気器給水B・冷却水B）", "受注者", "2式", "2式", "—", ""],
            ["電動機軸受（復水ポンプA）", "6310ZZ", "受注者", "1台分", "1台分", "—", ""],
            ["脱気器 マンホールガスケット", "うず巻形", "受注者", "2枚", "2枚", "—", ""]]
    t = Table([[P(c, STS) for c in r] for r in rows], colWidths=[44 * mm, 40 * mm, 18 * mm, 16 * mm, 16 * mm, 16 * mm, 32 * mm], repeatRows=1)
    t.setStyle(TableStyle(GRID + [("BACKGROUND", (0, 0), (-1, 0), HEAD)]))
    el.append(t)
    el.append(P("※本様式には金額を記載しない。受変電設備 年次点検は協力会社【架空】風早電設が実施（点検記録は別冊）。", STS))
    el.append(Spacer(1, 3 * mm))
    el.append(P("6. 写真（写真代替図）", STH))
    cap = ["減速機3 低速段歯面", "出力側軸受 外輪剥離", "脱気器 内部"]
    t = Table([[P("【架空】見本用の写真代替図<br/><br/><br/><br/>" + f"写真{i + 1} {cap[i]}", STS) for i in range(3)]],
              colWidths=[60 * mm] * 3, rowHeights=[34 * mm])
    t.setStyle(TableStyle(GRID + [("ALIGN", (0, 0), (-1, -1), "CENTER")]))
    el.append(t)
    import copy
    for _ in range(2):
        Doc(path, meta).build(copy.deepcopy(el))
        meta["pages"] = pdf_pages(path)
    return stem + ".pdf"


if __name__ == "__main__":
    print(make_c23())
    print(make_c23c())
