# -*- coding: utf-8 -*-
"""工事報告書（全年度、文字のある PDF）
 sv1 文章型（docx→PDF）：2009–2017（2017 は sv2 が sv1 の様式で作成）
 sv2 2018 様式（reportlab）：2018–
 メーカー出来形・写真帳：基幹改良 R7（R8 は施工中で未作成）
"""
import os, subprocess, datetime as dt, random
import world as W, contracts as C
from rec import Rec

def wid(c, t):
    return [f"{c.cid}:{t.tid}:{x}" for x in t.tags] if t.tags else [f"{c.cid}:{t.tid}"]

def jdate(d):
    if d.year >= 2019 and d >= dt.date(2019, 5, 1):
        return f"令和{'元' if d.year == 2019 else d.year - 2018}年{d.month}月{d.day}日"
    return f"平成{d.year - 1988}年{d.month}月{d.day}日"

K1_TEXTS = [
    ("T02", "(1) ガス化炉", "流動砂約 6m³ を抜き出して炉内を清掃し、散気ノズル 96 本の摩耗量を測定した。交換基準を超えた 24 本を新品に交換し、取付部を溶接した。炉内耐火物に有害な損傷はなかった。閉止後、珪砂 6 号を約 6m³ 投入した。"),
    ("T03", "(2) 旋回溶融炉", "炉内点検の結果、出滓口まわり（約 1.4m²）と炉頂部（約 0.8m²）の耐火物に溶損があり、はつりのうえ出滓口をれんが積み、炉頂部をキャスタブルで補修した。炉頂部のスタッド 18 本を植え替えた。鉄皮に異常はなかった。側壁は今回補修範囲外としたが、下段の減りが進んでおり次回以降の課題とする。"),
    ("T04", "(3) ボイラ", "水管・過熱器・エコノマイザのダストを除去し、肉厚を測定した（水管 120 点、過熱器 64 点、エコノマイザ 48 点）。二次過熱器下部ベンド No.14 で 2.6mm（新製時 4.0mm）を確認したため、組合と協議のうえ肉盛溶接を追加で行い、PT・UT で欠陥のないことを確認した。10 月 14 日に定期事業者検査（内部検査）を行い、10 月 21 日に水圧試験（4.6MPa、30 分保持）で漏れのないことを確認した。"),
    ("T05", "(4) スートブロワ", "長抜差式 LR-3、LR-4 を取り外して分解し、ランス管とグランドパッキンを交換した。"),
    ("T06", "(5) 不燃物抜出装置A", "スクリューを引き抜いて摩耗を測定し（外径 φ350 → φ331）、工場で羽根の硬化肉盛を行った。軸受・軸封を交換して組み込み、芯出しを行った。"),
    ("T07", "(6) 誘引通風機", "軸受（負荷側・反負荷側）を交換し、羽根車のダストを除去した。芯出し後の振動は 2.1mm/s であった。"),
    ("T08", "(7) 流動空気押込送風機", "軸受・羽根車を点検し、グリスを交換した。異常なし。"),
    ("T09", "(8) バグフィルタ", "6 室のろ布を目視点検し、試験用に 3 本を抜き取って予備ろ布を取り付けた。強度試験は外部機関に依頼した（結果は別紙）。"),
    ("T10", "(9) 減温塔", "噴霧ノズル 6 本を交換した。"),
    ("T11", "(10) 触媒反応塔", "第 1 層から触媒サンプルを採取した（性能試験は外部機関に依頼）。"),
    ("T12", "(11) スラグ水砕コンベヤ", "チェーン伸びは 1.1% で交換基準（2%）未満。摩耗したスクレーパ 8 枚を交換した。"),
]
K1_PARTS = [("ガス化炉 散気ノズル", "SCS13", "28 本", "24 本", "4 本"), ("減温塔 噴霧ノズル", "二流体", "8 個", "6 個", "2 個"),
            ("スートブロワ ランス管", "STBA24 φ60.5", "2 本", "2 本", "－"), ("グランドパッキン", "", "4 組", "2 組", "2 組"),
            ("スクレーパ", "W600 耐摩耗鋼", "10 枚", "8 枚", "2 枚"), ("不燃物抜出装置 軸受", "22220", "2 個", "2 個", "－"),
            ("誘引通風機 軸受", "22324", "2 個", "2 個", "－"), ("ろ布（抜取部）", "PTFE φ160×6000", "3 本", "3 本", "在庫から使用・補充発注済"),
            ("珪砂 6 号", "", "9 t", "約 9 t", "－"), ("溶接棒", "STBA24 用", "5 kg", "約 2 kg", "残り保管"), ("耐火れんが・キャスタブル", "丸越築炉 持込", "一式", "一式", "－")]
K1_TAGMAP = {"(1)": "T02", "散気": "T02", "減温塔": "T10", "ランス": "T05", "グランド": "T05", "スクレーパ": "T12", "不燃物": "T06", "誘引": "T07", "ろ布": "T09",
             "珪砂": "T02", "溶接棒": "T04", "耐火": "T03"}

def task_text(c, t):
    """日報の工程から報告の文章を書く（日報の連結ではなく、作業・確認した状態・結果）"""
    st = [s[1] for s in t.steps]
    main = [x for x in st if not x.startswith("【追加】")]
    add = [x.replace("【追加】", "") for x in st if x.startswith("【追加】")]
    finds = []
    for x in main:
        if "（" in x and "）" in x:
            finds.append(x[x.index("（") + 1:x.rindex("）")])
    nm = W.BY_TAG[t.tags[0]].name if t.tags else t.name
    import contracts as CC
    if t.tags and (W.BY_TAG[t.tags[0]].cls not in CC.WORKABLE or len(t.tags) > 1):
        seen = []
        for x in main:
            x2 = x
            if x2 not in seen: seen.append(x2)
        s = "、".join(seen[:6]) + "を行った。"
        if add:
            s += "整備中に " + "、".join(add) + " が必要となり、SPC と協議のうえ追加で施工した。"
        return s
    if t.attr == "REPLACE":
        s = f"既設の{nm}を撤去し、新品を据え付けて芯出し・単体試運転を行った。"
    elif t.attr == "INSPECT":
        s = "、".join(dict.fromkeys(x.split("（")[0] for x in main)) + "を行った。"
    else:
        if not finds:
            rr = random.Random(f"{c.cid}{t.tid}")
            finds = [CC.fmt_find(W.BY_TAG[t.tags[0]].cls, rr)]
        tail = rr2 = random.Random(f"{t.tid}{c.fy}").choice(["組立後、芯出しを行った。", "組み立てて芯出しし、単体運転で異常のないことを確認した。", "復旧後の試運転で振動・温度とも問題なかった。", "組立・芯出しのうえ復旧した。"])
        s = f"{nm}を分解点検した。" + "、".join(dict.fromkeys(finds)) + "。" + tail
    if finds and t.attr == "INSPECT":
        s += "結果：" + "、".join(dict.fromkeys(finds)) + "。"
    if add:
        s += "整備中に " + "、".join(add) + " が必要となり、SPC と協議のうえ追加で施工した。"
    return s

# ============================================================== sv1 文章型
def sv1(c, pdf_path, workdir, fname):
    from docx import Document
    from docx.shared import Pt, Mm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    rec = Rec(fname)
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Mm(210), Mm(297)
    sec.left_margin = sec.right_margin = Mm(22); sec.top_margin = sec.bottom_margin = Mm(20)
    st = doc.styles["Normal"]; st.font.name = "IPAGothic"; st.font.size = Pt(10.5)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), "IPAGothic")
    who = "sv1" if c.fy <= 2016 else "sv2"
    rep_d = c.end + dt.timedelta(days=21)
    doc.core_properties.title = f"【架空】ときわ北部クリーンセンター {c.short} 完了報告書（{C.fy_label(c.fy)}）"
    doc.core_properties.author = f"{who}（【架空】ときわ北部エコサービス 保全グループ）"
    doc.core_properties.created = dt.datetime(rep_d.year, rep_d.month, rep_d.day, 10)
    def para(t, size=10.5, bold=False, align=None, space=4):
        p = doc.add_paragraph(); r_ = p.add_run(t); r_.font.size = Pt(size); r_.bold = bold
        if align is not None: p.alignment = align
        p.paragraph_format.space_after = Pt(space); return p
    def fixw(t, mms):
        t.autofit = False
        lay = OxmlElement("w:tblLayout"); lay.set(qn("w:type"), "fixed"); t._tbl.tblPr.append(lay)
        for gc, w in zip(t._tbl.tblGrid.findall(qn("w:gridCol")), mms): gc.set(qn("w:w"), str(int(w * 56.7)))
        for row in t.rows:
            for cc, w in zip(row.cells, mms): cc.width = Mm(w)
    def table(rows, widths_, header=True):
        t = doc.add_table(rows=len(rows), cols=len(rows[0])); t.style = "Table Grid"; fixw(t, widths_)
        for i, row in enumerate(rows):
            for j, v in enumerate(row): t.cell(i, j).text = str(v)
        if header:
            tr = t.rows[0]._tr; trPr = tr.get_or_add_trPr(); el = OxmlElement("w:tblHeader"); el.set(qn("w:val"), "true"); trPr.append(el)
        return t
    h1 = lambda t: para(t, 12, True, space=6)
    # 表紙（題は数えない）
    for _ in range(5): para("")
    para("【架空】", 12, False, WD_ALIGN_PARAGRAPH.CENTER)
    para("ときわ北部クリーンセンター", 18, True, WD_ALIGN_PARAGRAPH.CENTER)
    para(c.short, 22, True, WD_ALIGN_PARAGRAPH.CENTER, 12)
    para("完　了　報　告　書", 22, True, WD_ALIGN_PARAGRAPH.CENTER, 60)
    for _ in range(6): para("")
    para(f"{jdate(rep_d)[:-len(str(rep_d.day)) - 1]}", 12, False, WD_ALIGN_PARAGRAPH.CENTER)
    para("【架空】ときわ北部エコサービス株式会社", 12, False, WD_ALIGN_PARAGRAPH.CENTER)
    para("保全グループ", 12, False, WD_ALIGN_PARAGRAPH.CENTER)
    doc.add_page_break()
    para("【架空】ときわ北部環境組合　施設課　御中", 10.5, False, space=10)
    rec.add("item", {"text": "【架空】ときわ北部環境組合　施設課　御中"}, seq=True)
    h1("1．工事概要")
    sc = "、".join(dict.fromkeys(W.BY_TAG[x].sub for t in c.tasks for x in t.tags))[:60]
    rows = [("工事件名", c.title), ("工　　期", f"{jdate(c.start)}（炉停止）〜{jdate(c.end)}（定格運転確認）" if c.kind == "teishu" else f"{jdate(c.start)}〜{jdate(c.end)}"),
            ("施　　工", f"{W.SUB1}（職長 x1）" + ("／耐火物：" + W.REFR if any(t.company == W.REFR for t in c.tasks) else "")),
            ("監　　督", f"{W.SPC} 保全グループ {who}"), ("対象範囲", (f"{c.furn}号炉 " if c.furn else "") + sc)]
    if c.cid == "T2015-2":
        rows[4] = ("対象範囲", "2号炉 ガス化炉、旋回溶融炉、ボイラ、排ガス処理設備、通風設備、不燃物・スラグ設備")
    table(rows, [28, 138], header=False)
    for a, b in rows: rec.add("item", {"label": a, "value": b}, seq=True)
    para("")
    h1("2．工程")
    if c.cid == "T2015-2":
        pt = "工程は当初計画のとおり 10 月 5 日に炉を停止し、3 日間の冷却ののち 10 月 8 日から開放・点検整備に着手した。ボイラの肉厚測定で過熱器に減肉が見つかり、肉盛補修を追加したため、ボイラ閉止と水圧試験が当初の 10 月 19 日から 10 月 21 日にずれた。復旧作業と並行して行い、10 月 26 日の点火、10 月 30 日の定格運転確認は予定どおりとなった。"
    else:
        adds = [t for t in c.tasks if any(s[1].startswith("【追加】") for s in t.steps)]
        ms = {t[1:]: d for d, t in c.milestones}
        pt = f"{c.start.month} 月 {c.start.day} 日に炉を停止し、冷却ののち開放・点検整備を行った。"
        if adds:
            pt += f"{adds[0].name}で追加補修が生じ 1 日延びたが、"
        if "点火" in ms:
            pt += f"{ms['点火'].month} 月 {ms['点火'].day} 日に点火、{c.end.month} 月 {c.end.day} 日に定格運転を確認した。"
        if c.kind == "dcs":
            pt = "盤の工場製作・試験ののち、1 月 14 日から現地工事に着手した。共通系を先に切り替え、1号炉・2号炉はそれぞれの停止期間に切り替えた。2 月 21 日に引き渡した。"
    para(pt); rec.add("item", {"section": "工程", "text": pt}, seq=True)
    h1("3．実施内容")
    texts = []
    if c.cid == "T2015-2":
        texts = [(tid, a, b) for tid, a, b in K1_TEXTS]
    else:
        n = 0
        for t in c.tasks:
            if not t.tags: continue
            n += 1
            texts.append((t.tid, f"({n}) {t.name}", task_text(c, t)))
    for tid, a, b in texts:
        para(a, 10.5, True, space=2); para(b, 10.5, space=6)
        t = [x for x in c.tasks if x.tid == tid][0]
        rec.add("item", {"section": "実施内容", "heading": a, "text": b, "equipment": [W.BY_TAG[x].name for x in t.tags]}, work=wid(c, t), seq=True)
    if c.kind == "teishu":
        h1("4．点検結果（肉厚測定 抜粋）")
        f_ = c.furn; y = c.fy; smin = C.sh_min(f_, y)
        if c.cid == "T2015-2":
            rows = [("測定部位", "点数", "新製時(mm)", "最小値(mm)", "判定"), ("蒸発管（前壁）", "40", "4.5", "4.1", "良"), ("蒸発管（側壁）", "80", "4.5", "4.0", "良"),
                    ("二次過熱器 下部ベンド", "32", "4.0", "2.6（No.14）", "肉盛補修"), ("二次過熱器 直管部", "32", "4.0", "3.4", "良"), ("エコノマイザ", "48", "4.0", "3.7", "良")]
        else:
            rr = random.Random(y * 3 + f_)
            rows = [("測定部位", "点数", "新製時(mm)", "最小値(mm)", "判定"),
                    ("蒸発管（前壁）", "40", "4.5", f"{4.4 - (y - 2009) * 0.02 - rr.uniform(0, 0.1):.1f}", "良"),
                    ("蒸発管（側壁）", "80", "4.5", f"{4.4 - (y - 2009) * 0.025 - rr.uniform(0, 0.1):.1f}", "良"),
                    ("二次過熱器 下部ベンド", "32", "4.0", f"{smin:.1f}", "要監視" if smin < 3.0 else "良"),
                    ("二次過熱器 直管部", "32", "4.0", f"{min(4.0, smin + 0.6):.1f}", "良"),
                    ("エコノマイザ", "48", "4.0", f"{3.9 - (y - 2009) * 0.015:.1f}", "良")]
        table(rows, [52, 20, 28, 34, 32])
        bt = [x for x in c.tasks if x.tid == "T04"]
        for row in rows[1:]:
            rec.add("item", dict(zip(["part", "points", "new_mm", "min_mm", "judge"], row)), work=wid(c, bt[0]) if bt else None, seq=True)
        para("")
    h1(("5" if c.kind == "teishu" else "4") + "．使用部品・材料")
    rows = [("品名", "仕様", "数量", "使用", "予備品へ")]
    if c.cid == "T2015-2":
        rows += K1_PARTS
    else:
        for tid, nm, spec, q, used, unit, price, ven in c.parts:
            if "試験" in nm or "工事" in nm: continue
            sp = q - used
            rows.append((nm, spec, f"{q:g} {unit}", f"{used:g} {unit}", f"{sp:g} {unit}" if sp else "－"))
        if len(rows) == 1: rows.append(("（支給品なし）", "", "", "", ""))
    table(rows, [46, 36, 20, 20, 44])
    for row in rows[1:]:
        rec.add("item", dict(zip(["name", "spec", "qty", "used", "spare"], row)), seq=True)
    para("")
    h1(("6" if c.kind == "teishu" else "5") + "．今後の課題")
    if c.cid == "T2015-2":
        kad = ["・二次過熱器の減肉は今回 1 本であったが、高温側ベンド部で今後も発生が見込まれる。次回の定期事業者検査で再測定し、材質変更を含めて検討する。",
               "・旋回溶融炉の側壁下段の減りが進んでいる。次回定期整備で残存厚を測定し、部分補修の範囲を決める。",
               "・不燃物抜出装置は 2 年ごとの分解整備を継続する（B は平成28年度）。"]
    else:
        kad = []
        if c.kind == "teishu":
            kad.append(f"・旋回溶融炉の側壁下段 残存厚は最小 {C.melter_thk(c.furn, c.fy)}mm。次回も測定を続ける。")
            kad.append(f"・過熱器の最小肉厚は {C.sh_min(c.furn, c.fy):.1f}mm。高温側ベンド部を重点に測定する。")
        if any(any(s[1].startswith("【追加】") for s in t.steps) for t in c.tasks):
            kad.append("・追加補修した機器は次回定期整備で再点検する。")
        if c.kind == "dcs":
            kad.append("・旧 DCS の予備品は廃棄せず、切替後 1 年間保管する。")
    for k in kad:
        para(k); rec.add("item", {"section": "今後の課題", "text": k}, seq=True)
    para("以上", align=WD_ALIGN_PARAGRAPH.RIGHT)
    fp = sec.footer.paragraphs[0]; fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = fp.add_run()
    for tag, txt in (("begin", None), (None, "PAGE"), ("end", None)):
        if tag: e = OxmlElement("w:fldChar"); e.set(qn("w:fldCharType"), tag); run._r.append(e)
        else: e = OxmlElement("w:instrText"); e.text = txt; run._r.append(e)
    docx_path = os.path.join(workdir, c.cid + ".docx")
    doc.save(docx_path)
    subprocess.run(["soffice", "-env:UserInstallation=file:///tmp/lo_F09", "--headless", "--convert-to", "pdf", "--outdir", workdir, docx_path],
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=300)
    os.replace(os.path.join(workdir, c.cid + ".pdf"), pdf_path)
    rec.save()

# ============================================================== sv2 2018 様式
def _rl():
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    try: pdfmetrics.getFont("IPAG")
    except KeyError:
        pdfmetrics.registerFont(TTFont("IPAG", "/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf"))

K2_DET = [
    ("T02", "1. ガス化炉", "流動砂を抜き出して炉内を清掃し、散気ノズルの摩耗を測定した。交換基準を超えるものはなく、耐火物にも異常はなかった。閉止後に珪砂 6 号 9t を投入した。"),
    ("T03", "2. 旋回溶融炉", "側壁下段の残存厚は最小 118mm で、2023 年度の 2号炉（95mm）より約 2 年分厚い。令和8年度の基幹改良で全面更新されるため、今回は出滓口の損傷れんがを 0.4m² 積み替える最小補修とした。"),
    ("T04", "3. ボイラ", "定期事業者検査（内部検査）を 5/21 に行い、5/23 の水圧試験で漏れのないことを確認した。過熱器 No.9 ベンドで 2.5mm を測定。計算必要厚 2.1mm を上回るため継続使用とし、基幹改良での更新まで毎年測定する。"),
    ("T05", "4. バグフィルタ", "2018 年度交換のろ布 252 本を全数交換した。第4室で旧ろ布の抜取り時にケージの曲がり 3 本を確認し、予備品のケージに交換した（予備品は補充発注済）。交換前の差圧 1.9kPa は立上げ後 0.9kPa となった。"),
    ("T06", "5. 不燃物抜出装置B", "スクリュー外径は φ336 で、羽根先端に硬化肉盛を行い φ349 に戻した。軸受・軸封を交換した。"),
    ("T11", "6. 前処理破砕機 1号", "2号機（予備）に切り替えて受入を続けながら、カッタ刃 40 枚と軸受 4 個を交換した。"),
]
K2_SUMM = [("T02", "1号 ガス化炉本体（1-210-01）", "点検", "開放、散気ノズル摩耗測定、耐火物点検", "異常なし", "R8 定修"),
           ("T03", "1号 旋回溶融炉本体（1-310-01）", "点検・補修", "残存厚測定、出滓口れんが部分積替え 0.4m²", "側壁下段 最小 118mm", "R8 基幹改良で全面更新"),
           ("T04", "1号 ボイラ本体／過熱器／エコノマイザ", "点検", "肉厚測定、定期事業者検査（5/21）、水圧試験", "過熱器 No.9 最小 2.5mm（必要厚 2.1mm）", "R8 基幹改良で過熱器更新"),
           ("T05", "1号 バグフィルタ本体（1-520-01）", "整備", "ろ布 252 本 全数交換、ケージ 3 本交換", "差圧 1.9→0.9kPa", "ろ布 2031 年度目安"),
           ("T06", "1号 不燃物抜出装置B（1-220-02）", "整備", "分解、軸受・軸封交換、羽根先端肉盛", "外径 φ336（肉盛後 φ349）", "2027 年度"),
           ("T07", "1号 定量供給機A（1-110-04）", "整備", "減速機点検、駆動チェーン交換", "良", "2028 年度"),
           ("T08", "1号 アンモニア水注入ポンプA（1-540-04）", "整備", "ダイヤフラム・弁座・弁体交換", "良", "2028 年度"),
           ("T09", "1号 スラグ水砕コンベヤ（1-320-02）", "点検", "チェーン伸び計測", "1.8%（基準 2%）", "R8 停止中に交換"),
           ("T10", "1号 誘引通風機（1-620-01）", "点検", "軸受点検、振動測定", "4.2mm/s（管理値 7.1）", "R8 基幹改良で更新"),
           ("T11", "前処理破砕機 1号（0-102-01）", "整備", "カッタ刃 40 枚・軸受 4 個交換", "良", "2028 年度")]
K2_PARTS = [["ろ布", "PTFE φ160×6000", "264 本", "252 本", "12 本（予備品）", "PO25-0018"],
            ["リテーナ（ケージ）", "SUS304 φ155×6000", "12 本", "3 本", "9 本＋補充 3 本", "PO25-0018 / 0041"],
            ["不燃物抜出装置 軸受", "22220 CAKE4", "2 個", "2 個", "－", "PO25-0021"], ["軸封パッキン", "グランド 一式", "1 式", "1 式", "－", "PO25-0021"],
            ["駆動チェーン", "RS100 142L", "1 本", "1 本", "－", "PO25-0021"], ["ダイヤフラム・弁座キット", "30L/h 用", "2 組", "1 組", "1 組", "PO25-0022"],
            ["破砕機カッタ刃", "t=50", "40 枚", "40 枚", "－", "PO25-0024"], ["破砕機 軸受", "23228", "4 個", "4 個", "－", "PO25-0024"],
            ["硬化肉盛溶接棒", "HF-600 φ4.0", "10 kg", "6 kg", "4 kg", "PO25-0025"], ["珪砂 6 号", "フレコン 1.5t", "6 袋", "6 袋", "－", "PO25-0028"],
            ["耐火れんが・キャスタブル", "丸越築炉 持込", "一式", "一式", "－", "PO25-0027"]]

def cut(s, n):
    if len(s) <= n: return s
    k = s.rfind("、", 0, n)
    return s[:k] + " ほか" if k > 0 else s[:n]

def next_rec(t, c):
    e = W.BY_TAG[t.tags[0]]
    ins, ovh, life = W.POLICY.get(e.cls, W.DEFAULT_POLICY)
    if e.cls in ("gasifier", "melter", "boiler", "bf"): return "次年度 定修"
    if ovh: return f"{c.fy + ovh} 年度"
    return "状態監視"

def sv2(c, pdf_path, fname, po_of=None):
    _rl()
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.lib import colors
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer, PageBreak
    rec = Rec(fname)
    sN = ParagraphStyle("n", fontName="IPAG", fontSize=9, leading=13)
    sS = ParagraphStyle("s", fontName="IPAG", fontSize=8, leading=11)
    sH = ParagraphStyle("h", fontName="IPAG", fontSize=11, leading=15, spaceBefore=6, spaceAfter=4)
    sT = ParagraphStyle("t", fontName="IPAG", fontSize=15, leading=20, alignment=1)
    P = lambda t, s=sS: Paragraph(t, s)
    rep_d = c.end + dt.timedelta(days=14)
    story = [P("【架空】ときわ北部クリーンセンター", sN), P(f"{c.fy}年度　{c.short}　工事報告書", sT), Spacer(1, 4 * mm)]
    kind = "第2期運営委託 年度整備（上限内）" if c.fy >= 2024 else ("第1期 長期包括運営委託（整備込み）" if c.kind == "teishu" else "第1期 長期包括運営委託（更新工事）")
    cont = f"{W.SUB1.replace('株式会社', '')}（職長 x1）" + ("、耐火物 【架空】丸越築炉" if any(t.company == W.REFR for t in c.tasks) else "")
    if c.kind == "dcs": cont = f"{W.MAKER.replace('株式会社', '')}（据付 北嶺プラント工業）"
    head = [("工事件名", c.title, "報告日", f"{rep_d.year}/{rep_d.month}/{rep_d.day}"),
            ("工期", f"{c.start.year}/{c.start.month}/{c.start.day}" + ("（停止）" if c.kind == "teishu" else "") + f" 〜 {c.end.month}/{c.end.day}" + ("（定格確認）" if c.kind == "teishu" else ""), "作成", "sv2（保全グループ）"),
            ("施工", cont, "提出先", "【架空】ときわ北部環境組合 施設課"), ("契約", kind, "確認", "所長")]
    t = Table([[P(a), P(b), P(cc), P(d)] for a, b, cc, d in head], colWidths=[18 * mm, 92 * mm, 16 * mm, 50 * mm])
    t.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.5, colors.black), ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#DDEBF7")),
                           ("BACKGROUND", (2, 0), (2, -1), colors.HexColor("#DDEBF7")), ("VALIGN", (0, 0), (-1, -1), "MIDDLE")]))
    for a, b, cc, d in head:
        rec.add("item", {"label": a, "value": b, "label2": cc, "value2": d}, seq=True)
    story += [t, Spacer(1, 10 * mm), P("【架空】ときわ北部環境組合　施設課　御中", sN), Spacer(1, 4 * mm),
              P("標記工事について、下記のとおり完了しましたので報告します。", sN), Spacer(1, 4 * mm),
              P("記", ParagraphStyle("c", fontName="IPAG", fontSize=10, alignment=1))]
    rec.add("item", {"text": "標記工事について、下記のとおり完了しましたので報告します。"}, seq=True)
    adds = [x for x in c.tasks if any(s[1].startswith("【追加】") for s in x.steps)]
    if c.cid == "T2025-1":
        rng = "1号炉 ガス化炉・旋回溶融炉・ボイラ・バグフィルタ・不燃物・給じん・脱硝・スラグ設備、共通 前処理破砕機1号"
        res = "全項目完了。工程変更 1 件（BF ケージ交換により 1 日延長、全体工程に影響なし）"
    else:
        rng = (f"{c.furn}号炉 " if c.furn else "") + "・".join(dict.fromkeys(W.BY_TAG[x].sub for tk in c.tasks for x in tk.tags))[:50]
        res = "全項目完了。" + (f"追加補修 {len(adds)} 件（{adds[0].name}、全体工程に影響なし）" if adds else "工程変更なし")
    kv = [("1", "工事件名", c.title), ("2", "工　　期", f"{c.start.year}年{c.start.month}月{c.start.day}日 〜 {c.end.year}年{c.end.month}月{c.end.day}日"),
          ("3", "実施範囲", rng), ("4", "結　　果", res), ("5", "添　　付", "(1) 実施結果サマリー　(2) 詳細　(3) 使用部品・材料" + ("　(4) 肉厚測定記録" if c.kind == "teishu" else ""))]
    t = Table([[P(a), P(b), P(cc)] for a, b, cc in kv], colWidths=[8 * mm, 22 * mm, 140 * mm])
    t.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LINEBELOW", (0, 0), (-1, -1), 0.3, colors.grey)]))
    for a, b, cc in kv: rec.add("item", {"no": a, "label": b, "value": cc}, seq=True)
    story += [t, Spacer(1, 16 * mm)]
    st = Table([[P("所長"), P("保全G"), P("作成")], ["", "", ""], [P("（印）"), P("（印）"), P("sv2")]], colWidths=[22 * mm] * 3,
               rowHeights=[6 * mm, 14 * mm, 6 * mm], hAlign="RIGHT")
    st.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.5, colors.black), ("ALIGN", (0, 0), (-1, -1), "CENTER")]))
    story += [st, PageBreak(), P("■ 添付(1) 実施結果サマリー", sH)]
    summ = [["No.", "機器（タグ）", "区分", "実施内容", "結果・判定", "次回"]]
    links = []
    if c.cid == "T2025-1":
        for i, (tid, a, b, cc, d, e) in enumerate(K2_SUMM, 1):
            summ.append([str(i), a, b, cc, d, e]); links.append(tid)
    else:
        i = 0
        for tk in c.tasks:
            if not tk.tags: continue
            i += 1
            e0 = W.BY_TAG[tk.tags[0]]
            lab = (f"{e0.furn}号 " if e0.furn else "共通 ") + e0.name + f"（{e0.tag}）" if len(tk.tags) == 1 else (tk.name if "台" in tk.name else tk.name + f"（{len(tk.tags)} 台）")
            kb = {"INSPECT": "点検", "OVERHAUL": "整備", "REPLACE": "更新", "UPGRADE": "改良"}.get(tk.attr, "")
            st_ = [s[1] for s in tk.steps]
            finds = [x[x.index("（") + 1:x.rindex("）")] for x in st_ if "（" in x and "）" in x]
            content = cut("、".join(dict.fromkeys(x.split("（")[0].replace(e0.name + " ", "") for x in st_ if not x.startswith("【追加】"))), 46)
            result = (cut("、".join(dict.fromkeys(finds)), 34) or "良") + ("／追加補修あり" if tk in adds else "")
            summ.append([str(i), lab, kb, content, result, next_rec(tk, c)]); links.append(tk.tid)
    for row, tid in zip(summ[1:], links):
        tk = [x for x in c.tasks if x.tid == tid][0]
        rec.add("item", dict(zip(["no", "equipment", "class", "content", "result", "next"], row)), work=wid(c, tk), seq=True)
    tbl = Table([[P(x) for x in row] for row in summ], colWidths=[10 * mm, 42 * mm, 13 * mm, 49 * mm, 35 * mm, 27 * mm], repeatRows=1)
    tbl.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.5, colors.black), ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E79")),
                             ("TEXTCOLOR", (0, 0), (-1, 0), colors.white), ("VALIGN", (0, 0), (-1, -1), "MIDDLE")]))
    story += [tbl, Spacer(1, 3 * mm)]
    if c.cid == "T2025-1":
        nt = "工程：BF 第4室でケージ曲がり 3 本が見つかり交換したため、BF ろ布交換の完了が 5/26 → 5/27 となった。全体工程（6/2 点火・6/6 定格）への影響はなし。"
    else:
        nt = "工程：" + (f"{adds[0].name}で追加補修があり 1 日延長。" if adds else "予定どおり。") + "全体工程への影響はなし。"
    story.append(P(nt, sN)); rec.add("item", {"text": nt}, seq=True)
    story += [PageBreak(), P("■ 添付(2) 詳細", sH)]
    dets = K2_DET if c.cid == "T2025-1" else [(tk.tid, f"{n}. {tk.name}", task_text(c, tk)) for n, tk in enumerate([x for x in c.tasks if x.tags], 1)]
    for tid, a, b in dets:
        story += [P(a, sN), P(b, sN), Spacer(1, 2 * mm)]
        tk = [x for x in c.tasks if x.tid == tid][0]
        rec.add("item", {"heading": a, "text": b}, work=wid(c, tk), seq=True)
    story.append(P("■ 添付(3) 使用部品・材料", sH))
    parts = [["品名", "仕様", "手配数", "使用", "予備・在庫", "発注"]]
    if c.cid == "T2025-1":
        parts += K2_PARTS
    else:
        for tid, nm, spec, q, used, unit, price, ven in c.parts:
            if "試験" in nm: continue
            sp = q - used
            parts.append([nm, spec, f"{q:g} {unit}", f"{used:g} {unit}", f"{sp:g} {unit}（予備品）" if sp else "－", (po_of or {}).get((c.cid, tid, nm), "")])
        if len(parts) == 1: parts.append(["（支給品なし）", "", "", "", "", ""])
    for row in parts[1:]: rec.add("item", dict(zip(["name", "spec", "qty", "used", "spare", "order"], row)), seq=True)
    t = Table([[P(x) for x in row] for row in parts], colWidths=[36 * mm, 34 * mm, 16 * mm, 16 * mm, 32 * mm, 30 * mm], repeatRows=1)
    t.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.5, colors.black), ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#DDEBF7"))]))
    story += [t, Spacer(1, 3 * mm)]
    if c.kind == "teishu":
        story.append(P("■ 添付(4) 肉厚測定記録（過熱器 抜粋、単位 mm）", sH))
        y = c.fy; f_ = c.furn
        y1, y2 = y - 8, y - 4
        rows = [["管No.", "部位", "新製時", str(y1), str(y2), str(y), "判定"]]
        rr = random.Random(y * 11 + f_)
        smin = C.sh_min(f_, y)
        for n in range(1, 25):
            v = smin if n == (9 if f_ == 1 else 14) else round(rr.uniform(smin + 0.2, smin + 1.0), 1)
            v2 = round(v + rr.uniform(0.1, 0.3), 1); v1 = round(v2 + rr.uniform(0.2, 0.5), 1)
            rows.append([str(n), "下部ベンド" if n % 2 else "直管部", "4.0", f"{min(v1, 4.0):.1f}", f"{min(v2, 4.0):.1f}", f"{min(v, 4.0):.1f}", "要監視" if v < 2.8 else "良"])
        bt = [x for x in c.tasks if x.tid == "T04"][0]
        for row in rows[1:]:
            rec.add("item", dict(zip(["tube", "part", "new", f"y{y1}", f"y{y2}", f"y{y}", "judge"], row)), work=f"{c.cid}:T04:{W.find(f_, '過熱器').tag}", seq=True)
        t = Table([[P(x) for x in r_] for r_ in rows], colWidths=[14 * mm, 26 * mm, 18 * mm, 18 * mm, 18 * mm, 18 * mm, 20 * mm], repeatRows=1)
        t.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.5, colors.black), ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#DDEBF7"))]))
        story.append(t)
    def onpage(cv, d):
        cv.setFont("IPAG", 7.5)
        cv.drawString(18 * mm, 10 * mm, "【架空】ときわ北部エコサービス株式会社 保全グループ　様式 MR-02（2018改）")
        cv.drawRightString(192 * mm, 10 * mm, f"{d.page}")
    doc = SimpleDocTemplate(pdf_path, pagesize=A4, leftMargin=16 * mm, rightMargin=16 * mm, topMargin=14 * mm, bottomMargin=16 * mm,
                            title=f"【架空】ときわ北部クリーンセンター {c.fy}年度 {c.short} 工事報告書", author="sv2（【架空】ときわ北部エコサービス 保全グループ）")
    doc.build(story, onFirstPage=onpage, onLaterPages=onpage)
    rec.save()

# ============================================================== メーカー（R7）
def maker(c, pdf_path, fname):
    import report_gen as RG
    RG.maker_k3(pdf_path)
    rec = Rec(fname)
    rows = [("発注者", "【架空】ときわ北部環境組合"), ("受注者", "【架空】鷹羽エンジニアリング株式会社　現場代理人 m1"),
            ("契約工期", "令和7年4月18日〜令和10年3月31日（継続費 3 か年）"), ("令和7年度施工", "2号炉　令和7年9月1日（停止）〜令和8年1月16日（引渡し）"),
            ("変更", "第1回変更（R7.12.5）：2号炉施工 12/19→1/16 に延長（過熱器管材の納期遅れ）"), ("提出日", "令和8年2月6日")]
    for a, b in rows: rec.add("item", {"label": a, "value": b}, seq=True)
    data = [("溶融炉", "旋回溶融炉 耐火物全面更新（高クロミア質）", "2-310-01", "一式（約 42m²）", "100%", 4), ("溶融炉", "鉄皮補修（当て板 追加 2.6m²）", "2-310-01", "一式", "100%", 3),
            ("溶融炉", "二次燃焼室 耐火物更新", "2-310-02", "一式", "100%", 5), ("溶融炉", "溶融炉バーナ 更新", "2-310-03", "1 基", "100%", 6),
            ("溶融炉", "炉体冷却水ポンプ A・B 更新", "2-310-06/07", "2 台", "100%", 7), ("ボイラ", "過熱器 一式更新（SUS310S 系）", "2-410-02", "一式", "100%", 9),
            ("ボイラ", "エコノマイザ 管群更新", "2-410-03", "一式", "100%", 11), ("通風", "誘引通風機 更新（高効率型）", "2-620-01", "1 台", "100%", 14),
            ("通風", "押込送風機・燃焼用送風機 更新", "2-610-01/02", "2 台", "100%", 15), ("電気計装", "インバータ盤 更新（誘引・押込）", "2-810-17/18", "2 面", "100%", 16),
            ("電気計装", "炉コントローラ盤 制御改造", "2-850-01/02", "一式", "100%", 17)]
    for kou, nm, tg, q, pr, k in data:
        tags = c.tasks[k][2]
        rec.add("item", {"kou": kou, "content": nm, "equipment": tg, "qty": q, "progress": pr},
                work=[f"{c.cid}:K{k:02d}:{t}" for t in tags], seq=True)
    rec.add("item", {"text": "※ 出来高・金額は別途 出来高報告書による。本書には記載しない。"}, seq=True)
    for ln in ["令和7年9月1日に2号炉を停止し、9月8日から溶融炉の耐火物解体と誘引通風機の撤去に着手した。", "（施工の経過 本文 8 行）",
               "定格負荷（95t/日相当）で 72 時間の連続運転を確認した。", "1号炉の施工（令和8年9月〜）では、鉄皮の減肉範囲を停止前に超音波で再調査する。",
               "過熱器管材は令和8年度分を 2 か月前倒しで手配済み。"]:
        if ln.startswith("（"): continue
        rec.add("item", {"section": "経過・結果・申し送り", "text": ln}, seq=True)
    for n, ph in enumerate(RG.PHOTOS, 1):
        rec.add("item", {"photo": n, "date": ph[0], "kou": ph[1], "place": ph[2], "content": ph[3], "image": f"写真 {n}（図形の枠、実写真なし）"}, seq=True)
    rec.save()
