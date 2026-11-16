# -*- coding: utf-8 -*-
"""工事報告書（文字のある PDF、reportlab）。原価は載せない。表はページをまたぐと見出しをくり返す。
m1 日誌型（2005–2011）：頭書き → 区域ごとに日付順の作業経過 → 該当日の後に表（肉厚・点検整備結果・交換部品） → 申し送り。
m1 一覧付き（2012–2018）：1 枚目に「整備結果一覧表」（帳票 1 枚）、2 ページ目から日誌型の経過。
g1（2019–）：1 枚目「工事完了報告書」の表、2 ページ目が続紙（作業経過）、その後に記録表・使用材料・所見。"""
import datetime as dt
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from reportlab.lib import colors  # noqa: E402
from reportlab.lib.pagesizes import A4  # noqa: E402
from reportlab.lib.styles import ParagraphStyle  # noqa: E402
from reportlab.lib.units import mm  # noqa: E402
from reportlab.pdfbase import pdfmetrics  # noqa: E402
from reportlab.pdfbase.ttfonts import TTFont  # noqa: E402
from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle  # noqa: E402
from common import FAC_K, LOCAL, MAKER, rng  # noqa: E402
import jobs as J  # noqa: E402
from model2 import thickness  # noqa: E402
from rec import Recs, fname, path, short  # noqa: E402
from works import BYID  # noqa: E402

pdfmetrics.registerFont(TTFont("IPAG", "/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf"))
S = ParagraphStyle("b", fontName="IPAG", fontSize=10, leading=16)
SS = ParagraphStyle("s", fontName="IPAG", fontSize=8.5, leading=12)
H1 = ParagraphStyle("h1", fontName="IPAG", fontSize=13, leading=20, spaceBefore=6, spaceAfter=4)
H2 = ParagraphStyle("h2", fontName="IPAG", fontSize=11, leading=17, spaceBefore=4, spaceAfter=2)
TT = ParagraphStyle("t", fontName="IPAG", fontSize=20, leading=30, alignment=1)
CC = ParagraphStyle("c", fontName="IPAG", fontSize=12, leading=20, alignment=1)
RT = ParagraphStyle("r", fontName="IPAG", fontSize=10, leading=16, alignment=2)
WD = "月火水木金土日"
M1_NAME = {"誘引通風機": "誘引送風機", "散気装置": "散気管", "バグフィルタ": "バグフィルター",
           "定量給じん機": "給塵スクリュー", "廃熱ボイラ": "ボイラ", "流動床炉本体": "焼却炉本体"}
RESULT10 = {   # H22（見本）の点検結果の書き方
    "1B-001": ("炉内点検・フリーボード部耐火物部分補修（追加）", "北面 剥離 1.2m² をはつり、キャスタブル施工", "補修済"),
    "1B-002": ("散気管 全数点検（120本）", "閉塞 2 本を清掃。摩耗は軽微（外径 33.2mm 以上）", "良"),
    "1B-003": ("分解整備", "スクリュー羽根 摩耗 最大 4mm、肉盛溶接・軸受交換", "良"),
    "1B-004": ("分解整備", "軸受交換、二重ダンパ シール交換。組立 1 日遅れ（9/16 完了）", "良"),
    "1C-005": ("チェーン伸び測定", "伸び 1.1%（交換基準 2%）", "良"),
    "1D-004": ("メカニカルシール交換", "シール面 摩耗による漏れ。交換後 漏れなし", "良"),
    "1D-017": ("塔内点検・噴霧ノズル交換（6本）", "内面 付着物軽微", "良"),
    "1E-007": ("ろ布外観点検・抜取り試験（4本）", "引張強さ 保持率 72%（H17 交換、6年目）", "要観察"),
    "1F-005": ("振動測定・軸受点検", "振動 2.8mm/s", "良"),
    "1L-005": ("端子締付け・リレー交換", "タイマーリレー 1 個 不良 交換", "良"),
    "2A-005": ("分解整備", "スクリュー羽根 摩耗 2mm、軸受交換", "良"),
    "2F-005": ("分解整備", "軸受・カップリング交換、羽根車点検（付着物除去）", "良"),
    "2E-013": ("補修", "ロータ摩耗によるシール不良、ロータ肉盛・シール交換", "良"),
    "0K-004": ("補修", "羽根の摩耗 肉盛補修", "良"),
}


class Doc:
    """flowable を積みながら、明細の行（seq）を正解に記録する。"""
    def __init__(self, fn):
        self.E = []
        self.R = Recs(fn)
        self.seq = 0

    def p(self, txt, st=S, item=None, wids=None):
        self.E.append(Paragraph(txt, st))
        if item is not None:
            self.R.add(None, self.seq, "item", item, wids, seq=True)
            self.seq += 1

    def table(self, head, rows, widths, kind_rows=None, wids_rows=None, title=None):
        if title:
            self.E.append(Paragraph(title, SS))
        data = [[Paragraph(str(c), SS) for c in head]] + [[Paragraph(str(c), SS) for c in row] for row in rows]
        t = Table(data, colWidths=widths, repeatRows=1)
        t.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.5, colors.black),
                               ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E7E6E6")),
                               ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                               ("TOPPADDING", (0, 0), (-1, -1), 2), ("BOTTOMPADDING", (0, 0), (-1, -1), 2)]))
        self.E.append(t)
        for k, row in enumerate(rows):
            f = {h: v for h, v in zip(head, row)}
            self.R.add(None, self.seq, (kind_rows or {}).get(k, "item"), f, (wids_rows or {}).get(k), seq=True)
            self.seq += 1


def _wa(d):
    if d >= dt.date(2019, 5, 1):
        n = d.year - 2018
        return f"令和{'元' if n == 1 else n}年{d.month}月{d.day}日"
    return f"平成{d.year - 1988}年{d.month}月{d.day}日"


def m1_name(e):
    nm = e["name"]
    for k, v in M1_NAME.items():
        nm = nm.replace(k, v)
    return nm


def m1_result(w, fy):
    e = BYID[w["eq"]]
    if fy == 2010 and w["eq"] in RESULT10:
        return RESULT10[w["eq"]]
    r = rng("res" + w["wid"])
    if e["base"] == "廃熱ボイラ":
        return ("開放・水洗・水管肉厚測定（84点）", f"最小 {thickness(e['line'], fy):.1f}mm（基準 3.2mm）。性能検査 合格", "良")
    if w["kind"] == "UPGRADE":
        return ("撤去・据付（部分改良）" if e.get("old_tag") else "新設（部分改良）", "据付・試運転 良好", "良")
    if w["kind"] == "REPLACE":
        return ("更新", w["label"], "良")
    if w["kind"] == "OVERHAUL":
        if "補修" in w["label"] or "追加" in w["label"]:
            return ("補修（追加）", r.choice(["パッキン劣化、交換", "軸受異音、軸受交換", "シール漏れ、交換", "摩耗部 肉盛補修"]), "良")
        if e["setsubi"] == "L":
            return ("清掃・端子締付け・絶縁抵抗測定", "絶縁抵抗 良好", "良")
        return ("分解整備", r.choice(["軸受・パッキン交換", "摺動部 摩耗部品交換", "インペラ点検・軸受交換", "シール・軸受交換"]), "良")
    if e["setsubi"] == "L":
        return ("外観・端子締付け確認", r.choice(["異常なし", "異常なし", "盤内清掃", "端子 増締め 2 箇所"]), "良")
    return ("外観・作動点検", r.choice(["異常なし", "異常なし", "異常なし", "軽微な油にじみ（清掃）", "給油実施"]), "良")


def parts_rows(works, fy):
    rows = []
    for w in works:
        e = BYID[w["eq"]]
        if w["kind"] == "OVERHAUL":
            what = "メカニカルシール" if "ポンプ" in e["base"] else ("軸受・パッキン" if e["setsubi"] != "L" else "端子・リレー類")
            rows.append([what, "", "一式", m1_name(e), "追加" if ("補修" in w["label"] or "追加" in w["label"]) else "持込"])
        elif w["kind"] == "REPLACE":
            rows.append(["ろ布（480本）" if "ろ布" in w["label"] else "更新品一式", "", "一式", m1_name(e), "持込"])
    rows.append(["流動砂（けい砂 6号）", "", "補充", "焼却炉本体", "市支給"])
    return rows


def area_tables(D_, job, area):
    """区域の最終日の後に出す表。"""
    fy = job["fy"]
    ws = [w for t in area["tasks"] for w in t["works"]]
    seen, rows, wr = set(), [], {}
    for w in sorted(ws, key=lambda w: w["eq"]):
        if w["eq"] in seen and w["kind"] == "INSPECT":
            continue
        seen.add(w["eq"])
        e = BYID[w["eq"]]
        what, res, judge = m1_result(w, fy)
        wr[len(rows)] = [w["wid"]]
        tag = e.get("old_tag") if (e.get("old_tag") and fy < e["year"]) else e["tag"]   # 更新前は竣工時の機器
        rows.append([tag or e["id"], m1_name(e), what, res, judge])
    D_.table(["機器No.", "機器", "実施内容", "結果", "判定"], rows, [16 * mm, 38 * mm, 40 * mm, 60 * mm, 14 * mm],
             wids_rows=wr, title=f"{area['label']} 点検・整備結果")
    pr = parts_rows(ws, fy)
    D_.table(["品名", "規格", "数量", "使用箇所", "区分"], pr, [38 * mm, 30 * mm, 16 * mm, 56 * mm, 30 * mm],
             title=f"{area['label']} 交換部品・使用材料")


def thick_table(D_, line, fy):
    th = thickness(line, fy)
    prev = thickness(line, fy - 3) if fy - 3 >= 2002 else th + 0.4
    rows = [["第1パス", "24", f"{th + 0.5:.1f}mm", f"{th + 0.8:.1f}mm", "3.2mm"],
            ["第2パス 上部", "24", f"{th + 0.3:.1f}mm", f"{th + 0.6:.1f}mm", "3.2mm"],
            ["第2パス 下部", "24", f"{th:.1f}mm", f"{th + 0.3:.1f}mm", "3.2mm"],
            ["エコノマイザ", "12", f"{th + 0.4:.1f}mm", f"{th + 0.7:.1f}mm", "3.0mm"]]
    D_.table(["測定箇所", "測定点数", "最小値", "平均値", "管理値"], rows, [32 * mm, 20 * mm, 24 * mm, 24 * mm, 24 * mm],
             title=f"{line}号炉 ボイラ水管 肉厚測定結果（超音波）　参考：H{fy - 3 - 1988} 測定 第2パス下部 最小 {prev:.1f}mm"
             + ("（H26 部分取替前）" if line == 1 and fy - 3 < 2014 <= fy else ""))


def diary(D_, job, area):
    days = job["days"]
    s0, s1 = area["span"]
    bo = next((t for t in area["tasks"] if t["name"].startswith("ボイラ開放")), None)
    d = s0
    while d <= s1:
        v = days.get(d)
        items = [it for it in (v["items"] if v else []) if it[0] == area["label"]]
        if items:
            txt = "。".join(f"{it[1]}：{it[2]}" if not it[2].startswith(it[1]) else it[2] for it in items) + "。"
            note = v.get("note") or ""
            if note and (area["label"] in note or v.get("sample")):
                txt += "<br/>※ " + note
            n = sum(it[3] for it in items)
            D_.p(f"<b>{d.month}月{d.day}日（{WD[d.weekday()]}）</b>　{txt}　<font size=8>（{n} 名）</font>",
                 item={"date": d, "text": txt.replace("<br/>", " "), "people": n},
                 wids=[x for it in items for x in it[7]])
        else:
            ops = [t for t in area["tasks"] if t["han"] == "運転" and t["as_"] <= d <= t["ae"]]
            if ops:
                D_.p(f"<b>{d.month}月{d.day}日（{WD[d.weekday()]}）</b>　{ops[0]['name'].replace('（運転業者）', '')}（運転業者）。",
                     item={"date": d, "text": ops[0]["name"]})
        if bo and d == bo["ae"] and area["line"] in (1, 2):
            thick_table(D_, area["line"], job["fy"])
        d += dt.timedelta(days=1)
    area_tables(D_, job, area)


def m1_page(title):
    def f(canvas, doc):
        canvas.saveState()
        canvas.setFont("IPAG", 8)
        if doc.page > 1:
            canvas.drawString(20 * mm, 287 * mm, title)
        canvas.drawCentredString(105 * mm, 10 * mm, f"- {doc.page} -")
        canvas.drawRightString(190 * mm, 10 * mm, MAKER)
        canvas.restoreState()
    return f


def handover(job):
    fy = job["fy"]
    out = []
    if job["cid"] == "P2008":
        return ["1・2号炉とも散気装置・不燃物抜出装置・砂循環設備を更新し、試運転で散気の均一性と砂循環量を確認しました。",
                "新設の不燃物冷却スクリューにより不燃物の排出温度が下がり、振動コンベヤの熱変形は解消される見込みです。",
                "撤去した旧機器の機器番号は、完成図書の番号のまま当社の撤去記録に残しています（改良図書は KR- 番号）。"]
    if job["cid"] == "R2014":
        return ["1号炉ボイラ第2パス下部の水管 6 本を部分取替しました。放射線透過試験・水圧試験とも合格です。",
                "取替えなかった管の最小肉厚は 3.7mm で、当面は毎年の測定で足ります。"]
    t1, t2 = thickness(1, fy), thickness(2, fy)
    for ln, th in ((1, t1), (2, t2)):
        if th < 3.8:
            yrs = max(0.0, (th - 3.2) / (0.13 if ln == 1 and fy <= 2013 else 0.1))
            out.append(f"{ln}号炉ボイラ第2パス下部の水管は最小 {th:.1f}mm です。"
                       + (f"現在の減肉速度では約 {yrs:.0f} 年で管理値に達するため、部分取替の計画をお願いします。" if yrs < 5 else "引き続き毎年測定します。"))
    bfage = {1: fy - max(y for y in (2005, 2011, 2016) if y <= fy), 2: fy - max(y for y in (2000, 2006, 2012, 2017) if y <= fy)}
    for ln, age in bfage.items():
        if age >= 5:
            out.append(f"{ln}号炉バグフィルターのろ布は {age + 1} 年目です。全数交換の時期を計画に入れてください。")
    if fy >= 2010:
        out.append(f"中央監視制御盤は製造後 {fy - 1989} 年を経過し、一部部品の供給が終了しています。")
    added = [t for a in job["areas"] for t in a["tasks"] if t.get("added")]
    for t in added[:2]:
        out.append(f"{t['name'].replace('（追加）', '')}は今回補修しましたが、次回停止時にも状態を確認します。")
    return out or ["特になし。"]


def m1_report(job):
    c = job["contract"]
    fy = job["fy"]
    with_list = fy >= 2012
    fn = fname("工事報告書", short(c["name"]), "m1様式（メーカー・一覧表＋作業経過）" if with_list else "m1様式（メーカー・作業経過日誌型）", "pdf")
    D_ = Doc(fn)
    last = max(a["span"][1] for a in job["areas"])
    sub = _wa(dt.date(fy + 1, 3, 18)) if c["kind"].startswith("定期") else _wa(last + dt.timedelta(days=10))
    title = f"【架空】{c['name']} 業務完了報告書（作業経過）"
    if with_list:
        D_.p("【架空】", SS)
        D_.p("整 備 結 果 一 覧 表", TT)
        D_.p(f"件名　{c['name']}", CC)
        D_.p(f"{MAKER}　現場代理人 m1　　{sub}", RT)
        rows, wr = [], {}
        for a in job["areas"]:
            groups = {}
            for t in a["tasks"]:
                for w in t["works"]:
                    s = BYID[w["eq"]]["setsubi"]
                    g = groups.setdefault(s, dict(n=0, oh=0, rp=0, ad=0, wids=[]))
                    g["n"] += 1
                    g["oh"] += w["kind"] == "OVERHAUL"
                    g["rp"] += w["kind"] in ("REPLACE", "UPGRADE")
                    g["ad"] += bool(t.get("added"))
                    g["wids"].append(w["wid"])
            from world import SETSUBI
            for s, g in groups.items():
                wr[len(rows)] = g["wids"]
                rows.append([a["label"], SETSUBI[s], g["n"], g["oh"] or "", g["rp"] or "", g["ad"] or "",
                             "追加補修あり" if g["ad"] else "良"])
        D_.table(["区域", "設備", "対象台数", "分解整備", "更新", "追加補修", "総合判定"], rows,
                 [20 * mm, 36 * mm, 18 * mm, 18 * mm, 16 * mm, 18 * mm, 40 * mm], wids_rows=wr)
        D_.p("※ 詳細は 2 ページ目以降の作業経過と各表による。", SS)
        D_.E.append(PageBreak())
    else:
        D_.p("【架空】", SS)
    D_.p(sub, RT)
    D_.p("いずみ川市長　様", S)
    D_.p(f"{MAKER}　いずみ川作業所<br/>現場代理人　m1", RT)
    D_.E.append(Spacer(1, 4 * mm))
    D_.p("業 務 完 了 報 告 書", TT)
    D_.p(f"件名　{c['name']}", CC)
    D_.E.append(Spacer(1, 3 * mm))
    order = "、".join(f"{a['label']}（{a['span'][0].month}月）" for a in sorted(job["areas"], key=lambda a: a["span"][0]))
    D_.p(f"標記業務を下記のとおり完了しましたので報告します。整備は {order} の順に行いました。以下、作業の経過を日を追って記し、"
         "点検・整備の結果と交換部品は区域ごとの経過の後に表で示します。", S)
    D_.p("記", CC)
    for k, a in enumerate(sorted(job["areas"], key=lambda a: a["span"][0]), 1):
        D_.p(a["label"] if a["label"] != "共通" else "受入設備・飛灰処理設備・共通設備", H1)
        D_.p(f"（{k}）作業経過", H2)
        diary(D_, job, a)
    D_.p("申し送り", H1)
    for x in handover(job):
        D_.p(x, S, item={"text": x})
    D_.p("機器No. は完成図書（H2.3）・改造図書（DX：H12、KR：H21）の機器番号。記録写真は工事記録写真（CD-R）に収録。", SS)
    D_.p("以上", RT)
    doc = SimpleDocTemplate(path(fn), pagesize=A4, leftMargin=20 * mm, rightMargin=20 * mm, topMargin=20 * mm,
                            bottomMargin=18 * mm, title=f"{FAC_K} {c['name']} 業務完了報告書", author=f"{MAKER}（架空）",
                            subject="【架空】試験用サンプル")
    doc.build(D_.E, onFirstPage=m1_page(title), onLaterPages=m1_page(title))
    D_.R.save()
    return fn


# ---------------------------------------------------------------- g1
def g1_page(canvas, doc):
    canvas.saveState()
    canvas.setFont("IPAG", 8)
    canvas.drawCentredString(105 * mm, 10 * mm, f"{doc.page}")
    if doc.page == 2:
        canvas.drawRightString(190 * mm, 287 * mm, "（続紙）")
    canvas.restoreState()


def nozzle_rows():
    r = rng("nozzle")
    per_col = {"A": 14, "B": 20, "C": 22, "D": 22, "E": 18}
    states = ["閉塞"] * 31 + ["摩耗"] * 52 + ["割れ"] * 13
    r.shuffle(states)
    rows, k, no = [], 0, 0
    for col, nrep in per_col.items():
        idx = set(r.sample(range(24), nrep))
        for i in range(24):
            no += 1
            if i in idx:
                st = states[k]
                k += 1
                od = round(r.uniform(30.1, 31.9), 1) if st == "摩耗" else round(r.uniform(31.5, 33.0), 1)
                rows.append([no, f"{col}-{i + 1:02d}", st, f"{od}", "取替"])
            else:
                rows.append([no, f"{col}-{i + 1:02d}", "良", f"{round(r.uniform(32.3, 33.6), 1)}", "継続使用"])
    return rows


def nozzle_spec(job):
    R = J.J23_RESULT
    job.update(summary=["流動砂抜出・炉内清掃", "風箱開放、散気ノズル全数点検（120本）", f"散気ノズル取替 {R['replaced']} 本（市支給品）",
                        "風箱内清掃・ヘッダ継手パッキン交換 2 箇所（追加、予備日充当）", "風箱閉止・流動砂張込、通風試験"],
               result=(f"点検 120 本のうち 閉塞 {R['ab']['閉塞']}・摩耗 {R['ab']['摩耗']}・割れ {R['ab']['割れ']} の計 {R['replaced']} 本を取替、"
                       f"{R['kept']} 本は継続使用。通風試験で風箱圧力・流動状態とも良好。"),
               supply=(f"散気ノズル 120 本支給（R4 購入品）。使用 {R['replaced']} 本、残 {R['spare_left']} 本は予備品倉庫へ返納。"
                       f"流動砂 新砂 {R['new_sand']}t（再使用砂 {R['reuse_sand']}t と合わせて張込）"),
               tables=[("通風試験結果（10/27）", ["項目", "判定基準", "工事前（R5.9 運転データ）", "工事後", "判定"],
                        [["風箱圧力（流動空気 定格）", "メーカー値 ±10%", "15.9kPa（高め、変動 ±1.2）", "13.4kPa（変動 ±0.3）", "良"],
                         ["炉床 流動状態（目視・砂層温度差）", "砂層温度差 50℃ 以内", "最大 92℃", "冷間のため立上げ後確認", "—"],
                         ["立上げ後 砂層温度差（10/30 運転業者記録）", "50℃ 以内", "—", "最大 28℃", "良"]]),
                       ("散気ノズル点検・取替記録（取替基準：閉塞・割れ・外径 32.0mm 以下、新品 34.0mm）",
                        ["No", "位置（列-番号）", "状態", "外径(mm)", "処置"], nozzle_rows())],
               materials=[["散気ノズル", "TFB-75 用 ねじ込み式", f"{R['replaced']} 本", "市支給", f"支給 120 本、残 {R['spare_left']} 本返納"],
                          ["流動砂（けい砂 6号）", "", f"{R['new_sand']} t", "市支給", f"再使用砂 {R['reuse_sand']}t と合わせ張込"],
                          ["ヘッダ継手パッキン", "耐熱 ジョイントシート", "2 枚", "持込", "追加作業"],
                          ["溶接棒（回り止め溶接）", "φ3.2", "5 kg", "持込", ""], ["ねじ部 焼付防止剤", "", "2 缶", "持込", ""]],
               shoken=["取り替えたノズルは中央部（B～D列）に多く、閉塞はノズル孔への砂・溶融物の固着によるもの。"
                       "継続使用の 24 本も外径 32.3～33.6mm であり、令和6年度末の停止までは使用できると判断する。",
                       "風箱内の砂たまりはヘッダ継手 2 箇所からの砂漏れによる。パッキンを交換した。",
                       "1号炉の散気ノズルは本工事の対象外。予備品 24 本は1号炉の不具合時に使用できる。",
                       "炉内点検の際、減温塔入口ダクトは外観のみ確認した（対象外、異常なし）。"])


def g1_report(job):
    c = job["contract"]
    if job["cid"] == "L20230":
        nozzle_spec(job)
    fn = fname("工事報告書", short(c["name"]), "g1様式（地元業者・表型＋記録表）", "pdf")
    D_ = Doc(fn)
    rows = job["daily"]
    hrs = sum(x[2] * x[3] for x in rows)
    pdays = sum(x[2] for x in rows)
    s0 = min(a["span"][0] for a in job["areas"])
    s1 = max(a["span"][1] for a in job["areas"])
    sub = s1 + dt.timedelta(days=13)
    D_.p("【架空】", SS)
    D_.p("工 事 完 了 報 告 書", TT)
    D_.E.append(Spacer(1, 4 * mm))
    D_.p(f"{_wa(sub)}　　いずみ川市長　様", S)
    D_.p(f"{LOCAL}　職長 g1", RT)
    D_.E.append(Spacer(1, 3 * mm))
    stops = "、".join(f"{a['label']} {a['span'][0].month}/{a['span'][0].day}～{a['span'][1].month}/{a['span'][1].day}" for a in job["areas"])
    form = [["工 事 名", c["name"]],
            ["工事場所", "いずみ川清掃センター " + "・".join(a["label"] for a in job["areas"])],
            ["工　　期", f"{_wa(job['cdate'])} ～ {_wa(job['climit'])}"],
            ["現場作業", f"{stops}　延べ {pdays} 人日（{hrs:g} h）"],
            ["対象機器", job["target"]],
            ["工事概要", "<br/>".join(f"{k}. {x}" for k, x in enumerate(job["summary"], 1))],
            ["結　　果", job["result"]],
            ["支 給 品", job["supply"]],
            ["添付資料", "続紙（作業経過）、記録表、使用材料、写真（別添）"]]
    t = Table([[Paragraph(a, S), Paragraph(b, S)] for a, b in form], colWidths=[28 * mm, 146 * mm])
    t.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.5, colors.black),
                           ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#E7E6E6")), ("VALIGN", (0, 0), (-1, -1), "MIDDLE")]))
    D_.E.append(t)
    D_.R.add(None, D_.seq, "item", {"工事名": c["name"], "工期": form[2][1], "現場作業": form[3][1], "対象機器": job["target"],
                                    "工事概要": job["summary"], "結果": job["result"], "支給品": job["supply"],
                                    "manDays": pdays, "manDaysText": f"{pdays} 人日"},
             [w["wid"] for w in job["works"]], seq=True)
    D_.seq += 1
    D_.E.append(Spacer(1, 6 * mm))
    sign = Table([["監督員確認", "検査員確認", "受付"], ["i2", "", ""]], colWidths=[28 * mm] * 3, rowHeights=[7 * mm, 16 * mm])
    sign.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.5, colors.black), ("FONT", (0, 0), (-1, -1), "IPAG", 9),
                              ("ALIGN", (0, 0), (-1, -1), "CENTER"), ("VALIGN", (0, 0), (-1, -1), "MIDDLE")]))
    sign.hAlign = "RIGHT"
    D_.E += [sign, PageBreak()]
    D_.p("1. 工事概要（続き）　作業経過", H2)
    D_.table(["月日", "作業内容", "人員"], [[f"{x[0].month}/{x[0].day}", x[4], f"{x[2]}"] for x in rows],
             [16 * mm, 142 * mm, 16 * mm], wids_rows={k: x[7] for k, x in enumerate(rows)})
    for k, (ttl, head, trs) in enumerate(job["tables"], 2):
        D_.p(f"{k}. {ttl}", H2)
        w = 174 / len(head)
        D_.table(head, trs, [w * mm] * len(head))
    D_.p(f"{len(job['tables']) + 2}. 使用材料", H2)
    D_.table(["品名", "規格", "数量", "区分", "備考"], job["materials"], [40 * mm, 40 * mm, 20 * mm, 20 * mm, 54 * mm])
    D_.p(f"{len(job['tables']) + 3}. 所見", H2)
    for k, x in enumerate(job["shoken"], 1):
        D_.p(f"({k}) {x}", S, item={"text": x})
    doc = SimpleDocTemplate(path(fn), pagesize=A4, leftMargin=18 * mm, rightMargin=18 * mm, topMargin=18 * mm,
                            bottomMargin=18 * mm, title=f"{FAC_K} {c['name']} 工事完了報告書", author=f"{LOCAL}（架空）",
                            subject="【架空】試験用サンプル")
    doc.build(D_.E, onFirstPage=g1_page, onLaterPages=g1_page)
    D_.R.save()
    return fn


def write(job):
    return g1_report(job) if job["vendor"] == "g1" else m1_report(job)
