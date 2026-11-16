# -*- coding: utf-8 -*-
"""発注台帳。H17–H25：市 i1 の Excel（年度ごとに 1 ファイル）。H26–R6：財務会計システムの支出負担行為一覧 CSV（年度ごと）。
金額は市の契約額（委託料・工事請負費・需用費・役務費）。工事全体の原価ではない。"""
import csv
import datetime as dt
import io
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from openpyxl import Workbook  # noqa: E402
from openpyxl.comments import Comment  # noqa: E402
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side  # noqa: E402
from common import FAC_K, FONT, LOCAL, MAKER, OPER, rng, set_props  # noqa: E402
from rec import Recs, path  # noqa: E402
from works import CONTRACTS  # noqa: E402

thin = Side(style="thin")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
F = Font(name=FONT, size=9)
FB = Font(name=FONT, size=9, bold=True)
CE = Alignment(horizontal="center", vertical="center", wrap_text=True)
LE = Alignment(horizontal="left", vertical="center", wrap_text=True)
D = dt.date


def wa(d):
    if d is None:
        return ""
    if d >= D(2019, 5, 1):
        return f"R{d.year - 2018}.{d.month}.{d.day}"
    return f"H{d.year - 1988}.{d.month}.{d.day}"


def year_items(fy, jobs):
    """その年度の契約・購入（正解）。(種別, 起案, 契約, 件名, 方法, 相手, 金額, 期限, 検収, 科目, 備考, 契約ID)"""
    r = rng(f"orders{fy}")
    y = fy
    it = []
    opm = 128_000_000 + (fy - 2005) * 1_700_000 if fy < 2014 else 168_000_000 + (fy - 2014) * 1_200_000
    if fy in (2022, 2023):
        opm = 178_200_000
    it.append(("委託", D(y, 3, 10), D(y, 4, 1), "いずみ川清掃センター運転業務委託", "随意契約", OPER, opm, D(y + 1, 3, 31), "毎月", "委託料", "長期継続", ""))
    for nm, unit, q, sec in (("消石灰購入（単価契約）", 31.5 if fy < 2014 else 44, "予定 380t" if fy < 2014 else "予定 360t", "需用費(薬品)"),
                             ("活性炭購入（単価契約）", 420 if fy < 2014 else 528, "予定 18t", "需用費(薬品)"),
                             ("キレート剤購入（単価契約）", 504 if fy < 2014 else 560, "予定 42t", "需用費(薬品)"),
                             ("灯油購入（単価契約）", 66 + (fy - 2005) * 1.5 if fy < 2014 else 88, "円/L", "需用費(燃料)")):
        it.append(("単価", D(y, 3, 12), D(y, 4, 1), nm, "指名競争" if "灯油" not in nm else "随意契約",
                   "せせらぎ化成販売(株)" if "灯油" not in nm else "いずみ川石油協同組合", round(unit, 1), D(y + 1, 3, 31), "毎月", sec, f"単価 {q}", ""))
    it.append(("委託", D(y, 3, 15), D(y, 4, 1), "電気保安管理業務委託", "随意契約", "いずみ川電気保安協同組合",
               1_386_000 if fy < 2014 else 1_584_000, D(y + 1, 3, 31), D(y + 1, 3, 31), "委託料", "月次・年次点検", ""))
    it.append(("委託", D(y, 4, 5), D(y, 4, 15), "ばい煙・ダイオキシン類測定業務委託", "指名競争", "(株)清流分析センター",
               2_919_000 if fy < 2014 else 3_410_000, D(y + 1, 2, 28), D(y + 1, 3, 4), "委託料", "ばい煙 年2回・DXN 年1回", ""))
    it.append(("委託", D(y, 4, 5), D(y, 4, 15), "作業環境測定業務委託（ダイオキシン類）", "随意契約", "(株)清流分析センター",
               693_000 if fy < 2014 else 770_000, D(y + 1, 3, 31), D(y + 1, 3, 22), "委託料", "6か月以内ごと 年2回", ""))
    it.append(("委託", D(y, 5, 10), D(y, 5, 20), "消防用設備等点検業務委託", "随意契約", "あおば防災設備(株)",
               525_000 if fy < 2014 else 594_000, D(y + 1, 3, 31), D(y + 1, 3, 25), "委託料", "機器点検 2回・総合 1回", ""))
    it.append(("委託", D(y, 5, 24), D(y, 6, 1), "非常用発電機保守点検業務委託", "随意契約", "紫雲電機(株)",
               367_500 if fy < 2014 else 418_000, D(y + 1, 3, 31), D(y + 1, 3, 10), "委託料", "", ""))
    if fy % 2 == 1:
        it.append(("委託", D(y, 6, 5), D(y, 6, 15), "計量機定期検査（代検査）業務委託", "随意契約", "若葉計量器(株)",
                   441_000 if fy < 2014 else 495_000, D(y, 8, 31), D(y, 8, 25), "委託料", "2基", ""))
    if (fy - 2008) % 3 == 0:
        it.append(("委託", D(y, 5, 15), D(y, 6, 1), "精密機能検査業務委託", "指名競争", "(株)清流分析センター",
                   1_785_000 if fy < 2014 else 1_980_000, D(y + 1, 3, 15), D(y + 1, 3, 15), "委託料", "3年に1回", ""))
    it.append(("物品", D(y, 6, 2), D(y, 6, 10), "流動砂（けい砂6号）購入", "見積合せ", "川瀬砂利工業(株)",
               588_000 if fy < 2014 else 418_000, D(y, 8, 31), D(y, 8, 27), "需用費(消耗品)", "定修時支給分" if fy < 2019 else "補充用", ""))
    # 手数料（性能検査）
    for nm, amt, when in (("ボイラー性能検査手数料（1号炉）", 47_000, None), ("ボイラー性能検査手数料（2号炉）", 47_000, None),
                          ("第一種圧力容器性能検査手数料（2基）", 38_000, None)):
        it.append(("手数料", D(y, 7, 1), None, nm, "—", "(社)産業安全検査協会" if fy < 2013 else "(公社)産業安全検査協会",
                   amt, None, None, "役務費(手数料)", "", ""))
    if fy % 2 == 0:
        it.append(("手数料", D(y, 7, 1), None, "ごみクレーン性能検査手数料（2基）", "—",
                   "(社)産業安全検査協会" if fy < 2013 else "(公社)産業安全検査協会", 62_000, None, None, "役務費(手数料)", "", ""))
    # 工事・委託（工事ごと）
    for j in jobs:
        c = j["contract"]
        if j["cid"].startswith("T"):
            base = 92_400_000 + r.randint(0, 12) * 525_000 if fy <= 2016 else (58_800_000 if fy == 2017 else 52_500_000)
            if fy == 2010:
                base = 98_700_000
            it.append(("委託", D(y, 4, 20), D(y, 6, 1), c["name"], "随意契約", MAKER, base, D(y + 1, 3, 18), D(y + 1, 3, 18),
                       "委託料", "特命（メーカー）" + ("。H29から点検範囲を縮小" if fy >= 2017 else ""), j["cid"]))
            added = [t for a in j["areas"] for t in a["tasks"] if t.get("added")]
            if added:
                d0 = min(t["as_"] for t in added)
                amt = 1_260_000 if fy == 2010 else 315_000 * (2 + len(added))
                it.append(("委託", d0 + dt.timedelta(days=10), d0 + dt.timedelta(days=18), c["name"] + "（変更）", "変更契約", MAKER,
                           amt, D(y + 1, 3, 18), D(y + 1, 3, 18), "委託料", "追加：" + "・".join(t["name"].replace("（追加）", "") for t in added)[:60],
                           j["cid"]))
        elif j["cid"] == "P2008" and fy in (2008, 2009):
            if fy == 2008:
                it.append(("工事", D(2008, 5, 12), D(2008, 6, 20), c["name"], "指名競争", MAKER, 486_150_000, D(2010, 3, 19),
                           None, "工事請負費", "継続費 H20 237,300,000／H21 248,850,000（市単独費）", "P2008"))
            else:
                it.append(("工事", None, None, "（継続）部分改良工事 平成21年度分", "—", MAKER, 248_850_000, D(2010, 3, 19),
                           D(2010, 3, 19), "工事請負費", "H20.6.20 契約分の年割", "P2008"))
        elif j["cid"] == "R2014":
            it.append(("工事", D(2014, 7, 14), D(2014, 8, 1), c["name"], "随意契約", MAKER, 18_360_000, D(2014, 11, 28),
                       D(2014, 11, 21), "工事請負費", "定期整備と同じ停止期間で施工", "R2014"))
        elif j["cid"].startswith("L"):
            lim = j["climit"]
            it.append(("工事", j["cdate"] - dt.timedelta(days=20), j["cdate"], c["name"],
                       "指名競争" if j["amount"] >= 3_000_000 else "随意契約（見積合せ）", LOCAL, j["amount"], lim,
                       lim - dt.timedelta(days=6), "工事請負費",
                       {"L20230": "散気ノズル・流動砂は市支給", "L20220": "予備ろ布（H28購入）使用",
                        "L20210": "No.2 減速機を転用"}.get(j["cid"], ""), j["cid"]))
    if fy >= 2019:
        it.append(("委託", D(y, 4, 12), D(y, 4, 25), "ごみクレーン年次自主検査業務委託", "随意契約", "黒潮クレーン(株)", 1_188_000,
                   D(y + 1, 3, 31), D(y, 9, 29) if fy % 2 == 0 else D(y, 11, 30), "委託料",
                   "性能検査受検補助含む" if fy % 2 == 0 else "", f"K{fy}"))
    # 予備品・修繕・その他
    extra = {
        2010: [("修繕", D(2010, 8, 2), D(2010, 8, 10), "プラットホーム出入口扉修繕", "見積合せ", LOCAL, 378_000, D(2010, 9, 10), D(2010, 9, 8), "需用費(修繕料)", "No.2 シャッタ巻上機（0A-005）", ""),
               ("物品", D(2010, 12, 6), D(2010, 12, 15), "ごみクレーンワイヤロープ購入", "見積合せ", "黒潮クレーン(株)", 714_000, D(2011, 1, 31), D(2011, 1, 21), "需用費(消耗品)", "予備 1 組（2号機用 0A-016）", "")],
        2016: [("物品", D(2016, 10, 3), D(2016, 10, 20), "バグフィルタろ布購入（予備品）", "随意契約（特命）", MAKER, 3_110_400, D(2017, 2, 28), D(2017, 2, 24), "需用費(消耗品)", "PTFE 48本 停止まで延命用", ""),
               ("物品", D(2016, 10, 3), D(2016, 10, 20), "中央監視制御盤予備基板購入", "随意契約（特命）", "紫雲電機(株)", 3_596_400, D(2017, 3, 17), D(2017, 3, 15), "需用費(消耗品)", "CPU基板 2・I/O基板 3（生産中止前の最終手配）", ""),
               ("物品", D(2016, 11, 7), D(2016, 11, 21), "誘引通風機インバータユニット購入（予備品）", "随意契約（特命）", "紫雲電機(株)", 2_160_000, D(2017, 3, 17), D(2017, 3, 10), "需用費(消耗品)", "160kW 1台", ""),
               ("物品", D(2016, 11, 7), D(2016, 11, 21), "混練機パドル購入（予備品）", "見積合せ", MAKER, 864_000, D(2017, 2, 28), D(2017, 2, 17), "需用費(消耗品)", "1式", "")],
        2020: [("修繕", D(2020, 6, 8), D(2020, 6, 9), "中央監視制御盤修繕", "随意契約（緊急）", "紫雲電機(株)", 572_000, D(2020, 6, 30), D(2020, 6, 12), "需用費(修繕料)", "CPU基板故障 予備基板へ交換", "S2020a"),
               ("工事", D(2020, 9, 14), D(2020, 10, 1), "計装用空気圧縮機更新工事", "随意契約（見積合せ）", "紫雲電機(株)", 6_380_000, D(2020, 11, 30), D(2020, 11, 26), "工事請負費", "No.1 故障・修理不能", "S2020b")],
        2022: [("物品", D(2022, 11, 7), D(2022, 11, 21), "散気ノズル購入", "随意契約（特命）", MAKER, 4_488_000, D(2023, 3, 31), D(2023, 3, 17), "需用費(消耗品)", "TFB-75用 120本 2号炉補修用予備品 納期 R5.3", "")],
        2023: [("物品", D(2023, 9, 25), D(2023, 10, 4), "流動砂（けい砂6号）購入（補修工事用）", "見積合せ", "川瀬砂利工業(株)", 352_000, D(2023, 11, 24), D(2023, 10, 24), "需用費(消耗品)", "8t 2号炉補修工事用（支給）", "")],
    }
    it += extra.get(fy, [])
    if fy < 2019 and fy not in (2010,) and r.random() < 0.7:
        it.append(("修繕", D(y, 8, 1 + r.randint(0, 20)), D(y, 8, 25), r.choice(["プラットホーム出入口扉修繕", "洗車装置修繕", "ピット放水銃修繕", "エアカーテン修繕"]),
                   "見積合せ", LOCAL, r.randint(18, 60) * 10_500, D(y, 10, 31), D(y, 10, 20), "需用費(修繕料)", "", ""))
    return it


# ---------------------------------------------------------------- Excel（i1）
def excel(fy, items):
    wb = Workbook()
    ws = wb.active
    ws.title = f"H{fy - 1988}"
    ws["A1"] = f"{FAC_K}　発注台帳（平成{fy - 1988}年度）"
    ws["A1"].font = Font(name=FONT, size=13, bold=True)
    ws["A2"] = "環境施設課 i1　（契約・支出負担行為の控え）"
    ws["A2"].font = F
    heads = ["No", "起案日", "契約日", "件名", "契約方法", "相手方", "金額（税込・円）", "履行期限", "検収日", "予算科目", "備考"]
    for j, (h, w) in enumerate(zip(heads, [4, 9, 9, 42, 9, 26, 14, 9, 9, 13, 30]), 1):
        c = ws.cell(row=4, column=j, value=h)
        c.font, c.alignment, c.border = FB, CE, BOX
        c.fill = PatternFill("solid", fgColor="D9D9D9")
        ws.column_dimensions[c.column_letter].width = w
    fn = f"{FAC_K}_発注台帳_平成{fy - 1988}年度_市i1のExcel.xlsx"
    R = Recs(fn)
    r = 5
    prev = None
    blank_done = False
    for k, x in enumerate(items, 1):
        if x[0] == "手数料" and not blank_done:      # 手数料の前に 1 行空ける（i1 の癖）
            r += 1
            blank_done = True
        kihyo, keiyaku = wa(x[1]), wa(x[2])
        if prev and prev[0] == "単価" and x[0] == "単価" and x[1] == prev[1]:
            kihyo = keiyaku = "〃"
        vals = [k, kihyo, keiyaku, x[3], x[4], x[5], x[6], wa(x[7]) if isinstance(x[7], dt.date) else x[7],
                wa(x[8]) if isinstance(x[8], dt.date) else (x[8] or ""), x[9], x[10]]
        for j, v in enumerate(vals, 1):
            c = ws.cell(row=r, column=j, value=v)
            c.font, c.border = F, BOX
            c.alignment = LE if j in (4, 6, 11) else CE
        ws.cell(row=r, column=7).number_format = "#,##0" if x[6] >= 1000 else "#,##0.0"
        if x[0] == "単価":
            ws.cell(row=r, column=7).comment = Comment("単価（円）。年間支出額は支出簿を見ること", "i1")
        R.add(ws.title, r, "item", {h: v for h, v in zip(heads, vals) if v not in (None, "")},
              _wids(x[11]))
        prev = x
        r += 1
    tot = sum(x[6] for x in items if x[6] >= 1000)
    ws.cell(row=r + 1, column=6, value="計（単価契約を除く）").font = FB
    ws.cell(row=r + 1, column=7, value=tot).number_format = "#,##0"
    R.add(ws.title, r + 1, "total", {"金額（税込・円）": tot, "memo": "計（単価契約を除く）"})
    ws.print_title_rows = "4:4"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_area = f"A1:K{r + 1}"
    set_props(wb, f"{FAC_K} 発注台帳 平成{fy - 1988}年度")
    wb.save(path(fn))
    R.save()
    return fn


WBYC = {}


def _wids(cid):
    if not cid:
        return None
    if not WBYC:
        from works import actual_works
        for w in actual_works():
            WBYC.setdefault(w["contract"], []).append(w["wid"])
    return WBYC.get(cid)


# ---------------------------------------------------------------- CSV（財務会計システム）
HEAD = ["年度", "会計", "款", "項", "目", "節", "細節", "所属コード", "所属名", "伝票番号", "起票日", "負担行為日",
        "件名", "契約方法", "債権者番号", "債権者名", "負担行為額", "変更額", "支出済額", "支出日", "履行期限", "摘要"]
SETSU = {"委託料": ("13", "委託料"), "工事請負費": ("15", "工事請負費"), "需用費(薬品)": ("10", "需用費（医薬材料費）"),
         "需用費(燃料)": ("10", "需用費（燃料費）"), "需用費(消耗品)": ("10", "需用費（消耗品費）"),
         "需用費(修繕料)": ("10", "需用費（修繕料）"), "役務費(手数料)": ("11", "役務費（手数料）")}


def vendor_full(n):
    return n.replace("(株)", "株式会社").replace("(公社)", "公益社団法人").replace("(社)", "社団法人")


def vno(n):
    r = rng("v" + n)
    return f"{r.randint(1000, 21999):07d}"


def csv_year(fy, items):
    buf = io.StringIO()
    w = csv.writer(buf, quoting=csv.QUOTE_ALL, lineterminator="\r\n")
    fn = f"{FAC_K}_発注台帳_{'令和' + ('元' if fy == 2019 else str(fy - 2018)) if fy >= 2019 else '平成' + str(fy - 1988)}年度_財務会計システムCSV.csv"
    R = Recs(fn)
    w.writerow(["【架空】いずみ川市 財務会計システム　支出負担行為一覧"])
    w.writerow([f"出力日時：{fy + 1}/04/08 09:{10 + fy % 40:02d}:17", "出力者：u0412", f"検索条件：年度={fy} 所属=040320 目=ごみ処理施設費"])
    w.writerow(HEAD)
    line = 4
    for k, x in enumerate(items, 1):
        setsu, saisetsu = SETSU[x[9]]
        futan = x[2] or x[1]
        unit = x[0] == "単価"
        amt = int(x[6] * {"消石灰": 360_000, "活性炭": 18_000, "キレート": 42_000, "灯油": 60_000}[next(t for t in ("消石灰", "活性炭", "キレート", "灯油") if t in x[3])]) if unit else int(x[6])
        paid = int(amt * (0.97 if unit else 1.0))
        payday = (x[8] + dt.timedelta(days=21)) if isinstance(x[8], dt.date) else D(fy + 1, 4, 25)
        row = [str(fy), "一般会計", "衛生費", "清掃費", "ごみ処理施設費", setsu, saisetsu, "040320", "環境部 清掃センター",
               f"{(fy - 2018) if fy >= 2019 else fy - 1988:02d}-{3400 + k:05d}", futan.strftime("%Y/%m/%d"),
               futan.strftime("%Y/%m/%d"), x[3], x[4].replace("指名競争", "指名競争入札") if x[4] != "—" else "",
               vno(x[5]), vendor_full(x[5]), str(amt), "0", str(paid), payday.strftime("%Y/%m/%d"),
               x[7].strftime("%Y/%m/%d") if isinstance(x[7], dt.date) else "",
               (x[10] + ("　" if x[10] else "") + (f"単価 {x[6]}" if unit else "")).strip()]
        w.writerow(row)
        R.add(None, line, "item", dict(zip(HEAD, row)), _wids(x[11]))
        line += 1
    with open(path(fn), "wb") as f:
        f.write(buf.getvalue().encode("cp932"))
    R.save()
    return fn


def write_all(jobs):
    out = []
    for fy in range(2005, 2025):
        js = [j for j in jobs if (j["fy"] == fy) or (j["cid"] == "P2008" and fy == 2009)]
        items = year_items(fy, js)
        out.append(excel(fy, items) if fy <= 2013 else csv_year(fy, items))
    return out
