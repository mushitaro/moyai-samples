# -*- coding: utf-8 -*-
"""発注台帳（本作り）
 (1) 財務会計システムの CSV（年度ごと、Shift_JIS・CRLF、負担行為ブロックと支出命令ブロック）
 (2) 部品手配表（年度末の版。R1〜R5 は l1 様式、R6〜 は l2 が列を足した様式。R8 は 11 月時点）"""
import os, csv, io, datetime as dt
from openpyxl import Workbook
from openpyxl.comments import Comment
from common import *
import world as W
from gf_order_base import KAMOKU, CRED, SECT, SECT_N, TEHAI, TEHAI_R6

TODAY = dt.date(2026, 11, 13)

def tax(d):
    return 1.08 if d < dt.date(2019, 10, 1) else 1.10

TEIKI = {2019: 41800000, 2020: 42350000, 2021: 42900000, 2022: 43120000, 2023: 43890000, 2024: 44550000, 2025: 45320000, 2026: 45980000}
GAS = {2019: 3410000, 2020: 3465000, 2021: 3465000, 2022: 3520000, 2023: 3575000, 2024: 3575000, 2025: 3630000, 2026: 3630000}
LIME = {y: 3960000 + 110000 * (y - 2019) for y in range(2019, 2027)}
CARB = {y: 2860000 + 88000 * (y - 2019) for y in range(2019, 2027)}
OIL = {y: 2200000 + 110000 * ((y - 2019) % 4) for y in range(2019, 2027)}

# 補修工事・物品購入・業務委託（契約 id → 行の元）
SPECIAL_ROWS = {
    2019: [("需用費", "消耗品費", "スートブロワランス購入", "随意契約（特命）", "朝凪", "2019/10/01", 902000, "2020/03/13", "図番AN-S15-120 1本", "P"),
           ("需用費", "消耗品費", "誘引通風機軸受購入", "見積合せ", "機工", "2019/09/20", 343200, "2019/10/18", "2個×2", "P"),
           ("委託料", "検査委託", "令和元年度さくら坂焼却場精密機能検査業務委託", "指名競争入札", "ひがた", "2019/07/08", 2860000, "2019/12/20", "", "L")],
    2020: [("工事", "", "令和2年度さくら坂焼却場1号炉バグフィルタろ布取替工事", "指名競争入札", "機工", "2020/05/11", 10780000, "2020/07/22", "工期 R2.5.11〜R2.7.10", "R2020a"),
           ("委託料", "システム委託", "令和2年度さくら坂焼却場中央監視装置監視端末更新業務委託", "指名競争入札", "IT", "2020/09/01", 8250000, "2021/02/10", "監視端末3台・ソフト移植", "R2020b"),
           ("需用費", "消耗品費", "ごみクレーン予備機巻上ワイヤ購入", "随意契約（特命）", "朝凪", "2020/06/01", 2178000, "2020/11/13", "2本", "P")],
    2021: [("工事", "", "令和3年度さくら坂焼却場2号炉バグフィルタろ布取替工事", "指名競争入札", "機工", "2021/09/01", 10890000, "2021/11/26", "工期 R3.9.1〜R3.11.30", "R2021a"),
           ("需用費", "消耗品費", "温水熱交換器管栓購入", "見積合せ", "鉄工", "2021/10/26", 50600, "2021/11/19", "20個（緊急）", "P"),
           ("需用費", "消耗品費", "減温塔噴霧ノズルチップ購入", "随意契約（特命）", "朝凪", "2021/05/20", 330000, "2021/09/30", "16個", "P")],
    2022: [("工事", "", "令和4年度さくら坂焼却場1号ボイラ水管補修工事", "指名競争入札", "機工", "2022/05/09", 11550000, "2022/07/29", "工期 R4.5.9〜R4.7.15", "R2022a"),
           ("工事変更", "", "令和4年度さくら坂焼却場1号ボイラ水管補修工事（変更）", "指名競争入札", "機工", "2022/06/21", 660000, None, "水管 2本追加（計8本）", "R2022a"),
           ("委託料", "検査委託", "令和4年度さくら坂焼却場精密機能検査業務委託", "指名競争入札", "ひがた", "2022/07/05", 2970000, "2022/12/16", "", "L"),
           ("需用費", "消耗品費", "バグフィルタろ布（予備品）購入", "見積合せ", "東雲", "2022/09/12", 1540000, "2022/11/18", "PTFE φ150×6m 20本", "P"),
           ("需用費", "消耗品費", "1号炉燃焼火格子片購入", "随意契約（特命）", "朝凪", "2022/10/14", 13068000, "2023/03/24", "図番AN-G60-203 90枚 納期R5.3.17", "P2022a")],
    2023: [("工事", "", "令和5年度さくら坂焼却場1号炉火格子補修工事", "指名競争入札", "機工", "2023/05/22", 6380000, "2023/07/24", "工期 R5.5.22〜R5.7.31 火格子片は市支給", "R2023a"),
           ("需用費", "消耗品費", "減温塔噴霧ノズルチップ購入", "随意契約（特命）", "朝凪", "2023/06/01", 352000, "2023/10/13", "16個", "P")],
    2024: [("需用費", "消耗品費", "誘引通風機羽根車購入", "随意契約（特命）", "朝凪", "2024/05/20", 9900000, "2025/03/28", "図番AN-F07-IMP2 1個 納期R7.3.14", "P2024a"),
           ("工事", "", "令和6年度さくら坂焼却場飛灰混練機パドル取替工事", "指名競争入札", "機工", "2024/10/01", 4620000, "2024/12/13", "工期 R6.10.1〜R6.12.20", "R2024a"),
           ("需用費", "消耗品費", "誘引通風機軸受購入", "見積合せ", "機工", "2024/11/15", 184800, "2024/12/20", "2個", "P")],
    2025: [("工事", "", "令和7年度さくら坂焼却場2号炉火格子補修工事", "指名競争入札", "機工", "2025/09/01", 4950000, "2025/11/21", "工期 R7.9.1〜R7.11.28 火格子片は市支給", "R2025a"),
           ("委託料", "検査委託", "令和7年度さくら坂焼却場精密機能検査業務委託", "指名競争入札", "ひがた", "2025/07/07", 3080000, "2025/12/19", "", "L"),
           ("需用費", "消耗品費", "減温塔噴霧ノズルチップ購入", "随意契約（特命）", "朝凪", "2025/05/26", 374000, "2025/09/26", "16個", "P")],
    2026: [],
}
CRED2 = dict(CRED)
CRED2["IT"] = ("30655", "【架空】さくら坂情報システム株式会社")
CRED2["鉄工"] = ("20214", "【架空】さくら坂鉄工所")

def rows_for(y):
    R = []
    seq = {"n": 30}
    def no(kind):
        seq["n"] += 7
        return ("R%02d-%04d" % (y - 2018, seq["n"])) if kind != "支出命令" else ("R%02d-S%04d" % (y - 2018, seq["n"] + 300))
    def add(kind, d, km, sai, title, how, cr, amt, memo="", ref=""):
        if dt.date(*map(int, d.split("/"))) > TODAY:
            return
        R.append(dict(fy=y, kind=kind, no=no(kind), d=d, km=km, sai=sai, title=title, how=how, cr=cr, amt=amt, memo=memo, ref=ref))
    N = W.nendo_name(y)
    add("支出負担行為", "%d/04/01" % y, "委託料", "施設維持管理委託", "%sさくら坂焼却場定期整備業務委託" % N, "随意契約", "機工", TEIKI[y], "単年度 部分払2回", "T%d" % y)
    add("支出負担行為", "%d/04/01" % y, "委託料", "施設維持管理委託", "%sさくら坂焼却場電気保安管理業務委託" % N, "随意契約", "みずなら", 1980000, "", "L")
    add("支出負担行為", "%d/04/08" % y, "委託料", "測定委託", "%sさくら坂焼却場排ガス等測定業務委託" % N, "指名競争入札", "かもめ", GAS[y], "ばいじん等 年6回/炉、DXN 年1回/炉", "L")
    add("支出負担行為", "%d/04/08" % y, "委託料", "測定委託", "%sさくら坂焼却場作業環境測定業務委託" % N, "随意契約", "かもめ", 616000, "DXN 年2回", "L")
    add("支出負担行為", "%d/04/15" % y, "需用費", "薬品費", "消石灰単価契約（R%d）" % (y - 2018) if y > 2019 else "消石灰単価契約（R1）", "単価契約", "青葉", LIME[y], "見込額", "")
    add("支出負担行為", "%d/04/15" % y, "需用費", "薬品費", "活性炭・キレート剤単価契約（R%d）" % (y - 2018) if y > 2019 else "活性炭・キレート剤単価契約（R1）", "単価契約", "青葉", CARB[y], "見込額", "")
    add("支出負担行為", "%d/04/15" % y, "需用費", "燃料費", "灯油単価契約（R%d）" % (y - 2018) if y > 2019 else "灯油単価契約（R1）", "単価契約", "若葉", OIL[y], "見込額", "")
    u1 = W.nendo_name(y)
    add("支出負担行為", "%d/05/30" % y, "役務費", "手数料", "ボイラー性能検査手数料（1号）", "", "検査", 69000, "", "L")
    add("支出負担行為", "%d/10/20" % y, "役務費", "手数料", "ボイラー性能検査手数料（2号）", "", "検査", 69000, "", "L")
    add("支出負担行為", "%d/10/20" % y, "役務費", "手数料", "第一種圧力容器性能検査手数料（温水熱交換器2基）", "", "検査", 78000, "", "L")
    if y in (2019, 2021, 2023, 2025):
        add("支出負担行為", "%d/10/31" % y, "役務費", "手数料", "クレーン性能検査手数料（ごみクレーン2基）", "", "検査", 132000, "", "L")
    if y in (2020, 2022, 2024, 2026):
        add("支出負担行為", "%d/08/26" % y, "役務費", "手数料", "計量器定期検査手数料（トラックスケール2基）", "", "計量", 26400, "", "L")
    add("支出負担行為", "%d/11/04" % y, "委託料", "点検委託", "%sさくら坂焼却場消防用設備等点検業務委託" % N, "見積合せ", "防災", 440000, "", "L")
    for km, sai, title, how, cr, d, amt, pay, memo, ref in SPECIAL_ROWS[y]:
        if km == "工事変更":
            add("変更負担行為", d, "工事", sai, title, how, cr, amt, memo, ref)
            continue
        add("支出負担行為", d, km, sai, title, how, cr, amt, memo, ref)
    # 支出命令
    add("支出命令", "%d/10/31" % y, "委託料", "施設維持管理委託", "%sさくら坂焼却場定期整備業務委託（部分払1）" % N, "", "機工", TEIKI[y] // 2, "", "T%d" % y)
    add("支出命令", "%d/03/29" % (y + 1), "委託料", "施設維持管理委託", "%sさくら坂焼却場定期整備業務委託（部分払2）" % N, "", "機工", TEIKI[y] - TEIKI[y] // 2, "", "T%d" % y)
    tot_by = {}
    for km, sai, title, how, cr, d, amt, pay, memo, ref in SPECIAL_ROWS[y]:
        tot_by[title.replace("（変更）", "")] = tot_by.get(title.replace("（変更）", ""), (km, sai, cr, ref, 0, None))
        k0 = tot_by[title.replace("（変更）", "")]
        tot_by[title.replace("（変更）", "")] = (k0[0] if km != "工事変更" else k0[0], k0[1], cr, ref, k0[4] + amt, pay or k0[5])
    for title, (km, sai, cr, ref, amt, pay) in tot_by.items():
        if pay:
            add("支出命令", pay, km if km != "工事変更" else "工事", sai, title, "", cr, amt, "完成払" if km == "工事" else ("完了払" if km == "委託料" else "納入検収後"), ref)
    R.sort(key=lambda x: (x["d"], x["no"]))
    import random as _r
    rr = _r.Random(y)
    cnt = {"A": 20 + rr.randint(0, 15), "S": 300 + rr.randint(0, 80)}
    for x in R:
        k = "S" if x["kind"] == "支出命令" else "A"
        cnt[k] += rr.randint(3, 11)
        x["no"] = ("R%02d-S%04d" if k == "S" else "R%02d-%04d") % (y - 2018, cnt[k])
    # 変更負担行為は元の番号に枝番
    for x in R:
        if x["kind"] == "変更負担行為":
            base = [z for z in R if z["kind"] == "支出負担行為" and z["title"] == x["title"].replace("（変更）", "")][0]
            x["no"] = base["no"] + "-1"
    return R

def fname_csv(y):
    return "【架空】さくら坂焼却場_発注台帳_%s_財務会計CSV.csv" % W.nendo_name(y)

def build_csv(y):
    R = rows_for(y)
    recs = []
    buf = io.StringIO()
    w = csv.writer(buf, lineterminator="\r\n")
    line = [0]
    def wr(row):
        w.writerow(row); line[0] += 1
    a = "%d/04/01" % y; b = min(dt.date(y + 1, 3, 31), dt.date(2026, 11, 12)).strftime("%Y/%m/%d")
    wr(["【架空】さくら坂市 財務会計システム 支出負担行為・支出命令一覧", "所属:%s %s" % (SECT, SECT_N), "年度:%d 起票日:%s-%s" % (y, a, b), "出力日:2026/11/12 09:%02d" % (40 + y - 2019)])
    head = ["年度", "伝票区分", "伝票番号", "起票日", "所属コード", "所属名", "款", "項", "目", "節", "節名称", "細節", "件名", "契約方法", "債権者コード", "債権者名", "金額", "支払予定日", "摘要"]
    blocks = [("＜支出負担行為（変更負担行為を含む）＞", [x for x in R if x["kind"] != "支出命令"]), ("＜支出命令＞", [x for x in R if x["kind"] == "支出命令"])]
    for bi, (label, rs) in enumerate(blocks):
        if bi:
            wr([])
        wr([label, "件数:%d" % len(rs)]); wr(head)
        tot = 0
        for x in rs:
            k = KAMOKU[x["km"]]; c = CRED2[x["cr"]]
            pay = ""
            if x["kind"] == "支出命令":
                dd = dt.date(*map(int, x["d"].split("/"))) + dt.timedelta(15)
                pay = dd.strftime("%Y/%m/%d")
            wr([x["fy"], x["kind"], x["no"], x["d"], SECT, SECT_N, k[0], k[1], k[2], k[3], k[4], x["sai"], x["title"], x["how"], c[0], c[1], x["amt"], pay, x["memo"]])
            recs.append(dict(r=line[0], kind="item", fields={"年度": x["fy"], "伝票区分": x["kind"], "伝票番号": x["no"], "起票日": x["d"], "節": k[4], "細節": x["sai"],
                                                             "件名": x["title"], "契約方法": x["how"], "債権者": c[1], "金額": x["amt"], "支払予定日": pay, "摘要": x["memo"], "contract": x["ref"]}))
            tot += x["amt"]
        wr(["", "小計"] + [""] * 14 + [tot, "", ""])
        recs.append(dict(r=line[0], kind="subtotal", fields={"ブロック": label, "金額": tot}))
    fn = os.path.join(OUT, fname_csv(y))
    with open(fn, "wb") as f:
        f.write(buf.getvalue().encode("cp932"))
    return fn, recs, R

def actual_cost_kyen(y):
    """計画書の決算（千円）：CSV の負担行為（変更含む）から"""
    R = [x for x in rows_for(y) if x["kind"] != "支出命令"]
    c = {"定期整備業務委託": 0, "補修工事（請負）": 0, "物品購入（部品）": 0, "法定検査・測定等委託": 0}
    for x in R:
        if x["ref"].startswith("T"):
            c["定期整備業務委託"] += x["amt"]
        elif x["ref"].startswith("R"):
            c["補修工事（請負）"] += x["amt"]
        elif x["ref"].startswith("P"):
            c["物品購入（部品）"] += x["amt"]
        elif x["ref"] == "L":
            c["法定検査・測定等委託"] += x["amt"]
    return {k: round(v / 1000) for k, v in c.items()}


# ---------------- 部品手配表 ----------------
TEHAI_R7 = [
    ("R7", "減温塔噴霧ノズルチップ", "AN-D14-NZ2", "1・2号 減温塔", "16個", "朝凪炉工", "R7.4.21", "R7.5.12", 340000, "R7.5.26", "R7.9.30", "R7.9.26", "長納期", "廃止までの最終購入", "R07-0093", "16", "6", "R8 まで"),
    ("R7", "火格子片（2号補修）", "AN-G60-203", "2号 燃焼火格子", "30枚", "市倉庫", "—", "—", 0, "—", "—", "—", "中止", "予備 18＋1号取外し品の良品 12 を支給（購入なし）", "", "0", "0", "R7 補修工事で使用"),
]
TEHAI_R8 = [
    ("R8", "誘引通風機 軸受（1号予備）", "球面ころ軸受 22220", "1号 誘引通風機", "2個", "地元商社", "R8.7.1", "R8.7.8", 172000, "", "", "", "通常", "廃止まで 1 年余りのため発注見送り（l2）", "", "2", "0", ""),
]

def fname_tehai(y):
    return "【架空】さくら坂焼却場_発注台帳_部品手配表_%s版_%s.xlsx" % (W.nendo_name(y), "l1様式" if y <= 2023 else "l2列追加様式")

def build_tehai(y):
    new = y >= 2024
    recs = []
    wb = Workbook(); ws = wb.active; ws.title = "手配表"
    cell(ws, "A1", "【架空】さくら坂焼却場　部品手配表（朝凪炉工 ほか）", F(14, True))
    cell(ws, "A2", ("作成 R1.8 施設管理課 l1　／　区分：○生産中止　△長納期（3か月以上）" if not new else
                    "作成 R1.8 施設管理課 l1　／　R6.4 l2：負担行為番号・在庫・廃止までの必要数・使用予定の列を追加、区分を「中止・長納期・通常」に") +
         ("　（%s 時点）" % ("R8.11" if y == 2026 else "%s末" % W.nendo_name(y))), F(9))
    if not new:
        heads = ["No", "品名", "図番・型式", "使用機器", "数量", "手配先", "見積依頼", "見積受領", "見積額\n（税抜・円）", "発注日", "回答納期", "納入日", "区分", "備考"]
        widths = [4, 20, 18, 22, 7, 14, 9, 9, 12, 9, 9, 9, 5, 34]
    else:
        heads = ["No", "品名", "図番・型式", "使用機器", "数量", "手配先", "見積依頼", "見積受領", "見積額\n（税抜・円）", "発注日", "負担行為\n番号", "回答納期", "納入日", "在庫・\n予備", "廃止まで\nの必要数", "区分", "使用予定", "備考"]
        widths = [4, 20, 18, 22, 7, 14, 9, 9, 12, 9, 10, 9, 9, 6, 7, 7, 16, 34]
    for i, h in enumerate(heads):
        col = get_column_letter(1 + i)
        added = new and h.split("\n")[0] in ("負担行為", "在庫・", "廃止まで", "使用予定")
        cell(ws, "%s4" % col, h, F(9, True), CENTER, B_ALL, fill("FFF2CC") if added else fill("D9E1F2"))
        ws.column_dimensions[col].width = widths[i]
    ws.row_dimensions[4].height = 30
    data = [r for r in TEHAI + TEHAI_R6 + TEHAI_R7 + TEHAI_R8 if int(r[0][1:]) + 2018 <= y]
    r = 5; cur = None; no = 0
    for row in data:
        fy = row[0]
        if fy != cur:
            cur = fy
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=len(heads))
            cell(ws, "A%d" % r, "%s年度" % fy.replace("R", "令和").replace("令和1", "令和元"), F(9, True), LEFT, B_ALL, fill("EDEDED"))
            r += 1
        no += 1
        (fy, name, zu, kiki, qty, saki, mi, mj, amt, hd, nk, nn, kb, memo, futan, zaiko, need, yotei) = row
        if new and kb in ("○", "△", ""):
            kb = {"○": "中止", "△": "長納期", "": "通常"}[kb]
        if not new and kb in ("中止", "長納期", "通常"):
            kb = {"中止": "○", "長納期": "△", "通常": ""}[kb]
        vals = [no, name, zu, kiki, qty, saki, mi, mj, amt if amt else "—", hd, futan, nk, nn, zaiko, need, kb, yotei, memo] if new else \
               [no, name, zu, kiki, qty, saki, mi, mj, amt if amt else "—", hd, nk, nn, kb, memo]
        for i, v in enumerate(vals):
            col = get_column_letter(1 + i)
            al = LEFT if heads[i] in ("品名", "図番・型式", "使用機器", "備考", "使用予定", "手配先") else CENTER
            c = cell(ws, "%s%d" % (col, r), v, F(9, color="C00000" if v in ("○", "中止") else "000000"), al, B_HAIR)
            if isinstance(v, int) and i == 8:
                c.number_format = "#,##0"; c.alignment = RIGHT
        recs.append(dict(sheet="手配表", r=r, kind="item", fields={h.replace("\n", ""): v for h, v in zip(heads, vals)}))
        r += 1
    r += 2
    cell(ws, "B%d" % r, "※朝凪炉工 窓口：部品営業 m4（R1.4 から整備撤退、部品のみ受付）。見積は FAX→PDF で保存（共有フォルダ ¥手配¥見積）", F(8))
    if new:
        cell(ws, "B%d" % (r + 1), "※R6 から負担行為番号を書く（財務会計の一覧と突き合わせるため）。R5 以前は分かる分だけ記入。l2", F(8, color="0000CC"))
        ws["B6"].comment = Comment("R1 の見積額は税抜。R2 以降も税抜で統一", "l2")
    page(ws, "landscape", 8 if new else 9, 1, 0, "4:4")
    props(wb, "【架空】さくら坂焼却場 部品手配表 %s版" % W.nendo_name(y), "【架空】さくら坂市 施設管理課 " + ("l2" if y >= 2023 else "l1"))
    fn = os.path.join(OUT, fname_tehai(y))
    wb.save(fn)
    return fn, recs
