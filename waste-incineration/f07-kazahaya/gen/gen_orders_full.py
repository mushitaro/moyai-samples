# -*- coding: utf-8 -*-
"""発注台帳：SPC 会計システム【架空】KAIKEI-21 の発注明細一覧 CSV（Shift_JIS・CRLF）、2015–2026 年度。
2022 年度までは品目コード列なし、2023 年度のシステム改修から予算科目の右に品目コード。
2018・2023 年度は見本の行をそのまま使い、見本に無かった支給品の行を足す。"""
import csv
import io
import math
from datetime import date, timedelta
from common import *
from contracts_full import get
from supplies import all_supplies
from works import WORKS, TASK_WORKS
from world import BY_CODE, h, actual, ALL
import sample_orders as SO

V = dict(SO.V)
V.update({"V0540": "【架空】かざはや溶接工業", "V0550": "【架空】北斗足場"})
LINKS = []
WAREKI = lambda y: "平成{}年度".format(y - 1988) if y <= 2018 else ("令和元年度" if y == 2019 else f"{y}年度")


def tax_rate(d):
    return 0.10 if d >= date(2019, 10, 1) else 0.08


def amount_of(c):
    md = 0
    for t in c["tasks"]:
        n = {"SPC運転": 0}.get(t["crew"], 4)
        md += n * max(1, (t["ae"] - t["as_"]).days)
    base = {"1": 42000000, "2": 38000000, "0": 9000000, "T": 62000000}[c["unit"]]
    v = base + md * 61000
    return int(round(v / 100000) * 100000)


def gen_rows(fy):
    """(発注日, 部門, 予算科目, 品目コード, 取引先, 件名, 仕様, 数量, 単位, 単価, 納期, 検収日, 契約区分, 担当, 備考, links)"""
    rows = []
    who = "sp1" if fy <= 2019 else "sp2"
    sp4 = "sp1" if fy <= 2019 else "sp4"
    y0 = date(fy, 4, 1)
    ye = date(fy + 1, 3, 31)
    nm = WAREKI(fy) if fy <= 2019 else f"{fy}年度"
    k = fy - 2015
    rows += [
        (date(fy, 4, 2 + k % 3), "保全", "委託費", "I-3100", "V0412", f"排ガス等測定業務委託（{nm}）", "ばい煙 炉別年6回・DXN 年1回・作業環境 年2回", 1, "式", 5980000 + 50000 * k, ye, ye - timedelta(days=2), "委託", who, "", []),
        (date(fy, 4, 2 + k % 3), "管理", "委託費", "I-3300", "V0520", f"電気保安管理業務委託（{nm}）", "月次点検・年次点検立会", 1, "式", 3700000 + 30000 * k, ye, ye - timedelta(days=2), "委託", sp4, "", []),
        (date(fy, 4, 2 + k % 3), "管理", "委託費", "I-3400", "V0530", f"消防用設備等点検業務委託（{nm}）", "機器点検2回・総合点検1回", 1, "式", 1100000 + 8000 * k, ye, ye - timedelta(days=9), "委託", sp4, "", []),
        (date(fy, 4, 12 + k % 3), "運転", "薬品費", "Y-4110", "V0601", "消石灰（高反応）", "年間単価契約 ローリー納入", 1480 - 5 * k, "t", 34000 + 600 * k, ye, ye, "単価契約", "sp3" if fy >= 2020 else who, "数量は年間予定", []),
        (date(fy, 4, 12 + k % 3), "運転", "薬品費", "Y-4120", "V0602", "粉末活性炭", "年間単価契約", 120 - k % 3, "t", 280000 + 4000 * k, ye, ye, "単価契約", "sp3" if fy >= 2020 else who, "数量は年間予定", []),
        (date(fy, 4, 12 + k % 3), "運転", "薬品費", "Y-4130", "V0603", "アンモニア水（25%）", "年間単価契約", 205 + k % 4, "t", 46000 + 700 * k, ye, ye, "単価契約", "sp3" if fy >= 2020 else who, "", []),
        (date(fy, 4, 12 + k % 3), "運転", "薬品費", "Y-4140", "V0603", "飛灰処理用キレート剤", "年間単価契約", 160 + k % 5, "t", 242000 + 3000 * k, ye, ye, "単価契約", "sp3" if fy >= 2020 else who, "", []),
    ]
    # 請負（定修）と変更
    for c in get():
        if c["fy"] != fy:
            continue
        S, E = c["period"]
        cname = c["name"] if fy >= 2020 else c["name"].replace(f"{fy}年度", WAREKI(fy))
        ws_ = [w for t in c["tasks"] for w in TASK_WORKS.get((c["key"], t["no"]), [])]
        kenshu = E + timedelta(days=21 + h(c["key"], 9))
        rows.append((max(S - timedelta(days=42 + h(c["key"], 10)), date(fy, 4, 5 + h(c["key"], 6))), "保全", "修繕費", "K-1100", "V0102", cname,
                     f"契約{c['contract_no']} 工期{S.month}/{S.day}～{E.month}/{E.day}", 1, "式", amount_of(c), E, kenshu, "請負", who,
                     "受変電年次点検（風早電設）含む" if c["unit"] == "0" else "", ws_))
        for (what, area) in c.get("changes", []):
            ev = [t for t in c["tasks"] if t.get("event") and t["event"][0] == "追加"]
            d = ev[0]["event"][1] + timedelta(days=8) if ev else S + timedelta(days=15)
            rows.append((d, "保全", "修繕費", "K-1100", "V0102", "（変更）" + cname, f"変更契約1 {what} {area}㎡", 1, "式",
                         int(round(area * 1320000 / 10000) * 10000), E, kenshu, "変更", who, "請負の増額", ws_))
    # 支給品
    for s in all_supplies():
        if not (y0 <= s["order_date"] <= ye):
            continue
        vendor = s["vendor"]
        kenshu = s["due"] + timedelta(days=3)
        if kenshu > ye:
            note = "支給品 翌年度検収"
            kenshu = None
        else:
            note = "支給品"
        acct = "修繕費"
        code = "B-2410" if vendor in ("V0330", "V0520") else "B-2310"
        ws_ = [w["id"] for w in WORKS if w["code"] in s["codes"] and w["work"] == s["work"] and w["contract_key"] == s["contract"]]
        rows.append((s["order_date"], "保全", acct, code, vendor, s["name"], s["spec"], s["qty"], s["unit"], s["price"], s["due"], kenshu,
                     "購入", who, note, ws_))
    # 委託（法定・試験）
    if fy % 2 == 0:
        rows.append((date(fy, 8, 6), "管理", "委託費", "I-3200", "V0610", "計量機 定期検査（計量法）", "搬入1・搬入2・搬出", 1, "式", 486000 + 4000 * k,
                     date(fy, 9, 28), date(fy, 10, 1), "委託", sp4, "", [w["id"] for w in WORKS if w["code"].startswith("0UK-0") and w["start"][:4] == str(fy) and BY_CODE[w["code"]]["base"] == "計量機"]))
    cranes = [w["id"] for w in WORKS if BY_CODE[w["code"]]["base"] in ("ごみクレーン", "灰クレーン", "タービン室天井クレーン") and w["start"][:4] == str(fy)]
    rows.append((date(fy, 8, 6 + k % 5), "保全", "委託費", "I-3200", "V0460",
                 "ごみクレーン・灰クレーン " + ("性能検査・年次自主検査" if fy % 2 else "年次自主検査"), "1号・2号・灰クレーン・タービン室", 1, "式",
                 (2860000 if fy % 2 else 1680000) + 20000 * k, date(fy, 10, 31), date(fy, 11, 6), "委託", who, "", cranes))
    rows.append((date(fy, 9, 8), "保全", "委託費", "I-3200", "V0450", "脱気器 性能検査", "検査手数料・立会", 1, "式", 98000, date(fy, 10, 20),
                 date(fy, 10, 22), "委託", who, "", []))
    for c in get():
        if c["fy"] != fy:
            continue
        for t in c["tasks"]:
            if t["crew"] == "分析" and t["works"]:
                rows.append((t["ps"] - timedelta(days=30), "保全", "委託費", "I-3100", "V0412", f"ろ布・触媒性能試験（{c['unit']}号炉）",
                             "抜取試料の強度・通気度／脱硝率", 1, "式", 860000 + 20000 * k, t["ps"] + timedelta(days=37), t["ps"] + timedelta(days=40),
                             "委託", who, "", TASK_WORKS.get((c["key"], t["no"]), [])))
    if fy in (2017, 2020, 2026):
        rows.append((date(fy, 10, 2), "管理", "委託費", "I-3500", "V0480", f"精密機能検査業務委託（{nm}）", "廃棄物処理法施行規則 3年毎", 1, "式",
                     2600000 + 30000 * k, date(fy + 1, 2, 28), date(fy + 1, 2, 27), "委託", sp4, "", []))
    if fy in (2022,):
        rows.append((date(fy, 7, 9), "保全", "委託費", "I-3600", "V0620", "蒸気タービン定期事業者検査 技術員派遣", "SV 1名×15日 開放点検指導", 1, "式", 3480000,
                     date(fy, 10, 9), date(fy, 10, 14), "委託", who, "", []))
    # 減温塔ノズルチップ（予備品補充）、直営作業の部品
    if fy % 2 == 0:
        rows.append((date(fy, 6, 18), "保全", "修繕費", "B-2310", "V0101", "減温塔 噴霧ノズルチップ（予備品補充）", "二流体 SUS316", 12, "個",
                     38500 + 500 * k, date(fy, 8, 24), date(fy, 8, 27), "購入", who, "定修で使用", []))
    for w in WORKS:
        if w["contract_key"] or w["start"][:4] == "" or not (y0.isoformat() <= w["start"] <= ye.isoformat()):
            continue
        e = BY_CODE[w["code"]]
        if w["t"] == "X" and e["base"] in ("純水装置",):
            rows.append((date.fromisoformat(w["start"]) - timedelta(days=40), "保全", "修繕費", "B-2310", "V0701", "純水装置 RO膜", "8インチ 6本", 6, "本",
                         268000, date.fromisoformat(w["start"]) - timedelta(days=5), date.fromisoformat(w["start"]) - timedelta(days=3), "購入", who, "直営交換", [w["id"]]))
        if w["t"] == "X" and e["base"] in ("ごみクレーン",) and "ワイヤ" in w["work"]:
            rows.append((date.fromisoformat(w["start"]) - timedelta(days=30), "保全", "修繕費", "B-2310", "V0460", f"{e['name']} ワイヤロープ",
                         "φ28 IWRC 6×Fi(29) 交換工事共", 1, "式", 2600000, date.fromisoformat(w["start"]), date.fromisoformat(w["start"]) + timedelta(days=4), "購入", who, "", [w["id"]]))
    # 事後保全（故障修理）
    bdm = [e for e in ALL if not rule_any(e) and e["kind"] in ("機器", "電動機")]
    for i in range(2 + fy % 3):
        e = bdm[h(f"bdm{fy}{i}", len(bdm))]
        d = date(fy, 4, 1) + timedelta(days=30 + h(f"bd{fy}{i}", 300))
        rows.append((d, "保全", "修繕費", "B-2310", "V0701", f"{e['name']} 修理", bdm_text(e), 1, "式",
                     90000 + 10000 * h(f"p{fy}{i}", 30), d + timedelta(days=18), d + timedelta(days=21), "修理", who, "事後保全", []))
    rows.append((date(fy, 11, 13), "保全", "消耗品費", "S-5100", "V0702", "保護具・作業用品（下期）", "防じんマスク用フィルタ・耐熱手袋 ほか", 1, "式",
                 168000 + 2000 * k, date(fy, 11, 24), date(fy, 11, 27), "購入", "sp3" if fy >= 2020 else who, "", []))
    return rows


def bdm_text(e):
    if e["kind"] == "電動機":
        return "軸受交換（異音）"
    if "ポンプ" in e["base"]:
        return "メカシール交換（漏れ）"
    if "扉" in e["base"] or "シャッター" in e["spec"]:
        return "開閉センサ交換（不動作）"
    if "コンベヤ" in e["base"]:
        return "チェーン部分交換"
    return "部品交換（作動不良）"


DT_BDM = ["軸受交換（異音）", "メカシール交換（漏れ）", "センサ交換（不動作）", "モータ巻替え", "パッキン交換"]


def rule_any(e):
    from world import rule
    return bool(rule(e, True))


def fmt_rows(rows, fy, cutoff=None):
    rows = [r for r in rows if cutoff is None or r[0] <= cutoff]
    rows.sort(key=lambda r: r[0])
    seq = {}
    out = []
    for r in rows:
        (d, dept, acct, code, v, nm, spec, q, unit, up, due, kenshu, kind, who, note, links) = r
        ym = d.strftime("%y%m")
        seq[ym] = seq.get(ym, 0) + 2 + (h(nm, 3))
        if cutoff and kenshu and kenshu > cutoff:
            kenshu = None
        no = f"PO{ym}-{seq[ym]:04d}"
        out.append((no, d, dept, acct, code, v, nm, spec, q, unit, up, due, kenshu, kind, who, note, links))
    return out


def sample_rows(fy):
    rows = []
    src = SO.FY2018 if fy == 2018 else SO.FY2023
    for r in src:
        if fy == 2023:
            (no, d, dept, acct, code, vcd, nm, spec, q, unit, up, due, kenshu, kind, who, note) = r
        else:
            (no, d, dept, acct, vcd, nm, spec, q, unit, up, due, kenshu, kind, who, note) = r
            code = ""
        pd = lambda s: date(*map(int, s.split("/"))) if s else None
        rows.append((no, pd(d), dept, acct, code, vcd, nm, spec, q, unit, up, pd(due), pd(kenshu), kind, who, note, []))
    return rows


def link_sample(rows, fy):
    """見本の行にも作業を結ぶ（件名・品名で）。"""
    cs = [c for c in get() if c["fy"] == fy]
    out = []
    for r in rows:
        nm = r[6]
        links = []
        for c in cs:
            base = c["name"].replace(f"{fy}年度", "")
            if base in nm:
                links = [w for t in c["tasks"] for w in TASK_WORKS.get((c["key"], t["no"]), [])]
        for s in all_supplies():
            if s["name"].split("（")[0][:8] in nm and s["order_date"].year in (fy, fy + 1):
                links = [w["id"] for w in WORKS if w["code"] in s["codes"] and w["work"] == s["work"] and w["contract_key"] == s["contract"]]
        out.append(r[:16] + (links,))
    return out


def run():
    files = []
    for fy in range(2015, 2027):
        with_code = fy >= 2023
        cutoff = date(2026, 11, 10) if fy == 2026 else None
        gen = fmt_rows(gen_rows(fy), fy, cutoff)
        if fy in (2018, 2023):
            base = link_sample(sample_rows(fy), fy)
            names = {r[6] for r in base}
            extra = [r for r in gen if r[13] == "購入" and r[15].startswith("支給品") and not any(r[6][:10] in n for n in names)]
            rows = base + extra
            rows.sort(key=lambda r: (r[1], r[0]))
        else:
            rows = gen
        fname = f"{FN}_発注台帳_{fy}年度_会計システムCSV" + ("品目コード付" if with_code else "") + ".csv"
        h_old = ["伝票番号", "発注日", "部門", "予算科目", "取引先コード", "取引先名", "件名・品名", "仕様・摘要", "数量", "単位", "単価",
                 "金額(税抜)", "消費税", "納期", "検収日", "契約区分", "担当", "備考"]
        head = h_old[:4] + (["品目コード"] if with_code else []) + h_old[4:]
        buf = io.StringIO()
        w = csv.writer(buf, lineterminator="\r\n")
        out_d = {2018: "2019/04/08", 2023: "2024/04/09", 2026: "2026/11/10"}.get(fy, f"{fy + 1}/04/{8 + fy % 3:02d}")
        w.writerow([f"#【架空】かざはや環境パートナーズ株式会社 会計システム KAIKEI-21 発注明細一覧 {fy}年度 出力日 {out_d} "
                    f"出力者 {'sp1' if fy <= 2019 else 'sp2'}" + ("（年度途中）" if fy == 2026 else "")])
        w.writerow(head)
        r_no = 3
        for r in rows:
            (no, d, dept, acct, code, vcd, nm, spec, q, unit, up, due, kenshu, kind, who, note, links) = r
            amt = q * up
            tax = math.floor(amt * tax_rate(d))
            fd = lambda x: x.strftime("%Y/%m/%d") if x else ""
            vals = [no, fd(d), dept, acct] + ([code] if with_code else []) + [vcd, V[vcd], nm, spec, q, unit, up, amt, tax, fd(due), fd(kenshu), kind, who, note]
            w.writerow(vals)
            rec(fname, "item", {"no": no, "orderDate": d.isoformat(), "dept": dept, "account": acct, **({"itemCode": code} if with_code else {}),
                                "vendorCode": vcd, "vendor": V[vcd], "name": nm, "spec": spec, "qty": q, "unit": unit, "unitPrice": up,
                                "amount": amt, "tax": tax, "due": due.isoformat() if due else None,
                                "inspected": kenshu.isoformat() if kenshu else None, "kind": kind, "person": who, "note": note}, r=r_no)
            LINKS.append(dict(file=fname, r=r_no, works=links))
            r_no += 1
        with open(outpath(fname), "w", encoding="cp932", newline="") as f:
            f.write(buf.getvalue())
        files.append(fname)
    return files


if __name__ == "__main__":
    for f in run():
        print(f)
