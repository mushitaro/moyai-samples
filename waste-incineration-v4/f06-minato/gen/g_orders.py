# -*- coding: utf-8 -*-
"""発注台帳（市の財務会計システムの出力 CSV、年度ごと）
H16〜H22 汎用機の契約データ抽出（ローマ字の項目名、和暦 YYMMDD、金額ゼロ埋め）
H23〜H30 旧財務会計システム（Shift_JIS、全項目引用符、和暦 H30.04.02）
R1〜     新財務会計システム（UTF-8、予算科目が先頭、区分＝契約・変更・年割・前金払）
金額は契約額（税込）。工事の実績原価ではない。
"""
import csv
import datetime as dt
import io
import os
from collections import defaultdict

import ans
import core
from xlhelp import OUT

D = dt.date.fromisoformat
VC = {"汐見重工環境": "0020311", "汐見環境サービス": "0020388", "港南工業": "0030552", "臨海プラント工事": "0030610",
      "北浜フィルター工業": "0041207", "しおかぜ環境分析センター": "0050093", "湾岸電設保安": "0050217", "臨海防災設備": "0050340",
      "みなと計量検定協会": "0050411", "港湾クレーン検査協会": "0050455", "湊環境技術コンサルタント": "0050588"}


def tax_rate(d):
    if d < dt.date(2014, 4, 1):
        return 5
    if d < dt.date(2019, 10, 1):
        return 8
    return 10


def r10k(x):
    return int(round(x / 10000.0)) * 10000


def contracts_of(fy):
    """年度の契約の行（市の発注）：dict(no, name, kind, how, vendor, date, s, e, amount, change, works, extra)"""
    W = core.get()
    rows = []
    seq = defaultdict(int)
    def no(kind):
        seq[kind] += 1
        return seq[kind]
    def amt_of(c, base, k):
        ws = [w for w in W.works if w["contract"] == c["id"]]
        return r10k((base + sum(w["cost"] for w in ws) * 1000 * k) * (1 + tax_rate(dt.date(fy, 4, 1)) / 100)), ws
    c = W.contracts[f"{fy}-TEISEI"]
    a, ws = amt_of(c, 160e6 + (fy - 2004) * 2.5e6, 0.55)
    chg = None
    if fy == 2022:
        chg = (a + 2640000, "2022-12-20", "変更1：2号炉集じん器ケージ交換・再加熱器施栓 追加")
    elif core.rng("chg", fy).random() < 0.25:
        chg = (a + r10k(core.rng("chg2", fy).uniform(0.8e6, 3.5e6)), f"{fy + 1}-01-{core.rng('chg3', fy).randrange(10, 28)}", "変更1：追加補修")
    rows.append(dict(kind="委託", how="随意契約", name=c["name"], vendor="汐見重工環境", date=f"{fy}-04-01", s=f"{fy}-04-01", e=f"{fy + 1}-03-31",
                     amount=a, change=chg, works=ws, cid=c["id"], cat="委託料"))
    if fy >= core.ZENRO_SPLIT_FY:
        c = W.contracts[f"{fy}-ZENRO"]
        a, ws = amt_of(c, 24e6, 0.6)
        rows.append(dict(kind="委託", how="随意契約", name=c["name"], vendor="汐見重工環境", date=f"{fy}-09-28", s=f"{fy}-10-01", e=f"{fy}-12-24",
                         amount=a, change=None, works=ws, cid=c["id"], cat="委託料"))
    if fy in core.TURB_YEARS:
        c = W.contracts[f"{fy}-TURB"]
        a = r10k((96e6 + (fy - 2006) * 0.8e6) * (1 + tax_rate(dt.date(fy, 8, 30)) / 100))
        ws = [w for w in W.works if w["contract"] == c["id"]]
        rows.append(dict(kind="委託", how="随意契約", name=c["name"], vendor="汐見重工環境", date=f"{fy}-08-30", s=f"{fy}-09-01", e=f"{fy + 1}-01-31",
                         amount=a, change=None, works=ws, cid=c["id"], cat="委託料"))
    for c in sorted([c for c in W.contracts.values() if c["fy"] == fy and c["kind"] in ("KOSHIN", "HOSHU")], key=lambda c: c["id"]):
        ws = [w for w in W.works if w["contract"] == c["id"]]
        d0 = min(D(w["start"]) for w in ws) - dt.timedelta(days=40 if c["kind"] == "KOSHIN" else 4)
        if c["kind"] == "HOSHU" and fy == 2022:
            d0 = dt.date(2022, 11, 24)
        k = 1.25 if c["kind"] == "KOSHIN" else 1.6
        a = r10k(sum(w["cost"] for w in ws) * 1000 * k * (1 + tax_rate(d0) / 100))
        if fy == 2022 and c["kind"] == "HOSHU":
            a = 14960000
        rows.append(dict(kind="工事", how="随意契約（緊急）" if c["kind"] == "HOSHU" else "指名競争", name=c["name"], vendor="汐見重工環境",
                         date=d0.isoformat(), s=(d0 + dt.timedelta(days=1)).isoformat(), e=(max(D(w["end"]) for w in ws) + dt.timedelta(days=10)).isoformat(),
                         amount=a, change=None, works=ws, cid=c["id"], cat="工事請負費"))
    lc = W.contracts.get(f"{fy}-LOCAL")
    if lc:
        ws = [w for w in W.works if w["contract"] == lc["id"]]
        byv = defaultdict(list)
        for w in ws:
            byv[w["vendor"]].append(w)
        for v, wl in sorted(byv.items()):
            e0 = core.get().by_tag[wl[0]["tag"]]
            nm = f"{core.wa_fy(fy)} みなと臨海工場 {e0['base_name']}" + ("ほか" if len(wl) > 1 else "") + "補修工事"
            d0 = min(D(w["start"]) for w in wl) - dt.timedelta(days=14)
            a = r10k(max(300000, sum(w["cost"] for w in wl) * 1000 * 1.1) * (1 + tax_rate(d0) / 100))
            rows.append(dict(kind="工事", how="指名競争" if a > 1300000 else "随意契約", name=nm, vendor=v, date=d0.isoformat(), s=(d0 + dt.timedelta(days=1)).isoformat(),
                             e=(max(D(w["end"]) for w in wl) + dt.timedelta(days=7)).isoformat(), amount=a, change=None, works=wl, cid=lc["id"], cat="工事請負費"))
    # ろ布（H30 から市が購入して支給）
    for w in W.works:
        if w["fy"] == fy and fy >= 2018 and w["kind"] == "R" and core.get().by_tag[w["tag"]]["cls"] == "bf":
            u = w["tag"][0]
            d0 = D(w["start"]) - dt.timedelta(days=88)
            tanka = 14200 if fy < 2022 else 15400
            rows.append(dict(kind="物品", how="指名競争", name=f"集じん器用ろ布（{u}号炉）", vendor="北浜フィルター工業", date=d0.isoformat(),
                             s=(d0 + dt.timedelta(days=1)).isoformat(), e=(D(w["start"]) - dt.timedelta(days=10)).isoformat(),
                             amount=int(680 * tanka * (1 + tax_rate(d0) / 100)), change=None, works=[w], cid=None, cat="需用費",
                             item=("集じん器用ろ布", "PTFE φ150×6000 縫製品", 680, "本", tanka)))
    # 基幹改良
    if fy == 2022:
        rows.append(dict(kind="工事", how="一般競争（総合評価）", name="みなと臨海工場 基幹的設備改良工事", vendor="汐見重工環境", date="2022-04-28",
                         s="2022-04-28", e="2026-03-31", amount=8976000000, change=None, works=[w for w in W.works if w["contract"] == "KIKAN"],
                         cid="KIKAN", cat="工事請負費（継続費）", kikan=True))
    if 2022 <= fy <= 2025:
        nen = {2022: 1980000000, 2023: 2310000000, 2024: 2420000000, 2025: 2266000000}[fy]
        rows.append(dict(kind="工事", how="", name=f"みなと臨海工場 基幹的設備改良工事（{fy} 年度 年割）", vendor="汐見重工環境", date="",
                         s=f"{fy}-04-01" if fy > 2022 else "2022-04-28", e=f"{fy + 1}-03-31", amount=nen, change=None,
                         works=[w for w in W.works if w["contract"] == "KIKAN" and w["fy"] == fy], cid="KIKAN", cat="工事請負費（継続費）", nenwari=True))
    # 定例の委託
    def itaku(name, v, base, how="随意契約", m=4, d=1):
        dd = dt.date(fy, m, d)
        rows.append(dict(kind="委託", how=how, name=f"{core.wa_fy(fy)} {name}", vendor=v, date=dd.isoformat(), s=dd.isoformat(), e=f"{fy + 1}-03-31",
                         amount=r10k(base * (1 + tax_rate(dd) / 100)), change=None, works=[], cid=None, cat="委託料"))
    itaku("みなと臨海工場 排ガス等測定分析業務委託", "しおかぜ環境分析センター", 6.2e6 + (fy - 2004) * 0.04e6, "指名競争")
    itaku("みなと臨海工場 作業環境測定業務委託", "しおかぜ環境分析センター", 1.35e6)
    itaku("みなと臨海工場 受変電設備保守点検業務委託", "湾岸電設保安", 5.1e6 + (fy - 2004) * 0.02e6)
    itaku("みなと臨海工場 消防用設備等点検業務委託", "臨海防災設備", 2.0e6, "指名競争", 5, 1)
    if (fy - 2007) % 3 == 0:
        itaku("みなと臨海工場 精密機能検査業務委託", "湊環境技術コンサルタント", 3.6e6, "指名競争", 12, 5)
    if fy % 2 == 0:
        itaku("ごみ計量機定期検査", "みなと計量検定協会", 0.2e6, "随意契約", 6, 10)
        itaku("クレーン性能検査（ごみクレーン1〜3号・保守用）", "港湾クレーン検査協会", 0.28e6, "随意契約", 7, 1)
    else:
        itaku("クレーン性能検査（灰クレーン）", "港湾クレーン検査協会", 0.07e6, "随意契約", 7, 1)
    # 運転管理（長期継続：H16〜H17、H18〜H22、H23〜H27、H28〜R2、R3〜R7、R8〜）
    periods = [(2004, 2005), (2006, 2010), (2011, 2015), (2016, 2020), (2021, 2025), (2026, 2030)]
    p = next(x for x in periods if x[0] <= fy <= x[1])
    rows.append(dict(kind="委託", how="長期継続", name=f"{core.wa_fy(fy)} みなと臨海工場 運転管理業務委託（年度分）", vendor="汐見環境サービス",
                     date=f"{p[0] - 1}-03-26", s=f"{fy}-04-01", e=f"{fy + 1}-03-31",
                     amount=r10k((520e6 + (fy - 2004) * 6e6) * (1 + tax_rate(dt.date(fy, 4, 1)) / 100)), change=None, works=[], cid=None, cat="委託料"))
    return rows


def paid(row, fy):
    """支払日・支払額（2026 年度は出力日時点で未払いのものあり）"""
    e = D(row["e"]) if row["e"] else None
    if not e:
        return "", ""
    pay = e + dt.timedelta(days=25)
    if pay > dt.date(2026, 11, 10):
        return "", ""
    amt = row["change"][0] if row["change"] else row["amount"]
    if row.get("kikan"):
        return "", ""
    if row.get("nenwari"):
        return pay.isoformat(), int(row["amount"] * 0.2) if row["s"] == "2022-04-28" else row["amount"]
    return pay.isoformat(), amt


def wa6(d):
    if not d:
        return "000000"
    d = D(d) if isinstance(d, str) else d
    y = d.year - 1988 if d < dt.date(2019, 5, 1) else d.year - 2018
    return f"{y:02d}{d.month:02d}{d.day:02d}"


def wadot(d):
    if not d:
        return ""
    d = D(d) if isinstance(d, str) else d
    if d < dt.date(2019, 5, 1):
        return f"H{d.year - 1988}.{d.month:02d}.{d.day:02d}"
    return f"R{d.year - 2018}.{d.month:02d}.{d.day:02d}"


def build(fy):
    rows = contracts_of(fy)
    if fy <= 2010:
        return build_hanyo(fy, rows)
    if fy <= 2018:
        return build_old(fy, rows)
    return build_new(fy, rows)


def link(fn, r, row):
    ans.rec(fn, "", r, "item", row["_fields"], works=[w["id"] for w in row["works"]] or None, check=row["_check"])


def build_hanyo(fy, rows):
    fn = f"【架空】みなと臨海工場_発注台帳_{core.wa_fy(fy)}_汎用機契約データ.csv"
    ans.start(fn, False, "発注台帳", fy, "汎用機 契約データ抽出（H16〜H22）", "市 財務会計（汎用機）")
    buf = io.StringIO()
    w = csv.writer(buf, lineterminator="\r\n")
    w.writerow([f"*【架空】ミナトシ ザイムカイケイ（汎用機）　契約データ抽出　{core.wa_fy(fy)}　課4120"])
    w.writerow(["NENDO", "KACD", "KEIYAKUNO", "KENMEI", "KBN", "HOHO", "GYOSHACD", "GYOSHAMEI", "KEIYAKUBI", "KOKIFROM", "KOKITO",
                "KINGAKU", "ZEIGAKU", "HENKOGO", "SHIHARAIBI"])
    kb = {"委託": "2", "工事": "1", "物品": "3"}
    r = 3
    for i, row in enumerate(rows, start=1):
        amt = row["amount"]
        tax = round(amt * tax_rate(D(row["date"])) / (100 + tax_rate(D(row["date"]))))
        pd_, pa = paid(row, fy)
        vals = [f"{fy - 1988:02d}", "4120", f"{fy - 1988:02d}{i:04d}", row["name"].replace("平成", "H").replace("年度 ", "年度"), kb[row["kind"]],
                {"随意契約": "Z", "指名競争": "S", "一般競争（総合評価）": "I", "長期継続": "C", "随意契約（緊急）": "Z"}.get(row["how"], "Z"),
                VC[row["vendor"]][2:], row["vendor"], wa6(row["date"]), wa6(row["s"]), wa6(row["e"]), f"{amt:012d}", f"{tax:010d}",
                f"{row['change'][0]:012d}" if row["change"] else "000000000000", wa6(pd_)]
        w.writerow(vals)
        row["_fields"] = dict(年度=fy, 契約番号=vals[2], 件名=vals[3], 区分=row["kind"], 業者=row["vendor"], 契約日=row["date"],
                              期間自=row["s"], 期間至=row["e"], 金額=amt, 消費税=tax, 変更後=row["change"][0] if row["change"] else None, 支払日=pd_ or None)
        row["_check"] = [vals[2]]
        link(fn, r, row)
        r += 1
    with open(os.path.join(OUT, fn), "wb") as f:
        f.write(buf.getvalue().encode("cp932"))
    return fn


OLD_HDR = ["年度", "課コード", "課名", "契約番号", "契約件名", "契約区分", "契約方法", "業者コード", "業者名", "契約日", "工期自", "工期至",
           "契約金額", "うち消費税", "変更回数", "変更後金額", "支払済額", "支払日", "検査日", "備考"]


def build_old(fy, rows):
    fn = f"【架空】みなと臨海工場_発注台帳_{core.wa_fy(fy)}_旧財務会計システム出力.csv"
    ans.start(fn, False, "発注台帳", fy, "旧財務会計システム（H23〜H30）", "市 財務会計（旧）")
    buf = io.StringIO()
    w = csv.writer(buf, lineterminator="\r\n", quoting=csv.QUOTE_ALL)
    w.writerow([f"【架空】みなと市財務会計システム　契約台帳（課別）　{core.wa_fy(fy)}　出力日：{wadot(dt.date(fy + 1, 4, 26))}"] + [""] * (len(OLD_HDR) - 1))
    w.writerow(OLD_HDR)
    r = 3
    pre = {"委託": "施", "工事": "施", "物品": "物"}
    for i, row in enumerate(rows, start=1):
        final = row["change"][0] if row["change"] else row["amount"]
        tax = round(final * tax_rate(D(row["date"])) / (100 + tax_rate(D(row["date"]))))
        pd_, pa = paid(row, fy)
        no = f"H{fy - 1988}-{pre[row['kind']]}-{i:03d}"
        bk = ""
        if row.get("item"):
            bk = f"{row['item'][2]}本 支給品"
        if row["how"] == "長期継続":
            bk = "長期継続契約の年度分"
        vals = [str(fy - 1988), "4120", "環境局施設課", no, row["name"].replace(" ", ""), row["kind"], row["how"], VC[row["vendor"]][2:],
                row["vendor"] + "（株）" if row["vendor"] not in ("みなと計量検定協会", "港湾クレーン検査協会") else row["vendor"],
                wadot(row["date"]), wadot(row["s"]), wadot(row["e"]), str(row["amount"]), str(tax), "1" if row["change"] else "0",
                str(row["change"][0]) if row["change"] else "", str(pa) if pa != "" else "", wadot(pd_) if pd_ else "",
                wadot(D(row["e"]) - dt.timedelta(days=5)) if pd_ else "", bk + ((" " + row["change"][2]) if row["change"] else "")]
        w.writerow(vals)
        row["_fields"] = dict(年度=fy, 契約番号=no, 件名=vals[4], 区分=row["kind"], 方法=row["how"], 業者=vals[8], 契約日=row["date"],
                              工期自=row["s"], 工期至=row["e"], 金額=row["amount"], 消費税=tax, 変更後=row["change"][0] if row["change"] else None,
                              支払済額=pa if pa != "" else None, 支払日=pd_ or None, 備考=vals[19])
        row["_check"] = [no]
        link(fn, r, row)
        r += 1
    with open(os.path.join(OUT, fn), "wb") as f:
        f.write(buf.getvalue().encode("cp932"))
    return fn


NEW_HDR = ["予算科目コード", "予算科目名", "会計年度", "区分", "契約番号", "変更番号", "事業名", "契約件名", "契約種別", "契約方法",
           "相手方コード", "相手方名称", "契約締結日", "履行期間開始", "履行期間終了", "金額（税込）", "うち消費税", "完了検査日",
           "支払日", "支払額", "品名", "規格", "数量", "単位", "単価", "担当"]
KAMOKU = {"委託料": ("04-01-02-13", "委託料（清掃工場）"), "工事請負費": ("04-01-02-15", "工事請負費（清掃工場）"),
          "需用費": ("04-01-02-11", "需用費（修繕・消耗品）"), "工事請負費（継続費）": ("04-01-02-15", "工事請負費（基幹改良・継続費）")}


def build_new(fy, rows):
    fn = f"【架空】みなと臨海工場_発注台帳_{core.wa_fy(fy)}_新財務会計システム出力.csv"
    ans.start(fn, False, "発注台帳", fy, "新財務会計システム（R1〜）", "市 財務会計（新）")
    out = []
    i = 0
    for row in rows:
        km = KAMOKU[row["cat"]]
        jig = "清掃工場基幹的設備改良事業（継続費 R4〜R7）" if "継続費" in row["cat"] else "みなと臨海工場維持管理事業"
        if row.get("nenwari"):
            no = "R4-4120-0004"
            kb = "年割"
        elif row.get("kikan"):
            no = "R4-4120-0004"
            kb = "契約"
        else:
            i += 1
            no = f"R{fy - 2018}-4120-{(100 + i) if row['kind'] == '物品' else i:04d}"
            kb = "契約"
        d0 = D(row["date"]) if row["date"] else D(row["s"])
        tax = round(row["amount"] * tax_rate(d0) / (100 + tax_rate(d0)))
        pd_, pa = paid(row, fy)
        it = row.get("item") or ("", "", "", "", "")
        tanto = core.city(fy) + ("・技術 g2" if "継続費" in row["cat"] else "")
        vals = [km[0], km[1], str(fy), kb, no, "", jig, row["name"], row["kind"], row["how"], VC[row["vendor"]], row["vendor"] + "株式会社"
                if row["vendor"] not in ("みなと計量検定協会", "港湾クレーン検査協会") else row["vendor"],
                row["date"].replace("-", "/"), row["s"].replace("-", "/"), row["e"].replace("-", "/"), row["amount"], tax,
                (D(row["e"]) - dt.timedelta(days=5)).isoformat().replace("-", "/") if pd_ else "", pd_.replace("-", "/") if pd_ else "", pa,
                it[0], it[1], it[2], it[3], it[4], tanto]
        out.append((vals, row, dict(会計年度=fy, 区分=kb, 契約番号=no, 件名=row["name"], 種別=row["kind"], 相手方=vals[11], 契約締結日=row["date"] or None,
                                    開始=row["s"], 終了=row["e"], 金額=row["amount"], 消費税=tax, 支払日=pd_ or None, 支払額=pa if pa != "" else None,
                                    品名=it[0] or None, 数量=it[2] or None, 単価=it[4] or None)))
        if row["change"]:
            v2 = list(vals)
            v2[3], v2[5], v2[7], v2[12], v2[15] = "変更", "1", row["name"] + f"（{row['change'][2]}）", row["change"][1].replace("-", "/"), row["change"][0]
            v2[16] = round(row["change"][0] * tax_rate(d0) / (100 + tax_rate(d0)))
            v2[17] = v2[18] = v2[19] = ""
            out.append((v2, dict(row, works=[]), dict(会計年度=fy, 区分="変更", 契約番号=no, 変更番号=1, 件名=v2[7], 金額=row["change"][0])))
        if row.get("kikan") and fy == 2022:
            v3 = list(vals)
            v3[3], v3[7], v3[12], v3[13], v3[14], v3[15], v3[16] = "前金払", row["name"] + "（前金払）", "", "", "", 792000000, ""
            v3[17], v3[18], v3[19] = "", "2022/05/31", 792000000
            out.append((v3, dict(row, works=[]), dict(会計年度=fy, 区分="前金払", 契約番号=no, 件名=v3[7], 支払日="2022-05-31", 支払額=792000000)))
    buf = io.StringIO()
    w = csv.writer(buf, lineterminator="\r\n")
    w.writerow(["帳票ID=KK0312", "帳票名=【架空】契約一覧（みなと市財務会計システム）", "所属=4120 環境局施設課", f"会計年度={fy}",
                f"出力日時={'2026-11-10 09:05' if fy == 2026 else f'{fy + 1}-05-10 14:22'}"])
    w.writerow(NEW_HDR)
    r = 3
    for vals, row, fields in out:
        w.writerow(vals)
        ans.rec(fn, "", r, "item", fields, works=[x["id"] for x in row["works"]] or None, check=[vals[4], str(vals[15]) if vals[15] != "" else None])
        r += 1
    w.writerow([f"件数={len(out)}"])
    with open(os.path.join(OUT, fn), "wb") as f:
        f.write(buf.getvalue().encode("utf-8"))
    return fn


def main():
    return [build(fy) for fy in core.FY_RANGE]


if __name__ == "__main__":
    print(main()[-1])
