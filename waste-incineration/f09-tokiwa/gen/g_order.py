# -*- coding: utf-8 -*-
"""発注台帳（年度ごとの CSV）：SPC 旧購買（2009–2015）／新購買（2016–2026）／組合 財務会計（R7・R8、基幹改良分）"""
import csv, io, datetime as dt, random, zlib
from collections import OrderedDict
import world as W, contracts as C
import order_gen as OG
from ledger_gen import hankaku
from rec import Rec

VEN = {"【架空】鷹羽エンジニアリング": "0201", "【架空】北嶺プラント工業": "0310", "【架空】丸越築炉": "0322", "【架空】峰岡送風機": "0405",
       "【架空】八潮ポンプ": "0215", "【架空】黒松破砕機": "0402", "【架空】大鳥フィルター": "0601", "【架空】環境分析センター": "0560",
       "【架空】三國クレーン": "0450", "【架空】東凪タービン": "0470", "【架空】青柳計測": "0480", "【架空】杉原計量器": "0490",
       "【架空】白浜石灰工業": "0112", "【架空】炭化工業所": "0130", "【架空】港北化成": "0141", "【架空】三和珪砂": "0118", "【架空】笹倉溶射工業": "0340",
       "【架空】ときわ環境サービス": "0720"}

def fy_of(d):
    return d.year if d.month >= 4 else d.year - 1

def lines_for_fy(y):
    """(発注日, 取引先, 品名, 規格, 数量, 単位, 単価, 納期, 区分, 契約 cid, task tid, tag, 備考)"""
    r = random.Random(W.SEED + y)
    out = []
    esc = 1 + (y - 2009) * 0.018
    for m, n in ((4, 0), (6, 0), (8, 0), (10, 0), (12, 0), (2, 1)):
        d = dt.date(y + n, m, 3 + r.randint(0, 10))
        out.append((d, "【架空】白浜石灰工業", "消石灰", "JIS特号 ローリー", 30, "t", int(20000 * esc / 100) * 100, d + dt.timedelta(days=4), "薬品", None, None, "", ""))
    for m, n in ((5, 0), (9, 0), (1, 1)):
        d = dt.date(y + n, m, 6 + r.randint(0, 12))
        out.append((d, "【架空】炭化工業所", "粉末活性炭", "フレコン 500kg", 6, "袋", int(92000 * esc / 1000) * 1000, d + dt.timedelta(days=8), "薬品", None, None, "", ""))
        out.append((d + dt.timedelta(days=2), "【架空】港北化成", "アンモニア水 25%", "ローリー", 8, "t", int(39000 * esc / 1000) * 1000, d + dt.timedelta(days=7), "薬品", None, None, "", ""))
    for c in C.ALL:
        if c.fy != y or c.kind == "kikan": continue
        od = c.start - dt.timedelta(days=35)
        md = sum(sum(cr.values()) for t in c.tasks for _, _, cr in t.steps)
        sub = int(md * 40000 * (1 + (y - 2009) * 0.015) / 100000) * 100000 if c.kind == "teishu" else None
        for tid, nm, spec, q, used, unit, price, ven in c.parts:
            t = [x for x in c.tasks if x.tid == tid][0]
            tag = t.tags[0] if t.tags else ""
            kb = "外注" if ("工事" in nm or "補修" in nm) else ("委託" if "試験" in nm else "工事用")
            dd = od if kb != "委託" else c.start - dt.timedelta(days=25)
            note = f"{used:g}使用+予備{q - used:g}" if q > used else ""
            out.append((dd, ven, nm, spec, q, unit, price, c.start - dt.timedelta(days=3), kb, c.cid, tid, tag, note))
        if sub:
            out.append((od + dt.timedelta(days=4), "【架空】北嶺プラント工業", c.short, "工事仕様書による（足場・開放・整備・復旧）", 1, "式", sub, c.end, "外注", c.cid, None, "",
                        "第2期 年度整備上限内" if y >= 2024 else ""))
        for t in c.tasks:
            for d, txt, cr in t.steps:
                if txt.startswith("【追加】"):
                    out.append((d, "【架空】北嶺プラント工業", f"{c.short} 追加工事（{W.BY_TAG[t.tags[0]].name}）", txt.replace("【追加】", "")[:30], 1, "式",
                                int(r.uniform(180, 520)) * 1000, c.end, "外注", c.cid, t.tid, t.tags[0], "変更(追加)"))
        if c.kind == "teishu" and any("珪砂投入" in s[1] for t in c.tasks for s in t.steps):
            out.append((od + dt.timedelta(days=6), "【架空】三和珪砂", "珪砂 6号", "フレコン 1.5t", 6, "袋", int(32000 * esc / 1000) * 1000, c.start, "工事用", c.cid, "T02",
                        c.tasks[2].tags[0] if len(c.tasks) > 2 and c.tasks[2].tags else "", "流動砂 補充"))
    out.sort(key=lambda x: x[0])
    return [x for x in out if fy_of(x[0]) == y or x[9]]

def old_csv(y, path, fname):
    rec = Rec(fname)
    buf = io.StringIO()
    buf.write(f"#【架空】ときわ北部エコサービス 購買管理 発注明細出力 期間 {y}/04/01-{y + 1}/03/31 出力日 {y + 1}/04/0{3 + y % 4}\r\n")
    w = csv.writer(buf, lineterminator="\r\n")
    w.writerow(["HACNO", "HACDT", "TORICD", "TORINM", "HINMEI", "KIKAKU", "SURYO", "TANI", "TANKA", "KINGAKU", "NOUKI", "YOSAN", "KOJINO", "BIKO"])
    hy = y - 1988
    if y == 2015:
        src = [(dt.datetime.strptime(a, "%Y%m%d").date(), "T" + cd, nm, hin, kik, q, u, tk, dt.datetime.strptime(nk, "%Y%m%d").date(), yo, ko, bi)
               for (a, cd, nm, hin, kik, q, u, tk, nk, yo, ko, bi) in OG.OLD]
        rows = []
        for i, (d, cd, nm, hin, kik, q, u, tk, nk, yo, ko, bi) in enumerate(src, 1):
            work = None
            rows.append(([f"H27-{i:04d}", d.strftime("%Y%m%d"), cd, hankaku(nm.replace("【架空】", "").replace("株式会社", "")) + "(ｶ", hin, kik, q, u, tk, q * tk, nk.strftime("%Y%m%d"), yo, ko, bi],
                         "T2015-2" if ko == "K27-02" else ("T2015-1" if ko == "K27-01" else None)))
    else:
        rows = []
        for i, (d, ven, nm, spec, q, u, tk, nk, kb, cid, tid, tag, note) in enumerate(lines_for_fy(y), 1):
            yo = {"薬品": f"S-0-{hy}-11"}.get(kb, f"S-{cid[-1] if cid else 0}-{hy}-01")
            ko = f"K{hy}-{int(cid[-1]):02d}" if cid and cid.startswith("T") else ""
            nm2 = hankaku(nm.replace("ガス化炉", "ｶﾞｽ化炉")) if r_half(i) else nm
            rows.append(([f"H{hy}-{i:04d}", d.strftime("%Y%m%d"), "T" + VEN.get(ven, "0999"), hankaku(ven.replace("【架空】", "")) + "(ｶ", nm2, spec, q, u, tk, round(q * tk),
                          nk.strftime("%Y%m%d"), yo, ko, note], (cid, tid, tag)))
    for k, (row, link) in enumerate(rows):
        w.writerow(row)
        fields = dict(zip(["HACNO", "HACDT", "TORICD", "TORINM", "HINMEI", "KIKAKU", "SURYO", "TANI", "TANKA", "KINGAKU", "NOUKI", "YOSAN", "KOJINO", "BIKO"], row))
        rec.add("item", fields, r=k + 3, work=work_ids(link))
    with open(path, "w", encoding="cp932", newline="") as fp:
        fp.write(buf.getvalue())
    rec.save()

def r_half(i):
    return i % 3 == 0

def work_ids(link):
    if not link: return None
    if isinstance(link, str):
        c = [x for x in C.ALL if x.cid == link][0]
        return [f"{c.cid}:{t.tid}" + (f":{t.tags[0]}" if t.tags else "") for t in c.tasks if t.company == W.SUB1][:3]
    cid, tid, tag = link
    if not cid: return None
    if tid:
        return [f"{cid}:{tid}:{tag}" if tag else f"{cid}:{tid}"]
    return [f"{cid}:外注一式"]

PO_OF = {}

def new_csv(y, path, fname):
    rec = Rec(fname)
    buf = io.StringIO()
    w = csv.writer(buf, quoting=csv.QUOTE_ALL, lineterminator="\r\n")
    w.writerow([f"【架空】ときわ北部エコサービス株式会社 購買システム 発注明細一覧", f"期間:{y}/04/01～{y + 1}/03/31", f"出力:{y + 1}/04/06 sv2"])
    w.writerow(["行種別", "発注番号", "明細No", "発注日／品目コード", "仕入先コード／品名", "仕入先名／規格・型式", "工事件名／数量", "勘定科目／単位",
                "合計金額(税抜)／単価", "消費税合計／金額(税抜)", "承認者／納入予定日", "設備タグ", "発注区分", "備考"])
    tax = 0.08 if y < 2019 or (y == 2019 and False) else 0.10
    yy = y % 100
    groups = OrderedDict()
    if y == 2025:
        for row in OG.NEW:
            groups.setdefault(row[0], []).append(row + (("T2025-1" if row[13] else None),))
    else:
        n = 0; lastkey = None
        for (d, ven, nm, spec, q, u, tk, nk, kb, cid, tid, tag, note) in lines_for_fy(y):
            key = (d, ven, cid)
            if key != lastkey:
                n += 1; lastkey = key
            po = f"PO{yy:02d}-{n * 3 + 9:04d}"
            c = [x for x in C.ALL if x.cid == cid][0] if cid else None
            ko = c.title if c else ""
            ka = {"薬品": "薬品費", "委託": "委託費"}.get(kb, "修繕費")
            kb2 = {"薬品": "通常"}.get(kb, kb)
            tg = tag
            if tg and n % 7 == 3: tg = ""                      # 設備タグ未入力（揺れ）
            elif tg and n % 11 == 5: tg = tg.replace("-0", "-")  # 0 埋めなし（揺れ）
            ln = len(groups.get(po, [])) + 1
            groups.setdefault(po, []).append((po, ln, d.strftime("%Y/%m/%d"), "V" + VEN.get(ven, "0999"), ven, f"P-{zlib.crc32(nm.encode()) % 9000 + 1000}", nm, spec, q, u, tk,
                                              nk.strftime("%Y/%m/%d"), ka, ko, tg, kb2, note, (cid, tid, tag)))
            if cid and tid: PO_OF[(cid, tid, nm)] = po
    r = 3
    for po, rows in groups.items():
        tot = sum(round(x[8] * x[10]) for x in rows)
        h = rows[0]
        appr = "所長" if tot >= 1000000 else "sv2"
        tax = 0.10 if h[2] >= "2019/10/01" else 0.08
        hrow = ["H", po, "", h[2], h[3], h[4], h[13], h[12], tot, round(tot * tax), appr, "", "", ""]
        w.writerow(hrow)
        rec.add("parent", {"行種別": "H", "発注番号": po, "発注日": h[2], "仕入先コード": h[3], "仕入先名": h[4], "工事件名": h[13], "勘定科目": h[12],
                           "合計金額": tot, "消費税合計": round(tot * tax), "承認者": appr}, r=r)
        r += 1
        for x in rows:
            amt = round(x[8] * x[10])
            w.writerow(["D", x[0], x[1], x[5], x[6], x[7], x[8], x[9], x[10], amt, x[11], x[14], x[15], x[16]])
            link = x[17]
            if isinstance(link, str) or link is None:
                link = ("T2025-1", None, None) if link else None
            rec.add("item", {"行種別": "D", "発注番号": x[0], "明細No": x[1], "品目コード": x[5], "品名": x[6], "規格・型式": x[7], "数量": x[8], "単位": x[9],
                             "単価": x[10], "金額": amt, "納入予定日": x[11], "設備タグ": x[14], "発注区分": x[15], "備考": x[16]}, r=r, work=work_ids(link))
            r += 1
    with open(path, "w", encoding="utf-8-sig", newline="") as fp:
        fp.write(buf.getvalue())
    rec.save()

def union_csv(fyw, path, fname):
    rec = Rec(fname)
    buf = io.StringIO()
    w = csv.writer(buf, lineterminator="\r\n")
    w.writerow(["【架空】ときわ北部環境組合 財務会計システム 支出負担行為・支出状況一覧", "事業:基幹的設備改良事業", f"年度:{fyw}", "出力日:R08.11.10" if fyw == "令和08" else "出力日:R08.04.08", "出力者:u1"])
    w.writerow(["年度", "予算科目", "負担行為日", "負担行為番号", "件名", "相手方", "契約方法", "負担行為額", "支出額", "支出日", "区分", "担当課", "摘要"])
    r = 3
    for (fy, d, no, nm, aite, meth, fu, sh, sd, kb, te) in OG.UNI:
        if fy != fyw: continue
        kamoku = "01-02-01-02-15 一般/衛生費/清掃費/施設整備費/工事請負費" if "監理" not in nm else "01-02-01-02-13 一般/衛生費/清掃費/施設整備費/委託料"
        row = [fy, kamoku, d, no, nm, aite, meth, fu if fu else "", sh if sh else "", sd, kb, "施設課", te]
        w.writerow(row)
        cid = "KK2025" if fy == "令和07" else "KK2026"
        rec.add("item", dict(zip(["年度", "予算科目", "負担行為日", "負担行為番号", "件名", "相手方", "契約方法", "負担行為額", "支出額", "支出日", "区分", "担当課", "摘要"], row)),
                r=r, work=[f"{cid}:契約"] if "監理" not in nm else None)
        r += 1
    with open(path, "w", encoding="cp932", newline="") as fp:
        fp.write(buf.getvalue())
    rec.save()
