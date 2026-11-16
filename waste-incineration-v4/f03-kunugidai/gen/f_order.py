# -*- coding: utf-8 -*-
"""発注台帳（全年度）。旧財務会計システム CSV（H13〜H25、CP932・和暦）、新財務会計システム CSV（H26〜、UTF-8 BOM・西暦）、
直営用 部品発注控（H24〜、k2 作成・R4 から k3、年度ごとにシートを複写）。金額は契約額（税込）で、工事の原価ではない。"""
import csv, io, random
import datetime as dt
from openpyxl import Workbook
from openpyxl.comments import Comment
import xlhelp as X
from rec import Rec
from model import WORKS, tax, workday, add_workdays, CUTOFF
from world import SEED, wareki, CLS

F = X.font
VEND = {"汐見重工環境": ("40102", "汐見重工環境（株）", "汐見重工環境株式会社"), "槇野": ("40211", "槇野環境分析センター（株）", "槇野環境分析センター株式会社"),
        "結城": ("40987", "結城電気管理事務所", "結城電気管理事務所"), "楓": ("41120", "楓防災設備（株）", "楓防災設備株式会社"),
        "篠目": ("40455", "篠目クレーン（株）", "篠目クレーン株式会社"), "白縫": ("40630", "白縫フェルト工業（株）", "白縫フェルト工業株式会社"),
        "杠葉": ("40731", "杠葉ポンプ製作所（株）", "杠葉ポンプ製作所株式会社"), "苅安": ("41377", "（有）苅安工業", "有限会社 苅安工業"),
        "鶉野": ("41502", "鶉野環境技術（株）", "鶉野環境技術株式会社"), "沖津": ("40815", "沖津化学（株）", "沖津化学株式会社"),
        "椿": ("40833", "椿石油（株）", "椿石油株式会社"), "柊": ("40920", "柊分析計器（株）", "柊分析計器株式会社"), "恩田": ("41011", "恩田精機（株）", "恩田精機株式会社")}

HIKAE_ITEMS = {"pump_s": ("メカニカルシール", "%s 用", "組", 85000), "pump_m": ("メカニカルシール", "%s 用", "組", 187000),
               "conv": ("チェーン", "%s 用 RS80", "m", 42000), "misc": ("消耗部品", "%s 用", "式", 36000), "inst": ("検出器", "%s 用", "台", 98000),
               "inst_a": ("電極", "%s 用", "本", 55000), "camera": ("カメラ", "%s 用 ケーシング付", "台", 160000), "door": ("油圧ホース", "%s 用", "本", 28000),
               "fan": ("Vベルト", "%s 用 B-65", "本", 2310), "comp": ("吸込フィルタ", "%s 用", "個", 7700), "hyd": ("作動油", "VG46 20L", "缶", 11000),
               "feeder": ("摺動板", "%s 用", "枚", 64000), "conv2": ("鋼板", "SS400 t9 1219×2438", "枚", 38500)}
CONSUM = [("グランドパッキン", "PTFE 10角", "m", 4400, "苅安"), ("溶接棒", "B-14 φ3.2 5kg", "箱", 6600, "苅安"), ("ボルト・ナット", "SUS M16×60", "組", 330, "苅安"),
          ("pH 電極", "柊 HP-2", "本", 55000, "柊"), ("オイルシール", "TC 45-62-8", "個", 1320, "苅安"), ("グリス", "リチウム 16kg", "缶", 14300, "苅安"),
          ("Vベルト", "B-65", "本", 2310, "苅安"), ("吸込フィルタ", "恩田 AF-55", "個", 7700, "恩田")]


def _hno(fy, n):
    return ("%s-%03d" % (wareki(fy, True), n))


def build_hikae(model):
    """控の行と、作業 id → 控No の対応"""
    rows = {}
    pmap = {}
    for fy in range(2012, 2027):
        rng = random.Random("%s-hk-%d" % (SEED, fy))
        lst = []
        dws = sorted({w["id"]: w for w in WORKS.values() if w["contract"] == "DIRECT" and w["fy"] == fy and w["kind"] in ("OVERHAUL", "REPAIR", "REPLACE")}.values(),
                     key=lambda w: w["start"])
        for w in dws:
            e = model.tag[w["tag"]]
            key = e.cls if e.cls in HIKAE_ITEMS else "misc"
            if w["tag"] == "1AS-001":
                key = "conv2"
            nm, spec, unit, price = HIKAE_ITEMS[key]
            q = 1 if unit in ("組", "台", "枚", "式") else rng.randint(2, 6)
            od = w["start"] - dt.timedelta(days=rng.randint(3, 20))
            if w["tag"] == "1BL-021" and fy == 2024:
                lst.append(dict(date=dt.date(2024, 9, 27), name="メカニカルシール", spec="杠葉 MS-38B（BFP 用）", q=1, unit="組", amt=187000, vend="杠葉",
                                od=dt.date(2024, 9, 30), dd=dt.date(2024, 10, 11), use="1号 給水ポンプA", tag="1BL-021", st="使用済", memo="10/16 交換（直営）", wid=w["id"], grp="BFP"))
                lst.append(dict(date=dt.date(2024, 9, 27), name="メカニカルシール", spec="杠葉 MS-38B（BFP 用）", q=1, unit="組", amt=187000, vend="杠葉",
                                od=dt.date(2024, 9, 30), dd=dt.date(2024, 10, 11), use="（予備）給水ポンプ共通", tag="", st="予備", memo="棚B-3", wid=None, grp="BFP"))
                continue
            vend = "杠葉" if key in ("pump_s", "pump_m") else ("柊" if key in ("inst", "inst_a") else ("恩田" if key == "comp" else "苅安"))
            lst.append(dict(date=od, name=nm, spec=spec % e.name if "%s" in spec else spec, q=q, unit=unit, amt=price * q, vend=vend,
                            od=od + dt.timedelta(days=1), dd=od + dt.timedelta(days=rng.randint(2, 12)),
                            use=("%d号 " % e.unit if e.unit else "") + e.name, tag=e.tag, st=rng.choice(["使用済", "使用済", "一部使用"]),
                            memo="", wid=w["id"], grp=None))
        for _ in range(rng.randint(5, 9)):
            nm, spec, unit, price, vend = rng.choice(CONSUM)
            d = dt.date(fy, rng.choice([4, 5, 6, 7, 8, 9, 10, 11, 12]), rng.randint(1, 27))
            if fy == 2026 and d > CUTOFF:
                continue
            q = rng.randint(1, 10)
            e = rng.choice([x for x in model.eqs if not x.idle and x.cls in ("pump_s", "conv", "fan", "comp", "misc")])
            lst.append(dict(date=d, name=nm, spec=spec, q=q, unit=unit, amt=price * q, vend=vend, od=d, dd=d + dt.timedelta(days=rng.randint(1, 5)),
                            use=("%d号 " % e.unit if e.unit else "") + e.name, tag=e.tag, st=rng.choice(["使用済", "在庫", "予備"]), memo="", wid=None, grp=None))
        lst.sort(key=lambda x: x["date"])
        for i, x in enumerate(lst, 1):
            x["no"] = _hno(fy, i * 3 - rng.randint(0, 2))
            if x["wid"]:
                pmap[x["wid"]] = x["no"]
        rows[fy] = lst
    return rows, pmap


def write_hikae(rows, path, fname):
    rec = Rec(fname)
    wb = Workbook()
    wb.remove(wb.active)
    X.props(wb, "【架空】くぬぎ台環境センター 直営補修 部品発注控", "くぬぎ台市 環境センター k2（R4〜 k3）", "発注台帳", "H24〜 年度ごとにシートをコピー")
    for fy in sorted(rows):
        lst = rows[fy]
        ws = wb.create_sheet(wareki(fy, True))
        newcols = fy >= 2022
        hdr = ["控No", "依頼日", "品名", "規格・型式", "数量", "単位", "金額(税込)", "発注先", "発注日", "納品日", "使用機器"] + (["タグNo"] if newcols else []) + ["状態", "備考"] + (["財務契約番号"] if newcols else [])
        wid = [8, 10, 14, 22, 5, 5, 10, 14, 10, 10, 18] + ([9] if newcols else []) + [11, 22] + ([16] if newcols else [])
        X.widths(ws, wid)
        X.put(ws, 1, 1, "【架空】直営補修　部品発注控（%s）" % wareki(fy), F(size=13, bold=True), X.NOWRAP)
        X.put(ws, 2, 1, "くぬぎ台環境センター 施設係" + ("　※少額の物品は財務システムの契約番号を書く（単価契約・消耗品は空欄）" if newcols else ""), F(size=9), X.NOWRAP)
        for i, h in enumerate(hdr, 1):
            X.put(ws, 4, i, h, F(bold=True, size=9), X.CENTER, X.fill("E2EFDA" if h not in ("タグNo", "財務契約番号") else "FFF2CC"))
        if newcols:
            ws.cell(4, 12).comment = Comment("R4 追加（k3）台帳統合に合わせて", "k3")
            ws.cell(4, len(hdr)).comment = Comment("R4 追加（k3）二重計上を避けるため", "k3")
        r = 5
        tot = 0
        for x in lst:
            vals = [x["no"], x["date"].strftime("%Y/%m/%d"), x["name"], x["spec"], x["q"], x["unit"], x["amt"], x["vend"] + ("ポンプ製作所" if x["vend"] == "杠葉" else ("工業" if x["vend"] == "苅安" else ("分析計器" if x["vend"] == "柊" else "精機"))),
                    "単価契約" if x["name"] == "吸込フィルタ" else x["od"].strftime("%Y/%m/%d"), x["dd"].strftime("%Y/%m/%d"), x["use"]]
            if newcols:
                vals.append(x["tag"])
            vals += [x["st"], x["memo"]]
            if newcols:
                vals.append(x.get("csvno", ""))
            for i, v in enumerate(vals, 1):
                X.put(ws, r, i, v, F(size=8), X.CENTER if i in (1, 2, 5, 6, 9, 10) else X.LEFT, nf="#,##0" if i == 7 else None)
            f = dict(zip(hdr, vals))
            f["金額(税込)"] = x["amt"]
            rec.add(ws.title, r, "item", f, works=[x["wid"]])
            tot += x["amt"]
            r += 1
        r += 2
        X.put(ws, r, 6, "計", F(bold=True, size=9), X.RIGHT)
        X.put(ws, r, 7, tot, F(bold=True, size=9), X.RIGHT, nf="#,##0")
        rec.add(ws.title, r, "total", {"計": tot})
        X.box(ws, 4, 1, r, len(hdr), inner=X.hair)
        X.page(ws, "landscape", "A4", 1, 0, titles="4:4")
    wb.save(path)
    rec.save()


def _wdate(d):
    return "%s.%02d.%02d" % ("H%d" % (d.year - 1988), d.month, d.day)


def year_rows(fy, model, hikae):
    """その年度の契約の行：dict(no, name, kind, method, vend, cdate, a, b, amt, amt2, paid, budget, note, ck, insp)"""
    rng = random.Random("%s-csv-%d" % (SEED, fy))
    out = []
    infl = 1 + (fy - 2001) * 0.012
    t = lambda d: tax(fy, d)
    a0, b0 = dt.date(fy, 4, 1), dt.date(fy + 1, 3, 31)
    for nm, vk, base, meth, note in [("排ガス等測定業務委託", "槇野", 3500, "指名競争", "単価契約 ばい煙 年6回×3炉・DXN"),
                                     ("作業環境測定業務委託", "槇野", 520, "随意契約", "DXN 年2回"),
                                     ("電気工作物保安管理業務委託", "結城", 1340, "随意契約", ""),
                                     ("消防用設備等点検業務委託", "楓", 380, "随意契約", ""),
                                     ("クレーン保守点検業務委託", "篠目", 2650, "随意契約", "年次自主検査含む")]:
        amt = int(round(base * infl * (1 + t(a0)) / 10.0)) * 10000
        out.append(dict(name="%s環境センター%s" % (wareki(fy), nm), kind="委託", method=meth, vend=vk, cdate=a0, a=a0, b=b0, amt=amt, amt2=None,
                        budget="委託料", note=note, ck="ITAKU"))
    if fy % 3 == 2:
        out.append(dict(name="%s環境センター精密機能検査業務委託" % wareki(fy), kind="委託", method="指名競争", vend="鶉野", cdate=dt.date(fy, 7, 10),
                        a=dt.date(fy, 7, 10), b=dt.date(fy + 1, 2, 28), amt=int(round(2400 * infl * (1 + t(a0)) / 10.0)) * 10000, amt2=None, budget="委託料",
                        note="廃掃法施行規則 4条の5", ck="ITAKU"))
    for nm, vk, base in [("消石灰", "沖津", 9800), ("活性炭", "沖津", 4200), ("苛性ソーダ（25%）", "沖津", 6100), ("アンモニア水（25%）", "沖津", 1900),
                         ("キレート剤", "沖津", 3300), ("灯油", "椿", 5600)]:
        if fy < 2002 and nm in ("活性炭", "アンモニア水（25%）", "キレート剤"):
            continue
        out.append(dict(name="%s環境センター用%s購入（単価契約）" % (wareki(fy), nm), kind="物品", method="指名競争", vend=vk, cdate=a0, a=a0, b=b0,
                        amt=int(round(base * infl * rng.uniform(0.9, 1.1) * (1 + t(a0)) / 10.0)) * 10000, amt2=None, budget="需用費（薬品費）" if nm != "灯油" else "需用費（燃料費）",
                        note="単価契約・見込数量", ck="BUPPIN"))
    for c in model.by_fy(fy):
        if c.kind in ("kairyo_part",):
            continue
        meth = c.method if c.kind != "kairyo" else c.method
        out.append(dict(name=c.name, kind="委託" if c.kind in ("teishu", "kyotsu") or c.key.startswith("AMSTOP") else "工事", method="随意契約（特命）" if c.kind in ("teishu", "kyotsu", "hoshu") else meth,
                        vend="汐見重工環境", cdate=c.cdate, a=c.start, b=c.end, amt=c.amount, amt2=c.amount2,
                        budget="委託料" if c.kind in ("teishu", "kyotsu") or c.key.startswith("AMSTOP") else "工事請負費", note=c.note, ck=c.key))
        # ろ布を市が支給する定修（H26 から）
        if c.kind == "teishu" and fy >= 2014:
            for tk in c.tasks:
                for wid in tk.works:
                    w = WORKS[wid]
                    if model.tag[w["tag"]].cls == "bf" and w["kind"] == "OVERHAUL":
                        dd = c.start - dt.timedelta(days=3)
                        out.append(dict(name="%d号炉バグフィルタ用ろ布購入" % c.unit, kind="物品", method="指名競争", vend="白縫", cdate=c.start - dt.timedelta(days=100),
                                        a=c.start - dt.timedelta(days=100), b=dd, amt=int(round(36000 * infl * 300 * (1 + t(dd)) / 10000.0)) * 10000, amt2=None,
                                        budget="需用費（修繕料）", note="PPS/PTFE φ150×6000 300本（%d号炉定修へ支給、残は予備）" % c.unit, ck="BUPPIN", wid=wid))
    if fy == 2013:
        out.append(dict(name="環境センター　バグフィルタ用ろ布（予備品）購入", kind="物品", method="随意契約", vend="白縫", cdate=dt.date(2013, 12, 18),
                        a=dt.date(2013, 12, 18), b=dt.date(2014, 2, 28), amt=504000, amt2=None, budget="需用費", note="PPS φ150×6000 12本", ck="BUPPIN"))
    # 直営の部品（H24 から。1 件 3 万円以上は財務システムで契約）
    if fy >= 2012:
        grp = {}
        for x in hikae.get(fy, []):
            if x["amt"] < 30000:
                continue
            k = x.get("grp") or x["no"]
            grp.setdefault(k, []).append(x)
        for k, xs in grp.items():
            x = xs[0]
            amt = sum(y["amt"] for y in xs)
            out.append(dict(name="%s購入" % (x["name"] if x["name"] != "メカニカルシール" else "ボイラ給水ポンプ用メカニカルシール" if x.get("grp") else x["spec"].replace(" 用", "用") + x["name"]),
                            kind="物品", method="随意契約（少額）", vend=x["vend"], cdate=x["od"], a=x["od"], b=x["dd"], amt=amt, amt2=None,
                            budget="需用費（修繕料）", note="直営補修用" + ("（%d組、1組予備）" % len(xs) if len(xs) > 1 else ""), ck="HIKAE", hk=xs))
    out.sort(key=lambda r: (r["cdate"], r["name"]))
    if fy == 2026:
        out = [r for r in out if r["cdate"] <= dt.date(2026, 11, 10)]
    return out


def write_old(fy, rows, path, fname):
    rec = Rec(fname)
    buf = io.StringIO()
    w = csv.writer(buf, lineterminator="\r\n")
    buf.write("【架空】くぬぎ台市財務会計システム　契約台帳一覧　所属：2101 環境センター　%s　出力日 H%d.04.15\r\n" % (wareki(fy), fy + 1 - 1988))
    hdr = ["契約番号", "件名", "契約区分", "契約方法", "業者番号", "業者名", "契約年月日", "履行期間自", "履行期間至", "当初契約額", "変更後契約額", "支出額", "予算節", "備考"]
    w.writerow(hdr)
    for i, r in enumerate(rows, 1):
        code, nm, _ = VEND[r["vend"]]
        no = "%02d-2101-%04d" % (fy - 1988, i * 3)
        r["no"] = no
        bud = {"委託料": "13委託料", "工事請負費": "15工事請負費"}.get(r["budget"], "11需用費")
        vals = [no, r["name"], r["kind"], r["method"].replace("（特命）", "").replace("（少額）", ""), code, nm, _wdate(r["cdate"]), _wdate(r["a"]), _wdate(r["b"]),
                r["amt"], r["amt2"] or "", r["amt2"] or r["amt"], bud, r["note"]]
        vals = [v.encode("cp932", "replace").decode("cp932") if isinstance(v, str) else v for v in vals]
        w.writerow(vals)
        rec.add("", i + 2, "item", dict(zip(hdr, vals)), works=_works(r))
    buf.write("件数,%d\r\n" % len(rows))
    rec.add("", len(rows) + 3, "total", {"件数": len(rows)})
    with open(path, "wb") as f:
        f.write(buf.getvalue().encode("cp932"))
    rec.save()


def write_new(fy, rows, path, fname):
    rec = Rec(fname)
    buf = io.StringIO()
    w = csv.writer(buf, lineterminator="\r\n", quoting=csv.QUOTE_ALL)
    od = "2026/11/10 08:47" if fy == 2026 else "%d/04/08 09:12" % (fy + 1)
    buf.write("# 【架空】くぬぎ台市 財務会計システム（契約管理） 契約一覧 抽出条件：所管課=0601 環境センター／年度=%d 出力日時 %s\r\n" % (fy, od))
    hdr = ["契約番号", "年度", "所管課", "契約件名", "契約種別", "契約方法", "相手方コード", "相手方名称", "契約締結日", "履行期間開始", "履行期間終了",
           "当初契約金額", "変更契約金額", "変更回数", "支払済額", "予算科目", "検査完了日", "備考"]
    w.writerow(hdr)
    zen = str.maketrans("0123456789", "０１２３４５６７８９")
    for i, r in enumerate(rows, 1):
        code, _, nm = VEND[r["vend"]]
        no = "%d-0601-%05d" % (fy, i * 3 + 2)
        r["no"] = no
        done = r["b"] <= dt.date(2026, 9, 30) if fy == 2026 else True
        insp = add_workdays(r["b"], 5) if done else None
        name = r["name"].replace("くぬぎ台環境センター", "くぬぎ台環境センター")
        name = name.translate(zen)
        kind = {"委託": "業務委託", "工事": "工事", "物品": "物品"}[r["kind"]]
        vals = [no, str(fy), "環境センター", name, kind, r["method"].replace("指名競争", "指名競争入札").replace("一般競争入札", "一般競争入札"), "1" + code, nm,
                r["cdate"].strftime("%Y/%m/%d"), r["a"].strftime("%Y/%m/%d"), r["b"].strftime("%Y/%m/%d"), r["amt"], r["amt2"] or "",
                1 if r["amt2"] else 0, (r["amt2"] or r["amt"]) if done else 0, r["budget"], insp.strftime("%Y/%m/%d") if insp else "", r["note"]]
        w.writerow(vals)
        rec.add("", i + 2, "item", dict(zip(hdr, vals)), works=_works(r))
        for x in r.get("hk", []):
            x["csvno"] = no
    with open(path, "wb") as f:
        f.write(b"\xef\xbb\xbf" + buf.getvalue().encode("utf-8"))
    rec.save()


def _works(r):
    if r.get("wid"):
        return [r["wid"]]
    if r.get("hk"):
        return [x["wid"] for x in r["hk"] if x["wid"]]
    ck = r["ck"]
    return [w["id"] for w in WORKS.values() if w["contract"] == ck][:200]
