# -*- coding: utf-8 -*-
"""発注台帳：SPC 購買システムの年度別出力 CSV。
2016〜2019 年度：旧システム（CP932・CRLF・半角カナ・8 桁日付）。2020〜：HIBARI-P2（UTF-8 BOM・ISO 日付・購買依頼No・案件コード・費目）。
2024 年度は見本で承認された明細（orders_src.NEW_ROWS）をそのまま使う。金額は SPC の発注額（税抜）で工事の実績原価ではない。"""
import csv
import datetime as dt
import io
import random
import world as W
import works as WK
import contracts as K
import gen_contract as G
import b_plan
from orders_src import V_NEW, NEW_ROWS
from rec import Rec, path

D = dt.date
OLD_CD = {"V00012": "T0012", "V00013": "T0013", "V00220": "C0220", "V00310": "H0310", "V00405": "M0405", "V00511": "S0511",
          "V00602": "H0602", "V00715": "T0715", "V00820": "H0820", "V00901": "K0901", "V01003": "T1003", "V01201": "Y1201",
          "V01310": "M0405"}   # 旧システムでは軸受も丸仁経由
OLD_NM = {"T0012": "ﾀｶﾊﾞｴﾝｼﾞﾆｱﾘﾝｸﾞ(ｶ)", "T0013": "ﾀｶﾊﾞﾃﾞﾝｷｻｰﾋﾞｽ(ｶ)", "C0220": "ﾁｸｼﾌｨﾙﾀ(ｶ)", "H0310": "ﾊﾔﾌﾞｻﾎﾟﾝﾌﾟｾｲｻｸｼｮ(ｶ",
          "M0405": "ﾏﾙﾆｷｻﾞｲｼｮｳｶｲ(ｶ)", "S0511": "ｻｻﾞﾅﾐｶﾝｷｮｳﾌﾞﾝｾｷ(ｶ", "H0602": "ﾋﾊﾞﾘｾｷﾕ(ｶ)", "T0715": "ﾄｳｴｲﾋﾊｶｲｹﾝｻ(ｶ)",
          "H0820": "ﾋﾊﾞﾘﾃﾞﾝｾﾂ(ｶ)", "K0901": "ﾎｺﾀｹｲｷ(ｶ)", "T1003": "ﾀﾂﾐｷｺｳ(ｶ)", "Y1201": "ﾔﾏﾄﾔｸﾋﾝ(ｶ)"}

ZEN = "アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワヲンァィゥェォッャュョー"
HAN = "ｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖﾗﾘﾙﾚﾛﾜｦﾝｧｨｩｪｫｯｬｭｮｰ"
DAK = {"ガ": "ｶﾞ", "ギ": "ｷﾞ", "グ": "ｸﾞ", "ゲ": "ｹﾞ", "ゴ": "ｺﾞ", "ザ": "ｻﾞ", "ジ": "ｼﾞ", "ズ": "ｽﾞ", "ゼ": "ｾﾞ", "ゾ": "ｿﾞ",
       "ダ": "ﾀﾞ", "ヂ": "ﾁﾞ", "ヅ": "ﾂﾞ", "デ": "ﾃﾞ", "ド": "ﾄﾞ", "バ": "ﾊﾞ", "ビ": "ﾋﾞ", "ブ": "ﾌﾞ", "ベ": "ﾍﾞ", "ボ": "ﾎﾞ",
       "パ": "ﾊﾟ", "ピ": "ﾋﾟ", "プ": "ﾌﾟ", "ペ": "ﾍﾟ", "ポ": "ﾎﾟ", "ヴ": "ｳﾞ"}


def hankana(s):
    out = []
    for ch in s:
        if ch in DAK:
            out.append(DAK[ch])
        elif ch in ZEN:
            out.append(HAN[ZEN.index(ch)])
        else:
            out.append(ch)
    return "".join(out)


def vendor_for(name):
    if "軸受" in name:
        return "V01310"
    if "ろ布" in name:
        return "V00220"
    if "メカニカルシール" in name:
        return "V00310"
    if "作動油" in name or "タービン油" in name:
        return "V00602"
    if "チェーン" in name or "Vベルト" in name or "ベルト" in name:
        return "V00405"
    if "インバータ" in name:
        return "V00013"
    return "V00012"


def esc(y):
    return 1 + 0.018 * (y - 2016) + (0.06 if y >= 2022 else 0)


def rows_for(y, contracts):
    """年度 y の発注明細（中間形式）。"""
    r = random.Random(W.SEED + y * 3)
    out = []

    def add(date, kind, case, nm, sp, q, u, up, due, acc, vc, asset="", memo="", works=None, ans=None):
        out.append(dict(date=date, kind=kind, case=case, name=nm, spec=sp, qty=q, unit=u, price=up, due=due, ans=ans or due,
                        acc=acc, vendor=vc, asset=asset, memo=memo, works=works or []))

    yy = y % 100
    cs = [c for c in contracts if c["y"] == y]
    intasks = set()
    for c in cs:
        for t in c["tasks"]:
            for nm, ff in t[7]:
                intasks.add((nm, ff))
        st = c["stop"]
        case = f"M{yy:02d}-{c['f']:02d}-OH"
        rows_ = K.daily_rows(c["tasks"], y, seed_add={(2024, 1): 24, (2025, 2): 25}.get((y, c["f"]), y * 10 + c["f"]))
        mh = sum(m * h for v in rows_.values() for (t, d_, m, h) in v if t[3] != "SPC運転")
        amt = int(round(mh * 23500 * esc(y), -5))
        cw = sorted({w for t in c["tasks"] for (nm, ff) in t[7] for w in [WK.wid(W.eq(nm, ff), y)] if w})
        end = c["period"][1]
        fy0 = D(y, 4, 3)
        add(max(st - dt.timedelta(days=40), fy0), "外注", case, c["title"], "仕様書のとおり（支給品を除く）", 1, "式", amt, end, end, "V00012",
            memo=f"炉停止 {st.month}/{st.day}〜", works=cw)
        for p in G.parts_for(c):
            if (y, c["f"]) == (2025, 2) and "燃焼ストーカ" in p["name"]:
                continue                      # 2024 年度に先行発注済み（PR24-0101）
            due = st + dt.timedelta(days=r.choice([3, 5, 7, 8]))
            late = r.random() < 0.08
            acc = due + dt.timedelta(days=2) if late else due - dt.timedelta(days=r.choice([0, 1, 2]))
            e = p["eq"]
            memo = "支給品" + (f"　{p['used']}{p['unit']}使用・{p['spare']}{p['unit']}予備" if p["spare"] else "")
            if late:
                memo += "　納入遅れ 2日"
            add(max(st - dt.timedelta(days=55), fy0), "部品", case, p["name"], p["spec"], p["qty"], p["unit"], p["price"], due, acc,
                vendor_for(p["name"]), e.no, memo, [WK.wid(e, y)], ans=acc if late else due)
        for t in c["tasks"]:
            if len(t) > 10 and t[10] and "追加" in t[10]:
                nm, ff = t[7][0]
                e = W.eq(nm, ff)
                add(st + dt.timedelta(days=14), "外注", case, f"{c['title']}（追加）", f"{e.name} 追加補修", 1, "式",
                    int(round(r.choice([680000, 920000, 1240000]) * esc(y), -4)), end, end, "V00012", e.no, "変更契約", [WK.wid(e, y)])
    # 共通設備の経常修繕（定修の工程に入らないもの）
    for e in W.EQ:
        a = W.actual_work(e, y)
        if not a or (e.name, e.f) in intasks or (e.name, e.f, y) in W.CITY_WORKS:
            continue
        k = next((k for k in "XUOR" if k in a), None)
        if not k:
            continue
        if e.f != 0:
            continue
        m = W.oh_month(0, y, e.seq)
        yr = y + 1 if m <= 3 else y
        dd = D(yr, m, 1) + dt.timedelta(days=r.randint(0, 9))
        if y == 2026 and dd > D(2026, 11, 10):
            continue
        base = {"X": b_plan.COST_X, "U": b_plan.COST_X, "O": b_plan.COST_O, "R": b_plan.COST_R}[k].get(e.strat, 6000 if k in "XU" else 800)
        if (e.name, y) == ("DCS中央オペレータステーション", 2019):
            base = 38500
        vc = {"comp": "V01003", "scale": "V00901", "pump_r": "V00310", "pump_s": "V00310", "pump_c": "V00310", "panel": "V00013",
              "dcs": "V00013", "hv": "V00820", "tr": "V00820", "ups": "V00013", "dc": "V00013"}.get(e.strat, "V00012")
        verb = {"X": "更新", "U": "更新（HMI）", "O": "分解整備", "R": "補修"}[k]
        add(dd, "外注", f"M{yy:02d}-K", f"{e.name} {verb}", e.spec[:20], 1, "式", int(round(base * 1000 * esc(y), -4)),
            dd + dt.timedelta(days=45), dd + dt.timedelta(days=r.randint(30, 50)), vc, e.no, "", [WK.wid(e, y)])
    # 法定点検・測定
    legal = [(4, "排ガス測定 第1回（ばいじん・NOx・SOx・HCl）", "1・2号煙道", 380000, "V00511"),
             (4, "ダイオキシン類 作業環境測定 前期", "炉室・灰出し室ほか", 420000, "V00511"),
             (6, "クレーン " + ("性能検査（受検・立会）" if y % 2 == 1 else "年次自主検査"), "ごみクレーン2基・灰クレーン・タービン室・ホイスト", 860000 if y % 2 == 1 else 660000, "V00012"),
             (7, "消防用設備等点検 機器点検", "全館", 290000, "V00820"),
             (8, "非常用発電機 点検（負荷試験含む）", "500kVA", 340000, "V00820"),
             (10, "排ガス測定 第2回（ばいじん・NOx・SOx・HCl）", "1・2号煙道", 380000, "V00511"),
             (10, "ダイオキシン類 排ガス測定", "1・2号煙道", 640000, "V00511"),
             (10, "ダイオキシン類 作業環境測定 後期", "炉室・灰出し室ほか", 420000, "V00511"),
             (11, "受変電設備 年次点検", "全停電", 720000, "V00820"),
             (1, "消防用設備等点検 総合点検", "全館", 410000, "V00820")]
    if y % 2 == 0:
        legal.append((6, "計量機 定期検査（受検手続・立会）", "入口・出口 30t", 230000, "V00901"))
    if y in (2017, 2020, 2023, 2026):
        legal.append((9, "精密機能検査", "施設全体", 1650000, "V00511"))
    for m, nm, sp, up, vc in legal:
        yr = y + 1 if m <= 3 else y
        dd = D(yr, m, 5)
        if y == 2026 and dd > D(2026, 11, 10):
            continue
        add(dd, "法定点検", f"M{yy:02d}-L", nm, sp, 1, "式", int(round(up * esc(y), -3)), dd + dt.timedelta(days=50),
            dd + dt.timedelta(days=r.randint(30, 50)), vc)
    for m in (5, 8, 11, 2):
        yr = y + 1 if m <= 3 else y
        dd = D(yr, m, 6)
        if y == 2026 and dd > D(2026, 11, 10):
            continue
        add(dd, "薬品", f"M{yy:02d}-C", "消石灰", "JIS特号 バラ", r.choice([58, 60, 62, 64]), "t", int(round(23500 * esc(y), -2)),
            dd + dt.timedelta(days=24), dd + dt.timedelta(days=24), "V01201", memo=f"{m}月分")
        add(dd, "薬品", f"M{yy:02d}-C", "キレート剤", "液体 1t コンテナ", 4, "個", int(round(312000 * esc(y), -3)),
            dd + dt.timedelta(days=24), dd + dt.timedelta(days=24), "V01201", memo=f"{m}月分")
    dd = D(y + 1, 1, 15)
    if y < 2026:
        add(dd, "消耗品", f"M{yy:02d}-C", "タービン油", "VG32 200L", 2, "缶", int(round(98000 * esc(y), -3)), dd + dt.timedelta(days=23),
            dd + dt.timedelta(days=21), "V00602", "HB05-003")
    for o in out:
        if y == 2026 and o["acc"] > D(2026, 11, 10):
            o["acc"] = None
    return sorted(out, key=lambda o: (o["date"], o["case"]))


def sample_2024():
    out = []
    for (pr, dd, cs, kind, code, nm, sp, q, u, up, want, ans, acc, vc, asset, memo) in NEW_ROWS:
        ws = []
        if asset:
            for e in W.EQ:
                if e.no == asset:
                    w = WK.wid(e, 2024 if dd < "2025-04" else 2025)
                    if w:
                        ws.append(w)
        if kind == "外注" and "定期整備工事" in nm and "追加" not in nm:
            f = 1 if "1号炉" in nm else 2
            c = G.build(2024, f)
            ws = sorted({w for t in c["tasks"] for (n_, ff) in t[7] for w in [WK.wid(W.eq(n_, ff), 2024)] if w})
        if "追加" in nm:
            ws = [WK.wid(W.eq("誘引通風機", 1), 2024), WK.wid(W.eq("灰押出装置", 1), 2024)]
        out.append(dict(pr=pr, date=D.fromisoformat(dd), case=cs, kind=kind, code=code, name=nm, spec=sp, qty=q, unit=u, price=up,
                        due=D.fromisoformat(want), ans=D.fromisoformat(ans) if ans else None, acc=D.fromisoformat(acc) if acc else None,
                        vendor=vc, asset=asset, memo=memo, works=ws))
    return out


def fname(y):
    sysn = "旧購買システムCSV" if y <= 2019 else "HIBARI-P2購買システムCSV"
    return f"【架空】ひばり野クリーンセンター_発注台帳_{y}年度_{sysn}.csv"


def write_old(y, rows):
    fn = fname(y)
    R = Rec(fn)
    buf = io.StringIO(newline="")
    buf.write(f"【架空】ひばり野環境サービス㈱　購買システム　発注一覧　年度={y}　出力日={y + 1}0408\r\n")
    w = csv.writer(buf, lineterminator="\r\n")
    head = ["発注番号", "発注日", "発注区分", "工事番号", "品名", "規格", "数量", "単位", "単価", "金額", "納入期限", "業者CD", "業者名", "予算科目", "摘要"]
    w.writerow(head)
    n = (y % 100) * 10000 + 1
    line = 3
    for o in rows:
        kb = {"外注": "工事", "部品": "物品", "法定点検": "役務", "薬品": "物品", "消耗品": "物品"}[o["kind"]]
        if o["kind"] == "外注" and "分解整備" not in o["name"] and "補修" not in o["name"] and "工事" not in o["name"]:
            kb = "役務"
        kj = ""
        if o["case"].endswith("-OH"):
            kj = f"{y % 100:02d}-{int(o['case'][4:6])}OH"
        bud = {"外注": "修繕費", "部品": "修繕費", "法定点検": "法定点検費", "薬品": "薬品費", "消耗品": "消耗品費"}[o["kind"]]
        vc = OLD_CD[o["vendor"]]
        memo = o["memo"].replace("　", " ")
        row = [n, o["date"].strftime("%Y%m%d"), kb, kj, hankana(o["name"]), hankana(o["spec"]), o["qty"], o["unit"], o["price"],
               o["qty"] * o["price"], o["due"].strftime("%Y%m%d"), vc, OLD_NM[vc], bud, hankana(memo)]
        w.writerow(row)
        rt = [v.encode("cp932").decode("cp932") if isinstance(v, str) else v for v in row]
        R.add(None, line, "item", dict(zip(head, rt)), [x for x in o["works"] if x])
        n += 1
        line += 1
    with open(path(fn), "wb") as f:
        f.write(buf.getvalue().encode("cp932"))
    return fn


def write_new(y, rows, sample=False):
    fn = fname(y)
    R = Rec(fn)
    out_d = "2026-11-10 09:40" if y == 2026 else f"{y + 1}-04-08 10:12"
    endp = "2026-11-10" if y == 2026 else f"{y + 1}-03-31"
    buf = io.StringIO(newline="")
    buf.write(f"# 【架空】ひばり野環境サービス株式会社 購買管理システム HIBARI-P2 発注明細エクスポート 対象期間:{y}-04-01〜{endp} 出力:{out_d}\n")
    w = csv.writer(buf, lineterminator="\n", quoting=csv.QUOTE_NONNUMERIC)
    head = ["購買依頼No", "発注No", "行", "発注日", "部門", "案件コード", "費目", "品目コード", "品名", "仕様・型式", "数量", "単位",
            "単価", "金額", "消費税", "希望納期", "回答納期", "検収日", "仕入先コード", "仕入先名", "資産番号", "備考"]
    w.writerow(head)
    po = {}
    seq = (y % 100) * 1000000 + 100
    prn = 0
    line = 3
    tax = 0.08 if y < 2019 or (y == 2019 and False) else 0.10
    for o in rows:
        if sample:
            pr = o["pr"]
        else:
            key = (o["case"], o["date"], o["kind"] == "外注")
            if key not in po:
                prn += 1
            pr = f"PR{y % 100:02d}-{prn:04d}" if key not in po else po[key][2]
        k2 = pr
        if k2 not in po:
            seq += 1
            po[k2] = [f"PO{seq}", 0, pr]
        if not sample:
            po[(o["case"], o["date"], o["kind"] == "外注")] = po[k2]
        po[k2][1] += 1
        amt = o["qty"] * o["price"]
        code = o.get("code", "")
        row = [pr, po[k2][0], po[k2][1], o["date"].isoformat(), "保全課", o["case"], o["kind"], code, o["name"], o["spec"], o["qty"],
               o["unit"], o["price"], amt, int(amt * (0.10 if o["date"] >= D(2019, 10, 1) else 0.08)), o["due"].isoformat(),
               o["ans"].isoformat() if o["ans"] else "", o["acc"].isoformat() if o["acc"] else "", o["vendor"], V_NEW[o["vendor"]],
               o["asset"], o["memo"]]
        w.writerow(row)
        R.add(None, line, "item", {h: v for h, v in zip(head, row) if v not in ("", None)}, [x for x in o["works"] if x])
        line += 1
    with open(path(fn), "wb") as f:
        f.write(b"\xef\xbb\xbf" + buf.getvalue().encode("utf-8"))
    return fn


def build_all(contracts):
    out = []
    for y in range(2016, 2027):
        if y == 2024:
            out.append(write_new(y, sorted(sample_2024(), key=lambda o: (o["date"], o["pr"])), sample=True))
        elif y <= 2019:
            out.append(write_old(y, rows_for(y, contracts)))
        else:
            out.append(write_new(y, rows_for(y, contracts)))
    return out
