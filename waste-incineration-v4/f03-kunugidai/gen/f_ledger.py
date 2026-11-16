# -*- coding: utf-8 -*-
"""機器台帳。
旧：k2 が H26（2014）に再編した炉別シートの台帳（令和3年度末 時点の版）。H15（k1）の台帳 No を残す。
新：k3 が R4（2022）に 1 シートへ統合した台帳（令和8年10月 時点）。旧シート名・旧 No を持つ。
"""
import random, os
from openpyxl import Workbook
from openpyxl.comments import Comment
from world import SYS, MAKERS, KIKAN_FY, wareki, SEED
import xlhelp as X
from rec import Rec

F = X.font

# k2 台帳（H26 再編）の呼び名。k3 統合時（R4）に今の名前へ直した
K2_NAMES = {"誘引通風機": "誘引送風機", "バグフィルタ": "ろ過式集じん器", "湿式洗煙塔": "洗煙塔",
            "ごみクレーン 1号": "ごみクレーン No.1", "ごみクレーン 2号": "ごみクレーン No.2",
            "押込送風機": "押込送風機", "触媒反応塔": "脱硝反応塔", "蒸気式空気予熱器": "空気予熱器",
            "減温塔": "ガス減温塔", "灰溶融炉本体": "灰溶融炉"}

K1_SYS = {"RF": "受", "IN": "燃", "BL": "ボ", "GT": "ガ", "FD": "通", "AS": "灰", "FA": "灰", "ST": "発",
          "CH": "薬", "WW": "排", "UT": "用", "EL": "電", "IC": "計", "MS": "雑", "AM": "溶"}


def _month(e, fy, salt):
    rng = random.Random("%s-%s-%d" % (SEED, e.tag, fy + salt))
    return rng.choice([6, 9, 10, 11, 12, 1, 2, 3])


def _ym(fy, m):
    y = fy if m >= 4 else fy + 1
    if y >= 2019 and not (y == 2019 and m < 5):
        return "R%d.%d" % (y - 2018, m)
    return "H%d.%d" % (y - 1988, m)


def _last_rep(e, upto):
    ys = [x["fy"] for x in e.events if x["fy"] <= upto and x["kind"] in ("REPLACE",)]
    return max(ys) if ys else None


def _k1_no(e, seq):
    # H15（2003）に k1 が付けた番号。「1-燃-05」の形。共通は「C-」。H15 以後に追設したものは無い
    if e.install > 2003 and not e.idle:
        return ""
    if e.idle:
        return "C-溶-%02d" % seq if e.install <= 2003 else ""
    return "%s-%s-%02d" % ("C" if e.unit == 0 else e.unit, K1_SYS[e.sys], seq)


def make_old(eqs, path, upto=2021, fname=None):
    rec = Rec(fname or os.path.basename(path))
    asof = "R4.3.31" if upto == 2021 else "H26.4.1"
    eqs = [e for e in eqs if e.install <= upto]
    wb = Workbook()
    X.props(wb, "【架空】くぬぎ台環境センター 機器台帳（%s）" % ("令和3年度末" if upto == 2021 else "平成26年4月 再編"), "くぬぎ台市 環境センター k2",
            "機器台帳", "H26 再編（炉別シート）。H15 台帳 No を併記")
    ws0 = wb.active
    ws0.title = "目次"
    X.widths(ws0, [4, 22, 10, 40])
    X.put(ws0, 1, 1, "【架空】くぬぎ台環境センター　機器台帳", F(size=16, bold=True))
    X.put(ws0, 2, 1, "くぬぎ台市環境部 環境センター 施設係　　H26.4 再編（k2）／ %s 現在" % asof, F(size=10))
    sheets = [("1号炉", lambda e: e.unit == 1), ("2号炉", lambda e: e.unit == 2), ("3号炉", lambda e: e.unit == 3),
              ("共通", lambda e: e.unit == 0 and not e.idle), ("灰溶融（休止）", lambda e: e.idle)]
    X.put(ws0, 4, 2, "シート", F(bold=True), X.CENTER, X.fill("DDEBF7"))
    X.put(ws0, 4, 3, "台数", F(bold=True), X.CENTER, X.fill("DDEBF7"))
    X.put(ws0, 4, 4, "備考", F(bold=True), X.CENTER, X.fill("DDEBF7"))
    rr = 5
    tot = 0
    for sname, pred in sheets:
        g = [e for e in eqs if pred(e)]
        tot += len(g)
        X.put(ws0, rr, 2, sname, F(), X.LEFT)
        X.put(ws0, rr, 3, len(g), F(), X.RIGHT)
        X.put(ws0, rr, 4, {"灰溶融（休止）": "H25.10 休止。設備は残置（撤去未定）",
                            "共通": "受入・タービン発電・排水・電気・計装ほか"}.get(sname, "炉別 系列機器"), F(size=9), X.LEFT)
        rr += 1
    X.put(ws0, rr, 2, "計", F(bold=True), X.CENTER)
    X.put(ws0, rr, 3, tot, F(bold=True), X.RIGHT)
    X.box(ws0, 4, 2, rr, 4)
    rr += 2
    for t in ["・タグ No 1 件を 1 台とする（電動弁・調節弁・主要計装を含む。手動弁は含まない）。",
              "・H15 台帳 No は k1 が灰溶融追設時に Excel 化したときの番号。H16 以降の追加機器には無い。",
              "・更新したときは「更新年月」を書き換え、旧機器の記録は備考に残す。",
              "・灰溶融設備は休止中も台帳に残す（H25.10〜）。"]:
        X.put(ws0, rr, 2, t, F(size=9), X.LEFT)
        ws0.merge_cells(start_row=rr, start_column=2, end_row=rr, end_column=4)
        rr += 1
    X.page(ws0, "portrait", "A4", 1, 1)

    for rr2 in range(5, 10):
        if ws0.cell(rr2, 2).value:
            rec.add("目次", rr2, "item" if ws0.cell(rr2, 2).value != "計" else "total", {"シート": ws0.cell(rr2, 2).value, "台数": ws0.cell(rr2, 3).value})
    hdr = ["No", "H15台帳No", "タグNo", "機器名称", "系統", "型式・仕様", "電動機\n(kW)", "製作所", "設置\n年月", "更新\n年月", "常用\n予備", "法定", "備考"]
    wid = [5, 9, 9, 24, 14, 34, 6, 15, 7, 7, 5, 8, 26]
    for sname, pred in sheets:
        ws = wb.create_sheet(sname)
        X.widths(ws, wid)
        X.put(ws, 1, 1, "【架空】くぬぎ台環境センター　機器台帳　%s" % sname, F(size=13, bold=True))
        X.put(ws, 2, 1, "%s 現在" % asof, F(size=9))
        for i, h in enumerate(hdr, 1):
            X.put(ws, 3, i, h, F(bold=True, size=9), X.CENTER, X.fill("DDEBF7"))
        ws.row_dimensions[3].height = 26
        r = 4
        no = 0
        k1seq = {}
        g = [e for e in eqs if pred(e)]
        cur = None
        for e in g:
            setsu, keito = SYS[e.sys]
            if setsu != cur:
                # 設備の区切り行（H26 再編のときの見出し。中途半端に A:D だけ結合）
                X.put(ws, r, 1, "■" + setsu, F(bold=True, size=9), X.LEFT, X.fill("F2F2F2"))
                ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
                for c in range(1, len(hdr) + 1):
                    ws.cell(r, c).fill = X.fill("F2F2F2")
                cur = setsu
                r += 1
            no += 1
            k = K1_SYS[e.sys]
            k1seq[k] = k1seq.get(k, 0) + 1
            X.put(ws, r, 1, no, F(size=8), X.CENTER)
            X.put(ws, r, 2, _k1_no(e, k1seq[k]), F(size=8), X.CENTER)
            X.put(ws, r, 3, e.tag, F(size=8), X.CENTER)
            X.put(ws, r, 4, K2_NAMES.get(e.name, e.name), F(size=8), X.LEFT)
            X.put(ws, r, 5, keito, F(size=7), X.LEFT)
            X.put(ws, r, 6, e.spec, F(size=7), X.LEFT)
            X.put(ws, r, 7, e.kw if e.kw else None, F(size=8), X.RIGHT)
            X.put(ws, r, 8, MAKERS[e.maker], F(size=7), X.LEFT)
            X.put(ws, r, 9, "H3.3" if e.install == 1990 else _ym(e.install, 3 if e.install != 2003 else 3), F(size=8), X.CENTER)
            lr = _last_rep(e, upto)
            if lr:
                X.put(ws, r, 10, _ym(lr, 12 if lr in KIKAN_FY.values() else _month(e, lr, 0)), F(size=8), X.CENTER)
            X.put(ws, r, 11, {"常用": "常", "予備": "予"}.get(e.duty, ""), F(size=8), X.CENTER)
            X.put(ws, r, 12, "○" if e.legal else "", F(size=8), X.CENTER)
            rem = ""
            if e.origin == "kikan" and KIKAN_FY[e.unit] <= upto:
                rem = "基幹改良で更新（旧機 H3.3）"
            elif e.origin == "dxn":
                rem = "DXN 対策で追設"
            if e.idle:
                rem = "休止中"
                for c in range(1, len(hdr) + 1):
                    ws.cell(r, c).fill = X.fill("E7E6E6")
            if e.legal:
                ws.cell(r, 12).comment = Comment(e.legal, "k2")
            X.put(ws, r, 13, rem, F(size=7), X.LEFT)
            rec.add(sname, r, "item", {"No": no, "H15台帳No": _k1_no(e, k1seq[k]), "code": e.tag, "name": K2_NAMES.get(e.name, e.name), "系統": keito,
                                       "spec": e.spec, "kW": e.kw or None, "maker": MAKERS[e.maker],
                                       "設置年月": ws.cell(r, 9).value, "更新年月": ws.cell(r, 10).value, "常用予備": ws.cell(r, 11).value,
                                       "法定": "○" if e.legal else None, "memo": rem, "unit": sname, "comments": {ws.cell(r, 12).coordinate: e.legal} if e.legal else None})
            r += 1
        X.box(ws, 3, 1, r - 1, len(hdr), inner=X.hair)
        ws.freeze_panes = "D4"
        X.page(ws, "landscape", "A4", 1, 0, titles="3:3", area="A1:M%d" % (r - 1))
        ws.oddFooter.center.text = "%s  &P/&N" % sname
    wb.save(path)
    rec.save()


def make_new(eqs, path, fname=None):
    rec = Rec(fname or os.path.basename(path))
    upto = 2026
    wb = Workbook()
    X.props(wb, "【架空】くぬぎ台環境センター 機器台帳（令和8年10月）", "くぬぎ台市 環境センター k3",
            "機器台帳", "R4 炉別シートを 1 シートに統合（k3）。旧シート名・旧 No を併記")
    ws = wb.active
    ws.title = "機器台帳"
    hdr = ["連番", "号炉", "設備", "系統", "タグNo", "機器名称", "区分", "型式・主要仕様", "kW", "製作所", "設置\n年度", "更新\n年度",
           "重要度", "法定点検", "旧シート", "旧No", "H15台帳No", "備考"]
    wid = [5, 5, 13, 13, 8.5, 22, 5, 30, 5, 13, 6, 6, 5, 20, 8, 5, 8, 26]
    X.widths(ws, wid)
    X.put(ws, 1, 1, "【架空】くぬぎ台環境センター　機器台帳（全炉・共通 統合）", F(size=13, bold=True))
    X.put(ws, 2, 1, "R8.10.1 現在　くぬぎ台市環境部 環境センター 施設係（k3）　※R4.4 に炉別シートを統合。旧シート・旧No は R4.3 版", F(size=9))
    for i, h in enumerate(hdr, 1):
        X.put(ws, 3, i, h, F(bold=True, size=9), X.CENTER, X.fill("D9E1F2"))
    for i in (15, 16, 17):
        ws.cell(3, i).fill = X.fill("EDEDED")
    ws.row_dimensions[3].height = 26
    r = 4
    oldno = {}
    for sname, pred in [("1号炉", lambda e: e.unit == 1), ("2号炉", lambda e: e.unit == 2), ("3号炉", lambda e: e.unit == 3),
                        ("共通", lambda e: e.unit == 0 and not e.idle), ("灰溶融（休止）", lambda e: e.idle)]:
        n = 0
        k1seq = {}
        for e in eqs:
            if pred(e):
                n += 1
                k = K1_SYS[e.sys]
                k1seq[k] = k1seq.get(k, 0) + 1
                oldno[e.tag] = (sname, n, _k1_no(e, k1seq[k]))
    order = sorted(eqs, key=lambda e: (e.unit if e.unit else 9, e.idle, list(SYS).index(e.sys), e.tag))
    for i, e in enumerate(order, 1):
        setsu, keito = SYS[e.sys]
        X.put(ws, r, 1, i, F(size=7), X.CENTER)
        X.put(ws, r, 2, {0: "共通", 1: "1", 2: "2", 3: "3"}[e.unit], F(size=8), X.CENTER)
        X.put(ws, r, 3, setsu, F(size=7), X.LEFT)
        X.put(ws, r, 4, keito, F(size=7), X.LEFT)
        X.put(ws, r, 5, e.tag, F(size=8), X.CENTER)
        X.put(ws, r, 6, e.name, F(size=8), X.LEFT)
        X.put(ws, r, 7, "休止" if e.idle else (e.duty or ""), F(size=7), X.CENTER)
        X.put(ws, r, 8, e.spec, F(size=7), X.LEFT)
        X.put(ws, r, 9, e.kw if e.kw else None, F(size=7), X.RIGHT)
        X.put(ws, r, 10, MAKERS[e.maker], F(size=7), X.LEFT)
        X.put(ws, r, 11, e.install, F(size=8), X.CENTER)
        lr = _last_rep(e, upto)
        X.put(ws, r, 12, lr, F(size=8), X.CENTER)
        X.put(ws, r, 13, e.imp if not e.idle else "－", F(size=8), X.CENTER)
        X.put(ws, r, 14, e.legal, F(size=6), X.LEFT)
        sn, on, k1 = oldno[e.tag]
        X.put(ws, r, 15, sn, F(size=7, color="808080"), X.CENTER)
        X.put(ws, r, 16, on, F(size=7, color="808080"), X.CENTER)
        X.put(ws, r, 17, k1, F(size=7, color="808080"), X.CENTER)
        rem = ""
        if e.name in K2_NAMES and K2_NAMES[e.name] != e.name:
            rem = "R4 名称統一（旧：%s）" % K2_NAMES[e.name]
        if e.origin == "kikan":
            rem = ("基幹改良で更新　" + rem).strip()
        if e.idle:
            rem = "H25.10 休止・残置" + ("　" + rem if rem else "")
            for c in range(1, len(hdr) + 1):
                ws.cell(r, c).fill = X.fill("E7E6E6")
        if lr and lr >= 2022:
            rem = (rem + "　R%d 更新" % (lr - 2018)).strip()
        X.put(ws, r, 18, rem, F(size=7), X.LEFT)
        rec.add("機器台帳", r, "item", {"連番": i, "unit": ws.cell(r, 2).value, "設備": setsu, "系統": keito, "code": e.tag, "name": e.name,
                                       "区分": ws.cell(r, 7).value, "spec": e.spec, "kW": e.kw or None, "maker": MAKERS[e.maker], "設置年度": e.install,
                                       "更新年度": lr, "重要度": ws.cell(r, 13).value, "法定点検": e.legal, "旧シート": sn, "旧No": on, "H15台帳No": k1, "memo": rem})
        r += 1
    X.box(ws, 3, 1, r - 1, len(hdr), inner=X.hair)
    ws.auto_filter.ref = "A3:R%d" % (r - 1)
    ws.freeze_panes = "G4"
    X.page(ws, "landscape", "A3", 1, 0, titles="3:3", area="A1:R%d" % (r - 1))
    ws.oddFooter.center.text = "&P/&N"
    # 集計
    w2 = wb.create_sheet("集計")
    X.widths(w2, [22, 8, 8, 8, 8, 8])
    X.put(w2, 1, 1, "【架空】くぬぎ台環境センター　機器台帳 集計（R8.10.1）", F(size=12, bold=True))
    cols = ["1号炉", "2号炉", "3号炉", "共通", "計"]
    for j, c in enumerate(["設備"] + cols):
        X.put(w2, 3, 1 + j, c, F(bold=True), X.CENTER, X.fill("D9E1F2"))
    rr = 4
    seen = []
    for s in SYS:
        st = SYS[s][0]
        if st in seen:
            continue
        seen.append(st)
        vals = [sum(1 for e in eqs if SYS[e.sys][0] == st and e.unit == u) for u in (1, 2, 3, 0)]
        X.put(w2, rr, 1, st, F(), X.LEFT)
        for j, v in enumerate(vals + [sum(vals)]):
            X.put(w2, rr, 2 + j, v or None, F(), X.RIGHT)
        rr += 1
    vals = [sum(1 for e in eqs if e.unit == u) for u in (1, 2, 3, 0)]
    X.put(w2, rr, 1, "計", F(bold=True), X.CENTER)
    for j, v in enumerate(vals + [sum(vals)]):
        X.put(w2, rr, 2 + j, v, F(bold=True), X.RIGHT)
    X.put(w2, rr + 1, 1, "うち休止（灰溶融）", F(size=9), X.LEFT)
    X.put(w2, rr + 1, 5, sum(e.idle for e in eqs), F(size=9), X.RIGHT)
    X.put(w2, rr + 1, 6, sum(e.idle for e in eqs), F(size=9), X.RIGHT)
    X.put(w2, rr + 2, 1, "うち予備機", F(size=9), X.LEFT)
    for j, u in enumerate((1, 2, 3, 0)):
        X.put(w2, rr + 2, 2 + j, sum(1 for e in eqs if e.unit == u and e.duty == "予備"), F(size=9), X.RIGHT)
    X.put(w2, rr + 2, 6, sum(1 for e in eqs if e.duty == "予備"), F(size=9), X.RIGHT)
    X.box(w2, 3, 1, rr + 2, 6)
    for rr2 in range(4, rr + 3):
        if w2.cell(rr2, 1).value:
            rec.add("集計", rr2, "total" if w2.cell(rr2, 1).value == "計" else "subtotal",
                    {"設備": w2.cell(rr2, 1).value, "1号炉": w2.cell(rr2, 2).value, "2号炉": w2.cell(rr2, 3).value, "3号炉": w2.cell(rr2, 4).value,
                     "共通": w2.cell(rr2, 5).value, "計": w2.cell(rr2, 6).value})
    X.page(w2, "portrait", "A4", 1, 1)
    wb.save(path)
    rec.save()


def make_k1(eqs, path, fname):
    """k1 が H15（灰溶融追設時）に Excel 化した最初の台帳。1 機器を 3 行の仕様ブロックで書く（型式・仕様・電動機が縦に続く）。
    号炉ごとに縦に並べ、番号は「号炉-設備略号-連番」。タグ No はまだ使っていない（メーカー図番の機器番号は備考にだけ）。"""
    rec = Rec(fname)
    upto = 2003
    wb = Workbook()
    ws = wb.active
    ws.title = "機器台帳"
    X.props(wb, "【架空】くぬぎ台環境センター 機器台帳（平成15年度）", "くぬぎ台市 清掃事業所 k1", "機器台帳", "H15 Excel 化（灰溶融追設時）")
    X.widths(ws, [9, 26, 9, 34, 16, 8, 14])
    X.put(ws, 1, 1, "【架空】くぬぎ台環境センター　機器台帳", F(size=14, bold=True), X.NOWRAP)
    X.put(ws, 2, 1, "平成15年度（H16.3.31 現在）　清掃事業所 環境センター　作成 k1　※竣工図書の機器表から転記", F(size=9), X.NOWRAP)
    hdr = ["機器番号", "名　称", "項目", "内　容", "製作所", "設置", "備考（図番）"]
    for i, h in enumerate(hdr, 1):
        X.put(ws, 4, i, h, F(bold=True, size=9), X.CENTER, X.fill("E7E6E6"))
    r = 5
    for unit in (1, 2, 3, 0):
        X.put(ws, r, 1, "＜%s＞" % (("%d号炉" % unit) if unit else "共通"), F(bold=True, size=10), X.NOWRAP)
        rec.add("機器台帳", r, "parent", {"unit": ("%d号炉" % unit) if unit else "共通"})
        r += 1
        seq = {}
        for e in [x for x in eqs if x.unit == unit and x.install <= upto]:
            k = K1_SYS[e.sys]
            seq[k] = seq.get(k, 0) + 1
            no = _k1_no(e, seq[k])
            nm = K2_NAMES.get(e.name, e.name)
            if e.duty == "予備":
                nm += "（予備）"
            lines = [("型式", e.spec.split("、")[0]), ("仕様", "、".join(e.spec.split("、")[1:]) or "－")]
            if e.kw:
                lines.append(("電動機", "%skW" % (int(e.kw) if e.kw == int(e.kw) else e.kw)))
            X.put(ws, r, 1, no, F(size=8), X.CENTER)
            X.put(ws, r, 2, nm, F(size=9), X.LEFT)
            X.put(ws, r, 5, MAKERS[e.maker], F(size=8), X.LEFT)
            X.put(ws, r, 6, "H3.3" if e.install == 1990 else ("H%d" % (e.install - 1988 + 1)), F(size=8), X.CENTER)
            X.put(ws, r, 7, ("灰溶融 H16.3 完成" if e.sys == "AM" else ("DXN 対策 H%d" % (e.install - 1988) if e.origin == "dxn" else "")), F(size=7), X.LEFT)
            for i, (a, b) in enumerate(lines):
                X.put(ws, r + i, 3, a, F(size=8), X.CENTER)
                X.put(ws, r + i, 4, b, F(size=8), X.LEFT)
            n = len(lines)
            ws.merge_cells(start_row=r, start_column=1, end_row=r + n - 1, end_column=1)
            ws.merge_cells(start_row=r, start_column=2, end_row=r + n - 1, end_column=2)
            X.box(ws, r, 1, r + n - 1, 7, inner=X.hair)
            f = {"機器番号": no, "name": nm, "型式": lines[0][1], "仕様": lines[1][1], "電動機": lines[2][1] if n > 2 else None,
                 "maker": MAKERS[e.maker], "設置": ws.cell(r, 6).value, "memo": ws.cell(r, 7).value, "unit": ("%d号炉" % unit) if unit else "共通"}
            rec.add("機器台帳", r, "item", f, r_end=r + n - 1)
            r += n
        r += 1
    ws.freeze_panes = "C5"
    X.page(ws, "portrait", "A4", 1, 0, titles="4:4")
    wb.save(path)
    rec.save()
