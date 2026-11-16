# -*- coding: utf-8 -*-
"""発注台帳＝修繕執行管理表（H24〜R8、年度ごとに 1 ファイル）
 t1（H24〜H29）：予算表＋契約一覧、日付は 月/日
 t2（H30〜R4）：t1 の表に 区分・対象機器・号炉・計画No・変更後・支出負担行為No を差し込み（黄色の見出し）、日付は和暦
 t3（R5〜）：t2 の表をそのまま使う。R6 から対象機器の頭に管理番号を書く
 金額は町の契約額（税込）。請負額で、原価ではない。"""
import datetime as dt, random, zlib
import world as W, cw, sg
from xl import *

def plan_no():
    no = 0; out = {}
    for sysc in W.SYS_ORDER:
        for e in W.EQ:
            if e.sys != sysc: continue
            if e.added_2015: out[e.mid] = "%d-2" % no
            else: no += 1; out[e.mid] = str(no)
    return out
PNO = plan_no()

ORDER = {"MK": 0, "ITAKU": 1, "YAKUMU": 2, "BUY": 3, "EN": 4, "CRUP": 4, "LOCAL": 5}
ITAKU_ORDER = ["ELEC", "ANAL", "ENV", "FIRE", "EV", "JOKA"]

def md(d): return "" if not d else "%d/%d" % (d.month, d.day)
def wa(d):
    if not d: return ""
    if d >= dt.date(2019, 5, 1): return "R%d.%d.%d" % (d.year - 2018, d.month, d.day)
    return "H%d.%d.%d" % (d.year - 1988, d.month, d.day)

def rows_of(fy):
    cs = [c for c in cw.CONTRACTS.values() if c.fy == fy or (c.kind == "EN" and fy in (2015, 2016))]
    def key(c):
        k = ORDER[c.kind]
        sub = ITAKU_ORDER.index("".join(ch for ch in c.cid if not ch.isdigit())) if c.kind == "ITAKU" else 0
        return (k, sub, c.cdate or dt.date(fy, 4, 1))
    cs.sort(key=key)
    if fy == 2026:
        cs = [c for c in cs if c.cdate and c.cdate <= dt.date(2026, 9, 30)]
    return cs

def dates(c, fy):
    """契約日・着手・完了・完了届・検査・支払・備考"""
    if c.kind == "MK":
        s = sg.SCHED[c.cid]
        sub = dt.date(2022, 12, 15) if fy == 2022 else cw.wd_after(s.site[1], 30 if fy >= 2016 else 32)
        kensa = cw.wd_after(sub, 3)
        pay = kensa + dt.timedelta(days=28)
        note = "現地 %d/%d〜%d/%d" % (s.site[0].month, s.site[0].day, s.site[1].month, s.site[1].day)
        if fy == 2022: note = "現地10/11〜10/28 1号ろ布144本含む"
        if fy == 2026: return (c.cdate, c.start, c.end, None, None, None, "現地 10/13〜（予定）")
        return (c.cdate, c.start, c.end, sub, kensa, pay, note)
    if c.kind == "ITAKU":
        base = "".join(ch for ch in c.cid if not ch.isdigit())
        if base in ("ELEC", "EV", "JOKA"):
            return (c.cdate, c.start, c.end, None, None, "毎月" if base == "ELEC" else "四半期", c.note)
        r = random.Random(zlib.crc32(c.cid.encode()) % 1013)
        k = dt.date(fy + 1, 3, r.randint(5, 20)) if base != "ENV" else dt.date(fy + 1, 2, r.randint(14, 24))
        if fy == 2017:
            k = {"ANAL": dt.date(2018, 3, 9), "ENV": dt.date(2018, 2, 20), "FIRE": dt.date(2018, 3, 20)}[base]
        if fy == 2022:
            k = {"ANAL": dt.date(2023, 3, 8), "ENV": dt.date(2023, 2, 17), "FIRE": dt.date(2023, 3, 17)}[base]
        kensa = cw.wd_after(k, 2 if base != "FIRE" else 2)
        if fy == 2017 and base == "ANAL": kensa = dt.date(2018, 3, 12)
        if fy == 2022: kensa = {"ANAL": dt.date(2023, 3, 10), "ENV": dt.date(2023, 2, 21), "FIRE": dt.date(2023, 3, 22)}[base]
        pay = dt.date(fy + 1, 3, 27 if fy == 2017 else (28 if base != "FIRE" else 30)) if fy != 2017 or base != "FIRE" else dt.date(2018, 3, 30)
        if fy == 2026: return (c.cdate, c.start, c.end, None, None, None, c.note)
        return (c.cdate, c.start, c.end, k, kensa, pay, c.note)
    if c.kind == "YAKUMU":
        return (c.cdate, c.start, c.end, None, None, c.cdate + dt.timedelta(days=21), c.note)
    if c.kind == "BUY":
        return (c.cdate, None, c.end, c.kanryo, c.kensa, c.pay, c.note)
    if c.kind == "EN":
        if fy == 2015:
            return (c.cdate, c.start, c.end, None, None, dt.date(2016, 3, 25), "継続費 H27年割（出来高払）1号炉")
        return (c.cdate, c.start, c.end, dt.date(2017, 2, 24), dt.date(2017, 3, 3), dt.date(2017, 3, 28), "継続費 H28年割 2号炉・中央監視")
    if c.kind == "CRUP":
        s = sg.SCHED[c.cid]
        k = cw.wd_after(s.site[1], 20)
        return (c.cdate, c.start, c.end, k, cw.wd_after(k, 4), cw.wd_after(k, 4) + dt.timedelta(days=27), "全炉停止中 %d/%d〜%d/%d 港南クレーン整備 下請" % (s.site[0].month, s.site[0].day, s.site[1].month, s.site[1].day))
    # LOCAL
    return (c.cdate, c.site[0] if c.key not in ("H30-ASH", "R4-CH") else c.start, c.site[1], c.kanryo, c.kensa, c.pay, c.note or "")

def amount(c, fy):
    if c.kind == "EN": return c.split[fy]
    return c.amount

BUD = {"需用費（修繕料）": "修繕料", "需用費（消耗品費）": "消耗品費", "委託料": "委託料", "役務費（手数料）": "役務費", "工事請負費": "工事請負費"}

def budget(fy, cs):
    used = {}
    for c in cs: used[c.hi] = used.get(c.hi, 0) + ((c.chg or amount(c, fy)) if fy >= 2018 else amount(c, fy))
    fixed = {2017: [("需用費（修繕料）", 4800000, 0), ("需用費（消耗品費）", 1500000, 0), ("委託料", 22000000, 0), ("役務費（手数料）", 60000, 0)],
             2022: [("需用費（修繕料）", 6000000, 0), ("需用費（消耗品費）", 800000, 0), ("委託料", 24500000, 0), ("役務費（手数料）", 40000, 0)]}
    if fy in fixed: return fixed[fy], used
    out = []
    for nm, hi in BUD.items():
        u = used.get(hi, 0)
        if hi == "工事請負費" and not u: continue
        b = int((u * 1.12 + 300000) / 100000) * 100000 if hi not in ("役務費",) else 40000
        if hi == "消耗品費": b = max(b, 600000)
        hosei = 0
        if hi == "修繕料" and u > b: hosei = int((u - b) / 100000 + 1) * 100000
        out.append((nm, b, hosei))
    return out, used

def budget_block(ws, r, budget_rows, used, spans):
    hdr = ["費目", "当初予算", "補正", "予算計", "執行済", "残額"]
    for (a, b), h in zip(spans, hdr):
        merge(ws, r, a, r, b, h, F(9, True), AC, B_ALL, fill("E7E6E6"))
    for j, (nm, bud, hosei) in enumerate(budget_rows):
        u = used.get(BUD[nm], 0)
        vals = [nm, bud, hosei, bud + hosei, u, bud + hosei - u]
        for i, ((a, b), v) in enumerate(zip(spans, vals)):
            merge(ws, r + 1 + j, a, r + 1 + j, b, v, F(9), AL if i == 0 else AR, B_ALL)
            ws.cell(row=r + 1 + j, column=a).number_format = "#,##0"
        rec(ws.title, r + 1 + j, "item", dict(zip(hdr, vals)))
    return r + 1 + len(budget_rows)

def t1_sheet(fy):
    cs = rows_of(fy)
    fname = "【架空】しおさい町ごみ処理場_発注台帳_%s修繕執行管理表_t1様式.xlsx" % W.wareki(fy)
    rec_open(fname)
    wb = new_wb("【架空】しおさい町ごみ処理場 修繕執行管理表", "しおさい町 環境衛生課 t1", dt.datetime(2012, 4, 10, 8, 30), "t1")
    ws = wb.active; ws.title = W.wareki(fy)
    widths(ws, [4, 34, 26, 10, 6, 12, 7, 7, 7, 7, 7, 7, 24])
    put(ws, 1, 1, "【架空】%s　修繕執行管理表（ごみ処理場）" % W.wareki_long(fy), F(14, True))
    put(ws, 2, 1, "環境衛生課　ごみ処理場担当", F(9))
    bud, used = budget(fy, cs)
    r = budget_block(ws, 4, bud, used, [(2, 2), (3, 3), (4, 5), (6, 6), (7, 9), (10, 12)])
    r += 2
    hdr = ["No", "件名", "業者", "費目", "見積", "契約額（円）", "契約日", "着手", "完了", "完了届", "検査", "支払", "備考"]
    for i, h in enumerate(hdr, 1):
        put(ws, r, i, h, F(9, True), AC, B_ALL, fill("E7E6E6"))
    ws.row_dimensions[r].height = 22
    for k, c in enumerate(cs, 1):
        r += 1
        (kd, st, en, kt, ken, pay, nt) = dates(c, fy)
        vals = [k, c.short, c.vendor, c.hi, c.mitsu, amount(c, fy), md(kd), md(st), md(en), md(kt), md(ken), pay if isinstance(pay, str) else md(pay), nt]
        for i, v in enumerate(vals, 1):
            put(ws, r, i, v, F(9), AR if i == 6 else (AL if i in (2, 3, 13) else AC), B_ALL, fmt="#,##0" if i == 6 else None)
        ws.row_dimensions[r].height = 20
        rec(ws.title, r, "item", dict(zip(hdr, vals)), c.works)
        if c.cid == "BUY2017":
            comment(ws, r, 2, "予備品。H29定修で2号に8枚使用、残16枚（倉庫2F）", "t1")
            REC[CUR["stem"]]["records"][-1]["fields"]["comments"] = {"B%d" % r: "予備品。H29定修で2号に8枚使用、残16枚（倉庫2F）"}
    for _ in range(3):
        r += 1
        for i in range(1, 14): put(ws, r, i, None, None, AC, B_ALL)
    r += 1
    tot = sum(amount(c, fy) for c in cs)
    put(ws, r, 2, "合計", F(9, True), AR, B_ALL)
    put(ws, r, 6, tot, F(9, True), AR, B_ALL, fmt="#,##0")
    rec(ws.title, r, "total", {"件名": "合計", "契約額（円）": tot})
    r += 2
    put(ws, r, 1, "※日付は年度内の月/日。1〜3月は翌年。　作成 t1", F(8))
    page(ws, landscape=True, paper=9, area="A1:M%d" % r)
    return save(wb, fname)

def t2_sheet(fy):
    cs = rows_of(fy)
    who = "t2" if fy <= 2022 else "t3"
    fname = "【架空】しおさい町ごみ処理場_発注台帳_%s修繕執行管理表_t2列追加様式.xlsx" % W.wareki(fy)
    rec_open(fname)
    wb = new_wb("【架空】しおさい町ごみ処理場 修繕執行管理表", "しおさい町 環境衛生課 t1", dt.datetime(2012, 4, 10, 8, 30), who)
    ws = wb.active; ws.title = W.wareki(fy)
    cols = [("No", 4, 0), ("区分", 5, 1), ("件名", 32, 0), ("対象機器", 20 if fy >= 2024 else 18, 1), ("号炉", 5, 1), ("計画No", 6, 1), ("業者", 24, 0),
            ("費目", 9, 0), ("見積", 5, 0), ("契約額（円）", 11, 0), ("変更後（円）", 11, 1), ("契約日", 8, 0), ("支出負担\n行為No", 8, 1),
            ("着手", 8, 0), ("完了", 8, 0), ("完了届", 8, 0), ("検査", 8, 0), ("支払", 8, 0), ("備考", 26, 0)]
    widths(ws, [c[1] for c in cols])
    put(ws, 1, 1, "【架空】%s　修繕執行管理表（ごみ処理場）" % W.wareki_long(fy).replace("令和元", "令和元"), F(14, True))
    put(ws, 2, 1, "環境衛生課　ごみ処理場担当", F(9))
    put(ws, 2, 7, "黄色の列：H30 追加（t2）　計画No＝長寿命化計画 整備スケジュール表の No", F(8, color="806000"))
    bud, used = budget(fy, cs)
    r = budget_block(ws, 4, bud, used, [(3, 3), (4, 4), (5, 6), (7, 7), (8, 10), (11, 12)])
    r += 2
    for i, (h, w_, y) in enumerate(cols, 1):
        put(ws, r, i, h, F(9, True), AC, B_ALL, fill("FFE699") if y else fill("E7E6E6"))
    ws.row_dimensions[r].height = 26
    sfno = 0
    for k, c in enumerate(cs, 1):
        r += 1
        (kd, st, en, kt, ken, pay, nt) = dates(c, fy)
        kb = {"MK": "委託", "ITAKU": "委託", "YAKUMU": "役務", "BUY": "物品", "LOCAL": "修繕", "EN": "工事", "CRUP": "工事"}[c.kind]
        es = sorted({W.BYMID[w.mid].mid: W.BYMID[w.mid] for w in c.works if w.src != "legal"}.values(), key=lambda e: e.idx)
        if c.kind == "MK": eqn, gou, pno = "1・2号炉 全般" if fy not in (2015, 2016) else ("2号炉・共通" if fy == 2015 else "1号炉・共通"), "", "－（毎年）"
        elif c.kind == "ITAKU":
            base = "".join(ch for ch in c.cid if not ch.isdigit())
            eqn = {"ELEC": "受変電設備", "ANAL": "1・2号炉 排ガス", "ENV": "－", "FIRE": "消防用設備", "EV": "人荷用エレベーター", "JOKA": "合併処理浄化槽"}[base]
            gou = "共通" if base in ("EV", "JOKA") else ""
            pno = PNO[W.m(eqn)] if base in ("EV", "JOKA") else ""
        elif c.kind == "YAKUMU": eqn, gou, pno = "計量機", "共通", PNO[W.m("計量機")]
        elif c.kind == "BUY":
            if "ろ布" in c.short: eqn, gou, pno = "バグフィルタ（1・2号）", "1・2", "%s・%s" % (PNO[W.m("バグフィルタ", 1)], PNO[W.m("バグフィルタ", 2)])
            else: eqn, gou, pno = "燃焼火格子（1・2号）", "1・2", "%s・%s" % (PNO[W.m("燃焼火格子", 1)], PNO[W.m("燃焼火格子", 2)])
        else:
            if len(es) == 1:
                e = es[0]; eqn = e.name; gou = "共通" if e.furnace == 0 else "%d号" % e.furnace; pno = PNO[e.mid]
            else:
                eqn = "、".join(sorted({e.name for e in es})); fs = sorted({e.furnace for e in es})
                gou = "共通" if fs == [0] else "・".join(str(f) for f in fs if f) + "号"; pno = "・".join(PNO[e.mid] for e in es)
            if c.kind == "EN": eqn, gou, pno = "火格子・ガス冷却室・誘引通風機ほか", "1号" if fy == 2015 else "2号・共通", "（長寿命化）"
            if c.kind == "CRUP": eqn, gou, pno = "ごみクレーン操作盤・動力盤", "共通", "%s・%s" % (PNO[W.m("ごみクレーン操作盤")], PNO[W.m("ごみクレーン動力盤")])
        if fy >= 2024 and es and c.kind in ("LOCAL", "CRUP"):
            eqn = "、".join(e.mid for e in es) + " " + eqn  # t3：管理番号を頭に
        if fy in (2020, 2021) and c.kind == "LOCAL" and k % 2 == 0:
            pno = ""  # 計画表が止まっていた年：計画Noを入れていない行がある
        sfno += 1
        sf = "%02d-%04d" % (fy % 100, [412, 15, 133, 134, 16, 17, 18][min(k - 1, 6)] if fy == 2022 and k <= 7 else 100 + sfno * 37)
        if fy == 2022:
            sf = {"MK2022": "22-0412", "ELEC2022": "22-0015", "ANAL2022": "22-0133", "ENV2022": "22-0134", "FIRE2022": "22-0016", "EV2022": "22-0017", "JOKA2022": "22-0018",
                  "SCALE2022": "22-0655", "BUY2022": "22-0297", "L-R4-PA": "22-0251", "L-R4-WA": "22-0388", "L-R4-CH": "22-0420", "L-R4-DR": "22-0871", "L-R4-JO": "22-0950"}.get(c.cid, sf)
        if c.kind == "LOCAL" and c.chg:
            nt = nt or ""
        if c.key == "R4-CH" if c.kind == "LOCAL" else False:
            nt = "変更10/26 テール軸受追加 +165,000"
        if c.cid == "L-R4-PA": nt = "R1にも同じ修理"
        vals = [k, kb, c.short if c.kind != "EN" else "長寿命化改良工事（%s年割）" % W.wareki(fy), eqn, gou, pno, c.vendor, c.hi, c.mitsu, amount(c, fy), c.chg,
                wa(kd), sf, wa(st), wa(en), wa(kt), wa(ken), pay if isinstance(pay, str) else wa(pay), nt]
        for i, v in enumerate(vals, 1):
            al = AR if i in (10, 11) else (AL if i in (3, 4, 7, 19) else AC)
            put(ws, r, i, v, F(9 if i not in (12, 14, 15, 16, 17, 18) else 8, color="C00000" if (i == 11 and v) else None), al, B_ALL, fmt="#,##0" if i in (10, 11) else None)
        ws.row_dimensions[r].height = 22
        rec(ws.title, r, "item", {h: v for (h, _, _), v in zip(cols, vals)}, c.works)
        if c.key == "R4-CH" if c.kind == "LOCAL" else False:
            comment(ws, r, 3, "10/20 テール軸受損傷→追加。10/26 変更契約", "t2")
            REC[CUR["stem"]]["records"][-1]["fields"]["comments"] = {"C%d" % r: "10/20 テール軸受損傷→追加。10/26 変更契約"}
    for _ in range(2):
        r += 1
        for i in range(1, len(cols) + 1): put(ws, r, i, None, None, AC, B_ALL)
    r += 1
    t1_ = sum(amount(c, fy) for c in cs); t2_ = sum((c.chg or amount(c, fy)) for c in cs)
    put(ws, r, 3, "合計", F(9, True), AR, B_ALL)
    put(ws, r, 10, t1_, F(9, True), AR, B_ALL, fmt="#,##0")
    put(ws, r, 11, t2_, F(9, True, "C00000"), AR, B_ALL, fmt="#,##0")
    rec(ws.title, r, "total", {"件名": "合計", "契約額（円）": t1_, "変更後（円）": t2_})
    r += 2
    put(ws, r, 1, "※日付は年度内の月/日。1〜3月は翌年。　作成 t1", F(8, color="808080")); r += 1
    note = "※合計の右（赤）は変更後の額で集計。執行済は変更後の額。" + ("R2・R3 は整備計画表への記入なし（この表を正とする）　t2" if fy >= 2020 and fy <= 2022 else "　" + who)
    if fy == 2026: note = "※R8 は 9 月末までの契約を記入。定修（10月）の完了・支払は未記入。　t3"
    put(ws, r, 1, note, F(8, color="C00000"))
    ws.freeze_panes = "D%d" % (6 + len(bud) + 3)
    page(ws, landscape=True, paper=8, area="A1:%s%d" % (CL(len(cols)), r))
    return save(wb, fname)

def build_all():
    out = []
    for fy in range(2012, 2027):
        out.append(t1_sheet(fy) if fy <= 2017 else t2_sheet(fy))
    return out
