# -*- coding: utf-8 -*-
"""本作りの中心：契約ごとの「工事」（工程の予定・実績、日ごとの作業、結果）を作り、作業に日付と工数を入れる。
工程表 → 日報 → 報告書 → 発注 はすべてここから書く（§4.3 の 6〜7）。"""
import datetime as dt
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import rng  # noqa: E402
import jobs as J  # noqa: E402
from works import BYID, CONTRACTS, actual_works  # noqa: E402

D = dt.date
DAY = dt.timedelta(days=1)


# ------------------------------------------------------------------ 休日
def _nth_monday(y, m, n):
    d = D(y, m, 1)
    while d.weekday() != 0:
        d += DAY
    return d + dt.timedelta(weeks=n - 1)


def holidays(y):
    h = {D(y, 1, 1), D(y, 2, 11), D(y, 4, 29), D(y, 5, 3), D(y, 5, 4), D(y, 5, 5), D(y, 11, 3), D(y, 11, 23)}
    if y <= 2018:
        h.add(D(y, 12, 23))
    if y >= 2020:
        h.add(D(y, 2, 23))
    h.add(_nth_monday(y, 1, 2))
    h.add(_nth_monday(y, 9, 3))
    h.add(D(y, 9, 22) if y in (2012, 2016, 2020, 2024) else D(y, 9, 23))
    h.add(_nth_monday(y, 10, 2) if y not in (2020, 2021) else D(y, 7, 24 if y == 2020 else 23))
    h.add(_nth_monday(y, 7, 3) if y not in (2020, 2021) else D(y, 7, 23 if y == 2020 else 22))
    h.add(D(y, 3, 20) if y % 4 in (0, 1) else D(y, 3, 21))
    if y >= 2016:
        h.add(D(y, 8, 11))
    for m in (9,):                      # 国民の休日（祝日にはさまれた平日）
        for dd in range(1, 30):
            d = D(y, m, dd)
            if d - DAY in h and d + DAY in h and d.weekday() != 6:
                h.add(d)
    out = set(h)
    for d in h:
        if d.weekday() == 6:
            n = d + DAY
            while n in h:
                n += DAY
            out.add(n)
    return out


_HOL = {}


def is_off(d):
    if d.year not in _HOL:
        _HOL[d.year] = holidays(d.year)
    return d.weekday() == 6 or d in _HOL[d.year]


def add_work_days(d, n):
    """d から作業日を n 日数えた最後の日（d を含む）。"""
    k = 0
    while True:
        if not is_off(d):
            k += 1
            if k == n:
                return d
        d += DAY


def next_work(d):
    while is_off(d):
        d += DAY
    return d


# ------------------------------------------------------------------ 作業の束ね方（メーカー定期整備）
GROUP = [  # (判定, 作業項目, 班, 基本日数, 順位)
    (lambda e: e["base"] in ("流動床炉本体", "散気装置", "炉壁冷却空気ファン", "炉内監視カメラ装置"), "炉内点検・散気装置点検", "機械", 2, 2),
    (lambda e: e["base"] == "不燃物抜出装置", "不燃物抜出装置 整備", "機械", 3, 3),
    (lambda e: e["setsubi"] == "A" and e["line"] > 0, "給じん設備 点検", "機械", 2, 4),
    (lambda e: e["setsubi"] == "B", "バーナ・炉温調整設備 点検", "機械", 1, 4),
    (lambda e: e["setsubi"] == "C", "砂循環設備 点検", "機械", 3, 3),
    (lambda e: e["base"] in ("廃熱ボイラ", "エコノマイザ", "スートブロワ", "ボイラダスト排出装置", "ボイラダストコンベヤ"), "ボイラ開放・水洗・肉厚測定", "製缶", 4, 1),
    (lambda e: e["setsubi"] == "D" and "減温塔" in e["base"], "減温塔 点検", "機械", 2, 4),
    (lambda e: e["setsubi"] == "D", "ボイラ補機 点検", "機械", 2, 4),
    (lambda e: e["base"] == "バグフィルタ", "バグフィルタ 点検", "機械", 3, 3),
    (lambda e: e["setsubi"] == "E", "排ガス処理補機 点検", "機械", 2, 4),
    (lambda e: e["base"] in ("誘引通風機", "誘引通風機速度制御装置", "誘引通風機潤滑油ユニット", "誘引通風機軸受冷却ファン"), "誘引通風機 点検", "機械", 2, 5),
    (lambda e: e["setsubi"] == "F", "通風設備 点検", "機械", 2, 5),
    (lambda e: e["setsubi"] == "L", "電気設備 点検", "電気", 1, 6),
]
CGROUP = [
    (lambda e: e["base"] in ("ごみクレーン", "ごみクレーン制御盤"), "ごみクレーン 点検（年次・性能検査）", "機械", 4, "受入"),
    (lambda e: e["setsubi"] == "A", "受入設備・前処理破砕機 点検", "機械", 5, "受入"),
    (lambda e: e["setsubi"] in ("E", "K", "C", "B", "J", "N"), "飛灰処理・薬剤・雑設備 点検", "機械", 5, "灰"),
    (lambda e: e["setsubi"] == "L", "受変電設備 停電点検", "電気", 2, "全停"),
    (lambda e: e["base"] in ("脱気器", "高圧蒸気だめ"), "脱気器・高圧蒸気だめ 開放点検（性能検査）", "製缶", 4, "全停"),
    (lambda e: e["setsubi"] in ("G", "H"), "給水・復水設備 点検", "機械", 4, "全停"),
    (lambda e: e["setsubi"] == "M", "計装用空気・中央監視 点検", "電気", 3, "全停"),
]


def _group(e, groups):
    for g in groups:
        if g[0](e):
            return g
    return groups[-1]


def _task(no, name, eqs, han, ps, pe, as_, ae, works=(), added=False):
    return dict(no=no, name=name, eqs=list(eqs), han=han, ps=ps, pe=pe, as_=as_, ae=ae,
                works=list(works), added=added)


def maker_line(fy, line, start, works, r):
    """1 炉の定期整備の工程。start＝立下げの日（金曜）。"""
    T = []
    w0 = next_work(start + 3 * DAY)            # 冷却 2 日のあと
    T.append(_task(0, "立下げ・冷却（運転業者）", [], "運転", start, start + 2 * DAY, start, start + 2 * DAY))
    T.append(_task(0, "安全措置・足場仮設", [], "足場", w0, add_work_days(w0, 2), w0, add_work_days(w0, 2)))
    T.append(_task(0, "炉開放・流動砂抜出", [BYID_FIND(line, "流動床炉本体")], "機械", w0, add_work_days(w0, 3),
                   w0, add_work_days(w0, 3)))
    groups = {}
    for w in works:
        e = BYID[w["eq"]]
        g = _group(e, GROUP)
        if "補修（点検時発見）" in w["label"] or ("追加" in w["label"]):
            key = ("追加", e["base"])
            groups.setdefault(key, dict(name=f"{e['name'].split(' No.')[0]} 補修（追加）", han=g[2], days=2,
                                        rank=6, works=[], added=True))
        else:
            key = (g[1],)
            groups.setdefault(key, dict(name=g[1], han=g[2], days=g[3], rank=g[4], works=[], added=False))
        groups[key]["works"].append(w)
    rows = []
    for key, g in groups.items():
        oh = sum(1 for w in g["works"] if w["kind"] == "OVERHAUL")
        rp = sum(1 for w in g["works"] if w["kind"] == "REPLACE")
        days = min(g["days"] + oh + 3 * rp + len(g["works"]) // 4, 9)
        name = g["name"]
        if rp and not g["added"]:
            name = name.replace("点検", "ろ布全数交換") if "バグ" in name else name + "・更新"
        elif oh and not g["added"] and "整備" not in name:
            name = name + "・分解整備"
        rows.append((g["rank"], name, g["han"], days, g["works"], g["added"]))
    rows.sort(key=lambda x: (x[0], x[1]))
    lane_end = {}
    for (rank, name, han, days, ws, added) in rows:
        base = {1: add_work_days(w0, 2), 2: add_work_days(w0, 3), 3: add_work_days(w0, 4)}.get(rank, add_work_days(w0, 6))
        s0 = max(base, lane_end.get(rank, base))
        s0 = next_work(s0)
        pe = add_work_days(s0, days)
        slip = 1 if r.random() < 0.18 else 0
        if added:
            ps = None
            a0 = next_work(s0 + DAY)
            ae = add_work_days(a0, days)
            T.append(_task(0, name, [w["eq"] for w in ws], han, None, None, a0, ae, ws, True))
        else:
            ae = add_work_days(s0, days + slip)
            T.append(_task(0, name, [w["eq"] for w in ws], han, s0, pe, s0, ae, ws))
        lane_end[rank] = add_work_days(pe, 1) + DAY
    last_p = max(t["pe"] for t in T if t["pe"])
    last_a = max(t["ae"] for t in T)
    boiler = [w for w in works if BYID[w["eq"]]["base"] == "廃熱ボイラ"]
    kp = next_work(last_p + DAY)
    ka = next_work(max(last_a, last_p) + DAY)
    T.append(_task(0, "ボイラ性能検査 準備・受検", [BYID_FIND(line, "廃熱ボイラ")], "製缶", kp, add_work_days(kp, 2),
                   ka, add_work_days(ka, 2), boiler[:0]))
    rp_ = next_work(add_work_days(kp, 2) + DAY)
    ra_ = next_work(add_work_days(ka, 2) + DAY)
    T.append(_task(0, "復旧・流動砂張込", [], "機械", rp_, add_work_days(rp_, 2), ra_, add_work_days(ra_, 2)))
    tp = next_work(add_work_days(rp_, 2) + DAY)
    ta = next_work(add_work_days(ra_, 2) + DAY)
    T.append(_task(0, "通風試験・試運転", [], "機械", tp, tp, ta, ta))
    T.append(_task(0, "立上げ（運転業者）", [], "運転", tp + DAY, tp + 2 * DAY, ta + DAY, ta + 2 * DAY))
    for k, t in enumerate(T, 1):
        t["no"] = k
    return T


def BYID_FIND(line, base):
    for e in BYID.values():
        if e["line"] == line and e["base"] == base:
            return e["id"]
    raise KeyError(base)


def maker_common(fy, receive_start, zen_start, works, r):
    T = [_task(0, f"全炉停止（運転業者）", [], "運転", zen_start - DAY, add_work_days(zen_start, 5),
               zen_start - DAY, add_work_days(zen_start, 5))]
    groups = {}
    for w in works:
        e = BYID[w["eq"]]
        g = _group(e, CGROUP)
        if "補修（点検時発見）" in w["label"]:
            key = ("追加", e["base"])
            groups.setdefault(key, dict(name=f"{e['name'].split(' No.')[0]} 補修（追加）", han=g[2], days=2, when=g[4],
                                        works=[], added=True))
        else:
            groups.setdefault((g[1],), dict(name=g[1], han=g[2], days=g[3], when=g[4], works=[], added=False))
        groups[key if "補修（点検時発見）" in w["label"] else (g[1],)]["works"].append(w)
    cursor = {"受入": receive_start, "灰": add_work_days(receive_start, 6), "全停": zen_start}
    for key, g in sorted(groups.items(), key=lambda kv: (["受入", "灰", "全停"].index(kv[1]["when"]), kv[1]["added"])):
        s0 = next_work(cursor[g["when"]]) if g["when"] != "全停" else zen_start
        days = min(g["days"], 5) if g["when"] == "全停" else g["days"]
        pe = add_work_days(s0, days)
        if g["added"]:
            a0 = next_work(pe + DAY) if g["when"] != "全停" else s0
            T.append(_task(0, g["name"], [w["eq"] for w in g["works"]], g["han"], None, None, a0,
                           add_work_days(a0, 2), g["works"], True))
        else:
            T.append(_task(0, g["name"], [w["eq"] for w in g["works"]], g["han"], s0, pe, s0, pe, g["works"]))
            if g["when"] != "全停":
                cursor[g["when"]] = pe + DAY
    T.append(_task(0, "全炉立上げ（運転業者）", [], "運転", add_work_days(zen_start, 5) + DAY,
                   add_work_days(zen_start, 5) + DAY, add_work_days(zen_start, 5) + DAY, add_work_days(zen_start, 5) + DAY))
    T.sort(key=lambda t: (t["as_"], t["no"]))
    for k, t in enumerate(T, 1):
        t["no"] = k
    return T


def _friday_on_or_after(d):
    while d.weekday() != 4:
        d += DAY
    return d


def maker_job(fy):
    """定期整備業務委託 T{fy}。H22 は見本の工程（jobs.py）をそのまま使う。"""
    r = rng(f"job{fy}")
    W = [w for w in actual_works() if w["contract"] == f"T{fy}"]
    by = {0: [], 1: [], 2: []}
    for w in W:
        by[BYID[w["eq"]]["line"]].append(w)
    if fy == 2010:
        def conv(tasks):
            out = []
            for (no, name, eqs, han, ps, pe, as_, ae) in tasks:
                ws = [w for w in W if w["eq"] in eqs and not (w["eq"] == "1B-001" and (("追加" in name) != (w["kind"] == "OVERHAUL")))]
                if name.startswith("炉開放") or name.startswith("ボイラ性能検査"):
                    ws = []
                out.append(_task(no, name, eqs, han, ps, pe, as_, ae, ws, added=ps is None))
            return out
        areas = [dict(label="1号炉", line=1, tasks=conv(J.J22_L1), span=J.J22_L1_SPAN),
                 dict(label="2号炉", line=2, tasks=conv(J.J22_L2), span=J.J22_L2_SPAN),
                 dict(label="共通", line=0, tasks=conv(J.J22_CM), span=J.J22_CM_SPAN)]
    else:
        s1 = _friday_on_or_after(D(fy, 9, 1) + dt.timedelta(weeks=r.randint(0, 2)))
        s2 = _friday_on_or_after(D(fy, 10, 29) + dt.timedelta(weeks=r.randint(0, 1)))
        t1 = maker_line(fy, 1, s1, by[1], r)
        t2 = maker_line(fy, 2, s2, by[2], r)
        rec = next_work(_friday_on_or_after(D(fy, 10, 12)) + 3 * DAY)
        zen = next_work(s2 + dt.timedelta(days=10))
        while zen.weekday() != 0:
            zen += DAY
        tc = maker_common(fy, rec, zen, by[0], r)
        def span(ts):
            return (min(t["as_"] for t in ts), max(t["ae"] for t in ts))
        areas = [dict(label="1号炉", line=1, tasks=t1, span=span(t1)),
                 dict(label="2号炉", line=2, tasks=t2, span=span(t2)),
                 dict(label="共通", line=0, tasks=tc, span=span(tc))]
    return dict(cid=f"T{fy}", fy=fy, contract=CONTRACTS[f"T{fy}"], areas=areas, works=W, vendor="m1")


if __name__ == "__main__":
    for fy in (2006, 2010, 2017):
        j = maker_job(fy)
        for a in j["areas"]:
            print(fy, a["label"], a["span"], len(a["tasks"]), sum(len(t["works"]) for t in a["tasks"]))
        print(len(j["works"]), len({w["wid"] for a in j["areas"] for t in a["tasks"] for w in t["works"]}))
