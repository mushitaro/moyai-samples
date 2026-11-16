# -*- coding: utf-8 -*-
"""全年度（2001〜2026）の契約・工程・作業（正解の世界の「いつ・どの契約で・誰が」）。
world.py の履歴（年度単位の作業）を、契約ごとの工程の作業（task）に割り付け、日付と工数を持たせる。
工程を先に決め、日報はこの工程から作る（gen_daily_full.py）。
"""
import datetime as dt
import random
from collections import OrderedDict
from world import CLS, SYS, SEED, KIKAN_FY, DXN_FY, AM_FY, AM_STOP_FY, wareki, find
import sched as S

D = dt.date
FY0, FY1 = 2001, 2026
CUTOFF = D(2026, 9, 30)    # 2026 年度は上期の実績まで


# ---------------------------------------------------------------- 暦
def _nth_monday(y, m, n):
    d = D(y, m, 1)
    while d.weekday() != 0:
        d += dt.timedelta(days=1)
    return d + dt.timedelta(days=7 * (n - 1))


def holidays(y):
    h = {D(y, 1, 1), D(y, 2, 11), D(y, 4, 29), D(y, 5, 3), D(y, 5, 4), D(y, 5, 5), D(y, 11, 3), D(y, 11, 23)}
    if y <= 2018:
        h.add(D(y, 12, 23))
    if y >= 2020:
        h.add(D(y, 2, 23))
    if y >= 2016:
        h.add(D(y, 8, 11))
    h.add(_nth_monday(y, 1, 2))
    h.add(_nth_monday(y, 7, 3) if y >= 2003 else D(y, 7, 20))
    h.add(_nth_monday(y, 9, 3) if y >= 2003 else D(y, 9, 15))
    h.add(_nth_monday(y, 10, 2))
    h.add(D(y, 3, 20 if y % 4 in (0, 1) else 21))
    h.add(D(y, 9, 22 if y % 4 == 0 else 23))
    out = set(h)
    for d in h:
        if d.weekday() == 6:
            n = d + dt.timedelta(days=1)
            while n in h:
                n += dt.timedelta(days=1)
            out.add(n)
    return out


_HOL = {}


def is_hol(d):
    if d.year not in _HOL:
        _HOL[d.year] = holidays(d.year)
    return d in _HOL[d.year] or d in S.HOLI


def workday(d):
    return d.weekday() != 6 and not is_hol(d)


def add_workdays(d, n):
    """d から数えて n 日目の作業日（d が作業日なら d を 1 日目）"""
    c = 0
    while True:
        if workday(d):
            c += 1
            if c >= n:
                return d
        d += dt.timedelta(days=1)


def next_workday(d):
    d += dt.timedelta(days=1)
    while not workday(d):
        d += dt.timedelta(days=1)
    return d


def wdays(a, b):
    out = []
    d = a
    while d <= b:
        if workday(d):
            out.append(d)
        d += dt.timedelta(days=1)
    return out


def tax(fy, d=None):
    if fy <= 2013:
        return 0.05
    if fy <= 2018:
        return 0.08
    if fy == 2019 and d is not None and d < D(2019, 10, 1):
        return 0.08
    return 0.10


# ---------------------------------------------------------------- 人
def city_person(fy):
    return "k1" if fy <= 2010 else ("k2" if fy <= 2017 else "k3")


def agent(fy):
    return "m1" if fy <= 2012 else "m2"


def reporter(fy):
    return "r1" if fy <= 2007 else "r2"


def direct_lead(fy):
    return "d1" if fy <= 2016 else "d2"


# ---------------------------------------------------------------- 物
class Task:
    def __init__(self, tid, name, crew, plan, act, tags, kind="std", size=4, works=None, note="", cls=None):
        self.id, self.name, self.crew, self.plan, self.act, self.tags = tid, name, crew, plan, act, tags
        self.kind, self.size, self.works, self.note, self.cls = kind, size, works or [], note, cls

    @property
    def a(self):
        return (self.act or self.plan)[0]

    @property
    def b(self):
        return (self.act or self.plan)[1]


class Contract:
    def __init__(self, key, fy, unit, kind, name, vendor, start, end, **kw):
        self.key, self.fy, self.unit, self.kind, self.name, self.vendor = key, fy, unit, kind, name, vendor
        self.start, self.end = start, end
        self.tasks = []
        self.amount = kw.get("amount", 0)
        self.amount2 = kw.get("amount2", None)
        self.method = kw.get("method", "随意契約")
        self.budget = kw.get("budget", "委託料")
        self.cdate = kw.get("cdate", start - dt.timedelta(days=25))
        self.note = kw.get("note", "")
        self.docs = kw.get("docs", kind in ("teishu", "kyotsu", "hoshu"))
        self.outage = kw.get("outage")
        self.zentei = kw.get("zentei")
        self.daily = {}     # date -> dict(weather, note, rows=[(tid, crew, place, text, n, h, prog)])
        self.sample = kw.get("sample", None)
        self.cancel_after_cutoff = False


WORKS = OrderedDict()


def _wid():
    return "W%05d" % (len(WORKS) + 1)


def add_work(tag, kind, what, contract, a=None, b=None, vendor="", src="history", fy=None):
    wid = _wid()
    WORKS[wid] = dict(id=wid, tag=tag, kind=kind, what=what, contract=contract, start=a, end=b,
                      manDays=0.0, vendor=vendor, src=src, fy=fy, links=[])
    return wid


# ---------------------------------------------------------------- 期間
def outages(fy):
    """{unit: 停止日}。kairyo の炉は除く"""
    if fy == 2013:
        return {1: D(2013, 5, 13), 3: D(2013, 11, 5)}, D(2013, 9, 2)
    if fy == 2024:
        return {1: D(2024, 9, 2), 2: D(2024, 10, 7), 3: D(2024, 11, 11)}, D(2024, 10, 9)
    if fy == 2026:
        return {1: D(2026, 9, 1), 2: D(2026, 10, 5), 3: D(2026, 11, 9)}, D(2026, 10, 7)
    rot = (fy - 2024) % 3
    order = [1, 2, 3][rot:] + [1, 2, 3][:rot]
    if fy <= 2012:
        st = [_nth_monday(fy, 5, 2), _nth_monday(fy, 6, 3), _nth_monday(fy, 11, 1)]
        ky = _nth_monday(fy, 9, 1)
    else:
        st = [_nth_monday(fy, 9, 1), _nth_monday(fy, 10, 1), _nth_monday(fy, 11, 2)]
        ky = _nth_monday(fy, 10, 1) + dt.timedelta(days=2)
    res = {}
    for u, d in zip(order, st):
        while is_hol(d):
            d += dt.timedelta(days=1)
        res[u] = d
    for u in list(res):
        if (u in DXN_FY and DXN_FY[u] == fy) or (u in KIKAN_FY and KIKAN_FY[u] == fy and u):
            del res[u]
    return res, ky


CREW_OF = {"IN": "A", "AS": "A", "BL": "B", "GT": "C", "FD": "D", "UT": "D", "EL": "E", "IC": "E"}
DUR = {"stoker": 9, "furnace": 9, "hyd": 3, "feeder": 3, "pump_s": 2, "pump_m": 4, "fan_big": 6, "fan": 3,
       "burner": 2, "boiler": 3, "sh": 9, "eco": 9, "sb": 3, "sv": 3, "mov": 2, "cv": 2, "damper": 2,
       "inst": 1, "inst_a": 2, "anal": 3, "panel": 2, "dcs": 4, "conv": 3, "bf": 6, "cat": 7, "scrub": 14,
       "tower": 3, "tank": 2, "hx": 3, "camera": 1, "crane": 3, "hoist": 1, "turbine": 18, "gen": 5,
       "acc": 5, "tr": 2, "egen": 4, "comp": 2, "struct": 3, "misc": 2, "scale": 2, "door": 2, "hopper": 3}
SIZE = {"stoker": 6, "furnace": 5, "sh": 4, "bf": 6, "cat": 5, "scrub": 6, "turbine": 5, "fan_big": 4,
        "inst": 2, "inst_a": 2, "anal": 2, "panel": 3, "dcs": 3}
KNAME = {"OVERHAUL": None, "REPLACE": "更新", "UPGRADE": "改造", "REPAIR": "補修", "INSPECT": None}


def _cost(e, kind):
    c = CLS[e.cls]
    return {"OVERHAUL": c["coh"], "REPLACE": c["crep"], "UPGRADE": c["crep"], "REPAIR": 900, "INSPECT": 300}[kind]


def _act_name(e, x):
    c = CLS[e.cls]
    k = x["kind"]
    if k == "OVERHAUL":
        return c.get("oh_name", "分解整備")
    if k == "INSPECT":
        return c.get("ins_name", "点検")
    return KNAME[k]


def big(e, x):
    return x["kind"] in ("OVERHAUL", "REPLACE") and _cost(e, x["kind"]) >= 30000 and e.cls != "turbine"


# 見本で手書きした契約の工程 id と作業（tag, kind, 内容）
SAMPLE_WORKS = {
    "a03": [("1IN-005", "OVERHAUL", "火格子部分交換 24 枚・サイドシール交換")],
    "a04": [("1IN-004", "INSPECT", "点検"), ("1IN-006", "INSPECT", "点検")],
    "a05": [("1IN-003", "INSPECT", "点検"), ("1IN-001", "INSPECT", "点検")],
    "a06": [("1IN-011", "OVERHAUL", "耐火物部分補修 1.8m²")],
    "a07": [("1BL-001", "INSPECT", "肉厚測定"), ("1BL-002", "INSPECT", "肉厚測定"), ("1BL-003", "INSPECT", "肉厚測定")],
    "a08": [("1BL-004", "OVERHAUL", "分解整備"), ("1BL-005", "OVERHAUL", "分解整備")],
    "a09": [("1BL-012", "OVERHAUL", "吹出し調整"), ("1BL-013", "OVERHAUL", "吹出し調整"), ("1BL-014", "OVERHAUL", "吹出し調整")],
    "a10": [("1GT-001", "OVERHAUL", "噴霧ノズル交換 6 本")],
    "a11": [("1GT-014", "REPAIR", "破損ろ布 2 本交換（市予備品）")],
    "a12": [("1GT-027", "INSPECT", "点検・清掃")],
    "a13": [("1GT-035", "INSPECT", "内部点検")],
    "a14": [("1FD-006", "INSPECT", "点検・振動測定"), ("1FD-001", "INSPECT", "点検")],
    "a16": [("1BL-001", "INSPECT", "水圧試験")],
    "q03": [("2BL-001", "INSPECT", "肉厚測定"), ("2BL-002", "INSPECT", "肉厚測定 最小 3.1mm"), ("2BL-003", "INSPECT", "肉厚測定")],
    "q04": [("2GT-014", "OVERHAUL", "ろ布全数交換 288 本（市支給品）")],
    "q05": [("2IN-005", "OVERHAUL", "火格子部分交換 30 枚"), ("2IN-004", "INSPECT", "点検"), ("2IN-006", "INSPECT", "点検")],
    "q06": [("2IN-011", "OVERHAUL", "耐火物部分補修 2.5m²")],
    "q07": [("2GT-036", "OVERHAUL", "分解整備・羽根車交換・ライナ交換")],
    "q08": [("2BL-002", "REPAIR", "1次過熱器管 3 本 部分取替（追加）")],
    "q09": [("2GT-027", "INSPECT", "点検・触媒抜取り")],
    "q10": [("2BL-006", "OVERHAUL", "分解整備"), ("2BL-007", "OVERHAUL", "分解整備")],
    "q11": [("2GT-035", "INSPECT", "内部点検")],
    "q12": [("2FD-006", "OVERHAUL", "羽根車 肉盛補修・現地バランス、軸受交換")],
    "q13": [("2IN-003", "INSPECT", "点検"), ("2IN-001", "INSPECT", "点検")],
    "q15": [("2BL-001", "INSPECT", "水圧試験（定期事業者検査）")],
    "q16": [("2FD-001", "INSPECT", "点検"), ("2FD-002", "INSPECT", "点検")],
    "c01": [("0RF-013", "OVERHAUL", "巻上ワイヤロープ交換・ブレーキ整備")],
    "c02": [("0ST-001", "OVERHAUL", "開放点検（定期事業者検査）"), ("0ST-002", "INSPECT", "開放点検"), ("0ST-003", "INSPECT", "点検")],
    "c04": [("0ST-031", "INSPECT", "開放・第一種圧力容器 性能検査")],
    "c05": [("0ST-0%02d" % i, "INSPECT", "減速機 油交換・点検") for i in range(13, 19)],
    "c06": [("0MS-001", "INSPECT", "点検"), ("1FD-011", "INSPECT", "内筒 肉厚測定"), ("3FD-011", "INSPECT", "内筒 肉厚測定")],
    "c07": [("0EL-001", "INSPECT", "精密点検"), ("0EL-002", "INSPECT", "精密点検"), ("0EL-008", "INSPECT", "精密点検")],
    "c08": [("0UT-005", "INSPECT", "清掃・点検（散水板 割れ）"), ("0UT-006", "INSPECT", "清掃・点検")],
}
SAMPLE_KEYS = {("teishu", 1, 2013): "H25-1", ("teishu", 2, 2024): "R6-2", ("kyotsu", 0, 2024): "R6-C"}


class Model:
    def __init__(self, eqs):
        self.eqs = eqs
        self.tag = {e.tag: e for e in eqs}
        self.contracts = OrderedDict()
        self.direct = []     # 直営の作業 dict(date, tag, text, n, h, wid, part)
        self.rng = random.Random(SEED + 99)
        self.unrecorded = []
        self._build()

    # ------------------------------------------------------------
    def ckey(self, fy, kind, unit, extra=""):
        return "%s-%s-%s%s" % (fy, kind, unit, extra)

    def _build(self):
        rng = self.rng
        # 履歴を年度・号炉ごとに分ける
        for fy in range(FY0, FY1 + 1):
            outs, ky = outages(fy)
            # 改良工事（資料は完成図書側。CSV と台帳・計画にだけ出る）
            self._kairyo(fy)
            for u, st in sorted(outs.items(), key=lambda kv: kv[1]):
                self._teishu(fy, u, st)
            self._kyotsu(fy, ky)
        self._assign_rest()
        self._direct_routine()

    def _kairyo(self, fy):
        P = "くぬぎ台環境センター"
        if fy in (2001, 2002):
            u = 2 if fy == 2001 else 3
            k = "DXN-%d" % u
            self.contracts[k] = Contract(k, fy, u, "kairyo", "%s %s%d号炉ダイオキシン類対策工事" % (wareki(fy), P, u),
                                         "汐見重工環境", D(fy, 6, 1), D(fy + 1, 3, 15), method="随意契約（特命）", budget="工事請負費",
                                         amount=int(1450000000 * 1.05), docs=False, note="EP→BF、減温塔・触媒反応塔追設")
        if fy == 2003:
            k = "AM-2003"
            self.contracts[k] = Contract(k, fy, 0, "kairyo", "%s 灰溶融設備設置工事" % P, "汐見重工環境", D(2003, 5, 12),
                                         D(2004, 3, 19), method="一般競争入札", budget="工事請負費", amount=int(1980000000 * 1.05), docs=False,
                                         note="電気抵抗式 30t/日")
        if fy == 2011:
            k = "KIKAN-0"
            self.contracts[k] = Contract(k, fy, 0, "kairyo", "%s基幹的設備改良工事" % P, "汐見重工環境", D(2011, 7, 1),
                                         D(2015, 3, 20), method="一般競争入札", budget="工事請負費", amount=int(4620000000 * 1.05), docs=False,
                                         note="継続費 H23〜H26（共通→1号→2号→3号）")
        for u in (1, 2, 3):
            if KIKAN_FY[u] == fy:
                self.contracts["KIKAN-%d" % u] = Contract("KIKAN-%d" % u, fy, u, "kairyo_part", "基幹的設備改良工事（%d号炉）" % u,
                                                          "汐見重工環境", D(fy, 6, 3), D(fy, 12, 20), docs=False)
        if fy == 2013:
            k = "AMSTOP-2013"
            self.contracts[k] = Contract(k, fy, 0, "kairyo", "灰溶融設備休止措置業務委託", "汐見重工環境", D(2013, 10, 7),
                                         D(2013, 11, 29), cdate=D(2013, 9, 27), amount=6825000, docs=False, note="電力費削減のため休止")

    def _events(self, fy, unit, pred=lambda e: True):
        out = []
        for e in self.eqs:
            if e.unit != unit or not pred(e):
                continue
            for x in e.events:
                if x["fy"] == fy and not x.get("_used"):
                    out.append((e, x))
        return out

    # ------------------------------------------------------------ 炉別 定期整備
    def _teishu(self, fy, u, st):
        rng = random.Random("%d-%d-%d" % (SEED, fy, u))
        P = "くぬぎ台環境センター"
        key = self.ckey(fy, "T", u)
        sk = SAMPLE_KEYS.get(("teishu", u, fy))
        if sk:
            return self._sample_contract(sk, key, fy, u)
        evs = self._events(fy, u)
        if fy == 2026 and st > CUTOFF:
            # 下期の定修は契約のみ（予定）。作業の実績は記録時点の対象外
            c = Contract(key, fy, u, "teishu", "%s %s %d号炉定期整備業務委託" % (wareki(fy), P, u), "汐見重工環境",
                         st, st + dt.timedelta(days=33), docs=False)
            c.planned_only = True
            self.contracts[key] = c
            self._plan_only_tasks(c, st)
            for e, x in evs:
                x["_used"] = True
                self.unrecorded.append(dict(tag=e.tag, fy=fy, what=x["what"], reason="2026 年度 下期（記録時点の上期より後）"))
            return
        big_evs = [(e, x) for e, x in evs if big(e, x)]
        small = [(e, x) for e, x in evs if not big(e, x) and x.get("contract") != "DIRECT"]
        c = Contract(key, fy, u, "teishu", "%s %s %d号炉定期整備業務委託" % (wareki(fy), P, u), "汐見重工環境", st, st)
        c.outage = st
        self.contracts[key] = c
        tasks = []
        d0 = st
        stop_end = add_workdays(d0, 3) if workday(d0) else add_workdays(next_workday(d0), 3)
        tasks.append(Task("t01", "炉立下げ・冷却（市）", "市", (d0, d0 + dt.timedelta(days=2)), None, ["%dIN-011" % u], "city"))
        w0 = next_workday(d0 + dt.timedelta(days=2))
        tasks.append(Task("t02", "足場組立（炉内・ボイラ・BF）", "鳶", (w0, add_workdays(w0, 3)), None, [], "scaf", 6))
        # 標準の点検（毎年）
        std = [("B", "ボイラ水管・過熱器 肉厚測定", [("%dBL-%03d" % (u, i), "INSPECT", "肉厚測定") for i in (1, 2, 3)], 3, 4),
               ("A", "ストーカ・給じん装置 点検", [("%dIN-%03d" % (u, i), "INSPECT", "点検") for i in (3, 4, 5, 6)], 3, 4),
               ("C", "BF・触媒反応塔・洗煙塔 点検", [("%dGT-%03d" % (u, i), "INSPECT", "点検") for i in (14, 27, 35)], 4, 4),
               ("D", "誘引通風機・押込送風機 点検", [("%dFD-%03d" % (u, i), "INSPECT", "点検") for i in (6, 1)], 2, 3),
               ("E", "電気・計装 点検・校正", [("%dEL-%03d" % (u, i), "INSPECT", "点検") for i in (1, 2, 3)], 7, 3)]
        # 履歴の作業を班ごとにまとめる
        groups = OrderedDict()
        used_tags = set()
        for e, x in small:
            crew = "築炉" if e.cls == "furnace" else CREW_OF.get(e.sys, "D")
            if e.cls in ("pump_s", "pump_m", "fan", "comp") and e.sys in ("BL", "GT", "IN"):
                crew = "D"
            if e.cls in ("inst", "inst_a", "anal", "panel", "dcs", "camera"):
                crew = "E"
            g = (crew, e.cls if e.cls in ("stoker", "furnace", "sb", "sv", "bf", "fan_big", "sh", "cat", "boiler") else e.sys, x["kind"] if x["kind"] != "INSPECT" else "INSPECT")
            groups.setdefault(g, []).append((e, x))
        # 標準点検と重なる INSPECT は標準点検の作業に入れる
        std_tags = {t for s in std for t, _, _ in s[2]}
        seq = {}
        crews = OrderedDict()
        tid = 3
        cur = {cr: add_workdays(w0, 1) for cr in "ABCDE"}
        cur["築炉"] = add_workdays(w0, 4)
        cur["E"] = add_workdays(w0, 6)
        allt = []
        for crew, nm, ws, dur, n in std:
            allt.append((crew, nm, ws, dur, n, "std", None))
        for (crew, g, kind), lst in groups.items():
            lst2 = [(e, x) for e, x in lst if not (x["kind"] == "INSPECT" and e.tag in std_tags)]
            for e, x in lst:
                if x["kind"] == "INSPECT" and e.tag in std_tags:
                    x["_into_std"] = True
            if not lst2:
                continue
            names = [e.name for e, _ in lst2]
            act = _act_name(lst2[0][0], lst2[0][1])
            if len(names) > 3:
                nm = "%s 他%d台 %s" % ("・".join(names[:2]), len(names) - 2, act)
            else:
                nm = "%s %s" % ("・".join(names), act)
            dur = max(DUR.get(e.cls, 2) for e, _ in lst2) + (1 if len(lst2) > 3 else 0)
            if kind == "INSPECT":
                dur = max(1, dur // 2)
            n = max(SIZE.get(e.cls, 3) for e, _ in lst2)
            allt.append((crew, nm, [(e.tag, x["kind"], x["what"], x) for e, x in lst2], dur, n, "event", lst2[0][0].cls))
        # 追加補修（約 1 割の定修で、開放後の発見による）
        add = None
        if rng.random() < 0.12 and fy < 2026:
            cand = [e for e in self.eqs if e.unit == u and e.cls in ("eco", "tower", "conv", "damper", "hx", "sb")]
            ea = rng.choice(cand)
            xa = dict(fy=fy, kind="REPAIR", what="【追加】%s 補修（開放時に損傷を発見）" % ea.name, _added=True)
            ea.events.append(xa)
            add = (ea, xa)
        for k2 in list(cur):
            if k2 != "築炉":
                cur[k2 + "2"] = add_workdays(w0, 2)
        for crew, nm, ws, dur, n, kind, cls in allt:
            if crew in ("A", "B", "C", "D", "E") and kind == "event" and cur[crew + "2"] < cur[crew]:
                crew = crew + "2"
            a = cur.get(crew, add_workdays(w0, 1))
            b = add_workdays(a, dur)
            cur[crew] = next_workday(b)
            tasks.append(Task("t%02d" % tid, nm, crew, (a, b), None, [w[0] for w in ws], kind, n, note="", cls=cls))
            tasks[-1]._ws = ws
            tid += 1
        # 遅れ（約 15%）
        for t in tasks[2:]:
            if t.kind == "event" and rng.random() < 0.15:
                ext = rng.choice([1, 1, 2])
                t.act = (t.plan[0], add_workdays(t.plan[1], ext + 1))
                t.note = "%s +%d日" % (rng.choice(["部品手直し", "摩耗範囲拡大", "雨天で搬入遅れ", "溶接手直し", "付着物多く清掃追加"]), ext)
        if add:
            ea, xa = add
            mid = add_workdays(w0, 5)
            t = Task("t%02d" % tid, "【追加】%s 補修" % ea.name, CREW_OF.get(ea.sys, "D"), None, (mid, add_workdays(mid, 3)),
                     [ea.tag], "add", 3, note="開放後に発見・協議", cls=ea.cls)
            t._ws = [(ea.tag, "REPAIR", xa["what"], xa)]
            tasks.append(t)
            tid += 1
        work_end_p = max(t.plan[1] for t in tasks if t.plan)
        work_end_a = max(t.b for t in tasks)
        def tail(base):
            hp = next_workday(base)
            return hp
        hp_p, hp_a = tail(work_end_p), tail(work_end_a)
        tasks.append(Task("t%02d" % tid, "ボイラ水圧試験", "B", (hp_p, hp_p), (hp_a, hp_a) if hp_a != hp_p else None,
                          ["%dBL-001" % u], "std", 4)); tasks[-1]._ws = [("%dBL-001" % u, "INSPECT", "水圧試験")]; tid += 1
        tasks.append(Task("t%02d" % tid, "足場解体・清掃", "鳶", (hp_p, add_workdays(hp_p, 3)),
                          (hp_a, add_workdays(hp_a, 3)) if hp_a != hp_p else None, [], "scaf", 6)); tid += 1
        su_p = add_workdays(hp_p, 4) + dt.timedelta(days=0)
        su_a = add_workdays(hp_a, 4)
        tasks.append(Task("t%02d" % tid, "耐火物乾燥・立上げ（市）", "市", (su_p, su_p + dt.timedelta(days=1)),
                          (su_a, su_a + dt.timedelta(days=1)) if su_a != su_p else None, ["%dIN-011" % u], "city")); tid += 1
        tr_p = next_workday(su_p + dt.timedelta(days=1))
        tr_a = next_workday(su_a + dt.timedelta(days=1))
        tasks.append(Task("t%02d" % tid, "試運転・調整", "全", (tr_p, add_workdays(tr_p, 3)), (tr_a, add_workdays(tr_a, 3)) if tr_a != tr_p else None,
                          [], "trial", 3)); tid += 1
        c.tasks = tasks
        c.end = add_workdays(tasks[-1].plan[1], 5)
        c.cdate = st - dt.timedelta(days=rng.randint(20, 35))
        # 作業（world）を作る
        cost = 0
        for t in tasks:
            for w in getattr(t, "_ws", []):
                tag, kind, what = w[0], w[1], w[2]
                if len(w) > 3:
                    w[3]["_used"] = True
                wid = add_work(tag, kind, what, key, t.a, t.b, "汐見重工環境", fy=fy)
                t.works.append(wid)
                cost += _cost(self.tag[tag], kind)
        for e, x in evs:
            if x.get("_into_std"):
                x["_used"] = True
        for e, x in big_evs:
            self._hoshu(fy, u, e, x, c)
        base = 26000 + 12000 * (1 + (fy - 2001) * 0.012)
        c.amount = int(round((base + cost * 0.9) * 1000 * (1 + tax(fy, c.cdate)) / 10000.0)) * 10000
        if add:
            c.amount2 = c.amount + int(round(_cost(add[0], "REPAIR") * 1.6 * 1000 * (1 + tax(fy)) / 10000.0)) * 10000

    def _plan_only_tasks(self, c, st):
        u = c.unit
        w0 = next_workday(st + dt.timedelta(days=2))
        c.tasks = [Task("t01", "炉立下げ・冷却（市）", "市", (st, st + dt.timedelta(days=2)), None, [], "city"),
                   Task("t02", "足場組立・解体", "鳶", (w0, add_workdays(w0, 20)), None, [], "scaf", 6),
                   Task("t03", "定期整備（燃焼・ボイラ・排ガス・回転機・電計）", "全", (add_workdays(w0, 2), add_workdays(w0, 18)), None, [], "std", 20),
                   Task("t04", "水圧試験・立上げ・試運転", "市/全", (add_workdays(w0, 19), add_workdays(w0, 24)), None, [], "trial", 3)]
        for t in c.tasks:
            t.act = None
        c.end = add_workdays(w0, 28)

    def _hoshu(self, fy, u, e, x, parent):
        rng = random.Random("%d-%s-%d-h" % (SEED, e.tag, fy))
        P = "くぬぎ台環境センター"
        act = _act_name(e, x)
        key = self.ckey(fy, "H", u, "-" + e.tag)
        st = parent.outage if parent and parent.outage else D(fy, 10, 1)
        w0 = add_workdays(next_workday(st + dt.timedelta(days=2)), 2)
        dur = DUR.get(e.cls, 6) + 4
        name = "%s %s %s%s%s工事" % (wareki(fy), P, ("%d号炉" % u) if u else "", e.name.replace(" ", ""), act.replace("・", ""))
        c = Contract(key, fy, u, "hoshu", name, "汐見重工環境", w0, add_workdays(w0, dur + 6), method="随意契約（特命）",
                     budget="工事請負費", outage=st)
        steps = [("準備・養生・仮設", 2, 4), ("既設 撤去・はつり", 3, 5), ("%s 施工" % act, dur - 4, SIZE.get(e.cls, 5)),
                 ("検査（膜厚・ピンホール／RT 等）", 1, 3), ("復旧・片付け", 2, 4)]
        tasks = []
        a = w0
        for i, (nm, d, n) in enumerate(steps, 1):
            b = add_workdays(a, d)
            t = Task("h%02d" % i, "%s %s" % (e.name, nm) if i in (2, 3) else nm, "H", (a, b), None, [e.tag], "hoshu", n, cls=e.cls)
            tasks.append(t)
            a = next_workday(b)
        if rng.random() < 0.3:
            t = tasks[2]
            t.act = (t.plan[0], add_workdays(t.plan[1], 2))
            t.note = "施工範囲 増 +1日"
            for t2 in tasks[3:]:
                t2.act = (add_workdays(t2.plan[0], 2), add_workdays(t2.plan[1], 2))
        c.tasks = tasks
        c.end = add_workdays(tasks[-1].b, 6)
        wid = add_work(e.tag, x["kind"], x["what"] if x["what"] not in ("分解整備", "更新") else act, key, tasks[1].a, tasks[3].b,
                       "汐見重工環境", fy=fy)
        tasks[2].works.append(wid)
        x["_used"] = True
        c.amount = int(round(_cost(e, x["kind"]) * 1.05 * 1000 * (1 + tax(fy)) / 10000.0)) * 10000
        c.cdate = st - dt.timedelta(days=rng.randint(40, 70))
        self.contracts[key] = c

    # ------------------------------------------------------------ 全炉停止時共通
    def _kyotsu(self, fy, st):
        rng = random.Random("%d-ky-%d" % (SEED, fy))
        P = "くぬぎ台環境センター"
        key = self.ckey(fy, "C", 0)
        sk = SAMPLE_KEYS.get(("kyotsu", 0, fy))
        if sk:
            return self._sample_contract(sk, key, fy, 0)
        evs = self._events(fy, 0, lambda e: True)
        if fy == 2026:
            c = Contract(key, fy, 0, "kyotsu", "%s %s 全炉停止時共通設備整備業務委託" % (wareki(fy), P), "汐見重工環境",
                         st, st + dt.timedelta(days=24), docs=False)
            c.planned_only = True
            self.contracts[key] = c
            w0 = st
            c.tasks = [Task("c01", "共通設備 整備（タービン点検・受変電・脱気器 他）", "F", (w0, add_workdays(w0, 18)), None, [], "std", 10),
                       Task("c02", "全炉停止", "市", (w0 + dt.timedelta(days=7), w0 + dt.timedelta(days=14)), None, [], "city")]
            for e, x in evs:
                if not (e.idle and x["kind"] == "INSPECT"):
                    x["_used"] = True
                    self.unrecorded.append(dict(tag=e.tag, fy=fy, what=x["what"], reason="2026 年度 下期（記録時点の上期より後）"))
            return
        c = Contract(key, fy, 0, "kyotsu", "%s %s 全炉停止時共通設備整備業務委託" % (wareki(fy), P), "汐見重工環境", st, st)
        c.outage = st
        self.contracts[key] = c
        zs = add_workdays(st, 5)
        while zs.weekday() != 2:
            zs += dt.timedelta(days=1)
        ze = zs + dt.timedelta(days=7)
        c.zentei = (zs, ze)
        tasks = [Task("c01", "各炉 立下げ（市）／全炉停止", "市", (zs - dt.timedelta(days=1), ze), None, [], "city")]
        sat = zs
        while sat.weekday() != 5:
            sat += dt.timedelta(days=1)
        std = [("E", "受電停止・高圧受配電盤 精密点検（%d/%d 停電）" % (sat.month, sat.day), [("0EL-001", "INSPECT", "年次点検"), ("0EL-002", "INSPECT", "年次点検")], (add_workdays(sat, 0) - dt.timedelta(days=1), sat + dt.timedelta(days=1)), 5),
               ("F", "脱気器 開放・第一種圧力容器 性能検査", [("0ST-031", "INSPECT", "性能検査")], (zs, add_workdays(zs, 5)), 3)]
        if fy >= 2003 and fy <= AM_STOP_FY:
            std.append(("F", "灰溶融炉 点検", [("0AM-001", "INSPECT", "点検")], (add_workdays(st, 1), add_workdays(st, 4)), 4))
        # 共通の履歴作業
        small = []
        for e, x in evs:
            if x.get("_used"):
                continue
            if e.idle and fy > AM_STOP_FY:
                continue   # 休止設備の目視は直営
            if e.cls in ("crane", "hoist") and x["kind"] == "INSPECT":
                ck = "%d-CRANE" % fy
                wid = add_work(e.tag, "INSPECT", x["what"], ck, D(fy, 11, 10), D(fy, 11, 14), "篠目クレーン", fy=fy, src="itaku")
                WORKS[wid]["manDays"] = None
                self.unrecorded.append(dict(work=wid, tag=e.tag, fy=fy, what=x["what"],
                                            reason="クレーン保守点検業務委託（篠目クレーン）の点検記録は別綴じ。契約は CSV に出る"))
                x["_used"] = True
                continue
            if x.get("contract") == "DIRECT":
                continue
            if x.get("contract", "") and x["contract"] and x["contract"].split("-")[0] in ("KIKAN", "DXN", "AM"):
                continue
            if big(e, x):
                self._hoshu(fy, 0, e, x, c)
                continue
            if e.cls in ("pump_s", "misc", "inst", "camera", "door") and x["kind"] in ("OVERHAUL", "REPAIR") and rng.random() < 0.5:
                x["contract"] = "DIRECT"
                continue
            small.append((e, x))
        groups = OrderedDict()
        for e, x in small:
            crew = "T" if e.sys == "ST" and e.cls in ("turbine", "gen", "pump_s", "hx", "cv", "mov", "panel", "inst_a", "misc") and e.name.find("脱気") < 0 else (
                "E" if e.sys in ("EL", "IC") or e.cls in ("inst", "inst_a", "anal", "panel", "dcs", "camera", "tr") else "F")
            g = (crew, e.sys if e.cls != "turbine" else "turbine", x["kind"])
            groups.setdefault(g, []).append((e, x))
        cur = {"F": add_workdays(st, 1), "T": add_workdays(st, 1), "E": add_workdays(st, 2), "鳶": zs}
        tid = 2
        for crew, nm, ws, pl, n in std:
            t = Task("c%02d" % tid, nm, crew, pl, None, [w[0] for w in ws], "std", n)
            t._ws = ws
            tasks.append(t)
            tid += 1
        for (crew, g, kind), lst in groups.items():
            names = [e.name for e, _ in lst]
            act = _act_name(lst[0][0], lst[0][1])
            nm = ("%s 他%d台 %s" % ("・".join(names[:2]), len(names) - 2, act)) if len(names) > 3 else "%s %s" % ("・".join(names), act)
            dur = max(DUR.get(e.cls, 2) for e, _ in lst) + (1 if len(lst) > 3 else 0)
            if kind == "INSPECT":
                dur = max(1, dur // 2)
            a = cur[crew]
            # 全停が要る作業（タービン・受変電は別、ここは全停の中に寄せる）
            if any(e.sys in ("FD", "MS", "CH") for e, _ in lst):
                a = max(a, zs)
            b = add_workdays(a, dur)
            cur[crew] = next_workday(b)
            t = Task("c%02d" % tid, nm, crew, (a, b), None, [e.tag for e, _ in lst], "event", max(SIZE.get(e.cls, 3) for e, _ in lst),
                     cls=lst[0][0].cls)
            t._ws = [(e.tag, x["kind"], x["what"], x) for e, x in lst]
            if rng.random() < 0.12:
                t.act = (a, add_workdays(b, 2))
                t.note = "+1日（%s）" % rng.choice(["部品待ち", "芯出し手直し", "検査日の変更"])
            tasks.append(t)
            tid += 1
        tasks.append(Task("c%02d" % tid, "各炉 立上げ（市）", "市", (ze + dt.timedelta(days=1), ze + dt.timedelta(days=2)), None, [], "city"))
        c.tasks = tasks
        c.start = min(t.a for t in tasks)
        c.end = add_workdays(max(t.b for t in tasks), 4)
        c.cdate = c.start - dt.timedelta(days=rng.randint(20, 30))
        cost = 0
        for t in tasks:
            for w in getattr(t, "_ws", []):
                if len(w) > 3:
                    w[3]["_used"] = True
                wid = add_work(w[0], w[1], w[2], key, t.a, t.b, "汐見重工環境" if t.crew != "T" else "葦原タービン", fy=fy)
                t.works.append(wid)
                cost += _cost(self.tag[w[0]], w[1])
        c.amount = int(round((18000 + cost * 0.85) * 1000 * (1 + tax(fy, c.cdate)) / 10000.0)) * 10000

    # ------------------------------------------------------------ 見本の契約（手書きの工程と日報を使う）
    def _sample_contract(self, sk, key, fy, u):
        ct = S_CON[sk]
        c = Contract(key, fy, u, "teishu" if u else "kyotsu", ct["name"].replace("くぬぎ台環境センター", "くぬぎ台環境センター"),
                     "汐見重工環境", D(*ct["start"]), D(*ct["end"]))
        c.sample = sk
        if sk == "H25-1":
            for tid, kubun, name, crew, a, b, tags, note in S.H25_1["tasks"]:
                t = Task(tid, name, crew, (a, b), None, tags, "city" if crew == "市" else ("scaf" if crew == "鳶" else "std"), 4, note=note)
                t.kubun = kubun
                c.tasks.append(t)
            c.amount = 52500000
            c.cdate = D(2013, 5, 8)
            c.outage = D(2013, 5, 13)
        else:
            for cc, tid, name, crew, plan, act, tags in S.R6:
                if cc != sk:
                    continue
                c.tasks.append(Task(tid, name, crew, plan, act, tags,
                                    "city" if crew == "市" else ("scaf" if crew == "鳶" else ("add" if plan is None else "std")), 4))
            if sk == "R6-2":
                c.amount, c.amount2, c.cdate, c.outage = 59730000, 63415000, D(2024, 9, 13), D(2024, 10, 7)
            else:
                c.amount, c.cdate, c.zentei = 71280000, D(2024, 9, 13), (D(2024, 10, 16), D(2024, 10, 23))
        # その年度・その炉の履歴は手書きの工程に置き換える（汎用の履歴は使わない）
        for e, x in self._events(fy, u):
            if x.get("contract") == "DIRECT" or (e.idle and x["kind"] == "INSPECT"):
                continue
            x["_used"] = True
            x["_replaced"] = True
        sizes = {}
        for t in c.tasks:
            for tag, kind, what in SAMPLE_WORKS.get(t.id, []):
                wid = add_work(tag, kind, what, key, t.a, t.b, "汐見重工環境" if t.crew != "T" else "葦原タービン", fy=fy, src="sample")
                t.works.append(wid)
            if t.tags:
                t.cls = self.tag[t.tags[0]].cls
        self.contracts[key] = c
        # 見本の日報（手書き）を使う日
        if sk == "R6-2":
            c.daily_fixed = S.R6_2_DAILY
        elif sk == "R6-C":
            c.daily_fixed = S.R6_C_DAILY
        else:
            c.daily_fixed = {}
            c.daily_fixed_m1 = S.H25_1_DAILY

    # ------------------------------------------------------------ 残り（直営・改良工事）
    def _assign_rest(self):
        for e in self.eqs:
            for x in e.events:
                if x.get("_used"):
                    continue
                fy = x["fy"]
                c = x.get("contract") or ""
                pre = c.split("-")[0] if c else ""
                if pre in ("KIKAN", "DXN", "AM"):
                    ck = c if c in self.contracts else ("KIKAN-%d" % e.unit if pre == "KIKAN" else c)
                    cc = self.contracts.get(ck) or self.contracts.get("KIKAN-0")
                    a, b = (cc.start, cc.end) if cc else (D(fy, 6, 1), D(fy, 12, 1))
                    wid = add_work(e.tag, x["kind"], x["what"], ck, a, b, "汐見重工環境", fy=fy, src="kairyo")
                    WORKS[wid]["manDays"] = None
                    self.unrecorded.append(dict(work=wid, tag=e.tag, fy=fy, what=x["what"],
                                                reason="改良工事（%s）の記録は完成図書側に保管。本データには契約（CSV）・台帳・計画だけが出る" % ck))
                    x["_used"] = True
                    continue
                if e.unit and ((e.unit in DXN_FY and DXN_FY[e.unit] == fy) or KIKAN_FY.get(e.unit) == fy):
                    ck = ("DXN-%d" % e.unit) if DXN_FY.get(e.unit) == fy else "KIKAN-%d" % e.unit
                    cc = self.contracts.get(ck)
                    a, b = (cc.start, cc.end) if cc else (D(fy, 6, 1), D(fy, 12, 1))
                    wid = add_work(e.tag, x["kind"], x["what"], ck, a, b, "汐見重工環境", fy=fy, src="kairyo")
                    WORKS[wid]["manDays"] = None
                    self.unrecorded.append(dict(work=wid, tag=e.tag, fy=fy, what=x["what"],
                                                reason="改良工事の年度で定期整備の契約が無く、改良工事の中で実施（記録は完成図書側）"))
                    x["_used"] = True
                    continue
                if fy == 2000:
                    continue
                # 直営
                x["contract"] = "DIRECT"
                self._direct_event(e, x)
                x["_used"] = True

    def _direct_event(self, e, x):
        fy = x["fy"]
        rng = random.Random("%d-%s-%d-d" % (SEED, e.tag, fy))
        # 号炉の機器は、その炉の定修中に直営で行う。共通は年度内の平日
        cands = [c for c in self.contracts.values() if c.fy == fy and c.kind in ("teishu", "kyotsu") and (c.unit == e.unit or (e.unit == 0 and c.kind == "kyotsu"))]
        if cands and cands[0].tasks:
            c0 = cands[0]
            day = add_workdays(c0.tasks[0].a + dt.timedelta(days=3), rng.randint(1, 8))
        else:
            day = add_workdays(D(fy, rng.choice([5, 6, 7, 10, 11, 1, 2]), rng.randint(1, 20)), 1)
        if fy == 2026 and day > CUTOFF:
            day = add_workdays(D(2026, 6, 2), rng.randint(1, 30))
        days = 1 if e.cls in ("inst", "inst_a", "camera") else rng.choice([1, 2, 2, 3])
        dl = [day]
        for _ in range(days - 1):
            dl.append(next_workday(dl[-1]))
        wid = add_work(e.tag, x["kind"], x["what"], "DIRECT", dl[0], dl[-1], "市直営", fy=fy, src="direct")
        n = rng.choice([2, 3, 3, 4])
        for i, d in enumerate(dl):
            ph = "first" if i == 0 else ("last" if i == len(dl) - 1 else "mid")
            self.direct.append(dict(date=d, tag=e.tag, wid=wid, n=n, h=rng.choice([2.0, 3.0, 4.0, 6.0, 7.5]) if len(dl) > 1 else rng.choice([1.5, 2.0, 3.0, 4.0]),
                                    phase=ph, kind=x["kind"], what=x["what"], cls=e.cls))

    def _direct_routine(self):
        """直営班の定常作業（巡視・給油・清掃・計器の点検）。作業（world）には入れない＝点検の記録だけ"""
        rng = random.Random(SEED + 555)
        pool = [e for e in self.eqs if not e.idle and e.cls in ("pump_s", "pump_m", "conv", "fan", "comp", "misc", "inst_a", "hyd", "feeder")]
        for fy in range(FY0, FY1 + 1):
            d = D(fy, 4, 1)
            end = min(D(fy + 1, 3, 31), CUTOFF if fy == 2026 else D(fy + 1, 3, 31))
            while d <= end:
                if workday(d) and d.weekday() < 5 and rng.random() < 0.55:
                    e = rng.choice(pool)
                    self.direct.append(dict(date=d, tag=e.tag, wid=None, n=rng.choice([1, 2, 2]), h=rng.choice([0.5, 1.0, 1.5, 2.0]),
                                            phase="routine", kind="ROUTINE", what="", cls=e.cls))
                d += dt.timedelta(days=1)
        self.direct.sort(key=lambda r: (r["date"], r["tag"]))

    # ------------------------------------------------------------
    def by_fy(self, fy):
        return [c for c in self.contracts.values() if c.fy == fy]


S_CON = None


def build(eqs):
    global S_CON
    from world import CONTRACTS
    S_CON = CONTRACTS
    WORKS.clear()
    return Model(eqs)


def apply_direct_samples(m):
    """見本で手書きした直営の作業（H25.5 d1・R6.10 d2）を、その日の直営の行として使う。"""
    import sched as S
    fixed = {}
    # 対象の作業 id を探す
    def wid_of(tag, fy):
        for w in WORKS.values():
            if w["tag"] == tag and w["fy"] == fy and w["contract"] == "DIRECT":
                return w["id"]
        return None
    days = set()
    for mon, rows in S.R6_DIRECT_WEEKS:
        for d, tag, eq, txt, n, h in rows:
            days.add(d)
    m.direct = [r for r in m.direct if r["date"] not in days and not (r["wid"] and WORKS[r["wid"]]["fy"] == 2024 and WORKS[r["wid"]]["tag"] in ("1BL-021", "1AS-001", "0AM-001"))]
    for mon, rows in S.R6_DIRECT_WEEKS:
        for d, tag, eq, txt, n, h in rows:
            tg = tag if tag != "-" else ("1AS-001" if "当て板" in txt else "0EL-023")
            w = wid_of(tg, 2024) if tg in ("1BL-021", "1AS-001", "0AM-001") else None
            m.direct.append(dict(date=d, tag=tg, wid=w, n=n, h=h, phase="fixed", kind="", what="", cls=m.tag[tg].cls, text=txt, eqname=eq,
                                 showtag=(tag != "-")))
    for wid in [wid_of(t, 2024) for t in ("1BL-021", "1AS-001", "0AM-001")]:
        if wid:
            ds = [r["date"] for r in m.direct if r["wid"] == wid]
            WORKS[wid]["start"], WORKS[wid]["end"] = min(ds), max(ds)
    # H25.5 d1
    hd = {d for d, wx, rows in S.H25_DIRECT_DAYS}
    w13 = wid_of("1AS-001", 2013)
    m.direct = [r for r in m.direct if r["date"] not in hd and not (w13 and r["wid"] == w13)]
    tagmap = {"1号 灰押出装置": "1AS-001", "2号 給水ポンプB": "2BL-022", "ピット汚水ポンプ": "0RF-023", "灰溶融炉": "0AM-013"}
    m.h25_fixed = {}
    for d, wx, rows in S.H25_DIRECT_DAYS:
        m.h25_fixed[d] = wx
        for eq, txt, men, tm in rows:
            tg = tagmap[eq]
            a, b = tm.split("〜")
            hh = (int(b.split(":")[0]) * 60 + int(b.split(":")[1]) - int(a.split(":")[0]) * 60 - int(a.split(":")[1])) / 60.0
            n = 1 + int(men.split("外")[1][0]) if "外" in men else int(men[0])
            m.direct.append(dict(date=d, tag=tg, wid=w13 if tg == "1AS-001" else None, n=n, h=hh, phase="fixed", kind="", what="",
                                 cls=m.tag[tg].cls, text=txt, eqname=eq, men=men, tm=tm, showtag=True))
    if w13:
        WORKS[w13]["start"], WORKS[w13]["end"] = dt.date(2013, 5, 21), dt.date(2013, 5, 23)
    m.direct.sort(key=lambda r: (r["date"], r["wid"] is None, r["tag"]))
