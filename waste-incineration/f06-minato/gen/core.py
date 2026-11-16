# -*- coding: utf-8 -*-
"""F06 正解の世界（1 本）：機器 → 年度の作業（world.simulate の実績）→ 契約・停止期間 → 工程（作業項目）→ 日々の作業 → 作業（work）。

すべての帳票はここから書き分ける。world.simulate() が「その年度にその機器に何をしたか」を決め、
本モジュールが「どの契約の、どの停止期間の、どの工程項目で、何日に、誰が」を決める。
"""
import datetime as dt
import random
import zlib
from collections import defaultdict

import equipment as EQ
import world as W

D = dt.date.fromisoformat
TODAY = dt.date(2026, 11, 13)
FY_RANGE = range(2004, 2027)
TURB_YEARS = (2006, 2010, 2014, 2018, 2022, 2026)
ZENRO_SPLIT_FY = 2011  # H23 から全炉停止時の共通設備整備を別契約にした（財務会計システム導入と同時）


def rng(*keys):
    return random.Random(zlib.crc32("|".join(map(str, keys)).encode()) ^ W.SEED)


def nth_monday(y, m, n):
    d = dt.date(y, m, 1)
    while d.weekday() != 0:
        d += dt.timedelta(days=1)
    return d + dt.timedelta(days=7 * (n - 1))


def workdays_from(start, n, sundays=False):
    out, d = [], start
    while len(out) < n:
        if sundays or d.weekday() != 6:
            out.append(d)
        d += dt.timedelta(days=1)
    return out


# ---------------------------------------------------------------- 担当（年代で変わる）
def agent(fy):
    return "m1" if fy <= 2012 else ("m2" if fy <= 2019 else "m3")


def city(fy):
    return "h1" if fy <= 2007 else ("h2" if fy <= 2011 else ("h3" if fy <= 2016 else ("h4" if fy <= 2021 else "h5")))


def reporter(fy):
    return "r1" if fy <= 2016 else "r2"


TEAM_BASE = {
    "機A": ("機械班A", "汐見重工環境", 6), "機B": ("機械班B", "汐見重工環境", 5), "機C": ("機械班C", "汐見重工環境", 5),
    "電計": ("電気計装班", "汐見重工環境", 4), "港南": ("港南工業", "【架空】港南工業", 8), "築炉": ("臨海築炉", "【架空】臨海築炉", 4),
    "足場": ("みなと足場", "【架空】みなと足場工業", 5), "検査": ("検査課", "汐見重工環境", 2),
    "東雲": ("東雲タービンサービス", "【架空】東雲タービンサービス", 6), "保温": ("臨海保温", "【架空】臨海保温工業", 4),
    "溶接": ("溶接班", "汐見重工環境", 4), "改良": ("改良工事班", "汐見重工環境", 12), "据付": ("据付協力 浜町機工", "【架空】浜町機工", 10),
    "電工": ("電気工事 湾岸電工", "【架空】湾岸電工", 6),
}


def foreman(team, fy):
    base = {"機A": ["k1", "k1", "k11"], "機B": ["k2", "k12", "k12"], "機C": ["k6", "k6", "k16"], "電計": ["k3", "k13", "k3"],
            "港南": ["k4", "k4", "k14"], "築炉": ["k5", "k5", "k5"], "足場": ["k8", "k18", "k8"], "検査": ["i1", "i2", "i1"],
            "東雲": ["t1", "t1", "t1"], "保温": ["k9", "k9", "k19"], "溶接": ["k7", "k7", "k17"], "改良": ["m4", "m4", "m4"],
            "据付": ["k21", "k21", "k21"], "電工": ["k22", "k22", "k22"]}[team]
    return base[0 if fy <= 2010 else (1 if fy <= 2018 else 2)]


def team(tk, fy):
    nm, co, size = TEAM_BASE[tk]
    return dict(key=tk, name=nm, co=co, size=size, foreman=foreman(tk, fy))


# ---------------------------------------------------------------- 作業の振り分け
TURB_TAGS = {"C-ST-501", "C-GEN-502", "C-RG-503", "C-MSV-504", "C-GV-505", "C-NRV-506", "C-NRV-507", "C-P-510", "C-P-511",
             "C-HE-512", "C-HE-513", "C-TG-514", "C-GOV-515", "C-VM-516", "C-PNL-517", "C-EXC-518", "C-PNL-519", "C-GC-520",
             "C-EJ-521", "C-HU-522", "C-TK-523", "C-PNL-524"}
ZENRO_CLS = {"header", "dea", "gis", "tr_main", "hv_panel", "tr", "static_stack", "stack_in", "acc", "dcps", "battery", "ups",
             "eg"}
BIG_R_CLS = {"grate", "grate_wc", "sh", "sh2", "sh3", "eco", "rh", "battery", "ups"}
LOCAL_CLS = {"door", "agitator", "dewater", "silo", "tank", "mech_small"}

# 作業項目（工程の行）の種類：(名前, 班, 1 日あたり処理できる機器数の目安, 開始の目安日（定修開始からの日）)
TASKTYPE = {
    "feed": ("給じん装置・油圧ユニット 点検整備", "機A", 4, 8),
    "grate": ("火格子 点検・補修", "機A", 2, 8),
    "furnace": ("耐火物 点検・補修", "築炉", 1, 9),
    "boiler_ut": ("ボイラ水管・過熱器 肉厚測定・点検", "検査", 3, 7),
    "boiler_rep": ("ボイラ 補修・安全弁 吹出し試験", "機A", 2, 12),
    "sb": ("スートブロワ 整備", "機A", 3, 14),
    "fan": ("送風機・通風機・ダンパ 点検整備", "機A", 3, 16),
    "bf": ("集じん器 点検整備（ダンパ・ロータリバルブ・コンベヤ）", "機B", 4, 7),
    "bf_cloth": ("集じん器 ろ布全数交換", "港南", 1, 7),
    "gct": ("減温塔・薬剤供給装置 点検整備", "機B", 4, 10),
    "scrubber": ("洗煙塔 内部点検・循環ポンプ 整備", "機B", 3, 9),
    "rh": ("再加熱器 気密試験・点検", "機B", 1, 15),
    "scr": ("脱硝反応塔 点検・触媒サンプリング", "機B", 2, 10),
    "valve": ("電動弁・調節弁・手動弁 分解整備", "機B", 4, 14),
    "ash": ("灰出し設備 点検整備", "機A", 3, 18),
    "inst": ("主要計装 校正・ループ試験", "電計", 8, 14),
    "elec": ("電気設備 点検（MCC・現場盤・制御盤）", "電計", 5, 15),
    "c_crane": ("ごみクレーン・灰クレーン 年次点検整備", "機C", 1, 8),
    "c_mech": ("共通設備 機器点検整備（ポンプ・圧縮機・コンベヤほか）", "機C", 3, 9),
    "c_inst": ("共通設備 計装 校正", "電計", 8, 16),
    "c_elec": ("共通設備 電気・制御盤 点検", "電計", 6, 17),
    # 全炉停止
    "z_steam": ("蒸気だめ・主蒸気系統 弁 開放点検・安全弁 試験", "機C", 3, 1),
    "z_dea": ("脱気器・給水系統 開放点検", "機C", 2, 1),
    "z_hv": ("特高受変電・高圧配電盤 停電点検", "電工", 6, 1),
    "z_stack": ("煙突・内筒 点検", "足場", 2, 2),
    "z_acc": ("空冷復水器 点検", "機C", 2, 3),
    "z_turb": ("蒸気タービン 補機 点検（タービン停止中）", "機C", 4, 3),
    "z_inst": ("共通計装 校正（全停中）", "電計", 8, 2),
    "z_dcs": ("中央制御装置（DCS）点検", "電計", 4, 2),
}

FURN_TT = {
    1: lambda e: "feed",
    2: lambda e: {"grate": "grate", "grate_wc": "grate", "furnace": "furnace"}.get(e["cls"], "fan" if e["cls"] in ("fan_big", "fan_mid", "damper", "sah", "burner") else None),
    3: lambda e: "boiler_ut" if e["cls"] in ("boiler", "sh", "sh2", "sh3", "eco", "static_hot") else ("sb" if e["cls"].startswith("sootblower") else None),
    4: lambda e: "gct" if e["cls"] in ("gct", "feeder_s", "blower") or e["loop"] in (401, 402, 403, 404, 406, 407) else "bf",
    5: lambda e: "rh" if e["cls"] == "rh" else "scrubber",
    6: lambda e: "scr" if e["cls"] in ("scr",) or e["loop"] in (602, 603, 604, 625, 626, 627) else "fan",
    7: lambda e: "ash",
    8: lambda e: "elec",
}


def task_type(e, kind, contract_kind):
    """1 件の作業が入る作業項目"""
    c = e["cls"]
    if contract_kind == "TURB":
        return "t_main"
    if contract_kind == "ZENRO":
        if e["tag"] in TURB_TAGS:
            return "z_turb"
        if c in ("gis", "tr_main", "hv_panel", "tr", "dcps", "battery", "ups", "eg"):
            return "z_hv"
        if c in ("static_stack", "stack_in"):
            return "z_stack"
        if c in ("acc",):
            return "z_acc"
        if c == "dea" or (e["sysno"] == 4 and c in ("pump_bfp", "cv", "cv_hot", "hv")):
            return "z_dea"
        if c.startswith("inst"):
            return "z_inst"
        if c.startswith("dcs"):
            return "z_dcs"
        return "z_steam"
    if e["unit"] == "C":
        if c in ("crane", "crane_sp", "crane_ash", "crane_m"):
            return "c_crane"
        if c.startswith("inst"):
            return "c_inst"
        if c.startswith("panel") or c in ("dcs", "dcs_hmi", "weigh"):
            return "c_elec"
        return "c_mech"
    if c.startswith("inst"):
        return "inst"
    if c in ("mv", "mv_hp", "hv", "hv_hp", "hv_corr", "cv", "cv_hot", "cv_corr", "sv"):
        return "boiler_rep" if c == "sv" else "valve"
    if c == "bf" and kind == "R":
        return "bf_cloth"
    if c == "boiler" and kind == "O":
        return "boiler_rep"
    if c in ("sh", "sh2", "sh3", "eco") and kind == "O":
        return "boiler_rep"
    tt = FURN_TT[e["sysno"]](e)
    return tt or "fan"


def assign_contract(e, fy, kind, note):
    """(契約キー, 停止区分)。停止区分は '1'/'2'/'3'（号炉の定修）, 'Z'（全炉停止）, 'T'（タービン）, 'K'（基幹改良）"""
    c = e["cls"]
    if kind == "U":
        return ("KIKAN", "K")
    if "補修工事" in note:
        return (f"{fy}-HOSHU-{e['tag']}", e["unit"] if e["unit"] != "C" else "Z")
    if kind == "R" and c in BIG_R_CLS:
        return (f"{fy}-KOSHIN", e["unit"] if e["unit"] != "C" else "Z")
    if e["unit"] == "C":
        if e["tag"] in TURB_TAGS:
            if fy in TURB_YEARS:
                return (f"{fy}-TURB", "T")
            return (f"{fy}-ZENRO" if fy >= ZENRO_SPLIT_FY else f"{fy}-TEISEI", "Z")
        if c in ZENRO_CLS or (e["sysno"] == 4 and c in ("mv_hp", "hv_hp", "sv", "cv_hot", "header")) or c in ("dcs", "dcs_hmi") and kind == "I" and e["sysno"] == 11:
            return (f"{fy}-ZENRO" if fy >= ZENRO_SPLIT_FY else f"{fy}-TEISEI", "Z")
        if kind in ("O", "R") and c in LOCAL_CLS:
            return (f"{fy}-LOCAL", None)
        # 運転しながらできる共通設備：3 炉の定修のどれかに合わせる
        k = (e["sysno"] + e["loop"]) % 3 + 1
        return (f"{fy}-TEISEI", str(k))
    return (f"{fy}-TEISEI", e["unit"])


# ---------------------------------------------------------------- 停止期間（定修・全停・タービン・改良）
KIKAN_STOP = {2022: ("1", "2022-06-13", "2022-10-16"), 2023: ("2", "2023-06-12", "2023-10-15"),
              2024: ("3", "2024-06-10", "2024-10-13")}


def furnace_start(fy, u):
    if fy in KIKAN_STOP:
        ku, ks, ke = KIKAN_STOP[fy]
        if u == ku:
            return D(ks) + dt.timedelta(days=21)  # 改良停止期間の中で定修
        others = [x for x in "123" if x != ku]
        if u == others[0]:
            return nth_monday(fy, 11, 1)
        return nth_monday(fy + 1, 1, 3)
    if u == "1":
        return nth_monday(fy, 5, 2)
    if u == "2":
        return nth_monday(fy, 9, 1)
    return nth_monday(fy + 1, 1, 3)


def zenro_start(fy):
    return nth_monday(fy, 11, 2)


def turb_start(fy):
    return nth_monday(fy, 10, 4)


# ---------------------------------------------------------------- エピソード（日報・報告書・工程に跡が残る出来事）
EPISODES = [
    # (年度, 停止区分, 作業項目, 何日目, 記事, 遅れ日数, 報告書の所見)
    (2005, "2", "sb", 1, "スートブロワ L4 ランス 曲がり（初期不良、保証で交換手配）", 0, "2号炉 長抜差式スートブロワ L4 のランス曲がりは保証期間内につき無償交換とした。"),
    (2009, "3", "bf", 0, "降雪のため屋外作業 半日中止", 1, "降雪により 3号炉の工程を 1 日延長した。"),
    (2013, "1", "boiler_ut", 2, "1次過熱器 吊り金具 1 か所 割れ。溶接補修", 1, "1号炉 1次過熱器の吊り金具割れを溶接補修した（1 日延長）。"),
    (2016, "2", "scrubber", 1, "洗煙塔 ライニング ふくれ 2 か所、部分補修 約 0.6m2", 0, "2号炉 洗煙塔のライニングふくれ 2 か所を部分補修した。"),
    (2019, "3", "fan", 0, "誘引通風機 軸受振動 7.1mm/s を確認（運転中の傾向管理値超え）。次年度 軸受交換を提案", 0,
     "3号炉 誘引通風機の軸受振動が上昇しており、次年度の軸受交換を提案する。"),
    (2020, "1", "grate", 0, "緊急事態宣言に伴い協力会社の県外応援 2 名減。工程を 2 日延長", 2, "協力会社の人員減により 1号炉の工程を 2 日延長した。"),
    (2022, "2", "boiler_ut", 2, "2号炉 ボイラ左側壁 水冷壁管 4 本 管理値割れ（最小 3.4mm）。本委託範囲外のため市と協議（補修工事で対応）", 3,
     "2号炉 第2放射室 左側壁の水冷壁管 4 本が管理値未満。別途 補修工事で取替・肉盛を行い、立上げを 3 日繰下げた。"),
    (2022, "2", "rh", 0, "再加熱器 管漏れ 2 本（フィン付管 根元 腐食）。施栓（累計 5 本）", 1, "2号炉 再加熱器の管漏れ 2 本を施栓した（累計 5 本 / 96 本）。R8 の更新を推奨する。"),
    (2022, "2", "bf_cloth", 3, "ろ布ケージ 腐食 4 本、持参予備で交換", 1, "2号炉 集じん器のろ布ケージ 4 本を交換した（変更契約）。"),
    (2025, "Z", "z_dcs", 0, "DCS 共通部 切替（基幹改良と同時）。全炉停止を 12 日間とした", 0, "全炉停止を基幹改良（DCS 共通部切替）に合わせて 12 日間とした。"),
]


def episodes_for(fy, stop):
    return [x for x in EPISODES if x[0] == fy and x[1] == stop]


# ---------------------------------------------------------------- 本体
class World:
    def __init__(self):
        self.eq = EQ.build()
        self.by_tag = {e["tag"]: e for e in self.eq}
        self.contracts = {}
        self.windows = {}
        self.tasks = []
        self.works = []
        self.lines = []  # 日々の作業（日報の行のもと）
        self.booked = {"plan": {}, "act": {}}
        self._build()

    def _book(self, which, tk, start, n, sundays=False):
        bk = self.booked[which].setdefault(tk, set())
        out, d = [], start
        while len(out) < n:
            if (sundays or d.weekday() != 6) and d not in bk:
                out.append(d)
            d += dt.timedelta(days=1)
        bk.update(out)
        return out

    def _schedule(self, tasks, fy, stop, start, sundays=False, late_after=True):
        """予定と実績を班の空きで決める。late（復旧以降）は主作業の後。エピソードで実績が延びる"""
        eps = episodes_for(fy, stop)
        shift = 0
        for ep in eps:
            for t in tasks:
                if t["type"] == ep[2]:
                    t.setdefault("episodes", []).append(ep)
                    if ep[5]:
                        if ep[2] in ("rh", "bf_cloth"):
                            t["n_days"] += ep[5]
                        else:
                            shift += ep[5]
        late = {"restore", "hydro", "test", "startup"}
        for which, key, ndays in (("plan", "pdays", "plan_days"), ("act", "days", "n_days")):
            prev_end = None
            main_end = start
            for t in tasks:
                s0 = t["plan_start"]
                if t["type"] in late and late_after:
                    s0 = max(s0, (prev_end or main_end) + dt.timedelta(days=1))
                    if which == "act":
                        s0 = max(s0, prev_end + dt.timedelta(days=1)) if prev_end else s0
                sun = sundays or t["type"] == "startup"
                days = self._book(which, t["team"], s0, t[ndays], sundays=sun)
                t[key] = days
                if t["type"] in late:
                    prev_end = days[-1]
                else:
                    main_end = max(main_end, days[-1])
                    if t["type"] not in ("shutdown", "clean", "scaffold"):
                        prev_end = None
                if t["type"] == "restore" and which == "act" and shift:
                    pass
            if which == "act" and shift:
                pass
        # 実績の遅れ（補修工事などで復旧以降が繰り下がる）
        if shift:
            for t in tasks:
                if t["type"] in late:
                    bk = self.booked["act"][t["team"]]
                    for d in t["days"]:
                        bk.discard(d)
            last = None
            for t in tasks:
                if t["type"] in late:
                    s0 = t["days"][0] + dt.timedelta(days=shift)
                    if last:
                        s0 = max(s0, last + dt.timedelta(days=1))
                    t["days"] = self._book("act", t["team"], s0, t["n_days"], sundays=sundays or t["type"] == "startup")
                    last = t["days"][-1]
        for t in tasks:
            t["plan_start"], t["plan_end"] = t["pdays"][0], t["pdays"][-1]
            self._lines_for(t)

    # ---- 作業の元（年度×機器×種類）
    def _raw_events(self):
        out = []
        for e in self.eq:
            sim = W.simulate(e)
            for fy, acts in sim.items():
                if fy > 2026:
                    continue
                for (k, cost, note) in acts:
                    out.append(dict(tag=e["tag"], fy=fy, kind=k, cost=cost, note=note))
        return out

    def _contract(self, key, fy):
        if key in self.contracts:
            return self.contracts[key]
        kind = key.split("-")[1] if key != "KIKAN" else "KIKAN"
        names = {
            "TEISEI": f"{wa_fy(fy)} みなと臨海工場 定期整備業務委託",
            "ZENRO": f"{wa_fy(fy)} みなと臨海工場 全炉停止時共通設備整備業務委託",
            "TURB": f"{wa_fy(fy)} みなと臨海工場 蒸気タービン発電設備定期事業者検査業務委託",
            "KOSHIN": f"{wa_fy(fy)} みなと臨海工場 機器更新工事",
            "LOCAL": f"{wa_fy(fy)} みなと臨海工場 設備補修工事",
            "KIKAN": "みなと臨海工場 基幹的設備改良工事（令和4〜7年度）",
        }
        c = dict(id=key, fy=fy, kind=kind, name=names.get(kind, key), vendor="汐見重工環境", agent=agent(fy) if kind != "KIKAN" else "m4")
        if kind == "HOSHU":
            tag = key.split("-", 2)[2]
            e = self.by_tag[tag]
            unit = "共通" if e["unit"] == "C" else f"{e['unit']}号炉"
            c["name"] = f"{wa_fy(fy)} みなと臨海工場 {unit}{e['base_name']}補修工事".replace("ボイラ本体", "ボイラ水冷壁管").replace("共通ボイラ給水ポンプ No.2", "ボイラ給水ポンプNo.2")
        if kind == "LOCAL":
            c["vendor"] = "港南工業"
        self.contracts[key] = c
        return c

    def _build(self):
        ev = self._raw_events()
        # 年度×契約×停止区分 にまとめる
        groups = defaultdict(list)
        for x in ev:
            e = self.by_tag[x["tag"]]
            ck, stop = assign_contract(e, x["fy"], x["kind"], x["note"])
            x["contract"] = ck
            x["stop"] = stop
            groups[(ck, stop)].append(x)
        # 毎年の定期整備・全炉停止・タービンは作業の有無にかかわらず契約がある
        for fy in FY_RANGE:
            self._contract(f"{fy}-TEISEI", fy)
            if fy >= ZENRO_SPLIT_FY:
                self._contract(f"{fy}-ZENRO", fy)
            if fy in TURB_YEARS:
                self._contract(f"{fy}-TURB", fy)
        self._groups = groups
        for (ck, stop), xs in groups.items():
            fy = xs[0]["fy"] if ck != "KIKAN" else 2022
            self._contract(ck, fy)
        # 停止期間ごとに工程を作る
        for fy in FY_RANGE:
            for u in "123":
                xs = groups.get((f"{fy}-TEISEI", u), [])
                self._furnace_window(fy, u, xs)
            zx = groups.get((f"{fy}-ZENRO", "Z"), []) + groups.get((f"{fy}-TEISEI", "Z"), [])
            self._zenro_window(fy, zx)
            if fy in TURB_YEARS:
                self._turb_window(fy, groups.get((f"{fy}-TURB", "T"), []))
        for (ck, stop), xs in groups.items():
            kind = self.contracts[ck]["kind"]
            if kind in ("KOSHIN", "HOSHU"):
                self._small_window(ck, stop, xs)
            if kind == "LOCAL":
                self._local(ck, xs)
        self._kikan(groups.get(("KIKAN", "K"), []))
        # 2026 の状態
        for w in self.works:
            w["status"] = "done" if w["end"] and D(w["end"]) <= TODAY else ("in_progress" if w["start"] and D(w["start"]) <= TODAY else "planned")
        for i, w in enumerate(self.works):
            w["id"] = f"W{i + 1:05d}"
        for t in self.tasks:
            t["work_ids"] = [w["id"] for w in t["works"]]
        for ln in self.lines:
            ln["work_ids"] = [w["id"] for w in ln["works"]]

    # ---- 定修（号炉）
    def _furnace_window(self, fy, u, xs):
        key = f"{fy}-TEISEI"
        c = self.contracts[key]
        start = furnace_start(fy, u)
        wid = f"{fy}-{u}"
        r = rng("fw", fy, u)
        bytt = defaultdict(list)
        for x in xs:
            e = self.by_tag[x["tag"]]
            bytt[task_type(e, x["kind"], "TEISEI")].append(x)
        tasks = []
        no = 0
        def mk(tt, name, tk, s_off, n_days, works, plan_n=None):
            nonlocal no
            no += 1
            t = dict(id=f"T-{wid}-{no:02d}", contract=key, window=wid, no=f"{u}-{no:02d}", type=tt, name=name, team=tk,
                     plan_start=start + dt.timedelta(days=s_off), plan_days=plan_n or n_days, n_days=n_days, works=[], fy=fy)
            for x in works:
                t["works"].append(x)
            tasks.append(t)
            return t
        mk("shutdown", "立下げ立会・資機材搬入・養生", "機A", 0, 2, [])
        mk("clean", "炉内・ボイラ 灰落とし・清掃", "港南", 2, 3, [])
        mk("scaffold", "足場設置（炉内・ボイラ・集じん器）", "足場", 3, 3, [])
        order = ["boiler_ut", "bf_cloth", "grate", "furnace", "bf", "scrubber", "scr", "gct", "feed", "rh", "boiler_rep", "sb",
                 "valve", "fan", "inst", "elec", "ash", "c_crane", "c_mech", "c_inst", "c_elec"]
        team_free = defaultdict(lambda: 7)
        main_end = 10
        for tt in order:
            ws = bytt.get(tt, [])
            if not ws and tt not in ("grate", "furnace", "boiler_ut", "inst"):
                continue
            nm, tk, per, s0 = TASKTYPE[tt]
            n = max(1, -(-len(ws) // per)) if ws else 2
            if tt == "bf_cloth":
                n = 9
            if tt == "furnace":
                n = r.choice([4, 5, 5, 6])
            n = min(n, 11)
            s = max(s0, team_free[tk])
            t = mk(tt, nm, tk, s, n, ws)
            team_free[tk] = s + n + (1 if n > 3 else 0)
            main_end = max(main_end, s + n + n // 6 + 1)
        mk("restore", "足場解体・復旧・マンホール閉鎖", "足場", main_end, 3, [])
        mk("hydro", "ボイラ水圧試験", "機A", main_end + 4, 1, [])
        mk("test", "通風試験・各機器試運転", "機A", main_end + 5, 2, [])
        mk("startup", "立上げ立会・片付け", "機A", main_end + 7, 2, [])
        self._schedule(tasks, fy, u, start)
        self.windows[wid] = dict(id=wid, contract=key, fy=fy, unit=u, label=f"{u}号炉 定期整備", plan_start=start,
                                 tasks=tasks)
        self.tasks += tasks

    def _place(self, tasks, fy, stop, start):
        """予定 → 実績（エピソードによる遅れ）→ 日々の作業"""
        eps = episodes_for(fy, stop)
        delay_after = {}
        shift = 0
        for ep in eps:
            for t in tasks:
                if t["type"] == ep[2]:
                    t.setdefault("episodes", []).append(ep)
                    if ep[5]:
                        if ep[2] in ("rh", "bf_cloth"):
                            t["n_days"] += ep[5]  # その項目だけ延びる
                        else:
                            shift += ep[5]
        late_types = {"restore", "hydro", "test", "startup"}
        for t in tasks:
            pstart = t["plan_start"]
            pend_days = workdays_from(pstart, t["plan_days"], sundays=t["type"] in ("startup",))
            t["plan_start"], t["plan_end"] = pend_days[0], pend_days[-1]
            astart = pstart + dt.timedelta(days=(shift if t["type"] in late_types else 0))
            days = workdays_from(astart, t["n_days"], sundays=t["type"] in ("startup",))
            t["days"] = days
        for t in tasks:
            self._lines_for(t)

    def _lines_for(self, t):
        tm = team(t["team"], t["fy"])
        r = rng("ln", t["id"])
        n = len(t["days"])
        ws = t["works"]
        # 作業（機器）を日に振り分け
        per_day = defaultdict(list)
        for i, x in enumerate(ws):
            di = min(n - 1, int(i * n / max(1, len(ws))))
            per_day[di].append(x)
        work_days = defaultdict(list)
        for i, d in enumerate(t["days"]):
            persons = tm["size"] if t["type"] not in ("shutdown", "startup", "hydro", "test") else max(2, tm["size"] - r.choice([1, 2, 3]))
            if t["type"] in ("boiler_ut",):
                persons = 2
            persons = max(1, persons - (1 if r.random() < 0.2 else 0))
            hours = r.choice([8.0, 8.0, 8.0, 8.0, 9.0, 7.0]) if i not in (n - 1,) else r.choice([6.0, 7.0, 8.0])
            items = per_day.get(i, [])
            ln = dict(date=d, task=t, team=tm, persons=persons, hours=hours, items=items, works=[], day_index=i, n_days=n,
                      episode=None)
            for ep in t.get("episodes", []):
                if ep[3] == i:
                    ln["episode"] = ep
            self.lines.append(ln)
            for x in items:
                work_days[id(x)].append(ln)
        # 作業ごとの記録（開始・終了・工数）
        for x in ws:
            lns = work_days.get(id(x)) or [self.lines[-1]]
            e = self.by_tag[x["tag"]]
            nshare = sum(ln["persons"] * ln["hours"] / max(1, len(ln["items"])) for ln in lns)
            w = dict(tag=x["tag"], kind=x["kind"], fy=x["fy"], note=x["note"], cost=x["cost"], contract=t["contract"], task=t["id"],
                     start=lns[0]["date"].isoformat(), end=lns[-1]["date"].isoformat(), manHours=round(nshare, 1),
                     vendor=tm["co"].replace("【架空】", ""), team=tm["name"], unit=e["unit"])
            self.works.append(w)
            for ln in lns:
                ln["works"].append(w)
            t.setdefault("work_objs", []).append(w)
        t["works"] = t.get("work_objs", [])

    # ---- 全炉停止
    def _zenro_window(self, fy, xs):
        key = f"{fy}-ZENRO" if fy >= ZENRO_SPLIT_FY else f"{fy}-TEISEI"
        start = zenro_start(fy)
        wid = f"{fy}-Z"
        span = 12 if fy == 2025 else 7
        bytt = defaultdict(list)
        for x in xs:
            bytt[task_type(self.by_tag[x["tag"]], x["kind"], "ZENRO")].append(x)
        tasks = []
        no = 0
        def mk(tt, name, tk, s_off, n_days, works):
            nonlocal no
            no += 1
            t = dict(id=f"T-{wid}-{no:02d}", contract=key, window=wid, no=f"Z-{no:02d}", type=tt, name=name, team=tk,
                     plan_start=start + dt.timedelta(days=s_off), plan_days=n_days, n_days=n_days, works=list(works), fy=fy)
            tasks.append(t)
        mk("z_stop", "全炉停止 立会・隔離（主蒸気母管・給水）", "機C", 0, 1, [])
        for tt in ["z_steam", "z_dea", "z_hv", "z_stack", "z_acc", "z_turb", "z_inst", "z_dcs"]:
            ws = bytt.get(tt, [])
            if not ws and tt not in ("z_steam", "z_hv"):
                continue
            nm, tk, per, s0 = TASKTYPE[tt]
            n = max(1, min(span - 2, -(-len(ws) // per) if ws else 2))
            if tt == "z_dcs" and fy == 2025:
                n = 9
            mk(tt, nm, tk, s0, n, ws)
        mk("z_start", "復旧・各炉立上げ立会", "機C", span - 1, 1, [])
        for t in tasks:
            t["plan_days"] = t["n_days"]
        self._schedule(tasks, fy, "Z", start, sundays=True, late_after=False)
        self.windows[wid] = dict(id=wid, contract=key, fy=fy, unit="Z", label="全炉停止", plan_start=start, tasks=tasks)
        self.tasks += tasks

    # ---- タービン定期事業者検査
    def _turb_window(self, fy, xs):
        key = f"{fy}-TURB"
        start = turb_start(fy)
        wid = f"{fy}-T"
        seq = [("t_stop", "解列・停止、ターニング・冷却", "東雲", 0, 3), ("t_ins", "保温取外し・計器類取外し", "保温", 3, 2),
               ("t_open", "車室開放", "東雲", 7, 3), ("t_lift", "ロータ吊上げ", "東雲", 11, 1),
               ("t_main", "分解点検（ロータ・ノズル・ダイヤフラム・軸受）", "東雲", 12, 12), ("t_valve", "主塞止弁・蒸気加減弁・抽気逆止弁 分解点検", "機C", 14, 6),
               ("t_insp", "定期事業者検査（開放検査）", "東雲", 28, 1), ("t_asm", "組立（ロータ据付・車室閉鎖）", "東雲", 29, 10),
               ("t_ins2", "保温復旧・計器取付", "保温", 42, 4), ("t_run", "試運転（昇速・調速・非常停止試験）", "東雲", 49, 4),
               ("t_final", "総合検査・片付け", "東雲", 53, 1)]
        tasks = []
        valve_tags = {"C-MSV-504", "C-GV-505", "C-NRV-506", "C-NRV-507"}
        for i, (tt, nm, tk, off, n) in enumerate(seq, start=1):
            if tt == "t_main":
                ws = [x for x in xs if x["tag"] not in valve_tags]
            elif tt == "t_valve":
                ws = [x for x in xs if x["tag"] in valve_tags]
            else:
                ws = []
            t = dict(id=f"T-{wid}-{i:02d}", contract=key, window=wid, no=f"T-{i:02d}", type=tt, name=nm, team=tk,
                     plan_start=start + dt.timedelta(days=off), plan_days=n, n_days=n, works=ws, fy=fy)
            tasks.append(t)
        self._schedule(tasks, fy, "T", start, late_after=False)
        self.windows[wid] = dict(id=wid, contract=key, fy=fy, unit="T", label="蒸気タービン 定期事業者検査", plan_start=start, tasks=tasks)
        self.tasks += tasks

    # ---- 更新工事・補修工事（定修・全停の停止期間内）
    def _small_window(self, ck, stop, xs):
        c = self.contracts[ck]
        fy = c["fy"]
        if stop in ("1", "2", "3"):
            base = furnace_start(fy, stop)
        else:
            base = zenro_start(fy)
        if c["kind"] == "HOSHU" and fy == 2022:
            base = D("2022-11-07")
            seq_start = D("2022-11-28")
        else:
            seq_start = base + dt.timedelta(days=9)
        tag = xs[0]["tag"]
        e = self.by_tag[tag]
        if c["kind"] == "KOSHIN":
            allx = [x for (k2, s2), v in self._groups.items() if k2 == ck for x in v]
            names = []
            for x in sorted(allx, key=lambda x: -x["cost"]):
                nm = self.by_tag[x["tag"]]["base_name"].rstrip("AB")
                if nm not in names:
                    names.append(nm)
            c["name"] = f"{wa_fy(fy)} みなと臨海工場 " + "・".join(names[:2]) + ("ほか" if len(names) > 2 else "") + "更新工事"
        wid = f"{ck}-{stop}"
        tasks = []
        tk0 = "溶接" if c["kind"] == "HOSHU" else "据付"
        plan = [("k_prep", "準備・既設撤去", tk0, 0, 2),
                ("k_main", ("取替・肉盛溶接・非破壊検査" if c["kind"] == "HOSHU" else "新設据付・調整"), tk0, 2, 3 if c["kind"] == "HOSHU" else 6),
                ("k_fin", "検査・片付け", tk0, 5 if c["kind"] == "HOSHU" else 9, 1)]
        for i, (tt, nm, tk, off, n) in enumerate(plan, start=1):
            t = dict(id=f"T-{wid}-{i:02d}", contract=ck, window=wid, no=f"H-{i:02d}", type=tt, name=nm, team=tk,
                     plan_start=seq_start + dt.timedelta(days=off), plan_days=n, n_days=n, works=(xs if tt == "k_main" else []), fy=fy)
            tasks.append(t)
        self._schedule(tasks, fy, "H", seq_start, late_after=False)
        where = {"1": "1号炉", "2": "2号炉", "3": "3号炉", "Z": "共通"}.get(stop, stop)
        self.windows[wid] = dict(id=wid, contract=ck, fy=fy, unit=stop, label=f"{where} " + ("更新工事" if c["kind"] == "KOSHIN" else "補修工事"),
                                 plan_start=seq_start, tasks=tasks)
        self.tasks += tasks
        c["stop"] = stop

    # ---- 地元業者の補修（市が直接発注。工程表・日報・報告書の電子データはない）
    def _local(self, ck, xs):
        c = self.contracts[ck]
        r = rng("local", ck)
        for x in xs:
            e = self.by_tag[x["tag"]]
            m = r.choice([6, 7, 8, 10, 11])
            d0 = dt.date(c["fy"] if m >= 4 else c["fy"] + 1, m, r.randrange(1, 25))
            w = dict(tag=x["tag"], kind=x["kind"], fy=x["fy"], note=x["note"], cost=x["cost"], contract=ck, task=None,
                     start=d0.isoformat(), end=(d0 + dt.timedelta(days=r.choice([2, 3, 4]))).isoformat(), manHours=None,
                     vendor=r.choice(["港南工業", "臨海プラント工事"]), team=None, unit="C")
            self.works.append(w)

    # ---- 基幹改良
    def _kikan(self, xs):
        key = "KIKAN"
        c = self.contracts.setdefault(key, dict(id=key, fy=2022, kind="KIKAN", name="みなと臨海工場 基幹的設備改良工事（令和4〜7年度）",
                                               vendor="汐見重工環境", agent="m4"))
        byfy = defaultdict(list)
        for x in xs:
            byfy[x["fy"]].append(x)
        for fy in (2022, 2023, 2024, 2025):
            ws = byfy.get(fy, [])
            wid = f"K{fy}"
            if fy in KIKAN_STOP:
                u, s, e_ = KIKAN_STOP[fy]
                start = D(s)
                seq = [("u_stop", f"{u}号炉 立下げ・冷却・養生", "改良", 0, 4), ("u_scaf", "仮設足場・揚重設備 設置", "足場", 4, 8),
                       ("u_rm_sh", "過熱器（1〜3次）撤去", "据付", 12, 12), ("u_sh", "過熱器（1〜3次）据付・溶接", "改良", 24, 22),
                       ("u_ndt", "溶接部 非破壊検査・水圧試験", "検査", 46, 6), ("u_idf", "誘引通風機 撤去・据付（インバータ化）", "据付", 14, 20),
                       ("u_dcs", "DCS 炉別盤 撤去・据付・配線替え", "電工", 16, 30), ("u_loop", "計装ループ試験・インターロック試験", "電計", 50, 10),
                       ("u_rmv", "仮設撤去・復旧", "足場", 60, 8), ("u_trial", f"{u}号炉 試運転・性能確認", "改良", 98, 12)]
            else:
                start = zenro_start(fy)
                u = "C"
                seq = [("u_prep", "共通 DCS 仮設盤 設置・切替準備", "電工", -40, 20), ("u_dcs", "共通 DCS 切替（全炉停止中）", "電工", 0, 10),
                       ("u_gov", "タービン 調速装置・主塞止弁・加減弁 更新", "改良", 0, 10), ("u_loop", "ループ試験・総合試運転", "電計", 10, 8),
                       ("u_rmv", "旧盤 撤去・片付け", "電工", 20, 6)]
            tasks = []
            for i, (tt, nm, tk, off, n) in enumerate(seq, start=1):
                if tt in ("u_sh",):
                    w_ = [x for x in ws if self.by_tag[x["tag"]]["cls"] in ("sh", "sh2", "sh3")]
                elif tt == "u_idf":
                    w_ = [x for x in ws if self.by_tag[x["tag"]]["cls"] == "idf"]
                elif tt == "u_dcs":
                    w_ = [x for x in ws if self.by_tag[x["tag"]]["cls"] in ("dcs", "dcs_hmi", "panel")]
                elif tt == "u_gov":
                    w_ = [x for x in ws if self.by_tag[x["tag"]]["unit"] == "C" and self.by_tag[x["tag"]]["cls"] not in ("dcs", "dcs_hmi", "panel")]
                else:
                    w_ = []
                t = dict(id=f"T-{wid}-{i:02d}", contract=key, window=wid, no=f"K-{i:02d}", type=tt, name=nm, team=tk,
                         plan_start=start + dt.timedelta(days=off), plan_days=n, n_days=n, works=w_, fy=fy)
                tasks.append(t)
            self._schedule(tasks, fy, "K", start, late_after=False)
            self.windows[wid] = dict(id=wid, contract=key, fy=fy, unit=u, label=f"基幹改良 {wa_fy(fy)}（{u if u == 'C' else u + '号炉'}）",
                                     plan_start=start, tasks=tasks)
            self.tasks += tasks

    # ---- 便利
    def windows_of(self, fy, contract_kind=None):
        out = [w for w in self.windows.values() if w["fy"] == fy]
        if contract_kind:
            out = [w for w in out if self.contracts[w["contract"]]["kind"] == contract_kind]
        return out

    def window_dates(self, w):
        ds = [d for t in w["tasks"] for d in t["days"]]
        ps = [t["plan_start"] for t in w["tasks"]] + [t["plan_end"] for t in w["tasks"]]
        return min(ds), max(ds), min(ps), max(ps)


def wa_fy(fy):
    if fy >= 2019:
        n = fy - 2018
        return "令和元年度" if n == 1 else f"令和{n}年度"
    return f"平成{fy - 1988}年度"


_W = None


def get():
    global _W
    if _W is None:
        _W = World()
    return _W


if __name__ == "__main__":
    import time
    t0 = time.time()
    w = get()
    print("works", len(w.works), "tasks", len(w.tasks), "lines", len(w.lines), "contracts", len(w.contracts), round(time.time() - t0, 1))
    from collections import Counter
    print(Counter(c["kind"] for c in w.contracts.values()))
    for wid in ("2022-2", "2020-1", "2022-T", "2022-Z", "K2022"):
        win = w.windows[wid]
        print(wid, w.window_dates(win), len(win["tasks"]))
    print(Counter(x["status"] for x in w.works))
