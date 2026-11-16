# -*- coding: utf-8 -*-
"""2009–2026 年度の全契約と実際に行った作業（正解の世界）。
- SPC 定期整備工事：各年度 1号炉（春、2009 年度は秋）・2号炉（秋）。2025 年度は 1号のみ、2026 年度は 2号のみ（春）。
- DCS 更新工事（2019 年度）。
- 基幹的設備改良工事（組合発注）：R7 2号炉（施工済）、R8 1号炉（施工中、2026-11-13 時点）。
協力会社が施工しない小物（弁の一部・C 重要度の小物・タンク等）は SPC 保全班が直営で行い、協力会社の日報・報告書には載らない（world.json の undocumented）。
"""
import datetime as dt, random
import world as W, works as K
from works import Task

rnd = random.Random(W.SEED * 7)
D = dt.date

def nth_monday(y, m, n):
    d = D(y, m, 1)
    while d.weekday() != 0:
        d += dt.timedelta(days=1)
    return d + dt.timedelta(weeks=n - 1)

def kikan_fy(e):
    for e2, code, fy, cost, txt in W.kikan_items():
        if e2.tag == e.tag:
            return fy, code
    return None, None

REP_HIST = {r["tag"]: r for r in W.REP}
ATTR = {"I": "INSPECT", "O": "OVERHAUL", "P": "OVERHAUL", "R": "REPLACE", "U": "UPGRADE"}
HIST_CODE = {"INSPECT": "I", "OVERHAUL": "O", "REPLACE": "R", "UPGRADE": "U"}

def teishu_years(f):
    ys = [y for y in range(2009, 2027)]
    if f == 1:
        return [y for y in ys if y != 2026]
    return [y for y in ys if y != 2025]

def raw_code(e, y):
    """方針と代表機器の履歴から、その年度に行うべき作業（実施の前）"""
    if e.tag in REP_HIST:
        for yy, a, t in REP_HIST[e.tag]["history"]:
            if yy == y:
                return HIST_CODE[a]
        return None
    if e.cls in W.DCS_CLS:
        return "R" if y == 2019 else None
    m = W.policy_marks(e, y, y).get(y)
    kfy, kcode = kikan_fy(e)
    if kfy:
        if m == "R" and kfy - 3 <= y <= kfy + 14:
            m = None
        if m == "O" and kfy - 1 <= y <= kfy + 1:
            m = None
    return m

_ACT = {}
def actual_code(e, y):
    """実際に行った年度の作業。炉が基幹改良で止まる年度の SPC 分は翌年度の定修へ回す（1号の 2026 分は 2027 へ）"""
    key = (e.tag, y)
    if key in _ACT:
        return _ACT[key]
    c = raw_code(e, y)
    f = e.furn
    if e.tag in REP_HIST:
        _ACT[key] = c
        return c
    if f == 2 and y == 2025:
        c = None
    if f == 2 and y == 2026:
        c = c or raw_code(e, 2025)
    if f == 1 and y == 2026:
        c = None
    _ACT[key] = c
    return c

def actual_replacements(e, before):
    yrs = [y for y in range(2009, min(before, 2027)) if actual_code(e, y) == "R"]
    kfy, kcode = kikan_fy(e)
    if kfy and kfy < before and kfy <= 2025 and kcode == "R" or (kfy == 2025 and kcode == "U" and e.base in ("過熱器", "誘引通風機") and before > 2025):
        yrs.append(kfy)
    return sorted(set(yrs))

W.replaced_before = actual_replacements   # 台帳の更新年は実際の作業から

# ------------------------------------------------------------------ 作業の型（協力会社が施工するもの）
TEAM = {"gasifier": "炉内", "melter": "築炉", "chamber": "築炉", "boiler": "ボイラ", "sh": "ボイラ", "eco": "ボイラ", "sb": "ボイラ",
        "bf": "排ガス", "scr": "排ガス", "tower": "排ガス", "spray": "排ガス", "rh": "排ガス", "aph": "機械",
        "analyzer": "電気", "inv": "電気", "mcc": "電気", "hv": "電気", "lv": "電気", "tr_s": "電気", "ups": "電気", "dc": "電気", "relay": "電気",
        "pc": "電気", "disp": "電気", "itv": "電気", "scale": "電気"}
B24 = {"炉内": "A班", "築炉": "築炉", "ボイラ": "B班", "排ガス": "B班", "機械": "C班", "機械2": "C班", "電気": "C班", "仮設": "仮設"}

FIND = {
    "pump": ["インペラ摩耗軽微", "メカニカルシール摺動面に傷あり交換", "軸受すきま {a:.2f}mm（基準内）", "ケーシング内面に軽微な腐食、補修塗装", "ライナリング摩耗、交換"],
    "bfp": ["バランスディスク摩耗 {a:.2f}mm", "インペラ・ライナリング点検、良", "メカニカルシール交換"],
    "pump_d": ["ダイヤフラムに亀裂、交換", "弁座摩耗、交換"],
    "fan_l": ["羽根車ダスト付着、除去", "振動 {v:.1f}mm/s（整備後 {w:.1f}mm/s）", "軸受交換"],
    "fan_m": ["羽根車点検、良", "軸受グリス劣化、交換", "Vベルト交換"],
    "fan_s": ["Vベルト交換", "軸受交換"],
    "blower": ["ロータ・ギヤ点検、良", "オイル交換、Vベルト交換"],
    "nfscrew": ["スクリュー外径 φ{d:.0f}、羽根肉盛", "軸受・軸封交換", "水冷ジャケット 漏れなし"],
    "screw": ["スクリュー羽根摩耗 {t:.0f}mm、肉盛", "減速機オイル交換"],
    "feeder": ["スクリュー摩耗点検、良", "駆動チェーン伸び {p:.1f}%", "減速機点検、オイル交換"],
    "feeder_s": ["テーブル摩耗点検、良", "減速機オイル交換"],
    "chainconv": ["チェーン伸び {p:.1f}%", "スクレーパ摩耗、交換", "トラフ底板 減肉、当て板"],
    "conv": ["ベルト損傷なし", "ローラ交換 {n} 個", "スクレーパ交換"],
    "conv_d": ["スクリュー羽根摩耗、肉盛", "軸受交換"],
    "rv": ["ロータ・ケーシングすきま {a:.1f}mm、調整", "シール交換"],
    "elev": ["バケット交換 {n} 個", "チェーン伸び {p:.1f}%"],
    "shredder": ["カッタ刃摩耗、交換", "軸受交換", "スクリーン摩耗、補修"],
    "crane": ["ワイヤロープ素線切れなし", "バケット爪摩耗、肉盛", "走行車輪摩耗点検、良", "油圧シリンダ パッキン交換"],
    "burner": ["バーナチップ交換", "点火装置点検、良", "フレームアイ清掃"],
    "sb": ["ランス管交換", "グランドパッキン交換", "ノズル摩耗、交換"],
    "comp": ["エアエンド点検、良", "オイル・フィルタ交換", "アンローダ弁交換"],
    "kneader": ["パドル摩耗、肉盛", "軸封交換"],
    "spray": ["ノズル交換", "ノズル詰まり清掃"],
    "ro": ["RO 膜交換", "樹脂再生・補充"],
    "ct": ["充てん材清掃", "ファン軸受交換"],
    "mvalve": ["アクチュエータ点検、良", "グランドパッキン交換", "リミットスイッチ調整"],
    "cvalve": ["ポジショナ校正", "弁体・弁座点検、良", "ダイヤフラム交換"],
    "safety": ["吹出し試験 良（設定圧 調整なし）", "弁座すり合わせ"],
    "egen": ["燃料噴射弁点検", "冷却水・潤滑油交換", "負荷試験 良"],
    "turbine": ["動翼・静翼点検", "軸受点検", "ガバナ・保安装置試験"],
    "gen": ["固定子・回転子点検", "絶縁抵抗測定 良"],
    "gear": ["歯当たり確認 良", "軸受点検"],
    "acc": ["管束清掃", "ファン点検"],
    "hx": ["管束清掃", "ガスケット交換"],
    "aph": ["フィン清掃", "管漏れなし"],
    "magsep": ["マグネット点検、良"], "ecs": ["ロータ点検、ベルト交換"], "sieve": ["網交換", "防振ゴム交換"],
    "elev_": [],
}
DUR = {"pump": 2, "bfp": 3, "pump_d": 1, "fan_l": 3, "fan_m": 2, "fan_s": 1, "blower": 2, "nfscrew": 4, "screw": 2, "feeder": 2,
       "feeder_s": 1, "chainconv": 3, "conv": 2, "conv_d": 2, "rv": 1, "elev": 3, "shredder": 4, "crane": 3, "burner": 1, "sb": 2,
       "comp": 2, "kneader": 2, "spray": 2, "ro": 2, "ct": 2, "egen": 2, "turbine": 10, "gen": 4, "gear": 3, "acc": 3, "hx": 2,
       "aph": 2, "magsep": 1, "ecs": 1, "sieve": 1}
CREW = {"pump": {"機": 2}, "bfp": {"機": 3}, "pump_d": {"機": 1}, "fan_l": {"機": 3}, "fan_m": {"機": 2}, "fan_s": {"機": 2},
        "blower": {"機": 2}, "nfscrew": {"機": 3}, "screw": {"機": 2}, "feeder": {"機": 2}, "feeder_s": {"機": 1}, "chainconv": {"機": 2, "作": 1},
        "conv": {"機": 2}, "conv_d": {"機": 2}, "rv": {"機": 1}, "elev": {"機": 2, "作": 1}, "shredder": {"機": 3}, "crane": {"機": 2, "電": 1},
        "burner": {"機": 1}, "sb": {"機": 2}, "comp": {"機": 2}, "kneader": {"機": 2}, "spray": {"機": 2}, "ro": {"機": 2}, "ct": {"機": 2},
        "egen": {"機": 2, "電": 1}, "turbine": {"機": 5, "作": 2}, "gen": {"電": 2, "機": 2}, "gear": {"機": 3}, "acc": {"機": 2, "作": 2},
        "hx": {"機": 2}, "aph": {"機": 2}, "magsep": {"機": 1}, "ecs": {"機": 1}, "sieve": {"機": 2}}
WORKABLE = set(DUR)
VALVES = ("mvalve", "cvalve", "safety")

def fmt_find(cls, r):
    pool = FIND.get(cls) or ["点検、良"]
    s = r.choice(pool)
    return s.format(a=r.uniform(0.08, 0.30), v=r.uniform(3.0, 5.5), w=r.uniform(1.5, 2.5), d=r.uniform(330, 342), t=r.uniform(3, 9),
                    p=r.uniform(0.6, 1.7), n=r.randint(2, 8))

# ------------------------------------------------------------------ 状態の推移（測定値）
def melter_thk(f, y):
    if f == 2:
        return round(230 - (y - 2009) * (135 / 14))
    return round(230 - (y - 2009) * (112 / 16))

def sh_min(f, y):
    if f == 2:
        tab = {2009: 4.0, 2011: 3.7, 2013: 3.2, 2015: 2.6, 2017: 3.3, 2019: 2.9, 2021: 2.6, 2023: 2.3}
    else:
        tab = {2009: 4.0, 2011: 3.8, 2013: 3.4, 2015: 3.3, 2017: 3.1, 2019: 2.9, 2021: 2.7, 2023: 2.6, 2025: 2.5}
    ks = sorted(tab)
    v = tab[max(k for k in ks if k <= y)]
    return v

def bf_dp(f, y, last):
    return round(0.9 + (y - last) * 0.15, 1)

# ------------------------------------------------------------------ 契約
class Contract:
    def __init__(self, cid, fy, furn, kind, title, short, orderer, contractor, start, end, tasks, milestones=None, plan_end=None):
        self.cid, self.fy, self.furn, self.kind, self.title, self.short = cid, fy, furn, kind, title, short
        self.orderer, self.contractor, self.start, self.end = orderer, contractor, start, end
        self.tasks, self.milestones, self.plan_end = tasks, milestones or [], plan_end
        self.direct = []      # SPC 直営（どの資料にも書かない）
        self.parts = []       # (task tid, 品名, 規格, 手配数, 使用数, 単位, 単価, 取引先)
        self.status = "完了"

def fy_label(fy, wareki=False):
    if wareki or fy <= 2018:
        return f"平成{fy - 1988}年度" if fy <= 2018 else f"令和{fy - 2018}年度"
    return f"{fy}年度"

def working_days(a, b, sundays_off=True):
    return [d for d in W.daterange(a, b) if not (sundays_off and d.weekday() == 6)]

def make_teishu(f, y):
    if (f, y) == (2, 2015):
        c = Contract("K1", 2015, 2, "teishu", W.CONTRACTS["K1"]["title"], "2号炉定期整備工事", W.SPC, W.SUB1,
                     D(2015, 10, 5), D(2015, 10, 30), K.K1, K.K1_MILESTONES)
        c.cid = "T2015-2"
        add_direct(c, [t for tk in K.K1 for t in tk.tags])
        return c
    if (f, y) == (1, 2025):
        c = Contract("T2025-1", 2025, 1, "teishu", W.CONTRACTS["K2"]["title"], "1号炉定期整備工事", W.SPC, W.SUB1,
                     D(2025, 5, 12), D(2025, 6, 6), K.K2, K.K2_MILESTONES)
        add_direct(c, [t for tk in K.K2 for t in tk.tags])
        return c
    r = random.Random(W.SEED * 100 + y * 10 + f)
    if y == 2009:
        start = nth_monday(2009, 9, 2) if f == 1 else nth_monday(2009, 10, 3)
    elif f == 1:
        start = nth_monday(y, 5, 2)
    elif y == 2026:
        start = nth_monday(2026, 5, 2)
    else:
        start = nth_monday(y, 10, 1)
    end = start + dt.timedelta(days=25)
    wd = working_days(start + dt.timedelta(days=3), start + dt.timedelta(days=19))
    title = (f"{fy_label(y)} ときわ北部クリーンセンター {f}号炉定期整備工事")
    tasks = []
    tasks.append(Task("T00", "炉立下げ・冷却", [], "-", (start, start + dt.timedelta(days=2)), "-", "SPC運転", [], owner="SPC"))
    tasks.append(Task("T01", "仮設足場・養生", [], "-", (wd[0], wd[-2]), "仮設", W.SUB1, [
        (wd[0], r.choice(["ガス化炉・溶融炉まわり 足場組立、床養生", "炉室 足場組立・養生"]), {"と": 4, "作": r.randint(1, 2)}),
        (wd[-3], "足場解体（炉まわり）", {"と": 4}), (wd[-2], "足場解体・養生撤去・片付け", {"と": 3, "作": 2})]))
    team_day = {}
    def sched(team, ndays, first=1):
        i = max(team_day.get(team, first), first)
        days = wd[i:i + ndays]
        if len(days) < ndays:
            days = wd[-ndays - 2:-2]
        team_day[team] = i + ndays
        return days
    # 炉内（ガス化炉）
    g = W.find(f, "ガス化炉本体"); gc = actual_code(g, y)
    days = sched("炉内", 5 if gc == "O" else 3, 0)
    nwear = r.randint(14, 30) if gc == "O" else r.randint(0, 9)
    steps = [(days[0], "マンホール開放、流動砂抜出し（約 6m³）", {"機": 2, "作": 3}),
             (days[1], f"炉内清掃、散気ノズル摩耗測定（96 本中 {nwear} 本 交換基準超え）" if nwear else "炉内清掃、散気ノズル摩耗測定（交換基準超えなし）", {"機": 2, "作": 1})]
    if gc == "O":
        steps += [(days[2], f"散気ノズル交換 {nwear // 2} 本", {"機": 2, "溶": 1}), (days[3], f"散気ノズル交換 {nwear - nwear // 2} 本（計 {nwear} 本）、取付部溶接", {"機": 2, "溶": 1}),
                  (days[4], "炉内耐火物点検、異常なし", {"機": 1})]
    else:
        steps += [(days[2], r.choice(["炉内耐火物点検、異常なし", "炉内耐火物点検、目地の一部に軽微な欠け（経過観察）"]), {"機": 1})]
    steps.append((wd[-4], "マンホール閉止、珪砂投入（約 6m³）", {"機": 2, "作": 3}))
    tasks.append(Task("T02", "ガス化炉 開放点検" + ("・散気ノズル交換" if gc == "O" else ""), [g.tag], "OVERHAUL" if gc == "O" else "INSPECT",
                      (steps[0][0], steps[-1][0]), "炉内", W.SUB1, steps, note=f"散気ノズル {nwear} 本" if gc == "O" else ""))
    if gc == "O":
        c_parts = [("T02", "ガス化炉散気ノズル", "TH-GSN-160 SCS13", nwear + 4, nwear, "本", 38000 + (y - 2009) * 600, "【架空】鷹羽エンジニアリング")]
    else:
        c_parts = []
    # 溶融炉
    m = W.find(f, "旋回溶融炉本体"); mc = actual_code(m, y)
    thk = melter_thk(f, y)
    hist_txt = None
    for yy, a, t in REP_HIST[m.tag]["history"]:
        if yy == y: hist_txt = t
    if mc == "O":
        days = sched("築炉", 8, 1)
        area = r.choice([1.2, 1.6, 2.4, 3.0])
        steps = [(days[0], f"炉内温度確認、内部点検、残存厚測定（側壁下段 最小 {thk}mm）、補修範囲マーキング", {"築": 2}),
                 (days[1], "損傷部 耐火物はつり", {"築": 3, "作": 2}), (days[2], "はつり完了、鉄皮点検（異常なし）", {"築": 3, "作": 2}),
                 (days[3], "スタッド補修", {"築": 3, "溶": 1, "作": 1}), (days[4], "れんが積み", {"築": 4, "作": 2}),
                 (days[5], "れんが積み・目地詰め", {"築": 4, "作": 2}), (days[6], "キャスタブル施工、養生", {"築": 3, "作": 1}),
                 (days[7], "仕上がり確認（SPC 立会）、片付け", {"築": 2, "作": 1})]
        tasks.append(Task("T03", "旋回溶融炉 耐火物部分補修", [m.tag], "OVERHAUL", (days[0], days[-1]), "築炉", W.REFR, steps,
                          note=(hist_txt or "") + f" 補修面積 約 {area}m²"))
        c_parts.append(("T03", "溶融炉 耐火物部分補修工事", f"{area}m2 材料込", 1, 1, "式", int(2200000 + area * 1800000), "【架空】丸越築炉"))
    else:
        days = sched("築炉", 2, 2)
        steps = [(days[0], f"炉内点検、残存厚測定（側壁下段 最小 {thk}mm）", {"築": 2}), (days[1], "出滓口まわり 目視点検、目地補修", {"築": 2})]
        tasks.append(Task("T03", "旋回溶融炉 点検・残存厚測定", [m.tag], "INSPECT", (days[0], days[-1]), "築炉", W.REFR, steps))
        c_parts.append(("T03", "溶融炉 点検・目地補修", "材料込", 1, 1, "式", 380000, "【架空】丸越築炉"))
    # ボイラ
    b = [W.find(f, "ボイラ本体"), W.find(f, "過熱器"), W.find(f, "エコノマイザ")]
    insp = y % 2 == 1
    days = sched("ボイラ", 6, 0)
    smin = sh_min(f, y)
    steps = [(days[0], "マンホール・点検口開放", {"機": 2, "作": 1}), (days[1], "水管・過熱器 ダスト除去", {"機": 2, "作": 3}),
             (days[2], "エコノマイザ ダスト除去、ドラム内部清掃", {"機": 2, "作": 2}),
             (days[3], f"肉厚測定（水管・過熱器・エコノマイザ）、過熱器 最小 {smin}mm", {"機": 2})]
    if insp:
        steps.append((days[4], "定期事業者検査 内部検査（SPC 立会）", {"機": 2}))
    steps += [(wd[-6], "マンホール・点検口閉止", {"機": 2, "作": 1}), (wd[-5], "水張り・水圧試験（漏れなし）", {"機": 2})]
    tasks.append(Task("T04", "ボイラ 開放清掃・肉厚測定" + ("・定期事業者検査" if insp else ""), [x.tag for x in b], "INSPECT",
                      (days[0], wd[-5]), "ボイラ", W.SUB1, steps, note=f"過熱器 最小 {smin}mm"))
    # BF
    bf = W.find(f, "バグフィルタ本体"); bc = actual_code(bf, y)
    if bc == "O":
        days = sched("排ガス", 7, 1)
        steps = [(days[k], f"第{k + 1}室 旧ろ布抜取・新ろ布取付（42 本）", {"機": 2, "作": 4}) for k in range(6)]
        steps.append((days[6], "全室 点検口閉止、パルス弁作動確認", {"機": 2, "作": 1}))
        tasks.append(Task("T05", "バグフィルタ ろ布全数交換", [bf.tag], "OVERHAUL", (days[0], days[-1]), "排ガス", W.SUB1, steps))
        c_parts.append(("T05", "ろ布", "PTFE φ160×6000", 264, 252, "本", 24000 + (y - 2009) * 160, "【架空】大鳥フィルター"))
    else:
        days = sched("排ガス", 2, 1)
        steps = [(days[0], "各室点検口開放、ろ布目視点検（6 室）", {"機": 2, "作": 1}), (days[1], r.choice(["パルス弁作動確認、点検口閉止", "ろ布抜取（試験用 3 本）、予備ろ布取付、閉止"]), {"機": 2})]
        tasks.append(Task("T05", "バグフィルタ 点検", [bf.tag], "INSPECT", (days[0], days[-1]), "排ガス", W.SUB1, steps))
    # 触媒・減温塔（隔年）
    if y % 2 == f % 2:
        tw = W.find(f, "減温塔噴霧ノズル装置")
        days = sched("排ガス", 2)
        steps = [(days[0], "減温塔 噴霧ノズル 6 本取外し、摩耗点検", {"機": 2}), (days[1], "噴霧ノズル交換・取付", {"機": 2})]
        tasks.append(Task("T06", "減温塔 噴霧ノズル交換", [tw.tag], "OVERHAUL", (days[0], days[-1]), "排ガス", W.SUB1, steps))
        c_parts.append(("T06", "減温塔噴霧ノズル", "二流体 TH-SPN-25", 8, 6, "個", 64000, "【架空】鷹羽エンジニアリング"))
    else:
        sc = W.find(f, "触媒反応塔")
        days = sched("排ガス", 1)
        tasks.append(Task("T06", "触媒反応塔 点検", [sc.tag], "INSPECT", (days[0], days[0]), "排ガス", W.SUB1,
                          [(days[0], "触媒反応塔 開放、触媒サンプル採取、閉止", {"機": 2})]))
        c_parts.append(("T06", "脱硝触媒 性能試験", "サンプル1層", 1, 1, "式", 480000, "【架空】環境分析センター"))
    # 方針・履歴による整備（回転機ほか）
    pool = [e for e in W.EQS if (e.furn == f or (e.furn == 0 and f == (1 if y not in (2026,) else 2)))]
    items = []
    valves = []
    direct = []
    core = {g.tag, m.tag, bf.tag} | {x.tag for x in b}
    for e in pool:
        if e.tag in core or e.base in ("減温塔噴霧ノズル装置", "触媒反応塔"):
            continue
        cd = actual_code(e, y)
        if cd not in ("O", "R", "I", "P"):
            continue
        if e.cls in VALVES and cd in ("O", "R"):
            valves.append((e, cd)); continue
        if e.cls in WORKABLE and (e.crit in ("A", "B") or e.tag in REP_HIST) and (e.furn != 0 or e.crit == "A" or e.tag in REP_HIST):
            items.append((e, cd))
        elif cd != "I":
            direct.append((e, cd))
    # 協力会社の範囲は 12 件まで（代表機器→重要度 A→更新→B の順）。残りは SPC 保全班の直営
    items.sort(key=lambda x: (0 if x[0].tag in REP_HIST else 1, {"A": 0, "B": 1, "C": 2}[x[0].crit], 0 if x[1] == "R" else 1, x[0].idx))
    for e, cd in items[12:]:
        if cd != "I":
            direct.append((e, cd))
    items = items[:12]
    vmax = 9
    for e, cd in valves[vmax:]:
        direct.append((e, cd))
    valves = valves[:vmax]
    tno = 7
    for e, cd in sorted(items, key=lambda x: (TEAM.get(x[0].cls, "機械"), x[0].idx)):
        team = TEAM.get(e.cls, "機械" if tno % 2 else "機械2")
        hist = None
        if e.tag in REP_HIST:
            for yy, a, t in REP_HIST[e.tag]["history"]:
                if yy == y: hist = t
        n = DUR.get(e.cls, 2) + (1 if cd == "R" else 0)
        days = sched(team, n, 1)
        crew = CREW.get(e.cls, {"機": 2})
        lab = ("共通 " if e.furn == 0 else "") + e.name
        if cd == "R":
            steps = [(days[0], f"{e.name} 既設撤去", dict(crew))]
            for d in days[1:-1]:
                steps.append((d, f"{e.name} 新品据付・配管／配線接続", dict(crew)))
            steps.append((days[-1], f"{e.name} 芯出し・単体試運転", dict(crew)))
            name = f"{lab} 更新"; attr = "REPLACE"
        elif cd == "I":
            steps = [(days[0], f"{e.name} 点検（{hist or fmt_find(e.cls, r)}）", dict(crew))]
            name = f"{lab} 点検"; attr = "INSPECT"
        else:
            f1 = hist or fmt_find(e.cls, r)
            steps = [(days[0], f"{e.name} 取外し・分解", dict(crew))]
            for d in days[1:-1]:
                steps.append((d, f"{e.name} {r.choice(['部品点検・計測', '摩耗部品交換', '清掃・部品交換'])}（{f1}）" if d == days[1] else f"{e.name} 部品交換・組立", dict(crew)))
            steps.append((days[-1], f"{e.name} 組立・芯出し" if len(days) > 1 else f"{e.name} 分解点検・組立（{f1}）", dict(crew)))
            name = f"{lab} 分解整備"; attr = "OVERHAUL"
        tasks.append(Task(f"T{tno:02d}", name, [e.tag], attr, (days[0], days[-1]), team, W.SUB1, steps, note=hist or ""))
        c_parts += parts_for(f"T{tno:02d}", e, cd, y, r)
        tno += 1
    if valves:
        tags = [e.tag for e, _ in valves]
        n = max(2, min(5, (len(valves) + 2) // 3))
        days = sched("機械2", n, 2)
        names = [e.name for e, _ in valves]
        steps = []
        per = (len(names) + n - 1) // n
        for i, d in enumerate(days):
            seg = names[i * per:(i + 1) * per]
            if seg:
                steps.append((d, f"電動弁・調節弁 整備（{'、'.join(seg)}）", {"機": 2, "電": 1}))
        tasks.append(Task(f"T{tno:02d}", f"電動弁・調節弁 整備（{len(valves)} 台）", tags, "OVERHAUL", (steps[0][0], steps[-1][0]), "機械2", W.SUB1, steps))
        c_parts.append((f"T{tno:02d}", "弁用パッキン類", "一式", 1, 1, "式", 180000 + 12000 * len(valves), "【架空】鷹羽エンジニアリング"))
        tno += 1
    # 追加補修（ときどき起きる）：整備中の発見で 1 日延びる
    cand = [t for t in tasks if t.tid >= "T07" and t.tags and t.attr == "OVERHAUL" and t.team in ("機械", "機械2")]
    if cand and r.random() < 0.45:
        t = r.choice(cand)
        last = t.steps[-1][0]
        nd = [d for d in wd if d > last]
        if nd:
            add = r.choice(["軸受箱に割れ、補修溶接", "ケーシング当たり面 摩耗、肉盛補修", "据付ボルト 腐食、交換", "軸に摩耗、スリーブ取付"])
            t.steps.append((nd[0], f"【追加】{W.BY_TAG[t.tags[0]].name} {add}", {"機": 2, "溶": 1}))
            t.actual = (t.steps[0][0], nd[0]); t.note = (t.note + " " if t.note else "") + "追加補修 " + add
    tasks.append(Task(f"T{tno:02d}", "復旧・単体試運転", [], "-", (wd[-4], wd[-1]), "機械", W.SUB1, [
        (wd[-3], "単体試運転（整備機器）", {"機": 2, "電": 1}), (wd[-1], r.choice(["点検口閉止確認、片付け", "残材搬出・清掃"]), {"機": 1, "作": 2})]))
    tno += 1
    su = start + dt.timedelta(days=21)
    tasks.append(Task(f"T{tno:02d}", "炉立上げ・試運転", [], "-", (su, end), "-", "SPC運転", [
        (su + dt.timedelta(days=1), "立上げ立会（漏れ・異音確認）", {"機": 1}), (end, "定格運転確認", {"機": 1})], owner="SPC"))
    ms = [(start, "▼炉停止"), (su, "▼点火"), (end, "▼定格")]
    c = Contract(f"T{y}-{f}", y, f, "teishu", title, f"{f}号炉定期整備工事", W.SPC, W.SUB1, start, end, tasks, ms)
    c.parts = c_parts
    c.direct = [(e.tag, ATTR[cd]) for e, cd in direct]
    return c

PARTS = {
    "pump": [("メカニカルシール", "", 1, 68000), ("軸受", "6310", 2, 9800)], "bfp": [("メカニカルシール", "", 2, 210000), ("ライナリング", "", 2, 88000)],
    "pump_d": [("ダイヤフラム・弁座キット", "", 2, 46000)], "fan_l": [("軸受", "22324", 2, 148000)], "fan_m": [("軸受", "22218", 2, 64000)],
    "fan_s": [("Vベルト", "B-60", 3, 2400)], "blower": [("Vベルト", "B-70", 3, 2600), ("ギヤオイル", "20L", 1, 18000)],
    "nfscrew": [("軸受", "22220 CAKE4", 2, 86000), ("軸封パッキン", "グランド 一式", 1, 64000)], "screw": [("軸受", "22216", 2, 42000)],
    "feeder": [("駆動チェーン", "RS100 142L", 1, 62000)], "chainconv": [("スクレーパ", "W600 耐摩耗鋼", 10, 27000)],
    "conv": [("キャリアローラ", "φ89", 6, 6800)], "conv_d": [("軸受", "6312", 2, 12000)], "rv": [("シール", "", 2, 15000)],
    "elev": [("バケット", "", 8, 14000)], "shredder": [("破砕機カッタ刃", "t=50", 40, 82000), ("軸受", "23228", 4, 120000)],
    "crane": [("ワイヤロープ", "IWRC 6×Fi(29) φ22", 2, 380000)], "burner": [("バーナチップ", "", 2, 46000)],
    "sb": [("ランス管", "STBA24 φ60.5", 1, 412000), ("グランドパッキン", "", 2, 18500)], "comp": [("オイル・フィルタ セット", "", 1, 56000)],
    "kneader": [("パドル", "", 8, 32000)], "spray": [("噴霧ノズル", "", 4, 52000)], "ro": [("RO 膜エレメント", "", 4, 118000)],
    "ct": [("充てん材", "", 1, 240000)], "egen": [("燃料噴射ノズル", "", 6, 36000)], "turbine": [("軸受メタル", "", 2, 860000), ("パッキン類", "一式", 1, 420000)],
    "gen": [("ブラシ・保護具", "", 1, 80000)], "gear": [("潤滑油", "200L", 1, 140000)], "acc": [("ファンベルト", "", 6, 9800)],
    "hx": [("ガスケット", "", 2, 38000)], "aph": [], "magsep": [], "ecs": [("ベルト", "W600", 1, 98000)], "sieve": [("網", "目開き 3mm", 2, 64000)],
}
def parts_for(tid, e, cd, y, r):
    out = []
    if cd == "R":
        price = W.COST_R.get(e.cls, W.COST_O.get(e.cls, 300) * 6) * 1000 * 0.7
        out.append((tid, e.name + " 本体", e.spec[:24], 1, 1, "台", int(round(price, -3)), W.BY_TAG[e.tag].maker))
        return out
    if cd == "I":
        return out
    for nm, spec, q, p in PARTS.get(e.cls, []):
        spare = 1 if q <= 2 and r.random() < 0.3 else (2 if q >= 6 and r.random() < 0.5 else 0)
        out.append((tid, f"{e.base} {nm}" if nm not in e.base else nm, spec, q + spare, q, "個" if q < 6 else "枚" if "刃" in nm or "スクレーパ" in nm else "本",
                    int(p * (1 + (y - 2009) * 0.012)), W.BY_TAG[e.tag].maker))
    return out

def add_direct(c, contractor_tags):
    f, y = c.furn, c.fy
    s = set(contractor_tags)
    pool = [e for e in W.EQS if e.furn == f or (e.furn == 0 and f == 1)]
    for e in pool:
        if e.tag in s: continue
        cd = actual_code(e, y)
        if cd in ("O", "R", "P"):
            c.direct.append((e.tag, ATTR[cd]))

# ------------------------------------------------------------------ DCS 更新（2019 年度）
def make_dcs():
    s = D(2020, 1, 14); e_ = D(2020, 2, 21)
    wd = working_days(s, e_)
    tags_c = [e.tag for e in W.EQS if e.cls in W.DCS_CLS and e.furn == 0]
    tags_1 = [e.tag for e in W.EQS if e.cls in W.DCS_CLS and e.furn == 1]
    tags_2 = [e.tag for e in W.EQS if e.cls in W.DCS_CLS and e.furn == 2]
    def st(days, texts, crew):
        return [(d, t, dict(crew)) for d, t in zip(days, texts)]
    tasks = [
        Task("T01", "中央制御室 新オペレータ卓 据付", [t for t in tags_c if W.BY_TAG[t].cls == "dcsop"], "REPLACE", (wd[0], wd[4]), "電気", W.SUB1,
             st(wd[0:5], ["搬入・仮置き、既設卓の横に据付位置墨出し", "オペレータステーション No.1・No.2 据付", "No.3・No.4・ES 据付", "電源・LAN 配線", "画面確認（メーカー SV 立会）"], {"電": 3, "作": 1})),
        Task("T02", "共通コントローラ盤・I/O盤 更新", [t for t in tags_c if W.BY_TAG[t].cls != "dcsop"], "REPLACE", (wd[3], wd[10]), "電気", W.SUB1,
             st(wd[3:11], ["新盤 搬入", "新盤 据付・アンカー", "ケーブル 移設準備（ラベル付け）", "ケーブル 移設（共通系 1/2）", "ケーブル 移設（共通系 2/2）", "ループチェック", "ループチェック（残）", "旧盤 撤去"], {"電": 4})),
        Task("T03", "1号炉 炉コントローラ盤・I/O盤 更新（1号炉停止中）", tags_1, "REPLACE", (wd[11], wd[18]), "電気", W.SUB1,
             st(wd[11:19], ["1号炉 停止確認、旧盤 電源切離し", "旧盤 撤去・新盤 据付", "ケーブル 移設", "ケーブル 移設（残）", "ループチェック", "インターロック試験", "炉立上げ前 総合確認", "切替後 監視"], {"電": 4, "作": 1})),
        Task("T04", "2号炉 炉コントローラ盤・I/O盤 更新（2号炉停止中）", tags_2, "REPLACE", (wd[19], wd[26]), "電気", W.SUB1,
             st(wd[19:27], ["2号炉 停止確認、旧盤 電源切離し", "旧盤 撤去・新盤 据付", "ケーブル 移設", "ケーブル 移設（残）", "ループチェック", "インターロック試験", "炉立上げ前 総合確認", "切替後 監視"], {"電": 4, "作": 1})),
        Task("T05", "総合試験・旧設備撤去", [], "-", (wd[27], wd[-1]), "電気", W.SUB1,
             st(wd[27:], ["両炉 運転データ比較", "旧オペレータ卓 撤去", "撤去品 搬出", "竣工図書 確認", "引渡し"], {"電": 2, "作": 2})),
    ]
    c = Contract("D2019", 2019, 0, "dcs", "2019年度 ときわ北部クリーンセンター DCS更新工事", "DCS更新工事", W.SPC, W.MAKER, s, e_, tasks,
                 [(wd[11], "▼1号停止"), (wd[19], "▼2号停止"), (e_, "▼引渡")])
    c.parts = [("T01", "DCS 機器一式（オペレータステーション・コントローラ・I/O）", "【架空】青柳計測 TH-DCS-2", 1, 1, "式", 168000000, "【架空】鷹羽エンジニアリング"),
               ("T02", "DCS 据付・配線工事", "", 1, 1, "式", 24800000, "【架空】北嶺プラント工業"),
               ("T01", "DCS ソフト移行・試験", "", 1, 1, "式", 38500000, "【架空】鷹羽エンジニアリング")]
    return c

# ------------------------------------------------------------------ 基幹改良
def make_kikan(fy):
    if fy == 2025:
        rows = K.K3
        c = Contract("KK2025", 2025, 2, "kikan", "ときわ北部クリーンセンター基幹的設備改良工事（令和7〜9年度）", "基幹的設備改良工事（R7 2号炉）",
                     W.UNION, W.MAKER, D(2025, 8, 18), D(2026, 1, 16), rows, K.K3_EVENTS, plan_end=D(2025, 12, 19))
        return c
    shift = D(2026, 9, 7) - D(2025, 9, 1)
    rows = []
    for kou, name, tags, attr, ps, pe, as_, ae, crew in K.K3:
        tags1 = [t.replace("2-", "1-", 1) for t in tags]
        ps1, pe1 = ps + shift, pe + shift
        if name in ("鉄皮補修（スタッド・当て板）",):
            a1, b1 = ps1, pe1 + dt.timedelta(days=2)     # 事前調査どおり。当て板は 2 日延長のみ
        else:
            a1, b1 = ps1, pe1
        if name == "旋回溶融炉 耐火物築炉":
            a1 = ps1 + dt.timedelta(days=2); b1 = pe1 + dt.timedelta(days=2)
        rows.append((kou, name, tags1, attr, ps1, pe1, a1, b1, crew))
    ev = [(D(2026, 9, 7), "◆1号炉停止"), (D(2026, 9, 30), "▲鉄皮 事前調査どおり（当て板 2 日延長）")]
    c = Contract("KK2026", 2026, 1, "kikan", "ときわ北部クリーンセンター基幹的設備改良工事（令和7〜9年度）", "基幹的設備改良工事（R8 1号炉）",
                 W.UNION, W.MAKER, D(2026, 8, 24), D(2026, 12, 18), rows, ev, plan_end=D(2026, 12, 18))
    c.status = "施工中"
    return c

def build_all():
    cs = []
    for y in range(2009, 2027):
        for f in (1, 2):
            if y in teishu_years(f):
                cs.append(make_teishu(f, y))
        if y == 2019:
            cs.append(make_dcs())
        if y in (2025, 2026):
            cs.append(make_kikan(y))
    return cs

ALL = build_all()

if __name__ == "__main__":
    for c in ALL:
        print(c.cid, c.title, c.start, c.end, len(c.tasks), len(c.direct), len(c.parts))
