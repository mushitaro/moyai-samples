# -*- coding: utf-8 -*-
"""定期整備工事（SPC→鷹羽）の工程を、その年度・その炉の実作業（world.actual_work）から組み立てる。
2024 年度 1号炉・2025 年度 2号炉は見本で承認された手書きの工程（contracts.T2024/T2025）を使う。
工程 → 日報（contracts.daily_rows）→ 報告書・発注 はすべてここで作る工程から出す。"""
import datetime as dt
import math
import random
import world as W
import contracts as K

D = dt.date
CREW_MEN = {"機械1班": 6, "機械2班": 5, "機械3班": 4, "補修班": 4, "築炉班": 5, "鳶班": 5, "検査班": 3, "電気班": 2,
            "タービン班": 6, "SPC運転": 0}
CREW_NAME = dict(K.CREW_2025)
CREW_NAME["機械3班（応援）"] = "野分工業 機械3班（応援）"
CREW_NAME["機械1班（応援）"] = "野分工業 機械1班（応援）"
CREW_NAME["補修班（応援）"] = "野分工業 補修班（応援）"

# 作業の日数（予定）
DUR = {("furnace", "R"): 8, ("stoker_c", "R"): 7, ("stoker_d", "R"): 3, ("stoker_a", "R"): 3, ("feeder", "R"): 3,
       ("hydraulic", "O"): 3, ("burner", "O"): 2, ("spray", "R"): 1, ("conv", "R"): 2, ("fan_s", "O"): 2, ("fan_m", "O"): 3,
       ("fdf", "O"): 5, ("idf", "O"): 4, ("boiler", "R"): 3, ("sh", "R"): 4, ("eco", "R"): 3, ("sootb", "O"): 4,
       ("tower", "R"): 4, ("pump_r", "O"): 2, ("pump_s", "O"): 2, ("feeder_s", "O"): 2, ("blower", "O"): 2,
       ("bf", "X"): 6, ("bf", "R"): 2, ("cat", "R"): 2, ("cat", "X"): 5, ("duct", "R"): 3, ("struct", "R"): 3,
       ("ashpush", "R"): 4, ("bfp_r", "O"): 3, ("bfp_s", "O"): 3, ("bfp_r", "X"): 3, ("inv", "O"): 3, ("inv", "R"): 2,
       ("panel", "R"): 2, ("panel", "X"): 4, ("dcs", "R"): 3, ("turbine", "O"): 15, ("accfan", "O"): 3, ("acc", "R"): 4,
       ("vessel", "O"): 2, ("gate", "R"): 1, ("damper", "R"): 1, ("small", "O"): 1, ("small", "R"): 1}

# 部品（SPC 支給）：品名, 型式, 数量（関数）, 単位, 単価（2016 年価格）, 予備数
PARTS = {
    ("stoker_c", "R"): [("燃焼ストーカ 可動火格子", "TKW-RS105C-G 鋳鋼", lambda r: r.choice([36, 38, 42, 48, 96]), "個", 62000, 0)],
    ("stoker_d", "R"): [("乾燥ストーカ 可動火格子", "TKW-RS105D-G 鋳鋼", lambda r: r.choice([18, 24, 24, 30]), "個", 58000, 2)],
    ("stoker_a", "R"): [("後燃焼ストーカ 可動火格子", "TKW-RS105A-G 鋳鋼", lambda r: r.choice([16, 18, 20]), "個", 57000, 2)],
    ("feeder", "R"): [("給じん装置 プッシャライナ", "SCH13 t25", lambda r: 4, "枚", 86000, 0)],
    ("hydraulic", "O"): [("ストーカ油圧ポンプ シールキット", "14MPa 用", lambda r: 2, "組", 148000, 0), ("作動油", "ISO VG46 200L", lambda r: 3, "缶", 52000, 0)],
    ("burner", "O"): [("バーナ ノズルチップ・点火トランス", "120L/h 用", lambda r: 1, "式", 168000, 0)],
    ("spray", "R"): [("炉内水噴射ノズル", "二流体 3/8B", lambda r: 4, "本", 38000, 0)],
    ("conv", "R"): [("コンベヤ チェーン・スクレーパ", "", lambda r: 1, "式", 980000, 0)],
    ("fan_m", "O"): [("送風機 軸受", "6316ZZ", lambda r: 2, "個", 22000, 0), ("Vベルト", "C-150", lambda r: 4, "本", 6200, 0)],
    ("fdf", "O"): [("押込送風機 軸受", "22224 両側", lambda r: 2, "個", 142000, 0)],
    ("idf", "O"): [("誘引通風機 軸受", "22228CCK/W33 + スリーブ", lambda r: 2, "個", 158000, 1)],
    ("sh", "R"): [("過熱器管", "STBA24 φ38.1×t4.5 L=2.0m", lambda r: 4, "本", 96000, 1)],
    ("sootb", "O"): [("スートブロワ 補修部品（パッキン・チェーン）", "長抜差・定置 一式", lambda r: 1, "式", 1080000, 0)],
    ("tower", "R"): [("減温塔 噴霧ノズル", "二流体 1/2B", lambda r: 6, "本", 44000, 2)],
    ("pump_r", "O"): [("ポンプ メカニカルシール", "", lambda r: 1, "組", 128000, 0), ("ポンプ 軸受", "6309C3", lambda r: 2, "個", 9800, 0)],
    ("pump_s", "O"): [("ポンプ メカニカルシール", "", lambda r: 1, "組", 128000, 0), ("ポンプ 軸受", "6309C3", lambda r: 2, "個", 9800, 0)],
    ("feeder_s", "O"): [("定量供給機 補修部品", "テーブル・スクリュー", lambda r: 1, "式", 320000, 0)],
    ("blower", "O"): [("気送ブロワ 補修部品", "ルーツ用", lambda r: 1, "式", 98000, 0)],
    ("bf", "X"): [("BF ろ布", "PTFE φ165×6000", lambda r: 360, "本", 21000, 0), ("BF ろ布リテーナ", "SUS φ160×6000", lambda r: 30, "本", 9600, 0)],
    ("cat", "R"): [("脱硝触媒 ブロック", "ハニカム 1段分", lambda r: 24, "個", 410000, 0)],
    ("bfp_r", "O"): [("ボイラ給水ポンプ メカニカルシール", "SPF-55用", lambda r: 1, "組", 182000, 0), ("ボイラ給水ポンプ 軸受", "6312C3", lambda r: 2, "個", 17200, 0)],
    ("bfp_s", "O"): [("ボイラ給水ポンプ メカニカルシール", "SPF-55用", lambda r: 1, "組", 182000, 0), ("ボイラ給水ポンプ 軸受", "6312C3", lambda r: 2, "個", 17200, 0)],
    ("inv", "O"): [("インバータ 冷却ファン・平滑コンデンサ", "", lambda r: 1, "式", 690000, 0)],
    ("turbine", "O"): [("タービン 軸受メタル・パッキン類", "2,400kW 用", lambda r: 1, "式", 6800000, 0)],
    ("accfan", "O"): [("空冷ファン 減速機 軸受・オイルシール", "", lambda r: 1, "式", 268000, 0)],
}

VERB = {"I": "点検", "O": "分解整備", "R": "補修", "X": "更新", "U": "改良"}

STEP_O = [["{n} 縁切り・分解", "{n} 開放、縁切り"], ["{n} 分解点検・摩耗測定", "{n} 部品点検、寸法測定"],
          ["{n} {p} 交換", "{n} {p} 取替"], ["{n} 組立", "{n} 組付け"], ["{n} 芯出し・試運転", "{n} 試運転、振動・温度確認"]]
STEP_R = [["{n} 補修範囲 確認・板厚測定", "{n} 損耗部 確認、けがき"], ["{n} 不良部 撤去", "{n} 損耗部 はつり・切断"],
          ["{n} {p}", "{n} {p} 施工"], ["{n} 仕上げ・確認", "{n} 仕上げ、SPC 確認"]]
REPAIR = {"furnace": "キャスタブル 打設", "stoker_c": "可動火格子 取付", "stoker_d": "可動火格子 取付", "stoker_a": "可動火格子 取付",
          "feeder": "ライナ 取替・溶接", "spray": "ノズル 交換", "conv": "チェーン・スクレーパ 交換", "boiler": "減肉部 肉盛溶接",
          "sh": "過熱器管 部分取替", "eco": "腐食部 肉盛", "tower": "内筒 当て板・ノズル交換", "duct": "腐食部 当て板",
          "struct": "腐食部 当て板", "ashpush": "底板 当て板 溶接", "cat": "触媒 1段 入替", "bf": "破損ろ布 交換", "acc": "管束 洗浄・フィン補修",
          "gate": "シール 交換", "damper": "軸封 交換", "small": "部品交換", "panel": "部品交換", "inv": "部品交換", "dcs": "部品交換"}
_ST = [["{n} 可動火格子 取外し", "{n} 旧火格子 取外し"], ["{n} 取外し 続き・火格子受け 清掃", "{n} 火格子受け 点検・清掃"],
       ["{n} 新火格子 取付", "{n} 新火格子 組付け"], ["{n} 新火格子 取付（続き）", "{n} 組付け 続き"], ["{n} すき間調整・手動作動確認", "{n} すき間 測定・作動確認"]]
STEP_SPEC = {
    "stoker_c": _ST, "stoker_d": _ST, "stoker_a": _ST,
    "furnace": [["{n} 不良部 はつり（左側）", "{n} 損耗部 はつり"], ["{n} はつり（右側）・アンカー溶接", "{n} アンカー 溶接"],
                ["{n} 型枠 組立", "{n} 型枠"], ["{n} キャスタブル 打設", "{n} 打設"], ["{n} 打設（続き）・養生", "{n} 養生"],
                ["{n} 脱型・仕上げ", "{n} 仕上げ・補修面積 確認"]],
    "conv": [["{n} チェーン 伸び測定・取外し", "{n} スクレーパ 取外し"], ["{n} 新チェーン・スクレーパ 取付・テンション調整", "{n} 取付・試運転"]],
}
PARTS_WORD = {"idf": "軸受", "fdf": "軸受", "fan_m": "軸受・Vベルト", "hydraulic": "シール", "burner": "ノズルチップ", "sootb": "パッキン・チェーン",
              "pump_r": "メカシール・軸受", "pump_s": "メカシール・軸受", "bfp_r": "メカシール・軸受", "bfp_s": "メカシール・軸受",
              "feeder_s": "摩耗部品", "blower": "ロータ・軸受", "inv": "冷却ファン・コンデンサ", "turbine": "軸受メタル", "accfan": "減速機 軸受",
              "fan_s": "軸受", "small": "消耗部品", "vessel": "パッキン"}

# 年度・炉ごとの発見と特記（日報の特記・報告書の所見に出す。世界の状態と合わせた）
FINDINGS = {
    (2016, 1): "ボイラ水管 UT 最小 4.4mm。SPC 受託後 初の定修のため 全数 写真記録",
    (2016, 2): "2号ボイラ 前壁水管 UT 最小 4.1mm（他部位より薄い）。次回以降 重点測定",
    (2017, 1): "燃焼ストーカ 火格子すき間 3〜5mm 良",
    (2017, 2): "2号ボイラ 前壁 UT 最小 3.9mm。減肉速度 約0.15mm/年",
    (2018, 1): "押込送風機 軸受 部品納期遅れのため 分解整備を 2019 年度へ繰越（SPC 協議）",
    (2018, 2): "2号ボイラ 前壁 UT 最小 3.7mm",
    (2019, 1): "BF ろ布 360本 全数交換 完了。旧ろ布 は産廃処分",
    (2019, 2): "2号ボイラ 前壁 UT 最小 3.5mm。肉盛 範囲の検討を SPC に依頼",
    (2020, 1): "触媒 活性試験の結果 良好のため 全量更新を見送り（SPC 判断）",
    (2020, 2): "新型コロナ対応で 作業員を縮小、ストーカ油圧装置 分解整備を翌年度へ。煙突内筒 下部に腐食（SPC→市へ報告）",
    (2021, 1): "誘引通風機 振動 3.8mm/s（傾向監視）",
    (2021, 2): "触媒 1段目のみ交換（全量更新 取りやめ）",
    (2022, 1): "灰押出装置 底板 残厚 5.1mm（当て板範囲外）。誘引通風機 振動 4.1mm/s",
    (2022, 2): "2号ボイラ 前壁水管 UT 最小 3.2mm（必要厚 2.9mm）→ SPC から市へ報告",
    (2023, 1): "誘引通風機 振動 4.4mm/s 上昇傾向。燃焼ストーカ 火格子すき間 4〜9mm",
    (2023, 2): "前壁水管の取替は 市追加工事（鷹羽）で別途施工。燃焼ストーカ すき間 測定 → 交換を 2025 へ",
    (2026, 2): "2号 BF ろ布 全数交換（廃止まで 1 回）",
    (2026, 1): "1号過熱器 2次入口 部分取替。1号 BF ろ布 全数交換",
}
DELAY_REASON = ["固着ボルトの切断に時間を要し 1日延長", "部品の寸法違い 手直しで 1日延長", "雨天で搬入できず 1日延長",
                "追加の当て板が必要となり 1日延長", "立会日程の調整で 1日後ろへ"]


def crew_of(e):
    if e.strat == "furnace":
        return "築炉班"
    if e.f == 0:
        if e.strat in ("turbine", "gen", "vessel") or e.cat == "05" and e.strat == "small":
            return "タービン班"
        return "補修班"
    return {"01": "機械1班", "02": "機械1班", "03": "機械2班", "08": "機械2班", "04": "機械3班", "06": "機械3班",
            "07": "補修班", "10": "電気班", "11": "電気班"}[e.cat]


IN_OH_COMMON = ("turbine", "gen", "acc", "accfan", "vessel")


def oh_furnace(y):
    """その年度に秋の定修を行う炉。"""
    return 1 if y % 2 == 0 else 2


def contract_works(y, f):
    """(eq, attr集合) のリスト。炉別機器と、全停・秋定修に入れる共通機器。"""
    out = []
    for e in W.EQ:
        a = W.actual_work(e, y)
        if not a:
            continue
        if (e.name, e.f, y) in W.CITY_WORKS:
            a = a - {"R", "X"}
            if not a:
                continue
        if e.f == f:
            out.append((e, a))
        elif e.f == 0 and e.strat in IN_OH_COMMON and oh_furnace(y) == f:
            out.append((e, a))
    return out


def stop_start(y, f):
    if (y, f) == (2024, 1):
        return D(2024, 10, 1)
    if (y, f) == (2025, 2):
        return D(2025, 9, 29)
    m = W.oh_month(f, y)
    x = D(y, m, 1)
    while x.weekday() != 0:
        x += dt.timedelta(days=1)
    if m == 5:
        x += dt.timedelta(days=7)
    return x


def wd_after(x, n=1):
    """x から n 稼働日後（日曜を飛ばす）。"""
    while n > 0:
        x += dt.timedelta(days=1)
        if x.weekday() != 6:
            n -= 1
    return x


def span(start, n):
    """start（稼働日）から n 稼働日の (開始, 終了)。"""
    while start.weekday() == 6:
        start += dt.timedelta(days=1)
    e = start
    for _ in range(n - 1):
        e = wd_after(e)
    return start, e


def md(x):
    return f"{x.month}/{x.day}"


def build(y, f):
    """戻り値：dict(cid, title, period, stop, tasks)"""
    cid = f"OH{y}-{f}"
    turb = any(e.strat == "turbine" for e, a in contract_works(y, f) if "O" in a)
    title = f"{y}年度 ひばり野CC {f}号炉定期整備工事" + ("（蒸気タービン開放点検を含む）" if turb else "")
    if (y, f) == (2024, 1):
        tasks = K.T2024
    elif (y, f) == (2025, 2):
        tasks = K.T2025
    else:
        tasks = _gen(y, f)
    st = stop_start(y, f)
    allds = [x for t in tasks for x in K.task_days(t, y)]
    pstart = D(y, st.month - 1, 1) if st.month > 1 else st
    while pstart.weekday() >= 5:
        pstart += dt.timedelta(days=1)
    end = max(allds)
    pm = end.month + 1
    pend = D(y + (pm > 12), (pm - 1) % 12 + 1, 1) - dt.timedelta(days=1)
    pend = D(pend.year, pend.month, 28) if pend.weekday() >= 5 else pend
    while pend.weekday() >= 5:
        pend -= dt.timedelta(days=1)
    if (y, f) == (2024, 1):
        pstart, pend = D(2024, 9, 2), D(2024, 11, 29)
    if (y, f) == (2025, 2):
        pstart, pend = D(2025, 9, 1), D(2025, 11, 28)
    K.CONTRACTS[cid] = dict(title=title, orderer=W.SPC, contractor=W.MAKER, sub=W.SUB, kind="定期整備",
                            period=(pstart, pend), fy=y)
    return dict(cid=cid, title=title, period=(pstart, pend), stop=st, tasks=tasks, y=y, f=f)


def _gen(y, f):
    r = random.Random(W.SEED + y * 10 + f)
    works = contract_works(y, f)
    d0 = stop_start(y, f)
    legal = (f == 1 and y % 2 == 0) or (f == 2 and y % 2 == 1)
    T = []

    def add(area, item, crew, men, ps, pe, act, eqs, attr, steps):
        T.append([f"T{len(T) + 1:02d}", area, item, crew, men, (md(ps), md(pe)),
                  (md(act[0]), md(act[1])) if act else None, eqs, attr, steps])

    # 共通の前段
    add("共通", "炉立下げ・燃料切れ運転", "SPC運転", 0, d0, d0 + dt.timedelta(days=1), None, [("焼却炉本体", f)], "—",
        [f"{f}号炉 ごみ投入停止、燃切り運転", "燃切り完了、バーナ停止"])
    add("共通", "炉内冷却", "SPC運転", 0, d0 + dt.timedelta(days=1), d0 + dt.timedelta(days=3), None, [("焼却炉本体", f)], "—",
        ["押込送風機で強制冷却", "強制冷却継続", "炉内 60℃以下 確認"])
    c3 = wd_after(d0 + dt.timedelta(days=1))
    add("共通", "縁切り・電源遮断・施錠", "電気班", 2, c3, c3, None, [], "—",
        [f"{f}号炉 MCC 遮断・施錠、縁切り板 挿入（煙道・飛灰）"])
    s4 = span(d0 + dt.timedelta(days=3), 2)
    add("共通", "足場組立（炉内・ボイラ）", "鳶班", 5, s4[0], s4[1], None, [], "—", ["炉内足場 組立", "ボイラ 足場 組立、点検"])
    s5 = span(s4[1], 2)
    add("燃焼設備", "炉内清掃・クリンカ除去", "機械1班", 6, s5[0], s5[1], None, [("焼却炉本体", f)], "INSPECT",
        ["炉内 クリンカ はつり", "残灰除去・ストーカ下 清掃"])
    # 作業の組立て
    lanes = {}
    start = {"機械1班": wd_after(s5[1]), "機械2班": s5[0], "機械3班": s5[0], "補修班": wd_after(s5[0]), "築炉班": wd_after(s5[1], 2),
             "電気班": wd_after(s5[0], 4), "タービン班": d0 + dt.timedelta(days=2), "検査班": None}
    jobs = {}
    insp = {}
    for e, a in works:
        cw = crew_of(e)
        key = next((k for k in "XUOR" if k in a), None)
        if key:
            dur = DUR.get((e.strat, key), 2)
            jobs.setdefault(cw, []).append((e, key, dur))
        elif "I" in a and e.strat not in ("furnace",):
            insp.setdefault(cw, []).append(e)
    # ボイラ水洗は毎年、UT は検査班
    plan_tasks = []
    for cw in ("機械1班", "機械2班", "機械3班", "補修班", "築炉班", "電気班", "タービン班"):
        seq = []
        if cw == "機械2班":
            seq.append(("wash", None, 4))
        for e, key, dur in sorted(jobs.get(cw, []), key=lambda x: -x[2]):
            seq.append(("job", (e, key), dur))
        ie = insp.get(cw, [])
        for i in range(0, len(ie), 4):
            chunk = ie[i:i + 4]
            seq.append(("insp", chunk, max(1, math.ceil(len(chunk) / 2))))
        if cw == "機械2班" and legal:
            seq.append(("hydro", None, 1))
        plan_tasks.append((cw, seq))
    # 予定と実績を別々に詰める（実績は遅れを足す）
    delays = {}
    for cw, seq in plan_tasks:
        for i, (kind, obj, dur) in enumerate(seq):
            if kind == "job" and r.random() < 0.18:
                delays[(cw, i)] = r.choice(DELAY_REASON)
    end_all = {"p": d0, "a": d0}
    out = []
    for cw, seq in plan_tasks:
        if not seq:
            continue
        load = sum(x[2] for x in seq)
        nl = 2 if load > 15 and cw in ("機械1班", "機械3班", "補修班") else 1
        cur = {"p": [start[cw]] * nl, "a": [start[cw]] * nl}
        for i, (kind, obj, dur) in enumerate(seq):
            lane = min(range(nl), key=lambda k: cur["p"][k])
            res = {}
            for m in ("p", "a"):
                dd = dur + (1 if (m == "a" and (cw, i) in delays) else 0)
                s, e_ = span(cur[m][lane], dd)
                res[m] = (s, e_)
                cur[m][lane] = wd_after(e_)
                end_all[m] = max(end_all[m], e_)
            crew = cw if lane == 0 else cw + "（応援）"
            out.append((crew, kind, obj, res, delays.get((cw, i))))
    # 検査班 UT（水洗の後）
    wash = [o for o in out if o[1] == "wash"][0]
    ut_s = wd_after(wash[3]["p"][1])
    utp = span(ut_s, 2)
    uta = span(wd_after(wash[3]["a"][1]), 2)
    for crew, kind, obj, res, dl in out:
        if kind == "wash":
            add("燃焼ガス冷却設備", "ボイラ水洗・清掃", crew, 5, res["p"][0], res["p"][1], res["a"] if res["a"] != res["p"] else None,
                [("ボイラ本体", f), ("エコノマイザ", f)], "INSPECT", None)
    add("燃焼ガス冷却設備", "ボイラ水管・過熱器 肉厚測定（UT）", "検査班", 3, utp[0], utp[1], uta if uta != utp else None,
        [("ボイラ本体", f), ("過熱器", f)], "INSPECT", None)
    if legal:
        x = wd_after(utp[1], 2)
        add("燃焼ガス冷却設備", "ボイラ定期事業者検査（開放検査）", "機械2班", 2, x, x, None,
            [("ボイラ本体", f), ("蒸気ドラム", f), ("過熱器", f)], "INSPECT", None)
    for crew, kind, obj, res, dl in out:
        act = res["a"] if res["a"] != res["p"] else None
        if kind == "job":
            e, key = obj
            area = "共通（" + W.CATS[e.cat] + "）" if e.f == 0 else W.CATS[e.cat]
            item = f"{e.name} {VERB[key]}" if key != "X" or e.strat != "bf" else f"{e.name} ろ布全数交換"
            men = min(CREW_MEN[crew.replace("（応援）", "")], 2 if DUR.get((e.strat, key), 2) <= 2 and e.strat in ("pump_r", "pump_s", "blower", "small", "spray", "damper", "gate") else 9)
            if crew.endswith("（応援）"):
                men = max(2, men - 2)
            add(area, item, crew, men, res["p"][0], res["p"][1], act, [(e.name, e.f)], {"O": "OVERHAUL", "R": "OVERHAUL", "X": "REPLACE", "U": "UPGRADE"}[key],
                None)
            T[-1].append(dl)
        elif kind == "insp":
            names = "・".join(x.name for x in obj)
            area = W.CATS[obj[0].cat] if obj[0].f else "共通（" + W.CATS[obj[0].cat] + "）"
            add(area, f"{names} 点検", crew, 2 if len(obj) <= 2 else 3, res["p"][0], res["p"][1], act, [(x.name, x.f) for x in obj],
                "INSPECT", None)
        elif kind == "hydro":
            add("燃焼ガス冷却設備", "ボイラ水圧試験", crew, 4, res["p"][0], res["p"][1], act, [("ボイラ本体", f)], "INSPECT", None)
    # 後段
    ep, ea = end_all["p"], end_all["a"]
    for lab, crew, men, n in (("足場解体・炉内片付け", "鳶班", 5, 2), ("マンホール閉鎖・縁切り復旧", "機械1班", 4, 1), ("単体試運転・インタロック確認", "機械2班", 4, 2)):
        sp_ = span(wd_after(ep), n)
        sa_ = span(wd_after(ea), n)
        add("共通", lab, crew, men, sp_[0], sp_[1], sa_ if sa_ != sp_ else None, [], "—", None)
        ep, ea = sp_[1], sa_[1]
    for lab, n in (("炉立上げ・耐火物乾燥・昇温", 2), ("負荷運転確認", 3)):
        sp_ = (ep + dt.timedelta(days=1), ep + dt.timedelta(days=1 + n))
        sa_ = (ea + dt.timedelta(days=1), ea + dt.timedelta(days=1 + n))
        add("共通", lab, "SPC運転" if lab.startswith("炉") else "機械2班", 0 if lab.startswith("炉") else 2, sp_[0], sp_[1],
            sa_ if sa_ != sp_ else None, [("焼却炉本体", f)] if lab.startswith("炉") else [], "—", None)
        ep, ea = sp_[1], sa_[1]
    # 日々の記述
    for t in T:
        if t[9] is None:
            t[9] = _steps(t, y, f, r)
    return [tuple(t[:10]) + ((t[10],) if len(t) > 10 else ()) for t in T]


def _steps(t, y, f, r):
    days = [x for x in K.task_days(t, y) if x.weekday() != 6] if t[3] != "SPC運転" else K.task_days(t, y)
    n = len(days)
    item = t[2]
    v = (y + f) % 2
    eqn = [W.eq(nm, ff) for nm, ff in t[7]] if t[7] else []
    if item == "ボイラ水洗・清掃":
        base = ["第1放射室 付着灰 落とし", "第2・第3放射室 水洗", "過熱器部 水洗、エコノマイザ 清掃", "ホッパ部 灰出し、乾燥", "残り 清掃・乾燥"]
        return [base[min(i, 4)] if i < 4 else "清掃 続き（付着多い）" for i in range(n)]
    if "UT" in item:
        th = boiler_min(f, y)
        return [f"水管 前壁・側壁 UT 測定（{r.choice([180, 200, 216])}点）", f"過熱器 UT 測定。水管 最小 {th:.1f}mm"][:n] + ["再測定"] * max(0, n - 2)
    if "定期事業者検査" in item:
        return ["ドラム内部・水管内面 確認（SPC 立会）、UT 結果 説明"]
    if "水圧試験" in item:
        return ["水圧試験 4.5MPa 30分保持、漏れなし（SPC・検査員 立会）"]
    if item.endswith(" 点検"):
        names = item[:-3].split("・")
        out = []
        for i in range(n):
            chunk = names[i * 2:(i + 1) * 2] or names[-1:]
            res = r.choice(["異常なし", "異常なし", "軽微な摩耗、次回確認", "給油・増締め", "清掃・調整"])
            out.append("、".join(chunk) + f" 点検 {res}")
        return out
    if t[8] in ("OVERHAUL", "REPLACE", "UPGRADE") and eqn:
        e = eqn[0]
        nm = (f"{e.f}炉 " if e.f else "") + (W.ALIAS.get(e.name, [e.name])[0])
        key = "R" if "補修" in item else ("X" if t[8] == "REPLACE" else "O")
        if key == "X" and e.strat == "bf":
            out = []
            per = max(1, math.ceil(6 / n))
            for i in range(n):
                a, b = i * per + 1, min(6, (i + 1) * per)
                out.append((f"旧ろ布 撤去・新ろ布 取付 {a}室" if a == b else f"旧ろ布 撤去・新ろ布 取付 {a}〜{b}室") if a <= 6 else "パルス弁 作動確認・差圧確認")
            return out
        tpl = STEP_SPEC.get(e.strat, STEP_R) if key == "R" else STEP_O
        p = REPAIR.get(e.strat, "部品交換") if key == "R" else PARTS_WORD.get(e.strat, "消耗部品")
        out = []
        for i in range(n):
            j = min(len(tpl) - 1, int(i * len(tpl) / n))
            s = tpl[j][v].format(n=nm, p=p)
            if i > 0 and out and s == out[-1]:
                s += "（続き）"
            out.append(s)
        if len(t) > 10 and t[10]:
            out[-1] = out[-1] + f"（{t[10]}）"
        return out
    if item.startswith("足場解体"):
        return ["ボイラ 足場 解体", "炉内 足場 解体・搬出"][:n] + ["片付け"] * max(0, n - 2)
    if item.startswith("マンホール"):
        return ["マンホール 閉鎖（炉・ボイラ・BF）、縁切り板 撤去"] * n
    if item.startswith("単体試運転"):
        return ["送風機・ポンプ 単体試運転", "ストーカ・給じん 連動、インタロック確認（SPC 立会）"][:n]
    if item.startswith("炉立上げ"):
        return ["バーナ着火、耐火物乾燥", "昇温継続", "ごみ投入開始"][:n] + [""] * max(0, n - 3)
    if item.startswith("負荷運転"):
        return ["負荷運転 立会", "負荷運転 立会 継続", "定格負荷 確認", "SPC へ引渡し"][:n]
    return [item] * n


def boiler_min(f, y):
    """ボイラ水管の最小肉厚（UT）の正解値。"""
    if f == 1:
        return round(4.45 - 0.06 * (y - 2016), 1)
    if y <= 2022:
        return round(4.25 - 0.17 * (y - 2016), 1)
    return round(3.75 - 0.12 * (y - 2023), 1)       # 前壁取替後は残りの部位の最小


def all_contracts():
    out = []
    for y in range(2016, 2027):
        for f in (1, 2):
            if W.oh_month(f, y) == 5:
                out.append(build(y, f))
        for f in (1, 2):
            if W.oh_month(f, y) == 10:
                out.append(build(y, f))
    return out


# ---------------------------------------------------------------- SPC 支給部品（発注 CSV と報告書の材料表で共用）
def _sample_parts(case):
    from orders_src import NEW_ROWS
    out = []
    for row in NEW_ROWS:
        (pr, dd, cs, kind, code, nm, sp, q, u, up, want, ans, acc, vc, asset, memo) = row
        if cs == case and kind == "部品":
            used = q
            if "予備" in memo:
                import re
                m = re.search(r"(\d+)(枚|個|本)使用", memo)
                used = int(m.group(1)) if m else q
            out.append(dict(asset=asset, name=nm, spec=sp, qty=q, unit=u, price=up, used=used, spare=q - used, row=row))
    return out


def parts_for(c):
    """契約 c の SPC 支給部品。[{asset, name, spec, qty, unit, price, used, spare, eq}]"""
    y, f = c["y"], c["f"]
    if (y, f) == (2024, 1):
        return _sample_parts("M24-01-OH")
    if (y, f) == (2024, 2):
        return _sample_parts("M24-02-OH")
    r = random.Random(W.SEED + y * 100 + f * 7)
    esc = 1 + 0.018 * (y - 2016) + (0.06 if y >= 2022 else 0)
    out = []
    for t in c["tasks"]:
        if t[8] not in ("OVERHAUL", "REPLACE", "UPGRADE"):
            continue
        for nm, ff in t[7]:
            e = W.eq(nm, ff)
            a = W.actual_work(e, y) or set()
            key = next((k for k in "XUOR" if k in a), None)
            for p in PARTS.get((e.strat, key), []):
                q = p[2](r)
                if (y, f, e.strat) == (2025, 2, "stoker_c"):
                    q = 110
                pn = f"{e.name} {p[0].split(' ', 1)[-1]}" if e.strat in ("conv", "pump_r", "pump_s", "feeder_s", "blower", "fan_m") else (f"{e.name} " + p[0].split(" ", 1)[-1] if e.strat == "burner" else p[0])
                out.append(dict(asset=e.no, name=pn, spec=p[1], qty=q + p[5], unit=p[3], price=int(round(p[4] * esc, -2)),
                                used=q, spare=p[5], eq=e))
    return out
