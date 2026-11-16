# -*- coding: utf-8 -*-
"""本作りの正解の世界：2019–2026 年度の全契約・工程（予定／実績）・毎日の作業・実際の作業（works）。
帳票の生成器はすべてここから書き分ける。物語の今日 2026-11-13、乱数の種 1111。"""
import datetime as dt, random
import world as W
import plan_data as P
import r5grate as G

TODAY = dt.date(2026, 11, 13)
YEARS = list(range(2019, 2027))
WD = "月火水木金土日"

def D(y, m, d):
    return dt.date(y, m, d)

def wdays(a, b, sat=()):
    out, d = [], a
    while d <= b:
        if d.weekday() < 5 or d in sat:
            out.append(d)
        d += dt.timedelta(1)
    return out

def nxt(d, n=1):
    """n 稼働日後"""
    k = 0
    while k < n:
        d += dt.timedelta(1)
        if d.weekday() < 5:
            k += 1
    return d

# ---------------------------------------------------------------------------
# 年度ごとの停止日程（運転委託との調整で決めた実際の日程）
# ---------------------------------------------------------------------------
SHUT = {
    2019: dict(u1=(D(2019, 6, 3), D(2019, 6, 28)), u2=(D(2019, 10, 15), D(2019, 11, 11)), zen=(D(2019, 10, 30), D(2019, 10, 31)), crane=(D(2019, 11, 5), D(2019, 11, 7))),
    2020: dict(u1=(D(2020, 6, 1), D(2020, 6, 26)), u2=(D(2020, 10, 14), D(2020, 11, 9)), zen=(D(2020, 10, 28), D(2020, 10, 29)), crane=(D(2020, 11, 16), D(2020, 11, 18))),
    2021: dict(u1=(D(2021, 6, 7), D(2021, 7, 2)), u2=(D(2021, 10, 11), D(2021, 11, 8)), zen=(D(2021, 10, 27), D(2021, 10, 28)), crane=(D(2021, 11, 9), D(2021, 11, 11))),
    2022: dict(u1=(D(2022, 6, 6), D(2022, 7, 1)), u2=(D(2022, 10, 11), D(2022, 11, 7)), zen=(D(2022, 10, 26), D(2022, 10, 27)), crane=(D(2022, 11, 14), D(2022, 11, 16))),
    2023: dict(u1=(D(2023, 6, 5), D(2023, 6, 28)), u2=(D(2023, 10, 16), D(2023, 11, 10)), zen=(D(2023, 10, 26), D(2023, 10, 27)), crane=(D(2023, 11, 6), D(2023, 11, 9))),
    2024: dict(u1=(D(2024, 6, 3), D(2024, 6, 27)), u2=(D(2024, 10, 15), D(2024, 11, 15)), zen=(D(2024, 10, 24), D(2024, 10, 25)), crane=(D(2024, 11, 18), D(2024, 11, 20))),
    2025: dict(u1=(D(2025, 6, 2), D(2025, 6, 26)), u2=(D(2025, 10, 6), D(2025, 11, 7)), zen=(D(2025, 10, 23), D(2025, 10, 24)), crane=(D(2025, 11, 4), D(2025, 11, 6))),
    2026: dict(u1=(D(2026, 6, 1), D(2026, 6, 25)), u2=(D(2026, 10, 13), D(2026, 11, 20)), zen=(D(2026, 10, 27), D(2026, 10, 28)), crane=(D(2026, 11, 10), D(2026, 11, 12))),
}
CRANE_PERF = {2019, 2021, 2023, 2025}

# 測定値（年度 → 値）。帳票の数字はここから
MEAS = {
    ("grate", "1"): {2019: 8.6, 2020: 9.5, 2021: 10.6, 2022: 11.8, 2023: 16.5, 2024: 11.4, 2025: 12.2, 2026: 12.9},
    ("grate", "2"): {2019: 7.2, 2020: 8.0, 2021: 9.1, 2022: 10.2, 2023: 11.2, 2024: 12.4, 2025: 14.8, 2026: 9.6},
    ("boiler", "1"): {2019: 2.8, 2020: 2.6, 2021: 2.4, 2022: 2.7, 2023: 2.65, 2024: 2.6, 2025: 2.55, 2026: 2.5},
    ("boiler", "2"): {2019: 2.9, 2020: 2.85, 2021: 2.8, 2022: 2.75, 2023: 2.7, 2024: 2.7, 2025: 2.65, 2026: 2.6},
    ("idf", "1"): {2019: 1.4, 2020: 1.5, 2021: 1.6, 2022: 1.7, 2023: 1.8, 2024: 1.2, 2025: 1.3, 2026: 1.4},
    ("idf", "2"): {2019: 2.3, 2020: 2.8, 2021: 3.6, 2022: 4.5, 2023: 4.8, 2024: 5.0, 2025: 1.6, 2026: 1.9},
    ("bf", "1"): {2019: 1.8, 2020: 1.0, 2021: 1.1, 2022: 1.2, 2023: 1.3, 2024: 1.4, 2025: 1.5, 2026: 1.6},
    ("bf", "2"): {2019: 1.7, 2020: 1.9, 2021: 1.0, 2022: 1.1, 2023: 1.2, 2024: 1.3, 2025: 1.4, 2026: 1.5},
}

# ---------------------------------------------------------------------------
# 機器ごと・年度ごとの実際の記号（○◎△●※）
# ---------------------------------------------------------------------------
def mark(e, y):
    if y == 2019:
        n = e.name
        off = 0 if e.unit == "1" else 1
        if e.unit == "共通":
            return {"ごみクレーン": "○", "灰クレーン": "○", "中央監視装置": "○", "トラックスケール": "○",
                    "主灰共通コンベヤ": "○", "飛灰混練機": "○", "冷却塔": "○", "空冷式蒸気復水器": "○"}.get(n)
        if n in P.YEARLY:
            return "○"
        if n in P.ALT and (y + off) % 2 == 1:
            return "○"
        return None
    if y == 2026:
        return P.l2_plan(e).get(2026)
    return P.actual(e).get(y)

# 補修工事・購入・他業者に回す作業
SPECIAL = {
    ("1D-05", 2020): "R2020a", ("2D-05", 2021): "R2021a", ("1C-01", 2022): "R2022a", ("0L-01", 2020): "R2020b",
    ("1B-03", 2022): "P2022a", ("1B-03", 2023): "R2023a", ("2F-07", 2024): "P2024a", ("0G-08", 2024): "R2024a",
    ("2B-03", 2025): "R2025a",
}
ATTR = {"○": "INSPECT", "◎": "OVERHAUL", "△": "OVERHAUL", "●": "REPLACE"}

# 定修の作業のまとまり（班が 1 日で動く単位）。キー → (名前, 機器名の集合 or 区分, 日数の目安)
GROUPS_UNIT = [
    ("clean", "炉内清掃・クリンカ除去", ["焼却炉本体"], 2),
    ("feed", "受入供給・給じん装置 点検", ["ごみ投入ホッパ", "ホッパゲート", "ブリッジ解除装置", "給じん装置"], 1),
    ("grate", "火格子 摩耗測定・点検", ["乾燥火格子", "燃焼火格子", "後燃焼火格子"], 2),
    ("hyd", "火格子油圧装置 点検", ["火格子油圧ユニット", "油圧ポンプ", "油圧油冷却器"], 1),
    ("burn", "燃焼補機（落じん・バーナ・ファン）点検", ["落じんコンベヤ", "助燃バーナ", "再燃バーナ", "バーナ用燃焼ブロワ", "助燃油サービスタンク", "炉壁冷却ファン", "炉内ITV冷却ブロワ"], 1),
    ("boiler", "ボイラ 清掃・水管肉厚測定", ["廃熱ボイラ本体", "節炭器", "スートブロワ（長抜差式）", "スートブロワ（定置回転式）", "ブロータンク", "サンプリングクーラ", "ボイラダストコンベヤ", "主蒸気配管"], 3),
    ("bpump", "ボイラ給水・薬注ポンプ 点検", ["ボイラ給水ポンプ", "清缶剤注入ポンプ", "脱酸素剤注入ポンプ"], 1),
    ("fan", "送風機・誘引通風機 点検", ["押込送風機", "二次送風機", "蒸気式空気予熱器", "誘引通風機", "誘引通風機潤滑油ユニット"], 1),
    ("duct", "煙道・ダクト・煙突内筒 点検", ["燃焼空気ダクト", "二次空気ダクト", "燃焼空気吸込ダクト", "ボイラ出口煙道", "減温塔出口煙道", "BF出口煙道", "BFバイパス煙道", "誘引通風機出口煙道", "煙突内筒"], 1),
    ("gct", "減温塔 噴霧ノズル交換・内部点検", ["減温塔本体", "減温塔噴霧水ポンプ", "減温塔ダストコンベヤ"], 1),
    ("bf", "バグフィルタ・薬剤供給 点検", ["バグフィルタ本体", "BFホッパヒータ", "BF下部スクリューコンベヤ", "BFロータリフィーダ", "飛灰集合コンベヤ", "消石灰定量供給機", "消石灰吹込ブロワ", "活性炭定量供給機", "活性炭吹込ブロワ"], 2),
    ("ash", "灰押出機・主灰コンベヤ 点検", ["灰押出機", "主灰振動コンベヤ", "主灰コンベヤ"], 1),
    ("misc", "炉室ホイスト・冷却水ポンプ 点検", ["炉室電動ホイスト", "冷却水ブースタポンプ"], 1),
]
GROUPS_COMMON = [
    ("c_recv", "受入設備（計量機・扉・ダンピングボックス）点検", "A"),
    ("c_fuel", "灯油移送設備 点検", "B"),
    ("c_steam", "給水・蒸気・復水設備 点検", "C"),
    ("c_chem", "薬剤貯槽・集じん機 点検", "D"),
    ("c_heat", "余熱利用設備 点検", "E"),
    ("c_ash", "灰出し・飛灰処理設備 点検", "G"),
    ("c_water", "給水設備 点検", "H"),
    ("c_drain", "排水処理設備 点検", "J"),
    ("c_elec", "電気設備 外観点検", "K"),
    ("c_ctrl", "計装制御設備 点検", "L"),
    ("c_misc", "雑設備 点検", "M"),
]

PHR = {  # 機器名のキーワード → 点検の言い方（日報・報告書の文の種）
    "火格子": "片の摩耗・割れ・クリアランス測定",
    "焼却炉本体": "炉壁耐火物の目視・打音",
    "ホッパ": "内面摩耗・変形の目視",
    "ゲート": "開閉作動・油漏れ確認",
    "給じん": "シリンダ・摺動部の摩耗確認",
    "ブリッジ": "作動確認",
    "油圧ユニット": "作動油の汚れ・油量確認",
    "油圧ポンプ": "吐出圧・異音確認",
    "冷却器": "管内清掃・漏れ確認",
    "コンベヤ": "チェーン・ローラ・ケーシング点検",
    "バーナ": "ノズル清掃・着火確認",
    "ブロワ": "軸受・ベルト点検",
    "ファン": "軸受・羽根点検",
    "タンク": "外観・液面計確認",
    "ボイラ本体": "水管清掃・肉厚測定",
    "節炭器": "管外面清掃・腐食確認",
    "スートブロワ": "ランス・駆動部点検",
    "ポンプ": "軸封・軸受・吐出圧確認",
    "ブロータンク": "内面腐食確認",
    "クーラ": "漏れ確認",
    "配管": "保温外観・漏れ確認",
    "減温塔本体": "内部付着物・下部コーン腐食確認",
    "バグフィルタ": "ろ布抜取検査・差圧確認・リテーナ点検",
    "ヒータ": "絶縁抵抗測定",
    "フィーダ": "羽根・シール点検",
    "供給機": "定量性確認",
    "送風機": "軸受温度・振動・Vベルト点検",
    "通風機": "羽根車目視・振動測定・軸受点検",
    "予熱器": "フィン管の汚れ確認",
    "ダクト": "外面腐食・保温確認",
    "煙道": "内面腐食・ダスト堆積確認",
    "煙突": "内筒外面目視",
    "押出機": "ケーシング腐食・水封確認",
    "ホイスト": "ワイヤ・ブレーキ点検",
}
def phrase(name):
    for k, v in PHR.items():
        if k in name:
            return v
    return "外観・作動確認"

MARK_WORD = {"○": "点検", "◎": "分解整備", "△": "補修", "●": "取替"}


class Task:
    def __init__(self, no, name, team, unit, eqs, kind, plan, act, memo="", key=""):
        self.no, self.name, self.team, self.unit = no, name, team, unit
        self.eqs = eqs          # [(Equip, 記号)]
        self.kind = kind        # group key / special
        self.plan, self.act = plan, act
        self.memo = memo
        self.key = key

class Job:
    def __init__(self, cid):
        c = W.CON[cid]
        self.cid, self.c = cid, c
        self.fy = c["fy"]
        self.title = c["title"]
        self.sheets = []        # [(シート名, 期間(a,b), [Task])]
        self.days = []          # 日報の日 [dict]
        self.notes = []


def _place(tasks, start, rng, order_gap=None):
    """作業を稼働日に並べる（前の作業と半分重ねる）"""
    d = start
    out = []
    for t in tasks:
        n = t[3]
        days = []
        x = d
        while len(days) < n:
            if x.weekday() < 5:
                days.append(x)
            x += dt.timedelta(1)
        out.append(days)
        # 次の開始：今の作業の中ほど
        d = days[max(0, len(days) // 2)] if rng.random() < 0.6 else nxt(days[-1])
    return out


def build_T(y, rng):
    """定期整備業務委託の工程（号炉ごと）"""
    job = Job("T%d" % y)
    s = SHUT[y]
    for u in ("1", "2"):
        a, b = s["u%s" % u]
        cool_end = a + dt.timedelta(2)
        start = nxt(cool_end) if cool_end.weekday() < 5 else nxt(cool_end)
        tasks = []
        tasks.append(Task("1", "炉停止・冷却（運転委託）", "運転", u, [], "stop", [a, a + dt.timedelta(1), cool_end], [a, a + dt.timedelta(1), cool_end]))
        grp = []
        for key, nm, names, nd in GROUPS_UNIT:
            eqs = []
            for n in names:
                e = W.find(u, n)
                m = mark(e, y)
                if m == "※":
                    m = "○"   # 部品の購入の年も定修の点検はする
                if m in ("○", "◎", "△", "●") and not str(SPECIAL.get((e.no, y), "P")).startswith("R"):
                    eqs.append((e, m))
            if key == "clean" and not eqs:
                eqs = [(W.find(u, "焼却炉本体"), "○")]
            if not eqs:
                continue
            # ◎△● は別の行
            plain = [(e, m) for e, m in eqs if m == "○"]
            heavy = [(e, m) for e, m in eqs if m != "○"]
            if plain:
                grp.append((key, nm, plain, nd))
            for e, m in heavy:
                nd2 = 2 if m in ("◎", "△") else 3
                grp.append((key + "_" + e.no, "%s %s" % (e.name, MARK_WORD[m] if e.no != "2F-07" else "羽根車取替（R6 購入品）・軸受交換"), [(e, m)], nd2 if e.no != "2F-07" else 4))
        # 検査日：立上げの 2 稼働日前。作業はそれまでに収める（班は 2〜3 作業を並行）
        insp = b
        for _ in range(2):
            insp -= dt.timedelta(1)
            while insp.weekday() >= 5:
                insp -= dt.timedelta(1)
        if y == 2026 and u == "2":
            insp = D(2026, 11, 17)
        Wd = wdays(start, insp - dt.timedelta(1))
        # 同じ停止で炉内に入る別契約（補修工事）の終わり
        rep_end = None
        for (eno, yy), cid in SPECIAL.items():
            if yy == y and W.EQ[eno].unit == u and cid.startswith("R") and cid != "R2020b":
                rj = build_repair(cid)
                rep_end = max(max(t.act) for t in rj.sheets[0][2])
        refr_n = 3
        refr_days = [d for d in Wd if rep_end is None or d > rep_end][-refr_n:]
        Wm = [d for d in Wd if d < refr_days[0]] or Wd
        ng = len(grp)
        for i, g in enumerate(grp):
            nd = min(g[3], len(Wm))
            st = round(i * (len(Wm) - nd) / max(1, ng - 1))
            days = Wm[st:st + nd]
            act = list(days)
            memo = ""
            if g[0] in ("clean", "boiler", "bf", "grate") and rng.random() < 0.15:
                k = Wm.index(days[-1]) + 1
                if k < len(Wm):
                    act = days + [Wm[k]]
                    memo = rng.choice(["付着物が多く 1 日延長", "清掃に時間がかかり 1 日延長", "測定点を追加し 1 日延長"])
            tasks.append(Task(str(len(tasks) + 1), g[1], "A班", u, g[2], g[0], days, act, memo, key=g[0]))
        rmemo = "別契約（%s）の炉内作業の後" % ("補修工事") if rep_end else ""
        tasks.append(Task(str(len(tasks) + 1), "耐火物 点検・部分補修", "A班", u, [(W.find(u, "焼却炉本体"), "○")], "refr", refr_days, list(refr_days), rmemo))
        # 性能検査・復旧・立上げ
        tasks.append(Task(str(len(tasks) + 1), "ボイラ性能検査（受検）", "A班", u, [(W.find(u, "廃熱ボイラ本体"), "☆")], "binsp", [insp], [insp]))
        rest = [d for d in [insp + dt.timedelta(k) for k in range((b - insp).days + 1)] if d.weekday() < 5 or d == b]
        tasks.append(Task(str(len(tasks) + 1), "復旧・昇温・立上げ（運転委託）", "運転", u, [], "start", rest, list(rest)))
        if u == "2":
            # 共通設備（2号の停止中に）
            z1, z2 = s["zen"]
            cst = nxt(start, 2)
            for key, nm, cat in GROUPS_COMMON:
                eqs = []
                for e in W.EQUIP:
                    if e.unit == "共通" and e.cat == cat:
                        m = mark(e, y)
                        if m in ("○", "◎", "△", "●") and not str(SPECIAL.get((e.no, y), "P")).startswith("R") and e.name not in ("ごみクレーン", "灰クレーン"):
                            eqs.append((e, m))
                if key == "c_elec":
                    eqs = [(e, "○") for e in W.EQUIP if e.unit == "共通" and e.cat == "K" and e.name in ("受電用変圧器", "動力変圧器", "照明変圧器")] + [(e, m) for e, m in eqs if e.name not in ("受電用変圧器", "動力変圧器", "照明変圧器")]
                    tasks.append(Task(str(len(tasks) + 1), "全停電（受変電 年次点検、1号炉も停止）・" + nm, "電気", "共通", eqs, key, [z1, z2], [z1, z2], "電気保安法人（【架空】みずなら電設）と共同"))
                    continue
                if key == "c_heat":
                    he = W.find("共通", "温水熱交換器（温水プール用）")
                    if not any(e.no == he.no for e, m in eqs):
                        eqs = [(he, "☆")] + eqs
                if not eqs:
                    continue
                n = 1 + len(eqs) // 4
                days = []
                x = cst
                while len(days) < n:
                    if x.weekday() < 5 and not (z1 <= x <= z2):
                        days.append(x)
                    x += dt.timedelta(1)
                memo = "第一種圧力容器 性能検査（%s機を開放）" % ("常用" if y % 2 else "予備") if key == "c_heat" else ""
                tasks.append(Task(str(len(tasks) + 1), nm, "A班", "共通", eqs, key, days, list(days), memo))
                cst = days[-1] if rng.random() < 0.5 else nxt(days[-1])
            c1, c2 = s["crane"]
            ce = [(W.find("共通", "ごみクレーン"), mark(W.find("共通", "ごみクレーン"), y) or "○"), (W.find("共通", "灰クレーン"), mark(W.find("共通", "灰クレーン"), y) or "○")]
            cd = wdays(c1, c2)
            tasks.append(Task(str(len(tasks) + 1), "ごみ・灰クレーン 年次点検" + ("・性能検査（受検）" if y in CRANE_PERF else ""), "A班", "共通", ce, "crane", cd, list(cd),
                              {2020: "予備機 巻上ワイヤ・ブレーキ整備", 2023: "常用機 走行車輪取替（R4→R5 繰越）", 2022: "常用機 走行車輪は R5 へ繰越（予算）"}.get(y, "")))
        sheet_end = max(max(t.act) for t in tasks)
        job.sheets.append(("%s号炉" % u, (a, sheet_end), tasks))
    return job


# ---------------- 補修工事（請負）：作業を手で決める ----------------
def build_repair(cid):
    job = Job(cid)
    if cid == "R2023a":
        tasks = []
        for no, item, team, plan, act in G.TASKS:
            tasks.append(Task(no, item, team, "1", [(W.EQ["1B-03"], "●")], "grate_rep", plan, act,
                              {"3": "ボルト固着でガス切断。1 日延長", "4-1": "6/15 点検で偏摩耗。追加（支給品）", "7": "6/22 市立会 合格"}.get(no, "")))
        job.sheets.append(("1号炉火格子補修", (D(2023, 6, 8), D(2023, 6, 22)), tasks))
        return job
    spec = {
        "R2020a": ("1号炉", "1D-05", "●", [
            ("1", "準備工（足場・養生・安全教育）", [D(2020, 6, 8)], [D(2020, 6, 8)], ""),
            ("2", "ろ布・リテーナ 取外し（6 室 540 本）", wdays(D(2020, 6, 9), D(2020, 6, 11)), wdays(D(2020, 6, 9), D(2020, 6, 12)), "ダスト固着が多く 1 日延長"),
            ("3", "セルプレート・ホッパ内面 清掃・点検", [D(2020, 6, 12)], [D(2020, 6, 15)], "セル穴の変形なし"),
            ("4", "新ろ布 取付（540 本）・リテーナ再使用", wdays(D(2020, 6, 15), D(2020, 6, 17)), wdays(D(2020, 6, 15), D(2020, 6, 17)), "リテーナ曲り 14 本は手直し"),
            ("5", "パルス弁 作動確認・プレコート", [D(2020, 6, 18)], [D(2020, 6, 18)], ""),
            ("6", "市立会検査・片付け", [D(2020, 6, 19)], [D(2020, 6, 19)], "差圧 0.9kPa（通風試験）"),
        ]),
        "R2021a": ("2号炉", "2D-05", "●", [
            ("1", "準備工（足場・養生・安全教育）", [D(2021, 10, 11)], [D(2021, 10, 11)], ""),
            ("2", "ろ布・リテーナ 取外し（6 室 540 本）", wdays(D(2021, 10, 12), D(2021, 10, 14)), wdays(D(2021, 10, 12), D(2021, 10, 14)), ""),
            ("3", "セルプレート・ホッパ内面 清掃・点検", [D(2021, 10, 15)], [D(2021, 10, 15)], "C 室ホッパ内面に腐食、当て板は定修で"),
            ("4", "新ろ布 取付（540 本）", wdays(D(2021, 10, 18), D(2021, 10, 20)), wdays(D(2021, 10, 18), D(2021, 10, 20)), ""),
            ("5", "パルス弁 作動確認・プレコート", [D(2021, 10, 21)], [D(2021, 10, 21)], ""),
            ("6", "市立会検査・片付け", [D(2021, 10, 22)], [D(2021, 10, 22)], "差圧 0.9kPa"),
        ]),
        "R2022a": ("1号炉", "1C-01", "●", [
            ("1", "準備工（足場・養生）", [D(2022, 6, 9)], [D(2022, 6, 9)], ""),
            ("2", "第1パス前面 水管 肉厚再測定・取替範囲決定", [D(2022, 6, 10)], [D(2022, 6, 10)], "最小 2.3mm、6 本の予定を 8 本に（変更契約）"),
            ("3", "水管 切断・撤去（8 本）", wdays(D(2022, 6, 13), D(2022, 6, 14)), wdays(D(2022, 6, 13), D(2022, 6, 15)), "2 本追加の分 1 日延長"),
            ("4", "開先加工・新管 取付・溶接", wdays(D(2022, 6, 15), D(2022, 6, 17)), wdays(D(2022, 6, 16), D(2022, 6, 20)), "溶接工 2 名"),
            ("5", "溶接部 放射線透過試験（外注）", [D(2022, 6, 20)], [D(2022, 6, 21)], "全 16 継手 合格"),
            ("6", "水圧試験・市立会", [D(2022, 6, 21)], [D(2022, 6, 22)], "2.25MPa 10 分 漏れなし"),
            ("7", "保温復旧・片付け", [D(2022, 6, 22)], [D(2022, 6, 23), D(2022, 6, 24)], ""),
        ]),
        "R2024a": ("共通", "0G-08", "●", [
            ("1", "準備工・常用機 停止（予備機へ切替）", [D(2024, 11, 11)], [D(2024, 11, 11)], "予備機で運転継続"),
            ("2", "常用機 パドル取外し・軸点検", [D(2024, 11, 12)], [D(2024, 11, 12)], "軸の摩耗なし"),
            ("3", "新パドル取付（地元製作 二軸 24 枚）", wdays(D(2024, 11, 13), D(2024, 11, 14)), wdays(D(2024, 11, 13), D(2024, 11, 14)), ""),
            ("4", "試運転・市立会・片付け", [D(2024, 11, 15)], [D(2024, 11, 15)], "混練状態 良"),
        ]),
        "R2025a": ("2号炉", "2B-03", "●", [
            ("1", "準備工（資材搬入・養生）", [D(2025, 10, 9)], [D(2025, 10, 9)], ""),
            ("2", "火格子片 摩耗測定・取替片マーキング", [D(2025, 10, 10)], [D(2025, 10, 10)], "30 枚（予備 18＋1号取外し品の良品 12）"),
            ("3", "火格子片 取外し（30 枚）", wdays(D(2025, 10, 13), D(2025, 10, 14)), wdays(D(2025, 10, 13), D(2025, 10, 14)), ""),
            ("4", "火格子片 取付・ボルト締付", wdays(D(2025, 10, 15), D(2025, 10, 16)), wdays(D(2025, 10, 15), D(2025, 10, 16)), "再使用片は肉盛なしで使用"),
            ("5", "クリアランス調整・空作動試験・市立会", [D(2025, 10, 17)], [D(2025, 10, 17)], "合格"),
        ]),
    }[cid]
    sheet, eno, m, rows = spec
    tasks = [Task(no, nm, "B班", W.EQ[eno].unit, [(W.EQ[eno], m)], "rep_" + cid, p, a, memo) for no, nm, p, a, memo in rows]
    a0 = min(min(t.act) for t in tasks); a1 = max(max(t.act) for t in tasks)
    job.sheets.append((sheet + W.CON[cid]["short"].replace("1号炉", "").replace("2号炉", "").replace("1号", "").replace("2号", "")[:10], (a0, a1), tasks))
    return job


# ---------------- 毎日の作業（日報の元） ----------------
WEATHER = "晴晴晴曇曇雨"

def make_days(job, rng):
    """班ごとに、その日にした作業の行を作る"""
    days = {}
    for sname, (a, b), tasks in job.sheets:
        for t in tasks:
            if t.team in ("運転",):
                continue
            n = len(t.act)
            # 機器を日に割り振る
            eqs = list(t.eqs)
            for i, d in enumerate(sorted(t.act)):
                if job.fy == 2026 and d > TODAY:
                    continue
                part = eqs[i * len(eqs) // n:(i + 1) * len(eqs) // n] if eqs else []
                days.setdefault((t.team if t.team != "電気" else "A班", d), []).append((t, i, n, part))
    out = []
    for (team, d), items in sorted(days.items(), key=lambda x: x[0][1]):
        lines = []
        for t, i, n, part in items:
            lines += task_lines(job, t, i, n, part, rng)
        k = len(items)
        if team == "B班":
            crew = [("職長", 1), ("機械工", 2 + (1 if k > 1 or rng.random() < 0.5 else 0)), ("普通作業員", 1 + (rng.random() < 0.6))]
            if any("溶接" in t.name or "切断" in t.name for t, *_ in items):
                crew.insert(2, ("溶接工", 2 if job.cid == "R2022a" else 1))
        else:
            mech = min(5, 2 + k // 2 + (rng.random() < 0.4))
            crew = [("職長", 1), ("機械工", mech)]
            if any(t.kind in ("c_elec", "bf", "c_ctrl") for t, *_ in items) or rng.random() < 0.25:
                crew.append(("電工", 1))
            if any("補修" in t.name or "取替" in t.name or t.kind == "refr" for t, *_ in items):
                crew.append(("溶接工", 1))
            crew.append(("普通作業員", 1 + int(rng.random() < 0.5) + int(k > 3)))
        ot = 2 if rng.random() < 0.08 else 0
        out.append(dict(team=team, d=d, w=rng.choice(WEATHER), crew=crew, ot=ot, items=items, lines=lines))
    # R5 火格子補修は見本と同じ中身（r5grate）
    if job.cid == "R2023a":
        out = []
        for gd in G.DAYS:
            items = [(t, 0, 1, []) for t in job.sheets[0][2] if gd["d"] in t.act]
            out.append(dict(team="B班", d=gd["d"], w=gd["w"], crew=gd["crew"], ot=gd["ot"], items=items, lines=list(gd["work"]),
                            place=gd["place"], prog=gd["prog"], note=gd["note"]))
    job.days = out
    for dd in out:
        n = sum(k for _, k in dd["crew"])
        dd["md"] = n + (3 * dd["ot"] / 8.0 if dd["ot"] else 0)
    return out


def meas_line(job, e, y, rng):
    u = e.unit
    if e.name == "燃焼火格子":
        return "燃焼火格子 摩耗測定 240 枚 最大 %.1fmm" % MEAS[("grate", u)][y]
    if e.name == "廃熱ボイラ本体":
        return "ボイラ水管 肉厚測定 60 点 最小 %.2fmm（第1パス前面）" % MEAS[("boiler", u)][y] if y != 2022 or u != "1" else "ボイラ水管 肉厚測定 60 点（R4 取替部を除く）最小 %.2fmm" % MEAS[("boiler", u)][y]
    if e.name == "誘引通風機":
        return "誘引通風機 振動測定 %.1fmm/s" % MEAS[("idf", u)][y]
    if e.name == "バグフィルタ本体":
        return "BF ろ布抜取 各室 2 本・差圧 %.1fkPa" % MEAS[("bf", u)][y]
    return None


def task_lines(job, t, i, n, part, rng):
    y = job.fy
    L = []
    if t.kind.startswith("rep_"):
        import re
        mm = re.search(r"（[^）]*?(\d+) ?(本|枚|個)", t.name)
        base = t.name.split("（")[0]
        if i == 0:
            L.append(t.name)
        else:
            L.append(base + " つづき" + ("、完了" if i == n - 1 else ""))
        if mm and n > 1:
            N = int(mm.group(1)); done = round(N * (i + 1) / n)
            L.append("　進み：%d/%d %s" % (done, N, mm.group(2)))
        if t.memo and i == n - 1:
            L.append("　" + t.memo)
        return L
    if t.kind == "binsp":
        return ["%s号ボイラ 性能検査 受検（検査機関 立会）合格" % t.unit]
    if t.kind == "crane":
        if i == 0:
            return ["ごみクレーン（常用・予備）年次点検：ワイヤ・ブレーキ・走行車輪" + ("、" + t.memo if t.memo and "繰越" not in t.memo else "")]
        if i == n - 1 and y in CRANE_PERF:
            return ["クレーン性能検査 受検（ごみクレーン 2 基）合格、灰クレーン 年次点検"]
        return ["灰クレーン 年次点検・ごみクレーン 給電装置点検"]
    if t.kind == "c_elec":
        return ["全停電 %s：受変電 年次点検に立会、%s" % ("1 日目" if i == 0 else "2 日目", "・".join(e.name for e, m in part) + " 外観点検" if part else "復電後 各機器の起動確認")]
    if t.kind == "refr":
        return [["耐火物 打音点検（側壁上部）", "耐火物 部分補修（キャスタブル施工）", "耐火物 乾燥・養生、炉内片付け"][min(i, 2)]]
    if t.kind == "clean":
        return [["炉内清掃 クリンカ除去（燃焼段・後燃焼段）、入炉前 O2・CO 測定", "炉内清掃 つづき、落じんホッパ清掃"][min(i, 1)]]
    for e, m in part:
        ml = meas_line(job, e, y, rng)
        if m in ("◎", "△", "●"):
            word = {"◎": "分解整備", "△": "補修", "●": "取替"}[m]
            if e.no == "2F-07" and y == 2025:
                L.append(["誘引通風機 羽根車取外し", "新羽根車（R6 購入品）取付・芯出し", "軸受交換・潤滑油入替", "試運転 振動 1.6mm/s"][min(i, 3)])
                continue
            L.append("%s %s（%s）" % (e.name, word, ["分解・点検", "部品交換・組立", "試運転"][min(i, 2)] if n > 1 else "分解・部品交換・試運転"))
        elif ml:
            L.append(ml)
        else:
            L.append("%s %s" % (e.name, phrase(e.name)))
    if not part and n > 1:
        L.append(t.name + " つづき")
    return L


def build_all():
    rng = random.Random(1111)
    jobs = []
    for y in YEARS:
        jobs.append(build_T(y, rng))
        for c in W.CONTRACTS:
            if c["fy"] == y and c["id"].startswith("R") and c["id"] != "R2020b":
                jobs.append(build_repair(c["id"]))
    for j in jobs:
        make_days(j, rng)
    return jobs


def works_from(jobs):
    """実際の作業（機器 × 作業）。工数は日ごとの人工を、その日の作業の機器数で割り振る"""
    works = []
    idx = {}
    for j in jobs:
        for dd in j.days:
            tot = sum(max(1, len(p)) for t, i, n, p in dd["items"]) or 1
            for t, i, n, part in dd["items"]:
                share = dd["md"] * max(1, len(part)) / tot
                for e, m in (part or t.eqs[:1]):
                    attr = ATTR.get(m, "INSPECT")
                    k = (j.cid, e.no, attr, t.name, t.unit)
                    if k not in idx:
                        idx[k] = dict(id="W%04d" % (len(works) + 1), equip=e.no, attr=attr, start=dd["d"].isoformat(), end=dd["d"].isoformat(),
                                      manDays=0.0, vendor=W.LOCAL_S if t.team != "電気" else "みずなら電設", contract=j.cid, task=t.name)
                        works.append(idx[k])
                    w = idx[k]
                    w["end"] = max(w["end"], dd["d"].isoformat())
                    w["start"] = min(w["start"], dd["d"].isoformat())
                    w["manDays"] += share / max(1, len(part or t.eqs[:1]))
    for w in works:
        w["manDays"] = round(w["manDays"], 2)
    return works

if __name__ == "__main__":
    jobs = build_all()
    for j in jobs:
        print(j.cid, j.title, [ (s[0], len(s[2])) for s in j.sheets], "日報", len(j.days), "人工 %.1f" % sum(d["md"] for d in j.days))
    print(len(works_from(jobs)))
