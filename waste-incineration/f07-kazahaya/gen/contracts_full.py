# -*- coding: utf-8 -*-
"""本作り：2015–2026 年度の全契約工事（正解の世界）。
各年度：2号炉定期整備工事（5〜6月）、1号炉定期整備工事（9〜11月）、全炉停止時共通設備点検整備工事（1号炉定修の期間中）。
2026 年度は 蒸気タービン精密点検工事（全停の期間に別契約）を加える。
見本で承認された 3 契約（2018 1号炉・2023 1号炉・2023 全停）は contracts.py の手書きの工程をそのまま使う。
ほかは world.actual() の作業を班ごとの工程に組み立てる（班は順番に 1 作業ずつ、日曜休工、2020 年度までは祝日も休工）。
"""
from datetime import date, timedelta
from world import ALL, BY_CODE, actual, TEISHU, h, cost_of
import contracts as SAMPLE
from contracts import CREWS, HOLI

D = date
TODAY = D(2026, 11, 13)
# 2015–2026 の祝日（休工判定用。2020 年度までの工程で使う）
for y in range(2015, 2027):
    for md in [(1, 1), (2, 11), (4, 29), (5, 3), (5, 4), (5, 5), (8, 11), (11, 3), (11, 23)]:
        HOLI.add(D(y, *md))
HOLI.update({D(2015, 9, 21), D(2015, 9, 22), D(2015, 9, 23), D(2015, 10, 12), D(2016, 9, 19), D(2016, 9, 22), D(2016, 10, 10),
             D(2017, 9, 18), D(2017, 10, 9), D(2019, 9, 16), D(2019, 9, 23), D(2019, 10, 14), D(2019, 10, 22), D(2020, 9, 21),
             D(2020, 9, 22), D(2021, 9, 20), D(2021, 9, 23), D(2022, 9, 19), D(2022, 9, 23), D(2022, 10, 10), D(2024, 9, 16),
             D(2024, 9, 23), D(2024, 10, 14), D(2025, 9, 15), D(2025, 9, 23), D(2025, 10, 13), D(2026, 9, 21), D(2026, 9, 22),
             D(2026, 9, 23), D(2026, 10, 12), D(2015, 7, 20), D(2016, 7, 18), D(2019, 5, 6), D(2020, 5, 6), D(2024, 5, 6),
             D(2025, 5, 6)})

CREWS.update({
    "溶接": ("【架空】かざはや溶接工業", "y1", 4),
    "機械D": ("【架空】水樹テクノサービス", "t4", 3),
})


def off(d, holi):
    return d.weekday() == 6 or (holi and d in HOLI)


def addw(d, n, holi):
    """作業日を n 日進める（d が作業日でなければ次の作業日から）。"""
    while off(d, holi):
        d += timedelta(days=1)
    k = 1
    while k < n:
        d += timedelta(days=1)
        if not off(d, holi):
            k += 1
    return d


def nextw(d, holi):
    d += timedelta(days=1)
    while off(d, holi):
        d += timedelta(days=1)
    return d


def prevw(d, holi, n=1):
    for _ in range(n):
        d -= timedelta(days=1)
        while off(d, holi):
            d -= timedelta(days=1)
    return d


def people(y):
    return dict(spc="sp1" if y <= 2019 else "sp2", foreman="q1" if y <= 2018 else "q2",
                zentei_writer="q1" if y <= 2018 else ("q2" if y <= 2020 else "q3"))


def cname(y, unit):
    t = {"1": "1号炉定期整備工事", "2": "2号炉定期整備工事", "0": "全炉停止時共通設備点検整備工事", "T": "蒸気タービン精密点検工事"}[unit]
    return f"{y}年度 かざはやエコプラザ {t}", t


# ---------------------------------------------------------------- 作業 → 工程の組（班・日数）
def group_of(e, t, w):
    """(組キー, 組の名前, 班, 標準日数, 炉内作業か, 優先)"""
    b = e["base"]
    if b in ("乾燥火格子", "燃焼火格子", "後燃焼火格子"):
        if t == "X":
            return ("grateX", f"{b} 撤去・新品取付・調整", "機械A", 10, True, 0)
        if t == "R":
            return ("grateR", "火格子 部分交換", "機械A", 4, True, 1)
        return ("grateI", "火格子 摩耗測定", "機械A", 1, True, 2)
    if b == "焼却炉本体":
        if t == "R":
            return ("refR", "耐火物 点検・部分補修", "耐火", 6, True, 0)
        return ("refI", "炉内耐火物 点検", "耐火", 1, True, 5)
    if b in ("給じん装置", "主灰押出装置", "ホッパブリッジ解除装置"):
        return ("slide", "給じん装置・灰押出装置 摺動部・ライナ補修", "機械A", 4, True, 3)
    if "油圧シリンダ" in b:
        return ("cyl", "油圧シリンダ パッキン交換", "機械A", 2, False, 6)
    if "油圧ポンプ" in b or "油圧ユニット" in b:
        return ("hyd", "油圧装置 整備（ポンプ・作動油）", "機械B", 3, False, 6)
    if b == "誘引通風機":
        return ("idf", "誘引通風機 " + w, "機械B", 10 if "更新" in w else 8, False, 1)
    if b in ("押込送風機", "二次送風機", "排ガス再循環送風機", "炉側壁冷却ファン", "バーナ燃焼用送風機", "炉内監視テレビ冷却ファン"):
        return ("fan", "送風機 整備（" + b + "）", "機械B", 4 if "更新" in w else 3, False, 4)
    if b in ("ボイラ本体",):
        return ("legal", "ボイラ定期事業者検査", "機械A", 1, True, 9)
    if b in ("蒸気ドラム", "水管壁（炉頂・放射室）", "蒸発管群", "1次過熱器", "2次過熱器", "3次過熱器", "エコノマイザ"):
        if t == "I":
            return ("thick", "ボイラ 肉厚測定（水管壁・過熱器・ECO・ドラム）", "検査", 4, True, 2)
        return ("blrR", "ボイラ " + b + " " + w, "溶接", 7, True, 1)
    if b in ("ボイラ給水ポンプ", "減温塔噴霧水ポンプ", "灰押出水封補給水ポンプ", "灯油移送ポンプ", "再加熱器ドレン回収ポンプ"):
        return ("pump_" + b, f"{b}{e['sfx']} 分解整備", "機械B", 5 if b == "ボイラ給水ポンプ" else 3, False, 5)
    if b in ("長抜差型スートブロワ", "定置型スートブロワ", "過熱器減温器", "ボイラダスト排出装置", "エコノマイザ槌打装置"):
        return ("sb", "スートブロワ・ボイラ付属 整備", "機械C", 2, True, 6)
    if b == "バグフィルタ本体":
        if t == "X":
            return ("bfX", "バグフィルタ ろ布全数交換（720本）", "機械C", 6, False, 0)
        return ("test", "ろ布抜取試験・触媒性能試験（抜取）", "分析", 1, False, 7)
    if b == "触媒反応塔":
        if t == "X":
            return ("catX", "触媒反応塔 " + w, "機械C", 4, False, 1)
        return ("test", "ろ布抜取試験・触媒性能試験（抜取）", "分析", 1, False, 7)
    if b.startswith("バグフィルタ") or b in ("減温塔噴霧ノズルユニット", "アンモニア水注入ポンプ", "減温塔ダスト排出装置", "消石灰定量供給機",
                                         "活性炭定量供給機", "消石灰吹込ブロワ", "活性炭吹込ブロワ", "分析計サンプリング装置", "アンモニア水噴霧装置"):
        return ("gas", "排ガス処理設備 整備（BF付属・減温塔・薬剤吹込）", "機械C", 3, False, 3)
    if b in ("減温塔", "煙突内筒"):
        return ("inner", "減温塔・煙突内筒 内面点検・肉厚測定", "検査", 2, True, 4)
    if b == "排ガス再加熱器":
        return ("rh", "排ガス再加熱器 " + w, "機械C", 6 if "更新" in w else 1, False, 2)
    if e["kind"] == "電動機":
        if e["kw"] and e["kw"] >= 200:
            return ("mHV", "誘引通風機電動機 " + w, "電計", 3, False, 2)
        return ("motor", "電動機 整備・絶縁診断", "電計", 2, False, 6)
    if e["kind"] == "計装":
        return ("inst", "計装 更新・交換（熱電対・検出器・分析計）", "電計", 3, False, 4)
    if e["kind"] == "盤":
        return ("panel", "盤 精密点検", "電計", 1, False, 8)
    if b in ("主灰振動コンベヤ", "飛灰搬送コンベヤ", "ボイラ灰搬送コンベヤ", "落じん排出コンベヤ", "主灰磁選機", "バグフィルタ飛灰コンベヤ",
             "ボイラダスト搬出コンベヤ"):
        return ("conv", "コンベヤ類 整備", "機械C", 3, False, 5)
    return ("misc", "その他機器 分解整備", "機械C", 2, False, 8)


def _shorten(names, n=3):
    names = list(dict.fromkeys(names))
    return "・".join(names[:n]) + ("ほか" if len(names) > n else "")


# ---------------------------------------------------------------- 炉別定修
def furnace_contract(u, y):
    S, E = TEISHU[(u, y)]
    holi = y <= 2020
    P = people(y)
    name, short = cname(y, u)
    works = []
    for e in ALL:
        if e["unit"] != u:
            continue
        for (t, w, note) in actual(e).get(y, []):
            works.append((e, t, w))
    groups = {}
    for (e, t, w) in works:
        g = group_of(e, t, w)
        k = g[0]
        if k not in groups:
            groups[k] = dict(key=k, name=g[1], crew=g[2], dur=g[3], inside=g[4], prio=g[5], works=[])
        groups[k]["works"].append((e["code"], t, w))
    legal = groups.pop("legal", None)
    # 組の名前と日数を中身から整える
    for k, g in groups.items():
        n = len(g["works"])
        nm = [BY_CODE[c]["name"].replace(" 電動機", "電動機") for c, _, _ in g["works"]]
        if k in ("slide", "cyl", "sb", "gas", "conv", "misc", "motor", "inst", "panel", "hyd"):
            g["name"] = g["name"].split("（")[0].split(" 整備")[0] + "（" + _shorten(nm) + "）" + ("" if k in ("motor", "inst", "panel") else " 整備")
            g["dur"] = min(g["dur"] + (n - 1) // 2, 6)
        if k == "grateX":
            g["dur"] = 10
    # 撤去と取付の 2 工程に分ける（火格子更新）
    tasks = []
    T = lambda name, codes, crew, ps, pe, works=(), kind="main": dict(name=name, codes=list(codes), crew=crew, ps=ps, pe=pe,
                                                                         as_=ps, ae=pe, works=list(works), kind=kind)
    tasks.append(T("炉立下げ・冷却", [], "SPC運転", S - timedelta(days=1), S + timedelta(days=1), kind="frame"))
    d_open = addw(S + timedelta(days=2), 1, holi)
    tasks.append(T("縁切り・安全措置・マンホール開放", [], "電計", d_open, d_open, kind="frame"))
    tasks.append(T("足場組立（炉内・ボイラ・BF）", [], "足場", d_open, addw(d_open, 3, holi), kind="frame"))
    clean_s = nextw(d_open, holi)
    clean_e = addw(clean_s, 3, holi)
    clean_works = groups.pop("refI", {"works": []})["works"]
    tasks.append(T("炉内清掃・クリンカ除去" + ("・耐火物点検" if clean_works else ""), [c for c, _, _ in clean_works] or [f"{u}NS-010"],
                   "耐火", clean_s, clean_e, clean_works, kind="frame"))
    # 毎回の定例作業（計画の機器行には無い：清掃・開放点検・給油）
    groups["r_boiler"] = dict(key="r_boiler", name="ボイラ 伝熱管 清掃（水洗）・目視点検", crew="機械C", dur=3, inside=True, prio=3, works=[])
    groups["r_bf"] = dict(key="r_bf", name="バグフィルタ 室別開放点検・ろ布目視", crew="機械C", dur=2, inside=False, prio=4, works=[])
    groups["r_rot"] = dict(key="r_rot", name="回転機 定例点検・給油（送風機・ポンプ類）", crew="機械B", dur=2, inside=False, prio=7, works=[])
    if y == 2015:  # 竣工 1 年目：瑕疵担保期間中の全体点検（メーカー立会）
        groups["warranty"] = dict(key="warranty", name="竣工1年目 全体点検（各部目視・増締め・メーカー立会）", crew="機械C", dur=5,
                                  inside=True, prio=0, works=[])
    # 主作業を班ごとに並べる（前から）
    free = {}
    first_in = nextw(clean_e, holi)
    first_out = nextw(d_open, holi)
    order = sorted(groups.values(), key=lambda g: (g["prio"], g["key"]))
    last_in = clean_e
    last_any = clean_e
    for g in order:
        crew = g["crew"]
        parts = [("撤去", 3), ("新品取付・調整", 7)] if g["key"] == "grateX" else [("", g["dur"])]
        for (lab, du) in parts:
            start = max(free.get(crew, first_out), first_in if g["inside"] else first_out)
            end = addw(start, du, holi)
            nm = g["name"] if not lab else g["name"].replace("撤去・新品取付・調整", lab)
            kind = "main" if g["works"] else "routine"
            tasks.append(T(nm, [c for c, _, _ in g["works"]], crew, start, end, g["works"] if (lab != "撤去") else [], kind))
            free[crew] = nextw(end, holi)
            if g["inside"]:
                last_in = max(last_in, end)
            last_any = max(last_any, end)
    # 終わりの段取り（前から続ける）
    end_tasks = []
    scaf_s = nextw(last_in, holi)
    scaf_e = addw(scaf_s, 3, holi)
    end_tasks.append(T("足場解体", [], "足場", scaf_s, scaf_e, kind="frame"))
    after = max(scaf_e, last_any)
    if legal:
        hyd = addw(scaf_s, 2, holi)
        insp = nextw(hyd, holi)
        end_tasks.append(T("ボイラ水張り・水圧試験", [f"{u}BO-010"], "機械A", hyd, hyd, kind="frame"))
        end_tasks.append(T("ボイラ定期事業者検査（立会）", [f"{u}BO-010"], "機械A", insp, insp, legal["works"], kind="main"))
        after = max(after, insp)
    unit_s = nextw(after, holi)
    if u == "1":  # 全停（2 号炉も停止）が終わってから 1 号炉を立ち上げる
        z0, z1 = TEISHU[("0", y)]
        if y == 2026:
            z1 = D(2026, 10, 27)
        if unit_s <= z1:
            unit_s = nextw(z1 if y not in (2018, 2022) else z0 + timedelta(days=3), holi)
    if (u, y) == ("1", 2026):  # タービン精密点検中のため、1 号炉の立上げは 11/8（蒸気は ACC へ）
        unit_s = max(unit_s, D(2026, 11, 5))
    unit_e = nextw(unit_s, holi)
    link = nextw(unit_e, holi)
    rise_s = link + timedelta(days=1)
    rise_e = rise_s + timedelta(days=1)
    d_done = nextw(rise_e, holi)
    end_tasks.append(T("復旧・単体試運転", [], "機械A", unit_s, unit_e, kind="frame"))
    end_tasks.append(T("連動試運転", [], "機械A", link, link, kind="frame"))
    end_tasks.append(T("炉立上げ（昇温・ごみ投入）", [], "SPC運転", rise_s, rise_e, kind="frame"))
    end_tasks.append(T("完了確認・片付け", [], "機械A", d_done, d_done, kind="frame"))
    E = d_done
    tasks += end_tasks
    tasks.sort(key=lambda t: (t["ps"], 0 if t["kind"] == "frame" else 1))
    for i, t in enumerate(tasks, 1):
        t["no"] = i
    c = dict(key=f"C{y}{u}", unit=u, fy=y, name=name, short=short, period=(S, E), stop=S - timedelta(days=1), holi=holi,
             tasks=tasks, events=[], **P, contract_no=f"KZ-C-{y}-{'021' if u == '1' else '011'}", changes=[])
    if y not in (2018, 2023):  # 見本の発注台帳（2018・2023 年度）に変更契約が無い契約は計画どおり
        apply_events(c)
    return c


# ---------------------------------------------------------------- 全停
ZENTEI_SYS = ("YN", "DE", "KS")
TURBINE_BASES = ("蒸気タービン", "タービン減速機", "発電機", "タービン調速装置", "主蒸気止め弁油筒")


def zentei_scope(e):
    return e["unit"] == "0" and (e["sys"] in ZENTEI_SYS or e["base"] == "機器冷却水ポンプ")


def zentei_contract(y, turbine_only=False):
    S, E = TEISHU[("0", y)]
    holi = y <= 2020
    P = people(y)
    if turbine_only:
        name, short = cname(y, "T")
    else:
        name, short = cname(y, "0")
        if y == 2026:
            E = D(2026, 10, 27)
    works = []
    for e in ALL:
        if not zentei_scope(e):
            continue
        is_t = e["base"] in TURBINE_BASES
        if y == 2026 and (is_t != turbine_only):
            continue
        if turbine_only and not is_t:
            continue
        for (t, w, note) in actual(e).get(y, []):
            works.append((e, t, w))
    tasks = []
    T = lambda name, codes, crew, ps, pe, works=(), kind="main": dict(name=name, codes=list(codes), crew=crew, ps=ps, pe=pe,
                                                                         as_=ps, ae=pe, works=list(works), kind=kind)
    big_turbine = y in (2018, 2022) or turbine_only
    if turbine_only:
        tasks.append(T("タービン停止・ターニング・冷却", [], "SPC運転", S - timedelta(days=1), S, kind="frame"))
        dis = nextw(S, holi)
        tw = [(e["code"], t, w) for (e, t, w) in works]
        tasks.append(T("蒸気タービン 車室開放・ロータ吊出し", ["0YN-010"], "機械D", dis, addw(dis, 5, holi), [], "main"))
        s2 = nextw(addw(dis, 5, holi), holi)
        tasks.append(T("ロータ・ノズル・軸受 精密点検（非破壊検査含む）", ["0YN-010"], "SV", s2, addw(s2, 8, holi), [x for x in tw if x[0] == "0YN-010"]))
        tasks.append(T("減速機・調速装置・主蒸気止め弁油筒 分解点検", ["0YN-020", "0YN-110", "0YN-120"], "機械D",
                       nextw(addw(dis, 5, holi), holi), addw(nextw(addw(dis, 5, holi), holi), 6, holi),
                       [x for x in tw if x[0] in ("0YN-020", "0YN-110", "0YN-120")]))
        tasks.append(T("発電機 分解点検", ["0YN-150"], "電計2", s2, addw(s2, 5, holi), [x for x in tw if x[0] == "0YN-150"]))
        s3 = nextw(addw(s2, 8, holi), holi)
        tasks.append(T("組立・芯出し", ["0YN-010", "0YN-020"], "機械D", s3, addw(s3, 5, holi)))
        s4 = nextw(addw(s3, 5, holi), holi)
        tasks.append(T("定期事業者検査（立会）・試運転・併入", ["0YN-010"], "機械D", s4, E))
        for i, t in enumerate(tasks, 1):
            t["no"] = i
        c = dict(key=f"C{y}T", unit="T", fy=y, name=name, short=short, period=(S, E), stop=S - timedelta(days=1), holi=holi,
                 tasks=tasks, events=[], **P, contract_no=f"KZ-C-{y}-031", changes=[])
        return c
    tasks.append(T("2号炉立下げ・全炉停止／タービン停止・ターニング", [], "SPC運転", S - timedelta(days=1), S, kind="frame"))
    tasks.append(T("停電作業準備（仮設電源・縁切り）", [], "電計2", S, S, kind="frame"))
    dp = S + timedelta(days=1)
    de = [x for x in works if x[0]["sys"] == "DE" or x[0]["base"] == "発電機盤"]
    rest = [x for x in works if x not in de]
    gen = [x for x in de if x[0]["base"] == "非常用発電機"]
    de2 = [x for x in de if x not in gen]
    if de2:
        tasks.append(T("受変電設備 年次点検（全停電）", sorted({x[0]["code"] for x in de2}), "受変電", dp, dp,
                       [(x[0]["code"], x[1], x[2]) for x in de2]))
    if gen:
        tasks.append(T("非常用発電機 年次点検（負荷試験）" + ("・機関オーバーホール" if any(x[1] == "O" for x in gen) else ""),
                       ["0DE-270"], "受変電", dp, dp if not any(x[1] == "O" for x in gen) else addw(dp, 4, holi),
                       [(x[0]["code"], x[1], x[2]) for x in gen]))
    free = {}
    w0 = nextw(dp, holi)
    lim = prevw(E, holi, 2)
    if big_turbine:
        tb = [x for x in rest if x[0]["base"] in TURBINE_BASES]
        rest = [x for x in rest if x not in tb]
        te = prevw(E, holi)
        tasks.append(T("蒸気タービン 定期事業者検査（開放点検）", sorted({x[0]["code"] for x in tb}), "SV", w0, te,
                       [(x[0]["code"], x[1], x[2]) for x in tb]))
    # その他：機器ごとに 1〜2 日、班 機械D／電計2 に順番に
    for (e, t, w) in sorted(rest, key=lambda x: (x[0]["kind"] == "電動機", x[0]["code"])):
        crew = "電計2" if e["kind"] in ("電動機", "盤", "計装") else ("洗浄" if e["base"] == "空冷式復水器" else "機械D")
        du = 3 if e["base"] == "空冷式復水器" else (2 if t in "OXR" else 1)
        if e["base"] == "脱気器":
            du = 4
        st = max(free.get(crew, w0), w0)
        en = addw(st, du, holi)
        if en > lim:
            st, en = w0, addw(w0, du, holi)
            crew = "機械C" if crew == "機械D" else crew
        label = f"{e['name']} {w}"
        if e["base"] == "脱気器":
            label = "脱気器 開放・性能検査受検・復旧"
        tasks.append(T(label, [e["code"]], crew, st, en, [(e["code"], t, w)]))
        free[crew] = nextw(en, holi)
    re_ = prevw(E, holi) if not big_turbine else prevw(E, holi)
    tasks.append(T("復旧・通電確認・単体試運転", [], "電計2", re_, re_, kind="frame"))
    tasks.append(T("2号炉立上げ・タービン起動" if not big_turbine else "タービン起動・併入（2号炉はバイパス運転から）", [], "SPC運転", E, E,
                   kind="frame"))
    tasks.sort(key=lambda t: (t["ps"], 0 if t["kind"] == "frame" else 1))
    for i, t in enumerate(tasks, 1):
        t["no"] = i
    c = dict(key=f"C{y}0", unit="0", fy=y, name=name, short=short, period=(S, E), stop=S - timedelta(days=1), holi=holi,
             tasks=tasks, events=[], **P, contract_no=f"KZ-C-{y}-022", changes=[])
    if y not in (2018, 2023):
        apply_events(c)
    return c


# ---------------------------------------------------------------- 計画外のこと（遅れ・手直し・追加）
def tlabel(t):
    n = t["name"].split("（")[0]
    return n if len(n) <= 18 else n[:18]


def apply_events(c):
    """契約ごとに 0〜2 件。実績の終わりを延ばし、日報・報告・発注（変更契約）に引き継ぐ。"""
    k = h(c["key"], 100)
    holi = c["holi"]
    c["changes"] = []
    mains = [t for t in c["tasks"] if t["kind"] == "main" and t["crew"] not in ("SPC運転", "分析", "受変電", "SV")]
    if not mains:
        return
    ref = [t for t in mains if t["crew"] == "耐火"]
    if ref and k % 3 == 0:
        t = ref[0]
        area = round(1.2 + (k % 7) * 0.3, 1)
        d = addw(t["ps"], 3, holi)
        t["ae"] = addw(t["pe"], 2, holi)
        t["event"] = ("追加", d, f"側壁下部に耐火物の損傷 {area}m² を追加補修（変更協議・承認、耐火 2名増員）")
        c["events"].append((d, f"耐火物 側壁下部 損傷 {area}m² 追加補修 → 変更協議・承認"))
        c["changes"].append(("耐火物 追加補修", area))
    cand = [t for t in mains if t["crew"] in ("機械A", "機械B", "機械C", "機械D") and t["pe"] > t["ps"]]
    if cand and k % 4 in (1, 2):
        t = cand[k % len(cand)]
        d = t["pe"]
        t["ae"] = addw(t["pe"], 2, holi)
        what = {"機械A": "クリアランス超過", "機械B": "組立後の振動値超過", "機械C": "気密試験で漏れ", "機械D": "芯出し不良"}[t["crew"]]
        t["event"] = ("手直し", d, f"{what} → 再調整（手直し 1日）")
        c["events"].append((d, f"{tlabel(t)} {what} → 手直し"))
    if len(cand) > 2 and k % 5 == 3:
        t = cand[(k // 5) % len(cand)]
        if "event" not in t:
            t["as_"] = addw(t["ps"], 2, holi)
            t["ae"] = addw(t["pe"], 2, holi)
            t["event"] = ("遅れ", t["ps"], "交換部品の入荷遅れ（2日）→ 着手を繰下げ")
            c["events"].append((t["ps"], f"{tlabel(t)} 部品入荷遅れ → 2日繰下げ"))
    c["events"].sort()


# ---------------------------------------------------------------- 見本 3 契約を同じ形に
def from_sample(sc, unit):
    tasks = []
    for (no, nm, code, crew, ps, pe, as_, ae) in sc["tasks"]:
        codes = []
        for part in code.split(","):
            part = part.strip()
            if not part:
                continue
            if "〜" in part:
                a, b = part.split("〜")
                pre = a[:4]
                n0, n1 = int(a[4:]), int(b[4:])
                codes += [f"{pre}{n:03d}" for n in range(n0, n1 + 1, 10) if f"{pre}{n:03d}" in BY_CODE]
            else:
                codes.append(part)
        tasks.append(dict(no=no, name=nm, codes=codes, crew=crew, ps=ps, pe=pe, as_=as_, ae=ae, works=[],
                          kind="frame" if not codes or crew == "SPC運転" else "main"))
    y = sc["fy"]
    S, E = sc["period"]
    c = dict(key=f"C{y}{unit}", unit=unit, fy=y, name=sc["name"], short=sc["short"], period=sc["period"], stop=sc["stop"],
             holi=y <= 2020, tasks=tasks, events=list(sc.get("events", [])), **people(y),
             contract_no=f"KZ-C-{y}-{'021' if unit == '1' else '022'}", sample=sc["key"], changes=[])
    if sc["key"] == "C23":
        c["changes"] = [("耐火物 追加補修", 2.4)]
        for t in tasks:
            if t["no"] == 8:
                t["event"] = ("追加", D(2023, 10, 11), "燃焼段 左右側壁下部 2.4m² 追加補修（変更承認 10/12、耐火 2名増員）")
            if t["no"] == 7:
                t["event"] = ("手直し", D(2023, 10, 19), "クリアランス超過 3列（6・7・8列目）→ 手直し 10/19–20")
            if t["no"] == 17:
                t["event"] = ("遅れ", D(2023, 10, 24), "水銀計 校正ガス不良 → 10/25 再校正")
    if sc["key"] == "C23C":
        for t in tasks:
            if t["no"] == 8:
                t["event"] = ("手直し", D(2023, 10, 20), "減速機3 組立後の軸端振れ 0.08mm → 10/21 芯出し再調整")
    if sc["key"] == "C18":
        for t in tasks:
            if t["no"] == 5:
                t["event"] = ("追加", D(2018, 9, 7), "燃焼段 可動火格子の摩耗大（最大8.6mm）→ 測定点追加で1日延長")
    # 作業を割り当てる：その契約の対象（号炉・年度）の actual のうち、工程の対象ｺｰﾄﾞにあるもの
    for e in ALL:
        if unit == "0":
            if not zentei_scope(e):
                continue
        elif e["unit"] != unit:
            continue
        for (t, w, note) in actual(e).get(y, []):
            for tk in tasks:
                if e["code"] in tk["codes"] and tk["kind"] == "main":
                    tk["works"].append((e["code"], t, w))
                    break
    return c


def build_all():
    cs = []
    for y in range(2015, 2027):
        for u in ("2", "1"):
            if (y, u) == (2018, "1"):
                cs.append(from_sample(SAMPLE.C18, "1"))
            elif (y, u) == (2023, "1"):
                cs.append(from_sample(SAMPLE.C23, "1"))
            else:
                cs.append(furnace_contract(u, y))
        if y == 2023:
            cs.append(from_sample(SAMPLE.C23C, "0"))
        else:
            cs.append(zentei_contract(y))
        if y == 2026:
            cs.append(zentei_contract(y, turbine_only=True))
    import world
    for c in cs:
        world.TEISHU[(c["unit"], c["fy"])] = c["period"]
    for c in cs:
        S, E = c["period"]
        c["status"] = "完了" if E < D(2026, 7, 1) else ("完了（報告書未提出）" if E < TODAY else "施工中")
        c["cut"] = TODAY if c["status"] == "施工中" else None   # 施工中は 11/13 までの実績
    return cs


CONTRACTS = None


def get():
    global CONTRACTS
    if CONTRACTS is None:
        CONTRACTS = build_all()
    return CONTRACTS


if __name__ == "__main__":
    for c in get():
        nwork = sum(len(t["works"]) for t in c["tasks"])
        print(c["key"], c["name"], c["period"], len(c["tasks"]), nwork, c["status"], c["events"][:2])
