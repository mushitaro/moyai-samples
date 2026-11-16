# -*- coding: utf-8 -*-
"""契約ごとの工程（予定・実績）、日々の作業（日報の行）、所見・測定（報告書）を作る。
 見本の 3 契約（R4 定修、R4 チェーン、H30 灰出し）は sched.py の手書きデータを使う。"""
import datetime as dt, random, zlib
from collections import OrderedDict, defaultdict
import world as W, sched as S, cw
from cw import D, wd_after, wd_before, first_wd, is_work, weekday_jp

R = random.Random(W.SEED + 21)

def weather(d):
    r = random.Random(d.toordinal() * 7 + 3).random()
    return "晴" if r < .5 else ("曇" if r < .8 else ("雨" if r < .95 else "晴のち曇"))

class Task:
    def __init__(s, group, name, team, crew, plan, act=None, works=(), ref=False):
        s.group, s.name, s.team, s.crew, s.plan = group, name, team, crew, plan
        s.act = act if act is not None else plan
        s.works = list(works); s.ref = ref
class Line:
    def __init__(s, team, gou, text, crew, hours, works=()):
        s.team, s.gou, s.text, s.crew, s.hours, s.works = team, gou, text, crew, hours, list(works)
class Day:
    def __init__(s, date, lines, memo="", wx=None):
        s.date, s.lines, s.memo = date, lines, memo
        s.wx = wx or weather(date)
class Sched:
    def __init__(s, c):
        s.c = c; s.tasks = []; s.days = []; s.findings = []; s.meas = {}; s.materials = []; s.htasks = []
        s.stop = {}; s.startup = {}; s.common = []; s.teiden = None; s.note = []; s.photos = []

SCHED = {}

# ---------------- 状態の推移（測定値） ----------------
def grate_min(f, fy):
    if f == 1:
        if fy < 2015: return {2012: 20, 2013: 23, 2014: 19}[fy]
        t = {2015: 38.0, 2016: 35.6, 2017: 33.2, 2018: 30.8, 2019: 28.4, 2020: 26.0, 2021: 23.6, 2022: 21.0, 2023: 23.8, 2024: 21.5, 2025: 23.9, 2026: 22.6}
        return t[fy]
    if fy < 2016: return {2012: 23, 2013: 22.5, 2014: 18, 2015: 25}[fy]
    t = {2016: 38.0, 2017: 36.2, 2018: 34.4, 2019: 32.6, 2020: 30.8, 2021: 29.0, 2022: 26.0, 2023: 25.0, 2024: 23.6, 2025: 21.8, 2026: 24.0}
    return t[fy]

def idf_vib(f, fy):
    if f == 1:
        t = {2012: 5.2, 2013: 5.9, 2014: 6.8, 2015: 1.5, 2016: 1.6, 2017: 1.7, 2018: 1.8, 2019: 1.9, 2020: 2.0, 2021: 2.6, 2022: 3.1, 2023: 3.5, 2024: 3.8, 2025: 4.2}
    else:
        t = {2012: 3.6, 2013: 3.9, 2014: 4.2, 2015: 4.6, 2016: 1.4, 2017: 1.4, 2018: 1.5, 2019: 1.5, 2020: 1.6, 2021: 1.7, 2022: 1.8, 2023: 1.9, 2024: 2.0, 2025: 2.1}
    return t.get(fy, 2.0)

def fdf_vib(f, fy):
    base = 1.8 if f == 1 else 2.0
    v = base + 0.03 * ((fy * 7 + f) % 5)
    if f == 2 and fy == 2014: v = 3.9
    if f == 2 and fy == 2017: v = 1.6  # 軸受取替後
    return round(v, 1)

def bag_age(f, fy):
    last = 2008
    for (ff, y) in ((1, 2015), (2, 2016), (1, 2022), (2, 2023)):
        if ff == f and y <= fy: last = y
    return fy - last

def bf_dp(f, fy):
    if fy == 2022: return 1.5 if f == 1 else 1.6
    return round(0.8 + 0.12 * bag_age(f, fy), 1)

def hopper(f, fy):
    """ガス冷却室 下部ホッパ 最小板厚（mm）。測った年だけ値を返す"""
    if f == 2 and fy == 2022: return 3.4
    if f == 1 and fy == 2022: return 4.6
    if f == 1 and fy == 2025: return 4.2
    if f == 2 and fy == 2024: return 5.6  # 当て板の上から
    return None

def oil_nas(f, fy):
    return 8 + ((fy + f) % 3 == 0)

# ---------------- メーカー定修 ----------------
A_CREW, B_CREW, E_CREW = 4, 3, 2

def mk_text(w):
    e = W.BYMID[w.mid]
    return "%s %s" % (e.name, w.desc)

def mk_schedule(c):
    fy = c.fy
    s = Sched(c)
    works = [w for w in c.works]
    only = cw.MK_ONLY.get(fy)
    furn = [only] if only else [1, 2]
    nc = cw.MK_NC.get(fy, 3) if not only else 1
    fw = {f: [w for w in works if W.BYMID[w.mid].furnace == f] for f in (1, 2)}
    cwk = [w for w in works if W.BYMID[w.mid].furnace == 0]
    planned = []  # 2026 は予定だけ（作業としては存在しない）
    if fy == 2026:
        for e in W.EQ:
            sym = W.T3PLAN[e.mid].get(2026)
            if sym and e.furnace and sym in ("整", "補", "更"):
                planned.append(e)
    def a_seq(f):
        seq = [("開放・炉内／ガス冷却室清掃", []), ("火格子 点検・摩耗測定", [])]
        ws = fw[f]
        gr = [w for w in ws if W.BYMID[w.mid].name in ("燃焼火格子", "乾燥火格子", "後燃焼火格子") and w.kind != "INSPECT"]
        bf = [w for w in ws if W.BYMID[w.mid].name == "バグフィルタ"]
        other = [w for w in ws if w not in gr and w not in bf and W.BYMID[w.mid].name not in ("誘引通風機", "押込送風機", "ガス冷却室", "火格子駆動用油圧ユニット", "燃焼制御盤", "BFパルス用空気槽", "燃焼火格子")]
        insp_fixed = [w for w in ws if w not in gr and w not in bf and w not in other]
        if gr:
            seq.append(("燃焼火格子片 取替", gr))
        oth = list(other)
        for e in planned:
            if e.furnace == f and e.name not in ("燃焼火格子",):
                oth.append(e)
        while oth:
            pair = oth[:2]; oth = oth[2:]
            seq.append(("PAIR", pair))
        seq.append(("ガス冷却室 内部点検・噴射ノズル点検", [w for w in insp_fixed if W.BYMID[w.mid].name == "ガス冷却室"]))
        seq.append(("送風機・誘引通風機 点検・振動測定、油圧ユニット 点検", [w for w in insp_fixed if W.BYMID[w.mid].name in ("誘引通風機", "押込送風機", "火格子駆動用油圧ユニット")]))
        seq.append(("復旧・閉鎖（町立会）", []))
        return seq, bf, [w for w in insp_fixed if W.BYMID[w.mid].name in ("燃焼制御盤", "BFパルス用空気槽")]
    # 日程
    seqs = {f: a_seq(f) for f in furn}
    f1 = furn[0]
    L1 = 1 + len(seqs[f1][0])
    if fy in cw.MK_COMMON_START:
        cs = cw.MK_COMMON_START[fy]
        start = wd_before(cs, L1)
    else:
        start = cw.MK_START[fy]
    day = [start]
    for i in range(80): day.append(wd_after(day[-1]))
    s.stop[f1] = start - dt.timedelta(days=1)
    lines = defaultdict(list)  # date -> [Line]
    memo = defaultdict(list)
    tasks = []
    def T(group, name, team, crew, d1, d2, works=(), ref=False):
        tasks.append(Task(group, name, team, crew, [(d1, d2)], works=works, ref=ref))
    gname = {1: "1号炉", 2: "2号炉"}
    def gou(f): return "%d号" % f if f else "共通"
    def run_furnace(f, d0idx, cool_idx):
        seq, bf, insp2 = seqs[f]
        T(gname[f], "停止（立下げ %d/%d 22:00）・冷却" % (s.stop[f].month, s.stop[f].day), "町", "", day[cool_idx], day[cool_idx])
        idx = d0idx
        b_from = None
        for (nm, ws) in seq:
            d = day[idx]
            if nm == "PAIR":
                items = ws
                for it in items:
                    if isinstance(it, W.Work):
                        txt = mk_text(it); wk = [it]
                    else:
                        txt = "%s 整備（予定）" % it.name; wk = []
                    T(gname[f], txt, "機械A", 2, d, d, wk)
                    lines[d].append(Line("機械A", gou(f), txt, 2 if len(items) > 1 else 4, 8, wk))
                    for w in wk: w.start = w.end = d
                idx += 1; continue
            crew = A_CREW
            T(gname[f], nm, "機械A", crew, d, d, ws)
            if nm.startswith("開放"):
                b_from = idx + 1
                lines[d].append(Line("機械A", gou(f), "点検口・マンホール開放、炉内クリンカ除去、ガス冷却室内部清掃", crew, 6))
                lines[d].append(Line("機械A", gou(f), "ガス冷却室内部 足場組立", crew, 2))
            elif nm.startswith("火格子"):
                gm = grate_min(f, fy)
                lines[d].append(Line("機械A", gou(f), "乾燥・燃焼・後燃焼火格子 外観点検、火格子片 摩耗測定（全数）", crew, 8))
                memo[d].append("%d号 燃焼火格子 最小残厚 %smm（限度 22mm）。" % (f, gm))
            elif nm.startswith("燃焼火格子片"):
                for w in ws: w.start = w.end = d
                n = "".join(ch for ch in ws[0].desc if ch.isdigit()) or "数"
                lines[d].append(Line("機械A", gou(f), "燃焼火格子片 %s枚取替（町保管予備品）、取付ボルト締付確認" % n, crew, 6, ws))
                lines[d].append(Line("機械A", gou(f), "火格子 手動作動確認", crew, 2))
            elif nm.startswith("ガス冷却室"):
                for w in ws: w.start = w.end = d
                hp = hopper(f, fy)
                t = "ガス冷却室 内部点検、噴射ノズル点検" + ("、下部ホッパ板厚測定（UT）" if hp else "")
                lines[d].append(Line("機械A", gou(f), t, crew, 6, ws))
                lines[d].append(Line("機械A", gou(f), "足場解体", crew, 2))
                if hp: memo[d].append("下部ホッパ 最小 %.1fmm（原厚 6.0mm）。" % hp)
                if (f == 1 and fy < 2015) or (f == 2 and fy < 2016):
                    memo[d].append("側壁耐火物 剥離・脱落あり。")
            elif nm.startswith("送風機"):
                for w in ws: w.start = w.end = d
                lines[d].append(Line("機械A", gou(f), "押込送風機・二次送風機・誘引通風機 点検、振動測定、軸受グリース交換", 3, 6, [w for w in ws if "通風" in W.BYMID[w.mid].name or "送風" in W.BYMID[w.mid].name]))
                lines[d].append(Line("機械A", gou(f), "火格子駆動用油圧ユニット 作動油" + ("サンプリング・交換" if fy % 2 == 0 else "サンプリング（交換なし）"), 1, 6, [w for w in ws if "油圧" in W.BYMID[w.mid].name]))
                lines[d].append(Line("機械A", gou(f), "片付け", 4, 2))
                ins = [w for w in insp2 if "燃焼制御盤" in W.BYMID[w.mid].name]
                T(gname[f], "燃焼制御盤 校正・ループ確認", "電計", E_CREW, d, d, ins)
                for w in ins: w.start = w.end = d
                lines[d].append(Line("電計", gou(f), "燃焼制御盤 調節計校正、ループ確認（O2・炉温・ガス冷却室出口温度）", E_CREW, 8, ins))
                memo[d].append("%d号 誘引通風機 %.1fmm/s（反負荷側 水平）。" % (f, idf_vib(f, fy)) if fy != 2026 else "")
            elif nm.startswith("復旧"):
                lines[d].append(Line("機械A", gou(f), "炉内最終確認、マンホール・点検口閉鎖、ボルト増締め", crew, 6))
                lines[d].append(Line("機械A", gou(f), "資機材移動・片付け", crew, 2))
                rec_day = d
            idx += 1
        # B 班（バグフィルタ）
        bi = b_from + 1
        if bf and any("全数" in w.desc for w in bf):
            d1, d3 = day[bi], day[bi + 2]
            T(gname[f], "バグフィルタ ろ布全数取替（144本）", "機械B", B_CREW, d1, d3, bf)
            for w in bf: w.start, w.end = d1, d3
            lines[day[bi]].append(Line("機械B", gou(f), "バグフィルタ 上部点検口開放、旧ろ布取外し（144本）、袋詰め", B_CREW, 8, bf))
            lines[day[bi + 1]].append(Line("機械B", gou(f), "リテーナ点検、新ろ布取付 96本", B_CREW, 8, bf))
            lines[day[bi + 2]].append(Line("機械B", gou(f), "新ろ布取付 残48本、取付完了確認、払落し作動確認", B_CREW, 6, bf))
            pd = day[bi + 2]
        else:
            d1 = day[bi]
            T(gname[f], "バグフィルタ 点検" + ("・破損ろ布取替" if bf else ""), "機械B", B_CREW, d1, d1, bf)
            for w in bf: w.start = w.end = d1
            t = "バグフィルタ 点検（差圧・ろ布外観）" + ("、" + bf[0].desc if bf else "")
            lines[d1].append(Line("機械B", gou(f), t, B_CREW, 6, bf))
            memo[d1].append("%d号 BF 差圧 %.1fkPa（運転時）。" % (f, bf_dp(f, fy)))
            pd = d1
        ps = [w for w in insp2 if "空気槽" in W.BYMID[w.mid].name]
        T(gname[f], "パルス用空気槽 自主検査", "機械B", 1, pd, pd, ps)
        for w in ps: w.start = w.end = pd
        lines[pd].append(Line("機械B", gou(f), "パルス用空気槽 定期自主検査（外観・安全弁作動）", 1, 2, ps))
        return rec_day
    # 1 炉目
    lines[day[0]].append(Line("SV", "共通", "工程打合せ（町）、作業指示", 1, 8))
    lines[day[0]].append(Line("機械A", "共通", "資機材搬入、新規入場者教育、KY", A_CREW, 6))
    lines[day[0]].append(Line("機械A", gou(f1), "炉内温度確認、換気用送風機設置", A_CREW, 2))
    T("共通", "資機材搬入・新規入場者教育", "全員", 8, day[0], day[0])
    rec1 = run_furnace(f1, 1, 0)
    ri = day.index(rec1)
    # 共通日（全炉停止）
    cdays = [day[ri + 1 + i] for i in range(nc)]
    s.common = cdays
    s.teiden = cdays[1] if len(cdays) > 1 else cdays[0]
    cm = [w for w in cwk]
    crane_ws = [w for w in cm if W.BYMID[w.mid].name == "ごみクレーン"]
    tank_ws = [w for w in cm if W.BYMID[w.mid].name in ("空気槽No.1", "空気槽No.2", "保守用ホイスト")]
    cmon = [w for w in cm if W.BYMID[w.mid].name in ("中央監視操作盤",)]
    rest = [w for w in cm if w not in crane_ws + tank_ws + cmon]
    T("共通設備", "ごみクレーン 年次自主検査" + ("・性能検査受検" if fy % 2 == 1 else ""), "機械A", 3, cdays[0], cdays[0], crane_ws)
    T("共通設備", "空気槽No.1・No.2 自主検査／保守用ホイスト 年次", "機械A", 1, cdays[0], cdays[0], tank_ws)
    lines[cdays[0]].append(Line("機械A", "共通", "ごみクレーン 年次自主検査（ワイヤロープ・ブレーキ・リミット・走行横行レール）" + ("、性能検査 受検立会" if fy % 2 == 1 else ""), 3, 6, crane_ws))
    lines[cdays[0]].append(Line("機械A", "共通", "空気槽No.1・No.2 定期自主検査、保守用ホイスト 年次自主検査", 1, 6, tank_ws))
    for w in crane_ws + tank_ws: w.start = w.end = cdays[0]
    T("共通設備", "〔町発注〕受変電 年次点検（停電 9:00〜14:00）", "別途", "", s.teiden, s.teiden, ref=True)
    T("共通設備", "中央監視盤・共通計装盤 停電時点検", "電計", E_CREW, s.teiden, s.teiden, cmon)
    lines[s.teiden].append(Line("電計", "共通", "中央監視操作盤・共通計装盤 停電時点検・清掃、端子増締め", E_CREW, 6, cmon))
    lines[s.teiden].append(Line("電計", "共通", "復電後 計装電源確認", E_CREW, 2))
    for w in cmon: w.start = w.end = s.teiden
    for i, w in enumerate(rest):
        d = cdays[min(i, len(cdays) - 1)]
        T("共通設備", mk_text(w), "機械A", 2, d, d, [w])
        lines[d].append(Line("機械A", "共通", mk_text(w), 2, 6, [w])); w.start = w.end = d
    for lc in [x for x in cw.contracts_of(fy, ("LOCAL", "CRUP")) if getattr(x, "in_common", False)]:
        T("共通設備", "〔町発注〕%s（%s）" % (lc.short, lc.vendor.replace("【架空】", "").replace("有限会社 ", "").replace("株式会社", "")), "別途", "", lc.site[0], lc.site[1], ref=True)
    s.startup[f1] = cdays[-1]
    T(gname[f1], "立上げ・試運転", "SV", 1, cdays[-1], cdays[-1])
    lines[cdays[-1]].append(Line("SV", gou(f1), "%d号炉 立上げ前点検・立上げ立会" % f1, 1, 8))
    # 2 炉目
    if len(furn) == 2:
        f2 = 2
        s.stop[f2] = rec1
        rec2 = run_furnace(f2, ri + 2, ri + 1)
        r2 = day.index(rec2)
        su = day[r2 + 1]
    else:
        su = day[ri + nc + 1]
        r2 = ri + nc
    s.startup[furn[-1] if len(furn) == 2 else f1] = s.startup.get(f1) if len(furn) == 1 else su
    if len(furn) == 2:
        T("2号炉", "立上げ・試運転", "SV", 1, su, su)
        lines[su].append(Line("SV", "2号", "2号炉 立上げ・運転状態確認", 1, 8))
    end = su if len(furn) == 2 else cdays[-1]
    T("共通", "撤収・資機材搬出", "全員", 8, end, end)
    lines[end].append(Line("機械A", "共通", "資機材搬出、清掃", 3, 6))
    # SV 毎日
    for d in sorted(lines):
        if not any(l.team == "SV" for l in lines[d]):
            lines[d].insert(0, Line("SV", lines[d][0].gou, "作業指示・安全巡視", 1, 8))
    if fy == 2026:
        pass
    s.site = (day[0], end)
    s.tasks = tasks
    s.days = [Day(d, lines[d], "".join(memo[d])) for d in sorted(lines)]
    for t in s.tasks: t.group = t.group
    # 停電日を受変電の年次点検の日にする
    for w in cw.LEGAL_WORKS:
        if w.cid == "ELEC%d" % fy and "年次" in w.desc:
            w.start = w.end = s.teiden
    mk_findings(s, furn)
    return s

def judge_idf(v): return "A" if v < 2.8 else ("B" if v < 4.5 else "C")

def mk_findings(s, furn):
    fy = s.c.fy
    rows = []
    def add(gou, sysn, eq, j, obs, act, rec, works=()):
        rows.append(dict(gou=gou, sys=sysn, eq=eq, judge=j, obs=obs, act=act, rec=rec, works=list(works)))
    for f in furn:
        g = "%d号" % f
        ws = [w for w in s.c.works if W.BYMID[w.mid].furnace == f]
        gm = grate_min(f, fy)
        gw = [w for w in ws if W.BYMID[w.mid].name == "燃焼火格子" and w.kind != "INSPECT"]
        if gw:
            add(g, "燃焼設備", "燃焼火格子", "C", "火格子片 摩耗限度（最小 %smm）" % gm if "割れ" not in gw[0].desc else "火格子片 割れ", gw[0].desc, "次回 摩耗測定", gw)
        else:
            add(g, "燃焼設備", "燃焼火格子", "A" if gm >= 26 else "B", "摩耗 限度内（最小 %smm）" % gm, "－", "次回 摩耗測定")
        old = (f == 1 and fy < 2015) or (f == 2 and fy < 2016)
        add(g, "燃焼設備", "乾燥火格子・後燃焼火格子", "B" if old else "A", "摩耗進行（竣工時のまま）" if old else "異常なし", "－", "長寿命化で更新" if old else "－")
        add(g, "燃焼設備", "火格子駆動用油圧ユニット", "A", "作動油 汚損度 NAS %d級" % oil_nas(f, fy), "作動油交換・フィルタ取替" if fy % 2 == 0 else "－", "－")
        hp = hopper(f, fy)
        if old:
            add(g, "燃焼ガス冷却設備", "ガス冷却室", "C", "側壁耐火物 剥離・脱落", "－", "長寿命化工事で更新")
        elif hp:
            j = "D" if hp < 4.0 else "B"
            add(g, "燃焼ガス冷却設備", "ガス冷却室", j, "下部ホッパ 最小%.1fmm（原厚6.0）" % hp, "－（町へ報告）" if j == "D" else "－", "早期に当て板補修" if j == "D" else "板厚測定を継続")
        else:
            add(g, "燃焼ガス冷却設備", "ガス冷却室", "A", "内部 異常なし", "噴射ノズル清掃", "－")
        bfw = [w for w in ws if W.BYMID[w.mid].name == "バグフィルタ"]
        age = bag_age(f, fy)
        if bfw and "全数" in bfw[0].desc:
            add(g, "排ガス処理設備", "バグフィルタ", "A", "ろ布全数取替後 異常なし", "ろ布144本取替", "－", bfw)
        elif bfw:
            add(g, "排ガス処理設備", "バグフィルタ", "C", "差圧 %.1fkPa、破損ろ布" % bf_dp(f, fy), bfw[0].desc, "ろ布全数取替", bfw)
        else:
            add(g, "排ガス処理設備", "バグフィルタ", "B" if age >= 6 else "A", "差圧 %.1fkPa（ろ布 %d年経過）" % (bf_dp(f, fy), age), "－", "ろ布全数取替の計画" if age >= 6 else "－")
        add(g, "排ガス処理設備", "BFパルス用空気槽", "A", "定期自主検査 異常なし", "－", "次年度 自主検査")
        add(g, "通風設備", "押込送風機・二次送風機", "A", "振動 %.1fmm/s" % fdf_vib(f, fy), "グリース交換", "－")
        v = idf_vib(f, fy)
        add(g, "通風設備", "誘引通風機", judge_idf(v), "振動 %.1fmm/s" % v, "グリース交換", "運転中測定 継続" if judge_idf(v) != "A" else "－")
        add(g, "計装設備", "燃焼制御盤", "A", "校正値 許容内", "調節計校正", "－")
        for w in ws:
            e = W.BYMID[w.mid]
            if e.name in ("燃焼火格子", "バグフィルタ", "誘引通風機", "押込送風機", "ガス冷却室", "火格子駆動用油圧ユニット", "燃焼制御盤", "BFパルス用空気槽", "乾燥火格子", "後燃焼火格子"):
                continue
            add(g, W.SYS[e.sys], e.name, "C" if w.kind != "INSPECT" else "B", "摩耗・劣化" if w.kind != "INSPECT" else "点検", w.desc, "－", [w])
    cw_ = [w for w in s.c.works if W.BYMID[w.mid].furnace == 0]
    crane_note = "異常なし" if fy < 2021 else ("横行レール継目 摩耗軽微。制御装置 部品供給終了" if fy < 2024 else "制御装置更新後 異常なし")
    add("共通", "受入供給設備", "ごみクレーン", "B" if 2021 <= fy < 2024 else "A", crane_note, "－", "制御装置の計画的更新" if 2021 <= fy < 2024 else "－",
        [w for w in cw_ if W.BYMID[w.mid].name == "ごみクレーン"])
    add("共通", "圧縮空気設備", "空気槽No.1・No.2", "A", "定期自主検査 異常なし", "－", "次年度 自主検査", [w for w in cw_ if "空気槽" in W.BYMID[w.mid].name])
    add("共通", "雑設備", "保守用ホイスト", "A", "年次自主検査 異常なし", "－", "－", [w for w in cw_ if "ホイスト" in W.BYMID[w.mid].name])
    add("共通", "計装設備", "中央監視操作盤・共通計装盤", "A", "停電時点検 異常なし", "清掃・増締め", "－", [w for w in cw_ if "中央監視" in W.BYMID[w.mid].name])
    for w in cw_:
        e = W.BYMID[w.mid]
        if e.name in ("ごみクレーン", "空気槽No.1", "空気槽No.2", "保守用ホイスト", "中央監視操作盤"): continue
        add("共通", W.SYS[e.sys], e.name, "C" if w.kind != "INSPECT" else "B", "点検" if w.kind == "INSPECT" else "摩耗・劣化", w.desc, "－", [w])
    s.findings = rows
    s.meas["grate"] = [(f, gm) for f in furn for gm in [grate_min(f, fy)]]
    s.meas["vib"] = [(f, fdf_vib(f, fy), idf_vib(f, fy - 1) if fy > 2012 else None, idf_vib(f, fy)) for f in furn]
    s.meas["dp"] = [(f, bf_dp(f, fy)) for f in furn]
    s.meas["hopper"] = [(f, hopper(f, fy)) for f in furn if hopper(f, fy)]
    mats = []
    for w in s.c.works:
        e = W.BYMID[w.mid]
        if e.name == "燃焼火格子" and w.kind != "INSPECT":
            n = "".join(ch for ch in w.desc if ch.isdigit()) or "1"
            mats.append(("火格子片", "燃焼火格子用", "%s枚" % n, "町保管予備品", "%d号炉 燃焼火格子" % e.furnace, [w]))
        elif e.name == "バグフィルタ" and "全数" in w.desc:
            mats.append(("ろ布", "PTFE φ150×4,500", "144本", "本業務で納入", "%d号炉 バグフィルタ" % e.furnace, [w]))
        elif e.name == "バグフィルタ":
            mats.append(("ろ布", "PTFE φ150×4,500", "2本", "町保管予備品", "%d号炉 バグフィルタ" % e.furnace, [w]))
        elif w.kind in ("OVERHAUL", "REPLACE") and e.name not in ("火格子駆動用油圧ユニット",):
            mats.append(("交換部品一式", e.name + "用", "一式", "本業務で納入", ("%d号炉 " % e.furnace if e.furnace else "") + e.name, [w]))
    if fy % 2 == 0:
        mats.append(("作動油", "ISO VG46", "%dL" % (200 * len(furn)), "本業務で納入", "火格子駆動用油圧ユニット", []))
    mats.append(("グリース", "リチウム系", "一式", "本業務で納入", "各回転機", []))
    s.materials = mats

def mk_fixed_2022(c):
    """見本（R4）：sched.py の手書きデータ"""
    s = Sched(c)
    for (grp, task, team, plan, act) in S.MK_TASKS:
        p = [(D(2022, 10, a), D(2022, 10, b)) for a, b in plan]
        ws = []
        for w in c.works:
            if w.start and any(a <= w.start <= b for a, b in p):
                e = W.BYMID[w.mid]
                key = {"燃焼火格子片": "燃焼火格子", "灰押出機": "灰押出機", "バグフィルタ": "バグフィルタ", "ガス冷却室": "ガス冷却室",
                       "パルス用空気槽": "BFパルス用空気槽", "送風機": "送風機", "油圧": "油圧", "燃焼制御盤": "燃焼制御盤",
                       "ごみクレーン": "ごみクレーン", "空気槽No.1": "空気槽", "中央監視": "中央監視", "火格子 点検": "燃焼火格子"}
                for k, v in key.items():
                    if task.startswith(k) or k in task:
                        if v in e.name or (v == "送風機" and "送風機" in e.name) or (v == "油圧" and "油圧" in e.name) or (v == "空気槽" and ("空気槽" in e.name or "ホイスト" in e.name)):
                            if (grp == "1号炉" and e.furnace == 1) or (grp == "2号炉" and e.furnace == 2) or (grp == "共通設備" and e.furnace == 0):
                                ws.append(w)
                        break
        s.tasks.append(Task(grp, task, team, {"機械A": 4, "機械B": 3, "電計": 2, "SV": 1, "町": "", "全員": 8, "別途": ""}[team], p, works=ws, ref=(team == "別途")))
    for d in sorted(S.MK_DAILY):
        wx, rows, memo = S.MK_DAILY[d]
        ls = []
        for (team, gou, txt, p, h) in rows:
            ws = [w for w in c.works if w.start and w.start <= d <= (w.end or w.start) and (
                ("%s号" % W.BYMID[w.mid].furnace == gou) or (gou == "共通" and W.BYMID[w.mid].furnace == 0)) and
                any(k in txt for k in [W.BYMID[w.mid].name.replace("No.1", "").replace("BF", "").replace("用", "")[:4]])]
            ls.append(Line(team, gou, txt, p, h, ws))
        s.days.append(Day(d, ls, memo, wx))
    s.site = (D(2022, 10, 11), D(2022, 10, 28))
    s.common = [D(2022, 10, 19), D(2022, 10, 20), D(2022, 10, 21)]
    s.teiden = D(2022, 10, 20)
    s.stop = {1: D(2022, 10, 10), 2: D(2022, 10, 18)}
    s.startup = {1: D(2022, 10, 21), 2: D(2022, 10, 28)}
    return s

# ---------------- 地元（浜風機工）の修繕 ----------------
TPL = {
 ("pump", "REPLACE"): [("KY・電源ロック（町立会）", 1), ("旧ポンプ撤去", 2), ("新ポンプ据付・芯出し", 3), ("配管・配線接続", 2), ("試運転・片付け", 1)],
 ("pump", "OVERHAUL"): [("KY・電源ロック（町立会）", 1), ("ポンプ分解・点検", 3), ("メカニカルシール・軸受取替", 2), ("組立・芯出し", 1), ("試運転・片付け", 1)],
 ("fan", "OVERHAUL"): [("KY・電源ロック", 1), ("カップリング切離し・軸受取替", 4), ("芯出し", 1.5), ("試運転・振動確認", 1), ("片付け", .5)],
 ("fan", "REPLACE"): [("KY・電源ロック", 1), ("旧送風機撤去", 3), ("新送風機据付・芯出し", 4), ("ダクト・配線接続", 3), ("試運転・振動確認", 2), ("片付け", 1)],
 ("hyd", "OVERHAUL"): [("KY・油抜き", 1), ("シリンダ取外し", 2), ("パッキン取替", 2), ("取付・エア抜き", 2), ("作動確認", 1)],
 ("hyd", "REPLACE"): [("KY・油抜き", 1), ("旧油圧ポンプ撤去", 2), ("新油圧ポンプ取付", 3), ("配管接続・エア抜き", 2), ("圧力調整・作動確認", 2), ("片付け", 1)],
 ("door", "OVERHAUL"): [("KY・油抜き（町と投入規制調整）", 1), ("シリンダ取外し", 2), ("パッキン取替", 2), ("取付・エア抜き", 2), ("開閉作動確認", 1)],
 ("struct", "OVERHAUL"): [("足場組立", 3), ("ケレン・板厚確認", 3), ("当て板溶接", 6), ("塗装", 2), ("足場解体・片付け", 2)],
 ("cooler", "OVERHAUL"): [("足場組立・換気", 3), ("ケレン・板厚確認", 3), ("当て板溶接", 8), ("溶接部確認・塗装", 2), ("足場解体・片付け", 2)],
 ("conv", "OVERHAUL"): [("KY・電源ロック", 1), ("カバー開放・清掃", 2), ("スクリュー取外し", 2), ("肉盛溶接・仕上げ", 6), ("組付け・試運転", 3), ("片付け", 1)],
 ("ashconv", "OVERHAUL"): [("KY・電源ロック（町立会）", 1), ("灰抜き・清掃", 3), ("部品取外し", 2), ("取替・組付け", 4), ("張り調整・試運転", 2), ("片付け", 1)],
 ("comp", "REPLACE"): [("材料搬入", 2), ("旧機撤去", 2), ("新機据付", 3), ("配管・配線接続", 3), ("試運転", 2), ("片付け", 1)],
 ("mixer", "OVERHAUL"): [("KY・電源ロック", 1), ("混練機開放・清掃", 3), ("パドル取外し", 3), ("肉盛溶接", 8), ("組付け・試運転", 4), ("片付け", 1)],
 ("heater", "REPLACE"): [("KY・電源ロック", 1), ("ヒータ取外し", 2), ("エレメント取付", 3), ("絶縁測定・通電確認", 1), ("片付け", 1)],
}
DEF_R = [("KY・電源ロック", 1), ("旧品撤去", 2), ("新品取付", 3), ("接続・調整", 1), ("試運転・片付け", 1)]
DEF_O = [("KY・停止確認", 1), ("分解・点検", 2), ("部品取替", 3), ("組立・試運転", 2)]
EXTRA = ["取付ボルト腐食のため取替（追加）", "フランジガスケット取替（追加）", "据付ベース手直し（追加）", "配線端子 締直し・絶縁確認（追加）"]
MATS = {"pump": ("ポンプ", "同等品"), "fan": ("軸受", "両側"), "hyd": ("パッキン一式", ""), "door": ("パッキン一式", ""), "comp": ("空気圧縮機", "スクリュー式 15kW"),
        "struct": ("当て板", "SS400 t6"), "cooler": ("当て板", "SS400 t6"), "conv": ("溶接棒", "肉盛用"), "mixer": ("溶接棒", "肉盛用"), "heater": ("ヒータエレメント", "15kW"),
        "ashconv": ("スクレーパ・軸受", ""), "rv": ("ロータリーバルブ", "同等品"), "damper": ("ダンパ駆動部", "同等品"), "blower": ("ブロワ", "同等品"), "misc": ("交換部品", "一式")}

def place_local(c, used):
    """現地の日を決める。炉の機器はその炉の定修の中、灰出しは全炉停止の日、その他は平日"""
    fy = c.fy
    es = [W.BYMID[w.mid] for w in c.works]
    hours = 0
    for w in c.works:
        e = W.BYMID[w.mid]
        hours += sum(h for _, h in TPL.get((e.cls, w.kind), DEF_R if w.kind == "REPLACE" else DEF_O))
    ndays = max(1, int(-(-hours // 8)))
    mk = SCHED.get("MK%d" % fy)
    r = random.Random(zlib.crc32(c.cid.encode()) % 10007)
    if mk and all(e.furnace for e in es) and es[0].furnace in mk.stop:
        f = es[0].furnace
        st = wd_after(mk.stop[f] if f != 2 else mk.stop[f], 3)
        d0 = st
    elif mk and any(e.cls == "ashconv" for e in es):
        d0 = mk.common[0]; c.in_common = True
    else:
        while True:
            mo = r.choice([5, 6, 7, 8, 9, 11, 12, 1, 2] if fy < 2026 else [5, 6, 8, 9])
            y = fy if mo >= 4 else fy + 1
            d0 = first_wd(D(y, mo, r.randint(3, 24)))
            if d0.weekday() == 5: d0 = wd_after(d0)
            if d0 not in used: break
    days = [d0]
    while len(days) < ndays:
        n = wd_after(days[-1])
        if n.weekday() == 5: n = wd_after(n)
        days.append(n)
    return days

def local_schedule(c, used):
    s = Sched(c)
    if c.key == "H30-ASH":
        return h30_fixed(c)
    if c.key == "R4-CH":
        return r4ch_fixed(c)
    days = c.site_days if hasattr(c, "site_days") else (place_local(c, used) if not hasattr(c, "site") or not c.site else
                                                         [c.site[0]] + ([c.site[1]] if c.site[1] != c.site[0] else []))
    for d in days: used.add(d)
    r = random.Random(zlib.crc32(c.cid.encode()) % 9973)
    seq = [("材料搬入・準備", 1, [])]
    for w in c.works:
        e = W.BYMID[w.mid]
        tp = TPL.get((e.cls, w.kind), DEF_R if w.kind == "REPLACE" else DEF_O)
        pre = (("%d号 " % e.furnace) if e.furnace else "") + e.name + " " if len(c.works) > 1 else ""
        part = w.desc.split("（")[1].rstrip("）") if w.desc.startswith("故障修理（") else w.desc.split("（")[0]
        for k_ in ("取替", "修理", "補修", "整備", "故障修理", "分解整備"):
            part = part.replace(k_, "")
        part = part.split("・")[0].strip("・ ") or "部品"
        for (nm, h) in tp:
            if nm.startswith("KY") and len(seq) > 2: continue
            if "部品" in nm and part != "部品": nm = nm.replace("部品", part)
            seq.append((pre + nm, h, [w]))
    seq[1:] = seq[1:]
    # 追加作業（ときどき）
    extra = None
    if r.random() < .22 and len(seq) > 3:
        k = r.randrange(2, len(seq) - 1)
        extra = (EXTRA[r.randrange(len(EXTRA))], 2, seq[k][2])
    # 時間割り：8〜12・13〜17
    slots = []
    for d in days:
        slots += [(d, h) for h in (8, 9, 10, 11, 13, 14, 15, 16)]
    def lay(items):
        out = []; i = 0
        for (nm, h, ws) in items:
            n = max(1, int(round(h)))
            sl = slots[i:i + n]
            if not sl:
                sl = [(slots[-1][0], 16)];
            i += n
            spans = []
            for (d, hh) in sl:
                if spans and spans[-1][0] == d and spans[-1][2] == hh: spans[-1] = (d, spans[-1][1], hh + 1)
                elif spans and spans[-1][0] == d and spans[-1][2] == 12 and hh == 13: spans[-1] = (d, spans[-1][1], 14)
                else: spans.append((d, hh, hh + 1))
            out.append((nm, ws, spans))
        return out, i
    plan, np_ = lay(seq)
    act_items = list(seq)
    if extra:
        k = [x[0] for x in seq].index(seq[k][0]) + 1 if False else None
        act_items.insert(len(seq) - 1, extra)
    # 実績の枠が足りなければ最終日を延長（18時まで）
    if sum(max(1, int(round(h))) for _, h, _ in act_items) > len(slots):
        slots.append((days[-1], 17))
    act, na = lay(act_items)
    pmap = {nm: sp for nm, _, sp in plan}
    s.htasks = []
    for (nm, ws, sp) in act:
        s.htasks.append(dict(name=nm, team="A", plan=pmap.get(nm, []), act=sp, works=ws, extra=(extra is not None and nm == extra[0])))
    # 日報
    crew = 3
    weld = any(W.BYMID[w.mid].cls in ("struct", "cooler", "conv", "mixer") for w in c.works)
    for d in days:
        items = [h for h in s.htasks if any(x[0] == d for x in h["act"])]
        hrs = sum((x[2] - x[1]) - (1 if x[1] < 12 < x[2] else 0) for h in items for x in h["act"] if x[0] == d)
        hrs = min(max(hrs, 1), 9)
        texts = [h["name"] for h in items]
        lines = [Line("A", "", "、".join(texts), crew, hrs, sorted({w for h in items for w in h["works"]}, key=lambda w: w.mid))]
        memo = ""
        for h in items:
            if h["extra"]:
                memo = h["name"].replace("（追加）", "") + "。町に確認のうえ追加で実施。"
        s.days.append(Day(d, lines, memo))
        s.days[-1].weld = 1 if weld else 0
        s.days[-1].t0 = min(x[1] for h in items for x in h["act"] if x[0] == d)
        s.days[-1].t1 = max(x[2] for h in items for x in h["act"] if x[0] == d)
    s.extra = extra
    s.site = (days[0], days[-1])
    for w in c.works:
        w.start, w.end = days[0], days[-1]
        w.mh = sum(l.crew * l.hours for dd in s.days for l in dd.lines if w in l.works) / max(1, len(c.works))
    # 契約の日付
    if not getattr(c, "cdate", None):
        r2 = random.Random(zlib.crc32(c.cid.encode()) % 7919)
        c.cdate = wd_before(days[0], r2.randint(8, 16))
        c.start = wd_after(c.cdate, r2.randint(2, 4))
        c.end = days[-1] + dt.timedelta(days=r2.randint(25, 40))
        c.kanryo = wd_after(days[-1], r2.randint(1, 3))
        c.kensa = wd_after(c.kanryo, r2.randint(2, 4))
        c.pay = c.kensa + dt.timedelta(days=r2.randint(20, 30))
    if not getattr(c, "start", None): c.start = c.cdate
    if not getattr(c, "end", None): c.end = c.pay
    c.site = s.site
    # 材料
    for w in c.works:
        e = W.BYMID[w.mid]
        nm, sp = MATS.get(e.cls, ("交換部品", "一式"))
        s.materials.append((nm, sp, "1式" if "一式" in nm else "1", e.name, [w]))
    if extra:
        s.materials.append(("ボルト・ガスケット類", "", "一式", "追加作業", extra[2]))
    return s

def h30_fixed(c):
    s = Sched(c)
    for (task, spans) in S.H30_TASKS:
        s.tasks.append(Task("", task, "A", "", spans, works=c.works if any(k in task for k in ("リンク", "スプロケット", "スクレーパ", "張り")) else [], ref=task.startswith("（")))
    for (d, wx, tm, items, crew, mach, ky) in S.H30_DAILY:
        hrs = sum(h for _, h in crew)
        dd = Day(d, [Line("A", "", "、".join(items), len(crew), hrs, c.works if d >= D(2018, 10, 24) else [])], "", wx)
        dd.items, dd.crew, dd.mach, dd.ky, dd.tm = items, crew, mach, ky, tm
        s.days.append(dd)
    s.site = (D(2018, 10, 23), D(2018, 10, 25))
    s.materials = [("チェーンリンク（ピン・ブシュ共）", "RF450 P150", "8リンク", "灰出しコンベヤNo.1", c.works),
                   ("スクレーパ", "幅400", "6枚", "灰出しコンベヤNo.1", c.works), ("溶接棒", "肉盛用", "一式", "スプロケット", c.works)]
    s.extra = None
    return s

def r4ch_fixed(c):
    s = Sched(c)
    for (task, team, plan, act) in S.CH_TASKS:
        s.htasks.append(dict(name=task, team=team, plan=[(D(2022, 10, d), a, b) for d, a, b in plan], act=[(D(2022, 10, d), a, b) for d, a, b in act],
                             works=c.works if team in ("A", "B", "全") else [], extra=task.startswith("〔追加")))
    for (d, wx, tm, txt, a, b, weld, hrs, prog, note) in S.CH_DAILY:
        dd = Day(d, [Line("A", "", txt, a + b + weld, hrs, c.works)], note, wx)
        dd.tm, dd.a, dd.b, dd.w, dd.hrs, dd.prog = tm, a, b, weld, hrs, prog
        s.days.append(dd)
    s.site = (D(2022, 10, 18), D(2022, 10, 21))
    s.extra = ("〔追加〕テール軸受取替", 4, c.works)
    s.materials = [("スクレーパ用ローラチェーン", "RF450 P150 2条", "38.1m×2", "", c.works), ("ドライブスプロケット", "RF450 用 N=10", "2枚", "", c.works),
                   ("テールスプロケット", "RF450 用 N=10", "2枚", "", c.works), ("軸受ユニット", "UCP212", "2個", "追加（変更契約）", c.works),
                   ("取付ボルト・ナット", "M16 SUS", "一式", "", c.works), ("グリース", "", "一式", "", c.works)]
    return s

# ---------------- 長寿命化改良工事 ----------------
EN_TPL = [  # (作業, 班, 人数, 開始日オフセット, 日数)
 ("停止・冷却", "町", "", 0, 2),
 ("仮設・足場組立", "足場", 4, 2, 5),
 ("既設火格子（乾燥・燃焼・後燃焼）撤去", "機械A", 6, 7, 8),
 ("ガス冷却室 内面耐火物撤去", "耐火", 5, 7, 9),
 ("誘引通風機 撤去", "機械B", 4, 8, 3),
 ("火格子 据付（乾燥・燃焼・後燃焼）", "機械A", 6, 15, 24),
 ("ガス冷却室 耐火物施工・噴射ノズル取替", "耐火", 5, 17, 22),
 ("誘引通風機 据付・芯出し", "機械B", 4, 18, 9),
 ("誘引通風機用インバータ盤 据付・ケーブル布設", "電計", 3, 25, 9),
 ("バグフィルタ ろ布全数取替（144本）", "機械B", 4, 33, 5),
 ("燃焼制御盤・炉別計装盤 更新", "電計", 3, 35, 18),
 ("配管・ダクト復旧", "機械A", 6, 40, 8),
 ("足場解体・清掃", "足場", 4, 49, 4),
 ("単体試運転", "全員", 8, 54, 5),
 ("乾燥焚き", "SV", 2, 60, 4),
 ("総合試運転・性能確認", "SV", 3, 64, 6),
]
def en_schedule(c, fy):
    s = Sched(c)
    f = 1 if fy == 2015 else 2
    start = D(fy, 9, 1)
    days = [start]
    for i in range(80): days.append(wd_after(days[-1]))
    s.stop[f] = start - dt.timedelta(days=1)
    lines = defaultdict(list); memo = defaultdict(list)
    ws_all = [w for w in c.works if w.fy == fy]
    def pick(*names):
        return [w for w in ws_all if W.BYMID[w.mid].name in names and W.BYMID[w.mid].furnace in (f, 0)]
    wmap = {"火格子": pick("乾燥火格子", "燃焼火格子", "後燃焼火格子"), "ガス冷却室": pick("ガス冷却室"), "誘引通風機 据付": pick("誘引通風機"),
            "インバータ": pick("誘引通風機用インバータ盤"), "ろ布": pick("バグフィルタ"), "燃焼制御盤": pick("燃焼制御盤", "炉別計装盤")}
    tpl = list(EN_TPL)
    if fy == 2016:
        tpl.insert(11, ("中央監視操作盤・データロガ 更新（1号炉運転中・仮設監視）", "電計", 3, 30, 20))
        wmap["中央監視"] = [w for w in ws_all if W.BYMID[w.mid].name in ("中央監視操作盤", "データ処理装置")]
    for (nm, team, crew, off, n) in tpl:
        d1, d2 = days[off], days[off + n - 1]
        ws = []
        for k, v in wmap.items():
            if k in nm and ("撤去" not in nm): ws = v
        s.tasks.append(Task("%d号炉" % f if "中央監視" not in nm else "共通", nm, team, crew, [(d1, d2)], works=ws))
        for w in ws:
            w.start = min(w.start or d1, d1) if w.start and w.start.year == fy else d1
            w.end = d2
        if team == "町": continue
        for i in range(n):
            d = days[off + i]
            pct = int(round(100 * (i + 1) / n))
            t = nm + ("（進捗 %d%%）" % pct if n > 2 else "")
            lines[d].append(Line(team, "%d号" % f if "中央監視" not in nm else "共通", t, crew, 8, ws))
    for d in lines:
        lines[d].insert(0, Line("SV", "%d号" % f, "工程管理・安全巡視、町と打合せ" if d.weekday() == 0 else "工程管理・安全巡視", 1, 8))
    # 記事
    r = random.Random(fy)
    notes = ["既設火格子 撤去完了。", "耐火物 撤去量 想定より多い（側壁上部）。", "新火格子 搬入（トラック2台）。", "ガス冷却室 耐火物 施工 50%。",
             "インバータ盤 受電。", "ろ布 取付完了。", "単体試運転 異常なし。", "乾燥焚き 開始（町運転員と交替制）。", "性能確認 排ガス測定（分析会社）。"]
    keys = sorted(lines)
    for i, nt in enumerate(notes):
        memo[keys[min(len(keys) - 1, int(i * len(keys) / len(notes)) + 2)]].append(nt)
    s.days = [Day(d, lines[d], "".join(memo[d])) for d in sorted(lines)]
    s.site = (days[0], days[69])
    s.startup[f] = days[69]
    # 所見（報告書）
    rows = []
    for (nm, team, crew, off, n) in tpl:
        if team in ("町",): continue
        ws = []
        for k, v in wmap.items():
            if k in nm and "撤去" not in nm: ws = v
        rows.append(dict(gou="%d号" % f if "中央監視" not in nm else "共通", sys="", eq=nm, judge="完了", obs="%d/%d〜%d/%d" % (days[off].month, days[off].day, days[off + n - 1].month, days[off + n - 1].day),
                         act="施工完了", rec="－", works=ws))
    s.findings = rows
    s.materials = [("火格子（乾燥・燃焼・後燃焼）", "階段式 一式", "1炉分", "新設", "%d号炉" % f, wmap["火格子"]),
                   ("耐火物", "キャスタブル・れんが", "一式", "新設", "%d号炉 ガス冷却室" % f, wmap["ガス冷却室"]),
                   ("誘引通風機", "11,000m³N/h×4.5kPa 37kW", "1台", "新設", "%d号炉" % f, wmap["誘引通風機 据付"]),
                   ("インバータ盤", "37kW用", "1面", "新設", "%d号炉" % f, wmap["インバータ"]),
                   ("ろ布", "PTFE φ150×4,500", "144本", "新設", "%d号炉 バグフィルタ" % f, wmap["ろ布"]),
                   ("燃焼制御盤・炉別計装盤", "PLC", "各1面", "新設", "%d号炉" % f, wmap["燃焼制御盤"])]
    if fy == 2016:
        s.materials.append(("中央監視操作盤・データロガ", "PLC＋監視PC", "一式", "新設", "共通", wmap["中央監視"]))
    return s

def crup_schedule(c):
    s = Sched(c)
    mk = SCHED["MK2024"]
    cd = mk.common
    tpl = [("準備・仮設（仮設盤設置）", 0, 1), ("既設操作盤・動力盤 撤去", 1, 1), ("新盤 据付", 2, 1), ("配線・接続", 2, 2), ("試運転・調整、荷重試験", 4, 1)]
    for (nm, off, n) in tpl:
        d1, d2 = cd[off], cd[min(len(cd) - 1, off + n - 1)]
        s.tasks.append(Task("ごみクレーン", nm, "クレーン", 3, [(d1, d2)], works=c.works))
        s.tasks.append(Task("ごみクレーン", nm + "（電気）", "電計", 2, [(d1, d2)], works=c.works)) if "配線" in nm or "試運転" in nm else None
    for i, d in enumerate(cd):
        ls = [Line("SV", "共通", "工程管理・安全巡視", 1, 8)]
        for t in s.tasks:
            if t.plan[0][0] <= d <= t.plan[0][1]:
                ls.append(Line(t.team, "共通", t.name, t.crew, 8, c.works))
        s.days.append(Day(d, ls, "全炉停止中（ごみ投入なし）。" if i == 0 else ("荷重試験 合格。" if i == len(cd) - 1 else "")))
    for w in c.works: w.start, w.end = cd[0], cd[-1]
    s.site = (cd[0], cd[-1])
    s.findings = [dict(gou="共通", sys="受入供給設備", eq="ごみクレーン操作盤", judge="完了", obs="PLC・インバータ更新", act="更新", rec="－", works=[c.works[0]]),
                  dict(gou="共通", sys="受入供給設備", eq="ごみクレーン動力盤", judge="完了", obs="インバータ更新", act="更新", rec="－", works=[c.works[1]]),
                  dict(gou="共通", sys="受入供給設備", eq="ごみクレーン", judge="A", obs="荷重試験 定格3.2t 異常なし", act="－", rec="－", works=[])]
    s.materials = [("クレーン操作盤", "PLC 椅子式", "1面", "新設", "ごみクレーン", [c.works[0]]), ("クレーン動力盤", "インバータ 走行・横行・巻上", "1面", "新設", "ごみクレーン", [c.works[1]])]
    return s

def build_all():
    used = set()
    for c in cw.CONTRACTS.values():
        if c.kind == "MK":
            SCHED[c.cid] = mk_fixed_2022(c) if c.fy == 2022 else mk_schedule(c)
            c.site = SCHED[c.cid].site
    for c in cw.CONTRACTS.values():
        if c.kind == "EN":
            SCHED["EN2015"] = en_schedule(c, 2015); SCHED["EN2016"] = en_schedule(c, 2016)
        if c.kind == "CRUP":
            SCHED[c.cid] = crup_schedule(c); c.site = SCHED[c.cid].site; c.in_common = True
    for c in cw.CONTRACTS.values():
        if c.kind == "LOCAL":
            if c.docs:
                SCHED[c.cid] = local_schedule(c, used)
            else:
                if not getattr(c, "site", None):
                    days = place_local(c, used)
                    c.site = (days[0], days[-1])
                    for w in c.works: w.start, w.end = days[0], days[-1]
                r2 = random.Random(zlib.crc32(c.cid.encode()) % 7919)
                if not getattr(c, "cdate", None):
                    c.cdate = wd_before(c.site[0], r2.randint(8, 16)); c.start = c.site[0]
                    c.kanryo = wd_after(c.site[1], r2.randint(1, 3)); c.kensa = wd_after(c.kanryo, 2); c.pay = c.kensa + dt.timedelta(days=24)
                if not getattr(c, "start", None): c.start = c.site[0]
                if not getattr(c, "end", None): c.end = c.pay
                for w in c.works:
                    if not w.start: w.start, w.end = c.site
    for c in cw.CONTRACTS.values():
        c.in_mk_ref = None
        if c.kind == "LOCAL" and c.docs:
            mk = SCHED.get("MK%d" % c.fy)
            if mk and mk.site[0] <= c.site[0] <= mk.site[1]:
                es = [W.BYMID[w.mid] for w in c.works]
                if getattr(c, "in_common", False):
                    sp = [(mk.common[0], mk.common[-1])]; nm = "（メーカー定期点検　全炉停止）"
                elif es[0].furnace in mk.stop:
                    f = es[0].furnace
                    sp = [(mk.stop[f] + dt.timedelta(days=1), mk.startup[f])]; nm = "（メーカー定期点検　%d号炉停止）" % f
                else:
                    continue
                c.in_mk_ref = (nm, sp, [], True)
    # 町発注の別途工事（全炉停止中）を定修の工程表へ載せ直す
    for fy in range(2012, 2027):
        mk = SCHED.get("MK%d" % fy)
        if not mk: continue
        for lc in cw.contracts_of(fy, ("LOCAL", "CRUP")):
            if getattr(lc, "in_common", False) and not any(t.ref and lc.short in t.name for t in mk.tasks):
                mk.tasks.append(Task("共通設備", "〔町発注〕%s（%s）" % (lc.short, lc.vendor.replace("【架空】", "").replace("有限会社 ", "").replace("株式会社", "").strip()), "別途", "", [lc.site], ref=True))
    # 作業の番号
    allw = sorted(W.WORKS + cw.LEGAL_WORKS, key=lambda w: (w.fy, w.start or D(w.fy, 4, 1), w.mid))
    for i, w in enumerate(allw, 1):
        w.wid = "W%04d" % i
    return SCHED

build_all()
