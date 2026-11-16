# -*- coding: utf-8 -*-
"""契約と工程（正解の世界の続き）。
 - 契約：メーカー定修委託（毎年）、長寿命化改良工事（H27〜28）、クレーン制御装置更新（R6）、地元の修繕、物品、保守委託・役務。
 - 工程：契約ごとに、工程（予定・実績）と日々の作業（日報の行）と所見（報告書）を持つ。
 工程表 → 日報 → 報告書 → 管理表 はここから書く。"""
import datetime as dt, random
from collections import OrderedDict, defaultdict
import world as W, sched as S

D = dt.date
RND = random.Random(W.SEED + 11)

def is_work(d): return d.weekday() != 6  # 日曜休み
def wd_after(d, n=1):
    while n > 0:
        d += dt.timedelta(days=1)
        if is_work(d): n -= 1
    return d
def wd_before(d, n=1):
    while n > 0:
        d -= dt.timedelta(days=1)
        if is_work(d): n -= 1
    return d
def first_wd(d):
    while not is_work(d) or d.weekday() == 5: d += dt.timedelta(days=1)
    return d
def weekday_jp(d): return "月火水木金土日"[d.weekday()]

def tax(fy): return 1.05 if fy <= 2013 else (1.08 if fy <= 2018 else 1.10)
def r1000(x): return int(round(x / 1000.0)) * 1000

class C:
    def __init__(s, **k):
        s.__dict__.update(dict(works=[], site=None, koutei=False, nippo=False, report=False, sched=None,
                               chg=None, purchase=None, note="", short=None, mitsu="2", subcontract=None))
        s.__dict__.update(k)
    def __repr__(s): return "<C %s %s>" % (s.cid, s.short or s.name)

CONTRACTS = OrderedDict()
def add(c):
    CONTRACTS[c.cid] = c; return c

def wl(fy): return W.wareki_long(fy)
FULL = lambda fy, nm: "%s しおさい町ごみ処理場 %s" % (wl(fy), nm)

# ============ メーカー定修 ============
MK_START = {2012: D(2012,10,9), 2013: D(2013,10,8), 2014: D(2014,10,14), 2015: D(2015,12,1), 2016: D(2016,12,6),
            2017: D(2017,10,10), 2019: D(2019,10,8), 2020: D(2020,10,13), 2021: D(2021,10,12),
            2023: D(2023,10,10), 2024: D(2024,10,8), 2025: D(2025,10,14), 2026: D(2026,10,13)}
MK_COMMON_START = {2018: D(2018,10,22)}
MK_NC = {2014: 1, 2018: 5, 2024: 5}
MK_ONLY = {2015: 2, 2016: 1}   # 長寿命化の年は片炉だけ（もう片方は改良工事）
def mk_amount(fy):
    if fy == 2017: return 16740000
    if fy == 2022: return 18150000
    base = 15500000 + (fy - 2017) * 200000
    if fy in (2015, 2016): base -= 4200000
    return r1000(base * tax(fy))

for fy in range(2012, 2027):
    add(C(cid="MK%d" % fy, fy=fy, name=FULL(fy, "定期点検整備業務委託"), short="%s 定期点検整備業務委託" % W.wareki(fy),
          vendor=W.MAKER, ctype="委託", hi="委託料", mitsu="特命", amount=mk_amount(fy),
          cdate=first_wd(D(fy, 7, 1)), start=first_wd(D(fy, 7, 1)), end=D(fy + 1, 1, 31),
          koutei=True, nippo=(fy <= 2025), report=(fy <= 2025), kind="MK"))
add(C(cid="EN", fy=2015, name="しおさい町ごみ処理場 長寿命化改良工事（平成27〜28年度）", short="長寿命化改良工事（H27〜28 継続費）",
      vendor=W.MAKER, ctype="請負", hi="工事請負費", mitsu="指名3", amount=318600000, split={2015: 152280000, 2016: 166320000},
      cdate=D(2015, 6, 15), start=D(2015, 6, 15), end=D(2017, 3, 17), koutei=True, nippo=True, report=True, kind="EN"))
add(C(cid="CRUP", fy=2024, name=FULL(2024, "ごみクレーン制御装置更新工事"), short="ごみクレーン制御装置更新工事",
      vendor=W.MAKER, ctype="請負", hi="工事請負費", mitsu="特命", amount=27280000, subcontract=W.CRANE,
      cdate=D(2024, 6, 10), start=D(2024, 6, 10), end=D(2025, 1, 31), koutei=True, nippo=True, report=True, kind="CRUP"))

# 作業の契約キー → 契約
KEYMAP = {"H27-28-EN": "EN", "R6-CRUP": "CRUP"}
for w in W.WORKS:
    k = w.contract
    if k in KEYMAP: w.cid = KEYMAP[k]
    elif k.endswith("-MK"): w.cid = "MK%d" % w.fy
    else: w.cid = None

# ============ 地元の修繕（年 3〜6 件） ============
LOCAL_VENDORS = (W.LOCAL, W.SETSUBI, W.CRANE, W.ELEC_W)
ACT = {"REPLACE": "取替修繕", "OVERHAUL": "修繕", "INSPECT": "点検", "UPGRADE": "改造"}

def eqlabel(e):
    return ("%d号炉" % e.furnace if e.furnace else "") + e.name

def short_name(ws):
    es = [W.BYMID[w.mid] for w in ws]
    if len(ws) == 1:
        e = es[0]; w = ws[0]
        d = w.desc
        if w.src == "hand":
            d2 = d.split("（")[0].replace(" ", "").replace(e.name, "")
            if w.kind == "REPLACE" and d2.startswith(("ポンプ", "水中ポンプ", "機器", "空調機", "空気圧縮機", "ブロワ", "装置")):
                return eqlabel(e) + ("更新修繕" if "更新" in d2 else "取替修繕")
            return eqlabel(e) + d2 + ("" if d2.endswith(("修繕", "補修", "整備")) else "修繕")
        if "故障修理" in d: return eqlabel(e) + "修繕"
        if w.kind == "REPLACE": return eqlabel(e) + "取替修繕"
        tail = {"pump": "分解整備", "fan": "軸受取替修繕", "conv": "補修", "hyd": "整備", "door": "油圧シリンダ整備"}.get(e.cls, "修繕")
        return eqlabel(e) + tail
    bases = {e.name.rstrip("AB").replace("No.1", "").replace("No.2", "") for e in es}
    nm = bases.pop() if len(bases) == 1 else {"pump": "ポンプ類", "fan": "送風機類", "blower": "ブロワ類", "comp": "空気圧縮機類"}.get(es[0].cls, "機器類")
    fs = sorted(set(e.furnace for e in es))
    pre = "1・2号炉" if fs == [1, 2] else (("%d号炉" % fs[0]) if fs[0] else "")
    return pre + nm + ("取替修繕（%d台）" % len(ws) if ws[0].kind == "REPLACE" else "修繕（%d台）" % len(ws))

SAMPLE_LOCAL = {  # 見本と同じ名前・日付・金額にする
 "H30-ASH": dict(short="灰出しコンベヤ補修工事", amount=1188000, cdate=D(2018,10,15), start=D(2018,10,15), end=D(2018,11,30),
                 site=(D(2018,10,23), D(2018,10,25)), kanryo=D(2018,10,29), kensa=D(2018,11,7), pay=D(2018,11,30)),
 "R4-CH": dict(short="灰出しコンベヤチェーン取替修繕", amount=4180000, chg=4345000, cdate=D(2022,9,5), start=D(2022,9,12), end=D(2022,11,30),
               site=(D(2022,10,18), D(2022,10,21)), kanryo=D(2022,10,27), kensa=D(2022,10,31), pay=D(2022,11,28)),
 "R4-PA": dict(short="1号炉噴射水ポンプAメカニカルシール取替修繕", amount=286000, cdate=D(2022,6,8), site=(D(2022,6,14), D(2022,6,14)), kanryo=D(2022,6,15), kensa=D(2022,6,16), pay=D(2022,7,8)),
 "R4-WA": dict(short="洗車装置高圧洗浄ポンプ取替修繕", amount=132000, cdate=D(2022,8,10), site=(D(2022,8,23), D(2022,8,23)), kanryo=D(2022,8,24), kensa=D(2022,8,26), pay=D(2022,9,26)),
 "R4-DR": dict(short="プラットホーム投入扉No.3油圧シリンダパッキン取替修繕", amount=198000, cdate=D(2023,1,6), site=(D(2023,1,17), D(2023,1,17)), kanryo=D(2023,1,18), kensa=D(2023,1,20), pay=D(2023,2,17)),
 "R4-JO": dict(short="浄化槽ブロワA取替修繕", amount=88000, cdate=D(2023,2,24), site=(D(2023,3,7), D(2023,3,7)), kanryo=D(2023,3,8), kensa=D(2023,3,10), pay=D(2023,3,31)),
 "H29-地元1": dict(short="2号炉押込送風機軸受取替修繕", amount=345600, cdate=D(2017,6,14), site=(D(2017,6,20), D(2017,6,20)), kanryo=D(2017,6,21), kensa=D(2017,6,23), pay=D(2017,7,14)),
 "H29-地元3": dict(short="混練機パドル肉盛補修", amount=410400, cdate=D(2017,7,25), site=(D(2017,8,8), D(2017,8,10)), kanryo=D(2017,8,14), kensa=D(2017,8,16), pay=D(2017,9,8)),
 "H29-VCB": dict(short="受変電設備真空遮断器取替修繕", amount=1620000, cdate=D(2017,8,28), site=(D(2017,10,19), D(2017,10,19)), kanryo=D(2017,10,23), kensa=D(2017,10,25), pay=D(2017,11,20), note="停電日に施工 H28年次点検指摘"),
 "H29-CR": dict(short="ごみクレーンバケット油圧ユニット整備", amount=1296000, mitsu="特命", cdate=D(2017,9,1), site=(D(2017,10,18), D(2017,10,19)), kanryo=D(2017,10,20), kensa=D(2017,10,24), pay=D(2017,11,20), note="全炉停止中"),
 "H29-地元2": dict(short="灰出しコンベヤNo.2スクレーパ取替修繕", amount=216000, cdate=D(2017,10,2), site=(D(2017,10,18), D(2017,10,18)), kanryo=D(2017,10,19), kensa=D(2017,10,24), pay=D(2017,11,20), note="10枚"),
 "H29-地元4": dict(short="事務室空調機取替修繕", amount=594000, cdate=D(2018,1,10), site=(D(2018,2,14), D(2018,2,14)), kanryo=D(2018,2,15), kensa=D(2018,2,19), pay=D(2018,3,16)),
}

AMT_OVR = {"R5-地元": 1650000, "H26-地元1": 334800, "H27-地元1": 421200, "R2-地元": 880000, "H25-地元": 1260000,
           "R3-CR": 1452000, "R1-CR": 1650000, "H30-計量": 486000, "R7-計量": 1430000}
def local_contracts():
    groups = OrderedDict()
    for w in W.WORKS:
        if w.cid or w.vendor not in LOCAL_VENDORS: continue
        k = w.contract
        if k.endswith("-MK") or "計量" in k and k not in ("H30-計量", "R7-計量"): continue
        groups.setdefault((w.fy, k), []).append(w)
    for (fy, k), ws in groups.items():
        cid = "L-" + k
        vendor = ws[0].vendor
        c = C(cid=cid, fy=fy, key=k, vendor=vendor, ctype="請負", hi="修繕料", kind="LOCAL", works=[])
        sp = SAMPLE_LOCAL.get(k, {})
        c.short = sp.get("short") or short_name(ws)
        c.name = FULL(fy, c.short)
        est = 0
        for w in ws:
            e = W.BYMID[w.mid]
            a, b = W.COST_UNIT.get(e.cls, (0, 0))
            v = (a if w.kind == "REPLACE" else b) * 1000 * (0.45 if w.kind == "REPLACE" else 0.6)
            est += max(v, 90000)
        c.amount = sp.get("amount") or AMT_OVR.get(k) or r1000(est * tax(fy) * RND.uniform(.85, 1.15))
        c.chg = sp.get("chg")
        c.mitsu = sp.get("mitsu", "特命" if vendor == W.CRANE else "2")
        c.note = sp.get("note", "")
        for kk in ("cdate", "start", "end", "site", "kanryo", "kensa", "pay"):
            if kk in sp: setattr(c, kk, sp[kk])
        c.docs = vendor == W.LOCAL   # 浜風機工だけ工程表・日報・完了報告を出す（町が求めるのは請負額 10 万円以上の浜風の工事）
        c.koutei = c.nippo = c.report = c.docs
        for w in ws: w.cid = cid
        c.works = ws
        add(c)
local_contracts()

# ============ 保守委託・役務・物品（毎年） ============
def yearly_misc():
    for fy in range(2012, 2027):
        t = tax(fy)
        rows = [
         ("ELEC%d" % fy, "電気工作物保安管理業務委託", W.ELEC, "委託料", "特命", r1000(1260000 * t), (D(fy,4,1), D(fy+1,3,31)), "年次点検（停電）・月次点検"),
         ("ANAL%d" % fy, "排ガス等測定業務委託", W.ANAL, "委託料", "3", r1000(1800000 * t), (first_wd(D(fy,5,15)), D(fy+1,3,15)), "DXN 年1、ばいじん等 年2"),
         ("ENV%d" % fy, "作業環境測定業務委託", W.ANAL, "委託料", "3", r1000(300000 * t), (first_wd(D(fy,5,15)), D(fy+1,3,15)), "DXN 年2回"),
         ("FIRE%d" % fy, "消防用設備等点検業務委託", W.FIRE, "委託料", "2", r1000(250000 * t), (D(fy,4,1), D(fy+1,3,31)), ""),
         ("EV%d" % fy, "エレベーター保守点検業務委託", "【架空】昇和エレベータ株式会社", "委託料", "特命", r1000(420000 * t), (D(fy,4,1), D(fy+1,3,31)), "定期検査含む"),
         ("JOKA%d" % fy, "浄化槽維持管理業務委託", W.SETSUBI, "委託料", "2", r1000(280000 * t), (D(fy,4,1), D(fy+1,3,31)), "法定検査手数料は別"),
        ]
        if fy == 2017:
            rows[0] = ("ELEC2017", "電気工作物保安管理業務委託", W.ELEC, "委託料", "特命", 1360800, (D(2017,4,1), D(2018,3,31)), "月額113,400 年次点検10/19")
            rows[1] = ("ANAL2017", "排ガス等測定業務委託", W.ANAL, "委託料", "3", 1944000, (D(2017,5,15), D(2018,3,15)), "DXN 年1、ばいじん等 年2")
            rows[2] = ("ENV2017", "作業環境測定業務委託", W.ANAL, "委託料", "3", 324000, (D(2017,5,15), D(2018,3,15)), "DXN 年2回")
            rows[3] = ("FIRE2017", "消防用設備等点検業務委託", W.FIRE, "委託料", "2", 270000, (D(2017,4,1), D(2018,3,31)), "")
            rows[4] = ("EV2017", "エレベーター保守点検業務委託", "【架空】昇和エレベータ株式会社", "委託料", "特命", 453600, (D(2017,4,1), D(2018,3,31)), "定期検査含む")
            rows[5] = ("JOKA2017", "浄化槽維持管理業務委託", W.SETSUBI, "委託料", "2", 302400, (D(2017,4,1), D(2018,3,31)), "法定検査手数料は別")
        if fy == 2022:
            amts = [1386000, 1980000, 330000, 286000, 462000, 308000]
            dts = [(D(2022,4,1), D(2023,3,31)), (D(2022,5,16), D(2023,3,15)), (D(2022,5,16), D(2023,3,15)), (D(2022,4,1), D(2023,3,31)), (D(2022,4,1), D(2023,3,31)), (D(2022,4,1), D(2023,3,31))]
            notes = ["年次点検10/20停電", "", "", "R4は報告年でない", "", ""]
            rows = [(r[0], r[1], r[2], r[3], r[4], amts[i], dts[i], notes[i]) for i, r in enumerate(rows)]
        for (cid, nm, v, hi, mi, amt, (a, b), nt) in rows:
            add(C(cid=cid, fy=fy, name=FULL(fy, nm), short="%s %s" % (W.wareki(fy), nm), vendor=v, ctype="委託", hi=hi, mitsu=mi,
                  amount=amt, cdate=a, start=a, end=b, kind="ITAKU", note=nt))
        if fy % 2 == 0:
            d = first_wd(D(fy, 11, 9))
            add(C(cid="SCALE%d" % fy, fy=fy, name=FULL(fy, "計量機定期検査"), short="計量機定期検査手数料", vendor="【架空】しおさい計量協会",
                  ctype="役務", hi="役務費", mitsu="－", amount=12600 if fy >= 2019 else 12300, cdate=d, start=d, end=d, kind="YAKUMU", note="偶数年度"))
yearly_misc()

# 物品
add(C(cid="BUY2017", fy=2017, name=FULL(2017, "火格子片（予備品）購入"), short="火格子片（予備品）購入", vendor=W.MAKER, ctype="物品", hi="消耗品費",
      mitsu="特命", amount=1036800, cdate=D(2017,6,1), start=None, end=D(2017,9,29), kanryo=D(2017,9,29), kensa=D(2017,10,2), pay=D(2017,10,31),
      kind="BUY", note="24枚 町保管", purchase=("火格子片（燃焼火格子用）", 24, "枚")))
add(C(cid="BUY2022", fy=2022, name=FULL(2022, "バグフィルタ用ろ布（予備品）購入"), short="バグフィルタ用ろ布（予備品）購入", vendor=W.MAKER, ctype="物品", hi="消耗品費",
      mitsu="特命", amount=352000, cdate=D(2022,6,20), start=None, end=D(2022,8,31), kanryo=D(2022,8,26), kensa=D(2022,8,29), pay=D(2022,9,26),
      kind="BUY", note="10本 予備品 10/22 2号で2本使用 残8", purchase=("ろ布 PTFE φ150×4,500", 10, "本")))
add(C(cid="BUY2024", fy=2024, name=FULL(2024, "火格子片（予備品）購入"), short="火格子片（予備品）購入", vendor=W.MAKER, ctype="物品", hi="消耗品費",
      mitsu="特命", amount=1023000, cdate=D(2024,6,3), start=None, end=D(2024,9,30), kanryo=D(2024,9,20), kensa=D(2024,9,25), pay=D(2024,10,25),
      kind="BUY", note="22枚 R6定修で1号に18枚使用予定", purchase=("火格子片（燃焼火格子用）", 22, "枚")))
add(C(cid="BUY2014", fy=2014, name=FULL(2014, "火格子片（予備品）購入"), short="火格子片（予備品）購入", vendor=W.MAKER, ctype="物品", hi="消耗品費",
      mitsu="特命", amount=1512000, cdate=D(2014,6,2), start=None, end=D(2014,9,30), kanryo=D(2014,9,26), kensa=D(2014,9,30), pay=D(2014,10,24),
      kind="BUY", note="40枚 H26定修で36枚使用", purchase=("火格子片（燃焼火格子用）", 40, "枚")))

# 法定の点検（委託の中の作業）を作業として持つ
LEGAL_WORKS = []
def legal_works():
    for fy in range(2012, 2027):
        upto = D(2026, 9, 30) if fy == 2026 else D(fy + 1, 3, 31)
        def LW(mid, desc, cid, d, vendor):
            if d <= upto:
                w = W.Work(mid, fy, "INSPECT", desc, cid, start=d, end=d, vendor=vendor, src="legal"); w.cid = cid
                LEGAL_WORKS.append(w)
        LW(W.m("高圧受電盤"), "受変電設備 年次点検（停電）", "ELEC%d" % fy, None or D(fy, 10, 20), W.ELEC)
        LW(W.m("自動火災報知設備受信機"), "消防用設備等 機器点検", "FIRE%d" % fy, first_wd(D(fy, 9, 10)), W.FIRE)
        LW(W.m("自動火災報知設備受信機"), "消防用設備等 総合点検", "FIRE%d" % fy, first_wd(D(fy + 1, 3, 5)), W.FIRE)
        LW(W.m("非常用発電機"), "非常用発電機 点検（消防）", "FIRE%d" % fy, first_wd(D(fy, 9, 10)), W.FIRE)
        LW(W.m("人荷用エレベーター"), "エレベーター 定期検査", "EV%d" % fy, first_wd(D(fy, 7, 8)), "【架空】昇和エレベータ株式会社")
        LW(W.m("合併処理浄化槽"), "浄化槽 法定検査（11条）", "JOKA%d" % fy, first_wd(D(fy, 8, 20)), W.SETSUBI)
        if fy % 2 == 0:
            LW(W.m("計量機"), "計量法 定期検査", "SCALE%d" % fy, first_wd(D(fy, 11, 9)), "【架空】しおさい計量協会")
legal_works()
# world の「計量法 定期検査」作業は SCALE 契約へ
for w in W.WORKS:
    if w.contract.endswith("-計量") and w.kind == "INSPECT": w.cid = "SCALE%d" % w.fy
    if w.contract in ("H28-保安", "R4-保安"): w.cid = "ELEC%d" % w.fy
    if w.contract == "R5-消防": w.cid = "FIRE%d" % w.fy
W.WORKS = [w for w in W.WORKS if not (w.contract.endswith("-計量") and w.kind == "INSPECT")]  # LEGAL_WORKS に同じものがある
for w in W.WORKS:
    if w.contract in ("H28-保安", "R4-保安", "R5-消防"):
        w.legal_dup = True
W.WORKS = [w for w in W.WORKS if not getattr(w, "legal_dup", False)]

for w in W.WORKS:
    if w.cid is None:
        raise SystemExit("契約が決まらない作業: %r %s" % (w, w.contract))
    CONTRACTS[w.cid].works.append(w) if w not in CONTRACTS[w.cid].works else None
for w in LEGAL_WORKS:
    CONTRACTS[w.cid].works.append(w)

def contracts_of(fy, kinds=None):
    return [c for c in CONTRACTS.values() if c.fy == fy and (kinds is None or c.kind in kinds)]
