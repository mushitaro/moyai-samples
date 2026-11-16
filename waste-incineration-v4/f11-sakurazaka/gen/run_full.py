# -*- coding: utf-8 -*-
"""F11【架空】さくら坂焼却場 本作り：python3 run_full.py（full/gen で実行）。
files/ と answer/ を空にして全部作り直す。乱数の種は固定。"""
import os, glob, json, random, datetime as dt, shutil
from common import OUT, BASE
import world as W
import fullworld as FW
import plan_data as P
import gf_sched, gf_daily, gf_report, gf_order, gf_plan, gen_ledger, gf_answer as A

ANS = os.path.join(BASE, "answer", "files")
for d in (OUT, ANS):
    os.makedirs(d, exist_ok=True)
    for f in glob.glob(os.path.join(d, "*")):
        os.remove(f)

jobs = FW.build_all()
works = FW.works_from(jobs)
JOB = {j.cid: j for j in jobs}
manifest = []   # (file, 資料, 年度, 契約)

def add(fn, kind, fy, cid=None, recs=None, plan=False, extra=None):
    manifest.append(dict(file=os.path.basename(fn), 資料=kind, 年度=fy, 契約=cid))
    A.dump(fn, recs or [], plan, extra)

# ---- 台帳 ----
for fn in gen_ledger.main():
    add(fn, "機器台帳", 2019 if "令和元年度" in fn else 2025, None, A.ledger_records(fn))

# ---- 計画 ----
ca = {k: {} for k in P.COST_ROWS}
for y in range(2020, 2026):
    for k, v in gf_order.actual_cost_kyen(y).items():
        ca[k][y] = v

def done(e, y):
    for j in jobs:
        if j.fy != y:
            continue
        for s, ab, ts in j.sheets:
            for t in ts:
                if any(x.no == e.no for x, m in t.eqs) and t.act and max(t.act) <= FW.TODAY:
                    return True
    return False

LEGAL_DONE26 = {"boiler1", "boilerm", "crane_y", "pv1", "pv2", "pvs", "elec", "eg", "gas", "dxn", "dxnw", "scale", "fire"}
fn = gf_plan.build_old(); add(fn, "中長期整備計画", 2019, None, A.plan_records(fn), True)
fn = gf_plan.build_new(2023, cost_act=ca); add(fn, "中長期整備計画", 2023, None, A.plan_records(fn), True)
fn = gf_plan.build_new(2026, actual_fn=lambda e: {y: FW.mark(e, y) for y in range(2020, 2027) if FW.mark(e, y)}, done_fn=done, cost_act=ca, legal_done26=LEGAL_DONE26)
add(fn, "中長期整備計画", 2026, None, A.plan_records(fn), True)

# ---- 工程表・日報・報告書 ----
rng = random.Random(1111)
rnd = random.Random(2019)
scan_info = {}
for j in jobs:
    fn, recs = gf_sched.build(j, rng); add(fn, "工程表", j.fy, j.cid, recs)
    fn, recs, info = gf_daily.build(j, rnd)
    add(fn, "作業日報", j.fy, j.cid, recs, extra={"images": info} if info else None)
    if info:
        scan_info[os.path.basename(fn)] = info
    fn, recs = gf_report.build(j, rnd)
    if fn:
        add(fn, "工事報告書", j.fy, j.cid, recs)

# ---- 発注 ----
for y in range(2019, 2027):
    fn, recs, R = gf_order.build_csv(y); add(fn, "発注台帳", y, None, recs)
    fn, recs = gf_order.build_tehai(y); add(fn, "発注台帳", y, None, recs)

# ---- world.json ----
def inst_year(s):
    return 1988 if s.startswith("S63") else 2002
eq = []
for e in W.EQUIP:
    aliases = {"台帳": e.name, "計画": e.name, "日報（手書き）": gf_daily.abbr(e.name), "報告書": ("%s号炉 " % e.unit if e.unit in ("1", "2") else "") + e.name}
    for key, nm, names, nd in FW.GROUPS_UNIT:
        if e.name in names:
            aliases["工程表のまとまり"] = nm
    eq.append(dict(no=e.no, oldno=e.oldno, name=e.name, unit=e.unit, units=e.units, install_year=inst_year(e.inst), install=e.inst,
                   renew=e.renew, remark=e.remark, category=W.CAT[e.cat], aliases=aliases))
# 他業者・運転委託の作業（e1 の帳票には出ないもの）
extra_works = []
k = len(works)
def xw(equip, attr, s, e_, md, vendor, contract, task):
    global k
    k += 1
    extra_works.append(dict(id="W%04d" % k, equip=equip, attr=attr, start=s, end=e_, manDays=md, vendor=vendor, contract=contract, task=task))
xw("0L-01", "REPLACE", "2020-12-07", "2020-12-18", 18, "さくら坂情報システム", "R2020b", "監視操作端末 3 台更新・ソフト移植")
for y in range(2019, 2027):
    xw("", "INSPECT", "%d-05-20" % y, "%d-02-20" % (y + 1) if y < 2026 else "2026-11-10", None, "かもめ環境分析", "測定委託 %d" % y, "排ガス測定（各炉 年 6 回）・DXN 年 1 回・作業環境測定 年 2 回")
    xw("", "INSPECT", "%d-09-01" % y, "%d-03-01" % (y + 1) if y < 2026 else "2026-09-01", None, "さくら坂防災設備", "点検委託 %d" % y, "消防用設備等点検（機器 6 か月・総合 1 年）")
    if y in (2019, 2022, 2025):
        xw("", "INSPECT", "%d-09-10" % y, "%d-11-30" % y, None, "ひがた技術コンサルタント", "検査委託 %d" % y, "精密機能検査")
    if y in (2020, 2022, 2024, 2026):
        xw("0A-01", "INSPECT", "%d-09-10" % y, "%d-09-10" % y, None, "県計量協会", "手数料 %d" % y, "計量器定期検査")
# 帳票の行 → 作業
links = []
for p in glob.glob(os.path.join(ANS, "*.json")):
    d = json.load(open(p, encoding="utf-8"))
    for i, r in enumerate(d["records"]):
        f = r.get("fields", {})
        ids = []
        c = f.get("contract")
        if c and f.get("作業"):
            ids = [w["id"] for w in works if w["contract"] == c and w["task"] == f["作業"]]
        elif f.get("日付") and d["file"].startswith("【架空】さくら坂焼却場_作業日報"):
            cid = [m["契約"] for m in manifest if m["file"] == d["file"]][0]
            dd = f["日付"]
            ids = [w["id"] for w in works if w["contract"] == cid and w["start"] <= dd <= w["end"]]
        elif f.get("code") and f.get("marks"):
            for y, mk in f["marks"].items():
                if "済" in str(mk):
                    ids += [w["id"] for w in works if w["equip"] == f["code"] and w["start"][:4] in (y, str(int(y) + 1)) and int(w["start"][:4]) + (0 if int(w["start"][5:7]) >= 4 else -1) == int(y)]
        elif c and d["file"].endswith(".csv"):
            ids = sorted({w["id"] for w in works + extra_works if w["contract"] == c})[:50]
        if ids:
            key = "r" if "r" in r else "seq"
            links.append(dict(file=d["file"], sheet=r.get("sheet"), **{key: r.get(key)}, works=ids))
world = dict(
    profile=dict(id="f11", name=W.FAC, city=W.CITY, furnace="階段式ストーカ・全連続", capacity="60t/24h×2炉＝120t/日", completed="1988-03",
                 boiler="廃熱ボイラ 1.5MPa 飽和 5t/h（場内・温水プール）、発電なし", gas="減温塔＋消石灰・活性炭＋BF", idf="19,000m3N/h",
                 history=["2001–02 EP→BF ほか DXN 対策", "2009–11 延命化", "2019 年度 メーカー（朝凪炉工）整備撤退・部品供給のみ", "2027 年度末 廃止予定（2号炉 R9.9 末停止、1号炉 R9 年度末停止）"],
                 contracts="運転委託（2010–）、定期整備は地元整備会社へ単年度委託（2019–）、補修請負 年 0〜2 件、部品は市が直接購入",
                 period="2019–2026 年度（8 年）、記録の時点 2026-11-13（R8 は上期＋2号定修の途中まで）", equipment="198 行（欄）・273 台（炉別 60 行 82 台×2、共通 78 行 109 台）",
                 people=W.PEOPLE, today=FW.TODAY.isoformat(), seed=1111),
    equipment=eq,
    contracts=[dict(id=c["id"], fy=c["fy"], title=c["title"], kind=c["kind"], vendor=c["vendor"], start=c["start"].isoformat(), end=c["end"].isoformat()) for c in W.CONTRACTS],
    works=works + extra_works,
    links=links,
    unrecorded=[
        dict(what="中央監視装置 監視端末更新（R2020b）", why="委託先（さくら坂情報システム）の完了届は紙のみで、工程表・日報・報告書は e1 の様式で作っていない。CSV と手配表、計画に出る"),
        dict(what="排ガス・DXN・作業環境測定、消防点検、精密機能検査、計量器検査", why="測定・点検委託の業者の報告書は市の紙ファイル。CSV（負担行為・支出命令）と計画の法定点検の行だけに出る"),
        dict(what="ボイラー・クレーンの月例自主検査、受変電の月次点検", why="運転委託・電気保安法人が行い、e1 の日報には書かない。計画の法定点検の行（月例）だけ"),
        dict(what="令和8年度 定期整備 2号炉の 11/16 以降の作業（耐火物の残り、ボイラ性能検査 11/17、立上げ 11/20）", why="記録の時点（2026-11-13）より後。工程表は予定だけ、日報・報告書はまだ無い"),
        dict(what="令和8年度 定期整備業務委託の完了報告書", why="年度末（R9.3）提出のため記録の時点では未作成"),
    ],
    plan_gaps=[
        dict(file="中長期整備計画 l1 R1.12", gap="第一種圧力容器（温水熱交換器 2 基）の性能検査の行がない（l2 が R5.4 に追加）"),
        dict(file="中長期整備計画 l1 R1.12", gap="計量機定期検査を R3・R5・R7・R9 に置いた（前回 H30.11 から 3 年あく＝周期の超え）。実績は R2・R4・R6・R8"),
        dict(file="中長期整備計画 l2 R5.4・R8.11", gap="費用 R3 決算計が行の合計より 110 千円少ない（税抜の 1 件を混ぜた）"),
        dict(file="中長期整備計画 l2", gap="2号炉の R9 定修・ボイラ性能検査なし（R9.9 末停止のため。欠けではなく意図）"),
        dict(file="中長期整備計画 l2 R8.11", gap="「前回整備（記録）」列が R1 のまま更新されていない"),
    ],
)
with open(os.path.join(BASE, "answer", "world.json"), "w", encoding="utf-8") as f:
    json.dump(world, f, ensure_ascii=False, indent=1, default=str)
with open(os.path.join(BASE, "answer", "manifest.json"), "w", encoding="utf-8") as f:
    json.dump(manifest, f, ensure_ascii=False, indent=1)
print("files", len(manifest), "works", len(works) + len(extra_works), "links", len(links))
from collections import Counter
print(Counter(m["資料"] for m in manifest))
