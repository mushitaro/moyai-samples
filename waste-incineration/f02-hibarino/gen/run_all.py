# -*- coding: utf-8 -*-
"""【架空】ひばり野クリーンセンター（F02）本作り：2016〜2026 年度の全資料と正解を作り直す。
cd f02/full/gen && python3 run_all.py   （LibreOffice は専用プロファイル file:///tmp/lo_F02）"""
import glob
import json
import os
import shutil
import sys
import datetime as dt
import world as W
import works as WK
import contracts as K
import gen_contract as G
import rec
from xlhelp import normalize_xlsx

BASE = rec.BASE
for d in (rec.FILES, rec.ANS):
    if os.path.isdir(d):
        shutil.rmtree(d)
    os.makedirs(d)

WK.build_all()
contracts = G.all_contracts()
import b_plan, b_ledger, b_sched, b_daily, b_report, b_orders, b_city

files = {}
files["中長期整備計画"] = b_plan.build_all()
files["機器台帳"] = b_ledger.build_all()
files["工程表"], files["作業日報"], files["工事報告書"] = [], [], []
for c in contracts:
    files["工程表"].append(b_sched.build(c))
    files["作業日報"].append(b_daily.build(c))
    fn = b_report.build(c)
    if fn:
        files["工事報告書"].append(fn)
files["発注台帳"] = b_orders.build_all(contracts)
cf = b_city.build_all(contracts)
for fn in cf:
    for k in files:
        if f"_{k}_" in fn:
            files[k].append(fn)

# ---------------------------------------------------------------- 作業の日付・工数・業者・契約
for c in contracts:
    y = c["y"]
    rows = K.daily_rows(c["tasks"], y, seed_add=b_daily.seed_of(c))
    for t in c["tasks"]:
        ids = [w for w in (WK.wid(W.eq(n, f), y) for n, f in t[7]) if w]
        if not ids:
            continue
        ds = K.task_days(t, y)
        mh = sum(m * h for x, v in rows.items() for (tt, d_, m, h) in v if tt[0] == t[0])
        for w in ids:
            o = WK.WORKS[w]
            o["start"] = min(o["start"] or ds[0].isoformat(), ds[0].isoformat())
            o["end"] = max(o["end"] or ds[-1].isoformat(), ds[-1].isoformat())
            o["manDays"] = round((o["manDays"] or 0) + mh / 8 / len(ids), 2)
            o["vendor"] = "鷹羽エンジニアリング株式会社（施工 " + G.CREW_NAME.get(t[3], t[3]) + "）"
            o["contract"] = c["cid"] + "：" + c["title"]
for cid, cc in b_city.CITY.items():
    o = WK.WORKS[b_city.city_work(cid)]
    ds = [dt.date(cc["fy"], *K.d(k)) for k in cc["days"]]
    o["start"], o["end"] = min(ds).isoformat(), max(ds).isoformat()
    o["manDays"] = sum(sum(v[1].values()) for v in cc["days"].values())
    o["contract"] = cid + "：" + cc["title"]
# 経常修繕（発注 CSV から）
for fn in files["発注台帳"]:
    for r_ in rec.REG[fn].records:
        pass
linked = {}
for l in rec.LINKS:
    linked.setdefault(l["work"], []).append(l)
for w, o in WK.WORKS.items():
    if o["contract"]:
        continue
    fy = o["fy"]
    e = [e for e in W.EQ if e.no == o["equipment"]][0]
    m = W.oh_month(e.f, fy, e.seq)
    yr = fy + 1 if m <= 3 else fy
    o["start"] = o["end"] = None
    o["month"] = f"{yr}-{m:02d}"
    if o["attr"] == "INSPECT":
        o["vendor"] = "SPC 保全課・運転課（直営点検）"
        o["contract"] = "SPC 直営（長期包括運営委託の点検業務）"
    else:
        o["vendor"] = "発注 CSV の仕入先"
        o["contract"] = f"SPC 経常修繕（案件 M{fy % 100:02d}-K）"
future = [w for w, o in WK.WORKS.items() if o.get("month") and o["month"] > "2026-11"]
for w in future:
    del WK.WORKS[w]          # 2026-11-13 時点では未実施（2026 年度下期の予定）
unrec = []
for w, o in WK.WORKS.items():
    if w not in linked:
        unrec.append({"work": w, "reason": "SPC の直営点検（定修中の SPC 保全課の点検を含む）で、記録は SPC 保全システム（今回の資料の範囲外）。中長期整備計画の実績段は 2022 年度改訂（〜2021 実績）までで、s2 様式には実績段がない"
                      if o["attr"] == "INSPECT" else "経常修繕の発注はあるが資産番号の記載がない／計画の実績段が無い年度"})

GAPS = [
    {"file": b_plan.fname(2021), "where": "法定欄 計量機 定期検査 2024", "what": "2 年周期の○が抜け（2021・2022 年度改訂）。2023 年度改訂で復活。2024 年度は実施（発注 CSV に行あり）"},
    {"file": b_plan.fname(2022), "where": "法定欄 計量機 定期検査 2024", "what": "同上"},
    {"file": "2023〜2026 年度改訂（s2 様式）", "where": "法定欄 精密機能検査 2029", "what": "廃止前年として外し「要協議」→ 3 年に 1 回を超える"},
    {"file": "2018〜2022 年度改訂（s1 様式）", "where": "シート 年度別概算費用", "what": "「廃止前 撤去準備調査（H30追加）」行が合計に入っていない（2028/2029/2030 で 3,200/4,800/6,500 千円）"},
    {"file": "全版", "where": "2号押込送風機", "what": "分解整備 2017→2025（8 年、他の送風機の 6 年周期超え、廃止まで延伸）"},
    {"file": "2018〜2022 年度改訂（s1 様式）", "where": "資産番号列", "what": "更新で付け替わった番号（HB08-101R 等・計量機 R）を反映せず旧番号のまま（台帳の「更新・付替え」シートと食い違い）"},
]
profile = {"id": "f02", "name": W.FAC, "city": "ひばり野市（架空）", "spc": W.SPC, "maker": W.MAKER, "sub": W.SUB,
           "furnace": "逆送式ストーカ 105t/24h×2（210t/日）全連続", "completed": "1999-03", "boiler": "3MPa・300℃・11.5t/h",
           "power": "2,400kW", "flue_gas": "減温塔→BF→ガス再加熱器→低温触媒脱硝→誘引通風機", "contract": "2016–2030 長期包括運営委託（SPC）＋市の追加工事",
           "closure": "2018 年度に 2030 年度末廃止を決定", "period": "2016–2026 年度（11 年）", "as_of": "2026-11-13",
           "equipment_count": {"total": 262, "furnace1": 72, "furnace2": 72, "common": 118, "rule": "SPC 資産番号 1 件＝1 台（主機）。盤は系統ごと。弁・計装・建築設備は入れない"},
           "people": W.PEOPLE}
eqs = [{"asset": e.no, "asset_renumbered": {str(W.RENUM[(e.name, e.f)][0]): e.no + W.RENUM[(e.name, e.f)][1]} if (e.name, e.f) in W.RENUM else None,
        "ledger_name": e.name if e.f == 0 else f"{e.f}号{e.name}", "plan_name": e.name, "aliases": W.ALIAS.get(e.name, []),
        "furnace": e.f, "category": W.CATS[e.cat], "installed": e.install, "importance": e.imp, "role": e.role or None,
        "policy_2018": e.policy, "spec": e.spec} for e in W.EQ]
world = {"profile": profile, "equipment": eqs, "works": list(WK.WORKS.values()), "links": rec.LINKS, "unrecorded_works": unrec,
         "plan_gaps": GAPS,
         "contracts": [{"id": c["cid"], "title": c["title"], "period": [c["period"][0].isoformat(), c["period"][1].isoformat()],
                        "stop": c["stop"].isoformat(), "orderer": W.SPC, "contractor": W.MAKER} for c in contracts] +
                      [{"id": k, "title": v["title"], "period": [v["period"][0].isoformat(), v["period"][1].isoformat()], "orderer": "ひばり野市",
                        "contractor": W.MAKER} for k, v in b_city.CITY.items()],
         "files": files, "not_made": [{"what": "2026年度 1号炉定期整備工事 報告書", "reason": "2026-11-13 時点で未提出（工事は 11/11 まで、報告は 12 月）"}]}
with open(os.path.join(BASE, "answer", "world.json"), "w", encoding="utf-8") as fp:
    json.dump(world, fp, ensure_ascii=False, indent=1, default=str)
rec.dump()
for f in glob.glob(os.path.join(rec.FILES, "*.xlsx")):
    normalize_xlsx(f)
n = {k: len(v) for k, v in files.items()}
print(n, sum(n.values()), "files;", len(WK.WORKS), "works;", len(rec.LINKS), "links;", len(unrec), "unrecorded")
