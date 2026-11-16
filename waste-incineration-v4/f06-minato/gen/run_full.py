# -*- coding: utf-8 -*-
"""F06【架空】みなと臨海工場 本作り：全資料と正解を作り直す。  python3 run_full.py"""
import glob
import json
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import ans
import core
import g_daily
import g_ledger
import g_orders
import g_plan
import g_report
import g_sched
from xlhelp import OUT, WORK

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
KIND_EN = {"I": "INSPECT", "O": "OVERHAUL", "R": "REPLACE", "U": "UPGRADE"}


def clear():
    os.makedirs(OUT, exist_ok=True)
    os.makedirs(WORK, exist_ok=True)
    for d in (OUT, os.path.join(BASE, "answer", "files")):
        os.makedirs(d, exist_ok=True)
        for f in glob.glob(os.path.join(d, "*")):
            if os.path.isfile(f):
                os.remove(f)


def world_json():
    W = core.get()
    eqs = []
    for e in W.eq:
        import s_ledger as L
        eqs.append(dict(tag=e["tag"], 台帳名=e["name"], 号炉=e["unit"], 設備区分=e["setsubi"], 系統=e["keito"], 仕様=e["spec"], 重要度=e["imp"],
                        常用予備=e["duty"], 設置年度=e["installed"], 撤去年度=e["removed"], 基幹改良で新設=bool(e["added"]),
                        呼び名={"機器台帳・計画・日報・報告書": e["name"], "竣工図書機器表": e["name"], "系統図": L.dwg(e)[0],
                              "改良工事機器番号": L.kaimei(e) or None}))
    works = []
    for w in W.works:
        works.append(dict(id=w["id"], tag=w["tag"], kind=KIND_EN[w["kind"]], fy=w["fy"], start=w["start"], end=w["end"],
                          manHours=w["manHours"], manDays=round(w["manHours"] / 8, 2) if w["manHours"] else None, vendor=w["vendor"],
                          team=w["team"], contract=w["contract"], task=w["task"], status=w["status"], note=w["note"] or None))
    contracts = [dict(id=c["id"], fy=c["fy"], kind=c["kind"], name=c["name"], vendor=c["vendor"]) for c in W.contracts.values()]
    linked = {wid for l in ans.LINKS for wid in l["work_ids"]}
    unrec = []
    for w in W.works:
        if w["contract"].endswith("LOCAL"):
            unrec.append(dict(id=w["id"], reason="地元業者への補修工事（市が直接発注）。完成届は紙で施設課に保管され、電子の工程表・日報・報告書はない。発注台帳の契約行だけに出る"))
        elif w["status"] != "done":
            unrec.append(dict(id=w["id"], reason="2026-11-13 時点で未実施（予定）。工程表の予定にだけ出る。日報・報告書はまだない"))
        elif w["id"] not in linked:
            unrec.append(dict(id=w["id"], reason="資料の行に結び付いていない"))
    defects = [
        "中長期整備計画 H25〜H31 版：灰クレーン性能検査の行がない（法令点検の抜け。R2 版で追加）",
        "中長期整備計画 H25〜H31 版：精密機能検査の星取が H28 の次を H32 とし 3 年を超える（H31 版の法定行、全 h3 様式の版で同じ）",
        "中長期整備計画 H29〜H31 版：排水処理設備 R2 の小計に C-P-820A の費用が入っていない（費用の小計の食い違い）",
        "中長期整備計画 H29〜H31 版：H29 の実績・費用の欄が空（担当交代 h3→h4 の追記漏れ）",
        "中長期整備計画 H24 版：H23 版の内容を年だけ送った（h2→h3 の引継ぎ年。H23 年度中の点検結果が反映されていない）",
        "中長期整備計画 R4 版：前回実施の欄が R2 までしか更新されていない（h4→h5 の引継ぎ年）",
        "中長期整備計画 R2〜R8 版：旧様式の小計行（千円）が受入供給設備の下に 1 か所残る（値は H31 版のまま）",
        "中長期整備計画 R5〜R8 版・機器台帳 R6 版：1号炉の基幹改良（R4）の撤去・新設が反映されていない（R7 版台帳で追記）",
        "中長期整備計画 H20〜H31 版：基幹改良の年度の見込み（H28〜H31 版は R3〜R6）が実際（R4〜R7）と 1 年ずれる（計画と実績の差）",
    ]
    profile = dict(id="f06", 施設名="【架空】みなと臨海工場（みなと市 環境局）", 竣工="2004-03", 能力="200t/24h×3炉＝600t/日",
                   炉形式="並行揺動式ストーカ（水冷火格子の一部採用）・全連続", ボイラ="4MPa・400℃・22t/h×3", 発電="抽気復水タービン 11,000kW",
                   排ガス処理="減温塔＋消石灰・活性炭＋BF＋湿式洗煙＋再加熱＋触媒脱硝", 改良="2014 DCS一部更新、2022–2025 基幹的設備改良",
                   契約="運転委託（メーカー系 5 年）＋年度の定期整備業務委託（3 炉時期分割）＋全炉停止時共通設備整備（H23 から別契約）＋タービン定期事業者検査（4 年ごと別契約）＋補修・更新工事＋基幹改良（4 か年）",
                   対象期間="2004–2026 年度（23 年）", 記録の時点="2026-11-13", 管理機器="1,275 台（炉別 310×3＋共通 345、P&ID タグ単位）。基幹改良前 1,267")
    out = dict(profile=profile, equipment=eqs, contracts=contracts, works=works, links=ans.LINKS, unrecorded_works=unrec, plan_defects=defects,
               catalog=ans.CATALOG)
    with open(os.path.join(BASE, "answer", "world.json"), "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=0)
    return len(unrec)


def main():
    t0 = time.time()
    clear()
    W = core.get()
    n_ev = sum(1 for e in W.eq for y, a in __import__("world").simulate(e).items() if y <= 2026 for _ in a)
    assert n_ev == len(W.works), (n_ev, len(W.works))
    res = {}
    for name, mod in (("工程表", g_sched), ("作業日報", g_daily), ("工事報告書", g_report), ("発注台帳", g_orders), ("中長期整備計画", g_plan),
                      ("機器台帳", g_ledger)):
        t1 = time.time()
        res[name] = len(mod.main())
        print(name, res[name], round(time.time() - t1, 1), flush=True)
    ans.save()
    nu = world_json()
    print("unrecorded", nu, "total", round(time.time() - t0, 1))
    with open(os.path.join(BASE, "work", "counts.json"), "w", encoding="utf-8") as f:
        json.dump(res, f, ensure_ascii=False)


if __name__ == "__main__":
    main()
