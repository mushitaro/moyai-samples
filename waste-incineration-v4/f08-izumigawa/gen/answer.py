# -*- coding: utf-8 -*-
"""answer/world.json（機器・作業・結び・書いていない作業・計画の欠け・profile）。"""
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import ANS  # noqa: E402
import rec  # noqa: E402
import world  # noqa: E402
from w_report import M1_NAME  # noqa: E402
from works import CONTRACTS, EQ, actual_works, policy  # noqa: E402


def write(jobs):
    pol = policy()
    eqs = []
    for e in EQ:
        m1 = e["name"]
        for k, v in M1_NAME.items():
            m1 = m1.replace(k, v)
        eqs.append({"id": e["id"], "ledgerName": e["name"], "line": e["line"] or "共通", "setsubi": world.SETSUBI[e["setsubi"]],
                    "keitou": e["keitou"], "installedFY": e["year"], "drawingNo": e["tag"] or None,
                    "oldDrawingNo": e.get("old_tag"), "duty": e["duty"] or None, "importance": e["imp"],
                    "policy2016": pol[e["id"]],
                    "names": {"機器台帳・計画": e["name"], "m1報告書": m1}})
    removed = [{"drawingNo": r["tag"], "name": r["name"], "line": r["line"] or "共通", "removedFY": r["removed_fy"],
                "replacedBy": r["by"], "managed": False} for r in world.REMOVED]
    W = []
    linked = {w for l in rec.LINKS for w in l["works"]}
    for w in actual_works():
        W.append({"id": w["wid"], "equipment": w["eq"], "kind": w["kind"], "label": w["label"], "fy": w["fy"],
                  "start": rec.jd(w.get("start")), "end": rec.jd(w.get("end")), "manDays": w.get("manDays"),
                  "vendor": w["vendor"] or CONTRACTS.get(w["contract"], {}).get("vendor"), "contract": w["contract"]})
    nodoc = []
    for w in actual_works():
        if w["contract"].startswith("K"):
            nodoc.append({"work": w["wid"], "reason": "市の直接の点検委託（黒潮クレーン）。工程表・日報・報告書の見本は作っていない。計画の実績欄と CSV にだけある"})
        elif w["contract"].startswith("S"):
            nodoc.append({"work": w["wid"], "reason": "紫雲電機の修繕・更新（補修工事の外）。工程表・日報・報告書は無く、CSV・台帳・計画にだけある"})
    gaps = [
        {"file": "中長期整備計画 旧様式（H21・H24・H27）", "what": "法定点検に作業環境測定（ダイオキシン類）と第二種圧力容器定期自主検査の行が無い", "truth": "毎年実施（発注台帳に委託あり）"},
        {"file": "中長期整備計画 新様式 R4改訂", "what": "ごみクレーン性能検査の R6 が空き", "truth": "R6 に受検"},
        {"file": "中長期整備計画 旧様式 H27改訂 費用シート", "what": "燃焼設備の小計 H27 が機器行の和より 600 千円少ない", "truth": "小計の直し忘れ"},
        {"file": "中長期整備計画 新様式 R4改訂", "what": "R4〜R6 の実績が空き（R5 以降は更新なし）", "truth": "R4・R5・R6 にも作業あり（world の works）"},
    ]
    profile = {"id": "f08", "name": "【架空】いずみ川清掃センター", "city": "【架空】いずみ川市", "furnace": "流動床式焼却炉・全連続",
               "capacity": "75t/24h×2=150t/日", "completed": "1990-03", "stopped": "2025-03（2024年度末）",
               "boiler": "廃熱ボイラ 1.0MPa 飽和 6t/h（発電なし）", "idf": "24,000m3N/h", "period": [2005, 2024],
               "equipmentCount": {"perLine": 95, "common": 128, "total": 318},
               "people": {"i1": "市 2005–2013", "i2": "市 2014–（2023年度以降も i2）", "m1": "メーカー現場代理人 2005–2018",
                          "g1": "地元業者 職長 2019–2024"}}
    out = {"profile": profile, "equipment": eqs, "removedEquipment": removed, "works": W,
           "contracts": {k: v for k, v in CONTRACTS.items()}, "links": rec.LINKS,
           "worksWithoutDocs": nodoc, "planGaps": gaps,
           "unlinkedWorks": sorted({w["wid"] for w in actual_works()} - linked)}
    with open(os.path.join(ANS, "world.json"), "w", encoding="utf-8") as f:
        json.dump(rec.jd(out), f, ensure_ascii=False, indent=1)
