# -*- coding: utf-8 -*-
"""正解の書き出し：answer/files/<stem>.json と answer/world.json"""
import os, json, datetime as dt
import world as W, cw, sg
from xl import REC, LINKS, EQLINKS, OUT

ANS = os.path.abspath(os.path.join(OUT, "..", "answer"))

def js(o):
    if isinstance(o, (dt.date, dt.datetime)): return o.isoformat()
    raise TypeError(type(o))

PLAN_GAPS = [
 {"file": "【架空】しおさい町ごみ処理場_中長期整備計画_H27策定長寿命化計画_コンサル策定版.xlsx", "gap": "法定点検の行が無い（計量機定期検査・クレーン性能検査・DXN作業環境測定・精密機能検査など）。毎年の委託点検は表示省略"},
 {"file": "【架空】しおさい町ごみ処理場_中長期整備計画_H27策定長寿命化計画_コンサル様式t1t2追記.xlsx", "gap": "R2・R3（2020–2021）の実績が未記入（計画の追記が止まった期間）。R4 は定修委託とチェーン取替だけ記入"},
 {"file": "【架空】しおさい町ごみ処理場_中長期整備計画_H27策定長寿命化計画_コンサル様式t1t2追記.xlsx", "gap": "通風設備 H27・H28 の小計に t1 が追記したインバータ盤（各 4,800 千円）が入っていない（費用の小計の食い違い）"},
 {"file": "【架空】しおさい町ごみ処理場_中長期整備計画_R5〜R15_t3計画実績2列様式.xlsx", "gap": "精密機能検査を R7 の次に R11 と書いた（3年周期を超える）"},
 {"file": "【架空】しおさい町ごみ処理場_中長期整備計画_R5〜R15_t3計画実績2列様式.xlsx", "gap": "年度計に法定点検（委託料）を含まない。2026 は上期実績のみ（10月定修分の実が空欄）"},
 {"file": "【架空】しおさい町ごみ処理場_中長期整備計画_H27策定長寿命化計画_コンサル様式t1t2追記.xlsx", "gap": "2号炉バグフィルタ H34（R4）に「◎実」と書いたが、実際は破損ろ布2本だけ（全数はR5）"},
]

PROFILE = {
 "id": "f04", "name": "【架空】しおさい町ごみ処理場", "owner": "しおさい町（単独）", "built": "1995-03", "capacity": "20t/16h×2炉＝40t/16h（帳票では40t/日（16h））",
 "furnace": "階段式ストーカ・准連続（16h）", "heat_recovery": "なし（水噴射式ガス冷却）", "flue_gas": "消石灰・活性炭吹込＋バグフィルタ",
 "improvements": ["2001 DXN対策（EP→BF）", "2015–2016 長寿命化改良工事（火格子・ガス冷却室・誘引通風機・ろ布・計装）"], "closure": "2033年度末 廃止予定（広域化）",
 "contracts": "町直営運転。メーカー【架空】水樹プラントの定期点検整備業務委託（年1回）、地元業者【架空】浜風機工ほかの少額修繕（年3〜6件）",
 "period": "2012–2026 年度（2026 は計画＋上期実績）", "equipment_count": 256, "count_rule": "機器1台＝1台、並列同型機も各1台、盤1面1台、建築設備を含む。弁・計装品は除く",
 "staff": {"町": "t1（2012–2017）→ t2（2018–2022）→ t3（2023–）", "地元業者": "g1（–2018）→ g2（2019–）", "メーカー": "SV v1"},
 "plan_not_updated": "2020–2021 年度は計画（星取表）への実績追記が止まっている",
}

def write_all(made):
    os.makedirs(os.path.join(ANS, "files"), exist_ok=True)
    for stem, d in REC.items():
        with open(os.path.join(ANS, "files", stem + ".json"), "w", encoding="utf-8") as f:
            json.dump(d, f, ensure_ascii=False, indent=1, default=js)
    allw = W.WORKS + cw.LEGAL_WORKS
    linked = {l["work"] for l in LINKS}
    eqs = []
    names = {}
    for l in EQLINKS:
        st = REC[os.path.splitext(l["file"])[0]]
        rr = [x for x in st["records"] if x.get("sheet") == l["sheet"] and x.get("r") == l["r"]]
        if rr:
            nm = rr[0]["fields"].get("機器名称")
            for m in l["mids"]: names.setdefault(m, {})[l["file"]] = nm
    for e in W.EQ:
        al = {"台帳2024": e.name, "竣工図書・台帳2015": e.oname, "旧計画（コンサル）": "灰出しコンベヤ No.1" if e.name == "灰出しコンベヤNo.1" else e.name}
        if e.name == "誘引通風機": al["H24〜H27 報告書"] = "誘引送風機"
        if e.name == "バグフィルタ": al["日報"] = "BF"
        if e.name == "灰出しコンベヤNo.1": al["浜風機工の資料"] = "灰出しコンベヤ"
        eqs.append({"mid": e.mid, "name": e.name, "aliases": al, "furnace": e.furnace, "system": W.SYS[e.sys], "installed_fy": e.inst,
                    "importance": e.imp, "spec": e.spec, "drawing_no": e.dno, "added_2015": e.added_2015, "rows": names.get(e.mid, {})})
    works = []
    for w in sorted(allw, key=lambda w: w.wid):
        c = cw.CONTRACTS.get(w.cid)
        works.append({"id": w.wid, "mid": w.mid, "equipment": W.BYMID[w.mid].name, "furnace": W.BYMID[w.mid].furnace, "fy": w.fy,
                      "kind": w.kind if w.kind in ("INSPECT", "OVERHAUL", "REPLACE", "UPGRADE") else "OVERHAUL", "desc": w.desc,
                      "start": w.start, "end": w.end or w.start, "manHours": round(w.mh, 1) if w.mh else None, "vendor": w.vendor,
                      "contract": w.cid, "contract_name": c.name if c else None, "source": w.src})
    unrec = []
    for w in allw:
        if w.wid not in linked:
            unrec.append({"work": w.wid, "reason": "資料の行に出てこない（委託の中の点検で、管理表は契約1行・報告書に個別の記載なし）"})
    for fy in (2013, 2016, 2019, 2022, 2025):
        unrec.append({"work": None, "fy": fy, "desc": "精密機能検査", "reason": "環境衛生課の別予算（総務係）で発注。修繕執行管理表・各帳票に載らない。新計画の法定点検欄（R7）にだけ出る"})
    mk26 = sg.SCHED["MK2026"]
    unrec.append({"work": None, "fy": 2026, "desc": "令和8年度 定期点検整備（10/13〜11/2 予定）の作業", "reason": "記録時点（2026-11）で日報・報告書・計画の実績が未記入。工程表（予定）だけある"})
    contracts = []
    for c in cw.CONTRACTS.values():
        contracts.append({"id": c.cid, "fy": c.fy, "name": c.name, "short": c.short, "vendor": c.vendor, "type": c.ctype, "account": c.hi,
                          "amount": c.amount, "amount_changed": c.chg, "contract_date": getattr(c, "cdate", None), "site": getattr(c, "site", None),
                          "docs": bool(c.koutei), "works": [w.wid for w in c.works],
                          "amount_note": "町の契約額（税込）。材料・労務・経費込みの請負額で原価ではない" if c.kind != "EN" else "継続費 総額。年割は管理表"})
    world = {"profile": PROFILE, "today": "2026-11-13", "seed": W.SEED, "equipment": eqs, "works": works, "contracts": contracts,
             "links": LINKS, "unrecorded": unrec, "plan_gaps": PLAN_GAPS, "files": sorted(made)}
    with open(os.path.join(ANS, "world.json"), "w", encoding="utf-8") as f:
        json.dump(world, f, ensure_ascii=False, indent=1, default=js)
