# -*- coding: utf-8 -*-
"""answer/ を物差しの形（moyai.sample-world/0）へ書き直す。files/ には触れない。
入力：f09/answer_v1_backup/（旧正解の控え）と生成器の中の世界（world.py・contracts.py。資料は書かない）
出力：full/answer/world.json・world.extra.json・files/*.json
"""
import os, sys, json, glob, datetime as dt
HERE = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, HERE)
import world as W, contracts as C
from plan_gen import u_name
from ledger_gen import spc_name, ACC, motor_of

FULL = os.path.normpath(os.path.join(HERE, ".."))
OLD = os.path.join(FULL, "gen", "_answer_v1")
ANS = os.path.join(FULL, "answer")
FILES = os.path.join(FULL, "files")
SYM_ATTR = {"○": "INSPECT", "点": "INSPECT", "◇": "INSPECT", "法": "INSPECT", "◎": "OVERHAUL", "整": "OVERHAUL", "△": "OVERHAUL", "補": "OVERHAUL",
            "●": "REPLACE", "更": "REPLACE", "改": "UPGRADE", "◆": "UPGRADE"}
FUR = {0: "共通", 1: "1", 2: "2"}

def fy_of(d):
    d = dt.date.fromisoformat(d)
    return d.year if d.month >= 4 else d.year - 1

def main():
    old = json.load(open(os.path.join(OLD, "world.json"), encoding="utf-8"))
    # ---------------- assets
    by_tag = {e.tag: e for e in W.EQS}
    aid = {e.tag: e.spc_no for e in W.EQS}
    assets = []
    for e in W.EQS:
        fl = {0: "共通", 1: "1号炉", 2: "2号炉"}[e.furn]
        assets.append({"id": e.spc_no, "name": e.name, "tag": e.tag, "system": f"{e.group} > {e.sub}", "furnace": FUR[e.furn],
                       "installed": int(e.install[:4]), "class": e.crit, "part": None,
                       "names": {"SPC機器台帳": f"{fl}　{e.group} > {e.sub} > {spc_name(e)}",
                                 "SPC長期修繕計画（sv1）": f"{fl} > {e.group} > {e.sub} > {e.name}",
                                 "SPC長期修繕計画（sv2）": f"{e.group} > {e.sub} > {e.name}",
                                 "組合中長期整備計画": f"{e.group[:-2]}設備 > {e.sub} > {u_name(e)}",
                                 "組合機器台帳": f"{e.group[:-2]} > {e.sub} > {u_name(e)}"}})
        for j, a in enumerate(ACC.get(e.cls, []), 1):
            assets.append({"id": f"{e.spc_no}-a{j}", "name": a, "tag": f"{e.tag}-a{j}", "system": f"{e.group} > {e.sub}", "furnace": FUR[e.furn],
                           "installed": int(e.install[:4]), "class": "付属機器", "part": e.spc_no,
                           "names": {"SPC機器台帳": f"{fl}　{e.group} > {e.sub} > {spc_name(e)} > {a}"}})
        for j, m in enumerate(motor_of(e), 1):
            nm = "電動機" if j == 1 else f"電動機（{j}）"
            assets.append({"id": f"{e.spc_no}-M{j}", "name": nm, "tag": f"{e.tag}-M{j}", "system": f"{e.group} > {e.sub}", "furnace": FUR[e.furn],
                           "installed": int(e.install[:4]), "class": "電動機", "part": e.spc_no,
                           "names": {"SPC機器台帳": f"{fl}　{e.group} > {e.sub} > {spc_name(e)} > {nm}"}})
    # ---------------- works（実績）
    cfy = {c["id"]: c["fy"] for c in old["contracts"]}
    steps_days = {}
    for c in C.ALL:
        if c.kind == "kikan": continue
        for t in c.tasks:
            steps_days[f"{c.cid}:{t.tid}"] = len({s[0] for s in t.steps})
    works = []
    actual_by = {}   # (tag, fy) -> work id
    for w in old["works"]:
        key = w["id"].rsplit(":", 1)[0] if w["tag"] else w["id"]
        if w["contract"].startswith("KK"):
            st = {"完了": "done", "施工中": "in_progress", "未着手": "planned"}[w["status"]]
            days = None
            if w["start"]:
                endd = dt.date.fromisoformat(w["end"]) if w["end"] else W.TODAY
                days = sum(1 for d in W.daterange(dt.date.fromisoformat(w["start"]), endd) if d.weekday() != 6)
            nw = {"id": w["id"], "asset": aid.get(w["tag"]), "fy": cfy[w["contract"]], "status": st, "task": w["work"], "attr": w["attr"],
                  "start": w["start"] or w.get("plannedStart"), "end": w["end"] or (w.get("plannedEnd") if st == "planned" else None),
                  "manDays": w["manDays"], "days": days, "contract": w["contract"]}
        else:
            nw = {"id": w["id"], "asset": aid.get(w["tag"]), "fy": cfy[w["contract"]], "status": "done", "task": w["work"], "attr": w["attr"],
                  "start": w["start"], "end": w["end"], "manDays": w["manDays"], "days": steps_days.get(key), "contract": w["contract"]}
        works.append(nw)
        if w["tag"] and w["attr"]:
            actual_by.setdefault((w["tag"], nw["fy"]), nw["id"])
    # 契約全体の作業（発注の一式・財務の契約行を結ぶ先）
    for c in C.ALL:
        works.append({"id": f"{c.cid}:契約", "asset": None, "fy": c.fy, "status": "done" if c.status == "完了" else "in_progress", "task": c.title,
                      "attr": None, "start": c.start.isoformat(), "end": c.end.isoformat() if c.status == "完了" else None, "manDays": None, "days": None, "contract": c.cid})
    # SPC 保全班の直営（どの資料にも作業としては書いていない実績）
    for i, u in enumerate(old["undocumented"]):
        if not u["tag"] or not u["contract"]: continue
        e = by_tag[u["tag"]]
        wid = f"{u['contract']}:直営:{u['tag']}"
        works.append({"id": wid, "asset": e.spc_no, "fy": cfy[u["contract"]], "status": "done", "task": f"{e.name} SPC 直営 {u['attr']}", "attr": u["attr"],
                      "start": u["period"][0], "end": u["period"][1], "manDays": None, "days": None, "contract": u["contract"], "undocumented": True})
        actual_by.setdefault((u["tag"], cfy[u["contract"]]), wid)
    for u in old["undocumented"]:
        if u["contract"] is None and u["tag"]:
            e = by_tag[u["tag"]]
            works.append({"id": f"P2027:{u['tag']}", "asset": e.spc_no, "fy": 2026, "status": "postponed", "task": f"{e.name} SPC 定期整備分（2027 年度へ繰越）",
                          "attr": u["attr"], "start": None, "end": None, "manDays": None, "days": None})
    wid_set = {w["id"] for w in works}
    synth = {}
    def plan_work(tag, fy, sym, label):
        """星取表の印の行き先。実績があればそれ、無ければ予定（planned／not_done）の作業を作る"""
        if tag and (tag, fy) in actual_by:
            return actual_by[(tag, fy)]
        key = (tag or label, fy, SYM_ATTR.get(sym))
        if key in synth: return synth[key]
        st = "planned" if fy >= 2026 else "not_done"
        if tag is None and fy < 2026: st = "done"          # 法定点検（外部機関が実施、資料は計画の行のみ）
        if tag and fy == 2026 and by_tag[tag].furn == 1 and SYM_ATTR.get(sym) in ("OVERHAUL", "REPLACE"):
            st = "postponed"
        wid = f"PLAN:{tag or label}:{fy}:{SYM_ATTR.get(sym)}"
        works.append({"id": wid, "asset": aid.get(tag), "fy": fy, "status": st, "task": (by_tag[tag].name if tag else label) + f" {sym}",
                      "attr": SYM_ATTR.get(sym), "start": None, "end": None, "manDays": None, "days": None})
        synth[key] = wid
        return wid
    # ---------------- files/*.json と links
    links = []
    old_links = {}
    for l in old["links"]:
        old_links.setdefault((l["file"], l.get("sheet"), l.get("r", l.get("seq")), "seq" in l), []).append(l["work"])
    os.makedirs(os.path.join(ANS, "files"), exist_ok=True)
    name_tag = {}
    for e in W.EQS:
        name_tag[(u_name(e), {0: "共通", 1: "1号", 2: "2号"}[e.furn])] = e.tag
        name_tag[(e.name, FUR[e.furn])] = e.tag
    files_list = sorted(os.listdir(FILES))
    for fn in files_list:
        stem = os.path.splitext(fn)[0]
        a = json.load(open(os.path.join(OLD, "files", stem + ".json"), encoding="utf-8"))
        isplan = a["plan"]; ispdf = fn.endswith(".pdf"); iscsv = fn.endswith(".csv")
        out = []
        for rec in a["records"]:
            f = dict(rec["fields"])
            nr = {}
            if ispdf:
                nr = {"sheet": None, "r": rec["seq"] + 1, "seq": rec["seq"]}
            elif iscsv:
                nr = {"sheet": None, "r": rec["r"]}
            else:
                nr = {"sheet": rec["sheet"], "r": rec["r"]}
            nr["kind"] = rec["kind"]
            if isplan:
                if rec["kind"] in ("parent", "item"):
                    if "組合" in fn:
                        plans = f.pop("marks", {}) or {}
                        marks = f.pop("actual", {}) or {}
                        f["marks"] = marks; f["plans"] = plans
                    else:
                        f["marks"] = f.get("marks", {}) or {}
                        f["plans"] = dict(f["marks"])
                    f.setdefault("path", [])
                    if rec["kind"] == "item":
                        tag = f.get("code")
                        if not tag and "組合" in fn and f["path"] and f["path"][0] != "法定点検・測定":
                            tag = name_tag.get((f["path"][-1], f.get("furnace")))
                        label = None if tag else "法定:" + f["path"][-1]
                        yrs = set(f["plans"]) | set(f["marks"])
                        for y in sorted(yrs):
                            sym = f["marks"].get(y) or f["plans"].get(y)
                            w = plan_work(tag, int(y), sym, label)
                            links.append({"file": fn, "sheet": nr["sheet"], "r": nr["r"], "work": w, "year": int(y), "how": "direct"})
            else:
                if "manDays" in f and not isinstance(f["manDays"], (int, float)):
                    f.pop("manDays")
                if f.get("people") and "date" in f and "manDays" in f and rec["kind"] == "item":
                    f["days"] = 1
                if isinstance(f.get("people"), dict) and rec["kind"] in ("total", "subtotal"):
                    f["days"] = len(f["people"])
            nr["fields"] = f
            out.append(nr)
            key = (fn, rec.get("sheet") if not (ispdf or iscsv) else None, rec.get("seq") if ispdf else rec["r"], ispdf)
            for w in old_links.get(key, []):
                how = "direct"
                if w.endswith(":外注一式"):
                    w = w.split(":")[0] + ":契約"; how = "bundle"
                elif w.endswith(":契約"):
                    how = "bundle"
                if w not in wid_set:
                    w = w.split(":")[0] + ":契約"; how = "bundle"
                lk = {"file": fn, "sheet": nr["sheet"], "r": nr["r"], "work": w, "how": how}
                if "date" in f: lk["day"] = f["date"]
                links.append(lk)
        json.dump({"file": fn, "plan": isplan, "records": out}, open(os.path.join(ANS, "files", stem + ".json"), "w", encoding="utf-8"), ensure_ascii=False, indent=0)
    # 工程表の節目・炉状態の行は契約全体の行（annual）として結ぶ
    world = {"schema": "moyai.sample-world/0",
             "facility": {"id": "F09", "name": W.FAC},
             "profile": {"furnace": "流動床式ガス化溶融炉（ガス化炉＋旋回溶融炉）95t/24h×2炉", "furnaces": ["1", "2"], "common": "共通"},
             "maker": {"today": "2026-11-13"},
             "assets": assets, "works": works, "links": links, "errata": [],
             "files": [{"file": fn} for fn in files_list]}
    json.dump(world, open(os.path.join(ANS, "world.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=0)
    extra = {k: v for k, v in old.items() if k not in ("works", "links", "equipment")}
    extra["equipment_v1"] = old["equipment"]
    extra["notes"] = ["星取表：SPC の長期修繕計画は予定と実績を分けない様式（過去の年度も当初計画の印を残す）ので marks＝plans＝行に書いてある記号。",
                      "組合の中長期整備計画は同じ行に計画・実績の列があるので、marks＝実績列、plans＝計画列。",
                      "星取表の印は、同じ機器・同じ年度の実際の作業があればそれへ、無ければ予定の作業（2026 年度以降 planned、2025 年度以前 not_done、1号炉 2026 年度の整備は postponed）へ結んだ。法定点検の行は外部の検査（done）へ。",
                      "SPC 保全班の直営作業は works に undocumented: true で入れた（資料には作業として書いていない）。",
                      "SPC 台帳の付属機器・電動機の行は assets に part（親の id）付きで入れた。機器 648 台には数えない。",
                      "errata：omitMark・lastYearUnreflected に当たるわざとの欠けは無い。ほかの欠けは planGaps に残す。",
                      "PDF の記録は r＝seq＋1。工数 manDays は人日（1 人日＝8 時間）。"]
    json.dump(extra, open(os.path.join(ANS, "world.extra.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=0)
    print(len(assets), "assets", len(works), "works", len(links), "links")

if __name__ == "__main__":
    main()
