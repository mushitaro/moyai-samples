# -*- coding: utf-8 -*-
"""answer/ を物差しの形（moyai.sample-world/0）に書き直す。files/ には触らない（読みもしない）。
入力：answer_v1（引数で指定、既定は ../../answer_v1_backup）と生成器の正解データ（core：ファイルは書かない）
出力：answer/files/*.json、answer/world.json、answer/world.extra.json
"""
import json
import os
import sys
from collections import defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import core
import s_ledger as L

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
V1 = sys.argv[1] if len(sys.argv) > 1 else os.path.join(BASE, "gen", "_answer_v1")
OUTD = os.path.join(BASE, "answer")
KIND_EN = {"I": "INSPECT", "O": "OVERHAUL", "R": "REPLACE", "U": "UPGRADE"}
UN = {"1": "1", "2": "2", "3": "3", "C": "共通"}


def load(p):
    with open(p, encoding="utf-8") as f:
        return json.load(f)


def short(s):
    return {k: v for k, v in s.items() if isinstance(v, str) and len(v) <= 8}


def convert_file(a, fy):
    fn = a["file"]
    is_pdf_csv = fn.endswith((".pdf", ".csv"))
    plan = a.get("plan", False)
    h3 = "施設保全計画様式" in fn
    recs = []
    errata = []
    for rc in a["records"]:
        n = dict(rc)
        if is_pdf_csv:
            n["sheet"] = None
        if n.get("seq") is not None and n.get("r") is None:
            n["r"] = n["seq"] + 1
        f = dict(n.get("fields", {}))
        if "_check" in f:  # 念のため（v1 では記録の直下に置いていた）
            n["_check"] = f.pop("_check")
        if plan:
            if n["kind"] == "item" and "marks" not in f:
                n["kind"] = "subtotal"  # 費用集計の行（機器ではない）
            if n["kind"] in ("parent", "item"):
                f.setdefault("path", [])
                marks = f.get("marks") or {}
                if h3:
                    plans = dict(marks)
                    if f.get("code") is None and f.get("path", [None])[0] == "法定点検・検査":
                        act = {y: s for y, s in plans.items() if int(y) < fy and int(y) != 2017}
                    else:
                        act = dict(f.get("actual") or {})
                    f["plans"] = short(plans)
                    f["marks"] = short(act)
                    f.pop("actual", None)
                    if f.get("path", [None])[0] == "法定点検・検査" and "精密機能検査" in f["path"][-1]:
                        errata.append(dict(type="omitMark", file=fn, sheet=n["sheet"], r=n["r"], year=2019, works=[],
                                           note="精密機能検査：H28 の次を H32 と入力（H31 の印が抜け、3 年を超える）"))
                else:
                    f["marks"] = short(marks)
                    f["plans"] = short(marks)
                if n["kind"] == "parent":
                    f["marks"], f["plans"] = {}, {}
        else:
            if n["kind"] in ("total", "subtotal"):
                for k in ("manDays",):
                    if k in f:
                        f["manDaysTotal"] = f.pop(k)
            if "days" not in f:
                for k in ("日数",):
                    if isinstance(f.get(k), (int, float)):
                        f["days"] = f[k]
                if "planStart" in f and "planEnd" in f and f.get("planStart") and f.get("planEnd") and "days" not in f:
                    import datetime as dt
                    a0 = dt.date.fromisoformat(f.get("actStart") or f["planStart"])
                    b0 = dt.date.fromisoformat(f.get("actEnd") or f["planEnd"])
                    f["days"] = (b0 - a0).days + 1
        n["fields"] = f
        recs.append(n)
    out = dict(file=fn, plan=plan, records=recs)
    return out, errata


def main():
    W = core.get()
    w1 = load(os.path.join(V1, "world.json"))
    cat = {c["file"]: c for c in w1["catalog"]}
    os.makedirs(os.path.join(OUTD, "files"), exist_ok=True)
    errata = []
    recidx = {}
    files = []
    for fn in sorted(os.listdir(os.path.join(V1, "files"))):
        a = load(os.path.join(V1, "files", fn))
        fy = cat.get(a["file"], {}).get("年度") or 0
        b, er = convert_file(a, fy)
        errata += er
        files.append(a["file"])
        for rc in b["records"]:
            recidx[(a["file"], rc.get("sheet"), rc.get("r"), rc.get("seq"))] = rc
        with open(os.path.join(OUTD, "files", fn), "w", encoding="utf-8") as f:
            json.dump(b, f, ensure_ascii=False, indent=0)
    # 作業
    task_of = {t["id"]: t for t in W.tasks}
    days_of = defaultdict(set)
    for ln in W.lines:
        for w in ln["works"]:
            days_of[id(w)].add(ln["date"])
    works = []
    wid_by = defaultdict(list)
    for w in W.works:
        t = task_of.get(w["task"])
        status = "done" if w["status"] == "done" else ("done" if w["status"] == "in_progress" else "planned")
        works.append(dict(id=w["id"], asset=w["tag"], fy=w["fy"], status=status,
                          task=(t["name"] if t else W.contracts[w["contract"]]["name"]) + (f"（{w['note']}）" if w["note"] else ""),
                          attr=KIND_EN[w["kind"]], start=w["start"], end=w["end"],
                          manDays=round(w["manHours"] / 8, 2) if w["manHours"] else None,
                          days=len(days_of[id(w)]) or None, contract=w["contract"]))
        wid_by[(w["tag"], w["fy"])].append(w["id"])
    # 結び
    links = []
    contract_kind = {c["id"]: c["kind"] for c in W.contracts.values()}
    for l in w1["links"]:
        fn = l["file"]
        is_pdf_csv = fn.endswith((".pdf", ".csv"))
        sheet = None if is_pdf_csv else l.get("sheet")
        r_ = l.get("r") if l.get("r") is not None else (l["seq"] + 1 if l.get("seq") is not None else None)
        rc = recidx.get((fn, sheet, r_, l.get("seq")))
        fld = (rc or {}).get("fields", {})
        n = len(l["work_ids"])
        how = "direct"
        if fn.endswith(".csv"):
            how = "upgrade" if "基幹的設備改良" in str(fld.get("件名", "")) else "bundle"
        elif "年間工程" in fn and sheet == "年間工程":
            how = "rollup"
        elif fn.endswith(".pdf") and (sheet is None) and ("区分" in fld or "台数" in fld):
            how = "rollup"
        elif "工程表" in fn and n > 1:
            how = "rollup"
        for wid in l["work_ids"]:
            d = dict(file=fn, sheet=sheet, r=r_, work=wid, how=how)
            if l.get("seq") is not None:
                d["seq"] = l["seq"]
            if fld.get("date") and "作業日報" in fn:
                d["day"] = fld["date"]
            links.append(d)
    # 星取表の印 → 作業（同じ機器・同じ年度の作業）
    for fn in files:
        a = load(os.path.join(OUTD, "files", fn.rsplit(".", 1)[0] + ".json"))
        if not a["plan"]:
            continue
        for rc in a["records"]:
            if rc["kind"] != "item":
                continue
            tag = rc["fields"].get("code")
            if not tag:
                continue
            for y in sorted(set(rc["fields"].get("marks", {})) | set(rc["fields"].get("plans", {}))):
                for wid in wid_by.get((tag, int(y)), []):
                    links.append(dict(file=fn, sheet=rc["sheet"], r=rc["r"], work=wid, year=int(y),
                                      how="direct"))
            # H29（2017）の実績欄の空き（担当交代の追記漏れ）
            if "施設保全計画様式" in fn and cat[fn]["年度"] >= 2018 and wid_by.get((tag, 2017)):
                errata.append(dict(type="lastYearUnreflected", file=fn, sheet=rc["sheet"], r=rc["r"], year=2017,
                                   works=wid_by[(tag, 2017)], note="H29 の実績欄が空（h3→h4 の引継ぎ年）"))
    assets = []
    for e in W.eq:
        assets.append(dict(id=e["tag"], name=e["name"], tag=e["tag"], system=f"{e['setsubi']} > {e['keito']}", furnace=UN[e["unit"]],
                           installed=e["installed"], class_=None, part=None,
                           names={"機器台帳": f"{L.dwg(e)[0]} {L.dwg(e)[1]} > {e['name']}", "竣工図書機器表": f"{e['setsubi']} > {e['name']}",
                                  "中長期整備計画": f"{e['setsubi']} > {e['keito']} > {e['name']}", "工事報告書": e["name"]}))
        assets[-1]["class"] = e["imp"]
        del assets[-1]["class_"]
    world = {
        "schema": "moyai.sample-world/0",
        "facility": {"id": "F06", "name": "【架空】みなと臨海工場"},
        "profile": {"furnace": "並行揺動式ストーカ（水冷火格子の一部採用）・全連続 200t/24h×3炉"},
        "maker": {"today": "2026-11-13"},
        "assets": assets, "works": works, "links": links, "errata": errata,
        "files": [{"file": f} for f in files],
    }
    with open(os.path.join(OUTD, "world.json"), "w", encoding="utf-8") as f:
        json.dump(world, f, ensure_ascii=False, indent=0)
    extra = {k: v for k, v in w1.items() if k not in ("links", "works")}
    extra["works_v1"] = w1["works"]
    with open(os.path.join(OUTD, "world.extra.json"), "w", encoding="utf-8") as f:
        json.dump(extra, f, ensure_ascii=False, indent=0)
    print("files", len(files), "works", len(works), "links", len(links), "errata", len(errata))


if __name__ == "__main__":
    main()
