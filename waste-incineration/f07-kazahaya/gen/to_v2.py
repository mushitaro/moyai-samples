# -*- coding: utf-8 -*-
"""answer/ を物差しの形（moyai.sample-world/0）に書き直す変換器。files/ は読まない・書かない。
入力：f07/answer_v1_backup/answer/（旧正解の控え）と生成器の正解データ（world・contracts_full・works、メモリ上だけ）。
出力：full/answer/world.json、full/answer/world.extra.json、full/answer/files/*.json
  python3 to_v2.py
当て方（星取表）：
 - sp1 様式（2015–2019 年度版、1 年度 1 列で予定だけ）：その行の記号＝予定。marks＝plans＝書いてある記号。
 - sp2 様式（2020–2026 年度版、年度ごとに「計画」「実績」の 2 列が同じ行）：marks＝実績の列、plans＝計画の列。
 - 法定点検の行・DCS 一式の行も同じ決まり。記号は 1 字なので 8 字超えは無い。
工数：日報の旧 manDays は人×時間（h）。人日＝h÷8 に直し、書いてあった字は manDaysText に残す。"""
import glob
import json
import os
import shutil
import sys
from datetime import date, timedelta

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
ROOT = os.path.dirname(HERE)
OLD = os.path.join(ROOT, "gen", "_answer_v1")
NEW = os.path.join(ROOT, "answer")

from world import ALL, BY_CODE, rule, actual  # noqa: E402
from contracts_full import get, off, TODAY  # noqa: E402
from works import WORKS  # noqa: E402


def fy_of(s):
    d = date.fromisoformat(s)
    return d.year if d.month >= 4 else d.year - 1


def wdays(a, b, holi):
    n, d = 0, a
    while d <= b:
        if not off(d, holi):
            n += 1
        d += timedelta(days=1)
    return max(n, 1)


def main():
    old_world = json.load(open(os.path.join(OLD, "world.json"), encoding="utf-8"))
    files = sorted(os.path.basename(p) for p in glob.glob(os.path.join(ROOT, "files", "*")))
    cmap = {c["key"]: c for c in get()}
    # ---------------- works（実績）
    works = []
    for w in WORKS:
        c = cmap.get(w["contract_key"])
        days = 1
        if c:
            t = [t for t in c["tasks"] if t["no"] == w["task_no"]][0]
            end = t["ae"] if not c.get("cut") else min(t["ae"], c["cut"])
            days = wdays(t["as_"], end, c["holi"])
        works.append(dict(id=w["id"], asset=w["code"], fy=fy_of(w["start"]), status="done" if w["status"] == "完了" else "in_progress",
                          task=w["work"], attr=w["attr"], start=w["start"], end=w["end"], manDays=w["manDays"], days=days,
                          vendor=w["vendor"], contract=w["contract"]))
    done_key = {}
    for w in works:
        done_key.setdefault((w["asset"], w["fy"], w["task"]), w["id"])
    done_cy = {}
    for w in works:
        done_cy.setdefault((w["asset"], w["fy"]), []).append(w["id"])
    # 予定の作業（計画の印に対応する、まだ実施していない・しなかった作業）
    ATTR = {"I": "INSPECT", "L": "INSPECT", "O": "OVERHAUL", "R": "OVERHAUL", "X": "REPLACE", "U": "UPGRADE"}
    planned = {}

    def plan_work(code, y, t, wname):
        k = (code, y, wname)
        if k in done_key:
            return done_key[k]
        if k not in planned:
            if y >= 2027 or (y == 2026):
                st = "planned"
            else:
                st = "not_done"
            if code == "2BO-180" and y == 2022:
                st = "postponed"
            pid = f"P{y}-{code}-{t}{len(planned):04d}"
            planned[k] = dict(id=pid, asset=code, fy=y, status=st, task=wname, attr=ATTR[t], start=None, end=None, manDays=None, days=None)
        return planned[k]["id"]

    for e in ALL:
        for ver, rev in [(v, False) for v in range(2015, 2020)] + [(v, True) for v in range(2020, 2027)]:
            for y, lst in rule(e, rev, ver).items():
                if y < ver:
                    continue
                for (t, wname) in lst:
                    plan_work(e["code"], y, t, wname)
    # ---------------- answer/files
    pidx = {}
    for x in planned.values():
        pidx.setdefault((x["asset"], x["fy"]), []).append(x["id"])
    links = []
    newfiles = os.path.join(NEW, "files")
    if os.path.exists(newfiles):
        shutil.rmtree(newfiles)
    os.makedirs(newfiles)
    for p in sorted(glob.glob(os.path.join(OLD, "files", "*.json"))):
        a = json.load(open(p, encoding="utf-8"))
        fname = a["file"]
        is_plan = a.get("plan", False)
        sp2 = "sp2星取表" in fname
        recs = []
        for r in a["records"]:
            f = dict(r["fields"])
            nr = {k: v for k, v in r.items() if k in ("sheet", "r", "seq", "kind")}
            if fname.lower().endswith((".csv", ".pdf")):
                nr.pop("sheet", None)
            if "seq" in nr and "r" not in nr:
                nr["r"] = nr["seq"] + 1
            if is_plan and r["kind"] in ("item", "parent"):
                old_m = f.pop("marks", {}) or {}
                old_a = f.pop("marksActual", {}) or {}
                if sp2:
                    f["marks"], f["plans"] = old_a, old_m
                else:
                    f["marks"], f["plans"] = dict(old_m), dict(old_m)
                f.setdefault("path", [])
                code = f.get("code")
                if code and code in BY_CODE:
                    for y, m in f["plans"].items():
                        yy = int(y)
                        cands = done_cy.get((code, yy)) or pidx.get((code, yy), [])
                        for wid in cands[:3]:
                            links.append(dict(file=fname, sheet=r.get("sheet"), r=r["r"], work=wid, year=yy, how="direct"))
                    for y, m in f["marks"].items():
                        if not sp2:
                            continue
                        for wid in done_cy.get((code, int(y)), []):
                            links.append(dict(file=fname, sheet=r.get("sheet"), r=r["r"], work=wid, year=int(y), how="direct"))
            if "manDays" in f and "作業日報" in fname:
                h = f.get("manDays")
                f["manDaysText"] = f.get("manDaysText") or ("" if h is None else f"{h:g}")
                f["manDays"] = None if h is None else round(h / 8, 4)
                f.pop("manDaysUnit", None)
                if h is not None:
                    nr["_hours"] = h
            nr["fields"] = f
            recs.append(nr)
        out = {"file": fname, "plan": is_plan, "records": recs}
        with open(os.path.join(newfiles, os.path.basename(p)), "w", encoding="utf-8") as fh:
            json.dump(out, fh, ensure_ascii=False, indent=0)
    # 旧 links（工程表・日報・報告書・発注）を 1 作業 1 行に
    day_of = {}
    for p in glob.glob(os.path.join(newfiles, "*作業日報*.json")):
        a = json.load(open(p, encoding="utf-8"))
        for r in a["records"]:
            if "date" in r["fields"]:
                day_of[(a["file"], r["sheet"], r["r"])] = r["fields"]["date"]
    for l in old_world["links"]:
        f = l["file"]
        if "中長期整備計画" in f:
            continue
        if "発注台帳" in f:
            how = "bundle"
        else:
            how = "direct"
        for wid in l.get("works", []):
            d = dict(file=f, sheet=None if f.endswith((".csv", ".pdf")) else l.get("sheet"), r=l.get("r"), work=wid, how=how)
            if "seq" in l:
                d["seq"] = l["seq"]
                d["r"] = l["seq"] + 1
            if (f, l.get("sheet"), l.get("r")) in day_of:
                d["day"] = day_of[(f, l.get("sheet"), l.get("r"))]
            links.append(d)
    # 発注台帳の年間委託（annual）と支給品（rollup）を区別
    orders = {}
    for p in glob.glob(os.path.join(newfiles, "*発注台帳*.json")):
        a = json.load(open(p, encoding="utf-8"))
        for r in a["records"]:
            orders[(a["file"], r["r"])] = r["fields"]
    for l in links:
        if "発注台帳" in l["file"]:
            o = orders.get((l["file"], l["r"]), {})
            if o.get("kind") == "委託":
                l["how"] = "annual"
            elif o.get("kind") in ("購入",):
                l["how"] = "rollup"
            else:
                l["how"] = "bundle"
    # ---------------- errata（印の抜け）
    errata = []
    for p in glob.glob(os.path.join(newfiles, "*sp1長期修繕計画表.json")):
        a = json.load(open(p, encoding="utf-8"))
        for r in a["records"]:
            c = r["fields"].get("code")
            if c in ("0UK-010", "0UK-020", "0UK-030"):
                errata.append(dict(type="omitMark", file=a["file"], sheet=r["sheet"], r=r["r"], year=2020,
                                   works=done_cy.get((c, 2020), []), note="計量機 定期検査（計量法 2年毎）の 2020 年度の印の抜け"))
    # ---------------- assets
    from gen_report_full import maker_name
    import gen_ledger_full as L
    L.ASOF, L.VER = date(2026, 11, 5), 2026
    UN = {"1": "1", "2": "2", "0": "共通"}
    UNP = {"1": "1号炉", "2": "2号炉", "0": "共通"}
    assets = []
    for e in ALL:
        iy, im = L.install_ym(e)
        names = {"機器台帳（PM-NAVI CSV）": e["name"],
                 "中長期整備計画": " > ".join([e["cat"], UNP[e["unit"]], e["sub"], e["name"]]),
                 "工事報告書（メーカー様式）": maker_name(e["code"]),
                 "工事報告書（SPC様式）・作業日報": e["name"]}
        assets.append(dict(id=e["code"], name=e["name"], tag=e["code"], system=e["cat"] + " > " + e["sub"], furnace=UN[e["unit"]],
                           installed=iy, **{"class": e["imp"]}, part=e["motor_of"], names=names, kind=e["kind"]))
    pw = list(planned.values())
    world = {
        "schema": "moyai.sample-world/0",
        "facility": {"id": "F07", "name": "【架空】かざはやエコプラザ"},
        "profile": dict(old_world["profile"]),
        "maker": {"today": "2026-11-13"},
        "assets": assets,
        "works": works + pw,
        "links": links,
        "errata": errata,
        "files": [{"file": f} for f in files],
    }
    extra = {k: v for k, v in old_world.items() if k not in ("profile", "works", "links")}
    extra["works_v1"] = "旧 works は answer_v1_backup/answer/world.json"
    extra["notes"] = __doc__
    with open(os.path.join(NEW, "world.json"), "w", encoding="utf-8") as fh:
        json.dump(world, fh, ensure_ascii=False, indent=0)
    with open(os.path.join(NEW, "world.extra.json"), "w", encoding="utf-8") as fh:
        json.dump(extra, fh, ensure_ascii=False, indent=0)
    print("works", len(works), "planned", len(pw), "links", len(links), "errata", len(errata))


if __name__ == "__main__":
    main()
