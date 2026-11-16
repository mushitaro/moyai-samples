# -*- coding: utf-8 -*-
"""answer/ を物差しの形（moyai.sample-world/0）にそろえる。files/ は読まない・書かない。
 入力：answer_v1（今の answer/ の写し。無ければ answer/ を読む）と、生成器の当初計画データ（world.CPLAN、ファイルは書かない）
 出力：answer/world.json、answer/world.extra.json、answer/files/*.json
 python3 to_v2.py"""
import os, sys, json, glob, shutil, datetime as dt
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
FULL = os.path.abspath(os.path.join(HERE, ".."))
ANS = os.path.join(FULL, "answer")
SRC = os.path.join(FULL, "gen", "_answer_v1")  # 元の正解（v1）を残す場所

import world as W  # 当初計画（CPLAN）と機器だけ使う。ファイルは書かない

def load_src():
    if not os.path.isdir(SRC):
        shutil.copytree(ANS, SRC)
    world = json.load(open(os.path.join(SRC, "world.json"), encoding="utf-8"))
    files = {}
    for p in sorted(glob.glob(os.path.join(SRC, "files", "*.json"))):
        d = json.load(open(p, encoding="utf-8"))
        files[os.path.basename(p)] = d
    return world, files

FURN = {0: "共通", 1: "1", 2: "2"}
ORIG_PLAN = "【架空】しおさい町ごみ処理場_中長期整備計画_H27策定長寿命化計画_コンサル策定版.xlsx"
OLD_PLAN = "【架空】しおさい町ごみ処理場_中長期整備計画_H27策定長寿命化計画_コンサル様式t1t2追記.xlsx"
NEW_PLAN = "【架空】しおさい町ごみ処理場_中長期整備計画_R5〜R15_t3計画実績2列様式.xlsx"

def mid_of(path_last, gou):
    f = {"共通": 0, "1号": 1, "2号": 2}.get(gou, 0)
    try: return W.find(path_last, f).mid
    except KeyError: return None

def cplan_syms(mid):
    return {str(fy): s for fy, (s, _) in W.CPLAN.get(mid, {}).items()} if mid else {}

def is_mark(v):
    return v not in (None, "") and len(str(v)) <= 8

def conv_plan(d, extra_rows):
    """星取表の fields を path・marks・plans にそろえる"""
    fn = d["file"]
    for rc in d["records"]:
        f = rc["fields"]
        if rc["kind"] not in ("item", "parent"):
            continue
        f.setdefault("path", [])
        m = f.get("marks", {}) or {}
        if fn == NEW_PLAN:
            marks, plans = {}, {}
            for y, v in m.items():
                a, _, b = str(v).partition("/")
                if is_mark(a): plans[y] = a
                if is_mark(b): marks[y] = b
            rc["_written"] = m  # 書いてある 計/実 の字（比べない）
            f["marks"], f["plans"] = marks, plans
        elif fn == OLD_PLAN:
            mid = mid_of(f["path"][-1], f.get("号炉")) if rc["kind"] == "item" and f["path"] else None
            rc["_mid"] = mid
            f["marks"] = {y: v for y, v in m.items() if is_mark(v)}
            f["plans"] = {y: v for y, v in cplan_syms(mid).items() if is_mark(v)} if mid and not str(f.get("code", "")).endswith("-2") else {}
        else:  # 策定版：計画の行だけ
            mid = mid_of(f["path"][-1], f.get("号炉")) if rc["kind"] == "item" and f["path"] else None
            rc["_mid"] = mid
            f["marks"] = {y: v for y, v in m.items() if is_mark(v)}
            f["plans"] = dict(f["marks"])
        if rc["kind"] == "item" and fn != OLD_PLAN and fn != ORIG_PLAN:
            rc["_mid"] = f.get("管理番号")
    return d

def conv_other(d):
    pdf = d["file"].endswith(".pdf")
    for rc in d["records"]:
        if pdf:
            rc.pop("sheet", None)
            if "seq" in rc and "r" not in rc: rc["r"] = rc["seq"] + 1
        f = rc["fields"]
        if rc["kind"] in ("total", "subtotal") and "manDays" in f:
            rc["_check"] = {"manDays": f.pop("manDays")}  # 合計は物差しに数えさせない
        if "manDays" in f and not isinstance(f["manDays"], (int, float)):
            f.pop("manDays")
        if "日報" in d["file"] and "工事別" in d["file"] and rc["kind"] == "item" and "manDays" in f:
            f["days"] = 1  # 工事別日報は 1 日 1 行
    return d

def workdays(a, b):
    if not a or not b: return None
    a = dt.date.fromisoformat(a); b = dt.date.fromisoformat(b)
    n = 0; d = a
    while d <= b:
        if d.weekday() != 6: n += 1
        d += dt.timedelta(days=1)
    return n

def main():
    w1, files = load_src()
    os.makedirs(os.path.join(ANS, "files"), exist_ok=True)
    for p in glob.glob(os.path.join(ANS, "files", "*.json")): os.remove(p)
    byfile = {}
    for name, d in files.items():
        d = conv_plan(d, None) if d.get("plan") else conv_other(d)
        byfile[d["file"]] = d
        with open(os.path.join(ANS, "files", name), "w", encoding="utf-8") as fo:
            json.dump(d, fo, ensure_ascii=False, indent=1)
    # ---- 資産 ----
    assets = []
    for e in w1["equipment"]:
        names = {}
        for fn, nm in (e.get("rows") or {}).items():
            if nm: names[fn] = "%s > %s" % (e["system"], nm)
        names[OLD_PLAN] = "%s > %s" % (e["system"], e["aliases"].get("旧計画（コンサル）", e["name"]))
        if not e["added_2015"]: names[ORIG_PLAN] = names[OLD_PLAN]
        names[NEW_PLAN] = "%s > %s" % (e["system"], e["name"])
        for k, v in e["aliases"].items():
            if k not in ("台帳2024", "竣工図書・台帳2015", "旧計画（コンサル）"): names[k] = v
        assets.append({"id": e["mid"], "name": e["name"], "tag": e["drawing_no"], "system": e["system"], "furnace": FURN[e["furnace"]],
                       "installed": e["installed_fy"], "class": e["importance"], "part": None, "names": names})
    # ---- 結び（v1）に年度・日・種類を足す ----
    wk = {x["id"]: x for x in w1["works"]}
    ctype = {c["id"]: c for c in w1["contracts"]}
    rec_index = {}
    for fn, d in byfile.items():
        for rc in d["records"]:
            key = (fn, rc.get("sheet") if not fn.endswith(".pdf") else None, rc.get("r") if "r" in rc else None)
            rec_index.setdefault(key, rc)
            if fn.endswith(".pdf"): rec_index[(fn, None, rc.get("seq"), "seq")] = rc
    links = []
    for l in w1["links"]:
        fn = l["file"]; w = wk[l["work"]]
        if "seq" in l:
            rc = rec_index.get((fn, None, l["seq"], "seq")); r = l["seq"] + 1; sheet = None
        else:
            rc = rec_index.get((fn, l["sheet"], l["r"])); r = l["r"]; sheet = l["sheet"]
        nl = {"file": fn, "sheet": sheet, "r": r, "work": l["work"], "how": "direct"}
        if byfile[fn].get("plan"): nl["year"] = w["fy"]
        if "作業日報" in fn and rc:
            dd = rc["fields"].get("日付") or rc["fields"].get("月日")
            if dd: nl["day"] = dd
        if "発注台帳" in fn:
            c = ctype.get(w["contract"])
            if c:
                if c["id"].startswith(("ELEC", "ANAL", "ENV", "FIRE", "EV", "JOKA", "SCALE")): nl["how"] = "annual"
                elif c["id"] == "EN": nl["how"] = "upgrade"
                elif len(c["works"]) > 1: nl["how"] = "bundle"
        if "_written" in (rc or {}) and False: pass
        links.append(nl)
    # ---- 作業：実績 ----
    md = {}; days = {}
    for l in links:
        if "作業日報" not in l["file"]: continue
        rc = rec_index.get((l["file"], l["sheet"], l["r"]))
        if not rc or "manDays" not in rc["fields"]: continue
        n = sum(1 for x in links if x["file"] == l["file"] and x["sheet"] == l["sheet"] and x["r"] == l["r"])
        md[l["work"]] = md.get(l["work"], 0) + rc["fields"]["manDays"] / max(1, n)
        if l.get("day"): days.setdefault(l["work"], set()).add(l["day"])
    works = []
    for x in w1["works"]:
        works.append({"id": x["id"], "asset": x["mid"], "fy": x["fy"], "status": "done", "task": x["desc"], "attr": x["kind"],
                      "start": x["start"], "end": x["end"], "manDays": round(md[x["id"]], 2) if x["id"] in md else (round(x["manHours"] / 8.0, 2) if x.get("manHours") else None),
                      "days": len(days[x["id"]]) if x["id"] in days else workdays(x["start"], x["end"]), "vendor": x["vendor"], "contract": x["contract"]})
    # ---- 作業：予定・繰延べ・計画の印だけ（星取表から） ----
    done_by = {}
    for x in w1["works"]: done_by.setdefault((x["mid"], x["fy"]), []).append(x["id"])
    pid = 0
    shared = {}
    def P(mid, fy, status, task, attr, fn, rc, key=None):
        nonlocal pid
        if key and key in shared:
            links.append({"file": fn, "sheet": rc["sheet"], "r": rc["r"], "work": shared[key], "year": int(fy), "how": "direct"}); return
        pid += 1
        i = "P%04d" % pid
        if key: shared[key] = i
        works.append({"id": i, "asset": mid, "fy": int(fy), "status": status, "task": task, "attr": attr, "start": None, "end": None, "manDays": None, "days": None})
        links.append({"file": fn, "sheet": rc["sheet"], "r": rc["r"], "work": i, "year": int(fy), "how": "direct"})
    A = {"●": "REPLACE", "◎": "OVERHAUL", "○": "OVERHAUL", "△": "INSPECT", "更": "REPLACE", "整": "OVERHAUL", "補": "OVERHAUL", "点": "INSPECT"}
    errata = []
    for fn in (ORIG_PLAN, OLD_PLAN, NEW_PLAN):
        for rc in byfile[fn]["records"]:
            if rc["kind"] != "item": continue
            mid = rc.get("_mid")
            if not mid or mid not in W.BYMID: continue
            f = rc["fields"]
            for y, s in f.get("plans", {}).items():
                fy = int(y); sym = s[0]
                done = done_by.get((mid, fy), [])
                if fn == NEW_PLAN:
                    act = f["marks"].get(y, "")
                    if act == "済": continue
                    if act.startswith("→"): P(mid, fy, "postponed", "%s（計画）" % {"更": "更新", "整": "整備", "補": "補修", "点": "点検"}[sym], A[sym], fn, rc)
                    elif fy >= 2026: P(mid, fy, "planned", "%s（計画）" % {"更": "更新", "整": "整備", "補": "補修", "点": "点検"}[sym], A[sym], fn, rc)
                    else: P(mid, fy, "not_done", "（計画の印のみ）", A[sym], fn, rc)
                else:
                    # コンサルの当初計画：策定版と実績追記版で同じ予定を 1 件として持つ
                    if not done: P(mid, fy, "not_done" if fy <= 2026 else "planned", "当初計画 " + sym, A.get(sym, "OVERHAUL"), fn, rc, key=(mid, fy, sym))
            # 印の抜け：実績追記版で R2〜R4 に行った整備・更新の印が無い
            if fn == OLD_PLAN:
                for fy in (2020, 2021, 2022):
                    ws_ = [i for i in done_by.get((mid, fy), []) if wk[i]["kind"] != "INSPECT"]
                    if ws_ and not str(f["marks"].get(str(fy), "")).endswith("実"):
                        errata.append({"type": "omitMark", "file": fn, "sheet": rc["sheet"], "r": rc["r"], "year": fy, "works": ws_})
    world = {
        "schema": "moyai.sample-world/0",
        "facility": {"id": "F04", "name": "【架空】しおさい町ごみ処理場"},
        "profile": dict(w1["profile"]),
        "maker": {"today": "2026-11-13", "seed": w1.get("seed")},
        "assets": assets, "works": works, "links": links, "errata": errata,
        "files": [{"file": f} for f in w1["files"]],
    }
    with open(os.path.join(ANS, "world.json"), "w", encoding="utf-8") as fo:
        json.dump(world, fo, ensure_ascii=False, indent=1)
    extra = {k: w1[k] for k in ("unrecorded", "plan_gaps", "contracts", "equipment") if k in w1}
    extra["links_v1"] = w1["links"]
    extra["notes"] = ("marks/plans の当て方：コンサル策定版は marks＝plans＝書いてある記号。実績追記版は同じセルに予定と実績が入る様式なので marks＝セルに書いてある字（◎済・●実 など）、"
                      "plans＝コンサルの当初計画の記号（生成器の CPLAN）。t1 が足した 83-2・92-2 行は当初計画が無いので plans は空。R5〜R15 版は 計 の列＝plans、実 の列＝marks。"
                      "PDF の記録は r＝seq＋1（seq も残す）。合計行の manDays は _check に移した。工事別日報（g2）は 1 行＝1 日として days＝1。"
                      "星取表の予定（planned）・繰延べ（postponed）・印だけで実施しなかったもの（not_done）は P で始まる作業として足した。")
    with open(os.path.join(ANS, "world.extra.json"), "w", encoding="utf-8") as fo:
        json.dump(extra, fo, ensure_ascii=False, indent=1)
    print("assets %d works %d links %d errata %d files %d" % (len(assets), len(works), len(links), len(errata), len(world["files"])))

if __name__ == "__main__":
    main()
