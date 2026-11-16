# -*- coding: utf-8 -*-
"""answer/ を物差しの形（moyai.sample-world/0）にそろえる変換器。files/ は読むだけで書かない。
元にするのは旧正解の控え（f11/answer_v1_backup/）と生成器の中の正解データ（plan_data・fullworld の測定や停止日程は使わない）。
python3 to_v2.py"""
import os, json, glob, re, copy, datetime as dt, hashlib

GEN = os.path.dirname(os.path.abspath(__file__))
FULL = os.path.dirname(GEN)
V1 = os.path.join(FULL, "gen", "_answer_v1")
ANS = os.path.join(FULL, "answer")
FILES = os.path.join(FULL, "files")
TODAY = "2026-11-13"

import sys
sys.path.insert(0, GEN)
import world as W
import plan_data as P

ATTR = {"○": "INSPECT", "◎": "OVERHAUL", "△": "OVERHAUL", "●": "REPLACE", "※": "REPLACE", "☆": "INSPECT"}
WORD = {"○": "点検", "◎": "分解整備", "△": "補修", "●": "取替・更新", "※": "部品購入", "☆": "法定検査"}

def fy_of(d):
    y, m = int(d[:4]), int(d[5:7])
    return y if m >= 4 else y - 1

def hash_files():
    h = {}
    for f in sorted(os.listdir(FILES)):
        h[f] = hashlib.sha256(open(os.path.join(FILES, f), "rb").read()).hexdigest()
    return h

before = hash_files()
v1w = json.load(open(os.path.join(V1, "world.json"), encoding="utf-8"))
v1man = json.load(open(os.path.join(V1, "manifest.json"), encoding="utf-8"))
MAN = {m["file"]: m for m in v1man}

# ---------------- works ----------------
works = []
for w in v1w["works"]:
    w2 = dict(id=w["id"], asset=w["equip"] or None, fy=fy_of(w["start"]), status="done" if w["end"] <= TODAY else "planned",
              task=w["task"], attr=w["attr"], start=w["start"], end=w["end"], manDays=w["manDays"], days=None,
              _vendor=w["vendor"], _contract=w["contract"])
    if w["start"] <= TODAY < w["end"]:
        w2["status"] = "done"   # 記録の時点で途中（日数は済んだ分）
    works.append(w2)
# 日数：日報に出た日の数（e1 の契約の作業）
daycount = {}
nid = [len(works)]
def new_work(asset, fy, status, task, attr, start=None, end=None, contract=None, md=None, days=None):
    nid[0] += 1
    w = dict(id="W%04d" % nid[0], asset=asset, fy=fy, status=status, task=task, attr=attr, start=start, end=end, manDays=md, days=days, _contract=contract)
    works.append(w)
    return w

# 部品購入（物品購入の契約）と、工程表の予定だけの作業（R8 2号炉の 11/14 以降）を作業に足す
PURCH = [("1B-03", 2022, "1号炉燃焼火格子片購入（90 枚）", "2022-10-14", "2023-03-22", "P2022a"),
         ("2F-07", 2024, "誘引通風機羽根車購入（1 個）", "2024-05-20", "2025-03-14", "P2024a"),
         ("0A-07", 2020, "ごみクレーン予備機巻上ワイヤ購入（2 本）", "2020-06-01", "2020-10-27", "P"),
         ("1C-03", 2019, "スートブロワランス購入（1 本）", "2019-10-01", "2020-03-10", "P")]
for a, fy, t, s, e, c in PURCH:
    new_work(a, fy, "done", t, "REPLACE", s, e, c)
# R8 2号炉 定修の残り（記録の時点より後）
for t, a, s, e in [("耐火物 点検・部分補修（残り）", "2B-01", "2026-11-16", "2026-11-16"), ("ボイラ性能検査（受検）", "2C-01", "2026-11-17", "2026-11-17")]:
    new_work(a, 2026, "planned", t, "INSPECT", s, e, "T2026")

def works_for(asset, fy, attr=None, status=None):
    out = [w for w in works if w["asset"] == asset and w["fy"] == fy and (attr is None or w["attr"] == attr) and (status is None or w["status"] == status)]
    return out

# ---------------- answer/files ----------------
os.makedirs(os.path.join(ANS, "files"), exist_ok=True)
for f in glob.glob(os.path.join(ANS, "files", "*.json")):
    os.remove(f)
links = []
errata = []
PLAN_TAKE = []   # 星取表の当て方の記録
LEGAL_ROW = {}

def mk_link(file, sheet, r, work, how="direct", year=None, day=None):
    l = dict(file=file, sheet=sheet, r=r, work=work, how=how)
    if year is not None: l["year"] = year
    if day is not None: l["day"] = day
    links.append(l)

def plan_marks_of(fname, e_no, legal_key=None):
    """その版の予定の記号（年度→記号）"""
    if legal_key:
        row = [x for x in P.LEGAL if x[0] == legal_key][0]
        y1, ya, y2 = row[4], row[5], row[6]
        if "l1様式" in fname:
            return {y: "☆" for y in (y1 or [])}
        d = {y: "☆" for y in (y1 or []) if y <= 2022}
        d.update({y: "☆" for y in y2 if y >= 2023})
        return d
    e = W.EQ[e_no]
    if "l1様式" in fname:
        return P.l1_plan(e)
    d = {y: m for y, m in P.l1_plan(e).items() if y <= 2022}
    d.update({y: m for y, m in P.l2_plan(e).items() if y >= 2023})
    if e_no == "2F-07":
        d[2024] = "○"
    return d

LEGAL_KEYS = {x[1]: x[0] for x in P.LEGAL}
legal_works = {}

def legal_work(key, name, y, status):
    k = (key, y)
    if k not in legal_works:
        legal_works[k] = new_work(None, y, status, name, "INSPECT", None, None, "法定")
    return legal_works[k]

def plan_target(asset, y, mark, rec_file):
    """印に当たる作業（無ければ予定・未実施・繰越の作業を作る）"""
    attr = ATTR.get(mark, "INSPECT")
    done = [w for w in works_for(asset, y, attr) if w["status"] == "done"]
    if not done and mark == "○":
        done = [w for w in works_for(asset, y) if w["status"] == "done"]
    if done:
        return done
    if y >= 2026 and (y > 2026 or True):
        st = "planned"
    later = [w for w in works if w["asset"] == asset and w["fy"] > y and w["attr"] == attr and w["status"] == "done" and attr != "INSPECT"]
    if y < 2026:
        st = "postponed" if later else "not_done"
    key = (asset, y, attr, st)
    ex = [w for w in works if w["asset"] == asset and w["fy"] == y and w["attr"] == attr and w["status"] == st]
    if ex:
        return ex[:1]
    return [new_work(asset, y, st, "%s（計画）" % WORD.get(mark, "点検"), attr)]

for p in sorted(glob.glob(os.path.join(V1, "files", "*.json"))):
    d = json.load(open(p, encoding="utf-8"))
    fname = d["file"]
    ext = fname.rsplit(".", 1)[1]
    recs = []
    for rec in d["records"]:
        r2 = copy.deepcopy(rec)
        f = r2["fields"]
        chk = {}
        for k in ("contract", "task", "原文", "_equip", "sheet") if not d["plan"] else ():
            if k in f:
                chk[k] = f.pop(k)
        if ext in ("pdf", "csv"):
            r2.pop("sheet", None)
            r2["sheet"] = None
        if ext == "pdf":
            r2["r"] = rec["seq"] + 1
        if chk:
            r2["_check"] = chk
        recs.append(r2)
    if d["plan"]:
        for r2 in recs:
            f = r2["fields"]
            path = f.get("path") or []
            if f.get("work") == "費用":
                r2["kind"] = "subtotal" if r2["kind"] == "item" else r2["kind"]   # 費用の行は機器の行ではない
                f["marks"], f["plans"] = {}, {}
                continue
            if r2["kind"] == "parent":
                f.setdefault("marks", {}); f.setdefault("plans", {})
                continue
            code = f.get("code") or None
            lk = LEGAL_KEYS.get(path[-1]) if path and path[0].startswith("法定") else None
            written = {int(y): v for y, v in f.get("marks", {}).items()}
            if code and len(path) == 4:   # 部品手当ての子行
                pm = {2022: "※"} if code == "1B-03" else {2024: "※"}
                if "l1様式" in fname:
                    pm = {}
            elif code:
                pm = plan_marks_of(fname, code)
            elif lk:
                pm = plan_marks_of(fname, None, lk)
            else:
                pm = {}
            # 実績の印＝「済」の付いた欄、予定の印＝その版の予定（l1 は全部が予定）
            marks = {str(y): v for y, v in written.items() if len(str(v)) <= 8}
            plans = {str(y): v for y, v in pm.items()}
            f["marks"], f["plans"] = marks, plans
            # 結び
            ys = sorted(set(int(y) for y in marks) | set(int(y) for y in plans))
            for y in ys:
                mk = str(marks.get(str(y)) or plans.get(str(y)) or "")
                base = mk.replace("済", "")
                if not base:
                    continue
                if lk:
                    st = "done" if ("済" in mk or ("l1様式" in fname and y <= 2025)) else ("planned" if y >= 2026 else "not_done")
                    yrs_done = [x for x in P.LEGAL if x[0] == lk][0][5]
                    if y in yrs_done or (y == 2026 and lk in {"boiler1", "boilerm", "crane_y", "pv1", "pv2", "pvs", "elec", "eg", "gas", "dxn", "dxnw", "scale", "fire"}):
                        st = "done"
                    elif y < 2026:
                        st = "not_done"
                    w = legal_work(lk, path[-1], y, st)
                    mk_link(fname, r2["sheet"], r2["r"], w["id"], "direct", year=y)
                    continue
                asset = code
                for w in plan_target(asset, y, base[0], fname):
                    mk_link(fname, r2["sheet"], r2["r"], w["id"], "direct", year=y)
    out = dict(file=fname, plan=d["plan"], records=recs)
    if d.get("images"):
        out["images"] = d["images"]
    json.dump(out, open(os.path.join(ANS, "files", os.path.basename(p)), "w", encoding="utf-8"), ensure_ascii=False, indent=1, default=str)

# 印の抜け（omitMark）：l1 の計量機定期検査は R2・R4・R6・R8 に印が無い（実績はその年）
l1 = [m["file"] for m in v1man if "l1様式" in m["file"] and "中長期" in m["file"]][0]
d1 = json.load(open(os.path.join(ANS, "files", os.path.splitext(l1)[0] + ".json"), encoding="utf-8"))
for rec in d1["records"]:
    pth = rec["fields"].get("path") or []
    if pth and pth[-1].startswith("計量機"):
        for y in (2020, 2022, 2024, 2026):
            w = legal_work("scale", pth[-1], y, "done")
            errata.append(dict(type="omitMark", file=l1, sheet=rec["sheet"], r=rec["r"], year=y, works=[w["id"]]))
            mk_link(l1, rec["sheet"], rec["r"], w["id"], "direct", year=y)

# ---------------- ほかの資料の結び ----------------
def wl(contract):
    return [w for w in works if w.get("_contract") == contract]

for p in sorted(glob.glob(os.path.join(ANS, "files", "*.json"))):
    d = json.load(open(p, encoding="utf-8"))
    if d["plan"]:
        continue
    fname = d["file"]
    kind = MAN[fname]["資料"]
    cid = MAN[fname]["契約"]
    for rec in d["records"]:
        f = rec["fields"]; chk = rec.get("_check", {})
        r = rec["r"]; sh = rec.get("sheet")
        if kind == "工程表" and rec["kind"] == "item":
            c = chk.get("contract") or cid
            ws = [w for w in wl(c) if w["task"] == f.get("作業")]
            for w in ws:
                mk_link(fname, sh, r, w["id"], "direct")
            if f.get("実績"):
                f["days"] = len(f["実績"])
        elif kind == "作業日報" and rec["kind"] == "item" and f.get("日付"):
            day = f["日付"]
            todays = [w for w in wl(cid) if w["start"] and w["start"] <= day <= w["end"]]
            if f.get("欄") == "作業内容":
                txt = chk.get("原文") or f.get("作業内容") or ""
                hit = [w for w in todays if w["asset"] and (W.EQ[w["asset"]].name in txt or W.EQ[w["asset"]].name[:4] in txt)]
                if hit:
                    for w in hit[:4]:
                        mk_link(fname, sh, r, w["id"], "direct", day=day)
                elif todays:
                    for w in todays[:12]:
                        mk_link(fname, sh, r, w["id"], "rollup", day=day)
            else:
                for w in todays[:12]:
                    mk_link(fname, sh, r, w["id"], "rollup", day=day)
        elif kind == "工事報告書" and rec["kind"] == "item":
            if f.get("区分") == "サマリー":
                c = cid
                ws = [w for w in wl(c) if w["task"] == f.get("作業")] or [w for w in wl(c) if w["asset"] in (f.get("対象機器") if isinstance(f.get("対象機器"), list) else [])]
                for w in ws:
                    mk_link(fname, None, r, w["id"], "direct")
            elif f.get("区分") == "工程実績" and f.get("日付"):
                day = f["日付"]
                for w in [w for w in wl(cid) if w["start"] and w["start"] <= day <= w["end"]][:12]:
                    mk_link(fname, None, r, w["id"], "rollup", day=day)
            elif f.get("区分") == "所見" and chk.get("_equip"):
                for w in works_for(chk["_equip"], 2019, status="done")[:3]:
                    mk_link(fname, None, r, w["id"], "direct")
        elif kind == "発注台帳" and fname.endswith(".csv") and rec["kind"] == "item":
            c = chk.get("contract") or ""
            if c.startswith("T"):
                for w in wl(c):
                    mk_link(fname, None, r, w["id"], "annual")
            elif c.startswith("R"):
                for w in wl(c):
                    mk_link(fname, None, r, w["id"], "bundle")
            elif c.startswith("P"):
                for w in wl(c):
                    mk_link(fname, None, r, w["id"], "direct")
    json.dump(d, open(p, "w", encoding="utf-8"), ensure_ascii=False, indent=1, default=str)

# 日数（作業の行が日報に出た日の数）＝ start〜end の稼働日で数える
for w in works:
    if w.get("start") and w.get("end") and w["days"] is None:
        a = dt.date.fromisoformat(w["start"]); b = dt.date.fromisoformat(w["end"])
        n = sum(1 for i in range((b - a).days + 1) if (a + dt.timedelta(i)).weekday() < 5)
        w["days"] = n or 1

# ---------------- world.json ----------------
eq = {e["no"]: e for e in v1w["equipment"]}
assets = []
ledger = [m["file"] for m in v1man if m["資料"] == "機器台帳"]
plans = [m["file"] for m in v1man if m["資料"] == "中長期整備計画"]
for e in v1w["equipment"]:
    names = {}
    for lf in ledger:
        names[lf] = "%s > %s" % (e["category"], e["name"])
    for pf in plans:
        names[pf] = "%s > %s > %s" % (e["category"], e["name"], e["unit"])
    names["作業日報（手書き）"] = e["aliases"].get("日報（手書き）")
    names["工事報告書"] = e["aliases"].get("報告書")
    if e["aliases"].get("工程表のまとまり"):
        names["工程表"] = e["aliases"]["工程表のまとまり"]
    assets.append(dict(id=e["no"], name=e["name"], tag=e["oldno"], system=e["category"], furnace=e["unit"], installed=e["install_year"],
                       class_=None, part=None, names=names, units=e["units"], remark=e["remark"]))
for a in assets:
    a["class"] = a.pop("class_")
wout = []
for w in works:
    x = {k: v for k, v in w.items() if not k.startswith("_")}
    x["_contract"] = w.get("_contract"); x["_vendor"] = w.get("_vendor")
    wout.append(x)
world = {
    "schema": "moyai.sample-world/0",
    "facility": {"id": "F11", "name": W.FAC},
    "profile": dict(v1w["profile"]),
    "maker": {"today": TODAY, "seed": 1111, "tool": "f11/full/gen（run_full.py → to_v2.py）"},
    "assets": assets,
    "works": wout,
    "links": links,
    "errata": errata,
    "files": [{"file": m["file"], "資料": m["資料"], "年度": m["年度"], "契約": m["契約"]} for m in v1man],
}
json.dump(world, open(os.path.join(ANS, "world.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1, default=str)
extra = {"contracts": v1w["contracts"], "unrecorded": v1w["unrecorded"], "plan_gaps": v1w["plan_gaps"],
         "equipment_v1": v1w["equipment"], "links_v1_count": len(v1w["links"]),
         "plan_how": "星取表：marks＝その行に書いてある記号（l2 版の「○済」などは実績、記号だけの欄は予定のまま）。plans＝その版の予定（l1 版は l1 の当初計画、l2 版は R2〜R4 が l1・R5〜R9 が R5.4 改訂の計画）。費用の行は kind を subtotal にした。",
         "errata_other": "第一種圧力容器の行の欠け（行そのものが無い）・R3 決算計の 110 千円・前回整備列の未更新は plan_gaps に残す"}
json.dump(extra, open(os.path.join(ANS, "world.extra.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1, default=str)
if os.path.exists(os.path.join(ANS, "manifest.json")):
    pass
after = hash_files()
assert before == after, "files/ が変わった"
from collections import Counter
print("works", len(works), Counter(w["status"] for w in works), "links", len(links), Counter(l["how"] for l in links), "errata", len(errata))
