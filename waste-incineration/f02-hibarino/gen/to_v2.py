# -*- coding: utf-8 -*-
"""answer/ を物差しの形（moyai.sample-world/0）にそろえる変換器。files/ は読まず・書かない。
入力：answer_v1（今の answer/ の写し。無ければ answer/ を読む）＋生成器の正解データ（world.py・contracts・gen_contract）。
出力：answer/world.json、answer/world.extra.json、answer/files/*.json
python3 to_v2.py"""
import datetime as dt
import glob
import json
import os
import re
import shutil
import world as W

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(BASE, "gen", "_answer_v1")
if not os.path.isdir(SRC):
    SRC = os.path.join(BASE, "answer")
OUT = os.path.join(BASE, "answer")


def load(p):
    return json.load(open(p, encoding="utf-8"))


w1 = load(os.path.join(SRC, "world.json"))
A = {}
for p in glob.glob(os.path.join(SRC, "files", "*.json")):
    d = load(p)
    A[d["file"]] = d

EQ = {e.no: e for e in W.EQ}
FUR = {0: "共通", 1: "1", 2: "2"}


def workdays(s, e):
    if not s or not e:
        return None
    a, b = dt.date.fromisoformat(s), dt.date.fromisoformat(e)
    n = 0
    while a <= b:
        if a.weekday() != 6:
            n += 1
        a += dt.timedelta(days=1)
    return n


def kind_of(fn):
    for k in ("中長期整備計画", "機器台帳", "工程表", "作業日報", "工事報告書", "発注台帳"):
        if f"_{k}_" in fn:
            return k
    return None


# ---------------------------------------------------------------- works
works = {}
for o in w1["works"]:
    e = EQ[o["equipment"]]
    verb = {"INSPECT": "点検", "OVERHAUL": "補修" if o.get("repair") else "分解整備", "REPLACE": "更新", "UPGRADE": "改良"}[o["attr"]]
    if o["id"].endswith("-CITY"):
        verb = "補修（市追加工事）"
    task = f"{e.full} {verb}"
    if e.strat == "bf" and o["attr"] == "REPLACE":
        task = f"{e.full} ろ布全数交換"
    works[o["id"]] = {"id": o["id"], "asset": o["equipment"], "fy": o["fy"], "status": "done", "task": task, "attr": o["attr"],
                      "start": o.get("start"), "end": o.get("end"), "manDays": o.get("manDays"), "days": workdays(o.get("start"), o.get("end")),
                      "_vendor": o.get("vendor"), "_contract": o.get("contract"), "_month": o.get("month")}
# 市工事の工数：RT 夜間（R5 10/20 検査員 3 名 17:00〜20:30）を足す
for wid, o in works.items():
    if wid.endswith("-CITY") and o["fy"] == 2023:
        o["manDays"] = round(o["manDays"] + 3 * 3.5 / 8, 2)

# 法定点検の作業（印の抜けの結び先）
LEGAL_W = {
    "L2024-計量機定期検査": {"id": "L2024-計量機定期検査", "asset": "HB01-001", "fy": 2024, "status": "done", "task": "計量機 定期検査（入口・出口）",
                          "attr": "INSPECT", "start": "2024-07-10", "end": "2024-07-10", "manDays": None, "days": 1},
    "L2029-精密機能検査": {"id": "L2029-精密機能検査", "asset": None, "fy": 2029, "status": "planned", "task": "精密機能検査（施設全体）",
                        "attr": "INSPECT", "start": None, "end": None, "manDays": None, "days": None},
}
works.update(LEGAL_W)

SYM_ATTR = {"○": "INSPECT", "◎": "OVERHAUL", "△": "OVERHAUL", "●": "REPLACE", "点": "INSPECT", "整": "OVERHAUL", "補": "OVERHAUL",
            "換": "REPLACE", "改": "UPGRADE"}


def plan_work_id(asset, y, sym):
    """計画の印の結び先。実績があればそれ、無ければ予定（planned）か未実施（not_done／postponed）の作業を起こす。"""
    wid = f"W{y}-{asset}"
    if wid in works:
        return wid
    cw = f"W{y}-{asset}-CITY"
    if cw in works:
        return cw
    pid = f"P{y}-{asset}"
    if pid not in works:
        e = EQ[asset]
        if y > 2026 or (y == 2026 and not W.actual_work(e, 2026)):
            st = "planned"
        else:
            st = "postponed" if f"W{y + 1}-{asset}" in works else "not_done"
        attr = SYM_ATTR.get(sym[0], "INSPECT")
        works[pid] = {"id": pid, "asset": asset, "fy": y, "status": st, "task": f"{e.full} {sym}（計画）", "attr": attr,
                      "start": None, "end": None, "manDays": None, "days": None}
    return pid


# ---------------------------------------------------------------- files/*.json
links = []
os.makedirs(os.path.join(OUT, "files"), exist_ok=True)
for p in glob.glob(os.path.join(OUT, "files", "*.json")):
    os.remove(p)
old_links = {}
for l in w1["links"]:
    old_links.setdefault((l["file"], l.get("sheet"), l.get("r"), l.get("seq")), []).append(l["work"])

for fn, d in A.items():
    kind = kind_of(fn)
    is_pdf = fn.endswith(".pdf")
    is_csv = fn.endswith(".csv")
    recs = []
    plan_rows = {}
    if d["plan"]:
        for rr in d["records"]:
            if rr["fields"].get("計実") == "計" and rr["kind"] == "item":
                plan_rows[tuple(rr["fields"]["path"])] = rr["fields"].get("marks") or {}
    for rr in d["records"]:
        f = dict(rr["fields"])
        nr = {"sheet": None if (is_pdf or is_csv) else rr.get("sheet"), "kind": rr["kind"], "fields": f}
        if "r" in rr:
            nr["r"] = rr["r"]
        if "seq" in rr:
            nr["seq"] = rr["seq"]
            nr["r"] = rr["seq"] + 1          # PDF：r ＝ seq＋1
        key = (fn, rr.get("sheet"), rr.get("r"), rr.get("seq"))
        wl = old_links.get(key, [])
        if d["plan"]:
            sheet = rr.get("sheet")
            if sheet == "年度別概算費用" or (f.get("path") or [""])[0] == "年度計画額":
                if nr["kind"] == "item":
                    nr["kind"] = "subtotal"          # 費用の集計行（機器の行ではない）
            hm = f.pop("hiddenMarks", None)
            if hm:
                nr["_hiddenMarks"] = hm            # 隠し列（H29 版の旧計画）。印として数えない
            if nr["kind"] in ("item", "parent"):
                marks = {k: v for k, v in (f.get("marks") or {}).items() if len(v) <= 8}
                f["marks"] = marks
                if f.get("計実") == "実":
                    f["plans"] = dict(plan_rows.get(tuple(f["path"]), {}))
                else:
                    f["plans"] = dict(marks)
                code = f.get("code")
                asset = code.rstrip("R") if code else None
                if code and asset in EQ and nr["kind"] == "item":
                    if f.get("計実") == "実":
                        for wid in wl:
                            links.append({"file": fn, "sheet": sheet, "r": rr["r"], "work": wid, "year": works[wid]["fy"], "how": "direct"})
                    else:
                        for y, sym in marks.items():
                            wid = plan_work_id(asset, int(y), sym)
                            links.append({"file": fn, "sheet": sheet, "r": rr["r"], "work": wid, "year": int(y), "how": "direct"})
                wl = []
        else:
            # 工数・日数
            if kind == "作業日報":
                if isinstance(f.get("days"), dict):          # 週報式：日ごとの欄
                    f["daily"] = f.pop("days")
                if nr["kind"] == "item" and "作業内容" in f or "作業項目" in f and "daily" in f:
                    if "daily" in f and "作業項目" in f:
                        f["days"] = len(f["daily"])
                        f["manDays"] = round(float(f.get("週工数", 0)) / 8, 2)
                        f["manDaysText"] = f"{f.get('週工数')}（人・時）"
                    elif "時間" in f:                         # 1日1枚：人員×時間
                        f["days"] = 1
                        f["manDays"] = round(f["人員"] * f["時間"] / 8, 2)
                        f["manDaysText"] = f"{f.get('工数')}（人・時）"
                    elif "午前" in f or "午後" in f:          # 2016 様式：午前・午後の○と人員
                        hrs = (4 if f.get("午前") else 0) + (4 if f.get("午後") else 0) + (1.5 if f.get("備考") == "残業 1.5h" else 0)
                        f["days"] = 1
                        f["manDays"] = round(f["人員"] * hrs / 8, 2)
                        f["manDaysText"] = f"{f['人員']}人（午前{f.get('午前') or '－'}・午後{f.get('午後') or '－'}）"
                    elif "職種別人員" in f:                   # 市工事のメーカー日報
                        tot = f["人員計"]
                        extra = 3 * 3.5 if f.get("年月日") == "2023-10-20" else 0
                        f["days"] = 1
                        f["manDays"] = round((tot * 8 + extra) / 8, 2)
                        f["manDaysText"] = f"{tot}人" + ("（RT 検査員 3名 20:30まで）" if extra else "")
                else:
                    f.pop("manDays", None)
                    f.pop("manDaysText", None)
                    if isinstance(f.get("days"), (dict, list)):
                        f["daily"] = f.pop("days")
            elif kind == "工程表" and nr["kind"] == "item":
                s = f.get("実績開始") or f.get("予定開始") or (f.get("開始") or "")[:10]
                e = f.get("実績終了") or f.get("予定終了") or (f.get("終了") or "")[:10]
                if s and e:
                    f["days"] = workdays(s, e)
            # 結び
            for wid in wl:
                how = "direct"
                if is_csv and (("定期整備工事" in str(f.get("品名", "")) and "追加" not in str(f.get("品名", ""))) ):
                    how = "bundle"
                if kind == "発注台帳" and fn.endswith(".xlsx"):
                    how = "bundle"
                if is_pdf and f.get("表") == "工事概要":
                    how = "bundle"
                if kind == "作業日報" and "daily" in f and "作業項目" in f:
                    for day in sorted(f["daily"]):
                        links.append({"file": fn, "sheet": rr.get("sheet"), "r": rr["r"], "work": wid, "day": day, "how": "direct"})
                    continue
                lk = {"file": fn, "sheet": None if (is_pdf or is_csv) else rr.get("sheet"), "work": wid, "how": how}
                if "r" in rr:
                    lk["r"] = rr["r"]
                else:
                    lk["seq"] = rr["seq"]
                    lk["r"] = rr["seq"] + 1
                day = f.get("日付") or f.get("年月日")
                if kind == "作業日報" and day:
                    lk["day"] = day
                links.append(lk)
        recs.append(nr)
    stem = os.path.splitext(fn)[0]
    with open(os.path.join(OUT, "files", stem + ".json"), "w", encoding="utf-8") as fp:
        json.dump({"file": fn, "plan": d["plan"], "records": recs}, fp, ensure_ascii=False, indent=1)

# ---------------------------------------------------------------- errata（印の抜け）
errata = []
for fn, d in A.items():
    if not d["plan"]:
        continue
    for rr in d["records"]:
        f = rr["fields"]
        if f.get("計実") == "実":
            continue
        nm = (f.get("path") or ["", ""])[-1]
        m = re.search(r"(\d{4})年度改訂", fn)
        v = int(m.group(1))
        if nm == "計量機 定期検査" and v in (2021, 2022):
            errata.append({"type": "omitMark", "file": fn, "sheet": rr["sheet"], "r": rr["r"], "year": 2024, "works": ["L2024-計量機定期検査"]})
        if nm == "精密機能検査" and v >= 2023:
            errata.append({"type": "omitMark", "file": fn, "sheet": rr["sheet"], "r": rr["r"], "year": 2029, "works": ["L2029-精密機能検査"]})

# ---------------------------------------------------------------- assets
assets = []
for q in w1["equipment"]:
    e = EQ[q["asset"]]
    cat = q["category"]
    nm_plan = e.name if e.f == 0 else f"{e.name}（{e.f}号）"
    names = {"機器台帳": q["ledger_name"], "中長期整備計画": f"{cat} > {nm_plan}",
             "工程表": e.name, "作業日報": (f"{e.f}炉 " if e.f else "") + (W.ALIAS.get(e.name, [e.name])[0]), "発注台帳": e.name}
    assets.append({"id": e.no, "name": q["ledger_name"], "tag": e.no, "system": cat, "furnace": FUR[e.f],
                   "installed": int(e.install[:4]), "class": e.imp, "part": None, "names": names,
                   "_renumbered": q.get("asset_renumbered"), "_aliases": q.get("aliases"), "_spec": e.spec, "_role": e.role or None,
                   "_policy2018": e.policy})

files_list = [{"file": f} for f in sorted(A)]
world = {"schema": "moyai.sample-world/0",
         "facility": {"id": "F02", "name": W.FAC},
         "profile": dict(w1["profile"], furnace="逆送式ストーカ 105t/24h×2（210t/日）全連続　炉数 2（1号・2号）＋共通"),
         "maker": {"today": "2026-11-13"},
         "assets": assets,
         "works": [{k: v for k, v in o.items()} for o in works.values()],
         "links": links,
         "errata": errata,
         "files": files_list}
extra = {k: w1[k] for k in w1 if k not in ("profile", "equipment", "works", "links", "files")}
extra["plan_gaps_other"] = [g for g in w1.get("plan_gaps", [])]
extra["notes"] = [
    "星取表の印：旧様式（計・実の 2 段）は 計 の行で marks＝plans＝計画の記号、実 の行で marks＝実績、plans＝同じ機器の 計 の行の記号。",
    "新様式（2023〜）は 1 行に 1号・2号の小列。同じ r に 1号・2号の 2 記録を置き、marks＝plans＝その炉の記号（実績段なし）。",
    "計 の印の結び先：同じ機器・年度の実績作業があればそれ。無ければ 2026 年度以降の予定は planned、過去で翌年度に実績があれば postponed、それ以外は not_done の作業を起こした（P〈年度〉-〈資産番号〉）。",
    "隠し列（H29 版 2027〜2030）の記号は _hiddenMarks に移し、印として数えない。費用シート・年度計画額の行は kind を subtotal にした。",
    "日報の manDays は人日（1 人日＝8 時間）。1日1枚は 人員×時間÷8、2016 様式は 午前・午後を各 4 時間（残業 1.5h）、週報式は 週工数÷8、市工事は 人員×8 時間（R5 10/20 の RT 夜間を加算）÷8。元の字は manDaysText。",
    "発注 CSV の定期整備工事一式（外注）と市の契約一覧、報告書の工事概要は how=bundle。",
]
with open(os.path.join(OUT, "world.json"), "w", encoding="utf-8") as fp:
    json.dump(world, fp, ensure_ascii=False, indent=1)
with open(os.path.join(OUT, "world.extra.json"), "w", encoding="utf-8") as fp:
    json.dump(extra, fp, ensure_ascii=False, indent=1)
st = {}
for o in works.values():
    st[o["status"]] = st.get(o["status"], 0) + 1
print("files", len(A), "assets", len(assets), "works", len(works), st, "links", len(links), "errata", len(errata))
