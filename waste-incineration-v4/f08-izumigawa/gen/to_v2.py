# -*- coding: utf-8 -*-
"""answer/ を物差しの形（moyai.sample-world/0）に書き直す変換器。files/ は読みも書きもしない。
入力：f08/answer_v1_backup/（旧正解）と生成器の中の正解データ（world・works・model。メモリ上だけ）。
出力：full/answer/world.json・full/answer/world.extra.json・full/answer/files/*.json
    python3 -I to_v2.py"""
import datetime as dt
import glob
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import world  # noqa: E402
from model import is_off  # noqa: E402
from works import CONTRACTS, EQ, actual_works  # noqa: E402

GEN = os.path.dirname(os.path.abspath(__file__))
FULL = os.path.dirname(GEN)
OLD = os.path.join(FULL, "gen", "_answer_v1")
ANS = os.path.join(FULL, "answer")
FILES = os.path.join(FULL, "files")

ATTR = {"○": "INSPECT", "◎": "OVERHAUL", "●": "REPLACE", "△": "OVERHAUL", "▲": "OVERHAUL", "◇": "REPLACE"}
LINE = {0: "共通", 1: "1", 2: "2"}


def d(s):
    return dt.date.fromisoformat(s) if s else None


def workdays(a, b):
    if not a or not b:
        return None
    a, b = d(a), d(b)
    n, x = 0, a
    while x <= b:
        if not is_off(x):
            n += 1
        x += dt.timedelta(days=1)
    return n or (b - a).days + 1


def half(day):
    try:
        return 0.5 if d(day).weekday() == 5 else 1.0
    except Exception:
        return 1.0


def assets():
    out = []
    first = {}
    for e in EQ:
        first.setdefault((e["line"], e["base"]), e["id"])
    for e in EQ:
        par = world.PARENT.get(e["base"])
        part = first.get((e["line"], par)) if par else None
        m1 = e["name"]
        for k, v in {"誘引通風機": "誘引送風機", "散気装置": "散気管", "バグフィルタ": "バグフィルター",
                     "定量給じん機": "給塵スクリュー", "廃熱ボイラ": "ボイラ", "流動床炉本体": "焼却炉本体"}.items():
            m1 = m1.replace(k, v)
        S = world.SETSUBI[e["setsubi"]]
        out.append({"id": e["id"], "name": e["name"], "tag": e["tag"] or None, "system": f"{S}/{e['keitou']}",
                    "furnace": LINE[e["line"]], "installed": 1990 if e["year"] == 1989 else e["year"] + 1,
                    "class": e["imp"], "part": part,
                    "names": {"中長期整備計画": f"{S} > {'共通' if e['line'] == 0 else str(e['line']) + '号炉'} > {e['keitou']} > {e['name']}",
                              "機器台帳": f"{S} > {e['keitou']} > {e['base']}",
                              "工事報告書（メーカー）": m1}})
    return out


def how_for(fn, n_works, contract=None):
    if "_発注台帳_" in fn:
        if contract and contract.startswith("T"):
            return "annual"
        if contract == "P2008":
            return "upgrade"
        return "direct" if n_works == 1 else "bundle"
    return "direct" if n_works == 1 else "bundle"


def main():
    W = actual_works()
    wby = {w["wid"]: w for w in W}
    works = []
    for w in W:
        st, en = (w.get("start").isoformat() if w.get("start") else None), (w.get("end").isoformat() if w.get("end") else None)
        works.append({"id": w["wid"], "asset": w["eq"], "fy": w["fy"], "status": "done", "task": w["label"], "attr": w["kind"],
                      "start": st, "end": en, "manDays": w.get("manDays"), "days": workdays(st, en),
                      "vendor": w["vendor"] or CONTRACTS.get(w["contract"], {}).get("vendor"), "contract": w["contract"]})
    actual_idx = {}
    for w in W:
        actual_idx.setdefault((w["eq"], w["fy"]), []).append(w["wid"])
    planned = {}          # 計画の印で、実際に行わなかったもの → status not_done（共通の ID で版をまたぐ）

    def plan_work(eq, fy, sym):
        if (eq, fy) in actual_idx:
            return actual_idx[(eq, fy)][0]
        key = (eq, fy)
        if key not in planned:
            pid = f"P{len(planned) + 1:04d}"
            planned[key] = pid
            works.append({"id": pid, "asset": eq, "fy": fy, "status": "not_done", "task": f"計画の印 {sym}", "attr": ATTR.get(sym, "INSPECT"),
                          "start": None, "end": None, "manDays": None, "days": None})
        return planned[key]

    links, errata, files_out = [], [], []
    old_world = json.load(open(os.path.join(OLD, "world.json"), encoding="utf-8"))
    old_links = {}
    for l in old_world["links"]:
        old_links[(l["file"], l.get("sheet"), l.get("r", l.get("seq")))] = l["works"]
    contract_of = {w["wid"]: w["contract"] for w in W}
    os.makedirs(os.path.join(ANS, "files"), exist_ok=True)
    for f in glob.glob(os.path.join(ANS, "files", "*.json")):
        os.remove(f)
    for path in sorted(glob.glob(os.path.join(OLD, "files", "*.json"))):
        old = json.load(open(path, encoding="utf-8"))
        fn = old["file"]
        is_pdf, is_csv = fn.endswith(".pdf"), fn.endswith(".csv")
        recs = []
        for rec in old["records"]:
            f = dict(rec["fields"])
            seq = rec.get("seq")
            r = rec["r"] if "r" in rec else seq + 1          # PDF は出てくる順の番号（1 から）
            sheet = None if (is_pdf or is_csv) else rec.get("sheet")
            kind = rec["kind"]
            new = {"sheet": sheet, "r": r, "kind": kind, "fields": f}
            if seq is not None:
                new["seq"] = seq
            key = (fn, rec.get("sheet"), rec.get("r", seq))
            wids = old_links.get(key, [])
            if old["plan"]:
                if kind == "item" and "cost" in f and "marks" not in f:
                    new["kind"] = "cost"                    # 費用シートの機器行（星取表と同じ機器。印は無い）
                    recs.append(new)
                    continue
                if kind in ("item", "parent"):
                    mk = f.pop("marks", {}) or {}
                    act = f.pop("marksActual", None)
                    if act is None:                          # 旧様式：計画だけの行
                        f["marks"], f["plans"] = dict(mk), dict(mk)
                    else:                                    # 新様式：同じ行の別の列（実績・計画）
                        f["marks"], f["plans"] = dict(act), dict(mk)
                    f["marks"] = {y: s for y, s in f["marks"].items() if len(str(s)) <= 8}
                    f["plans"] = {y: s for y, s in f["plans"].items() if len(str(s)) <= 8}
                    code = f.get("code")
                    for y in sorted(set(f["marks"]) | set(f["plans"])):
                        if code:
                            wid = plan_work(code, int(y), f["plans"].get(y) or f["marks"].get(y))
                            links.append({"file": fn, "sheet": sheet, "r": r, "work": wid, "year": int(y), "how": "direct"})
                recs.append(new)
                continue
            # ---- 計画以外：manDays（人日、1 人日＝8 時間）と days
            if "manHours" in f:                              # 地元業者の週報（人×時間）
                f["manDays"] = round(f["manHours"] / 8, 3)
                f["days"] = 1 if kind == "item" else None
                if f["days"] is None:
                    f.pop("days")
            elif "crewText" in f:                            # メーカー日報 H22–（作業行の職種人数）
                n = sum(int(x) for x in re.findall(r"\d+", f["crewText"]))
                f["manDays"] = n * half(f.get("date"))
                f["manDaysText"] = f["crewText"]
                f["days"] = 1
            elif "people" in f and "date" in f:              # メーカー日報 H17–H21・報告書の経過（人数）
                f["manDays"] = f["people"] * half(f["date"])
                f["manDaysText"] = f"{f['people']}名"
                f["days"] = 1
            elif "actualStart" in f:                         # メーカー工程表
                f["days"] = workdays(f.get("actualStart"), f.get("actualEnd"))
            elif "manDays" in f and "工事名" in f:          # 地元業者の報告書 1 枚目
                pass
            if "_発注台帳_" not in fn and "工事名" in f:
                hrs = re.search(r"（([\d.]+) h）", f.get("現場作業", ""))
                if hrs:
                    f["manDaysText"] = f.get("manDaysText")
                    f["manDays"] = round(float(hrs.group(1)) / 8, 3)
            if "人員" in f and "月日" in f and "作業内容" in f:   # g1 報告書の続紙（作業経過の表）
                f["manDaysText"] = f"{f['人員']}名"
                f["manDays"] = float(f["人員"])
                f["days"] = 1
            f = {k: v for k, v in f.items() if v is not None}
            new["fields"] = f
            day = f.get("date")
            for wid in wids:
                if wid not in wby:
                    continue
                L = {"file": fn, "sheet": sheet, "r": r, "work": wid,
                     "how": "rollup" if ("区域" in f and "対象台数" in f) else how_for(fn, len(wids), contract_of.get(wid))}
                if day:
                    L["day"] = day
                links.append(L)
            recs.append(new)
        with open(os.path.join(ANS, "files", os.path.basename(path)), "w", encoding="utf-8") as fo:
            json.dump({"file": fn, "plan": old["plan"], "records": recs}, fo, ensure_ascii=False, indent=1)
        files_out.append({"file": fn})
        # わざと入れた印の抜け（R4 改訂のクレーン性能検査 R6）
        if old["plan"] and "令和4年度改訂" in fn:
            for rec in recs:
                p = rec["fields"].get("path") or []
                if len(p) > 1 and "クレーン性能検査" in p[-1]:
                    errata.append({"type": "omitMark", "file": fn, "sheet": rec["sheet"], "r": rec["r"], "year": 2024,
                                   "works": [w["wid"] for w in W if w["fy"] == 2024 and w["eq"] in ("0A-015", "0A-016")]})
    profile = old_world["profile"]
    out = {"schema": "moyai.sample-world/0",
           "facility": {"id": "F08", "name": "【架空】いずみ川清掃センター"},
           "profile": {"furnace": profile["furnace"], "capacity": profile["capacity"], "boiler": profile["boiler"],
                       "period": profile["period"], "stopped": profile["stopped"]},
           "maker": {"today": "2026-11-13"},
           "assets": assets(), "works": works, "links": links, "errata": errata,
           "files": sorted(files_out, key=lambda x: x["file"])}
    with open(os.path.join(ANS, "world.json"), "w", encoding="utf-8") as fo:
        json.dump(out, fo, ensure_ascii=False, indent=1)
    extra = {k: v for k, v in old_world.items() if k not in ("works", "links")}
    extra["note"] = ("v1 の world.json のうち新しい形に無い欄。equipment は v1 の機器（各資料での名前・図書番号・2016 区分）、"
                     "removedEquipment は撤去した竣工時の機器（管理対象外）、planGaps は errata に入れなかった計画の欠け、"
                     "worksWithoutDocs は資料に工程表・日報・報告書の無い作業")
    extra["planGapsNotInErrata"] = [g for g in old_world["planGaps"] if "クレーン" not in g["what"]]
    with open(os.path.join(ANS, "world.extra.json"), "w", encoding="utf-8") as fo:
        json.dump(extra, fo, ensure_ascii=False, indent=1)
    print("files", len(files_out), "works", len(works), "(not_done", len(planned), ")", "links", len(links), "errata", len(errata))


if __name__ == "__main__":
    main()
