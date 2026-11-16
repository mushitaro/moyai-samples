# -*- coding: utf-8 -*-
"""answer/ を物差しの形（moyai.sample-world/0、common/answer-v2.md）にそろえる変換器。
files/ には触れない（読みもしない）。生成器は走らせない。
元の answer（v1）は初回に build/answer_v1/ へ写し、以後は必ずそこから変換する（何度走らせても同じ結果）。
  python3 gen/to_v2.py
"""
import os, sys, json, glob, shutil, re, datetime as dt
from collections import defaultdict, OrderedDict

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ANS = os.path.join(BASE, "answer")
V1 = os.path.join(BASE, "gen", "_answer_v1")
CUT = 2026          # この年度以後で実績の無い印は「予定」
SYM_ATTR = {"○": ("INSPECT", "点検・測定"), "◎": ("OVERHAUL", "分解整備・部分取替"), "◇": ("OVERHAUL", "補修"), "△": ("OVERHAUL", "補修"),
            "●": ("REPLACE", "更新"), "(●)": ("REPLACE", "更新（方針決定後に判断）"), "★": ("INSPECT", "法定点検")}
MATCH = {"INSPECT": ("INSPECT",), "OVERHAUL": ("OVERHAUL",), "REPLACE": ("REPLACE", "UPGRADE")}


def syms(s):
    out = []
    s = str(s)
    i = 0
    while i < len(s):
        if s.startswith("(●)", i):
            out.append("(●)")
            i += 3
        else:
            if s[i] in SYM_ATTR:
                out.append(s[i])
            i += 1
    return out


def ensure_v1():
    if not os.path.isdir(V1):
        os.makedirs(os.path.dirname(V1), exist_ok=True)
        shutil.copytree(ANS, V1)


def hours(tm):
    m = re.match(r"(\d+):(\d+)〜(\d+):(\d+)", str(tm))
    if not m:
        return None
    a = int(m.group(1)) * 60 + int(m.group(2))
    b = int(m.group(3)) * 60 + int(m.group(4))
    h = (b - a) / 60.0
    if a <= 12 * 60 and b >= 13 * 60:
        h -= 1.0        # 昼休憩 12:00〜13:00
    return h


def persons(s):
    s = str(s)
    m = re.search(r"外\s*(\d+)\s*名", s)
    if m:
        return 1 + int(m.group(1))
    m = re.search(r"(\d+)", s)
    return int(m.group(1)) if m else None


def main():
    ensure_v1()
    w1 = json.load(open(os.path.join(V1, "world.json"), encoding="utf-8"))
    works = OrderedDict((w["id"], dict(w)) for w in w1["works"])
    by_tag_fy = defaultdict(list)
    for w in works.values():
        by_tag_fy[(w["tag"], w["fy"])].append(w)
    files = sorted(os.path.basename(p)[:-5] for p in glob.glob(os.path.join(V1, "files", "*.json")))
    # v1 の結び（seq を r にする）
    links_v1 = defaultdict(list)
    for l in w1["links"]:
        key = (l["file"], l.get("sheet"), l.get("r", l.get("seq")))
        links_v1[key].append(l["work"])
    out_links = []
    pseudo = OrderedDict()
    day_of_work = defaultdict(set)

    def pseudo_work(tag, year, sym):
        attr, task = SYM_ATTR[sym]
        key = (tag, year, attr)
        if key not in pseudo:
            st = "not_done" if year < CUT else "planned"
            pseudo[key] = dict(id="P%05d" % (len(pseudo) + 1), tag=tag, fy=year, status=st, task=task, attr=attr, start=None, end=None,
                               manDays=None, days=None, src="plan")
        return pseudo[key]["id"]

    def actual_for(tag, year, sym):
        attr = SYM_ATTR[sym][0]
        lst = [w for w in by_tag_fy.get((tag, year), []) if w["kind"] in MATCH[attr]]
        return [w["id"] for w in lst]

    os.makedirs(os.path.join(ANS, "files"), exist_ok=True)
    for stem in files:
        a = json.load(open(os.path.join(V1, "files", stem + ".json"), encoding="utf-8"))
        fname = a["file"]
        is_csv = fname.endswith(".csv")
        is_pdf = fname.endswith(".pdf")
        scan = "手書き" in fname
        plan = bool(a.get("plan"))
        recs = []
        last_plan_marks = {}
        for rc in a["records"]:
            rc = dict(rc)
            f = dict(rc.get("fields", {}))
            if "r" not in rc:
                rc["r"] = rc.get("seq", 0)
            if is_pdf or is_csv:
                rc["sheet"] = None
            sheet = rc.get("sheet")
            # ---------- 星取表
            if plan and rc["kind"] in ("parent", "item"):
                f.setdefault("path", [])
                mk = {k: v for k, v in (f.get("marks") or {}).items() if len(str(v)) <= 8}
                code = f.get("code")
                if f.get("計実") == "実":
                    f["marks"] = mk
                    f["plans"] = last_plan_marks.get((sheet, code), {})
                else:
                    f["marks"] = mk
                    f["plans"] = dict(mk)
                    if code:
                        last_plan_marks[(sheet, code)] = dict(mk)
                # 印ごとの結び（年度つき）
                if code and rc["kind"] == "item":
                    for y, s in mk.items():
                        yi = int(y)
                        for sy in syms(s):
                            if f.get("計実") == "実":
                                ids = actual_for(code, yi, sy) or [w["id"] for w in by_tag_fy.get((code, yi), [])]
                            else:
                                ids = actual_for(code, yi, sy) or [pseudo_work(code, yi, sy)]
                            for wid in ids:
                                out_links.append(dict(file=fname, sheet=sheet, r=rc["r"], work=wid, year=yi, how="direct"))
            elif not plan:
                # ---------- 工数・日数
                if "人員" in f and "時間" in f and "manDays" not in f:
                    n = f["人員"] if isinstance(f["人員"], (int, float)) else persons(f["人員"])
                    h = hours(f["時間"])
                    if n and h:
                        f["manDays"] = round(n * h / 8.0, 2)
                        f["manDaysText"] = "%s・%s" % (f["人員"], f["時間"])
                if rc["kind"] == "item" and ("予定開始" in f or "実績開始" in f):
                    a0 = f.get("実績開始") or f.get("予定開始")
                    b0 = f.get("実績終了") or f.get("予定終了")
                    try:
                        f["days"] = (dt.date.fromisoformat(b0) - dt.date.fromisoformat(a0)).days + 1
                    except Exception:
                        pass
                # ---------- 結び
                key = (fname, a["records"][0].get("sheet") if False else rc.get("sheet") if not (is_pdf or is_csv) else None, rc["r"])
                ids = links_v1.get((fname, None if (is_pdf or is_csv) else sheet, rc["r"])) or links_v1.get((fname, "" if is_csv else sheet, rc["r"])) or []
                day = f.get("日付") or f.get("月日")
                for wid in ids:
                    w = works.get(wid)
                    how = "direct"
                    if is_csv or "発注控" in fname:
                        ck = (w or {}).get("contract", "")
                        nm = str(f.get("件名") or f.get("契約件名") or "")
                        if ck.split("-")[0] in ("KIKAN", "DXN", "AM", "AMSTOP"):
                            how = "upgrade"
                        elif "発注控" in fname or "購入" in nm:
                            how = "rollup"
                        elif "保守点検" in nm or "単価契約" in nm:
                            how = "annual"
                        else:
                            how = "bundle"
                    l = dict(file=fname, sheet=None if (is_pdf or is_csv) else sheet, r=rc["r"], work=wid, how=how)
                    if day and how == "direct" and ("日報" in fname):
                        l["day"] = day
                        day_of_work[wid].add(day)
                    out_links.append(l)
            rc["fields"] = f
            recs.append(rc)
        json.dump({"file": fname, "plan": plan, "records": recs}, open(os.path.join(ANS, "files", stem + ".json"), "w", encoding="utf-8"),
                  ensure_ascii=False, indent=0)

    # ---------- world.json
    eq = w1["equipment"]
    fur = lambda u: {"1号炉": "1", "2号炉": "2", "3号炉": "3"}.get(u, "共通")
    assets = []
    for e in eq:
        sub = e.get("subsystem")
        names = {"機器台帳（R8.10・k3）": "%s > %s > %s" % (e["system"], sub, e["name"]),
                 "機器台帳（H26.4・R3末・k2）": "%s > %s > %s" % ("灰溶融（休止）" if e["idle"] else e["unit"], e["system"], e["names"].get("k2台帳（H26・R3）", e["name"])),
                 "中長期整備計画（R1〜R8・k3）": "%s > %s > %s > %s" % ("灰溶融（休止）" if e["idle"] else e["unit"], e["system"], sub, e["name"]),
                 "中長期整備計画（H23・H30・k2）": "%s > %s > %s" % (e["system"], sub, e["name"])}
        if e["installFY"] <= 2003:
            names["機器台帳（H15・k1）"] = "%s > %s" % (e["unit"], e["names"].get("k1台帳（H15）", e["name"]) + ("（予備）" if e.get("duty") == "予備" else ""))
        for k, v in (e.get("aliases") or {}).items():
            names[k] = v
        assets.append(dict(id=e["tag"], name=e["name"], tag=e["tag"], system=e["system"], furnace=fur(e["unit"]), installed=e["installFY"],
                           class_=None, part=None, names=names))
        assets[-1]["class"] = ("休止" if e["idle"] else e["importance"])
        del assets[-1]["class_"]
    wout = []
    for w in works.values():
        days = len(day_of_work.get(w["id"], ())) or None
        if days is None and w["start"] and w["end"]:
            days = (dt.date.fromisoformat(w["end"]) - dt.date.fromisoformat(w["start"])).days + 1
        wout.append(dict(id=w["id"], asset=w["tag"], fy=w["fy"], status="done", task=w["what"], attr=w["kind"], start=w["start"], end=w["end"],
                         manDays=w["manDays"], days=days, contract=w.get("contract"), contractName=w.get("contractName"), vendor=w.get("vendor"),
                         detail=w.get("detail")))
    for p in pseudo.values():
        wout.append(dict(id=p["id"], asset=p["tag"], fy=p["fy"], status=p["status"], task=p["task"], attr=p["attr"], start=None, end=None,
                         manDays=None, days=None, note="中長期整備計画の印から（%s）" % ("実施されなかった予定" if p["status"] == "not_done" else "記録時点の予定")))
    # 2026 年度 下期の定修（契約済み・未実施）の作業も予定として持つ
    k = 0
    for u in w1["unrecorded"]:
        if "2026 年度 下期" in u.get("reason", "") and u.get("tag"):
            k += 1
            wout.append(dict(id="Q%05d" % k, asset=u["tag"], fy=2026, status="planned", task=u["what"], attr=_attr(u["what"]), start=None, end=None,
                             manDays=None, days=None, note=u["reason"]))
    world = OrderedDict()
    world["schema"] = "moyai.sample-world/0"
    world["facility"] = {"id": "F03", "name": "【架空】くぬぎ台環境センター"}
    world["profile"] = {"furnace": ["1", "2", "3"], "furnaceText": "100t/24h×3 炉（並行揺動式ストーカ・全連続）。共通設備は「共通」"}
    world["maker"] = {"today": "2026-11-13"}
    world["assets"] = assets
    world["works"] = wout
    world["links"] = out_links
    world["errata"] = []
    world["files"] = [{"file": x["file"]} for x in w1["files"]]
    json.dump(world, open(os.path.join(ANS, "world.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=0, default=str)
    extra = OrderedDict()
    extra["note"] = ("world.json（moyai.sample-world/0）に無い欄をここに移した。errata は該当なし：入れた計画の欠けは小計の食い違い・法定点検の行の抜け・"
                     "周期の超え・計画と実績の不一致で、印の抜け（omitMark）と前年度の実績の反映漏れ（lastYearUnreflected）は入れていない。")
    extra["profile"] = w1["profile"]
    extra["equipment"] = eq
    extra["contracts"] = w1["contracts"]
    extra["unrecorded"] = w1["unrecorded"]
    extra["planDefects"] = w1["planDefects"]
    extra["filesByKind"] = w1["files"]
    extra["howRule"] = {"direct": "その行が作業そのもの（日報・工程表・報告書・台帳の行、星取表の印）", "bundle": "定期整備・補修工事などの契約の行（CSV）",
                        "annual": "年間の委託（クレーン保守点検など）", "rollup": "部品・物品の購入の行（CSV・部品発注控）", "upgrade": "改良工事の契約の行"}
    extra["planMarkRule"] = ("k3 様式（計／実 2 段）：計の行は marks＝plans＝計の行の記号、実の行は marks＝実の行の記号・plans＝同じ機器の計の行の記号。"
                             "k1・k2 様式（予定だけ）は marks＝plans。印の結び：実績（同じ機器・同じ年度・記号に合う種類の作業）があればその作業、"
                             "無ければ 2026 年度より前は not_done、2026 年度以後は planned の作業（P 番号、版をまたいで機器・年度・種類で 1 件）。"
                             "k1（H13）版は系統単位の行で機器が 1 台に決まらないので印の結びを付けない。法定点検の行も機器に結ばない。")
    json.dump(extra, open(os.path.join(ANS, "world.extra.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=0, default=str)
    print("files %d / works %d (planned+not_done %d) / links %d" % (len(files), len(wout), len(pseudo) + k, len(out_links)))


def _attr(t):
    if "更新" in t:
        return "REPLACE"
    if any(x in t for x in ("点検", "測定", "検査", "診断", "試験")):
        return "INSPECT"
    return "OVERHAUL"


if __name__ == "__main__":
    main()
