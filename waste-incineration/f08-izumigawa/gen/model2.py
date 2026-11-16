# -*- coding: utf-8 -*-
"""日ごとの作業（日報の中身）、作業への日付・工数の書き戻し、メーカー以外の工事（部分改良・補修請負・地元業者）。"""
import datetime as dt
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import rng  # noqa: E402
import jobs as J  # noqa: E402
from model import (D, DAY, add_work_days, is_off, maker_job, next_work, _task, BYID_FIND,  # noqa: E402
                   _friday_on_or_after)
from works import BYID, CONTRACTS, actual_works  # noqa: E402

CREW = {"機械": (3, 5), "製缶": (3, 4), "築炉": (2, 4), "電気": (1, 2), "足場": (3, 4)}
WEATHER = ["晴", "晴", "晴", "曇", "曇", "雨", "晴のち曇", "曇時々雨"]


# ------------------------------------------------------------------ ボイラ水管の肉厚（正解の推移）
def thickness(line, fy):
    """第2パス下部の最小肉厚（mm）。1号炉は H26 部分取替後に別の箇所が最小になる。"""
    if line == 1:
        if fy <= 2013:
            return round(4.25 - 0.13 * (fy - 2005), 2)
        return round(3.70 - 0.04 * (fy - 2014), 2)
    if fy <= 2016:
        return round(4.40 - 0.10 * (fy - 2005), 2)
    return round(3.30 - 0.01 * (fy - 2016), 2)


def _wd(task):
    out, d = [], task["as_"]
    while d <= task["ae"]:
        if not is_off(d):
            out.append(d)
        d += DAY
    return out


def _mnames(ws):
    return [BYID[w["eq"]]["name"] for w in ws]


def _verb(w):
    return {"INSPECT": "点検", "OVERHAUL": "分解整備", "REPLACE": "更新", "UPGRADE": "据付"}[w["kind"]]


def item_text(t, k, n, r, line, fy):
    """作業 t の k 日目（全 n 日）の作業内容。機器を日ごとに割り振り、同じ文を並べない。"""
    nm = t["name"]
    ws = t["works"]
    if nm.startswith("安全措置"):
        a0 = r.choice(["安全措置（電源断・縁切り・札掛け）、KY。足場組立開始", "電源断・弁閉止・札掛け確認、KY。足場材搬入・組立",
                       "縁切り・施錠確認（運転業者立会）、足場組立開始"])
        a1 = r.choice(["足場組立 完了、点検", "足場組立 完了、手すり・昇降設備確認", "足場 完了、作業床点検"])
        return [a0, a1][min(k, 1)] if n > 1 else "安全措置、足場組立"
    if nm.startswith("炉開放"):
        tot = r.randint(21, 26)
        if k == 0:
            return f"マンホール開放、炉内温度確認（{r.randint(45, 62)}℃）、換気。流動砂抜出開始"
        if k == n - 1:
            return f"流動砂抜出 完了（計 {tot}t、ふるい分けて再使用分を保管）"
        return f"流動砂抜出（累計 約 {tot * (k + 1) // n}t）、炉床清掃"
    if nm.startswith("ボイラ開放"):
        seq = [r.choice(["ボイラ開放（マンホール 6 箇所）、ダスト払い", "ボイラ マンホール開放、炉内ダスト払い・搬出"]),
               r.choice(["ボイラ水洗（第1〜第3パス）", "ボイラ水洗、スケール・ダスト除去", "高圧水洗浄（第1〜第3パス）、排水処理"]),
               r.choice(["水管肉厚測定（超音波 84 点）", "水管肉厚測定 開始（84 点、超音波）"]),
               f"肉厚測定 完了（最小 {thickness(line, fy):.1f}mm／管理値 3.2mm）、スートブロワ点検"]
        if k < n - 1:
            return seq[min(k, 2)] if k < 3 else "エコノマイザ外観点検、ダスト付着状況確認"
        return seq[3]
    if nm.startswith("ボイラ性能検査"):
        return (r.choice(["性能検査準備（内部清掃・付属品点検）", "性能検査準備（安全弁整備・水面計点検）"]) if k == 0 else
                r.choice(["ボイラ性能検査 受検（合格）", "性能検査 受検、合格（検査証 裏書）", "ボイラ性能検査 受検 合格、指摘なし"]))
    if nm.startswith("復旧"):
        return (r.choice(["マンホール閉止、流動砂張込", "流動砂張込（再使用分＋補充）、マンホール閉止"]) if k == 0 else
                r.choice(["復旧 完了、保温・養生撤去", "保温復旧、足場解体、片付け", "復旧 完了、足場解体・清掃"]))
    if nm.startswith("通風試験"):
        return r.choice(["通風試験・試運転（運転業者立会）、異常なし", "通風試験、各ダンパ作動確認 良好", "試運転（押込・誘引 連動）、異常なし"])
    if not ws:
        return nm
    chunk = max(1, -(-len(ws) // n))
    part = ws[k * chunk:(k + 1) * chunk]
    if not part and not t.get("added"):     # 機器より日数が多い作業：段取りの文にする（同じ文を並べない）
        rest = k - -(-len(ws) // chunk)
        m = BYID[ws[rest % len(ws)]["eq"]]["name"]
        if "クレーン" in nm:
            seq = [f"{m} ワイヤロープ・ブレーキ点検", f"{m} 走行・横行装置 点検、給油", "性能検査 準備（荷重試験用 おもり手配）",
                   "クレーン性能検査 受検・年次自主検査記録整理"]
        else:
            seq = r.choice([[f"{m} 分解部品 清掃・寸法計測", f"{m} 組立・芯出し", f"{m} 試運転・作動確認", f"{nm} 手直し・点検記録整理"],
                            [f"{m} 部品 洗浄・摩耗測定", f"{m} 消耗品交換・組立", f"{m} 調整・試運転", f"{nm} 記録整理・片付け"],
                            [f"{m} 摺動部 点検・計測", f"{m} 締付け確認・給油", f"{m} 単体試運転", f"{nm} 手直し"]])
        return seq[rest % len(seq)]
    part = part or ws[-1:]
    if t.get("added"):
        e = BYID[part[0]["eq"]]
        return (f"{e['name']} 不具合部 分解・補修" if k == 0 else f"{e['name']} 補修 完了、試運転")
    bits = []
    for w in part:
        e = BYID[w["eq"]]
        v = _verb(w)
        if w["kind"] == "OVERHAUL" and k == n - 1:
            v = "組立・試運転"
        if w["kind"] == "REPLACE" and "ろ布" in w["label"]:
            v = f"ろ布交換（{min(480, 120 * (k + 1))}/480 本）"
        bits.append(f"{e['name']} {v}")
    return "、".join(bits)


def build_days(areas, fy, seed):
    """日 → {weather, items[(area, 設備・機器, text, n, prog, han, task_no, [wid])], note, tomorrow, crew}"""
    r = rng(f"days{seed}")
    days = {}
    for a in areas:
        for t in a["tasks"]:
            if t["han"] == "運転":
                continue
            wd = _wd(t)
            for k, d in enumerate(wd):
                lo, hi = CREW.get(t["han"], (2, 3))
                n = r.randint(lo, hi) + (1 if any(w["kind"] in ("OVERHAUL", "REPLACE", "UPGRADE") for w in t["works"]) else 0)
                prog = "完" if k == len(wd) - 1 else f"{int((k + 1) * 100 / len(wd)) // 10 * 10}%"
                lab = (BYID[t["works"][0]["eq"]]["base"] if t["works"] else t["name"].split(" ")[0].split("・")[0])
                txt = item_text(t, k, len(wd), r, a["line"], fy)
                dd = days.setdefault(d, dict(weather=r.choice(WEATHER), items=[], notes=[]))
                dd["items"].append((a["label"], lab, txt, n, prog, t["han"], t["no"], [w["wid"] for w in t["works"]]))
                if t.get("added") and k == 0:
                    e = BYID[t["works"][0]["eq"]] if t["works"] else None
                    if e:
                        dd["notes"].append(f"{a['label']} {e['name']} に不具合を確認、監督員へ報告し補修を追加")
                if t["pe"] and t["ae"] > t["pe"] and d == t["pe"]:
                    dd["notes"].append(f"{t['name']} は予定より遅れ、完了は {t['ae'].month}/{t['ae'].day} の見込み")
    for d, dd in days.items():
        crew = {}
        for it in dd["items"]:
            crew[it[5]] = crew.get(it[5], 0) + it[3]
        crew = {k: max(1, int(v * 0.85)) for k, v in crew.items()}
        crew["作業員"] = r.randint(2, 6)
        dd["crew"] = crew
        dd["note"] = "。".join(dd["notes"])
    ds = sorted(days)
    for k, d in enumerate(ds):
        nxt = ds[k + 1] if k + 1 < len(ds) else None
        days[d]["tomorrow"] = ("、".join(sorted({it[1] for it in days[nxt]["items"]}))[:60] if nxt else "完了")
    return days


def assign_mandays(job):
    """日報の人数から作業ごとの工数（人日）を割り振る。"""
    per_task = {}
    for d, dd in job["days"].items():
        for it in dd["items"]:
            per_task.setdefault((it[0], it[6]), 0)
            per_task[(it[0], it[6])] += it[3]
    for a in job["areas"]:
        for t in a["tasks"]:
            ws = t["works"]
            if not ws:
                continue
            tot = per_task.get((a["label"], t["no"]), 0)
            wt = [2 if w["kind"] in ("OVERHAUL", "REPLACE", "UPGRADE") else 1 for w in ws]
            for w, x in zip(ws, wt):
                w["start"], w["end"] = t["as_"], t["ae"]
                w["manDays"] = round(tot * x / sum(wt), 1)
                w["task"] = (job["cid"], a["label"], t["no"])


# ------------------------------------------------------------------ 部分改良（P2008）
def p2008_job():
    W = [w for w in actual_works() if w["contract"] == "P2008"]
    areas = []
    for line, start in ((1, D(2009, 1, 9)), (2, D(2010, 1, 8))):
        ws = [w for w in W if BYID[w["eq"]]["line"] == line]
        g = lambda keys: [w for w in ws if any(k in BYID[w["eq"]]["base"] for k in keys)]
        w0 = next_work(start + 3 * DAY)
        T = [_task(1, "立下げ・冷却（運転業者）", [], "運転", start, start + 2 * DAY, start, start + 2 * DAY),
             _task(2, "安全措置・足場仮設", [], "足場", w0, add_work_days(w0, 2), w0, add_work_days(w0, 2)),
             _task(3, "流動砂抜出・既設散気装置 撤去", [], "機械", w0, add_work_days(w0, 5), w0, add_work_days(w0, 5))]
        s = add_work_days(w0, 6)
        T.append(_task(4, "散気装置 据付（ヘッダ・ノズル 120 本）", [w["eq"] for w in g(["散気"])], "機械", s,
                       add_work_days(s, 8), s, add_work_days(s, 9 if line == 1 else 8), g(["散気"])))
        T.append(_task(5, "不燃物抜出装置 撤去・据付", [w["eq"] for w in g(["不燃物抜出"])], "機械", s,
                       add_work_days(s, 7), s, add_work_days(s, 7), g(["不燃物抜出"])))
        s2 = add_work_days(w0, 3)
        T.append(_task(6, "砂循環設備 撤去・据付", [w["eq"] for w in g(["砂", "振動", "不燃物排出", "冷却スクリュー"])],
                       "機械", s2, add_work_days(s2, 14), s2, add_work_days(s2, 15),
                       g(["砂", "振動", "不燃物排出", "冷却スクリュー"])))
        s3 = add_work_days(s2, 10)
        T.append(_task(7, "砂循環設備現場盤 新設・配線", [w["eq"] for w in g(["現場盤"])], "電気", s3,
                       add_work_days(s3, 6), s3, add_work_days(s3, 6), g(["現場盤"])))
        e = max(t["ae"] for t in T)
        ep = max(t["pe"] for t in T)
        T.append(_task(8, "復旧・流動砂張込", [], "機械", next_work(ep + DAY), add_work_days(next_work(ep + DAY), 2),
                       next_work(e + DAY), add_work_days(next_work(e + DAY), 2)))
        e2 = T[-1]["ae"]
        T.append(_task(9, "試運転調整（散気均一性・砂循環量）", [], "機械", next_work(T[-1]["pe"] + DAY),
                       add_work_days(next_work(T[-1]["pe"] + DAY), 3), next_work(e2 + DAY), add_work_days(next_work(e2 + DAY), 3)))
        e3 = T[-1]["ae"]
        T.append(_task(10, "立上げ（運転業者）", [], "運転", T[-1]["pe"] + DAY, T[-1]["pe"] + 2 * DAY, e3 + DAY, e3 + 2 * DAY))
        areas.append(dict(label=f"{line}号炉", line=line, tasks=T,
                          span=(start, e3 + 2 * DAY)))
    return dict(cid="P2008", fy=2008, contract=CONTRACTS["P2008"], areas=areas, works=W, vendor="m1")


def r2014_job(t2014):
    """H26 1号炉ボイラ水管補修工事（メーカー請負）。定期整備と同じ停止期間に、別の契約で行う。"""
    W = [w for w in actual_works() if w["contract"] == "R2014"]
    a1 = next(a for a in t2014["areas"] if a["line"] == 1)
    bo = next(t for t in a1["tasks"] if t["name"].startswith("ボイラ開放"))
    s = next_work(bo["ae"] + DAY)
    T = [_task(1, "水管切断・撤去（第2パス 6 本）", ["1D-001"], "製缶", s, add_work_days(s, 2), s, add_work_days(s, 2), W),
         _task(2, "新管 取付・溶接", ["1D-001"], "製缶", add_work_days(s, 3), add_work_days(s, 5), add_work_days(s, 3), add_work_days(s, 5)),
         _task(3, "放射線透過試験", ["1D-001"], "製缶", add_work_days(s, 6), add_work_days(s, 6), add_work_days(s, 6), add_work_days(s, 6)),
         _task(4, "水圧試験", ["1D-001"], "製缶", add_work_days(s, 7), add_work_days(s, 7), add_work_days(s, 7), add_work_days(s, 7))]
    end = T[-1]["ae"]
    # 定期整備の性能検査以降を、補修の後へずらす（同時にできない）
    shift = 0
    for t in a1["tasks"]:
        if t["name"].startswith("ボイラ性能検査") and t["as_"] <= end:
            shift = (end - t["as_"]).days + 1
    if shift:
        delta = dt.timedelta(days=shift)
        moving = False
        for t in a1["tasks"]:
            if t["name"].startswith("ボイラ性能検査"):
                moving = True
            if moving:     # 定期整備の工程も補修工事を見込んで組んだ（予定・実績とも）
                for k in ("ps", "pe", "as_", "ae"):
                    if t[k]:
                        t[k] = t[k] + delta if t["han"] == "運転" else next_work(t[k] + delta)
        a1["span"] = (a1["span"][0], max(t["ae"] for t in a1["tasks"]))
    return dict(cid="R2014", fy=2014, contract=CONTRACTS["R2014"],
                areas=[dict(label="1号炉", line=1, tasks=T, span=(T[0]["as_"], end))], works=W, vendor="m1")


def all_maker_jobs():
    jobs = [maker_job(fy) for fy in range(2005, 2019)]
    t14 = next(j for j in jobs if j["fy"] == 2014)
    extra = [p2008_job(), r2014_job(t14)]
    for j in jobs + extra:
        if j["fy"] == 2010 and j["cid"] == "T2010":
            j["days"] = build_days(j["areas"], j["fy"], j["cid"])
            for d, v in J.J22_DAILY.items():      # 見本の 1号炉 9/6〜9/14 は見本の日報のまま
                tno = {}
                items = []
                for (ln, eq, txt, n, prog) in v["items"]:
                    items.append((ln, eq, txt, n, prog, "機械", _guess_task(j, ln, eq), _wid_for(j, ln, eq)))
                j["days"][d] = dict(weather=v["weather"], items=items, note=v["note"], tomorrow=v["tomorrow"],
                                    crew=v["crew"], sample=True)
        else:
            j["days"] = build_days(j["areas"], j["fy"], j["cid"])
        assign_mandays(j)
    return jobs + extra


def _guess_task(j, ln, eq):
    a = next(a for a in j["areas"] if a["label"] == ln) if ln != "1号炉" or True else None
    a = next(a for a in j["areas"] if a["label"] == "1号炉")
    keys = {"流動床炉": "炉", "散気装置": "炉内", "廃熱ボイラ": "ボイラ", "エコノマイザ": "ボイラ", "スートブロワ": "ボイラ",
            "不燃物抜出装置": "不燃物", "砂循環設備": "砂循環", "バグフィルタ": "バグ", "給じん設備": "給じん",
            "ボイラ給水ポンプ": "ボイラ給水", "減温塔": "減温塔", "共通": "安全", "足場": "安全"}
    k = keys.get(eq, eq)
    for t in a["tasks"]:
        if t["name"].startswith(k) or k in t["name"]:
            return t["no"]
    return 0


def _wid_for(j, ln, eq):
    a = next(a for a in j["areas"] if a["label"] == "1号炉")
    no = _guess_task(j, ln, eq)
    t = next((t for t in a["tasks"] if t["no"] == no), None)
    return [w["wid"] for w in t["works"]] if t else []


if __name__ == "__main__":
    js = all_maker_jobs()
    W = actual_works()
    have = [w for w in W if w["contract"][:1] in "TPR"]
    print(len(js), sum(len(j["days"]) for j in js), sum(1 for w in have if "manDays" in w), len(have))
    print([(j["cid"], len(j["days"])) for j in js][:6])
