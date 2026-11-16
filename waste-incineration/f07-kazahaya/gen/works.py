# -*- coding: utf-8 -*-
"""実際に行った作業（正解）。契約の工程に載った作業は工程の実績日、載らない作業は SPC 直営・委託として日付を置く。"""
from datetime import date, timedelta
from world import ALL, BY_CODE, actual, TEISHU, h
from contracts_full import get, CREWS, zentei_scope, TODAY

ATTR = {"I": "INSPECT", "L": "INSPECT", "O": "OVERHAUL", "R": "OVERHAUL", "X": "REPLACE", "U": "UPGRADE"}
WORKS = []        # dict(id, code, attr, t, work, start, end, manDays, vendor, contract, task_no, recorded_in)
BY_ID = {}
TASK_WORKS = {}   # (contract key, task no) -> [work id]
UNRECORDED = []   # (work id, reason)


def workdays(a, b, holi):
    n, d = 0, a
    while d <= b:
        if not (d.weekday() == 6 or False):
            n += 1
        d += timedelta(days=1)
    return max(n, 1)


def build():
    if WORKS:
        return
    seen = set()
    cs = get()
    for c in cs:
        for t in c["tasks"]:
            if not t["works"]:
                continue
            comp, leader, n = CREWS[t["crew"]]
            a, b = t["as_"], t["ae"]
            done = not (c.get("cut") and b > c["cut"])
            md = round(n * workdays(a, b, c["holi"]) / len(t["works"]) * 2) / 2 if n else 0
            for (code, tt, w) in t["works"]:
                key = (c["fy"], code, tt, w)
                if key in seen:
                    continue
                seen.add(key)
                wid = f"W{c['fy']}-{code}-{tt}{len(WORKS):04d}"
                rec = dict(id=wid, code=code, attr=ATTR[tt], t=tt, work=w, start=a.isoformat(),
                           end=b.isoformat() if done else None, manDays=md, vendor=comp, contract=c["name"],
                           contract_key=c["key"], task_no=t["no"], status="完了" if done else "施工中")
                WORKS.append(rec)
                BY_ID[wid] = rec
                TASK_WORKS.setdefault((c["key"], t["no"]), []).append(wid)
    # 契約に載らない作業
    for e in ALL:
        for y, lst in actual(e).items():
            for (tt, w, note) in lst:
                key = (y, e["code"], tt, w)
                if key in seen:
                    continue
                seen.add(key)
                s0, _ = TEISHU[(e["unit"], y)]
                if e["base"] in ("ごみクレーン", "灰クレーン", "タービン室天井クレーン"):
                    vendor, why = "【架空】白峰クレーン株式会社", "クレーン検査は委託（発注台帳の委託行）。工事帳票には載らない"
                    d = date(y, 10, 20) if y < 2026 else date(y, 10, 20)
                elif e["base"] == "計量機":
                    vendor, why = "【架空】秤屋計測株式会社", "計量法の定期検査は委託（発注台帳）。工事帳票には載らない"
                    d = date(y, 9, 10)
                elif e["base"] in ("空気槽",):
                    vendor, why = "SPC 保全（直営）", "第二種圧力容器の定期自主検査は SPC 直営。記録は PM-NAVI のみ"
                    d = s0 + timedelta(days=2)
                elif e["unit"] in ("1", "2"):
                    vendor, why = "SPC 保全（直営）", "定修工事の帳票に記載が無い（受注者の記載漏れ。PM-NAVI には SPC が実績入力）"
                    d = s0 + timedelta(days=10)
                else:
                    vendor, why = "SPC 保全（直営）", "全停を要しない共通機器は SPC 保全班が直営で整備（予備機に切替えて随時）。記録は PM-NAVI と計画の実績欄のみ"
                    d = date(y, 1 + (h(e["code"] + w, 12) + 3) % 12 or 1, 10 + h(e["code"], 10))
                    if d.month < 4:
                        d = date(y + 1, d.month, d.day)
                    if y == 2026 and d > date(2026, 11, 1):
                        d = date(2026, 7, 10 + h(e["code"], 10))
                wid = f"W{y}-{e['code']}-{tt}{len(WORKS):04d}"
                done = not (y == 2026 and note == "施工中")
                rec = dict(id=wid, code=e["code"], attr=ATTR[tt], t=tt, work=w, start=d.isoformat(), end=d.isoformat() if done else None,
                           manDays=1.0 if tt in "IL" else 2.0, vendor=vendor, contract=None, contract_key=None, task_no=None,
                           status="完了" if done else "施工中")
                WORKS.append(rec)
                BY_ID[wid] = rec
                UNRECORDED.append((wid, why))


build()

if __name__ == "__main__":
    from collections import Counter
    print(len(WORKS), Counter(w["vendor"] for w in WORKS).most_common(8))
    print(len(UNRECORDED), Counter(r for _, r in UNRECORDED))
