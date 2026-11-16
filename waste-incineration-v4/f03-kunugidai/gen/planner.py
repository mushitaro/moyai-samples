# -*- coding: utf-8 -*-
"""計画の立て方（その版の作成時点で知り得た履歴だけを使う）。
style="k2"：周期＋前回整備から置く（平成 23 年度の長寿命化計画の作り方。k3 が H30 に引き継いで更新）。
style="k3"：k2 の置き方に、健全度・重要度・2 回目大規模改修の判断待ち（2029 年度以降の大きな更新は括弧付き）を足す。
記号  ○点検・測定  ◇補修  ◎分解整備・部分取替  ●更新  (●)方針決定後に判断
"""
from world import CLS, find

PENDING_FROM = 2029
LEVEL_ASOF = [2026]   # 2 回目の大規模改修か建替えかの判断が要る年度（k3 の注記）


def _known(e, as_of):
    return [x for x in e.events if x["fy"] < as_of]


def plan_tag(e, as_of, years, style):
    """{fy: [(sym, cost千円, 内容)]}"""
    c = CLS[e.cls]
    out = {}
    if e.idle and as_of > 2013:
        return out
    if e.install > as_of:
        return out
    kn = _known(e, as_of)
    base = e.install
    for x in kn:
        if x["kind"] in ("REPLACE", "UPGRADE"):
            base = max(base, x["fy"])
    last_rep = base
    last_oh = max([base] + [x["fy"] for x in kn if x["kind"] == "OVERHAUL"])
    h = e.health if style == "k3" else 0
    nrep = last_rep + c["rep"] + 2 * h if c.get("rep") else None
    noh = last_oh + c["oh"] + h if c.get("oh") else None
    first = years[0]
    if nrep is not None and nrep < first:
        if c["crep"] >= 50000:
            nrep = first + 5                 # 大物（ストーカ・BF・クレーン等）の期限超えは 5 年先に置いて検討（k2 の書き方を k3 も踏襲）
        else:
            nrep = first + "ABC".index(e.imp)   # 期限超えは重要度の順に振り分け
    if noh is not None and noh < first:
        noh = first if e.imp in ("A", "B") else first + 1
    for y in years:
        if nrep is not None and y == nrep:
            sym = "●"
            if style == "k3" and as_of >= 2024 and y >= PENDING_FROM and c["crep"] >= 5000:
                sym = "(●)"
            out.setdefault(y, []).append((sym, c["crep"], c.get("rep_name", "更新")))
            nrep = y + c["rep"]
            noh = y + c["oh"] if c.get("oh") else None
            continue
        if noh is not None and y == noh:
            out.setdefault(y, []).append(("◎", c["coh"], c.get("oh_name", "分解整備")))
            noh = y + c["oh"]
            continue
        if c.get("ins"):
            if c["ins"] == 1 or (y - base) % c["ins"] == 0:
                if style == "k3" or c["ins"] <= 2 or e.imp == "A":
                    out.setdefault(y, []).append(("○", 0, c.get("ins_name", "点検")))
    return out


# 代表機器の版ごとの計画（根拠記録 notes/02 と一致）。{(版, unit, name): {fy: [(sym, cost, what)]}}
OVR = {
    # ---- 令和8年度版（k3、2026.4 作成） ----
    (2026, 2, "誘引通風機"): {2026: [("○", 0, "振動診断")], 2027: [("○", 0, "振動診断")], 2028: [("◎", 4200, "軸受交換・バランス")],
                         2029: [("●", 9800, "羽根車更新")], 2030: [("○", 0, "振動診断")], 2031: [("○", 0, "振動診断")],
                         2032: [("◎", 4200, "軸受交換")], 2033: [("○", 0, "振動診断")], 2034: [("○", 0, "振動診断")], 2035: [("(●)", 38000, "本体更新")]},
    (2026, 2, "バグフィルタ"): {2027: [("○", 0, "ろ布強度試験")], 2029: [("○", 0, "ろ布強度試験")], 2031: [("◎", 16000, "ろ布全数交換")], 2034: [("(●)", 160000, "本体更新")]},
    (2026, 2, "過熱器"): {2026: [("○", 0, "肉厚測定")], 2027: [("●", 42000, "1次過熱器管 全数取替")], 2028: [("○", 0, "肉厚測定")],
                       2029: [("○", 0, "肉厚測定")], 2030: [("○", 0, "肉厚測定")], 2031: [("○", 0, "肉厚測定")], 2032: [("○", 0, "肉厚測定")],
                       2033: [("○", 0, "肉厚測定")], 2034: [("○", 0, "肉厚測定")], 2035: [("○", 0, "肉厚測定")]},
    (2026, 1, "燃焼ストーカ"): {2026: [("◎", 9500, "火格子部分交換 44枚")], 2027: [("○", 0, "摩耗測定")], 2029: [("◎", 9500, "火格子部分交換")],
                           2031: [("(●)", 98000, "ストーカ更新")]},
    (2026, 0, "ごみクレーン 1号"): {2026: [("◎", 3200, "ワイヤロープ交換")], 2029: [("◎", 3200, "ワイヤロープ交換")], 2032: [("◎", 3200, "ワイヤロープ交換")], 2033: [("(●)", 72000, "更新")]},
    (2026, 0, "ごみクレーン 2号"): {2027: [("◎", 3200, "ワイヤロープ交換")], 2030: [("◎", 3200, "ワイヤロープ交換")], 2033: [("(●)", 72000, "更新")]},
    (2026, 0, "蒸気タービン"): {2028: [("◎", 38000, "開放点検")], 2032: [("◎", 38000, "開放点検")]},
    (2026, 3, "湿式洗煙塔"): {2027: [("○", 0, "内面点検")], 2029: [("○", 0, "内面点検")], 2031: [("○", 0, "内面点検")], 2033: [("○", 0, "内面点検")], 2035: [("○", 0, "内面点検")]},
    (2026, 2, "洗煙循環ポンプ A"): {2028: [("◎", 1200, "分解整備")], 2032: [("◎", 1200, "分解整備")], 2034: [("●", 5200, "更新")]},
    (2026, 2, "洗煙循環ポンプ B"): {2026: [("◎", 1200, "分解整備")], 2030: [("◎", 1200, "分解整備")], 2034: [("●", 5200, "更新")]},
    (2026, 0, "非常用発電機"): {2027: [("●", 48000, "更新（500kVA→同容量）")]},
    (2026, 1, "ボイラ給水ポンプ A"): {2029: [("◎", 1200, "分解整備")], 2030: [("(●)", 5200, "更新")]},
    (2026, 1, "主灰押出装置"): {2026: [("●", 6500, "シュート更新")], 2030: [("◎", 900, "分解整備")]},
    # ---- 平成30年度版（k2 様式、k3 が 2018.5 改訂） ----
    (2018, 2, "過熱器"): {y: [("○", 0, "肉厚測定")] for y in range(2018, 2028) if y != 2023} | {2023: [("●", 38000, "1次過熱器管 取替")]},
    (2018, 3, "湿式洗煙塔"): {2018: [("◎", 30000, "ライニング更新")], 2020: [("○", 0, "内面点検")], 2022: [("○", 0, "内面点検")], 2024: [("○", 0, "内面点検")], 2026: [("○", 0, "内面点検")]},
    (2018, 0, "非常用発電機"): {2024: [("◎", 3500, "オーバーホール")]},
    (2018, 2, "誘引通風機"): {2021: [("◎", 4200, "軸受交換")], 2025: [("◎", 4200, "軸受交換")]},
    (2018, 0, "蒸気タービン"): {2020: [("◎", 36000, "開放点検")], 2024: [("◎", 36000, "開放点検")]},
}


def plan_all(eqs, as_of, years, style):
    res = {}
    for e in eqs:
        key = (as_of, e.unit, e.name)
        if key in OVR:
            res[e.tag] = {y: v for y, v in OVR[key].items() if y in years}
        else:
            res[e.tag] = plan_tag(e, as_of, years, style)
        if as_of == 2011 and e.origin == "kikan":
            from world import KIKAN_FY
            ky = KIKAN_FY[e.unit]
            if ky in years:
                res[e.tag] = {y: [t for t in v if t[0] not in ("●", "(●)")] for y, v in res[e.tag].items() if y > ky}
                res[e.tag] = {y: v for y, v in res[e.tag].items() if v}
                res[e.tag][ky] = [("●", CLS[e.cls]["crep"], "基幹的設備改良で更新")]
    if style == "k3":
        LEVEL_ASOF[0] = as_of
        _level(eqs, res, years)
    return res


def _level(eqs, res, years):
    """同じ号炉・同じ年度に大きな更新（2,000 万円以上）が 3 件以上重なったら、重要度の低いものを翌年度へ送る。"""
    by = {}
    for e in eqs:
        for y, items in res[e.tag].items():
            for s, cost, w in items:
                if s == "●" and cost >= 20000:
                    by.setdefault((e.unit, y), []).append((e, s, cost, w))
    for (u, y), lst in sorted(by.items(), key=lambda kv: kv[0]):
        if len(lst) <= 2:
            continue
        lst.sort(key=lambda t: ("ABC".index(t[0].imp), -t[2]))
        for e, s, cost, w in lst[2:]:
            if (2026, e.unit, e.name) in OVR or y + 1 not in years:
                continue
            res[e.tag][y] = [t for t in res[e.tag][y] if t[0] != "●"]
            if not res[e.tag][y]:
                del res[e.tag][y]
            res[e.tag].setdefault(y + 1, []).append(("●" if (y + 1 < PENDING_FROM or LEVEL_ASOF[0] < 2024) else "(●)", cost, w + "（平準化で翌年度）"))


ACT_SYM = {"OVERHAUL": "◎", "REPLACE": "●", "UPGRADE": "●", "REPAIR": "◇", "INSPECT": "○"}


def actual(e, fy):
    return [(ACT_SYM[x["kind"]], x["what"]) for x in e.events if x["fy"] == fy]
