# -*- coding: utf-8 -*-
"""F06 正解の世界：機器の状態・整備の実績（年度単位）と、各版の計画の見え方。

kind: I=INSPECT（点検・計測）, O=OVERHAUL（分解整備・補修）, R=REPLACE（更新・主要部品交換）, U=UPGRADE（基幹改良）
"""
import random
import zlib

import equipment as EQ

SEED = 20261113
TODAY = "2026-11-13"
FY0, FY_END = 2004, 2040

# 機器の種類ごとの保全方針（kind, 基準周期年, 1 回あたり概算 千円）
STRATEGY = {
    "static": [("I", 10, 300)], "static_hot": [("I", 2, 150)], "static_stack": [("I", 5, 1800)],
    "stack_in": [("I", 3, 400), ("O", 9, 6500)],
    "mech_small": [("O", 5, 400)], "feeder": [("I", 1, 200), ("O", 4, 2600)], "feeder_s": [("O", 3, 450)],
    "hyd": [("I", 2, 300), ("O", 6, 1800)], "hyd_drive": [("O", 4, 1200)],
    "pump_dup": [("O", 5, 650)], "pump_corr": [("O", 3, 1400)], "pump_dosing": [("O", 3, 250)],
    "pump_slurry": [("O", 2, 500)], "pump_bfp": [("O", 4, 3800)], "pump_single": [("O", 5, 700)],
    "hex": [("I", 3, 350)], "furnace": [("I", 1, 600), ("O", 2, 9500)],
    "grate": [("I", 1, 400), ("O", 3, 7800), ("R", 13, 52000)],
    "grate_wc": [("I", 1, 500), ("O", 2, 11000), ("R", 12, 68000)],
    "fan_big": [("I", 1, 300), ("O", 4, 3200)], "fan_mid": [("O", 4, 1100)], "blower": [("O", 3, 380)],
    "damper": [("O", 6, 350)], "damper_bf": [("O", 4, 420)], "damper_idf": [("O", 3, 600)],
    "sah": [("I", 1, 250), ("O", 8, 6800)], "cv": [("O", 5, 380)], "cv_hot": [("O", 3, 520)],
    "cv_corr": [("O", 3, 300)], "mv": [("O", 8, 420)], "mv_hp": [("O", 4, 900)], "hv": [],
    "hv_hp": [("O", 8, 600)], "hv_corr": [("O", 6, 350)], "sv": [("I", 1, 180)], "burner": [("I", 2, 300)],
    "conveyor": [("O", 3, 1300)], "rv": [("O", 3, 380)], "boiler": [("I", 1, 2200), ("O", 3, 8500)],
    "sh": [("I", 1, 300), ("R", 18, 38000)], "sh2": [("I", 1, 300), ("R", 18, 42000)],
    "sh3": [("I", 1, 300), ("R", 15, 48000)], "eco": [("I", 1, 200), ("R", 20, 30000)],
    "sootblower": [("O", 3, 700)], "sootblower_r": [("O", 5, 400)], "gct": [("I", 1, 300), ("O", 5, 3500)],
    "bf": [("I", 1, 600), ("R", 5, 26000)], "scrubber": [("I", 1, 500), ("O", 7, 24000)],
    "me": [("O", 2, 900)], "rh": [("I", 1, 300), ("R", 22, 36000)], "scr": [("I", 2, 800), ("R", 7, 19000)],
    "idf": [("I", 1, 300), ("O", 4, 4200)], "ash_ext": [("I", 1, 200), ("O", 3, 3800)],
    "panel": [("I", 3, 200)], "panel_l": [("I", 5, 80)], "panel_inv": [("I", 2, 350)],
    "dcs": [("I", 1, 900)], "dcs_hmi": [("I", 1, 150)],
    "inst": [("I", 3, 60)], "inst_hot": [("I", 1, 60), ("R", 4, 180)], "inst_wet": [("I", 1, 90)],
    "inst_main": [("I", 1, 420)], "weigh": [("I", 2, 650)], "door": [("O", 4, 900)],
    "crane": [("I", 1, 800), ("O", 2, 4200)], "crane_sp": [("I", 1, 700), ("O", 3, 3800)],
    "crane_ash": [("I", 1, 500), ("O", 3, 1800)], "crane_m": [("I", 1, 250)],
    "silo": [("I", 5, 600)], "tank": [("I", 5, 450)], "dea": [("I", 1, 400), ("O", 6, 4800)],
    "header": [("I", 2, 700)], "pure": [("O", 3, 2400)],
    "turbine": [("O", 4, 98000)], "generator": [("O", 4, 16000)], "gearbox": [("O", 4, 9000)],
    "tbv": [("O", 4, 5200)], "gov": [("I", 1, 600)], "acc": [("I", 2, 900)], "acc_fan": [("O", 4, 1100)],
    "comp": [("O", 2, 900)], "agitator": [("O", 5, 650)], "dewater": [("O", 3, 2600)],
    "kneader": [("O", 2, 2800)], "gis": [("I", 6, 9000)], "tr_main": [("I", 3, 1600)], "tr": [("I", 6, 500)],
    "hv_panel": [("I", 3, 400)], "dcps": [("I", 3, 300)], "battery": [("R", 13, 6800)],
    "ups": [("I", 2, 400), ("R", 10, 5200)], "eg": [("I", 1, 600), ("O", 6, 5600)], "ct": [("O", 4, 1600)],
}

# 基幹改良（2022–2025）の対象：cls と、改良で施工する年
def upgrade_year(e, version="actual"):
    """version: 'actual'/'2020'（決定済み） or '2019'（平成31年度版の見込み：R3 から 1 炉ずつ）"""
    if version == "none":
        return None
    if version == "2019":
        fy = {"1": 2021, "2": 2022, "3": 2023, "C": 2024}
    else:
        fy = {"1": 2022, "2": 2023, "3": 2024, "C": 2025}
    cls, u = e["cls"], e["unit"]
    if u in "123" and (cls in ("sh", "sh2", "sh3", "idf", "dcs") or e["tag"].endswith(("-IPB-811", "-IO-814", "-CP-812"))):
        return fy[u]
    if u == "C" and (cls in ("dcs", "dcs_hmi", "gov") or e["tag"] in ("C-MSV-504", "C-GV-505", "C-PNL-524", "C-HU-522")):
        if version == "2019" and cls == "dcs_hmi":
            return None  # H31 版では HMI は単独更新の想定
        return fy["C"]
    return None


def _rng(tag, salt=""):
    return random.Random(zlib.crc32((tag + salt).encode()) ^ SEED)


# 個別に決めた実績（代表機器ほか）。(tag, year, kind, note)
FIXED = {}
def _fix(tag, yrs, kind, note=""):
    for y in yrs:
        FIXED.setdefault(tag, []).append((y, kind, note))

# タービン・発電機の定期事業者検査（4 年ごと、別契約）
for t in ("C-ST-501", "C-GEN-502", "C-RG-503", "C-NRV-506", "C-NRV-507"):
    _fix(t, [2006, 2010, 2014, 2018, 2022, 2026], "O", "定期事業者検査（開放）")
for t in ("C-MSV-504", "C-GV-505"):
    _fix(t, [2006, 2010, 2014, 2018, 2022], "O", "定期事業者検査（開放）")
# ろ布交換（号炉ごとに年をずらす）
_fix("1-BF-412", [2010, 2016, 2021], "R", "ろ布全数交換")
_fix("2-BF-412", [2011, 2017, 2022], "R", "ろ布全数交換")
_fix("3-BF-412", [2012, 2018, 2023], "R", "ろ布全数交換")
# 脱硝触媒 1 層交換
_fix("1-SCR-601", [2012, 2019], "R", "触媒 1 層交換"); _fix("2-SCR-601", [2013, 2020], "R", "触媒 1 層交換")
_fix("3-SCR-601", [2014, 2021], "R", "触媒 1 層交換")
# 3 次過熱器：2 号炉は 2017 に部分管交換（減肉）
_fix("2-SH-304", [2017], "O", "減肉管 6 本取替")
_fix("1-SH-304", [2018], "O", "減肉管 4 本取替")
# 再加熱器 2 号炉：2022 管漏れ 2 本施栓
_fix("2-RH-515", [2022], "O", "管 2 本漏れ 施栓")
# 誘引通風機 3 号炉：2019 振動上昇→2020 軸受交換
_fix("3-IDF-607", [2020], "O", "軸受交換（振動上昇）")
# 2 号炉ボイラ水管：2022 減肉→補修工事（別契約）
_fix("2-B-301", [2022], "O", "水冷壁管 4 本 肉盛・1 本取替（補修工事）")
# 火格子全面更新（燃焼火格子 水冷）
_fix("1-GR-203", [2015], "R", "燃焼火格子 全面更新"); _fix("2-GR-203", [2016], "R", "燃焼火格子 全面更新")
_fix("3-GR-203", [2017], "R", "燃焼火格子 全面更新")
# ボイラ給水ポンプ No.2：H30 補修工事（メーカー）
_fix("C-P-403", [2018], "O", "補修工事（メカシール・軸受・カップリング）")
# 3 号炉 落じんコンベヤ：H28 見送り → H29 チェーン交換
_fix("3-CV-226", [2017], "O", "チェーン交換（H28 見送り分）")
# 予備ごみクレーン：2018 ワイヤ・走行車輪
_fix("C-CR-203", [2018, 2023], "O", "ワイヤロープ・走行車輪交換")
# 主変圧器：2019 油中ガス分析でアセチレン微量→2020 再分析、2025 内部点検
_fix("C-TR-1002", [2020], "I", "油中ガス再分析"); _fix("C-TR-1002", [2025], "O", "内部点検（全炉停止時）")
# 2014 DCS 一部更新（オペレータステーション）
for t in ("C-OPS-1111", "C-OPS-1112", "C-OPS-1113", "C-OPS-1114", "C-OPS-1115", "C-EWS-1116", "C-LOG-1117"):
    _fix(t, [2014], "R", "DCS 一部更新")
# 2 号炉 洗煙循環ポンプ（交互運転）：インペラ摩耗の進みが違う
_fix("2-P-502A", [2007, 2010, 2013, 2016, 2019, 2022, 2025], "O", "分解整備（インペラ・メカシール）")
_fix("2-P-502B", [2008, 2012, 2015, 2018, 2020, 2022, 2024], "O", "分解整備（インペラ・メカシール）")

FIXED_CLASSES_SKIP = {"C-ST-501", "C-GEN-502", "C-RG-503", "C-NRV-506", "C-NRV-507", "C-MSV-504", "C-GV-505",
                      "1-BF-412", "2-BF-412", "3-BF-412", "1-SCR-601", "2-SCR-601", "3-SCR-601",
                      "2-P-502A", "2-P-502B"}

PARAM_OVERRIDE = {"1-GR-203": {"R": 30}, "2-GR-203": {"R": 30}, "3-GR-203": {"R": 30}, "2-RH-515": {"R": 22}, "1-RH-515": {"R": 24}, "3-RH-515": {"R": 25},
                  "C-BAT-1033": {"R": 13}, "1-GR-202": {"R": 14}}

POSTPONE_YEARS = {2009: ("C", "B"), 2016: ("C",)}  # 予算の都合で O/R を 1 年送り


def _params(e):
    r = _rng(e["tag"])
    out = []
    for kind, iv, cost in STRATEGY.get(e["cls"], []):
        # 重要度 C の計装・盤は定期の点検を置かない（故障時対応）。C の整備も 4 割は事後保全
        if e["imp"] == "C" and kind == "I" and e["cls"] in ("inst", "panel_l", "panel", "inst_wet"):
            continue
        if e["imp"] == "C" and kind in ("O", "R") and r.random() < 0.4:
            continue
        if e["cls"] == "inst" and kind == "I":
            iv = r.choice([1, 2]) if e["imp"] == "A" else r.choice([3, 4, 4])
        if iv == 1:
            ive = 1
        else:
            wf = r.choice([0.8, 1.0, 1.0, 1.0, 1.25, 1.5]) if e["imp"] != "A" else r.choice([0.8, 1.0, 1.0, 1.0])
            ive = max(1, round(iv * wf))
        # 炉別の点検（周期 3 以上）は号炉の順番に寄せるが、個体差でずれる
        if e["tag"] in PARAM_OVERRIDE and kind in PARAM_OVERRIDE[e["tag"]]:
            ive = PARAM_OVERRIDE[e["tag"]][kind]
        phase = r.randrange(ive) if ive > 1 else 0
        if kind == "O" and ive >= 4:
            phase = max(1, phase)
        if e["unit"] in "123" and kind == "I" and ive >= 3 and r.random() < 0.7:
            phase = (int(e["unit"]) - 1) % ive
        out.append(dict(kind=kind, iv=ive, base_iv=iv, phase=phase, cost=cost * r.choice([0.8, 0.9, 1.0, 1.1, 1.3])))
    return out


def findings(e):
    """状態の発見（年, kind） 実績側のみ。A/B 機器にまれに起きる。"""
    r = _rng(e["tag"], "f")
    out = []
    if e["imp"] in ("A", "B") and STRATEGY.get(e["cls"]):
        n = r.choice([0, 0, 0, 1, 1, 2]) if e["imp"] == "A" else r.choice([0, 0, 0, 0, 1])
        for _ in range(n):
            out.append(r.randrange(2006, 2026))
    if e["cls"] == "hv" and r.random() < 0.15:
        out.append(r.randrange(2008, 2026))  # 手動弁は漏れが出たときだけ
    return sorted(out)


def _first_year(e, prm):
    y0 = max(FY0, e["installed"])
    if e["installed"] > FY0:
        y0 = e["installed"] + 1
    if prm["iv"] == 1:
        return y0
    if prm["kind"] == "R":
        return max(FY0, e["installed"]) + prm["iv"]
    return y0 + prm["phase"] + (1 if prm["kind"] != "I" else 0)


def simulate(e, cutoff=None, uversion="actual", until=FY_END, start_hist=None):
    """機器 1 台の年度ごとの作業。
    cutoff=None：実績（2004–2026、それ以降は 2026 年 11 月時点の計画）。
    cutoff=Y：Y 年度の版の計画として、Y-1 年度までは実績、Y 年度以降は計画（知り得た事だけ）。
    戻り値 {year: [(kind, cost, note)]}
    """
    prms = _params(e)
    fnd = findings(e)
    fixed = FIXED.get(e["tag"], [])
    uy_act = upgrade_year(e, "actual")
    uy_plan = upgrade_year(e, uversion)
    res = {}
    last = {}
    first_due = {i: _first_year(e, p) for i, p in enumerate(prms)}
    rm = e["removed"]
    for y in range(FY0, until + 1):
        is_plan = cutoff is not None and y >= cutoff
        if cutoff is None and y > 2026:
            is_plan = True
        if y < max(FY0, e["installed"]):
            continue
        if rm is not None and y >= rm and not (is_plan and uversion in ("2019", "none")):
            break
        if is_plan and uversion in ("2019", "none") and rm is not None and uy_plan is None and y >= rm:
            break
        uy = uy_plan if is_plan else uy_act
        acts = []
        if uy is not None and y == uy:
            acts.append(("U", 0, "基幹的設備改良"))
            for i in range(len(prms)):
                if prms[i]["kind"] in ("O", "R"):
                    last[i] = y
        if e["tag"] in FIXED_CLASSES_SKIP:
            for (fy, k, note) in fixed:
                if fy == y and (not is_plan or y >= 2026 or cutoff is None):
                    acts.append((k, _fixed_cost(e, k), note))
            # 計画側の見込み（固定実績の周期を延長）
            if is_plan:
                yrs = [fy for fy, _, _ in fixed]
                k0 = fixed[0][1]
                if len(yrs) >= 2:
                    iv = yrs[-1] - yrs[-2]
                    nxt = [fy for fy in yrs if fy < (cutoff or 2027)]
                    if nxt:
                        base = nxt[-1]
                        if y > base and (y - base) % iv == 0 and not any(a[0] == k0 for a in acts):
                            acts.append((k0, _fixed_cost(e, k0), "計画"))
            # 固定以外の点検は通常どおり
            for i, p in enumerate(prms):
                if p["kind"] == "I" and _due(y, i, p, last, first_due):
                    acts.append(("I", p["cost"], ""))
                    last[i] = y
        else:
            for i, p in enumerate(prms):
                if not _due(y, i, p, last, first_due):
                    continue
                if p["kind"] == "R" and uy is not None and y <= uy <= y + 5:
                    continue  # 改良で更新するので単独更新は見送る
                if not is_plan and p["kind"] in ("O", "R") and e["imp"] in POSTPONE_YEARS.get(y, ()):
                    if _rng(e["tag"], f"pp{y}").random() < 0.6:
                        continue  # 送り（翌年に due のまま残る）
                acts.append((p["kind"], p["cost"], ""))
                last[i] = y
            for (fy, k, note) in fixed:
                if fy == y and (not is_plan):
                    acts.append((k, _fixed_cost(e, k), note))
                    for i, p in enumerate(prms):
                        if p["kind"] == k or (k == "R" and p["kind"] == "O"):
                            last[i] = y
        # 状態の発見 → 翌年度に補修（計画の版では知り得た分だけ）
        for fy in fnd:
            known = (cutoff is None and fy <= 2026) or (cutoff is not None and fy < cutoff)
            if known and y == fy + 1 and STRATEGY.get(e["cls"]):
                if not any(a[0] in ("O", "R", "U") for a in acts):
                    acts.append(("O", _fixed_cost(e, "O"), "前年度点検の指摘による補修"))
        if acts:
            res[y] = acts
    return res


def _due(y, i, p, last, first_due):
    if i not in last:
        return y >= first_due[i]
    return y - last[i] >= p["iv"]


def _fixed_cost(e, kind):
    for k, iv, cost in STRATEGY.get(e["cls"], []):
        if k == kind:
            return cost
    for k, iv, cost in STRATEGY.get(e["cls"], []):
        return cost
    return 300


def all_equipment():
    return EQ.build()


if __name__ == "__main__":
    eq = all_equipment()
    tot = {}
    for e in eq:
        s = simulate(e)
        for y, a in s.items():
            tot[y] = tot.get(y, 0) + len(a)
    print(sorted(tot.items()))
    for t in ("2-SH-304", "2-BF-412", "C-ST-501", "2-RH-515", "1-TE-305", "2-HV-243"):
        e = [x for x in eq if x["tag"] == t]
        if e:
            print(t, {y: [k for k, _, _ in a] for y, a in simulate(e[0]).items()})
