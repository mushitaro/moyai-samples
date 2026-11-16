# -*- coding: utf-8 -*-
"""SPC 支給品（SPC が会計システムで発注し、定修工事に支給する部品）。発注台帳と報告書の「使用部品」で同じものを使う。"""
from datetime import timedelta, date
from world import BY_CODE
from contracts_full import get

# (作業名の一致, 品名, 仕様, 数量, 単位, 単価(円), 使用数, 取引先, 納期の前倒し日数)
RULES = [
    ("燃焼火格子更新", "{u}号炉 燃焼火格子（可動・固定）", "MZ-G150-C 高クロム鋳鋼 192枚＋予備8枚", 200, "枚", 307200, 192, "V0101", 150),
    ("火格子部分交換", "{u}号炉 {stage}火格子（部分交換用）", "MZ-G150-C 摩耗の大きい列", 48, "枚", 296000, 48, "V0101", 240),
    ("ろ布全数交換", "{u}号炉 バグフィルタ ろ布", "PTFE φ165×6,000 720本＋予備40本", 760, "本", 26800, 720, "V0101", 120),
    ("予備層触媒充填", "{u}号炉 脱硝触媒（予備層 1層分）", "ハニカム 低温活性型 24ユニット", 24, "ユニット", 1480000, 24, "V0101", 180),
    ("分析計更新", "{u}号炉 {name}", "【架空】青嵐計器 後継機 一式（据付は定修工事）", 1, "台", 0, 1, "V0330", 100),
    ("ポンプ分解整備", "{name} 分解整備部品", "メカシール2組（1組予備）・軸受・ライナリング", 1, "式", 1380000, 1, "V0210", 90),
    ("蓄電池更新", "直流電源装置 蓄電池", "MSE 600Ah 54セル", 54, "セル", 98000, 54, "V0520", 80),
    ("UPS蓄電池更新", "{name} 蓄電池", "小形制御弁式", 1, "式", 1900000, 1, "V0330", 60),
    ("減速機分解整備", "空冷復水器ファン減速機 軸受一式", "1台分", 1, "式", 742000, 1, "V0101", 75),
]
PRICE = {"HCl計": 8200000, "SOx・NOx・O2・CO計": 9800000, "ばいじん計": 6400000, "水銀連続分析計": 8650000, "NH3計": 7200000}


def supplies(c):
    """契約 c の支給品：[dict(name, spec, qty, unit, price, used, spare, vendor, order_date, due, code, work)]"""
    out = []
    S, E = c["period"]
    for t in c["tasks"]:
        for (code, tt, w) in t["works"]:
            e = BY_CODE[code]
            for (key, nm, spec, q, un, up, used, v, lead) in RULES:
                if key != w:
                    continue
                if key == "ポンプ分解整備" and e["base"] != "ボイラ給水ポンプ":
                    continue
                if key == "減速機分解整備" and e["base"] != "空冷復水器ファン減速機":
                    continue
                stage = e["base"].replace("火格子", "")
                price = PRICE.get(e["base"], up)
                if key == "火格子部分交換" and e["base"] == "後燃焼火格子":
                    q, used = 36, 36
                name = nm.format(u=e["unit"], stage=stage, name=e["name"])
                od = t["ps"] - timedelta(days=lead)
                due = t["ps"] - timedelta(days=12)
                out.append(dict(name=name, spec=spec, qty=q, unit=un, price=price, used=used, spare=q - used, vendor=v,
                                order_date=od, due=due, code=code, work=w, contract=c["key"]))
    # 同じ品名の減速機軸受はまとめる
    merged = {}
    for s in out:
        k = s["name"]
        if k in merged:
            merged[k]["qty"] += s["qty"]
            merged[k]["used"] += s["used"]
            merged[k]["codes"].append(s["code"])
        else:
            s["codes"] = [s["code"]]
            merged[k] = s
    return list(merged.values())


ALL_SUPPLIES = None


def all_supplies():
    global ALL_SUPPLIES
    if ALL_SUPPLIES is None:
        ALL_SUPPLIES = []
        for c in get():
            ALL_SUPPLIES += supplies(c)
    return ALL_SUPPLIES


if __name__ == "__main__":
    for s in all_supplies():
        print(s["contract"], s["name"], s["qty"], s["order_date"])
