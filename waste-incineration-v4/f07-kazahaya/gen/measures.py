# -*- coding: utf-8 -*-
"""測定値の正解（報告書・計画の備考・台帳メモで同じ値を使う）。"""
import zlib
from datetime import date


def _k(*a):
    return zlib.crc32("|".join(map(str, a)).encode()) % 1000 / 1000


def grate_wear(u, stage, y):
    """各年度定修時の最大摩耗量 mm（新品基準面から）。1号は 9〜10 月、2号は 5〜6 月に測る。"""
    t_meas = y + (0.75 if u == "1" else 0.4) - 2015.25     # 竣工 2015.3 からの年数
    if stage == "燃焼":
        repl = 2023.8 if u == "1" else 2024.45
        part = 2019.7 if u == "1" else 2020.4
        if y + (0.75 if u == "1" else 0.4) > repl:   # 更新後
            return round(2.4 * (y + (0.75 if u == "1" else 0.4) - repl) + 0.0, 1) or None
        hot = 2.46 * t_meas                                # 1・3 列目
        rest = 1.85 * t_meas
        if y + (0.75 if u == "1" else 0.4) > part + 0.05:
            hot = 2.5 * (y + (0.75 if u == "1" else 0.4) - part)
        v = max(hot, rest) * (1.0 if u == "1" else 0.94)
        return round(v, 1)
    if stage == "乾燥":
        v = 1.0 * t_meas if t_meas < 5 else 5.0 + 2.2 * (t_meas - 5)
        if u == "1" and y == 2018:
            return 3.9
        if u == "1" and y == 2023:
            return 9.6
        if u == "1" and y == 2024:
            return 11.8
        return round(v * (1.0 if u == "1" else 0.95), 1)
    if stage == "後燃焼":
        v = 1.45 * t_meas
        if (u == "1" and y > 2025) or (u == "2" and y > 2026):
            v = 1.45 * 1
        return round(v * (1.0 if u == "1" else 0.96), 1)
    return None


SH3 = {"1": {2017: 6.9, 2019: 6.6, 2021: 6.3, 2023: 5.6, 2025: 5.1},
       "2": {2016: 7.0, 2018: 6.8, 2020: 6.5, 2022: 6.1, 2024: 5.7, 2026: 5.4}}


BOILER = {  # 部位: (公称, 限界, 1号炉 2023 年の最小, 減肉 mm/年)  ※2023 1号炉 報告書（見本）の値に合わせる
    "水管壁（炉頂・放射室）": ("5.0+肉盛", 4.0, 4.9, 0.15), "蒸発管群": ("4.5", 3.5, 4.3, 0.0), "1次過熱器": ("5.0", 4.5, 4.9, 0.0),
    "2次過熱器": ("6.0", 4.5, 5.8, 0.05), "エコノマイザ": ("4.0", 3.0, 3.9, 0.0), "蒸気ドラム": ("32.0", 28.0, 31.8, 0.0)}


def boiler_min(u, part, y):
    """肉厚測定の最小値 mm、前回値、公称、限界。"""
    if part == "3次過熱器":
        return SH3[u].get(y), SH3[u].get(y - 2), "7.1", 4.5
    nom, lim, v23, rate = BOILER[part]
    off = 0.0 if u == "1" else 0.1
    f = lambda yy: round(v23 - rate * (yy - 2023) + off * (1 if rate else 0), 1)
    if part == "水管壁（炉頂・放射室）" and ((u == "1" and y > 2025) or (u == "2" and y > 2026)):
        f = lambda yy: 5.6   # 2025 肉盛補修後
    return f(y), (f(y - 2) if y - 2 >= 2015 else None), nom, lim


def boiler_judge(part, v, lim):
    rate = 0.35 if part == "3次過熱器" else BOILER[part][3]
    return "要観察" if (rate >= 0.1 and v - lim < 1.5) else "良"


IDF = {("1", 2019): (4.2, 1.8), ("1", 2021): (4.6, None), ("1", 2023): (5.8, 2.1), ("2", 2020): (4.0, 1.7), ("2", 2024): (5.2, 2.0),
       ("1", 2025): (2.4, None), ("2", 2026): (2.6, None)}


def vib(code, y):
    if code[1:] == "KT-100":
        return IDF.get((code[0], y), (3.0 + _k(code, y), 1.6 + _k(code, "a", y)))
    b = round(3.0 + 2.5 * _k(code, y), 1)
    a = round(1.2 + 0.8 * _k(code, "a", y), 1)
    return b, a


BF = {("1", 2018): (88, 92), ("1", 2020): (74, 79), ("1", 2023): (81, 86), ("1", 2025): (78, 83),
      ("2", 2019): (87, 91), ("2", 2021): (73, 77), ("2", 2024): (82, 86), ("2", 2026): (79, 84)}
CAT = {("1", 2017): (94, None), ("1", 2020): (93, 91), ("1", 2023): (92, 89), ("2", 2018): (94, None), ("2", 2021): (92, 90),
       ("2", 2024): (91, 88)}


def bf_strength(u, y):
    return BF.get((u, y), (85, 90))


def cat_rate(u, y):
    return CAT.get((u, y), (92, 90))


def pi_index(code, y):
    return round(2.6 + 2.0 * _k(code, y), 1), round(0.8 + 1.6 * _k(code, "t", y), 1)
