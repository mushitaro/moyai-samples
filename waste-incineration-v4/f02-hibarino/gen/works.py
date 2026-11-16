# -*- coding: utf-8 -*-
"""実際に行った作業（年度単位の正解）に ID を振る。日付・工数・業者・契約は工程・日報から後で埋める。"""
import world as W

ATTR = {"X": "REPLACE", "U": "UPGRADE", "O": "OVERHAUL", "R": "OVERHAUL", "I": "INSPECT"}
PRI = "XUORI"


def key_of(a):
    return next((k for k in PRI if k in a), None)


def wid(e, y):
    a = W.actual_work(e, y)
    if not a:
        return None
    return f"W{y}-{e.no}"


WORKS = {}   # id -> dict


def build_all():
    for e in W.EQ:
        for y in range(2016, 2027):
            a = W.actual_work(e, y)
            if not a:
                continue
            city = W.CITY_WORKS.get((e.name, e.f, y))
            if city:
                WORKS[f"W{y}-{e.no}-CITY"] = {"id": f"W{y}-{e.no}-CITY", "equipment": e.no, "name": e.full, "furnace": e.f, "fy": y,
                                              "attr": "OVERHAUL", "detail": "R", "repair": True, "start": None, "end": None,
                                              "manDays": None, "vendor": "鷹羽エンジニアリング株式会社", "contract": city}
                a = a - {"R", "X"}
            k = key_of(a)
            WORKS[wid(e, y)] = {"id": wid(e, y), "equipment": e.no, "name": e.full, "furnace": e.f, "fy": y,
                                "attr": ATTR[k], "detail": "".join(sorted(a)), "repair": k == "R",
                                "start": None, "end": None, "manDays": None, "vendor": None, "contract": None}
    return WORKS
