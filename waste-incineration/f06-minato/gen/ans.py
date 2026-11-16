# -*- coding: utf-8 -*-
"""正解の記録（answer/files/<stem>.json と、world.json の資料行⇔作業の結び）"""
import json
import os

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ADIR = os.path.join(BASE, "answer", "files")

FILES = {}
LINKS = []      # {file, sheet, r|seq, work_ids, contract}
CATALOG = []    # {file, 資料, 契約, 年度, 様式, 作成者}


def start(fname, plan, 資料, fy, 様式, 作成者, contract=None):
    FILES[fname] = dict(file=fname, plan=plan, records=[])
    CATALOG.append(dict(file=fname, 資料=資料, 年度=fy, 様式=様式, 作成者=作成者, 契約=contract))


def rec(fname, sheet, r, kind, fields, works=None, check=None, seq=None):
    d = {"sheet": sheet, "kind": kind, "fields": fields}
    if seq is not None:
        d["seq"] = seq
    else:
        d["r"] = r
    if check:
        d["_check"] = [str(x) for x in check if x not in (None, "")]
    FILES[fname]["records"].append(d)
    if works:
        l = dict(file=fname, sheet=sheet, work_ids=list(works))
        if seq is not None:
            l["seq"] = seq
        else:
            l["r"] = r
        LINKS.append(l)


def save():
    os.makedirs(ADIR, exist_ok=True)
    for fn, d in FILES.items():
        stem = os.path.splitext(fn)[0]
        with open(os.path.join(ADIR, stem + ".json"), "w", encoding="utf-8") as f:
            json.dump(d, f, ensure_ascii=False, indent=0)
