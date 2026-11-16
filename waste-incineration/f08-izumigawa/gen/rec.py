# -*- coding: utf-8 -*-
"""正解の書き出し（answer/files/<stem>.json）と、資料の行 → 作業 の結び（world.json の links）。"""
import datetime as dt
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import ANS, FAC_K, OUT  # noqa: E402

LINKS = []
FILES = []


def jd(v):
    if isinstance(v, dt.date):
        return v.isoformat()
    if isinstance(v, dict):
        return {k: jd(x) for k, x in v.items()}
    if isinstance(v, (list, tuple)):
        return [jd(x) for x in v]
    return v


class Recs:
    def __init__(self, fname, plan=False):
        self.file = fname
        self.plan = plan
        self.records = []

    def add(self, sheet, r, kind, fields, wids=None, seq=False):
        rec = {"sheet": sheet, ("seq" if seq else "r"): r, "kind": kind, "fields": jd(fields)}
        if sheet is None:
            rec.pop("sheet")
        self.records.append(rec)
        if wids:
            LINKS.append({"file": self.file, "sheet": sheet, ("seq" if seq else "r"): r, "works": sorted(set(wids))})

    def save(self):
        os.makedirs(os.path.join(ANS, "files"), exist_ok=True)
        stem = os.path.splitext(self.file)[0]
        with open(os.path.join(ANS, "files", stem + ".json"), "w", encoding="utf-8") as f:
            json.dump({"file": self.file, "plan": self.plan, "records": self.records}, f, ensure_ascii=False, indent=1)
        FILES.append(self.file)


def short(contract_name):
    return contract_name.replace("いずみ川清掃センター", "").replace(" ", "").replace("　", "")


def fname(doc, what, form, ext):
    return f"{FAC_K}_{doc}_{what}_{form}.{ext}"


def path(fn):
    os.makedirs(OUT, exist_ok=True)
    return os.path.join(OUT, fn)
