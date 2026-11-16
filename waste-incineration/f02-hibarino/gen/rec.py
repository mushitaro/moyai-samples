# -*- coding: utf-8 -*-
"""正解（answer/files/<stem>.json）と、資料の行 → 作業 の結び（world.json の links）を集める。"""
import json
import os

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FILES = os.path.join(BASE, "files")
ANS = os.path.join(BASE, "answer", "files")
REG = {}
LINKS = []


class Rec:
    def __init__(self, fname, plan=False):
        self.file = fname
        self.plan = plan
        self.records = []
        self.seq = 0
        REG[fname] = self

    def add(self, sheet, r, kind, fields, works=None):
        d = {"sheet": sheet, "r": r, "kind": kind, "fields": fields}
        self.records.append(d)
        for w in works or []:
            LINKS.append({"file": self.file, "sheet": sheet, "r": r, "work": w})
        return d

    def add_seq(self, kind, fields, works=None):
        d = {"seq": self.seq, "kind": kind, "fields": fields}
        self.records.append(d)
        for w in works or []:
            LINKS.append({"file": self.file, "seq": self.seq, "work": w})
        self.seq += 1
        return d


def path(name):
    return os.path.join(FILES, name)


def dump():
    os.makedirs(ANS, exist_ok=True)
    for f, r in REG.items():
        stem = os.path.splitext(f)[0]
        with open(os.path.join(ANS, stem + ".json"), "w", encoding="utf-8") as fp:
            json.dump({"file": f, "plan": r.plan, "records": r.records}, fp, ensure_ascii=False, indent=1, default=str)
