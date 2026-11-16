# -*- coding: utf-8 -*-
"""正解の記録（answer/files/<stem>.json）と、資料の行→作業の結び"""
import os, json
HERE = os.path.dirname(os.path.abspath(__file__))
ANS = os.path.normpath(os.path.join(HERE, "..", "answer", "files"))
LINKS = []          # (file, sheet, r/seq, work_id)

class Rec:
    def __init__(self, file, plan=False):
        self.file, self.plan, self.records, self.seq = file, plan, [], 0
    def add(self, kind, fields, sheet=None, r=None, work=None, seq=False):
        rec = {"kind": kind, "fields": {k: v for k, v in fields.items() if v not in (None, "", {}, [])}}
        if seq:
            rec = {"seq": self.seq, **rec}; self.seq += 1
        else:
            rec = {"sheet": sheet, "r": r, **rec} if sheet is not None else {"r": r, **rec}
        self.records.append(rec)
        ws = work if isinstance(work, (list, tuple)) else ([work] if work else [])
        for w in ws:
            LINKS.append({"file": self.file, "sheet": sheet, ("seq" if seq else "r"): rec.get("seq", r), "work": w})
        return rec
    def save(self):
        os.makedirs(ANS, exist_ok=True)
        stem = os.path.splitext(self.file)[0]
        with open(os.path.join(ANS, stem + ".json"), "w", encoding="utf-8") as fp:
            json.dump({"file": self.file, "plan": self.plan, "records": self.records}, fp, ensure_ascii=False, indent=0)
