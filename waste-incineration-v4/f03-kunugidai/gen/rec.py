# -*- coding: utf-8 -*-
"""正解の行（answer/files/<stem>.json）を、帳票を書くのと同時に記録する。"""
import os, json
import xlhelp as X

ANS = os.path.join(X.BASE, "answer", "files")
LINKS = []   # (file, sheet, r/seq, work id)


class Rec:
    def __init__(self, file, plan=False):
        self.file = file
        self.plan = plan
        self.records = []
        self.seq = 0

    def add(self, sheet, r, kind, fields, works=None, **extra):
        f = {k: v for k, v in fields.items() if v not in (None, "", [], {})}
        rec = {"sheet": sheet, "r": r, "kind": kind, "fields": f}
        rec.update(extra)
        self.records.append(rec)
        for w in works or []:
            if w:
                LINKS.append(dict(file=self.file, sheet=sheet, r=r, work=w))
        return rec

    def addp(self, kind, fields, works=None, **extra):
        """PDF：seq を振る（見出し・題・ページの外は呼ばない）"""
        f = {k: v for k, v in fields.items() if v not in (None, "", [], {})}
        rec = {"seq": self.seq, "kind": kind, "fields": f}
        rec.update(extra)
        self.records.append(rec)
        for w in works or []:
            if w:
                LINKS.append(dict(file=self.file, seq=self.seq, work=w))
        self.seq += 1
        return rec

    def save(self):
        os.makedirs(ANS, exist_ok=True)
        stem = os.path.splitext(self.file)[0]
        with open(os.path.join(ANS, stem + ".json"), "w", encoding="utf-8") as fh:
            json.dump({"file": self.file, "plan": self.plan, "records": self.records}, fh, ensure_ascii=False, indent=0, default=str)
