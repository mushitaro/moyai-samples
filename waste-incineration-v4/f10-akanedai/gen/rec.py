# -*- coding: utf-8 -*-
"""正解の行を集める。answer/files/<stem>.json に書く。"""
import json, os, datetime as dt
REG = {}
class Rec:
    def __init__(self, file, plan=False):
        self.file = file; self.plan = plan; self.records = []; REG[file] = self
    def add(self, sheet, r, kind, fields, works=None, seq=None):
        f = {k: (v.isoformat() if isinstance(v, (dt.date,)) else v) for k, v in fields.items() if v not in (None, '', [], {})}
        x = {'sheet': sheet, 'kind': kind, 'fields': f}
        if seq is not None: x['seq'] = seq
        else: x['r'] = r
        if works: x['_works'] = works
        self.records.append(x)
        return x
def dump(outdir):
    os.makedirs(outdir, exist_ok=True)
    for fn, rc in REG.items():
        stem = os.path.splitext(fn)[0]
        recs = [{k: v for k, v in x.items() if k != '_works'} for x in rc.records]
        json.dump({'file': fn, 'plan': rc.plan, 'records': recs}, open(os.path.join(outdir, stem + '.json'), 'w'), ensure_ascii=False, indent=0)
def links():
    out = []
    for fn, rc in REG.items():
        for x in rc.records:
            for w in x.get('_works', []):
                out.append({'file': fn, 'sheet': x['sheet'], 'r': x.get('r'), 'seq': x.get('seq'), 'work': w})
    return out
