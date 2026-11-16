# -*- coding: utf-8 -*-
"""正解の記録（answer/files/<stem>.json と、資料の行 → 作業 の結び）"""
import os, json, datetime as dt
from common import ANS

FILES = {}
LINKS = []   # dict(file, sheet, r|seq, work)


def start(fname, plan=False):
    stem = os.path.splitext(fname)[0]
    FILES[stem] = dict(file=fname, plan=plan, records=[])
    return stem


def _clean(v):
    if isinstance(v, (dt.date, dt.datetime)):
        return v.isoformat()
    if isinstance(v, dict):
        return {k: _clean(x) for k, x in v.items() if x is not None and x != '' and x != {} and x != []}
    if isinstance(v, (list, tuple)):
        return [_clean(x) for x in v]
    return v


def add(stem, sheet, r, kind, fields, seq=None, works=None):
    rec = dict(sheet=sheet, kind=kind, fields=_clean(fields))
    if seq is not None:
        rec = dict(seq=seq, kind=kind, fields=_clean(fields))
        if sheet:
            rec['page'] = sheet
    else:
        rec['r'] = r
    FILES[stem]['records'].append(rec)
    for w in (works or []):
        if w:
            LINKS.append(dict(file=FILES[stem]['file'], sheet=sheet, **({'seq': seq} if seq is not None else {'r': r}), work=w))
    return rec


def next_seq(stem):
    return sum(1 for x in FILES[stem]['records'] if 'seq' in x)


def write_all():
    d = os.path.join(ANS, 'files')
    os.makedirs(d, exist_ok=True)
    for stem, v in FILES.items():
        with open(os.path.join(d, stem + '.json'), 'w', encoding='utf-8') as fp:
            json.dump(v, fp, ensure_ascii=False, indent=1)
