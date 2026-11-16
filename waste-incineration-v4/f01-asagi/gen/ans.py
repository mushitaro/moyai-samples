# -*- coding: utf-8 -*-
"""正解（answer/files/<stem>.json）の記録係。"""
import json, os, datetime as dt
BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FILES = BASE + '/files'
ADIR = BASE + '/answer/files'
A = {}      # stem -> dict
LINKS = []  # どの資料のどの行がどの作業か（world.json 用）

def _clean(v):
    if isinstance(v, (dt.date, dt.datetime)):
        return v.isoformat()
    if isinstance(v, dict):
        return {k: _clean(x) for k, x in v.items() if x is not None}
    if isinstance(v, (list, tuple)):
        return [_clean(x) for x in v]
    return v

def start(fname, plan=False):
    stem = os.path.splitext(os.path.basename(fname))[0]
    A[stem] = dict(file=os.path.basename(fname), plan=plan, records=[])
    return stem

def rec(stem, kind, fields, sheet=None, r=None, seq=None, image=None, works=None, code=None):
    d = dict(kind=kind, fields=_clean(fields))
    if sheet is not None:
        d['sheet'] = sheet
    if r is not None:
        d['r'] = r
    if seq is not None:
        d['seq'] = seq
    if image is not None:
        d['image'] = image
    A[stem]['records'].append(d)
    if works or code:
        LINKS.append(dict(file=A[stem]['file'], sheet=sheet, r=r, seq=seq, image=image, works=list(works or []), code=code))
    return d

def drop(stem):
    A.pop(stem, None)

def write_all():
    os.makedirs(ADIR, exist_ok=True)
    for stem, d in A.items():
        json.dump(d, open(os.path.join(ADIR, stem + '.json'), 'w'), ensure_ascii=False, indent=1)
