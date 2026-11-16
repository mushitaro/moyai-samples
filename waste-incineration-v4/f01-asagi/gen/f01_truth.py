# -*- coding: utf-8 -*-
"""世界の正解（この施設で実際にあったこと）。帳票はここから書き分ける。
- EQUIP : 機器 1 台ずつ（台帳名、各資料での呼び名、号炉、設置年度）
- WORKS : 実際に行った作業（機器、INSPECT/OVERHAUL/REPLACE/UPGRADE、開始日・終了日、工数、業者、契約）
- REG   : どの資料のどの行がどの作業（または機器）か。各生成器が書くときに登録する
今回は書き出しはしない（dump() は用意だけ）。
"""
import datetime as dt
from f01_master import EQ, SYS, TODAY, wareki
from f01_history import BYTAG
from f01_contracts import (H23_F1, H23_F2, R7_F1, R7_F2, R7_CHANGES, CONTRACTS_H, MAKER, work_days)

TYPE = {'点検': 'INSPECT', '計測': 'INSPECT', '整備': 'OVERHAUL', '補修': 'OVERHAUL', '更新': 'REPLACE'}
KIKAN_TITLE = 'あさぎ市清掃工場 基幹的設備改良工事（平成29〜令和元年度）'

def _contract_name(fy, d, ct, text):
    if ct in ('K', 'X') and fy <= 2001:
        return 'あさぎ市清掃工場 ダイオキシン類恒久対策工事（平成12〜13年度）'
    if ct == 'D':
        return '平成24年度 あさぎ市清掃工場 分散制御システム更新工事'
    if ct == 'N':
        return '平成22年度 あさぎ市清掃工場 ごみピット放水銃設置工事'
    if ct == 'K':
        return KIKAN_TITLE
    if ct == 'H':
        for c in CONTRACTS_H.get(fy, []):
            if c['tag'] == d['tag']:
                return c['title']
        return '%s あさぎ市清掃工場%s%s補修工事' % (wareki(fy, False), '' if d['f'] == 0 else d['fname'], d['name'])
    return '%s あさぎ市清掃工場 定期整備業務委託' % wareki(fy, False)

def _vendor(fy, d, ct):
    if ct == 'H':
        for c in CONTRACTS_H.get(fy, []):
            if c['tag'] == d['tag']:
                return c['vendor']
        if d['prof'] in ('pit', 'stack_out'):
            return 'あさぎ建設株式会社'
        if d['prof'] == 'furnace':
            return '港東築炉工業株式会社'
    return MAKER

def _window(fy, d, ct):
    """見本の 2 年度は工程表のタスクから、その他の年度は炉別の定修期間（推定）から日付を決める。"""
    tasks = {2011: H23_F1 + H23_F2, 2025: R7_F1 + R7_F2}.get(fy)
    if tasks:
        for t in tasks:
            if d['tag'] in t['tags'] and (t['kind'] != 'sep' or ct == 'H'):
                e = R7_CHANGES.get(t['no'], {}).get('actual_e', t['e']) if fy == 2025 else t['e']
                return t['s'], e, t
    if ct == 'K':
        return {2000: (dt.date(2000, 9, 4), dt.date(2001, 3, 16)), 2001: (dt.date(2001, 9, 3), dt.date(2002, 2, 22)),
                2017: (dt.date(2017, 9, 4), dt.date(2018, 2, 23)), 2018: (dt.date(2018, 9, 3), dt.date(2019, 2, 22)),
                2019: (dt.date(2019, 6, 3), dt.date(2019, 12, 20))}[fy] + (None,)
    s = dt.date(fy, 6, 1) if d['f'] in (1, 0) else dt.date(fy, 10, 1)
    s += dt.timedelta(days=(7 - s.weekday()) % 7 + 7)   # 第 2 週の月曜から
    off = (sum(map(ord, d['tag'])) % 12)
    s2 = s + dt.timedelta(days=off)
    return s2, s2 + dt.timedelta(days=1 + off % 4), None

WORKS = []
_BY = {}
for d in EQ:
    for fy in sorted(d['act']):
        for i, (k, t, c, ct) in enumerate(d['act'][fy]):
            s, e, task = _window(fy, d, ct)
            if task and task['n'] != (0, 0):
                share = len([g for g in task['tags']]) or 1
                mh = round(sum(task['n']) * 8 * len(work_days(task)) / share, 1)
            else:
                mh = round(c * 0.077, 1)
            typ = 'UPGRADE' if ('インバータ' in t or '新設' in t) else TYPE[k]
            w = dict(id='W%d-%s-%d' % (fy, d['code'], i + 1), fy=fy, code=d['code'], tag=d['tag'], type=typ, kind_ja=k, text=t,
                     start=s, end=e, manhours=mh, vendor=_vendor(fy, d, ct), contract=_contract_name(fy, d, ct, t), contract_type=ct,
                     cost_k=c, status='完了' if e <= TODAY else '施工中')
            WORKS.append(w)
            _BY.setdefault((fy, d['tag']), []).append(w)
# 2026 の 2号炉（施工中、未計上）
for d in EQ:
    for i, (k, t, c) in enumerate(d.get('pending2026', [])):
        w = dict(id='W2026-%s-p%d' % (d['code'], i + 1), fy=2026, code=d['code'], tag=d['tag'], type=TYPE[k], kind_ja=k, text=t,
                 start=dt.date(2026, 10, 5), end=dt.date(2026, 11, 20), manhours=round(c * 0.077, 1), vendor=MAKER,
                 contract='令和8年度 あさぎ市清掃工場 定期整備業務委託', contract_type='T', cost_k=c, status='施工中')
        WORKS.append(w)
        _BY.setdefault((2026, d['tag']), []).append(w)

def work_ids(fy, tag, contract_type=None):
    return [w['id'] for w in _BY.get((fy, tag), []) if contract_type is None or w['contract_type'] == contract_type]

EQUIP = {d['code']: dict(code=d['code'], tag=d['tag'], ledger_name=d['name'], furnace=d['fname'], inst_fy=d['inst'], names={}) for d in EQ}

REG = []
_SEEN = set()
def ref(file, sheet, row, what, ids=(), code=None, name=None, col=None):
    """帳票の行（またはセル）と、正解の作業・機器を結びつける。name はその資料での呼び名。"""
    key = (file, sheet, row, col, what)
    if key in _SEEN:
        return
    _SEEN.add(key)
    REG.append(dict(file=file, sheet=sheet, row=row, col=col, what=what, works=list(ids), code=code))
    if code and name and code in EQUIP:
        EQUIP[code]['names'].setdefault(file, name)

def dump(path):
    import json
    def conv(o):
        return o.isoformat() if isinstance(o, (dt.date,)) else str(o)
    json.dump(dict(equip=EQUIP, works=WORKS, legal=LEGAL_WORKS, reg=REG), open(path, 'w'), ensure_ascii=False, indent=1, default=conv)

CUR = {'file': None}
def set_file(name):
    import os
    CUR['file'] = os.path.basename(name)
def r(sheet, row, what, ids=(), code=None, name=None, col=None):
    ref(CUR['file'], sheet, row, what, ids, code, name, col)

# 法令に基づく点検・検査・測定（年度ごとの実施。見本の契約・計画と結ぶ id は LGxx-年度）
from f01_master import LEGAL, legal_years
LEGAL_WORKS = []
for it in LEGAL:
    for y in legal_years(it, 2009, 2026):
        LEGAL_WORKS.append(dict(id='%s-%d' % (it[0], y), fy=y, name=it[1], basis=it[2], targets=it[3], times=it[5],
                                type='INSPECT', status='完了' if y <= 2025 else '一部実施（上期）'))
