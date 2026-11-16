# -*- coding: utf-8 -*-
"""answer/ を物差しの形（moyai.sample-world/0）に書き直す。files/ は読むだけ（星取表の網掛け・コメントを読む）。
元にするのは work/answer_v1/（v1 の正解の写し）と、生成器の正解データ（works.py をメモリ上で組み立てる。ファイルは書かない）。
使い方：python3 to_v2.py"""
import os, sys, json, glob, re, hashlib, datetime as dt
from openpyxl import load_workbook

HERE = os.path.dirname(os.path.abspath(__file__))
FULL = os.path.dirname(HERE)
FILES = os.path.join(FULL, 'files')
ANS = os.path.join(FULL, 'answer')
V1 = os.path.join(FULL, 'gen', '_answer_v1')
sys.path.insert(0, HERE)
import works as W   # メモリ上だけ

GREY = ('FFBFBFBF', '00BFBFBF')
Y2SYM = {'○': '点', '◎': '整', '●': '更', '△': '補'}


def hash_files():
    return {os.path.basename(p): hashlib.sha256(open(p, 'rb').read()).hexdigest() for p in sorted(glob.glob(os.path.join(FILES, '*')))}


def wd_count(spans):
    n = 0
    for s, e in spans:
        d0, d1 = dt.date.fromisoformat(s), dt.date.fromisoformat(e)
        n += len(W.workdays(d0, d1))
    return n


# ------------------------------------------------------------------ 作業（実績＋予定）
v1w = json.load(open(os.path.join(V1, 'world.json'), encoding='utf-8'))
works = []
WID = {}
for x in v1w['works']:
    days = sum(len(W.workdays(dt.date.fromisoformat(s), dt.date.fromisoformat(e))) for s, e in [(x['start'], x['end'])]) if x['contract'] != 'KIKAN' else None
    wo = W.WBY[x['id']]
    days = sum(len(W.workdays(s, e)) for s, e in wo['act']) or 1
    status = 'done'
    md = x['manDays']
    if x['contract'].startswith('T'):
        md = round(md * 7.5 / 8, 3)    # 定期整備は 8:30～17:00（実働 7.5h）。1 人日＝8h
    if not md:
        md = None                       # 別契約は人数の記録なし
    works.append(dict(id=x['id'], asset=(x['machines'][0] if x['machines'] else None), assets=x['machines'], fy=x['年度'], status=status,
                      task=x['作業'], attr=x['type'], start=x['start'], end=x['end'], manDays=md, days=days,
                      contract=x['contract'], vendor=x['vendor'], unrecorded=x.get('unrecorded', False)))
    WID[x['id']] = works[-1]
done_by = {}
for w in works:
    for a in w['assets']:
        done_by.setdefault((a, w['fy']), []).append(w)

PLANNED = {}


def planned_work(asset, fy, sym, status, name):
    key = (asset, fy, sym, status)
    if key in PLANNED:
        return PLANNED[key]['id']
    wid = f'P-{asset or "LEGAL"}-{fy}-{len(PLANNED)+1:04d}'
    attr = {'点': 'INSPECT', '○': 'INSPECT', '整': 'OVERHAUL', '◎': 'OVERHAUL', '補': 'OVERHAUL', '△': 'OVERHAUL', '更': 'REPLACE', '●': 'REPLACE', '★': 'UPGRADE'}.get(sym[:1], 'INSPECT')
    w = dict(id=wid, asset=asset, assets=[asset] if asset else [], fy=fy, status=status, task=f'{name} {sym}（計画）', attr=attr,
             start=None, end=None, manDays=None, days=None, contract='PLAN')
    PLANNED[key] = w
    works.append(w)
    return wid


def match_done(asset, fy, sym):
    lst = done_by.get((asset, fy), [])
    if not lst:
        return None
    s = sym[:1]
    want = {'点': None, '○': None, '整': {'OVERHAUL', 'REPLACE', 'UPGRADE'}, '◎': {'OVERHAUL', 'REPLACE', 'UPGRADE'}, '補': {'OVERHAUL', 'REPLACE', 'UPGRADE'},
            '△': {'OVERHAUL', 'REPLACE', 'UPGRADE'}, '更': {'REPLACE', 'UPGRADE'}, '●': {'REPLACE', 'UPGRADE'}, '★': {'UPGRADE'}}.get(s)
    for w in lst:
        if want is None or w['attr'] in want:
            return w['id']
    return lst[0]['id']


# ------------------------------------------------------------------ 資料ごとの記録
links = []
errata = []
names = {}   # asset -> {file: path}
os.makedirs(os.path.join(ANS, 'files'), exist_ok=True)
for p in glob.glob(os.path.join(ANS, 'files', '*.json')):
    os.remove(p)
v1links = {}
for l in v1w['links']:
    v1links.setdefault((l['file'], l.get('sheet'), l.get('r'), l.get('seq')), []).append(l['work'])

for jp in sorted(glob.glob(os.path.join(V1, 'files', '*.json'))):
    a = json.load(open(jp, encoding='utf-8'))
    fname = a['file']
    fp = os.path.join(FILES, fname)
    out = dict(file=fname, plan=a['plan'], records=[])
    is_pdf = fname.endswith('.pdf')
    wb = load_workbook(fp) if fname.endswith('.xlsx') else None
    y2 = '_y2様式' in fname
    V = None
    if a['plan']:
        m = re.search(r'令和(\d)年度', fname)
        V = int(m.group(1)) + 2018
    for rc in a['records']:
        f = dict(rc['fields'])
        nr = dict(kind=rc['kind'], fields=f)
        if is_pdf:
            nr['sheet'] = None
            nr['r'] = None
            nr['seq'] = rc['seq']
            if 'image' in f:
                nr['image'] = f.pop('image')
        else:
            nr['sheet'] = rc['sheet']; nr['r'] = rc['r']
        key = (fname, None if is_pdf else rc.get('sheet'), None if is_pdf else rc.get('r'), rc.get('seq'))
        wl = v1links.get(key, [])
        if a['plan']:
            # 星取表：marks＝実績（y1「済」・計画外、y2 網掛け）、plans＝予定（当初の記号）
            ws = wb[rc['sheet']]
            hdr_row = 6 if not y2 else 7
            ycol = {}
            for c in range(1, ws.max_column + 1):
                v = ws.cell(hdr_row, c).value
                if isinstance(v, str) and re.fullmatch(r'R\d+', v):
                    ycol[c] = int(v[1:]) + 2018
            marks, plans = {}, {}
            asset = f.get('code')
            comments = f.get('comments') or {}
            for c, y in ycol.items():
                cell = ws.cell(rc['r'], c)
                v = cell.value
                if not isinstance(v, str) or not v.strip() or len(v) > 8:
                    continue
                addr = cell.coordinate
                cmt = comments.get(addr, '')
                if not y2:
                    unplanned = '計画外' in cmt
                    base = v.replace('済', '')
                    if v.endswith('済'):
                        marks[str(y)] = v
                    if not unplanned:
                        plans[str(y)] = base
                else:
                    shaded = (cell.fill is not None and cell.fill.fgColor is not None and str(cell.fill.fgColor.rgb) in GREY)
                    unplanned = '計画外' in cmt
                    if shaded:
                        marks[str(y)] = v
                    if not unplanned:
                        plans[str(y)] = v
                # 結び
                if rc['kind'] == 'item':
                    nm = f['path'][-1]
                    if asset or f['path'][0].startswith('法定'):
                        if str(y) in marks:
                            wid = match_done(asset, y, v) if asset else None
                            if wid is None:
                                wid = planned_work(asset, y, v.replace('済', ''), 'done_unlinked', nm)
                                PLANNED[(asset, y, v.replace('済', ''), 'done_unlinked')]['status'] = 'done'
                            links.append(dict(file=fname, sheet=rc['sheet'], r=rc['r'], work=wid, year=y, how='direct'))
                        elif str(y) in plans:
                            past = y < V if not y2 else y <= V
                            if past:
                                st = 'postponed' if any(k in cmt for k in ('繰越', '先送り', '翌年度', '見直し')) else 'not_done'
                            else:
                                st = 'planned'
                            wid = planned_work(asset, y, v.replace('済', ''), st, nm)
                            links.append(dict(file=fname, sheet=rc['sheet'], r=rc['r'], work=wid, year=y, how='direct'))
            if rc['kind'] in ('item', 'parent'):
                f['marks'] = marks
                f['plans'] = plans
                if 'path' not in f:
                    f['path'] = [f.get('WBS', '')]
            if asset and rc['kind'] == 'item':
                names.setdefault(asset, {})[fname] = ' > '.join(f['path'])
            # 精密機能検査 R18 の印の抜け
            if y2 and f.get('path', [''])[-1] == '精密機能検査':
                wid = planned_work(None, 2036, '点', 'planned', '精密機能検査')
                PLANNED[(None, 2036, '点', 'planned')]['task'] = '精密機能検査 点（3年周期、計画から抜けた）'
                errata.append(dict(type='omitMark', file=fname, sheet=rc['sheet'], r=rc['r'], year=2036, works=[wid],
                                   note='3年周期の R18（2036）の印が無い（R6 改訂でコピー時に欠落、R7 版に継承）'))
        else:
            # 工数：人日（1 人日＝8 時間）
            if 'manDays' in f:
                if '時間' in f and f.get('人員'):
                    h = 0
                    t = f['時間']
                    a1, b1 = t.split('～')
                    h1, m1 = map(int, a1.split(':')); h2, m2 = map(int, b1.split(':'))
                    h = (h2 * 60 + m2 - h1 * 60 - m1) / 60
                    h = h - 1 if h > 6 else h
                    f['manDays'] = round(f['人員'] * h / 8, 3)
                    f['manDaysText'] = f"{f['人員']}人 {t}"
                    f['days'] = 1
                elif '工数' in f:
                    f['manDays'] = round(float(f['工数']) / 8, 3)
                    f['manDaysText'] = f.get('manDaysText') or f'{f["工数"]}'
                    if f.get('日付'):
                        f['days'] = 1
                elif 'WBS' in f and f.get('人員'):
                    f['manDays'] = float(f['人員'])   # 作業時間 8:00～17:00（昼 1h 除き 8h）
                    f['manDaysText'] = f'{f["人員"]}人（8:00～17:00）'
                    f['days'] = 1
                elif '作業員計' in f:
                    f['manDays'] = float(f['作業員計'])
                    f['days'] = 1
            if 'plan' in f or 'actual' in f:
                sp = f.get('actual') or f.get('plan')
                if sp and all(isinstance(x, list) and len(x) == 2 for x in sp):
                    f['days'] = wd_count(sp)
            if '日' in f and isinstance(f['日'], int):
                f['days'] = f['日']
            for wid in wl:
                ln = dict(file=fname, sheet=None if is_pdf else rc['sheet'], r=None if is_pdf else rc['r'], work=wid)
                if is_pdf:
                    ln['seq'] = rc['seq']
                if f.get('日付') and re.fullmatch(r'\d{4}-\d\d-\d\d', str(f['日付'])):
                    ln['day'] = f['日付']
                w = WID.get(wid)
                how = 'direct'
                if '発注台帳' in fname and len(wl) > 1:
                    how = 'upgrade' if w and w['contract'] == 'KIKAN' else 'annual'
                elif '年間工程' == rc.get('sheet') and len(wl) == 1 and w and w['contract'].startswith('T'):
                    how = 'rollup'
                ln['how'] = how
                links.append(ln)
            # 機器台帳の名前
            if fname.endswith('.xlsx') and '機器台帳' in fname and f.get('code'):
                nm = f.get('機器名称')
                names.setdefault(f['code'], {})[fname] = ' > '.join(x for x in [f.get('区分（値）') or f.get('区分'), nm] if x) if nm else None
        out['records'].append(nr)
    stem = os.path.splitext(fname)[0]
    json.dump(out, open(os.path.join(ANS, 'files', stem + '.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=1)

# 機器台帳刷新の名前（機器名称）・旧台帳（機器名称）
for jp in glob.glob(os.path.join(ANS, 'files', '*機器台帳*.json')):
    a = json.load(open(jp, encoding='utf-8'))
    for rc in a['records']:
        f = rc['fields']
        if f.get('code') and f.get('機器名称'):
            names.setdefault(f['code'], {})[a['file']] = ' > '.join(x for x in [rc['sheet'] if '刷新' in a['file'] else (f.get('区分（値）') or f.get('区分')), f['機器名称']] if x)

# ------------------------------------------------------------------ assets
assets = []
oldn = 0
for m in v1w['machines']:
    if m['code']:
        nm = dict(names.get(m['code'], {}))
        nm['工程表・日報'] = m['呼び名'].get('工程表・日報')
        assets.append(dict(id=m['code'], name=m['台帳名'], tag=m['code'], system=m['系統'], furnace={'1号炉': '1', '2号炉': '2', '共通': '共通'}[m['号炉']],
                           installed=int(m['設置'][:4]), class_=m.get('重要度_y2'), part=None, names=nm,
                           oldName=m['旧台帳名'], standby=m.get('常用予備'), status=m['状態'], renewedFY=m['更新年度']))
    else:
        oldn += 1
        assets.append(dict(id=f'OLD-S-{oldn:02d}', name=m['旧台帳名'], tag=None, system=m['系統'], furnace='共通', installed=2001, class_=None, part=None,
                           names={'機器台帳R3旧版': m['旧台帳名']}, status=m['状態']))
for a in assets:
    a['class'] = a.pop('class_')

files_list = [dict(file=os.path.basename(p)) for p in sorted(glob.glob(os.path.join(FILES, '*')))]
v1p = v1w['profile']
world = {
    'schema': 'moyai.sample-world/0',
    'facility': dict(id='F05', name=v1p['施設']),
    'profile': dict(furnace=2, **{k: v for k, v in v1p.items() if k != 'id'}),
    'maker': dict(today='2026-11-13', seed=v1p['乱数の種']),
    'assets': assets,
    'works': works,
    'links': links,
    'errata': errata,
    'files': files_list,
}
json.dump(world, open(os.path.join(ANS, 'world.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
extra = dict(note='v1 の world.json のうち v2 の形に無い欄', contracts=v1w['contracts'], unrecordedWorks=v1w['unrecordedWorks'],
             planNotDone=v1w['planNotDone'], planGaps=v1w['planGaps'], machines_v1=v1w['machines'], profile_v1=v1p,
             works_v1_extra={x['id']: {k: x.get(k) for k in ('契約件名', '停止', 'machines', 'plan', '発見', '部品', '追加')} for x in v1w['works']})
json.dump(extra, open(os.path.join(ANS, 'world.extra.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
print('assets', len(assets), 'works', len(works), 'planned', len(PLANNED), 'links', len(links), 'errata', len(errata))
