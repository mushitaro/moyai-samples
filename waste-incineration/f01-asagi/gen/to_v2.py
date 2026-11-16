# -*- coding: utf-8 -*-
"""answer/ を物差しの形（moyai.sample-world/0）にそろえる変換器。files/ は読まず・書かない。
入力：answer_v1（今までの answer/files/*.json と answer/world.json）。出力：answer/files/*.json・answer/world.json・answer/world.extra.json
python3 to_v2.py [v1 の answer のパス]
"""
import json, os, sys, glob, datetime as dt, re, shutil

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = sys.argv[1] if len(sys.argv) > 1 else BASE + '/gen/_answer_v1'
OUT = BASE + '/answer'
TODAY = dt.date(2026, 11, 13)

def jload(p):
    return json.load(open(p, encoding='utf-8'))

W1 = jload(SRC + '/world.json')
A1 = {os.path.splitext(os.path.basename(p))[0]: jload(p) for p in sorted(glob.glob(SRC + '/files/*.json'))}

def fy_of(d):
    d = dt.date.fromisoformat(d)
    return d.year if d.month >= 4 else d.year - 1

def disp(v):
    """星取表・計画表のセルに見えている字（数は 3 桁区切り）"""
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, int):
        return '{:,}'.format(v)
    if isinstance(v, float):
        return '{:,}'.format(int(v)) if v == int(v) else str(v)
    return str(v)

def path_list(p):
    return [x for x in str(p).split('/') if x]

# ------------------------------------------------------------------ 機器
eq = W1['equipment']
FUR = {'1号炉': '1', '2号炉': '2', '共通': '共通'}
assets = []
for code, e in eq.items():
    inst = e['inst_fy']
    assets.append(dict(id=code, name=e['ledger_name'], tag=e['tag'], system=e['system'], furnace=FUR.get(e['furnace'], e['furnace']),
                       installed=1997 if inst == 1996 else inst, class_=e['importance'], part=None, names={}))
for a in assets:
    a['class'] = a.pop('class_')
AS = {a['id']: a for a in assets}

# ------------------------------------------------------------------ 作業（実績）
works = []
WBY = {}      # (code, fy) -> [work id]（実績）
for w in W1['works']:
    status = 'done' if w['status'] == '完了' else 'in_progress'
    s, e = w.get('start'), w.get('end')
    days = (dt.date.fromisoformat(e) - dt.date.fromisoformat(s)).days + 1 if s and e else None
    works.append(dict(id=w['id'], asset=w['code'], fy=w['fy'], status=status, task=w['text'], attr=w['type'], start=s, end=e,
                      manDays=w['manDays'], days=days, vendor=w['vendor'], contract=w['contract'], contract_type=w['contract_type']))
    if w['contract_type'] != 'R':
        WBY.setdefault((w['code'], w['fy']), []).append(w['id'])
for L in W1['legal_inspections']:
    tg = L['targets'][0] if L['targets'] else None
    code = next((c for c, e in eq.items() if e['tag'] == tg), None) if tg else None
    works.append(dict(id=L['id'], asset=code, fy=L['fy'], status='done' if L['fy'] <= 2025 else 'in_progress', task=L['name'], attr='INSPECT',
                      start=None, end=None, manDays=None, days=None, basis=L['basis'], times_per_year=L['times_per_year']))
WID = {w['id']: w for w in works}

# ------------------------------------------------------------------ 予定の作業（星取表・計画表の予定の印から）
def plan_work(code, year, task):
    """予定の印に当たる作業。実績があればそれ、無ければ予定（planned）か、過去なら実施されなかった予定（not_done）"""
    done = WBY.get((code, year))
    if done:
        return done
    wid = 'PL-%s-%d' % (code or 'LEGAL', year)
    if wid not in WID:
        st = 'planned' if year >= 2026 else 'not_done'
        WID[wid] = dict(id=wid, asset=code, fy=year, status=st, task=task, attr='INSPECT' if task in ('○', '◇') else ('REPLACE' if task == '●' else 'OVERHAUL'),
                        start=None, end=None, manDays=None, days=None)
        works.append(WID[wid])
    return [wid]

# ------------------------------------------------------------------ 資料ごとの正解
links = []
os.makedirs(OUT + '/files', exist_ok=True)
for f in glob.glob(OUT + '/files/*.json'):
    os.remove(f)
v1_links = {}
for L in W1['doc_links']:
    key = (L['file'], L.get('sheet'), L.get('r'), L.get('seq'))
    v1_links.setdefault(key, []).append(L)

def how_for(file, works_):
    if '発注台帳' in file:
        cts = {WID[w].get('contract_type') for w in works_ if w in WID}
        if len(works_) > 3 and cts <= {'T'}:
            return 'annual'
        if len(works_) > 3:
            return 'bundle'
    if '工事報告書' in file and len(works_) > 3:
        return 'rollup'
    return 'direct'

plan_notes = []
for stem, A in A1.items():
    f = A['file']
    plan = bool(A.get('plan'))
    recs = []
    prev_plan_marks = {}   # path -> 計画行の marks（a3）
    for rc in A['records']:
        r2 = {k: v for k, v in rc.items() if k not in ('fields',)}
        if f.endswith('.pdf') and rc.get('seq') is not None:
            r2['r'] = rc['seq'] + 1   # PDF は r = seq + 1（seq は残す）
        flds = dict(rc['fields'])
        if f.endswith(('.pdf', '.csv')):
            r2.pop('sheet', None)
        if plan:
            p = path_list(flds.get('path', ''))
            flds['path'] = p
            raw = flds.get('marks') or {}
            m = {str(y): disp(v) for y, v in raw.items() if v is not None and len(disp(v)) <= 8}
            kubun = flds.get('区分')
            if '星取表' in f:
                if kubun == '計画':
                    flds['marks'] = m; flds['plans'] = dict(m)
                    prev_plan_marks[tuple(p)] = m
                elif kubun == '実績':
                    flds['marks'] = m; flds['plans'] = dict(prev_plan_marks.get(tuple(p), {}))
                else:
                    flds['marks'] = m; flds['plans'] = {}
            else:
                # a1・a2：同じ行の年度の列が予定（金額）。a2 R1 改訂版の H25〜H30 列だけは実績額
                act_cols = [str(y) for y in range(2013, 2019)] if '令和元年度改訂' in f else []
                flds['marks'] = {y: v for y, v in m.items() if y in act_cols}
                flds['plans'] = {y: v for y, v in m.items() if y not in act_cols}
            # 結び：印ごとに年度つき
            code = flds.get('code') if flds.get('code') in AS else None
            if code is None:
                for L in v1_links.get((f, rc.get('sheet'), rc.get('r'), rc.get('seq')), []):
                    if L.get('code') in AS:
                        code = L['code']
            if code and p:
                AS[code]['names'].setdefault(f, ' > '.join(p))
            if rc['kind'] == 'item':
                legal_id = None
                if p and p[0].startswith('法'):
                    nm = p[-1]
                    legal_id = next((L['id'].split('-')[0] for L in W1['legal_inspections'] if L['name'] == nm), None)
                actual_row = (kubun == '実績') or (not kubun and False)
                for y, sym in flds['marks'].items():
                    yy = int(y)
                    if legal_id:
                        ws_ = ['%s-%d' % (legal_id, yy)] if ('%s-%d' % (legal_id, yy)) in WID else plan_work(None, yy, sym)
                    elif code:
                        ws_ = WBY.get((code, yy)) if (kubun == '実績' or '令和元年度改訂' in f) else plan_work(code, yy, sym)
                        ws_ = ws_ or plan_work(code, yy, sym)
                    else:
                        continue
                    for w in ws_:
                        links.append(dict(file=f, sheet=rc.get('sheet'), r=rc.get('r'), work=w, year=yy, how='direct', col='実績' if kubun == '実績' else ('計画' if kubun else '実績額')))
                for y, sym in flds['plans'].items():
                    if kubun:   # 星取表は計画行の印を上で結んだ（実績行の plans は計画行の写し）
                        continue
                    yy = int(y)
                    if legal_id:
                        ws_ = ['%s-%d' % (legal_id, yy)] if ('%s-%d' % (legal_id, yy)) in WID else plan_work(None, yy, sym)
                    elif code:
                        ws_ = plan_work(code, yy, sym)
                    else:
                        continue
                    for w in ws_:
                        links.append(dict(file=f, sheet=rc.get('sheet'), r=rc.get('r'), work=w, year=yy, how='direct', col='予定'))
        else:
            # 工程表の行：日数
            if flds.get('start') and flds.get('end'):
                try:
                    flds['days'] = (dt.date.fromisoformat(flds['end']) - dt.date.fromisoformat(flds['start'])).days + 1
                except Exception:
                    pass
            for L in v1_links.get((f, rc.get('sheet'), rc.get('r'), rc.get('seq')), []):
                if L.get('code') in AS:
                    nm = flds.get('name') or flds.get('place')
                    if nm:
                        sysn = AS[L['code']]['system']
                        fur = {'1': '1号炉', '2': '2号炉'}.get(AS[L['code']]['furnace'], '共通')
                        AS[L['code']]['names'].setdefault(f, ' > '.join([sysn, fur, nm]) if '台帳' in f else ' > '.join([fur, nm]))
                ws_ = L.get('works') or []
                how = how_for(f, ws_)
                for w in ws_:
                    if w not in WID:
                        continue
                    lk = dict(file=f, sheet=None if f.endswith(('.pdf', '.csv')) else L.get('sheet'), r=L.get('r'), work=w, how=how)
                    if L.get('seq') is not None:
                        lk['seq'] = L['seq']
                        if f.endswith('.pdf'):
                            lk['r'] = L['seq'] + 1
                    if flds.get('date'):
                        lk['day'] = flds['date']
                    links.append(lk)
        r2['fields'] = flds
        recs.append(r2)
    out = dict(file=f, plan=plan, records=recs)
    for k in ('image_size',):
        if k in A:
            out[k] = A[k]
    json.dump(out, open(OUT + '/files/' + stem + '.json', 'w', encoding='utf-8'), ensure_ascii=False, indent=1)

# ------------------------------------------------------------------ 計画の欠け → errata（印の抜けだけ）
errata, gaps_rest = [], []
for g in W1['plan_gaps']:
    if g['kind'] == '法令点検の抜け' and 'H36' in g['where']:
        A = A1[os.path.splitext(g['file'])[0]]
        rc = next(x for x in A['records'] if 'ボイラー定期事業者検査（2号炉）' in str(x['fields'].get('path', '')))
        errata.append(dict(type='omitMark', file=g['file'], sheet=rc['sheet'], r=rc['r'], year=2024, works=['LG02-2024'], note=g['what']))
    else:
        gaps_rest.append(g)

# 名前の書き方：台帳 R1 改訂の名前を基本に、資料ごとの呼び名
for a in assets:
    a['names'] = dict(sorted(a['names'].items()))

PROFILE = W1['profile']
world = {
    'schema': 'moyai.sample-world/0',
    'facility': dict(id='F01', name='【架空】あさぎ市清掃工場'),
    'profile': dict(furnace=PROFILE['furnace'], capacity=PROFILE['capacity'], boiler=PROFILE['boiler'], power=PROFILE['power'], period=PROFILE['period'],
                    equipment_count=PROFILE['equipment_count'], counting=PROFILE['counting']),
    'maker': dict(today='2026-11-13'),
    'assets': assets,
    'works': works,
    'links': links,
    'errata': errata,
    'files': [dict(file=A['file']) for A in A1.values()],
}
json.dump(world, open(OUT + '/world.json', 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
extra = dict(profile=PROFILE, equipment=W1['equipment'], legal_inspections=W1['legal_inspections'], unrecorded_works=W1['unrecorded_works'],
             plan_gaps_other=gaps_rest, plan_gaps_all=W1['plan_gaps'], doc_links_v1=W1['doc_links'], generated=W1.get('generated'),
             notes=dict(
                 marks='星取表（a3）は計画・実績の 2 段：計画行は marks＝plans＝その行の記号、実績行は marks＝実績の記号・plans＝同じ機器の計画行の記号。a1・a2 は同じ行の年度の列に金額が入る様式で、a2 R1 改訂版の H25〜H30 列だけが実績額（marks）、ほかの列は予定（plans）。値はセルに見えている字（3 桁区切り）。8 字を越える字は印にしない',
                 links='予定の印は、その年度にその機器の実績があれば実績の作業へ、無ければ予定の作業（2026 年度以降 planned、2025 年度以前 not_done、id は PL-機器コード-年度）へ結んだ。法令点検は LGxx-年度',
                 how='発注台帳の定期整備委託の行（定修の作業を多数まとめる）は annual、改造工事・補修以外で多数をまとめる行は bundle、報告書の多数をまとめる行は rollup、ほかは direct',
                 installed='設置年は年度（竣工時の機器は 1997＝平成 9 年 3 月竣工）',
                 manDays='manDays は人日（1 人日＝8 時間）。帳票の字は manDaysText に残した',
                 status='done＝実績、in_progress＝R8 2号炉定修で施工中（記録時点 2026-11-13）、planned＝予定、not_done＝過去の予定で実施されなかったもの'))
json.dump(extra, open(OUT + '/world.extra.json', 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
print('assets', len(assets), 'works', len(works), 'links', len(links), 'errata', len(errata), 'files', len(world['files']))
