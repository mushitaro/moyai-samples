# -*- coding: utf-8 -*-
"""answer/ を物差しの形（moyai.sample-world/0）に書き直す。files/ は読まない・書かない。
入力：f10/answer_v1_backup/（旧正解の控え）。出力：f10/full/answer/。  python3 to_v2.py"""
import os, json, glob, shutil, datetime as dt, re
HERE = os.path.dirname(os.path.abspath(__file__))
FULL = os.path.dirname(HERE)
SRC = os.path.join(FULL, 'gen', '_answer_v1')
DST = os.path.join(FULL, 'answer')
SYM = {'○': '点検', '◇': '計測・診断', '◎': '分解整備', '△': '補修', '●': '更新', '★': '基幹改良で更新', '□': '法定点検'}
ATTR = {'○': 'INSPECT', '◇': 'INSPECT', '◎': 'OVERHAUL', '△': 'OVERHAUL', '●': 'REPLACE', '★': 'UPGRADE', '□': 'INSPECT'}

def fy_of(s):
    d = dt.date.fromisoformat(s); return d.year if d.month >= 4 else d.year - 1

def hours(a, b):
    h1, m1 = map(int, a.split(':')); h2, m2 = map(int, b.split(':'))
    t = (h2 * 60 + m2 - h1 * 60 - m1) / 60
    if h1 < 12 and h2 >= 13: t -= 1.0
    return round(t, 1)

def clean_mark(v):
    v = str(v)
    return v if len(v) <= 8 else None

def main():
    W = json.load(open(os.path.join(SRC, 'world.json')))
    man = json.load(open(os.path.join(SRC, 'manifest.json')))
    if os.path.exists(DST): shutil.rmtree(DST)
    os.makedirs(os.path.join(DST, 'files'))
    units = {u['tag']: u for u in W['units']}
    # ---------------------------------------------------------------- 作業
    works = []; byid = {}
    for w in W['works']:
        md = round(w['manHours'] / 8, 2) if w.get('manHours') else None
        days = (dt.date.fromisoformat(w['end']) - dt.date.fromisoformat(w['start'])).days + 1 if w.get('end') and w.get('start') else None
        st = 'done'
        if w['start'] > '2026-11-13': st = 'planned'
        elif w.get('status') == '施工中' or w.get('end') in (None, '2026-11-13') and w['start'] >= '2026-10-01': st = 'in_progress'
        x = dict(id=w['id'], asset=w['unit'], fy=fy_of(w['start']), status=st,
                 task=f"{SYM.get(w['symbol'], '点検')}（{w['unitName']}）" + (f"：{w['memo']}" if w.get('memo') else ''), attr=w['type'],
                 start=w['start'], end=w.get('end'), manDays=md, days=days, contract=w['contract'], vendor=w['vendor'])
        works.append(x); byid[x['id']] = x
    key = {}
    for x in works: key.setdefault((x['asset'], x['fy']), x)
    nid = [len(works)]
    def new_work(asset, fy, sym, status, memo):
        nid[0] += 1
        x = dict(id=f'W{nid[0]:05d}', asset=asset, fy=fy, status=status, task=f"{SYM.get(sym, sym)}（{units[asset]['name'] if asset in units else asset}）" + (f'：{memo}' if memo else ''),
                 attr=ATTR.get(sym, 'INSPECT'), start=None, end=None, manDays=None, days=None)
        works.append(x); byid[x['id']] = x; key[(asset, fy, status)] = x
        return x
    # どの資料にも無い実績（月例）を done の作業に
    for u in W['unrecordedWorks']:
        if (u['unit'], u['year']) not in key and u['symbol'] != '★':
            x = new_work(u['unit'], u['year'], u['symbol'], 'done', u['reason']); key[(u['unit'], u['year'])] = x
    def plan_work(asset, fy, sym, kind):
        """kind：計画の印 → 実施済なら done の作業、無ければ not_done（過去）／planned（今年度以降）。×（繰越）→ not_done"""
        d = key.get((asset, fy))
        if d and kind != 'x': return d
        st = 'not_done' if (kind == 'x' or fy < 2026) else 'planned'
        if (asset, fy, st) in key: return key[(asset, fy, st)]
        return new_work(asset, fy, sym, st, '計画の印（実施の記録なし）' if st == 'not_done' else '中長期整備計画の予定')
    # ---------------------------------------------------------------- 記録と結び
    links = []
    v1links = {}
    for l in W['links']: v1links.setdefault((l['file'], l['sheet'], l['r'], l['seq']), []).append(l['work'])
    clinks = {}
    for l in W['contractLinks']: clinks.setdefault((l['file'], l['r']), []).append(l['work'])
    contract_works = {}
    for w in W['works']: contract_works.setdefault(w['contractKey'], []).append(w['id'])
    legal_assets = {}
    notes = []
    for fn in sorted(glob.glob(os.path.join(SRC, 'files', '*.json'))):
        a = json.load(open(fn)); f = a['file']; ext = f.rsplit('.', 1)[1]
        recs = []
        plan = a['plan']
        # 同じシートの計画行（code ごと）
        planrow = {}
        if plan:
            for x in a['records']:
                if x['fields'].get('row') in ('計画', None) and 'marks' in x['fields'] and x['kind'] == 'item':
                    planrow[(x['sheet'], x['fields'].get('code') or tuple(x['fields'].get('path', [])))] = x['fields'].get('marks', {})
        for x in a['records']:
            y = dict(x); fl = dict(x['fields'])
            if ext in ('pdf', 'csv'): y['sheet'] = None
            if ext == 'pdf': y['r'] = x['seq'] + 1
            if plan and x['kind'] in ('item', 'parent'):
                row = fl.get('row'); ident = (x['sheet'], fl.get('code') or tuple(fl.get('path', [])))
                marks = {k: v for k, v in fl.get('marks', {}).items() if clean_mark(v)}
                if '済' in ''.join(marks.values()) or (row is None and 'code' in fl):   # j1：同じ行・同じ列（実施済は「済」）
                    act = {k: v.replace('済', '') for k, v in marks.items() if v.endswith('済')}
                    pl = {k: v for k, v in marks.items() if not v.endswith('済') and v != '×'}
                    xs = {k: v for k, v in marks.items() if v == '×'}
                    fl['marks'] = {**act, **xs}; fl['plans'] = pl
                elif row == '計画':
                    fl['plans'] = dict(marks); fl['marks'] = dict(marks)
                elif row == '実績':
                    fl['marks'] = dict(marks); fl['plans'] = dict(planrow.get(ident, {}))
                elif row is None and 'done' in fl:   # j3 集計・法定の 1 行（灰色＝実施済）
                    fl['plans'] = dict(marks); fl['marks'] = {k: v for k, v in marks.items() if k in fl['done']}
                else:
                    fl['marks'] = {}; fl['plans'] = {}
                fl['path'] = fl.get('path', [])
                # 結び
                asset = fl.get('code')
                if asset in units:
                    for yr, sym in fl['marks'].items():
                        if row == '計画':
                            w = plan_work(asset, int(yr), sym, 'p')
                        elif sym == '×':
                            w = plan_work(asset, int(yr), fl['plans'].get(yr, '○'), 'x')
                        else:
                            w = key.get((asset, int(yr))) or plan_work(asset, int(yr), sym, 'p')
                            if w['status'] in ('not_done', 'planned'):
                                w = new_work(asset, int(yr), sym, 'done', '計画の実績欄に記号（作業の記録は月例点検）'); key[(asset, int(yr))] = w
                        links.append(dict(file=f, sheet=y['sheet'], r=y['r'], work=w['id'], year=int(yr), how='direct'))
                    if row is None or row in ('実績',):
                        for yr, sym in fl['plans'].items():
                            if yr in fl['marks']: continue
                            w = plan_work(asset, int(yr), sym, 'p')
                            links.append(dict(file=f, sheet=y['sheet'], r=y['r'], work=w['id'], year=int(yr), how='direct'))
                elif fl['path'] and fl['path'][0] == '法定点検・測定等':
                    nm = fl['path'][-1]
                    la = legal_assets.setdefault(nm, f'L-{len(legal_assets)+1:02d}')
                    for yr, sym in {**fl['plans'], **fl['marks']}.items():
                        st = 'done' if (yr in fl['marks'] and sym != '×') else ('not_done' if int(yr) < 2026 or sym == '×' else 'planned')
                        k2 = (la, int(yr), st)
                        w = key.get(k2) or new_work(la, int(yr), '□', st, nm)
                        links.append(dict(file=f, sheet=y['sheet'], r=y['r'], work=w['id'], year=int(yr), how='direct'))
            if not plan:
                # 工数・日数
                if 'manDays' in fl and fl['manDays'] is not None:
                    fl['manDaysText'] = str(fl['manDays']) if 'manDaysText' not in fl else fl['manDaysText']
                    if fl.get('actStart') and fl.get('actEnd'): fl['days'] = (dt.date.fromisoformat(fl['actEnd']) - dt.date.fromisoformat(fl['actStart'])).days + 1
                    elif fl.get('planStart') and fl.get('planEnd'): fl['days'] = (dt.date.fromisoformat(fl['planEnd']) - dt.date.fromisoformat(fl['planStart'])).days + 1
                elif 'manHours' in fl:
                    fl['manDaysText'] = f"{fl['manHours']}"
                    fl['manDays'] = round(fl['manHours'] / 8, 2)
                    if x['kind'] == 'item': fl['days'] = 1
                elif 'persons' in fl and 'start' in fl and 'end' in fl:
                    fl['manDaysText'] = f"{fl['persons']}名 {fl['start']}〜{fl['end']}"
                    fl['manDays'] = round(fl['persons'] * hours(fl['start'], fl['end']) / 8, 2); fl['days'] = 1
                elif 'persons' in fl and x['kind'] == 'item':
                    fl['manDaysText'] = f"{fl['persons']}名"; fl['manDays'] = float(fl['persons']); fl['days'] = 1
                wl = v1links.get((f, x['sheet'], x.get('r'), x.get('seq')), [])
                day = fl.get('date') if '作業日報' in f else None
                for wid in wl:
                    links.append(dict(file=f, sheet=y['sheet'], r=y['r'], work=wid, day=day, how='direct'))
                if ext == 'csv' and (f, x.get('r')) in clinks:
                    for ck in clinks[(f, x['r'])]:
                        ws_ = contract_works.get(ck, [])
                        how = 'upgrade' if ck == '基幹改良' else 'annual' if ck.startswith('保守') else 'bundle' if len(ws_) > 1 else 'direct'
                        for wid in ws_: links.append(dict(file=f, sheet=None, r=x['r'], work=wid, how=how))
                if '副資材' in f and (f, x.get('r')) not in clinks: pass
            y['fields'] = fl
            y.pop('_works', None)
            recs.append(y)
        json.dump({'file': f, 'plan': plan, 'records': recs}, open(os.path.join(DST, 'files', os.path.basename(fn)), 'w'), ensure_ascii=False, indent=0)
    # 法定点検の「機器」
    assets = []
    for u in W['units']:
        sysname = u['system'] + ' > ' + u['equipment']
        names = {'機器台帳': u['name'], '中長期整備計画': ' > '.join([u['system'], u['furnace'], u['equipment'], u['aliases']['中長期整備計画']])}
        if '電動機台帳（H23）' in u['aliases']: names['機器台帳（H23 電動機台帳）'] = u['aliases']['電動機台帳（H23）']
        if '日報の場所' in u['aliases']: names['作業日報'] = u['aliases']['日報の場所']
        assets.append(dict(id=u['tag'], name=u['name'], tag=u['tag'], system=sysname, furnace={'1号炉': '1', '2号炉': '2'}.get(u['furnace'], '共通'),
                           installed=u['installed'], **{'class': u['importance']}, part=u['parent'], names=names, kind=u['kind'], duty=u['duty']))
    for nm, la in legal_assets.items():
        assets.append(dict(id=la, name=nm, tag=None, system='法定点検・測定等', furnace='共通', installed=None, **{'class': None}, part=None,
                           names={'中長期整備計画': '法定点検・測定等 > ' + nm}, kind='法定点検（機器ではない）', duty=None))
    world = {'schema': 'moyai.sample-world/0',
             'facility': {'id': 'F10', 'name': '【架空】あかね台資源化センター'},
             'profile': W['profile'], 'maker': {'today': '2026-11-13'},
             'assets': assets, 'works': works, 'links': links, 'errata': [],
             'files': [{'file': m['file']} for m in man]}
    json.dump(world, open(os.path.join(DST, 'world.json'), 'w'), ensure_ascii=False, indent=0)
    extra = {'contracts': W['contracts'], 'contractLinks': W['contractLinks'], 'unrecordedWorks': W['unrecordedWorks'], 'planGaps': W['planGaps'], 'notes': W['notes'] + [
        'errata は空：わざと入れた計画の欠けに「印の抜け（omitMark）」「前年度の実績の反映漏れ（lastYearUnreflected）」に当たるものが無いため。欠けの一覧は planGaps。',
        '星取表の marks・plans の当て方：計画の行は marks＝plans＝その行の記号。実績の行は marks＝その行の記号（×＝繰越を含む）、plans＝同じシートの同じ機器の計画の行の記号。'
        'j1（H23・H25）は同じ行・同じ列の様式で、「◎済」のように済の付いた記号を実績（marks、済を外す）、済の無い記号を予定（plans）とした。j3 の集計・法定の 1 行は灰色（done）の年を marks、全部の □ を plans。'
        '費用の行は marks・plans とも空。非表示の列の記号は hiddenMarks に残し、結びは付けない。',
        '作業の status：実施は done（施工中は in_progress）。計画の印で実施の記録が無い過去の年度は not_done、今年度以降は planned。実績の行の ×（繰越）は not_done。',
        '法定点検は機器ではないので assets に L-xx の項目として置いた（kind で区別）。',
        '工数：日報（R1 様式）は工数（人・h）÷8＝人日、H25 様式は人数×時間÷8、班別様式は人数（1 人 1 日）。工程表の延人日はそのまま人日。書いてある字は manDaysText。',
        '発注 CSV の行の結び：年間保守は annual、基幹改良は upgrade、複数の作業の補修工事・物品は bundle、1 件なら direct。副資材（単価契約）の行は機器の作業に結ばない。'],
        'v1_manifest': man}
    json.dump(extra, open(os.path.join(DST, 'world.extra.json'), 'w'), ensure_ascii=False, indent=0)
    shutil.copy(os.path.join(SRC, 'manifest.json'), os.path.join(DST, 'manifest.json'))
    print('assets', len(assets), 'works', len(works), 'links', len(links))

if __name__ == '__main__':
    main()
