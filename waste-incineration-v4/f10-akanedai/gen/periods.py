# -*- coding: utf-8 -*-
"""本作り：年度ごとの契約・施工期間・工程（正解）。工程表→日報→報告書→発注はここから作る。
期間 P = dict(key, fy, kind, contract, title, part, start, end, items, milestones, stop, groups, sv, made, other, status)
作業 it = dict(no, name, units[tag], co, team, ppl, plan, act, note, tmpl, sym, added)"""
import datetime as dt, random
import world, history as H, jobs

D = dt.date
CH = '千草環境'; SN = '東雲築炉'; AD = 'あかね電設'; MT = '丸鳥足場工業'; HK = 'ひかり工業'; CITY = '市運転班'
FY0, FY1 = 2011, 2026
TODAY = D(2026, 11, 13)
FACN = world.FACN

def wa(fy):
    return f'平成{fy-1988}年度' if fy < 2019 else ('令和元年度' if fy == 2019 else f'令和{fy-2018}年度')

def monday_on_or_after(d):
    return d + dt.timedelta((7 - d.weekday()) % 7)

def city_staff(fy):
    return 'j1' if fy <= 2015 else 'j2' if fy <= 2023 else 'j3'

def sv(fy):
    return 'p1' if fy <= 2017 else 'p2'

U = {}
def unit(furn, base):
    k = (furn, base)
    if k not in U:
        U[k] = next(u for u in world.UNITS if u['furn'] == furn and (u['base'] == base or u['name'] == base))
    return U[k]

# ------------------------------------------------------------------ 期間の日付
def outage_dates(fy, f):
    r = random.Random(f'out-{fy}-{f}')
    fixed = {(2018, 2): (D(2018, 6, 4), D(2018, 7, 6)), (2018, 1): (D(2018, 10, 1), D(2018, 11, 9)),
             (2021, 2): (D(2021, 6, 7), D(2021, 7, 9)), (2021, 1): (D(2021, 9, 27), D(2021, 11, 9)),
             (2022, 1): (D(2022, 9, 5), D(2022, 12, 22)), (2023, 2): (D(2023, 5, 29), D(2023, 9, 14)),
             (2026, 1): (D(2026, 10, 5), D(2026, 11, 20))}
    if (fy, f) in fixed: return fixed[(fy, f)]
    if f == 2:
        s = monday_on_or_after(D(fy, 6, 1)); L = r.randint(32, 35)
    else:
        s = monday_on_or_after(D(fy, 9, 24) + dt.timedelta(r.randint(0, 7))); L = r.randint(39, 44)
    return s, s + dt.timedelta(L - 1)

def allstop_dates(fy):
    if fy == 2024: return D(2024, 5, 7), D(2024, 5, 24)
    s = monday_on_or_after(D(fy, 5, 8)); return s, s + dt.timedelta(9)

# ------------------------------------------------------------------ 炉の定修の標準作業（R3 の 44 日を基準に伸縮）
STD = [  # key, 作業, 対象(base), 会社, 班, 人員, 始, 終, 条件, 型
    ('1', '立下げ・炉冷却', '溶融炉本体', CITY, '運転', 0, 0, 6, None, 'city'),
    ('2', '入槽前測定・安全措置', '溶融炉本体', CH, '機械A', 3, 7, 7, None, 'safety'),
    ('3', '足場組立（炉内・ボイラ・BF）', '溶融炉本体', MT, '足場', 4, 7, 9, None, 'scaffold'),
    ('4', '炉内残留物除去', '溶融炉本体', SN, '炉内', 6, 8, 11, None, 'residue'),
    ('5', '羽口点検・漏水試験', '主羽口装置', CH, '機械A', 3, 9, 10, None, 'tuyere'),
    ('6', '炉内耐火物点検・部分補修', '溶融炉本体', SN, '炉内', 5, 14, 23, None, 'refr'),
    ('7', 'ボイラ水洗', 'ボイラ本体', HK, '清掃', 6, 8, 11, None, 'wash'),
    ('8', '過熱器・蒸発管 肉厚測定', '2次過熱器', CH, '機械B', 2, 14, 16, 'sh', 'ut'),
    ('9', 'ボイラ水圧試験', 'ボイラ本体', CH, '機械A', 4, 24, 24, None, 'hydro'),
    ('10', 'ボイラ定期事業者検査（立会）', 'ボイラ本体', CH, '機械A', 2, 25, 25, 'legal', 'legal'),
    ('11', '減温塔点検・噴霧ノズル', '減温塔噴霧ノズル', CH, '機械B', 2, 9, 10, None, 'nozzle'),
    ('12', 'BF 内部点検', 'ろ過式集じん器本体', CH, '機械B', 3, 15, 18, None, 'bf'),
    ('13', '触媒反応塔 点検', '触媒反応塔', CH, '機械B', 2, 21, 21, None, 'cat'),
    ('14', '誘引通風機 点検', '誘引通風機', CH, '機械B', 2, 22, 22, None, 'idf'),
    ('17', '水砕槽清掃・掻揚コンベヤ点検', 'スラグ掻揚コンベヤ', HK, '清掃', 4, 8, 11, None, 'slag'),
    ('18', '送風機・燃焼空気送風機 点検', '羽口送風機', CH, '機械A', 2, 21, 22, None, 'fan'),
    ('19', '給じん装置（シールダンパ）点検', '上部シールダンパ', CH, '機械A', 3, 22, 24, None, 'feed'),
    ('20', '計装校正', '炉頂温度計', AD, '電計', 3, 21, 29, None, 'cal'),
    ('21', 'MCC 点検', '炉別動力制御盤（MCC）', AD, '電計', 3, 28, 29, None, 'mcc'),
    ('22', '足場解体', '溶融炉本体', MT, '足場', 4, 28, 29, None, 'scaffold2'),
    ('23', '復旧・マンホール閉止', '溶融炉本体', CH, '機械A', 4, 29, 31, None, 'close'),
    ('24', 'インターロック試験', '炉別制御ステーション（FCS）', CH, '機械A', 3, 31, 31, None, 'il'),
    ('25', '炉乾燥・立上げ', '溶融炉本体', CITY, '運転', 0, 33, 38, None, 'city'),
    ('26', 'ごみ投入・試運転', '溶融炉本体', CITY, '運転', 0, 39, 43, None, 'city'),
]
# 標準作業が受け持つ機器（base）
STD_UNITS = {'4': ['溶融炉本体', '炉体冷却ジャケット', '燃焼室本体'], '5': ['主羽口装置', '副羽口装置'], '6': ['溶融炉本体', '炉体冷却ジャケット', '燃焼室本体'],
             '7': ['ボイラ本体'], '8': ['1次過熱器', '2次過熱器', 'エコノマイザ'], '9': ['ボイラ本体'], '11': ['減温塔噴霧ノズル', '減温塔本体'],
             '12': ['ろ過式集じん器本体'], '13': ['触媒反応塔'], '14': ['誘引通風機'], '17': ['スラグ掻揚コンベヤ', '水砕槽'],
             '18': ['羽口送風機', '燃焼用空気送風機'], '19': ['上部シールダンパ', '下部シールダンパ'], '21': ['炉別動力制御盤（MCC）'],
             '24': ['炉別制御ステーション（FCS）']}

CLS = {'rot': 'OH', 'mech': 'OH', 'crane': 'OH', 'crane2': 'OH', 'safety': 'OH', 'egen': 'OH', 'turbine': 'OH', 'gen': 'OH', 'scale': 'OH',
       'motor': 'EL', 'valve': 'VL', 'instr': 'EL', 'analyzer': 'EL', 'panel': 'EL', 'hv': 'EL', 'dcs': 'EL',
       'static': 'ST', 'vessel': 'ST', 'furnace': 'ST', 'furnace2': 'ST', 'taphole': 'TP', 'tuyere': 'ST', 'boiler': 'ST',
       'sh': 'ST', 'sh2': 'ST', 'bf': 'ST', 'cat': 'ST', 'nozzle': 'ST', 'psa': 'OH'}
CLS_NAME = {'OH': '回転機・機械 整備', 'EL': '電気・計装 点検整備', 'VL': '電動弁・調節弁 整備', 'ST': '静機器 点検補修', 'TP': '出湯口 補修'}
SYMW = {'○': 'INSPECT', '◇': 'INSPECT', '◎': 'OVERHAUL', '△': 'OVERHAUL', '●': 'REPLACE', '★': 'UPGRADE'}

def events(furn, fy, pred=lambda u: True):
    out = []
    for u in world.UNITS:
        if u['furn'] != furn or not pred(u): continue
        if u.get('ident') == 'ak_new' and fy < 2024: continue
        h = H.truth(u)
        if fy in h:
            s, d = h[fy]
            if s == '★': continue   # 基幹改良の契約で施工
            out.append((u, s, d))
    return out

def scale(off, L):
    return round(off * (L - 1) / 43)

def mk(no, name, units, co, team, ppl, plan, act, note='', tmpl='ev', sym=None, added=False):
    return dict(no=no, name=name, units=units, co=co, team=team, ppl=ppl, plan=plan, act=act, note=note, tmpl=tmpl, sym=sym, added=added)

REPAIR_IN = {}   # (fy, furn) → 別契約の補修工事が受け持つ機器 base

def furnace_items(fy, f, s, e, ak=False):
    """標準作業＋その年の出来事のまとまり。実績は乱数で少しずらす。"""
    L = (e - s).days + 1
    r = random.Random(f'it-{fy}-{f}')
    ev = events(f, fy)
    evb = {u['base']: (u, sy, d) for (u, sy, d) in ev}
    legal_year = (fy % 2 == 1) if f == 1 else (fy % 2 == 0)
    items = []; covered = set()
    for key, name, base, co, team, ppl, a, b, cond, tmpl in STD:
        if ak and key in ('4', '5', '6'): continue  # 基幹改良の工程で施工
        if cond == 'sh' and not any(x in evb for x in ('1次過熱器', '2次過熱器', 'エコノマイザ')): continue
        if cond == 'legal' and not legal_year: continue
        if tmpl == 'cat' and '触媒反応塔' not in evb: continue
        if ak:
            a2, b2 = a, b
            if tmpl == 'city' and key == '25': a2, b2 = 33, 38
        A = s + dt.timedelta(scale(a, L) if not ak else min(a, L - 1))
        Bd = s + dt.timedelta(scale(b, L) if not ak else min(b, L - 1))
        if ak and key in ('22', '23', '24', '25', '26'):   # 改良工事の後ろに寄せる
            A = e - dt.timedelta(43 - a); Bd = e - dt.timedelta(43 - b)
        nm = name; units = [unit(f, base)['tag']]
        for bb in STD_UNITS.get(key, []):
            if bb in REPAIR_IN.get((fy, f), []): continue
            if bb in evb: covered.add(bb)
            try: t = unit(f, bb)['tag']
            except StopIteration: continue
            if t not in units: units.append(t)
        if tmpl == 'nozzle': nm = '減温塔点検・噴霧ノズル交換' if evb.get('減温塔噴霧ノズル', (0, ''))[1] == '●' else '減温塔内部点検'
        if tmpl == 'bf' and evb.get('ろ過式集じん器本体', (0, ''))[1] == '●':
            nm = 'BF ろ布全数交換'; Bd = A + dt.timedelta(5); ppl = 5
        if tmpl == 'idf' and evb.get('誘引通風機', (0, ''))[1] == '◎':
            nm = '誘引通風機 分解整備（軸受・羽根車）'; A = s + dt.timedelta(scale(21, L)); Bd = A + dt.timedelta(4); ppl = 4
        if tmpl == 'idf' and '誘引通風機' not in evb: nm = '誘引通風機 振動測定・外観点検'
        act = (A, Bd)
        note = ''
        if tmpl not in ('city', 'legal', 'hydro', 'il') and r.random() < 0.06:
            act = (A, Bd + dt.timedelta(1)); note = r.choice(['部品入荷待ちで 1日延長', '付着物が多く 1日延長', '雨天で 1日延長'])
        items.append(mk(key, nm, units, co, team, ppl, (A, Bd), act, note, tmpl, sym=None))
    # 出湯口
    if '出湯口' in evb and (fy, f) not in REPAIR_IN:
        u, sy, d = evb['出湯口']; covered.add('出湯口')
        A = s + dt.timedelta(scale(15, L)); Bd = A + dt.timedelta(4 if sy == '△' else 10)
        items.append(mk('7T', '出湯口 部分補修' if sy == '△' else '出湯口 再築', [u['tag']], SN, '炉内', 3 if sy == '△' else 5, (A, Bd), (A, Bd), '', 'tap', sy))
    # その他の出来事を系統×区分でまとめる
    groups = {}
    for (u, sy, d) in ev:
        if u['base'] in covered or u['base'] in sum(STD_UNITS.values(), []): continue
        if (fy, f) in REPAIR_IN and (u['base'] in REPAIR_IN[(fy, f)] or ('__analyzer__' in REPAIR_IN[(fy, f)] and u['cat'] == 'analyzer')): continue
        c = CLS.get(u['cat'], 'ST')
        groups.setdefault((u['sys'], u['sysname'], c), []).append((u, sy, d))
    k = 0
    for (sy_, sn_, c), lst in sorted(groups.items()):
        k += 1
        n = len(lst); dur = max(1, min(6, 1 + n // 2))
        st = scale(12, L) + (k * 3) % max(1, scale(16, L))
        A = s + dt.timedelta(st); Bd = A + dt.timedelta(dur - 1)
        if Bd > s + dt.timedelta(scale(30, L)): Bd = s + dt.timedelta(scale(30, L)); A = min(A, Bd)
        co = AD if c == 'EL' else CH
        nm = f"{sn_} {CLS_NAME[c]}（{n}台）"
        items.append(mk(f'E{k}', nm, [u['tag'] for (u, _, _) in lst], co, '電計' if c == 'EL' else '機械B', min(4, 1 + (n + 1) // 2),
                        (A, Bd), (A, Bd), '', 'ev_' + c, None))
    return items

def common_items(fy, s, e, ak=False):
    """全炉停止（共通設備）：受変電の年次点検（別契約）と共通機器の出来事。"""
    L = (e - s).days + 1
    items = [mk('1', '全炉停止・系統切離し', [unit(0, '特高受電設備（GIS）')['tag']], CITY, '運転', 0, (s, s), (s, s), '市直営', 'city'),
             mk('2', '共通動力制御盤・低圧配電盤 点検', [u['tag'] for u in world.UNITS if u['furn'] == 0 and u['sys'] == '11' and u['sub'] == '4'],
                AD, '電計', 3, (s + dt.timedelta(1), s + dt.timedelta(3)), (s + dt.timedelta(1), s + dt.timedelta(3)), '停電作業', 'mcc')]
    ev = events(0, fy, pred=lambda u: not (u['sys'] == '05' and u['sub'] in ('2', '3')) and not other_vendor(u, fy))
    groups = {}
    for (u, sy, d) in ev:
        c = CLS.get(u['cat'], 'ST')
        groups.setdefault((u['sys'], u['sysname'], c), []).append((u, sy, d))
    k = 0
    for (sy_, sn_, c), lst in sorted(groups.items()):
        k += 1; n = len(lst); dur = max(1, min(L - 3, 1 + n // 2))
        A = s + dt.timedelta(1 + (k * 2) % max(1, L - dur - 2)); Bd = A + dt.timedelta(dur - 1)
        items.append(mk(f'E{k}', f"{sn_} {CLS_NAME[c]}（{n}台）", [u['tag'] for (u, _, _) in lst], AD if c == 'EL' else CH,
                        '電計' if c == 'EL' else '機械', min(4, 1 + (n + 1) // 2), (A, Bd), (A, Bd), '', 'ev_' + c))
    items.append(mk('9', '復電・系統復旧', [unit(0, '特高受電設備（GIS）')['tag']], CITY, '運転', 0, (e, e), (e, e), '市直営', 'city'))
    other = [('受変電', 'R-1', '受変電設備 年次点検（全停）', '特高受電設備（GIS）', '電気保安協会', 4, (s, s + dt.timedelta(1)), (s, s + dt.timedelta(1)), '市の別契約')]
    return items, other

def turbine_items(fy, s):
    """2号定修の期間に行うタービン・空冷復水器の作業（片炉運転、タービンバイパス）。"""
    ev = events(0, fy, pred=lambda u: u['sys'] == '05' and u['sub'] in ('2', '3') and not other_vendor(u, fy))
    if not ev: return []
    its = []
    tb = [x for x in ev if x[0]['base'] in ('蒸気タービン', 'タービン減速機', '発電機')]
    rest = [x for x in ev if x not in tb]
    if tb:
        A = s + dt.timedelta(3); Bd = A + dt.timedelta(23)
        its.append(mk('T1', '蒸気タービン 開放点検（定期事業者検査）' if any(x[0]['base'] == '蒸気タービン' for x in tb) else 'タービン発電機 点検',
                      [x[0]['tag'] for x in tb], CH, 'タービン', 5, (A, Bd), (A, Bd), 'タービン停止・バイパス運転', 'turbine'))
    if rest:
        A = s + dt.timedelta(6); Bd = A + dt.timedelta(max(1, min(8, len(rest))) - 1)
        its.append(mk('T2', f'余熱利用設備 補機 整備（{len(rest)}台）', [x[0]['tag'] for x in rest], CH, 'タービン', 3, (A, Bd), (A, Bd), '', 'ev_OH'))
    return its

def milestones_for(items, s, e):
    by = {it['tmpl']: it for it in items}
    ms = [(s, 'ごみ投入停止')]
    if 'safety' in by: ms.append((by['safety']['act'][0], '入槽'))
    if 'legal' in by: ms.append((by['legal']['act'][0], '定期事業者検査'))
    elif 'hydro' in by: ms.append((by['hydro']['act'][0], '水圧試験'))
    c = [it for it in items if it['no'] == '25']
    if c: ms.append((c[0]['act'][0], '乾燥'))
    c = [it for it in items if it['no'] == '26']
    if c: ms.append((c[0]['act'][0], 'ごみ投入'))
    return ms

def stop_for(s, e, items):
    c25 = [it for it in items if it['no'] == '25']; c26 = [it for it in items if it['no'] == '26']
    op = c26[0]['act'][0] if c26 else e
    clo = [it for it in items if it['no'] == '23']
    cl = clo[0]['act'][0] if clo else op - dt.timedelta(10)
    return [(s, s + dt.timedelta(2), '立下げ'), (s + dt.timedelta(3), s + dt.timedelta(6), '冷却'), (s + dt.timedelta(7), cl - dt.timedelta(1), '開放・施工'),
            (cl, op - dt.timedelta(1), '閉止・乾燥・昇温'), (op, e, '運転')]

def groups_for(items):
    G = [('共通', ['1', '2', '3', '22', '23']), ('溶融炉・スラグ', ['4', '5', '5-2', '6', '7T', '17', '18']), ('ボイラ', ['7', '8', '8-2', '9', '10']),
         ('排ガス処理', ['11', '12', '13', '14']), ('電気・計装', ['19', '20', '21']), ('立上げ', ['24', '25', '26'])]
    ev = [it['no'] for it in items if it['no'].startswith(('E', 'T'))]
    return G[:5] + [('その他の機器（出来事）', ev)] + G[5:]

def from_jobs(J, f):
    """見本の H30・R3 の 1号炉定修：見本の作業をそのまま使い、見本に無かった出来事のまとまりを足す。"""
    items = []
    for no, name, base, co, team, ppl, pl, ac, note in J['items']:
        try: t = unit(f, base)['tag']
        except StopIteration: t = unit(f, '溶融炉本体')['tag']
        units = [t] + [unit(f, b)['tag'] for b in STD_UNITS.get(no.split('-')[0], []) if unit(f, b)['tag'] != t]
        tmpl = {'1': 'city', '25': 'city', '26': 'city', '23': 'city', '24': 'city'}.get(no, 'std') if co == CITY else 'std'
        items.append(mk(no, name, units, co, team, ppl, pl, ac, note, tmpl, added=pl is None))
    return items

def build():
    world.build()
    REPAIR_IN.clear()
    for c in REPAIRS:
        for (fy, f, bases) in c.get('in', []): REPAIR_IN.setdefault((fy, f), []).extend(bases)
    P = []
    for fy in range(FY0, FY1 + 1):
        ck = f'保守{fy}'
        # 全停
        s, e = allstop_dates(fy)
        its, oth = common_items(fy, s, e)
        P.append(dict(key=f'{fy}-0', fy=fy, kind='全停', contract=ck, part='全炉停止（共通設備）', start=s, end=e, items=its, other=oth,
                      milestones=[(s, '全停'), (e, '復電')], stop=[(s, e, '全炉停止')], groups=None, book=3))
        for f in (2, 1):
            s, e = outage_dates(fy, f)
            ak = (fy, f) in ((2022, 1), (2023, 2))
            if (fy, f) == (2018, 1): its = from_jobs(jobs.H30, 1) + [x for x in furnace_items(fy, f, s, e) if x['no'].startswith('E')]
            elif (fy, f) == (2021, 1): its = from_jobs(jobs.R3, 1) + [x for x in furnace_items(fy, f, s, e) if x['no'].startswith('E')]
            else: its = furnace_items(fy, f, s, e, ak=ak)
            if f == 2: its += turbine_items(fy, s)
            p = dict(key=f'{fy}-{f}', fy=fy, kind=f'{f}号定修', contract=ck, part=f'{f}号炉 定期整備', start=s, end=e, items=its, other=[],
                     milestones=milestones_for(its, s, e), stop=stop_for(s, e, its), groups=None, book=1 if f == 2 else 2)
            if (fy, f) == (2018, 1): p['milestones'] = jobs.H30['milestones']
            if (fy, f) == (2021, 1): p['milestones'] = jobs.R3['milestones']; p['stop'] = jobs.R3['stop']
            p['groups'] = groups_for(its)
            if fy == 2026 and f == 1: p['status'] = '施工中'
            P.append(p)
    # 実績の打ち切り（2026 年 1 号は 11/13 まで）
    for p in P:
        for it in p['items']:
            if it['act'] and it['act'][0] > TODAY: it['act'] = None
            elif it['act'] and it['act'][1] > TODAY: it['act'] = (it['act'][0], TODAY); it['note'] = (it['note'] + ' 施工中').strip()
    return P

# ------------------------------------------------------------------ 補修工事（請負、千草）。発注 CSV と工程表・日報・報告書
REPAIRS = [
    dict(key='R2012', fy=2012, name=f'平成24年度 {FACN} 2号ボイラエコノマイザ管漏れ補修工事', start=D(2013, 2, 12), end=D(2013, 2, 22),
         cs=D(2013, 2, 8), ce=D(2013, 3, 22), amount=9_660_000, units=[(2, 'エコノマイザ')], qa='q1',
         items=[('1', '2号炉 緊急停止・冷却（市）', 'エコノマイザ', CITY, 0, 0, 2), ('2', 'エコノマイザ 開放・漏れ箇所特定', 'エコノマイザ', CH, 3, 3, 4),
                ('3', '漏れ管 切替え（2本）・溶接', 'エコノマイザ', CH, 4, 5, 7), ('4', 'RT 検査・水圧試験', 'エコノマイザ', CH, 3, 8, 8),
                ('5', '復旧・立上げ（市）', 'エコノマイザ', CITY, 0, 9, 10)]),
    dict(key='R2014', fy=2014, name=f'平成26年度 {FACN} 2号溶融炉出湯口再築工事', in_=(2014, 2), units=[(2, '出湯口')], qa='q1', amount=31_320_000,
         items=[('1', '出湯口 解体・はつり', '出湯口', SN, 4, 8, 11), ('2', '出湯口れんが築造', '出湯口', SN, 4, 14, 20),
                ('3', 'キャスタブル施工・養生', '出湯口', SN, 3, 21, 23), ('4', '開孔・出湯確認', '出湯口', CH, 2, 31, 31)]),
    dict(key='R2017', fy=2017, name=f'平成29年度 {FACN} 排ガス分析計更新工事', in_=None, units='analyzer', qa='q1', amount=58_320_000,
         items=None),
    dict(key='R2020', fy=2020, name=f'令和2年度 {FACN} 2号溶融炉出湯口再築工事', in_=(2020, 2), units=[(2, '出湯口')], qa='q1', amount=34_100_000,
         items=[('1', '出湯口 解体・はつり', '出湯口', SN, 4, 8, 11), ('2', '出湯口れんが築造', '出湯口', SN, 4, 14, 19),
                ('3', 'キャスタブル施工・養生', '出湯口', SN, 3, 20, 22), ('4', '開孔・出湯確認', '出湯口', CH, 2, 31, 31)]),
    dict(key='R2021', fy=2021, name=f'令和3年度 {FACN} 1号溶融炉出湯口補修工事', sample=True, units=[(1, '出湯口')], qa='q1', amount=38_610_000),
    dict(key='R2024', fy=2024, name=f'令和6年度 {FACN} 1号ボイラ2次過熱器下部ベンド部分更新工事', in_=(2024, 1), units=[(1, '2次過熱器')], qa='q1', amount=27_830_000,
         items=[('1', '2次過熱器 下部ベンド 管切断・撤去（4本）', '2次過熱器', CH, 4, 14, 15), ('2', '新管 取付・溶接', '2次過熱器', CH, 4, 16, 19),
                ('3', 'RT 検査', '2次過熱器', CH, 2, 20, 20), ('4', '水圧試験（年間保守と同時）', '2次過熱器', CH, 2, 24, 24)]),
]
for c in REPAIRS:
    if c.get('in_'):
        c['in'] = [(c['in_'][0], c['in_'][1], [b for (_, b) in c['units']])]
    elif c['key'] == 'R2012':
        c['in'] = [(2012, 2, ['エコノマイザ'])]
    elif c['key'] == 'R2021':
        c['in'] = [(2021, 1, ['出湯口'])]
    elif c['key'] == 'R2017':
        c['in'] = [(2017, 2, ['__analyzer__']), (2017, 1, ['__analyzer__'])]

# ------------------------------------------------------------------ 他の業者の作業（市が直接発注。資料は財務 CSV だけ）
def other_vendor(u, fy):
    h = H.truth(u)
    if fy not in h: return None
    s, d = h[fy]
    if u['cat'] in ('crane', 'crane2') or (u['furn'] == 0 and 'クレーン' in u['name'] and u['kind'] in ('機器', '電動機', '盤')):
        if u['kind'] == '電動機' and s == '○': return None
        return 'あかね台クレーン整備株式会社'
    if u['base'] in ('空冷復水器ファン（C）', '空冷復水器ファン（E）') and (fy, u['base']) in ((2019, '空冷復水器ファン（C）'), (2023, '空冷復水器ファン（E）')):
        return '東和電機サービス株式会社'
    if u['cat'] == 'scale': return 'さくら計量器株式会社'
    if u['cat'] == 'egen': return '東和電機サービス株式会社'
    return None

# ------------------------------------------------------------------ 作業（正解）
WORKS = []
def add_work(tag, sym, s, e, contract, vendor, pkey, ino, desc=''):
    w = dict(id=f'W{len(WORKS)+1:05d}', unit=tag, type=SYMW.get(sym, 'INSPECT'), sym=sym, start=s, end=e, contract=contract,
             vendor=vendor, period=pkey, item=ino, manHours=None, desc=desc)
    WORKS.append(w); return w['id']

def assign_works(P, contract_vendor=lambda c: '千草環境システム株式会社'):
    bytag = {u['tag']: u for u in world.UNITS}
    for p in P:
        for it in p['items']:
            it['wids'] = []
            if it['co'] == CITY or not it['units']: continue
            rng = it['act'] or it['plan']
            if not rng: continue
            for t in it['units']:
                u = bytag[t]; h = H.truth(u)
                sym = it.get('sym') or (h.get(p['fy'], ('○', ''))[0])
                if sym == '★' and p['contract'] != '基幹改良': sym = '○'
                end = rng[1] if it['act'] and not ('施工中' in (it['note'] or '')) else (None if p.get('status') == '施工中' else rng[1])
                it['wids'].append(add_work(t, sym, rng[0], end, p['contract'], contract_vendor(p['contract']), p['key'], it['no'],
                                           h.get(p['fy'], ('', ''))[1]))

# ------------------------------------------------------------------ 基幹的設備改良工事（令和3〜6年度、千草・現場代理人 q2）
AKC = 'あかね台資源化センター 基幹的設備改良工事（令和3〜6年度）'
def ak_units(fy, pred=lambda u: True):
    return [u['tag'] for u in world.UNITS if world.ak_scope(u) and world.ak_scope(u)[0] == fy and pred(u)] + \
           [u['tag'] for u in world.UNITS if u.get('ident') == 'ak_new' and fy == (2023 if u['sys'] == '03' else 2024) and pred(u)]

def ak_periods():
    P = []
    # R3：設計と 1号定修での先行作業（見本の K-1〜K-3）
    its = []
    for ck, no, name, base, co, ppl, pl, ac, note in jobs.R3['other']:
        if ck != '基幹改良': continue
        its.append(mk(no, name, [unit(1, base)['tag']], co.replace('千草環境（改良工事）', CH), '改良', ppl, pl, ac, note, 'ak', '○'))
    P.append(dict(key='AK-2021', fy=2021, kind='基幹改良', contract='基幹改良', part='令和3年度 施工（1号炉定修時の先行作業）', start=D(2021, 10, 4), end=D(2021, 10, 28),
                  items=its, other=[], milestones=[(D(2021, 10, 5), '着手')], stop=[(D(2021, 10, 4), D(2021, 10, 28), '1号炉 停止中（年間保守の定修）')], groups=None))
    def furn_ak(fy, f, s, e):
        L = (e - s).days + 1
        tag = lambda b: unit(f, b)['tag']
        mot = [u['tag'] for u in world.UNITS if u['furn'] == f and u['base'] == '羽口冷却水ポンプ用電動機']
        pumps = [u['tag'] for u in world.UNITS if u['furn'] == f and u['base'] == '羽口冷却水ポンプ']
        X = [('A1', '足場組立（炉内全段）', ['溶融炉本体'], MT, 6, 8, 12), ('A2', '旧耐火物 解体（シャフト・朝顔・炉床）', ['溶融炉本体'], SN, 10, 13, 30),
             ('A3', '炉体冷却ジャケット 撤去・新設', ['炉体冷却ジャケット'], CH, 8, 25, 45), ('A4', '鉄皮 補修（肉盛・当て板）', ['溶融炉本体'], CH, 4, 20, 35),
             ('A5', '羽口冷却水ポンプ・電動機 更新', None, CH, 4, 30, 36), ('A6', '主羽口・副羽口 装置更新', ['主羽口装置', '副羽口装置'], CH, 6, 40, 52),
             ('A7', '炉体温度計 更新（24点）', ['炉体温度計'], AD, 3, 50, 56), ('A8', '築炉（炉床・朝顔・シャフト）', ['溶融炉本体'], SN, 10, 46, 78)]
        if f == 2: X.append(('A8-2', '出湯口 更新（炉床と一体）', ['出湯口'], SN, 4, 60, 72))
        X += [('A9', '冷却水系 耐圧試験', ['炉体冷却ジャケット'], CH, 4, 79, 80), ('A10', '足場解体', ['溶融炉本体'], MT, 6, 81, 84),
              ('A11', '炉乾燥・立上げ（市運転、改良工事立会）', ['溶融炉本体'], CITY, 0, 88, 102), ('A12', '性能確認運転', ['溶融炉本体'], CH, 3, 103, L - 1)]
        its = []
        for no, nm, bases, co, ppl, a, b in X:
            us = pumps + mot if bases is None else [tag(bb) for bb in bases]
            A = s + dt.timedelta(a); Bd = s + dt.timedelta(min(b, L - 1))
            act = (A, Bd + dt.timedelta(2)) if no == 'A8' and f == 1 else (A, Bd)
            note = '築炉材の入荷遅れで 2日延長' if (no == 'A8' and f == 1) else ''
            its.append(mk(no, nm, us, co, '改良', ppl, (A, Bd), act, note, 'ak', '★'))
        return its
    s, e = D(2022, 9, 5), D(2022, 12, 22)
    P.append(dict(key='AK-2022', fy=2022, kind='基幹改良', contract='基幹改良', part='令和4年度 施工（1号炉 炉体・羽口更新）', start=s, end=e, items=furn_ak(2022, 1, s, e), other=[],
                  milestones=[(s, '1号炉停止'), (s + dt.timedelta(46), '築炉着手'), (s + dt.timedelta(88), '乾燥'), (e, '引渡し')], stop=[(s, s + dt.timedelta(87), '1号炉 停止・改良工事'), (s + dt.timedelta(88), e, '乾燥・立上げ・性能確認')], groups=None))
    s, e = D(2023, 5, 29), D(2023, 9, 14)
    P.append(dict(key='AK-2023-2', fy=2023, kind='基幹改良', contract='基幹改良', part='令和5年度 施工（2号炉 炉体・羽口・出湯口更新）', start=s, end=e, items=furn_ak(2023, 2, s, e), other=[],
                  milestones=[(s, '2号炉停止'), (s + dt.timedelta(46), '築炉着手'), (s + dt.timedelta(88), '乾燥'), (e, '引渡し')], stop=[(s, s + dt.timedelta(87), '2号炉 停止・改良工事'), (s + dt.timedelta(88), e, '乾燥・立上げ・性能確認')], groups=None))
    s, e = D(2023, 6, 12), D(2023, 7, 28)
    psa = ak_units(2023, lambda u: u['sys'] == '03')
    byb = lambda pred: [t for t in psa if pred(next(u for u in world.UNITS if u['tag'] == t))]
    X = [('P1', '液体酸素への切替（市運転）', psa[:1], CITY, 0, 0, 0), ('P2', '旧PSA 撤去', byb(lambda u: u['kind'] in ('機器', '電動機')), CH, 6, 1, 8),
         ('P3', '基礎改修', psa[:1], CH, 4, 9, 14), ('P4', '機器据付（ブロワ・真空ポンプ・吸着塔・製品槽・圧縮機・冷却器）', byb(lambda u: u['kind'] in ('機器', '電動機')), CH, 8, 15, 28),
         ('P5', '配管・切替弁 据付', byb(lambda u: u['kind'] in ('調節弁', '電動弁')), CH, 5, 24, 34),
         ('P6', '制御盤・インバータ盤・濃度計 据付・配線', byb(lambda u: u['kind'] in ('盤', '計装')), AD, 4, 28, 38),
         ('P7', '試運転・調整（製品酸素 濃度・流量）', psa[:1], CH, 3, 39, 45), ('P8', 'PSA へ切替（市運転）', psa[:1], CITY, 0, 46, 46)]
    its = [mk(no, nm, us, co, '改良', ppl, (s + dt.timedelta(a), s + dt.timedelta(b)), (s + dt.timedelta(a), s + dt.timedelta(b)), '', 'ak', '★') for no, nm, us, co, ppl, a, b in X]
    P.append(dict(key='AK-2023-P', fy=2023, kind='基幹改良', contract='基幹改良', part='令和5年度 施工（酸素発生装置 更新）', start=s, end=e, items=its, other=[],
                  milestones=[(s, '液酸切替'), (s + dt.timedelta(39), '試運転'), (e, 'PSA切替')], stop=[(s, e, 'PSA 停止（液体酸素で運転）')], groups=None))
    s, e = D(2024, 5, 7), D(2024, 5, 24)
    dcs = ak_units(2024)
    f1 = [t for t in dcs if t.startswith('1-')]; f2 = [t for t in dcs if t.startswith('2-')]; c0 = [t for t in dcs if t.startswith('0-')]
    X = [('D1', '全炉停止（市運転）', c0[:1], CITY, 0, 0, 0), ('D2', '旧DCS 撤去・新盤据付（OPS・EWS・FCS-C）', c0, CH, 6, 1, 5),
         ('D3', '1号炉 FCS・I/O 切替', f1, AD, 5, 3, 8), ('D4', '2号炉 FCS・I/O 切替', f2, AD, 5, 6, 11), ('D5', '共通 I/O 切替', c0, AD, 4, 8, 12),
         ('D6', 'ループチェック', c0[:1] + f1[:1] + f2[:1], CH, 6, 10, 15), ('D7', 'インターロック試験（全系統）', c0[:1], CH, 4, 15, 16), ('D8', '立上げ（市運転）', c0[:1], CITY, 0, 17, 17)]
    its = [mk(no, nm, us, co, '改良', ppl, (s + dt.timedelta(a), s + dt.timedelta(b)), (s + dt.timedelta(a), s + dt.timedelta(b)), '', 'ak', '★') for no, nm, us, co, ppl, a, b in X]
    its[5]['act'] = (its[5]['act'][0], its[5]['act'][1] + dt.timedelta(1)); its[5]['note'] = '2号 I/O の不具合 2点で 1日延長'
    P.append(dict(key='AK-2024', fy=2024, kind='基幹改良', contract='基幹改良', part='令和6年度 施工（DCS 更新・切替）', start=s, end=e, items=its, other=[],
                  milestones=[(s, '全停'), (s + dt.timedelta(10), 'ループチェック'), (e, '立上げ')], stop=[(s, e, '全炉停止（改良工事で延長）')], groups=None))
    return P

def repair_periods(Pmain):
    by = {p['key']: p for p in Pmain}
    out = []
    for c in REPAIRS:
        if c['key'] == 'R2021':
            its = [mk(o[1], o[2], [unit(1, o[3])['tag']], o[4], '補修', o[5], o[6], o[7], o[8], 'tap', '◎') for o in jobs.R3['other'] if o[0] == 'R3出湯口']
            out.append(dict(key='R2021', fy=2021, kind='補修工事', contract='R2021', part='', start=D(2021, 10, 5), end=D(2021, 11, 6), items=its, other=[],
                            milestones=[(D(2021, 10, 5), '着手'), (D(2021, 10, 22), '養生完了'), (D(2021, 11, 6), '出湯確認')], stop=jobs.R3['stop'], groups=None, title=c['name']))
            continue
        if c['key'] == 'R2017':
            for f in (2, 1):
                p0 = by[f'2017-{f}']; s = p0['start']; L = (p0['end'] - s).days + 1
                an = [u['tag'] for u in world.UNITS if u['furn'] == f and u['cat'] == 'analyzer']
                its = [mk('1', '排ガス分析計 撤去・据付・配線', an, AD, '電計', 4, (s + dt.timedelta(scale(14, L)), s + dt.timedelta(scale(19, L))), (s + dt.timedelta(scale(14, L)), s + dt.timedelta(scale(19, L))), '', 'ev_EL', '●'),
                       mk('2', '校正・試運転（標準ガス）', an, AD, '電計', 2, (s + dt.timedelta(scale(33, L)), s + dt.timedelta(scale(36, L))), (s + dt.timedelta(scale(33, L)), s + dt.timedelta(scale(36, L))), '', 'cal', '●')]
                out.append(dict(key=f'R2017-{f}', fy=2017, kind='補修工事', contract='R2017', part=f'{f}号炉', start=s + dt.timedelta(scale(14, L)), end=s + dt.timedelta(scale(36, L)),
                                items=its, other=[], milestones=[(s + dt.timedelta(scale(14, L)), '着手')], stop=[], groups=None, title=c['name']))
            continue
        if c['key'] == 'R2012':
            s = c['start']
            its = [mk(no, nm, [unit(2, b)['tag']], co, '補修', ppl, (s + dt.timedelta(a), s + dt.timedelta(b_)), (s + dt.timedelta(a), s + dt.timedelta(b_)), '', 'ev_ST', '△')
                   for no, nm, b, co, ppl, a, b_ in c['items']]
            out.append(dict(key='R2012', fy=2012, kind='補修工事', contract='R2012', part='2号炉', start=s, end=c['end'], items=its, other=[],
                            milestones=[(s, '緊急停止')], stop=[], groups=None, title=c['name']))
            continue
        fy, f = c['in_']; p0 = by[f'{fy}-{f}']; s = p0['start']; L = (p0['end'] - s).days + 1
        sym = H.truth(unit(f, c['units'][0][1])).get(fy, ('△', ''))[0]
        its = [mk(no, nm, [unit(f, b)['tag']], co, '補修', ppl, (s + dt.timedelta(scale(a, L)), s + dt.timedelta(scale(b_, L))),
                  (s + dt.timedelta(scale(a, L)), s + dt.timedelta(scale(b_, L))), '', 'tap' if '出湯口' in nm else 'ev_ST', sym) for no, nm, b, co, ppl, a, b_ in c['items']]
        out.append(dict(key=c['key'], fy=fy, kind='補修工事', contract=c['key'], part=f'{f}号炉', start=its[0]['act'][0], end=its[-1]['act'][1], items=its, other=[],
                        milestones=[(its[0]['act'][0], '着手'), (its[-1]['act'][1], '完了')], stop=[], groups=None, title=c['name']))
    return out
