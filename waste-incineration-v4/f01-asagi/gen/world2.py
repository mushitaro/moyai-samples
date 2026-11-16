# -*- coding: utf-8 -*-
"""本作り用の世界（1997〜2026 年度）。工程（タスク）を先に作り、日報・報告書・発注はここから作る。"""
import datetime as dt, hashlib, random
from f01_master import EQ, SYS, TODAY, wareki
from f01_history import BYTAG, HOSHU, KIKAN
from f01_contracts import (H23, R7, H23_F1, H23_F2, R7_F1, R7_F2, R7_CHANGES, CONTRACTS_H, MAKER, OPCO, work_days, T,
                           teishu_amount, hoshu_amount)

D = dt.date
YEARS = list(range(1997, 2027))

def h(*a):
    return int(hashlib.md5('|'.join(map(str, a)).encode()).hexdigest()[:8], 16)
def R(*a):
    return random.Random(h(*a))

def monday_on_or_after(x):
    return x + dt.timedelta(days=(7 - x.weekday()) % 7)

# ------------------------------------------------------------ 人（担当と期間）
def city_person(fy):
    return 'a1' if fy <= 2007 else ('a2' if fy <= 2020 else 'a3')
def agent(fy):
    return 'm1' if fy <= 2009 else ('m2' if fy <= 2018 else 'm3')
def report_writer(fy):
    return 'r1' if fy <= 2014 else 'r2'
def operator_word(fy):
    return '運転委託先' if fy >= 2006 else '市運転班'

# ------------------------------------------------------------ 様式（プロファイルの年）
def sched_style(fy):
    return 'hand' if fy <= 2003 else ('m1' if fy <= 2011 else 'm2')
def daily_style(fy):
    return 'hand' if fy <= 2003 else ('m1d' if fy <= 2010 else ('m2' if fy <= 2020 else 'm3'))
def report_style(fy):
    return 'r1' if fy <= 2015 else 'r2'
def order_style(fy):
    return 'c1' if fy <= 2008 else ('old' if fy <= 2019 else 'new')

# ------------------------------------------------------------ 定修の期間（炉別）
SPECIAL_LONG = {  # 改造工事で炉を長く止めた年度：定修はその停止の頭で行う
    (2000, 1): (D(2000, 9, 4), D(2001, 3, 16)), (2001, 2): (D(2001, 9, 3), D(2002, 2, 22)),
    (2017, 1): (D(2017, 9, 4), D(2018, 2, 23)), (2018, 2): (D(2018, 9, 3), D(2019, 2, 22)),
}
def windows(fy):
    """{1: (停止日, 完了日), 2: (...), 'Z': (全停 開始, 終了)}"""
    if fy == 2011:
        return {1: H23['f1'], 2: H23['f2'], 'Z': H23['zentei']}
    if fy == 2025:
        return {1: R7['f1'], 2: R7['f2'], 'Z': R7['zentei']}
    r = R('win', fy)
    w = {}
    s1 = monday_on_or_after(D(fy, 5, 25) + dt.timedelta(days=7 * r.randint(0, 4)))
    s2 = monday_on_or_after(D(fy, 9, 20) + dt.timedelta(days=7 * r.randint(0, 4)))
    if fy == 2026:
        s1, s2 = D(2026, 6, 1), D(2026, 10, 5)
    w[1] = (s1, s1 + dt.timedelta(days=32))
    w[2] = (s2, s2 + dt.timedelta(days=32 if fy != 2026 else 46))
    for (y, f), per in SPECIAL_LONG.items():
        if y == fy:
            w[f] = per
            o = 2 if f == 1 else 1
            so = monday_on_or_after(D(fy, 6, 1) + dt.timedelta(days=7 * r.randint(0, 2)))
            w[o] = (so, so + dt.timedelta(days=32))
    # 全停は 1号炉（長期停止の年度は 運転している側の定修）期間の 3 週目
    base = w[1] if (fy, 1) not in SPECIAL_LONG else w[2]
    zs = base[0] + dt.timedelta(days=15)
    zlen = 3 if fy not in (2019, 2012) else (9 if fy == 2019 else 5)
    w['Z'] = (zs, zs + dt.timedelta(days=zlen - 1))
    return w

# ------------------------------------------------------------ 班の割り当て
def team_of(d):
    k = d['kind']
    if k in ('炉体', '火格子', '給じん装置', 'ホッパ', 'ゲート', '灰押出装置', 'バーナ', '油圧装置', 'ピット', '投入扉', '扉', '破砕機'):
        return 'A'
    if k in ('盤', '変圧器', '監視装置', '電動弁', '発電機', '空調機', '計量機'):
        return 'D'
    if k in ('ボイラ', '減温塔', 'バグフィルタ', '煙道', '煙突', 'スートブロワ', '空気作動弁', '空気予熱器', 'タンク', 'サイロ', '水槽', '蒸気だめ',
             '脱気器', '熱交換器', '空冷復水器', '減圧減温装置', '集じん器', '吸着塔', 'ろ過装置', '水処理装置', '冷却塔'):
        return 'B'
    return 'C'

ZENTEI_SYS = ('J', 'L')
def needs_zentei(d):
    return d['f'] == 0 and (d['sys'] in ZENTEI_SYS or d['tag'] in ('C-02', 'C-03', 'SL-03', 'EWS-01', 'DCS-01'))

def task_days(cost, kind):
    if kind in ('点検', '計測'):
        return 1 if cost < 600 else 2
    return max(1, min(8, 1 + int(cost / 2200)))

def events_for(fy, f=None, cts=('T',)):
    out = []
    for d in EQ:
        if f is not None and d['f'] != f:
            continue
        evs = list(d['act'].get(fy, []))
        if fy == 2026 and d.get('pending2026'):
            evs += [(k, t, c, 'T') for (k, t, c) in d['pending2026']]
        for i, (k, t, c, ct) in enumerate(evs):
            if ct in cts:
                out.append((d, k, t, c, ct, i))
    return out

def _wd_add(x, n, allow_sun=False):
    """作業日を n 日進める（日曜休工）"""
    while n > 0:
        x += dt.timedelta(days=1)
        if x.weekday() != 6 or allow_sun:
            n -= 1
    return x

def _first_wd(x):
    while x.weekday() == 6:
        x += dt.timedelta(days=1)
    return x

def gen_furnace_tasks(fy, f, s, e, pre):
    """1 炉の定修タスク（段取り＋機器作業）。pre はタスク No の接頭辞"""
    op = operator_word(fy)
    L = []
    no = [0]
    def nx():
        no[0] += 1
        return '%s%02d' % (pre, no[0])
    grp = '%d号炉' % f
    L.append(T(nx(), grp, '準備工（資機材搬入・仮設電源・新規入場者教育）', s - dt.timedelta(days=3), s + dt.timedelta(days=2), 'A', (1, 3)))
    L.append(T(nx(), grp, '焼却停止・立下げ（%s操作）' % op, s, s, '運転', (0, 0), 'stop'))
    L.append(T(nx(), grp, '炉冷却（自然冷却・誘引通風機運転）', s, s + dt.timedelta(days=2), '運転', (0, 0)))
    o = _first_wd(s + dt.timedelta(days=3))
    L.append(T(nx(), grp, '安全措置（O2・CO測定、電源ロック）・マンホール開放', o, o, 'A', (2, 2)))
    L.append(T(nx(), grp, '足場組立（炉内・ボイラ）', o, _wd_add(o, 1), 'A', (1, 6)))
    L.append(T(nx(), grp, '炉内清掃・クリンカ除去', o, _wd_add(o, 2), 'A', (1, 7)))
    L.append(T(nx(), grp, 'ボイラ清掃（水管付着ダスト除去）', _wd_add(o, 1), _wd_add(o, 5), 'B', (1, 5)))
    w0 = _wd_add(o, 3)
    evs = [x for x in events_for(fy, f) if x[4] == 'T']
    bys = {}
    for x in evs:
        bys.setdefault(team_of(x[0]), []).append(x)
    lastday = w0
    for team in 'ABCD':
        cur = w0
        items = sorted(bys.get(team, []), key=lambda x: (x[0]['sys'], x[0]['code']))
        i = 0
        while i < len(items):
            chunk = [items[i]]
            # 同じ設備の小物（点検・計測）は 1 タスクにまとめる（2〜3 台）
            while i + 1 < len(items) and len(chunk) < 3 and items[i + 1][1] in ('点検', '計測') and chunk[0][1] in ('点検', '計測') \
                    and items[i + 1][0]['sys'] == chunk[0][0]['sys']:
                i += 1
                chunk.append(items[i])
            i += 1
            dd = max(task_days(x[3], x[1]) for x in chunk)
            names = '・'.join(dict.fromkeys(x[0]['name'] for x in chunk))
            work = '・'.join(dict.fromkeys(x[2].replace('（基幹改良）', '') for x in chunk))
            item = '%s %s' % (names, work) if len(chunk) == 1 else '%s（%s）' % (names, work)
            ss = _first_wd(cur)
            ee = _wd_add(ss, dd - 1)
            n = (1, 1 if dd <= 2 else (2 if dd <= 5 else 4))
            if team == 'A' and dd >= 4:
                n = (1, 5)
            L.append(T(nx(), grp, item, ss, ee, team, n, tags=tuple(x[0]['tag'] for x in chunk)))
            L[-1]['evs'] = [(x[0]['tag'], x[5]) for x in chunk]
            cur = _wd_add(ss, max(1, dd // 2)) if R('ov', fy, f, team, i).random() < 0.5 else _wd_add(ee, 1)
            lastday = max(lastday, ee)
    lim = e - dt.timedelta(days=11)
    if lastday > lim:   # 作業が多い年は工程に収まるよう圧縮しない：完了日を延ばす
        e = lastday + dt.timedelta(days=11)
    t1 = _wd_add(max(lastday, w0), 1)
    L.append(T(nx(), grp, '足場解体・炉内片付け', t1, _wd_add(t1, 2), 'A', (1, 6)))
    t2 = _wd_add(t1, 3)
    L.append(T(nx(), grp, 'マンホール閉鎖・単体試運転', t2, _wd_add(t2, 2), 'A', (3, 3)))
    st = e - dt.timedelta(days=4)
    L.append(T(nx(), grp, '立上げ（昇温・耐火物乾燥）', st, st + dt.timedelta(days=1), '運転', (1, 0), 'start'))
    L.append(T(nx(), grp, '試運転（負荷運転）・性能確認', st + dt.timedelta(days=2), e, '運転', (2, 0)))
    return L, e

def gen_common_tasks(fy, wins, pre, part):
    """共通設備のタスク。part=1: 1号炉期間（全停含む）、part=2: 2号炉期間"""
    evs = [x for x in events_for(fy, 0) if x[4] == 'T']
    L = []
    no = [0]
    def nx():
        no[0] += 1
        return '%s%02d' % (pre, no[0])
    zs, ze = wins['Z']
    zpart = 1 if not ((fy, 1) in SPECIAL_LONG) else 2
    if part == zpart:
        L.append(T(nx(), '全停', '%s 立下げ（全停準備）' % ('2号炉' if zpart == 1 else '1号炉'), zs - dt.timedelta(days=1), zs - dt.timedelta(days=1), '運転', (0, 0), 'stop'))
        zl = [x for x in evs if needs_zentei(x[0])]
        byt = {}
        for x in zl:
            byt.setdefault(team_of(x[0]), []).append(x)
        for team in 'BCD':
            items = byt.get(team, [])
            for j in range(0, len(items), 4):
                chunk = items[j:j + 4]
                names = '・'.join(dict.fromkeys(x[0]['name'] for x in chunk))
                L.append(T(nx(), '全停', '%s 点検・整備' % names if len(names) < 40 else '%s ほか %d 台 点検・整備' % (chunk[0][0]['name'], len(chunk) - 1),
                           zs, ze, team, (2, 2) if team == 'D' else (1, 2), tags=tuple(x[0]['tag'] for x in chunk)))
                L[-1]['evs'] = [(x[0]['tag'], x[5]) for x in chunk]
        L.append(T(nx(), '全停', '%s 立上げ' % ('2号炉' if zpart == 1 else '1号炉'), ze + dt.timedelta(days=1), ze + dt.timedelta(days=1), '運転', (0, 0), 'start'))
    rest = [x for x in evs if not needs_zentei(x[0]) and (h('cp', fy, x[0]['tag']) % 2 + 1) == part]
    f = part if (fy, part) not in SPECIAL_LONG else part
    s, e = wins[part]
    cur = _first_wd(s + dt.timedelta(days=8))
    for x in sorted(rest, key=lambda x: x[0]['code']):
        dd = task_days(x[3], x[1])
        ss = _first_wd(cur)
        ee = _wd_add(ss, dd - 1)
        L.append(T(nx(), '共通', '%s %s' % (x[0]['name'], x[2]), ss, ee, team_of(x[0]), (1, 1 if dd <= 2 else 3), tags=(x[0]['tag'],)))
        L[-1]['evs'] = [(x[0]['tag'], x[5])]
        cur = _wd_add(ss, 1)
        if cur > e - dt.timedelta(days=12):
            cur = _first_wd(s + dt.timedelta(days=9))
    return L

# ------------------------------------------------------------ 契約
def place(fy, d):
    """補修工事を施工する停止期間：炉の機器はその炉の定修、全停が要る共通機器は全停、その他の共通は 2号炉期間"""
    w = windows(fy)
    if d['f'] in (1, 2):
        return d['f'], w[d['f']]
    if needs_zentei(d):
        return 'Z', w['Z']
    return 2, w[2]

def hoshu_contracts(fy):
    """補修工事（請負）。見本の 2 年度は承認済みの契約をそのまま使う"""
    if fy in CONTRACTS_H:
        out = []
        for c in CONTRACTS_H[fy]:
            c = dict(c)
            out.append(c)
        return out
    out = []
    for (code, k, t, c) in HOSHU.get(fy, []):
        d = next(e for e in EQ if e['code'] == code)
        vendor = 'あさぎ建設株式会社' if d['prof'] in ('pit', 'stack_out') else ('港東築炉工業株式会社' if d['prof'] == 'furnace' else MAKER)
        s0, e0 = place(fy, d)[1]
        title = '%s あさぎ市清掃工場%s%s%s工事' % (wareki(fy, False), '' if d['f'] == 0 else d['fname'], d['name'].replace('（', '').replace('）', ''),
                                             '補修' if k != '更新' else '更新')
        cdate = s0 - dt.timedelta(days=30 + h('cd', fy, code) % 25)
        out.append(dict(title=title, tag=d['tag'], vendor=vendor, method='指名競争入札', cdate=cdate, s=cdate + dt.timedelta(days=3),
                        e_=e0 + dt.timedelta(days=25)))
    return out

def hoshu_tasks(fy, c):
    """補修工事の工程（自社の工程表用）"""
    d = BYTAG[c['tag']]
    k, t, cost, ct = next(x for x in d['act'].get(fy, []) if x[3] == 'H')
    key, (s0, e0) = place(fy, d)
    # 定修の工程表の「別契約」行と同じ期間にする
    st = c.get('_s') or (_first_wd(s0) if key == 'Z' else _first_wd(_wd_add(_first_wd(s0 + dt.timedelta(days=3)), 3)))
    dd = max(3, min(12, 2 + int(cost / 1500)))
    en = _wd_add(st, dd - 1)
    if c.get('_e'):
        en = c['_e']
        dd = max(3, len([1 for i in range((en - st).days + 1) if (st + dt.timedelta(days=i)).weekday() != 6]))
    c['_s'], c['_e'] = st, en
    nm = d['name']
    L = [T('1', '準備', '施工計画書提出・資材手配', c['s'], c['s'] + dt.timedelta(days=10), '事務', (1, 0)),
         T('2', '準備', '資機材搬入・養生', st - dt.timedelta(days=1), st - dt.timedelta(days=1), 'A', (1, 2))]
    a = max(1, dd // 4)
    L.append(T('3', '施工', '%s 撤去・はつり（劣化部）' % nm if k == '補修' else '%s 既設撤去' % nm, st, _wd_add(st, a - 1), 'A', (1, 3), tags=(d['tag'],)))
    b0 = _wd_add(st, a)
    L.append(T('4', '施工', '%s %s' % (nm, t), b0, _wd_add(b0, max(1, dd - a - 2) - 1), 'A', (1, 4), tags=(d['tag'],)))
    c0 = _wd_add(b0, max(1, dd - a - 2))
    L.append(T('5', '検査', '出来形・寸法検査（市立会）', c0, c0, 'A', (2, 1), tags=(d['tag'],)))
    L.append(T('6', '片付け', '復旧・清掃・搬出', _wd_add(c0, 1), _wd_add(c0, 1), 'A', (1, 2)))
    for x in L:
        x['evs'] = [(d['tag'], 0)] if x['tags'] else []
    return L

def special_contracts(fy):
    """改造工事（DXN 対策・DCS 更新・放水銃設置・基幹改良）"""
    out = []
    if fy in (2000, 2001):
        f = 1 if fy == 2000 else 2
        s, e = SPECIAL_LONG[(fy, f)]
        out.append(dict(kind='X', title='あさぎ市清掃工場 ダイオキシン類恒久対策工事（%s・%d号炉）' % (wareki(fy, False), f), vendor=MAKER, f=f, s=s, e_=e,
                        cdate=D(fy, 6, 23) if fy == 2000 else D(2000, 6, 23), method='随意契約（継続費）'))
    if fy == 2010:
        out.append(dict(kind='N', title='平成22年度 あさぎ市清掃工場 ごみピット放水銃設置工事', vendor='あさぎ防災設備株式会社', f=0, s=D(2010, 11, 1), e_=D(2011, 2, 25),
                        cdate=D(2010, 10, 22), method='指名競争入札'))
    if fy == 2012:
        out.append(dict(kind='D', title='平成24年度 あさぎ市清掃工場 分散制御システム更新工事', vendor=MAKER, f=0, s=D(2012, 9, 3), e_=D(2013, 2, 28),
                        cdate=D(2012, 8, 24), method='随意契約'))
    if fy in (2017, 2018, 2019):
        per = {2017: SPECIAL_LONG[(2017, 1)], 2018: SPECIAL_LONG[(2018, 2)], 2019: (D(2019, 6, 3), D(2019, 12, 20))}[fy]
        out.append(dict(kind='K', title='あさぎ市清掃工場 基幹的設備改良工事（%s施工分）' % wareki(fy, False), vendor=MAKER, f={2017: 1, 2018: 2, 2019: 0}[fy],
                        s=per[0], e_=per[1], cdate=D(2017, 6, 30), method='随意契約（継続費 H29〜R1）'))
    return out

def special_tasks(fy, c):
    ct = c['kind']
    evs = [(d, k, t, cst, x, i) for (d, k, t, cst, x, i) in events_for(fy, None, (ct,))]
    s, e = c['s'], c['e_']
    L = [T('1', '準備', '施工計画・製作図承認・機器製作', max(c['cdate'] + dt.timedelta(days=7), s - dt.timedelta(days=60)), s - dt.timedelta(days=1), '事務', (2, 0)),
         T('2', '準備', '仮設・養生・既設撤去準備', s, _wd_add(s, 4), 'A', (1, 4))]
    cur = _wd_add(s, 5)
    byg = {}
    for x in evs:
        byg.setdefault((x[0]['sys'], x[0]['f']), []).append(x)
    no = 3
    for key in sorted(byg):
        items = byg[key]
        for j in range(0, len(items), 4):
            chunk = items[j:j + 4]
            dd = 6 + int(sum(x[3] for x in chunk) / 5000)
            dd = min(dd, 45)
            names = '・'.join(dict.fromkeys(x[0]['name'] for x in chunk))
            work = chunk[0][2].replace('（基幹改良）', '').replace('（ダイオキシン類恒久対策）', '').replace('（DCS更新工事）', '')
            L.append(T(str(no), '施工', '%s %s %s' % (chunk[0][0]['fname'], names, work), cur, _wd_add(cur, dd - 1), team_of(chunk[0][0]),
                       (2, 6 if dd > 6 else 3), tags=tuple(x[0]['tag'] for x in chunk)))
            L[-1]['evs'] = [(x[0]['tag'], x[5]) for x in chunk]
            no += 1
            cur = _wd_add(cur, max(2, dd // 2))
            if cur > e - dt.timedelta(days=25):
                cur = _wd_add(s, 7 + no % 5)
    L.append(T(str(no), '試験', '単体試運転・総合試運転', e - dt.timedelta(days=20), e - dt.timedelta(days=6), 'D', (3, 2)))
    L.append(T(str(no + 1), '検査', '完成検査（市）・引渡し', e - dt.timedelta(days=3), e, '事務', (2, 0)))
    return L

# ------------------------------------------------------------ 年度ごとの工程（定修）
_CACHE = {}
def teishu(fy):
    if fy in _CACHE:
        return _CACHE[fy]
    wins = windows(fy)
    if fy == 2011:
        P = dict(f1=H23_F1, f2=H23_F2)
    elif fy == 2025:
        P = dict(f1=R7_F1, f2=R7_F2)
    else:
        t1, e1 = gen_furnace_tasks(fy, 1, wins[1][0], wins[1][1], '1-' if fy >= 2012 else '1')
        t2, e2 = gen_furnace_tasks(fy, 2, wins[2][0], wins[2][1], '2-' if fy >= 2012 else '2')
        wins[1] = (wins[1][0], e1) if (fy, 1) not in SPECIAL_LONG else wins[1]
        wins[2] = (wins[2][0], e2) if (fy, 2) not in SPECIAL_LONG else wins[2]
        c1 = gen_common_tasks(fy, wins, 'C-' if fy >= 2012 else 'C', 1)
        c2 = gen_common_tasks(fy, wins, 'D-' if fy >= 2012 else 'D', 2)
        P = dict(f1=t1 + c1, f2=t2 + c2)
        # 補修工事（別契約）を参考行として入れる
        for c in hoshu_contracts(fy):
            hoshu_tasks(fy, c)
            d = BYTAG[c['tag']]
            pk = place(fy, d)[0]
            key = 'f1' if pk in (1, 'Z') else 'f2'
            if pk == 'Z' and (fy, 1) in SPECIAL_LONG:
                key = 'f2'
            P[key].append(T('別%d' % (len([x for x in P[key] if x['grp'] == '別契約']) + 1), '別契約', '【別契約】' + c['title'].split('清掃工場')[-1],
                            c['_s'], c['_e'], '別', (0, 0), 'sep', tags=(c['tag'],)))
            P[key][-1]['evs'] = []
    P['wins'] = wins
    _CACHE[fy] = P
    return P

def hoshu_all(fy):
    """補修工事と工程（定修の工程表の別契約行と同じ期間）"""
    out = []
    P = teishu(fy)
    for c in hoshu_contracts(fy):
        if fy in (2011, 2025):
            sep = next(t for t in P['f1'] + P['f2'] if t['kind'] == 'sep' and c['tag'] in t['tags'])
            c['_s'], c['_e'] = sep['s'], sep['e']
        else:
            sep = next(t for t in P['f1'] + P['f2'] if t['kind'] == 'sep' and c['tag'] in t['tags'])
            c['_s'], c['_e'] = sep['s'], sep['e']
        L = hoshu_tasks(fy, c)
        out.append((c, L))
    return out

def teishu_period(fy):
    """定期整備業務委託の契約日・工期（両炉の停止の早い方の前に契約し、遅い方の完了後 2 か月を工期末にする）"""
    if fy == 2011:
        return H23['contract_date'], H23['period'][0], H23['period'][1]
    if fy == 2025:
        return R7['contract_date'], R7['period'][0], R7['period'][1]
    w = teishu(fy)['wins']
    first = min(w[1][0], w[2][0]); last = max(w[1][1], w[2][1])
    return first - dt.timedelta(days=25), first - dt.timedelta(days=22), last + dt.timedelta(days=60)
