# -*- coding: utf-8 -*-
"""機器ごとの整備方針・実績（正解）と、各版の計画（その時点で知り得た情報だけから作る）。
記号：○点検 ◎分解整備 △補修 ●更新 ★基幹改良で更新 ◇計測（j3 の様式から。j2 の版では ○ と書く）"""
import random
import world

PREC = {'●': 6, '★': 6, '◎': 5, '△': 4, '◇': 3, '○': 2}
Y0, Y1 = 2011, 2038

def _r(u, salt):
    return random.Random(f"{world.SEED}-{u['tag']}-{salt}")

def params(u):
    """(記号, 周期, 初回位相) のリストと ●（更新）の年。周期 None は計画しない。"""
    r = _r(u, 'p')
    c, imp, kind = u['cat'], u['imp'], u['kind']
    P = []; life = None
    if c == 'rot':
        if imp == 'A':
            P = [('◎', 4, r.randint(0, 3)), ('○', 4, None)]
        elif imp == 'B':
            P = [('◎', 6, (0 if u['duty'] == '常用' else 3) + r.randint(0, 2))]
        else:
            if r.random() < 0.6: P = [('◎', r.choice([8, 10]), r.randint(0, 7))]
        if r.random() < 0.25 and imp != 'A': life = 2006 + r.randint(22, 30)
    elif c == 'motor':
        kw = u['kw'] or 0
        if kw >= 30: P = [('○', 4 if imp == 'A' else 6, r.randint(0, 3))]
        if kw >= 30 and r.random() < 0.5: life = 2006 + r.randint(25, 31)
        elif r.random() < 0.08: life = 2006 + r.randint(20, 28)
    elif c == 'valve':
        if imp == 'A': P = [('◎', r.choice([4, 6]), r.randint(0, 5))]
        elif r.random() < 0.5: P = [('◎', r.choice([8, 10]), r.randint(0, 9))]
        if r.random() < 0.3: life = 2006 + r.randint(23, 30)
    elif c == 'instr':
        if r.random() < 0.35: life = 2006 + r.randint(18, 26)
    elif c == 'analyzer':
        P = [('◎', 3, r.randint(0, 2))]
        life = 2017 if u['furn'] else 2018          # 1 回目の更新（H29 排ガス分析計更新工事／共通は H30）
    elif c == 'panel':
        if imp == 'A': P = [('○', 3, r.randint(0, 2))]
        if r.random() < 0.5: life = 2006 + r.randint(26, 31)
    elif c == 'hv':
        P = [('○', 3, r.randint(0, 2))]
        if r.random() < 0.3: life = 2006 + r.randint(28, 32)
    elif c == 'dcs':
        pass
    elif c == 'static':
        if r.random() < 0.55: P = [('△', r.choice([5, 6, 8]), r.randint(0, 7))]
    elif c == 'furnace':
        P = [('△', 1, 0)]
    elif c == 'furnace2':
        P = [('△', 2, r.randint(0, 1))]
    elif c == 'tuyere':
        P = [('○', 1, 0), ('△', 3, r.randint(0, 2))]
    elif c == 'taphole':
        P = [('△', 1, 0), ('◎', 6, None)]
    elif c == 'boiler':
        P = [('○', 1, 0)]
    elif c == 'sh':
        P = [('◇', 2, r.randint(0, 1)), ('△', 6, r.randint(0, 5))]
    elif c == 'sh2':
        P = [('◇', 2, r.randint(0, 1))]
    elif c == 'bf':
        P = [('○', 2, 1), ('●', 5, None)]
    elif c == 'cat':
        P = [('◇', 2, 0)]
    elif c == 'crane':
        P = [('◎', 3, 0 if u['duty'] == '常用' else 1)]
    elif c == 'crane2':
        P = [('◎', 8, r.randint(0, 7))]
    elif c == 'turbine':
        P = [('◎', 4, None)]
    elif c == 'gen':
        P = [('◎', 8, None), ('○', 4, None)]
    elif c == 'vessel':
        P = [('△', 6, r.randint(0, 5))]
    elif c == 'psa':
        pass
    elif c == 'safety':
        P = [('◎', 2, r.randint(0, 1))]
    elif c == 'egen':
        P = [('◎', 6, 2)]
    elif c == 'scale':
        P = [('◎', 5, r.randint(0, 4))]
        life = 2006 + 25
    elif c == 'nozzle':
        P = [('●', 2, r.randint(0, 1))]
    elif c == 'mech':
        P = [('◎', 4 if imp == 'A' else 6, r.randint(0, 5))]
    return P, life

# 代表機器の実績（正解）を個別に書く。キー：(furn, base)。値：{年: (記号, 内容)}
REP = {
    (1, '出湯口'): {2011: ('△', '出湯口れんが部分補修'), 2012: ('△', '同'), 2013: ('◎', '出湯口再築（定修内）'), 2014: ('△', '部分補修'),
                    2015: ('△', '部分補修'), 2016: ('△', '部分補修'), 2017: ('△', '部分補修'), 2018: ('△', 'スリーブ部分補修'),
                    2019: ('△', '水冷スリーブ交換'), 2020: ('△', '部分補修（開孔回数増）'),
                    2021: ('◎', '出湯口再築・冷却ジャケット交換（R3 補修工事）'), 2022: ('△', '部分補修'), 2023: ('△', '部分補修'),
                    2024: ('△', '部分補修'), 2025: ('△', '部分補修')},
    (2, '出湯口'): {2014: ('◎', '出湯口再築'), 2020: ('◎', '出湯口再築'), 2023: ('★', '基幹改良で炉床と一体に更新')},
    (1, '主羽口装置'): {2015: ('△', '羽口 No.3 漏水 交換'), 2018: ('△', '羽口 No.2・No.6 漏水 交換（市予備品使用）'),
                       2021: ('△', '羽口 No.5 漏水 交換（市予備品使用）'), 2022: ('★', '全数更新（基幹改良）')},
    (0, 'PSA原料空気ブロワ'): {2014: ('◎', '分解整備'), 2018: ('◎', '分解整備'), 2023: ('★', '更新（基幹改良）')},
    (0, 'PSA吸着塔'): {2016: ('△', '吸着剤一部補充'), 2023: ('★', '更新（基幹改良）')},
    (2, '2次過熱器'): {2013: ('◇', '肉厚測定'), 2015: ('◇', '肉厚測定'), 2017: ('◇', '肉厚測定 最小4.4mm'), 2019: ('◇', '最小4.1mm'),
                       2021: ('◇', '最小3.8mm'), 2023: ('◇', '最小3.5mm'), 2025: ('◇', '最小3.2mm（限界3.0mm）')},
    (1, '2次過熱器'): {2012: ('◇', '肉厚測定'), 2014: ('◇', '肉厚測定'), 2016: ('◇', '肉厚測定'), 2018: ('◇', '最小4.3mm'),
                       2021: ('△', '肉厚測定、下部ベンド2本 肉盛溶接（追加）'), 2023: ('◇', '最小3.9mm'),
                       2024: ('△', '下部ベンド部 管4本 部分更新'), 2025: ('◇', '最小4.2mm（更新部除く）')},
    (1, '誘引通風機'): {2014: ('◎', '分解整備'), 2016: ('○', '点検'), 2018: ('◎', '軸受交換・羽根車点検'), 2020: ('○', '点検'),
                       2022: ('◎', '羽根車バランス修正・軸受交換'), 2024: ('○', '点検')},
    (0, '1号ごみクレーン'): {2013: ('◎', '巻上ワイヤ交換'), 2016: ('◎', '巻上ワイヤ交換'), 2019: ('◎', '巻上ワイヤ交換'),
                          2020: ('△', '走行減速機分解'), 2022: ('◎', '巻上ワイヤ交換'), 2025: ('◎', '巻上ワイヤ交換')},
    (0, '蒸気タービン'): {2013: ('◎', '開放点検'), 2017: ('◎', '開放点検'), 2021: ('◎', '開放点検'), 2025: ('◎', '開放点検')},
    (1, 'ろ過式集じん器本体'): {2011: ('○', 'ろ布抜取検査'), 2013: ('○', '抜取検査'), 2015: ('●', 'ろ布全数交換'), 2017: ('○', '抜取検査'),
                         2019: ('○', '抜取検査'), 2020: ('●', 'ろ布全数交換'), 2022: ('○', '抜取検査'), 2024: ('○', '抜取検査 引張強度 残存 72%')},
    (1, '触媒反応塔'): {2012: ('◇', '活性測定'), 2014: ('◇', '活性測定'), 2016: ('●', '上層触媒交換'), 2018: ('◇', '活性測定'),
                       2020: ('◇', '活性測定'), 2022: ('◇', '活性測定'), 2024: ('◇', '活性測定 下層 0.68')},
}
# 空冷復水器ファンは個別の実績で（C・E のみ大きく）
# 生成した実績に重ねる個別の出来事（補修工事・突発）
EXTRA = {(2, 'エコノマイザ'): {2012: ('△', '管漏れ補修（H25.2 緊急停止、補修工事）')},
         (0, '2号ごみクレーン走行装置'): {2021: ('△', '走行減速機補修（補修工事）')},
         (0, '1号ごみクレーン走行装置'): {2020: ('△', '走行減速機分解（補修工事）')}}
ACC = {'空冷復水器ファン（C）': {2019: ('◎', '減速機分解（軸受異音）')}, '空冷復水器ファン（E）': {2023: ('◎', '減速機分解')}}

def rep_key(u):
    return (u['furn'], u['base'])

def project(u, hist, start, end, j3=False):
    """hist（{年: 記号}）の最後の実施から周期で先を引く。start 以後 end まで。"""
    P, life = params(u)
    out = {}
    mid = [x for x in P if x[0] == '○' and any(q[0] == '◎' and q[1] == x[1] for q in P)]
    for sym, cyc, ph in P:
        if not cyc or (sym, cyc, ph) in mid: continue
        done = [y for y, s in hist.items() if s == sym or (sym == '◎' and s in ('●', '★')) or (sym in ('○', '◇') and s in ('◎', '●', '★'))]
        last = max(done) if done else None
        if last is None:
            base = 2006 + (ph if ph is not None else 0)
            y = base
            while y < start: y += cyc
        else:
            y = last + cyc
            if y < start: y = start  # 周期超え → 最初の年へ
        while y <= end:
            if sym == '○' and any(out.get(k) in ('◎',) for k in [y]):
                pass
            if PREC.get(out.get(y, ''), 0) < PREC[sym]: out[y] = sym
            y += cyc
    for sym, cyc, ph in mid:
        ys = [y for y, s in out.items() if s == '◎']
        for y in ys + ([min(ys) - cyc] if ys else []):
            m = y + cyc // 2
            if start <= m <= end and m not in out: out[m] = '○'
    if life:
        rep = [y for y, s in hist.items() if s in ('●', '★')]
        L = life if not rep else max(rep) + (life - 2006)
        if start <= L <= end: out[L] = '●'
    # ○ と ◎ が同じ年のときは ◎。◎ の年に ○ を落とした分は残す
    if not j3:
        out = {y: ('○' if s == '◇' else s) for y, s in out.items()}
    return out

def ak_plan(u):
    a = world.ak_scope(u)
    return a

_TC = {}
def truth(u):
    if u['tag'] in _TC: return _TC[u['tag']]
    h = _truth(u)
    for y, v in EXTRA.get((u['furn'], u['base']), {}).items(): h[y] = v
    _TC[u['tag']] = h
    return h

def _truth(u):
    """2011–2026 の実績（正解）。2026 は上期（〜9月）と 1 号定修（10〜11月、施工中）の区別を持つ。"""
    k = rep_key(u)
    acc = ACC.get(u['base'])
    if k in REP or acc:
        h = dict((REP.get(k) or acc))
        if u['cat'] == 'tuyere':   # 羽口は毎定修の漏水点検（○）を下地に
            for y in list(range(Y0, 2026)):
                if y not in h and y != 2022: h[y] = ('○', '漏水点検')
        hs = {y: v[0] for y, v in h.items()}
        # 2026
        p26 = project(u, hs, 2026, 2026, j3=True)
        for y, s in p26.items(): h[y] = (s, '計画（R8）')
        return h
    r = _r(u, 't')
    h = {}
    hs = {}
    for y in range(Y0, 2027):
        p = project(u, hs, y, y, j3=True)
        a = world.ak_scope(u)
        if a and a[0] == y:
            h[y] = ('★', '基幹改良で更新'); hs[y] = '★'
            continue
        if y in p:
            s = p[y]
            x = r.random()
            if s in ('◎', '△', '○', '◇') and x < 0.07 and y < 2026:
                continue  # 先送り（翌年の投影で拾われる）
            h[y] = (s, ''); hs[y] = s
        elif u['cat'] in ('rot', 'valve', 'mech', 'static') and r.random() < 0.025 and y < 2026:
            h[y] = ('△', '故障補修'); hs[y] = '△'
        a = world.ak_scope(u)
        if a and a[0] == y:
            h[y] = ('★', '基幹改良で更新'); hs[y] = '★'
    return h

def hist_upto(h, y):
    return {k: v[0] for k, v in h.items() if k <= y}

def plan_version(u, h, made_year, start, end, j3):
    """made_year 年度当初に作った版：前年度までの実績から先を引く。基幹改良の予定は既知のものを ★ で載せる。"""
    hs = hist_upto(h, made_year - 1)
    p = project(u, hs, start, end, j3=j3)
    a = world.ak_scope(u)
    if a and made_year >= 2019 and a[0] >= start and a[0] >= made_year - 0:
        p = {y: s for y, s in p.items() if y < a[0]}
        p[a[0]] = '★'
        hs2 = dict(hs); hs2[a[0]] = '★'
        after = project(u, hs2, a[0] + 1, end, j3=j3)
        p.update(after)
    # 代表機器の版ごとの判断（その時点の情報）
    k = rep_key(u)
    if k == (1, '出湯口') and made_year == 2019:
        p[2019] = '◎'; p[2021] = '△'; p[2022] = '△'
    if k == (2, '2次過熱器'):
        if made_year == 2023: p[2028] = '●'
        if made_year >= 2024: p[2027] = '●'; p.pop(2028, None)
    if k == (1, '触媒反応塔') and made_year >= 2024: p[2027] = '●'
    if k == (1, 'ろ過式集じん器本体'):
        if made_year == 2023: p[2025] = '●'
        if made_year == 2024: p[2025] = '●'
        if made_year >= 2026: p[2026] = '●'
    if k == (0, '1号ごみクレーンインバータ盤') and made_year >= 2024: p[2027] = '●'
    if k == (0, '蒸気タービン'):
        for y in range(start, end + 1):
            if (y - 2013) % 4 == 0: p[y] = '◎'
    if k == (1, '出湯口') and made_year >= 2023:
        p[2028] = '◎'
        if p.get(2027) == '◎': p[2027] = '△'
        if p.get(2033) == '◎': p[2033] = '△'; p[2034] = '◎'
    return {y: s for y, s in p.items() if start <= y <= end}

COST = {  # 記号ごとの概算（千円）。cat ごとの規模
    'rot': {'◎': 800, '△': 300, '●': 4500, '○': 0}, 'motor': {'●': 1200}, 'valve': {'◎': 250, '●': 900},
    'instr': {'●': 600}, 'analyzer': {'◎': 400, '●': 6500}, 'panel': {'●': 9000, '○': 150}, 'hv': {'○': 400, '●': 28000},
    'static': {'△': 1200}, 'furnace': {'△': 18000, '●': 520000}, 'furnace2': {'△': 6000}, 'tuyere': {'△': 3500, '○': 0},
    'taphole': {'△': 2500, '◎': 26000}, 'boiler': {'○': 0}, 'sh': {'◇': 600, '△': 4500, '○': 600},
    'sh2': {'◇': 600, '○': 600, '△': 5000, '●': 68000}, 'bf': {'○': 300, '●': 24000}, 'cat': {'◇': 800, '○': 800, '●': 46000},
    'crane': {'◎': 3800, '△': 1800, '●': 32000}, 'crane2': {'◎': 1500}, 'turbine': {'◎': 42000}, 'gen': {'◎': 9000, '○': 1200},
    'vessel': {'△': 1500}, 'safety': {'◎': 350}, 'egen': {'◎': 5500}, 'scale': {'◎': 2200, '●': 18000}, 'nozzle': {'●': 900},
    'mech': {'◎': 1200, '△': 500}, 'psa': {}, 'dcs': {},
}
def cost(u, s):
    if s == '★': return None
    c = COST.get(u['cat'], {}).get(s)
    if c is None: return 0 if s in ('○', '◇') else 500
    kw = u.get('kw') or 0
    if u['cat'] in ('rot', 'motor') and kw:
        c = int(c * (0.6 + min(kw, 200) / 60))
    return c
