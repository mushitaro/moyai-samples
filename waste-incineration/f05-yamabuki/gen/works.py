# -*- coding: utf-8 -*-
"""正解の作業（2021〜2025 年度、2026 は計画のみ）。
定期整備業務委託（毎年度）・基幹的設備改良工事（R4〜R5）・組合の別契約。
定期整備の作業は、その年度の計画（y1/y2）と状態・履歴から選び、停止（1号中間・2号中間・全炉停止）へ割り付ける。
R4 全炉停止・R7 全炉停止は見本で承認された工程・日報（schedules.py）をそのまま使う。
"""
import random, zlib, datetime as dt
from collections import defaultdict
from common import SEED
from world import MACHINES, old_machines
import schedules as S
import plan_data as PD

D = dt.date
MBY = {m['tag']: m for m in MACHINES}
OLDBY = {}
for _m in old_machines():
    if _m.get('tag'):
        OLDBY[_m['tag']] = _m['old']

HOL = {D(2021, 7, 22), D(2021, 7, 23), D(2021, 9, 20), D(2021, 9, 23), D(2021, 11, 3), D(2021, 11, 23),
       D(2022, 7, 18), D(2022, 9, 19), D(2022, 9, 23), D(2022, 10, 10), D(2022, 11, 3), D(2022, 11, 23),
       D(2023, 7, 17), D(2023, 9, 18), D(2023, 10, 9), D(2023, 11, 3), D(2023, 11, 23),
       D(2024, 2, 12), D(2024, 2, 23), D(2024, 7, 15), D(2024, 10, 14), D(2024, 11, 4),
       D(2025, 7, 21), D(2025, 10, 13), D(2025, 11, 3), D(2025, 11, 24)}


def is_off(d):
    return d.weekday() == 6 or d in HOL


def workdays(a, b, sat=True):
    d = a
    res = []
    while d <= b:
        if not is_off(d) and (sat or d.weekday() < 5):
            res.append(d)
        d += dt.timedelta(1)
    return res


def R(key):
    return random.Random(SEED ^ zlib.crc32(str(key).encode()))


# ---------------------------------------------------------------- 契約
CONTRACTS = {}


def contract(key, **kw):
    CONTRACTS[key] = dict(key=key, **kw)
    return CONTRACTS[key]


TEIKI = {}
for y, d0, d1, sv, amt in ((2021, D(2021, 4, 26), D(2022, 3, 18), 'n1', 38940000), (2022, D(2022, 4, 28), D(2023, 3, 17), 'n1', 39600000),
                           (2023, D(2023, 4, 27), D(2024, 3, 15), 'n1', 41800000), (2024, D(2024, 4, 25), D(2025, 3, 14), 'n2', 42900000),
                           (2025, D(2025, 4, 25), D(2026, 3, 13), 'n2', 44000000)):
    TEIKI[y] = contract(f'T{y}', name=f'令和{y-2018}年度 やまぶき広域処理組合第二工場 定期整備業務委託', vendor='【架空】朝凪炉工株式会社',
                        period=(d0, d1), sv=sv, amount=amt, kind='委託')
KIKAN = contract('KIKAN', name='やまぶき広域処理組合第二工場 基幹的設備改良工事（令和4〜5年度）', vendor='【架空】朝凪炉工株式会社',
                 period=(D(2022, 6, 27), D(2024, 3, 15)), sv='n3', amount=1954700000, kind='請負')

# ---------------------------------------------------------------- 停止
STOPS = {
    2021: dict(mid1=(D(2021, 6, 11), D(2021, 6, 14), D(2021, 6, 18), D(2021, 6, 19)),
               mid2=(D(2021, 7, 9), D(2021, 7, 12), D(2021, 7, 16), D(2021, 7, 17)),
               all=(D(2021, 10, 15), D(2021, 10, 18), D(2021, 10, 29), D(2021, 10, 31)), blackout=D(2021, 10, 24)),
    2022: dict(mid2=(D(2022, 6, 10), D(2022, 6, 13), D(2022, 6, 17), D(2022, 6, 18)),
               all=(D(2022, 10, 14), D(2022, 10, 17), D(2022, 10, 29), D(2022, 10, 30)), blackout=D(2022, 10, 23)),
    2023: dict(mid1=(D(2023, 6, 9), D(2023, 6, 12), D(2023, 6, 16), D(2023, 6, 17)),
               all=(D(2023, 10, 13), D(2023, 10, 16), D(2023, 10, 27), D(2023, 10, 29)), blackout=D(2023, 10, 22)),
    2024: dict(mid1=(D(2024, 6, 7), D(2024, 6, 10), D(2024, 6, 14), D(2024, 6, 15)),
               mid2=(D(2024, 7, 5), D(2024, 7, 8), D(2024, 7, 12), D(2024, 7, 13)),
               all=(D(2024, 10, 11), D(2024, 10, 15), D(2024, 10, 26), D(2024, 10, 27)), blackout=D(2024, 10, 20)),
    2025: dict(mid1=(D(2025, 6, 6), D(2025, 6, 9), D(2025, 6, 13), D(2025, 6, 14)),
               mid2=(D(2025, 7, 4), D(2025, 7, 7), D(2025, 7, 11), D(2025, 7, 12)),
               all=(D(2025, 10, 10), D(2025, 10, 14), D(2025, 10, 25), D(2025, 10, 26)), blackout=D(2025, 10, 19)),
}
EVLABEL = {'mid1': '1号炉中間点検', 'mid2': '2号炉中間点検', 'all': '全炉停止'}
TEAMCAP = {'機械1班': 5, '機械2班': 4, '機械3班（応援）': 5, '電計班': 3, '外注(クレーン)': 3, '外注(計器)': 2, '外注(足場)': 3, '外注(回転機)': 2, '外注(制御)': 2}

# ---------------------------------------------------------------- 作業
WORKS = []
WBY = {}


def add_work(**kw):
    y = kw.get('year')
    n = sum(1 for w in WORKS if w['contract'] == kw['contract']) + 1
    kw['id'] = f"{kw['contract']}-{n:03d}"
    kw.setdefault('days', {})      # date -> [(text, place, team, persons, time, progress)]
    kw.setdefault('finding', '')
    kw.setdefault('measure', None)
    kw.setdefault('parts', [])
    kw.setdefault('rev', 0)
    kw.setdefault('docs', [])      # 書かれた資料の位置
    WORKS.append(kw)
    WBY[kw['id']] = kw
    return kw


def place_of(tags):
    if not tags:
        return '全域'
    m = MBY[tags[0]]
    k = m['furnace']
    s = m['sys']
    if k:
        return {'B': f'{k}号炉', 'F': f'{k}号IDF室' if m['cls'] in ('idf', 'damper') and 'IDF' in m['name'] + '誘引' or '誘引' in m['name'] else ('送風機室' if m['cls'] == 'fan' else f'{k}号炉'),
                'C': f'{k}号減温塔', 'G': f'{k}号BF', 'A': f'{k}号 灰出し', 'O': '燃料ポンプ室', 'I': f'{k}号炉', 'E': '電気室', 'S': '中央制御室'}[s]
    return {'R': 'クレーン' if m['cls'] in ('crane', 'bucket') else 'プラットホーム', 'G': '薬剤サイロ', 'A': '灰出し室', 'O': '燃料ポンプ室',
            'W': 'ポンプ室', 'D': '排水処理室', 'K': '空気圧縮機室', 'E': '電気室', 'I': '各所', 'S': '中央制御室', 'F': '煙突', 'M': '工作室'}[s]


def short(m):
    """日報・工程表での呼び名（現場の略し方）"""
    n = m['name']
    n = n.replace('誘引通風機', 'IDF').replace('バグフィルタ', 'BF').replace('（A）', '(A)').replace('（B）', '(B)')
    n = n.replace('（1号）', '1号').replace('（2号）', '2号')
    k = m['furnace']
    return (f'{k}号 ' if k else '') + n


# 分類ごとの作業の型：(作業名の後ろ, 班, 人数, 日数, [日ごとの文], 属性)
def recipe(m, sym, year, r):
    c, n = m['cls'], short(m)
    T = lambda *a: list(a)
    if sym in ('点', '○', '測'):
        wt = 'INSPECT'
        if c == 'static_hot':
            return ('肉厚測定', '機械2班', 2, 1, T(f'{n} 肉厚測定（超音波厚さ計）'), wt)
        if c == 'tower':
            return ('内部点検・肉厚測定', '機械2班', 3, 2, T(f'{n} 開放・内面付着物除去', f'{n} 肉厚測定・ノズル点検、復旧'), wt)
        if c == 'bf':
            if m.get('renew') and year > m['renew']:
                return ('ろ布抜取試験', '機械2班', 2, 1, T(f'{n} ろ布 5本 抜取・試験体送付'), wt)
            return ('ろ布点検・破損ろ布交換', '機械2班', 4, 2, T(f'{n} 開放・ろ布 全数目視点検', f'{n} 破損ろ布交換・リークテスト'), wt)
        if c in ('inst_other', 'inst_temp'):
            return ('校正', '電計班', 1, 1, T(f'{n} 校正'), wt)
        if c == 'inst_anal':
            return ('年次点検・校正', '外注(計器)', 2, 1, T(f'{n} 年次点検・校正（メーカー）'), wt)
        if c == 'tank':
            return ('内部点検', '機械2班', 2, 1, T(f'{n} 水抜き・内部点検'), wt)
        if c == 'burner':
            return ('点検', '電計班', 1, 1, T(f'{n} ノズル清掃・点火確認'), wt)
        return ('点検', '機械2班', 1, 1, T(f'{n} 点検'), wt)
    if sym in ('更', '●'):
        wt = 'REPLACE'
        if c == 'inst_temp':
            return ('熱電対交換', '電計班', 1, 1, T(f'{n} 熱電対交換'), wt)
        if c in ('pump', 'pump_s'):
            return ('更新', '機械2班', 3, 2, T(f'{n} 撤去・据付', f'{n} 配管接続・試運転'), wt)
        return ('更新', '機械2班', 2, 1, T(f'{n} 更新'), wt)
    if sym in ('補', '△'):
        wt = 'OVERHAUL'
        if c == 'grate':
            return ('火格子片 部分取替', '機械1班', 5, 2, T(f'{n} 割れ・摩耗片 取替', f'{n} 取替続き・作動確認'), wt)
        if c == 'furnace':
            return ('炉内耐火物 部分補修', '機械1班', 4, 4, T(f'{n} 耐火物 はつり', f'{n} 型枠', f'{n} キャスタブル施工', f'{n} 養生'), wt)
        if c == 'ashpusher':
            return ('摺動板 計測・取替', '機械1班', 4, 2, T(f'{n} 摺動板 計測・旧品撤去', f'{n} 摺動板 取付・作動確認'), wt)
        if c == 'bf':
            return ('ろ布点検・破損ろ布交換', '機械2班', 4, 2, T(f'{n} 開放・ろ布 全数目視点検', f'{n} 破損ろ布交換・リークテスト'), wt)
        return ('補修', '機械2班', 2, 1, T(f'{n} 補修'), wt)
    # 整・◎
    wt = 'OVERHAUL'
    if c in ('pump', 'pump_s'):
        return ('分解整備', '機械2班', 2, 2, T(f'{n} 分解・点検', f'{n} メカシール・軸受交換、組立・試運転'), wt)
    if c in ('fan', 'fan_s'):
        return ('軸受交換', '機械2班', 3, 2, T(f'{n} 開放・旧軸受抜取', f'{n} 新軸受組込・芯出し・試運転'), wt)
    if c == 'idf':
        return ('軸受交換', '機械2班', 4, 3, T(f'{n} ケーシング開放・軸受抜取', f'{n} 新軸受組込', f'{n} 芯出し・試運転'), wt)
    if c == 'compressor':
        return ('定期整備', '機械2班', 2, 2, T(f'{n} 分解（エア・オイルセパレータ）', f'{n} 部品交換・組立・試運転'), wt)
    if c in ('conveyor',):
        return ('整備', '機械2班', 2, 1, T(f'{n} チェーン・スプロケット点検、張り調整・給脂'), wt)
    if c in ('chem', 'blower', 'bf_aux'):
        return ('整備', '機械2班', 2, 1, T(f'{n} 分解点検・消耗品交換'), wt)
    if c == 'damper':
        return ('軸封・軸受交換', '機械2班', 2, 1, T(f'{n} 軸封パッキン・軸受交換'), wt)
    if c in ('hyd_dev', 'feeder'):
        return ('分解整備', '機械1班', 2, 2, T(f'{n} シリンダ取外し・分解', f'{n} パッキン交換・組付・作動確認'), wt)
    if c == 'hydraulic':
        return ('作動油・フィルタ交換', '機械1班', 2, 1, T(f'{n} 作動油抜取・フィルタ交換・給油'), wt)
    if c == 'door':
        return ('整備', '機械1班', 2, 1, T(f'{n} シリンダパッキン交換・作動確認' if '投入扉' in m['name'] else f'{n} 駆動部点検・チェーン交換'), wt)
    if c == 'bucket':
        return ('油圧シリンダ・爪整備', '外注(クレーン)', 3, 1, T(f'{n} 油圧シリンダ・爪 点検整備'), wt)
    if c == 'crane':
        return ('ワイヤロープ交換', '外注(クレーン)', 3, 2, T(f'{n} 旧ワイヤ取外し', f'{n} 新ワイヤ取付・試運転'), wt)
    return ('整備', '機械2班', 2, 1, T(f'{n} 整備'), wt)


def finding_for(m, sym, year, r):
    c = m['cls']
    if c in ('fan', 'idf'):
        v = round(r.uniform(1.2, 2.6), 1)
        return f'振動 {v}mm/s', ('軸受振動', '4.5mm/s 以下', f'{v}mm/s', '良')
    if c in ('pump',):
        if sym in ('整', '◎'):
            return 'メカシール 摩耗、交換', None
        return '', None
    if c == 'static_hot':
        v = round(r.uniform(2.6, 3.0) if '予熱' in m['name'] else r.uniform(4.6, 5.4), 1)
        base = '原厚 3.2mm' if '予熱' in m['name'] else '原厚 6.0mm'
        return f'最小肉厚 {v}mm', ('肉厚（最小）', base, f'{v}mm', '良')
    if c == 'grate' and sym in ('点', '○'):
        v = r.randint(74, 92) if (m.get('renew') and year > m['renew']) else r.randint(62, 70)
        return f'残厚 平均{v}%', ('火格子片 残厚（平均）', '原厚比 70%以上', f'{v}%', '良' if v >= 70 else '要注意')
    if c == 'bf':
        if m.get('renew') and year > m['renew']:
            v = r.randint(82, 93)
            return f'ろ布引張強さ 新品比 {v}%', ('ろ布引張強さ', '新品比 60%以上', f'{v}%', '良')
        return '', None
    if c == 'inst_anal' and sym in ('点', '○'):
        return 'スパン調整', ('校正', '－', 'スパン調整', '良')
    return '', None


PARTS = {
    'pump': ('メカニカルシール・軸受', '{n}用', '1式'), 'fan': ('軸受', '{n}用', '2個'), 'idf': ('軸受', '{n}用 2個組', '1式'),
    'compressor': ('定期整備部品', '{n}用 12,000hキット', '1式'), 'conveyor': ('チェーン給脂・調整材', '{n}用', '1式'),
    'damper': ('軸封パッキン・軸受', '{n}用', '1式'), 'hyd_dev': ('シリンダパッキン', '{n}用', '1式'), 'feeder': ('シリンダパッキン', '{n}用', '1式'),
    'hydraulic': ('作動油・フィルタ', 'ISO VG46 400L・フィルタ', '1式'), 'door': ('シリンダパッキン', '{n}用', '1式'),
    'inst_temp': ('熱電対', '{n}用', '1本'), 'chem': ('消耗品', '{n}用', '1式'), 'blower': ('Vベルト・フィルタ', '{n}用', '1式'),
    'bf_aux': ('消耗品', '{n}用', '1式'), 'bucket': ('油圧シリンダパッキン・爪', '{n}用', '1式'),
}


def sym_of(s):
    return {'○': '点', '◎': '整', '●': '更', '△': '補'}.get(s, s)


# ---------------------------------------------------------------- 計画 → その年度の作業
DEFER = {
    # (tag, year): 理由（実施しなかった計画）
    ('1G-08', 2021): '予算減額のため先送り',
    ('2B-02', 2022): 'R4全炉停止で点検、状態良好のため見送り',
    ('0R-04', 2022): 'R4全炉停止で点検、状態良好のため見送り',
    ('2B-05', 2022): '点検の結果良好、R5基幹改良で更新のため見送り',
    ('2B-07', 2022): '点検の結果良好、R5基幹改良で更新のため見送り',
    ('0W-08', 2022): 'R3にB号機を整備済のためA号機を実施',
    ('1F-09', 2023): '振動 1.4mm/s で安定のため見送り',
    ('0K-02', 2022): '点検のみ（運転時間が整備周期に達していない）',
    ('2B-12', 2025): '点検の結果 残厚 13mm、取替不要',
    ('2B-12', 2024): '点検の結果 残厚 14mm、取替不要',
}
KIKAN_DEFER = {2022: 1, 2023: 2}
OVERFLOW_LOG = []
# 決まった作業・生成した作業の発見と値の上書き（(年度, 停止, 作業名) → 項目）。他の資料（R4 報告書など）が引く値
RFIX = {
    (2021, 'all', '火格子 点検・計測'): None,
}
RFIX = {
    (2021, 'mid1', '火格子 点検・計測'): dict(finding='燃焼段 残厚 平均71%・最小58%、脱落 3片', measure=('火格子片 残厚（平均／最小）', '原厚比 70%以上', '71%／58%', '要注意')),
    (2021, 'mid2', '火格子 点検・計測'): dict(finding='燃焼段 残厚 平均68%・最小55%、割れ 8片', measure=('火格子片 残厚（平均／最小）', '原厚比 70%以上', '68%／55%', '要注意')),
    (2021, 'mid1', '誘引通風機 点検・振動測定'): dict(finding='振動 1.4mm/s', measure=('軸受振動', '4.5mm/s 以下', '1.4mm/s', '良')),
    (2021, 'mid2', '誘引通風機 点検・振動測定'): dict(finding='振動 1.9mm/s', measure=('軸受振動', '4.5mm/s 以下', '1.9mm/s', '良')),
    (2021, 'all', '1号 火格子 点検・計測'): None,
    (2021, 'mid2', '減温塔 噴霧ノズル点検・清掃'): dict(finding='下部胴 最小 4.9mm', measure=('下部胴 肉厚（最小）', '原厚 6.0mm', '4.9mm', '要注意')),
    (2021, 'mid1', '1号 BF ろ布点検・破損ろ布交換'): dict(finding='破損 1本 交換', parts=[('ろ布', 'φ150×5,000 PTFE', '1本', '組合予備品', '1号 BF')]),
    (2021, 'all', '2号 BF ろ布点検・破損ろ布交換'): dict(finding='差圧 1.8kPa（設計 1.5）、破損 2本 交換、予備品残 22本',
                                                 measure=('差圧', '1.5kPa 以下', '1.8kPa', '要注意'), parts=[('ろ布', 'φ150×5,000 PTFE', '2本', '組合予備品', '2号 BF')]),
    (2022, 'mid2', '2号 BF ろ布点検・破損ろ布交換'): dict(finding='差圧 1.9kPa、破損 2本 交換', measure=('差圧（6月）', '1.5kPa 以下', '1.9kPa', '要注意'),
                                                  parts=[('ろ布', 'φ150×5,000 PTFE', '2本', '組合予備品', '2号 BF')]),
    (2024, 'all', '誘引通風機 点検・振動測定', '2F-09'): dict(finding='振動 3.9mm/s 上昇、反負荷側軸受に異音', measure=('軸受振動', '4.5mm/s 以下／R4 2.1', '3.9mm/s', '要注意')),
    (2024, 'mid2', '2号 灰押出装置 点検'): dict(finding='摺動板 残厚 14mm（R4取替品）、取替不要', measure=('摺動板 残厚', '原厚 16mm', '14mm', '良')),
    (2024, 'all', 'ごみクレーン1号 ワイヤロープ交換'): dict(parts=[('ワイヤロープ', 'IWRC 6×Fi(29) φ22 85m', '2本', '組合支給品', 'ごみクレーン 1号')]),
}
RFIX = {k: v for k, v in RFIX.items() if v}
# 作業の書き方の上書き（(年度, 機器番号) → 項目）
OVERRIDE = {
    (2021, '1C-01'): dict(name='1号 減温塔 下部胴 当て板補修（計画外）', team='機械2班', persons=3, dur=2,
                         texts=['1号減温塔 下部胴 肉厚測定（24点）・減肉部確認', '1号減温塔 下部胴 当て板溶接（300×300 t6）'],
                         finding='下部胴 最小 4.1mm（限界 4.5mm）、当て板補修 1か所', measure=('下部胴 肉厚（最小）', '原厚 6.0mm／限界 4.5mm', '4.1mm', '不良（当て板）'),
                         parts=[('当て板', 'SS400 t6 300×300', '1枚', '当社手配', '1号 減温塔 下部胴')]),
    (2021, '0K-01'): dict(finding='運転時間 12,300h、OH後 吐出 0.70MPa／82℃', measure=('吐出圧力・温度（OH後）', '0.7MPa／90℃以下', '0.70MPa／82℃', '良')),
    (2023, '1B-12'): dict(name='1号 灰押出装置 摺動板 計測・部分取替', finding='摺動板 残厚 11mm、端部 2枚 取替', measure=('摺動板 残厚', '原厚 16mm', '11mm', '部分取替')),
}
# 計画に無かったが行った作業（報告書・日報に出る）
EXTRA = {
    2021: [('1C-01', '補', 'extra'), ('0W-08', '整', 'extra'), ('0R-12', '整', 'extra'), ('2B-06', '補', 'extra'), ('1G-05', '点', 'extra:mid1')],
    2022: [('2G-05', '点', 'extra:mid2')],
    2024: [('2B-12', '点', 'extra:mid2')],
}   # 基幹改良で停止中の号炉の、改良範囲外の計画は翌年度へ


def wtypes_in(tag, year):
    return {w['wtype'] for w in WORKS if w['year'] == year and tag in w['machines']}


def compat(sym, wts):
    sym = sym_of(sym)
    if not wts:
        return False
    if sym == '点':
        return True
    if sym in ('整', '補'):
        return bool(wts & {'OVERHAUL', 'REPLACE', 'UPGRADE'})
    if sym == '更':
        return bool(wts & {'REPLACE', 'UPGRADE'})
    if sym == '★':
        return 'UPGRADE' in wts
    return False


_Y1V = {}
NOTDONE_WHY = {}


def y1v(m, V):
    """y1 計画の V 年度版（前年度までの見送りのうち「先送り・翌年度へ」を翌年度へ送る）。'_carry' に繰越の注記"""
    key = (m.get('tag') or m['old'], m['furnace'], V)
    if key in _Y1V:
        return _Y1V[key]
    marks = dict(PD.y1_marks(m))
    carry = {}
    if m.get('tag') and V > 2021:
        prev = y1v(m, V - 1)
        marks.update({k: v for k, v in prev.items() if k != '_carry'})
        carry.update(prev.get('_carry', {}))
        y = V - 1
        sy = prev.get(y)
        if sy and sy != '★' and not compat(sy, wtypes_in(m['tag'], y)):
            why = DEFER.get((m['tag'], y)) or NOTDONE_WHY.get((m['tag'], y), '')
            if ('先送り' in why or '翌年度' in why) and (y + 1) not in marks and y + 1 <= 2030:
                marks[y + 1] = sy
                carry[y + 1] = f'R{y-2018}から繰越'
    marks['_carry'] = carry
    _Y1V[key] = marks
    return marks


def plan_items(year):
    """(tag, 記号, 由来)。R3〜R5 は y1 のその年度版、R6 は y1 R5 版の R6 列（＋繰越）、R7 は y2 R6 改訂"""
    if year <= 2024:
        res = []
        for m in old_machines():
            if not m.get('tag'):
                continue
            sy = y1v(m, year).get(year)
            if sy and sy != '★':
                res.append((m['tag'], sy, 'y1'))
        return res
    items = []
    for m in MACHINES:
        mk = PD.y2_marks(m).get(year)
        if mk and (mk[0] != '点' or m['cls'] not in ('grate', 'idf', 'crane', 'dcs', 'tower')):
            items.append((m['tag'], mk[0], 'y2'))
    return items


def done_decision(tag, sym, year, src):
    m = MBY[tag]
    if (tag, year) in DEFER:
        return False, DEFER[(tag, year)]
    if src == 'y1' and year in KIKAN_DEFER and m['furnace'] == KIKAN_DEFER[year] and not m.get('kikan') and m['cls'] != 'inst_anal':
        return False, '基幹改良工事で停止中のため翌年度へ'
    if m['cls'] in ('scale', 'hv', 'trans', 'gen', 'battery', 'inverter', 'panel', 'camera', 'pc', 'dcs'):
        return None, '別契約'
    return True, ''


# ---------------------------------------------------------------- スケジューラ
def schedule_event(items, days, teamcap=TEAMCAP, allow_help=False):
    """items を順に、班の人数の範囲で最初に入る日へ。入らないものを返す（最後の手段で詰める）"""
    used = defaultdict(int)
    over = []
    for it in items:
        cap = teamcap.get(it['team'], 4)
        p = min(it['persons'], cap)
        s0 = it.get('earliest', 0)
        found = None
        for s in range(s0, len(days) - it['dur'] + 1):
            if all(used[(it['team'], days[s + k])] + p <= cap for k in range(it['dur'])):
                found = s
                break
        if found is None and it['team'] in ('機械1班', '機械2班'):
            for alt in [t for t in ('機械1班', '機械2班', '機械3班（応援）') if t != it['team']]:
                if alt.startswith('機械3') and not allow_help:
                    continue
                cap2 = teamcap[alt]
                p = min(it['persons'], cap2)
                for s in range(s0, len(days) - it['dur'] + 1):
                    if all(used[(alt, days[s + k])] + p <= cap2 for k in range(it['dur'])):
                        found = s
                        it['team'] = alt
                        break
                if found is not None:
                    break
        if found is None:
            if it.get('force'):
                found = max(0, len(days) - it['dur'])
            else:
                over.append(it)
                continue
        for k in range(it['dur']):
            used[(it['team'], days[found + k])] += min(it['persons'], teamcap.get(it['team'], 4))
        it['dlist'] = days[found:found + it['dur']]
    # 全炉停止でも入らないもの：最も空いている日に詰める
    return over


def spans_of(dlist, alldays):
    """日のリスト → 連続区間（休日をまたぐ所は続ける）"""
    if not dlist:
        return []
    res = []
    s = e = dlist[0]
    for d in dlist[1:]:
        between = [x for x in alldays if e < x < d]
        if not between:
            e = d
        else:
            res.append((s, e)); s = e = d
    res.append((s, e))
    return res


# ---------------------------------------------------------------- 年度ごとの定期整備を組み立てる
def routine(ev, year, fur_stopped):
    """毎回の決まった作業（計画の 点 と重なるものはここで持つ）"""
    out = []
    ks = fur_stopped
    if ev == 'all':
        out.append(dict(kubun='準備', name='炉内足場 架設', team='外注(足場)', persons=3, dur=1, tags=[], wt='INSPECT',
                        texts=['炉内足場 架設（' + '・'.join(f'{k}号' for k in ks) + '）'], first=True))
    for k in ks:
        out.append(dict(kubun=f'{k}号炉', name='炉内清掃・点検', team='機械1班', persons=3, dur=1, tags=[f'{k}B-09'], wt='INSPECT',
                        texts=[f'{k}号 炉内清掃（クリンカ除去）・耐火物目視点検']))
        out.append(dict(kubun=f'{k}号炉', name='火格子 点検・計測', team='機械1班', persons=3, dur=1, tags=[f'{k}B-05', f'{k}B-06', f'{k}B-07'], wt='INSPECT',
                        texts=[f'{k}号 火格子 全段 点検・残厚計測'], grate=True))
        out.append(dict(kubun=f'{k}号炉', name='誘引通風機 点検・振動測定', team='機械2班', persons=2, dur=1, tags=[f'{k}F-09'], wt='INSPECT',
                        texts=[f'{k}号IDF 軸受給脂・振動測定']))
        out.append(dict(kubun=f'{k}号炉', name='減温塔 噴霧ノズル点検・清掃', team='機械2班', persons=2, dur=1, tags=[f'{k}C-01'], wt='INSPECT',
                        texts=[f'{k}号減温塔 噴霧ノズル取外し・清掃']))
        out.append(dict(kubun=f'{k}号炉', name='計装 校正', team='電計班', persons=2, dur=2,
                        tags=[f'{k}I-01', f'{k}I-02', f'{k}I-06', f'{k}I-07', f'{k}I-08', f'{k}I-09'], wt='INSPECT',
                        texts=[f'{k}号 温度計・圧力計 校正', f'{k}号 流量計・差圧計 校正']))
        out.append(dict(kubun=f'{k}号炉', name='盤点検（MCC・炉前盤）', team='電計班', persons=1, dur=1, tags=[f'{k}E-01', f'{k}E-02'], wt='INSPECT',
                        texts=[f'{k}号MCC・炉前盤 点検・絶縁抵抗測定']))
    if ev == 'all':
        out.append(dict(kubun='共通', name='ごみクレーン 年次自主検査（1・2号）', team='外注(クレーン)', persons=3, dur=2, tags=['0R-14', '0R-15'], wt='INSPECT',
                        texts=['ごみクレーン1号 年次自主検査', 'ごみクレーン2号 年次自主検査']))
        out.append(dict(kubun='共通', name='全停電 立会（受変電年次点検：別契約）', team='電計班', persons=2, dur=1, tags=[], wt='INSPECT',
                        texts=['全停電 立会（停電・復電操作）'], blackout=True))
    for k in ks:
        pass
    last = dict(kubun='全体', name='単体試運転・インターロック確認', team='電計班', persons=3, dur=1, tags=[], wt='INSPECT',
                texts=['各機器 単体試運転・インターロック確認'], last=2)
    last2 = dict(kubun='全体', name='総合試運転' + ('（空運転）・組合確認' if ev == 'all' else '・復旧'), team='機械1班', persons=3, dur=1, tags=[], wt='INSPECT',
                 texts=['総合試運転・組合確認・片付け' if ev == 'all' else '復旧・試運転・片付け'], last=1)
    out += [last, last2]
    if ev == 'all':
        out.append(dict(kubun='準備', name='炉内足場 解体', team='外注(足場)', persons=3, dur=1, tags=[], wt='INSPECT',
                        texts=['炉内足場 解体'], last=3))
    return out


def build_teiki(year):
    st = STOPS[year]
    c = TEIKI[year]
    events = [e for e in ('mid1', 'mid2', 'all') if e in st]
    # 全炉停止時に開放する炉（基幹改良で停止中の炉は本工事側）
    stopped_all = [1, 2]
    if year == 2022:
        stopped_all = [2]
    if year == 2023:
        stopped_all = [1]
    # 計画の作業を停止へ割り付け
    assign = defaultdict(list)
    skipped = []
    hand_tags = set()
    if year == 2022:
        hand_tags = {t for x in S.R4_ALL for t in x[6]}
    if year == 2025:
        hand_tags = {t for x in S.R7_ALL for t in x[6]}
    extra = EXTRA.get(year, [])
    for tag, sym, src in plan_items(year) + extra:
        m = MBY[tag]
        if tag in hand_tags and not src.startswith('extra'):
            continue
        ok, why = done_decision(tag, sym, year, src) if not src.startswith('extra') else (True, '')
        if ok is None:
            NOTDONE_WHY[(tag, year)] = why
            continue
        if ok is False:
            skipped.append((tag, sym, why))
            NOTDONE_WHY[(tag, year)] = why
            continue
        if ok is None:
            NOTDONE_WHY[(tag, year)] = why
        s = sym_of(sym)
        k = m['furnace']
        heavy = m['cls'] in ('grate', 'furnace', 'ashpusher', 'idf', 'tower', 'bf', 'hydraulic', 'static_hot') or s in ('補',)
        if m['cls'] in ('crane',) and s == '点':
            continue  # 年次は決まった作業
        if m['cls'] == 'inst_anal' and s == '点':
            ev = f'mid{k}' if f'mid{k}' in st else 'all'
            if k not in stopped_all and ev == 'all':
                continue   # 基幹改良側の試運転で実施
        elif m['cls'] == 'grate' and s == '点':
            continue
        elif k and heavy and k in stopped_all:
            ev = 'all'
        elif k:
            ev = f'mid{k}' if f'mid{k}' in st else ('all' if k in stopped_all else None)
        else:
            if m['cls'] in ('pump', 'pump_s', 'fan_s', 'tank', 'static', 'chem', 'blower', 'bf_aux', 'conveyor') and m['cls'] != 'crane':
                ev = 'mid1' if 'mid1' in st else ('mid2' if 'mid2' in st else 'all')
            else:
                ev = 'all'
        if src.startswith('extra:'):
            ev = src.split(':')[1]
        if ev == 'all' and year in (2022, 2025):
            ev = [e for e in events if e.startswith('mid')][0] if k == 0 else (f'mid{k}' if f'mid{k}' in st else 'mid2')
        if ev is None:
            skipped.append((tag, sym, '停止なし'))
            continue
        assign[ev].append((tag, s, src))
    result = {}
    built = {}
    for ev in events:
        ks = [int(ev[-1])] if ev.startswith('mid') else stopped_all
        items = routine(ev, year, ks)
        grp_anal = defaultdict(list)
        grp_grate = defaultdict(list)
        extra_items = []
        for tag, s, src in assign[ev]:
            m = MBY[tag]
            if m['cls'] == 'inst_anal':
                grp_anal[m['furnace']].append(tag)
                continue
            if m['cls'] == 'grate' and s == '補':
                grp_grate[m['furnace']].append(tag)
                continue
            r = R((tag, year, s))
            suf, team, pers, dur, texts, wt = recipe(m, s, year, r)
            it = dict(kubun=f"{m['furnace']}号炉" if m['furnace'] else '共通', name=f"{short(m)} {suf}",
                      team=team, persons=pers, dur=dur, tags=[tag], wt=wt, texts=texts, sym=s, src=src)
            ov = OVERRIDE.get((year, tag))
            if ov:
                it.update(ov)
            extra_items.append(it)
        for k, tags in grp_grate.items():
            secs = '・'.join(MBY[t]['name'][4:-2] for t in sorted(tags))
            extra_items.append(dict(kubun=f'{k}号炉', name=f'{k}号 火格子片 部分取替（{secs}）', team='機械1班', persons=5, dur=2 + (len(tags) > 1),
                                    tags=sorted(tags), wt='OVERHAUL', sym='補', src='plan',
                                    texts=[f'{k}号 火格子 割れ・摩耗片 取外し', f'{k}号 新火格子片 取付', f'{k}号 火格子 作動確認・残片取替']))
        for k, tags in grp_anal.items():
            extra_items.append(dict(kubun=f'{k}号炉', name='排ガス分析計・ばいじん計・流量計 年次点検', team='外注(計器)', persons=2, dur=2,
                                    tags=sorted(tags), wt='INSPECT', texts=[f'{k}号 排ガス分析計・O2計 年次点検', f'{k}号 ばいじん計・排ガス流量計 校正'], sym='点'))
        pos = len(items) - (3 if ev == 'all' else 2)
        items[pos:pos] = extra_items
        built[ev] = items
    overflow = []
    for ev in events:
        down, w0, w1, up = st[ev]
        days = workdays(w0, w1)
        ks = [int(ev[-1])] if ev.startswith('mid') else stopped_all
        items = built[ev]
        if ev == 'all' and overflow:
            pos = len(items) - 3
            items[pos:pos] = overflow
            overflow = []
        bl = st.get('blackout') if ev == 'all' else None
        wdays_nobl = [d for d in days if d != bl]
        main = [it for it in items if not it.get('first') and not it.get('last') and not it.get('blackout')]
        for it in items:
            if it.get('first'):
                it['dlist'] = wdays_nobl[:1]
            if it.get('blackout'):
                it['dlist'] = [bl]
        if ev == 'all':
            tail = {1: wdays_nobl[-1:], 2: wdays_nobl[-2:-1], 3: wdays_nobl[-3:-2]}
        else:
            tail = {1: wdays_nobl[-1:], 2: wdays_nobl[-1:]}
        for it in items:
            if it.get('last'):
                it['dlist'] = tail[it['last']]
        body = wdays_nobl[(1 if ev == 'all' else 0):max(1, len(wdays_nobl) - (2 if ev == 'all' else 1))]
        for it in main:
            if it['dur'] > len(body):
                it['dur'] = len(body)
            it['texts'] = (it['texts'] * 3)[:it['dur']] if len(it['texts']) < it['dur'] else it['texts'][:it['dur']]
        for it in main:
            if it.get('sym') == '補':
                it['earliest'] = min(2 if ev == 'all' else 1, max(0, len(body) - it['dur']))
        prio = lambda it: (0 if it.get('grate') or '清掃' in it['name'] else (2 if it.get('sym') == '補' else 1))
        main.sort(key=prio)
        over = schedule_event(main, body, allow_help=True)
        if over and ev != 'all' and 'all' in events:
            for it in over:
                items.remove(it)
                if it['tags'] and MBY[it['tags'][0]]['furnace'] == 0 or True:
                    overflow.append(it)
        elif over:
            OVERFLOW_LOG.append((year, ev, [it['name'] for it in over]))
            for it in over:
                it['force'] = True
            schedule_event(over, body)
        result[ev] = dict(year=year, ev=ev, label=EVLABEL[ev], down=down, w0=w0, w1=w1, up=up, days=days, blackout=bl, items=items,
                          stopped=ks)
    return result, skipped


# ---------------------------------------------------------------- 日報の行（生成した停止）
WEATHER = ['晴', '晴', '晴', '曇', '曇', '雨', '晴', '曇', '晴のち曇', '曇のち雨']


def make_days(evinfo, sv):
    """作業 → 日ごとの行。実績の日は計画と同じ（例外は別に書く）"""
    year = evinfo['year']
    rows = defaultdict(list)
    for it in evinfo['items']:
        dl = it['dlist']
        for i, d in enumerate(dl):
            txt = it['texts'][min(i, len(it['texts']) - 1)]
            prog = '完了' if i == len(dl) - 1 else f'{int(100 * (i + 1) / len(dl))}%'
            tm = '7:30～16:30' if it.get('blackout') else ('8:30～15:00' if it.get('last') == 1 else '8:30～17:00')
            rows[d].append(dict(text=txt, place=place_of(it['tags']) if it['tags'] else ('電気室' if it.get('blackout') else '全域'),
                                team=it['team'], persons=it['persons'], time=tm, prog=prog, wid=it.get('wid')))
    out = {}
    alld = sorted(rows)
    for j, d in enumerate(alld):
        r = R((year, evinfo['ev'], d))
        notes = []
        for it in evinfo['items']:
            if it['dlist'] and it['dlist'][-1] == d and it.get('finding'):
                notes.append(f"{short(MBY[it['tags'][0]]) if it['tags'] else ''} {it['finding']}".strip())
        nxt = alld[j + 1] if j + 1 < len(alld) else None
        tom = ' / '.join(list(dict.fromkeys(x['text'].split('（')[0] for x in rows[nxt]))[:4]) if nxt else '－'
        out[d] = dict(w=r.choice(WEATHER), rows=rows[d], note='。'.join(notes[:3]) + ('。' if notes else ''), tomorrow=tom)
    return out


# ---------------------------------------------------------------- 定期整備 全年度
TEIKI_EV = {}
TEIKI_SKIP = {}


def build_all_teiki():
    for year in (2021, 2022, 2023, 2024, 2025):
        evs, skipped = build_teiki(year)
        TEIKI_SKIP[year] = skipped
        c = f'T{year}'
        if year == 2022:
            evs['all'] = hand_r4_all()
        if year == 2025:
            evs['all'] = hand_r7_all()
        for ev, info in evs.items():
            if info.get('hand'):
                continue
            for it in info['items']:
                r = R((year, ev, it['name'], tuple(it['tags'])))
                if it['tags'] and not it.get('finding'):
                    fd, ms = finding_for(MBY[it['tags'][0]], it.get('sym', '点'), year, r)
                    it['finding'], it['measure'] = fd, ms
                fx = RFIX.get((year, ev, it['name'], it['tags'][0] if it['tags'] else '')) or RFIX.get((year, ev, it['name']))
                if fx:
                    it.update(fx)
                sp = spans_of(it['dlist'], info['days'])
                w = add_work(contract=c, year=year, event=ev, kubun=it['kubun'], name=it['name'], team=it['team'], persons=it['persons'],
                             plan=sp, act=sp, machines=it['tags'], wtype=it['wt'], vendor=vendor_of(it['team']),
                             finding=it.get('finding', ''), measure=it.get('measure'), sym=it.get('sym'), src=it.get('src'))
                m0 = MBY[it['tags'][0]] if it['tags'] else None
                for pp in it.get('parts', []):
                    w['parts'].append(pp)
                if m0 is not None and m0['cls'] in PARTS and it['wt'] != 'INSPECT' and not it.get('parts'):
                    pn, spc, q = PARTS[m0['cls']]
                    w['parts'].append((pn, spc.format(n=m0['name']), q, '当社手配', short(m0)))
                it['wid'] = w['id']
            info['daily'] = make_days(info, TEIKI[year]['sv'])
        TEIKI_EV[year] = evs
    apply_exceptions()


def vendor_of(team):
    return {'外注(クレーン)': '【架空】みさき起重機サービス有限会社', '外注(計器)': '【架空】東雲計器株式会社', '外注(足場)': '【架空】有限会社 ひぐれ足場',
            '外注(回転機)': '【架空】あけぼの回転機株式会社', '外注(制御)': '【架空】かもめシステム株式会社'}.get(team, '【架空】朝凪炉工株式会社')


def apply_exceptions():
    """生成した停止の中の、計画と違った日・追加（日報・工程表の実績と合わせる）"""
    # R3：1号減温塔 当て板補修（計画外、全炉停止）、1号火格子 脱落 3 片（中間点検）
    pass


# ---------------------------------------------------------------- 見本で承認した R4・R7 全炉停止
def hand_r4_all():
    st = STOPS[2022]['all']
    items = []
    for (iid, kubun, nm, team, plan, act, ms, rev) in S.R4_ALL:
        wt = 'OVERHAUL' if any(k in nm for k in ('取替', '補修', '交換', '整備')) else 'INSPECT'
        items.append(dict(iid=iid, kubun=kubun, name=nm, team=team, plan=plan, act=act, tags=ms, rev=rev, wt=wt))
    info = dict(year=2022, ev='all', label='全炉停止', down=st[0], w0=st[1], w1=st[2], up=st[3], days=workdays(st[1], st[2]),
                blackout=STOPS[2022]['blackout'], items=items, stopped=[2], hand=True)
    # 作業
    pers = {}
    for d, rec in S.R4_DAILY.items():
        for (wk, place, team, n, tm, prog) in rec['rows']:
            pass
    for it in items:
        persons = {'機械1班': 5, '機械2班': 4, '電計班': 2, '外注(足場)': 3, '外注(クレーン)': 3, '外注(計器)': 2, '全班': 7}[it['team']]
        w = add_work(contract='T2022', year=2022, event='all', kubun=it['kubun'], name=it['name'], team=it['team'], persons=persons,
                     plan=it['plan'], act=it['act'], machines=it['tags'], wtype=it['wt'], vendor=vendor_of(it['team']), rev=it['rev'])
        it['wid'] = w['id']
        HAND_IID[('R4', it['iid'])] = w['id']
    R4_FIX(items)
    info['daily'] = S.R4_DAILY
    return info


HAND_IID = {}


def R4_FIX(items):
    """R4 の発見・測定・部品（報告書と同じ値）"""
    def W(iid):
        return WBY[HAND_IID[('R4', iid)]]
    W('c').update(finding='燃焼段 残厚 平均64%・最小52%、割れ・摩耗 24片', measure=('火格子片 残厚 平均／最小', 'R3：68%／55%', '64%／52%', '取替（部分）'))
    W('d')['parts'].append(('火格子片', 'ASR-30R 燃焼段用', '24片', '当社手配', '2号炉 燃焼段'))
    W('e').update(finding='側壁 キャスタブル剥離 約0.8m³'); W('e')['parts'].append(('キャスタブル', '高アルミナ質', '1,800kg', '当社手配', '2号炉 側壁'))
    W('g').update(finding='摺動板 残厚 9mm（原厚16mm）、穴あき1か所', measure=('摺動板 残厚', '原厚 16mm', '9mm（穴あき1）', '取替（追加）'))
    W('g2')['parts'].append(('摺動板', '灰押出装置用 SS400 t16', '4枚', '当社手配（追加）', '2号炉 灰押出装置'))
    W('g2')['wtype'] = 'OVERHAUL'
    W('h').update(finding='振動 2.1mm/s', measure=('軸受振動', '4.5mm/s 以下／R3 1.9', '2.1mm/s', '良'))
    W('j').update(finding='下部胴 最小 4.8mm', measure=('下部胴 肉厚（最小）', 'R3：4.9mm', '4.8mm', '要注意'))
    W('k').update(finding='破損ろ布 6本（予備品残14本）'); W('k')['parts'].append(('ろ布', 'φ150×5,000 PTFE', '6本', '組合予備品', '2号炉 BF'))
    W('k')['wtype'] = 'OVERHAUL'
    W('p')['parts'].append(('ワイヤロープ', 'IWRC 6×Fi(29) φ22 85m', '2本', '組合支給品', 'ごみクレーン 1号'))
    W('p')['wtype'] = 'OVERHAUL'
    W('q2').update(finding='1号・2号とも合格', measure=('性能検査', '－', '合格', '良'))
    W('s')['parts'].append(('メカニカルシール・軸受', '冷却水ポンプ用', '1式', '当社手配', '機器冷却水ポンプ(A)'))
    W('s')['wtype'] = 'OVERHAUL'
    W('f')['parts'].append(('油圧フィルタ', '炉油圧ユニット用', '1組', '当社手配', '2号炉'))
    W('r')['parts'].append(('作動油', 'ISO VG46', '20L', '当社手配', '投入扉 油圧ユニット（補充）'))
    W('n').update(finding='スパン調整 CO・HCl', measure=('校正', '－', 'スパン調整 CO・HCl', '良'))


def hand_r7_all():
    st = STOPS[2025]['all']
    items = []
    for (kubun, nm, team, n, plan, act, ms) in S.R7_ALL:
        wt = 'INSPECT'
        if any(k in nm for k in ('取替', '補修', '整備', '交換')) and '取替判定' not in nm:
            wt = 'OVERHAUL'
        if '熱電対交換' in nm:
            wt = 'REPLACE'
        items.append(dict(kubun=kubun, name=nm, team=team, persons=n, plan=plan, act=act, tags=ms, wt=wt))
    info = dict(year=2025, ev='all', label='全炉停止', down=D(2025, 10, 10), w0=D(2025, 10, 14), w1=D(2025, 10, 25), up=D(2025, 10, 26),
                days=workdays(D(2025, 10, 14), D(2025, 10, 25)), blackout=D(2025, 10, 19), items=items, stopped=[1, 2], hand=True)
    for i, it in enumerate(items, 1):
        team = it['team'].replace('外注 ', '外注(').replace('足場', '足場)').replace('計器', '計器)').replace('クレーン', 'クレーン)') if it['team'].startswith('外注 ') else it['team']
        vend = vendor_of(team) if team.startswith('外注(') else ('【架空】あけぼの回転機株式会社 ほか' if '外注' in team else vendor_of(team))
        w = add_work(contract='T2025', year=2025, event='all', kubun=it['kubun'], name=it['name'], team=it['team'], persons=it['persons'],
                     plan=it['plan'], act=it['act'], machines=it['tags'], wtype=it['wt'], vendor=vend)
        it['wid'] = w['id']
        HAND_IID[('R7', i)] = w['id']
    def W(i):
        return WBY[HAND_IID[('R7', i)]]
    W(2).update(finding='1号 燃焼段 摩耗・割れ 18片、乾燥段・後燃焼段 良好', measure=('火格子片 残厚（平均）', '原厚比 70%以上', '燃焼段 73%', '取替（部分）'))
    W(3)['parts'].append(('火格子片', 'ASR-30 燃焼段用', '18片', '当社手配', '1号炉 燃焼段'))
    W(4).update(finding='ろ布 5本 抜取 → 引張強さ 新品比 88%', measure=('ろ布引張強さ', '新品比 60%以上', '88%', '良'))
    W(5).update(finding='振動 1.6mm/s', measure=('軸受振動', '4.5mm/s 以下', '1.6mm/s', '良'))
    W(6)['parts'].append(('熱電対', 'K 煙突入口温度計用', '1本', '当社手配', '1号 煙突入口温度計'))
    W(8).update(finding='2号 火格子 残厚 平均86%', measure=('火格子片 残厚（平均）', '原厚比 70%以上', '86%', '良'))
    W(10)['parts'].append(('キャスタブル', '高アルミナ質', '1,400kg', '当社手配', '2号炉 側壁・後壁'))
    W(11).update(finding='摺動板 残厚 13mm、取替不要', measure=('摺動板 残厚', '原厚 16mm', '13mm', '良'))
    W(12).update(finding='旧軸受 外輪はく離。交換後 4.8mm/s → 再バランス後 2.2mm/s', measure=('軸受振動（交換後）', '4.5mm/s 以下', '2.2mm/s（再バランス後）', '良'),
                 wtype='OVERHAUL')
    W(12)['parts'].append(('軸受', '2号IDF用 22320 2個組', '1式', '当社手配', '2号 誘引通風機'))
    W(13)['parts'].append(('軸封パッキン', '2号IDF入口ダンパ用', '1式', '当社手配', '2号 IDF入口ダンパ'))
    W(14)['parts'].append(('熱電対', 'K 減温塔出口温度計用', '1本', '当社手配', '2号 減温塔出口温度計'))
    W(15).update(finding='2号 噴霧ノズル 1本 先端摩耗のため交換'); W(15)['parts'].append(('噴霧ノズル', '二流体ノズル', '1本', '当社手配（予備品）', '2号 減温塔'))
    W(18)['parts'].append(('ワイヤロープ', 'IWRC 6×Fi(29) φ22 85m', '2本', '組合支給品', 'ごみクレーン 2号'))
    W(19).update(finding='1号・2号 異常なし', measure=('年次自主検査', '－', '異常なし', '良'))
    W(21)['parts'].append(('シリンダパッキン', '投入扉用', '3組', '当社手配', '投入扉 No.1・2・4'))
    W(22)['parts'].append(('作動油・フィルタ', 'ISO VG46 200L', '1式', '当社手配', '投入扉 油圧ユニット'))
    W(23).update(finding='噴射水供給ポンプ(A) 軸スリーブ摩耗、予備品で対応'); W(23)['parts'].append(('メカニカルシール・軸受', 'ポンプ(A) 3台分', '1式', '当社手配', 'ポンプ(A) 3台'))
    W(17).update(finding='スパン調整', measure=('校正', '－', 'スパン調整', '良'))
    info['weekly'] = S.R7_WEEKLY
    return info


# ---------------------------------------------------------------- 基幹的設備改良工事
def kikan_tags(wbs, k):
    t = {
        '1.1': ['B-05', 'B-06', 'B-07'], '1.2': ['B-08'], '1.3': ['B-05', 'B-06', 'B-07'], '1.4': ['B-05', 'B-06', 'B-07'], '1.5': ['B-09'],
        '2.1': ['C-01'], '2.2': ['C-01'], '2.3': ['C-01'], '2.4': ['C-02', 'C-03'], '2.5': ['C-01'],
        '3.1': ['G-05'], '3.2': ['G-05'], '3.3': ['G-06', 'G-07'], '3.4': ['G-05'],
        '4.2': ['S-01', 'E-03', 'E-04'], '4.3': ['I-03', 'I-04', 'I-07'],
    }.get(wbs, [])
    res = [f'{k}{x}' for x in t]
    if wbs == '4.1':
        res = [m['tag'] for m in MACHINES if m['furnace'] == 0 and m['cls'] == 'dcs']
    return res


SUBCON = {'仮設班': '【架空】有限会社 ひぐれ足場', '機械A班': '【架空】かもしか工業株式会社', '機械B班': '【架空】ほしぞら鉄工株式会社',
          '機械C班': '【架空】白浜フィルタ工業株式会社', '築炉班': '【架空】いぶき築炉株式会社', '配管班': '【架空】さざれ配管工業有限会社',
          '保温班': '【架空】ゆきやなぎ保温株式会社', '電計班': '【架空】かもめシステム株式会社・さざなみ電機株式会社', '全班': '【架空】朝凪炉工株式会社（全協力会社）', '－': '－'}

# 基幹改良 R5（2号炉＋DCS切替）
KIKAN_R5 = [
    ('0', '準備工', None, None, [], 'H'),
    ('0.1', '仮設・資材置場 再設置', '仮設班', 4, [(D(2023, 9, 11), D(2023, 9, 15))], 'N'),
    ('0.2', '2号炉 立下げ・冷却（運転委託）', '－', None, [(D(2023, 9, 19), D(2023, 9, 22))], 'C'),
    ('0.3', '足場・養生', '仮設班', 6, [(D(2023, 9, 25), D(2023, 9, 29))], 'C'),
    ('1', '2号炉 火格子更新', None, None, [], 'H'),
    ('1.1', '既設火格子・油圧配管 解体撤去', '機械A班', 8, [(D(2023, 9, 26), D(2023, 10, 7))], 'C'),
    ('1.2', '炉油圧ユニット 更新', '機械B班', 4, [(D(2023, 10, 10), D(2023, 10, 20))], 'N'),
    ('1.3', '新火格子 搬入（炉室ホイスト）', '機械A班', 6, [(D(2023, 10, 23), D(2023, 10, 25))], 'C'),
    ('1.4', '新火格子 据付・芯出し', '機械A班', 8, [(D(2023, 10, 26), D(2023, 11, 15))], 'C'),
    ('1.5', '炉床まわり耐火物 施工', '築炉班', 6, [(D(2023, 11, 6), D(2023, 11, 17))], 'C'),
    ('1.6', '据付検査（組合立会）', '－', None, [(D(2023, 11, 16), D(2023, 11, 16))], 'K'),
    ('2', '2号炉 減温塔更新', None, None, [], 'H'),
    ('2.1', '既設下部胴・ノズル 解体（25t ラフター）', '機械B班', 5, [(D(2023, 10, 2), D(2023, 10, 7))], 'N'),
    ('2.2', '新下部胴 据付・溶接', '機械B班', 5, [(D(2023, 10, 23), D(2023, 11, 2))], 'N'),
    ('2.3', '二流体ノズル・噴射水配管', '配管班', 4, [(D(2023, 11, 6), D(2023, 11, 14))], 'N'),
    ('2.4', '噴射水ポンプ(A)(B) 更新', '配管班', 3, [(D(2023, 10, 23), D(2023, 10, 28))], 'N'),
    ('2.5', '保温・外装', '保温班', 4, [(D(2023, 11, 15), D(2023, 11, 21))], 'N'),
    ('2.6', '【追加】2号煙道 減肉部補修', '機械B班', 3, [(D(2023, 11, 6), D(2023, 11, 9))], 'N'),
    ('3', '2号炉 バグフィルタ更新', None, None, [], 'H'),
    ('3.1', '既設ろ布・リテーナ 撤去', '機械C班', 5, [(D(2023, 9, 27), D(2023, 10, 3))], 'N'),
    ('3.2', '本体ケーシング・ホッパ 更新', '機械C班', 6, [(D(2023, 10, 4), D(2023, 11, 1))], 'N'),
    ('3.3', '逆洗装置・加温装置 据付', '機械C班', 4, [(D(2023, 11, 2), D(2023, 11, 14))], 'N'),
    ('3.4', 'ろ布 装着（192本）', '機械C班', 6, [(D(2023, 11, 20), D(2023, 11, 23))], 'N'),
    ('4', '計装・制御', None, None, [], 'H'),
    ('4.2', '2号炉 FCS・BF盤・減温塔盤 更新', '電計班', 4, [(D(2023, 11, 20), D(2023, 11, 25))], 'N'),
    ('4.3', '計装（温度・差圧）取替', '電計班', 2, [(D(2023, 11, 20), D(2023, 11, 24))], 'N'),
    ('4.4', 'ループチェック', '電計班', 4, [(D(2023, 11, 27), D(2023, 12, 1))], 'C'),
    ('5', '試運転', None, None, [], 'H'),
    ('5.1', '足場解体・片付け', '仮設班', 6, [(D(2023, 11, 27), D(2023, 11, 29))], 'N'),
    ('5.2', '空運転・耐火物乾燥', '全班', 6, [(D(2023, 11, 30), D(2023, 12, 3))], 'C'),
    ('5.3', 'ごみ投入 試運転・排ガス確認', '全班', 6, [(D(2023, 12, 4), D(2023, 12, 8))], 'C'),
    ('5.4', '出来形確認（組合）', '－', None, [(D(2023, 12, 8), D(2023, 12, 8))], 'M'),
    ('6', 'DCS 切替（全炉停止）', None, None, [], 'H'),
    ('6.1', '1・2号炉 立下げ（運転委託）', '－', None, [(D(2024, 2, 2), D(2024, 2, 3))], 'C'),
    ('6.2', '旧DCS 撤去・新DCS 結線切替', '電計班', 6, [(D(2024, 2, 5), D(2024, 2, 9))], 'C'),
    ('6.3', 'ループチェック・連動試験（1・2号）', '電計班', 6, [(D(2024, 2, 10), D(2024, 2, 15))], 'C'),
    ('6.4', '立上げ・総合試運転', '全班', 6, [(D(2024, 2, 16), D(2024, 2, 22))], 'C'),
    ('6.5', '完成検査（組合）', '－', None, [(D(2024, 3, 8), D(2024, 3, 8))], 'M'),
]
KIKAN_R5_RANGE = (D(2023, 9, 11), D(2023, 12, 16))
KIKAN_DCS_RANGE = (D(2024, 1, 29), D(2024, 3, 9))

# 基幹改良 日報の文（作業の段取りが日ごとに進む）
KSTEP = {
    '0.1': ['仮設事務所 据付', '資材置場 敷鉄板', '仮設電源 引込', '仮設トイレ・休憩所', '安全看板・区画'],
    '0.3': ['炉室 足場 組立（下段）', '足場 組立（上段）', '開口部養生・防炎シート', '粉じん養生・負圧集じん機設置', '足場点検（組合確認）'],
    '1.1': ['火格子駆動部 切離し', '後燃焼段 火格子片 撤去', '後燃焼段 フレーム 切断', '燃焼段 火格子片 撤去', '燃焼段 フレーム 切断', '乾燥段 火格子片 撤去',
            '乾燥段 フレーム 切断', '油圧配管 撤去', '油圧シリンダ 撤去', '撤去材 搬出', '炉床 清掃・墨出し'],
    '1.2': ['旧油圧ユニット 切離し', '旧ユニット 撤去', 'ベース 補修', '新ユニット 据付', '油圧配管 敷設', '配管 フラッシング', '耐圧試験', '作動確認'],
    '1.3': ['新火格子 フレーム 搬入', '燃焼段 火格子片 荷揚げ', '乾燥段・後燃焼段 火格子片 荷揚げ'],
    '1.4': ['乾燥段 フレーム 据付', '乾燥段 レベル調整', '燃焼段 フレーム 据付', '燃焼段 レベル調整', '後燃焼段 フレーム 据付', '後燃焼段 レベル調整',
            '可動火格子 駆動軸 据付', '火格子片 取付（乾燥段）', '火格子片 取付（燃焼段）', '火格子片 取付（後燃焼段）', 'サイドシール 取付', '油圧シリンダ 取付',
            '芯出し', '駆動 単体試験', 'ストローク調整', '据付寸法 記録', '手直し'],
    '1.5': ['炉床まわり 旧耐火物 はつり', 'アンカー 溶接', '型枠', 'キャスタブル 打設（左側）', 'キャスタブル 打設（右側）', '脱型', '目地・仕上げ', '自然乾燥'],
    '2.1': ['下部胴 切離し準備', 'ラフター 据付・下部胴 吊り下し', 'ノズル・配管 撤去', '旧ホッパ 撤去', '撤去材 搬出', '据付面 清掃'],
    '2.2': ['新下部胴 吊込み', '芯出し', '溶接（周継手）', '溶接（縦継手）', '非破壊検査', 'ホッパ 取付', 'マンホール 取付', '手直し・塗装', '気密試験'],
    '2.3': ['二流体ノズル 取付', '噴射水配管 敷設', '噴霧空気配管 敷設', '配管 耐圧試験', 'ノズル 噴霧試験', '調整'],
    '2.4': ['旧ポンプ 撤去', 'ベース 据付', '新ポンプ(A) 据付', '新ポンプ(B) 据付', '配管接続', '単体試運転'],
    '2.5': ['保温材 取付（下部）', '保温材 取付（上部）', '外装板 取付', '外装板 取付続き', '仕上げ'],
    '2.6': ['煙道 減肉部 調査・UT', '当板 製作', '当板 溶接', '塗装・保温復旧'],
    '3.1': ['ろ布 抜取（第1・2室）', 'ろ布 抜取（第3・4室）', 'リテーナ 撤去', '旧ろ布 梱包・搬出', '内部 清掃'],
    '3.2': ['旧ケーシング 切断（第1室）', '旧ケーシング 切断（第2室）', '旧ケーシング 切断（第3・4室）', '旧ホッパ 撤去', '新ホッパ 据付', '新ケーシング 吊込み（第1・2室）',
            '新ケーシング 吊込み（第3・4室）', 'ケーシング 溶接', 'ダクト 接続', 'セル板 取付', '点検歩廊', '気密試験', '手直し'],
    '3.3': ['逆洗ヘッダ 取付', '電磁弁 取付（48個）', '逆洗用空気槽 据付', '加温ヒータ 取付', '循環ファン 据付', '配線', '単体試験'],
    '3.4': ['ろ布 装着（第1室）', 'ろ布 装着（第2室）', 'ろ布 装着（第3室）', 'ろ布 装着（第4室）・プレコート'],
    '4.1': ['DCS盤 搬入', '盤 据付（共通FCS）', '盤 据付（ネットワーク盤）', 'OPS 据付', '電源 配線', '信号 配線', '盤内 確認', '単体通電'],
    '4.2': ['FCS 盤 撤去・据付', 'BF盤 更新', '減温塔盤 更新', '配線 接続', '通電確認'],
    '4.3': ['熱電対 取替', '差圧伝送器 取替', '導圧配管', '配線', '確認'],
    '4.4': ['ループチェック（燃焼系）', 'ループチェック（通風系）', 'ループチェック（減温塔）', 'ループチェック（BF）', 'インターロック試験'],
    '5.1': ['足場 解体（上段）', '足場 解体（下段）', '片付け・清掃'],
    '5.2': ['空運転（送風機・IDF）', '耐火物乾燥 昇温（200℃）', '昇温（400℃）', '昇温（600℃）・保持'],
    '5.3': ['ごみ投入 開始・燃焼調整', '燃焼調整・減温塔噴霧調整', 'BF 差圧確認・逆洗調整', '排ガス確認（自主測定）', '連続運転確認'],
    '6.2': ['旧DCS 盤 切離し', '旧CRT・グラフィックパネル 撤去', '新DCS 結線切替（1号炉）', '新DCS 結線切替（2号炉）', '新DCS 結線切替（共通）'],
    '6.3': ['ループチェック（1号炉）', 'ループチェック（2号炉）', 'ループチェック（共通）', '連動試験（1号炉）', '連動試験（2号炉）', '帳票確認'],
    '6.4': ['1号炉 立上げ', '1号炉 燃焼調整', '2号炉 立上げ', '2号炉 燃焼調整', '総合試運転', '総合試運転', '運転員 操作教育'],
}

KIKAN_EV = {}


def build_kikan():
    for key, wbs_list, k in (('R4', S.KIKAN_R4, 1), ('R5', KIKAN_R5, 2)):
        for (wbs, nm, team, n, spans, kind) in wbs_list:
            if kind == 'H' or not spans:
                continue
            kk = k if not wbs.startswith('6') else 0
            tags = kikan_tags(wbs, k) if not wbs.startswith('6') else [m['tag'] for m in MACHINES if m['furnace'] == 0 and m['cls'] == 'dcs']
            if wbs == '4.1' and key == 'R5':
                tags = []
            wt = 'UPGRADE' if tags else 'INSPECT'
            if wbs.startswith('0') or wbs in ('5.1',):
                wt = 'INSPECT'
            if wbs == '2.6':
                wt = 'OVERHAUL'
            year = spans[0][0].year if spans[0][0].month >= 4 else spans[0][0].year - 1
            w = add_work(contract='KIKAN', year=year, event=f'kikan{key}', kubun=wbs, name=nm, team=team, persons=n or 0,
                         plan=spans, act=spans, machines=tags, wtype=wt, vendor=SUBCON.get(team or '－', '－'), wbs=wbs, kind=kind, rev=1 if '追加' in nm else 0)
            KIKAN_EV.setdefault(key, []).append(w)


# ---------------------------------------------------------------- 組合の別契約（支出負担行為に出る作業）
OTHER = []


def build_other():
    def o(year, name, vendor, d0, d1, tags, wt, contract_key):
        w = add_work(contract=contract_key, year=year, event='other', kubun='別契約', name=name, team='－', persons=0, plan=[(d0, d1)], act=[(d0, d1)],
                     machines=tags, wtype=wt, vendor=vendor)
        OTHER.append(w)
        return w
    for y in range(2021, 2026):
        bl = STOPS[y]['blackout']
        o(y, '受変電設備 年次点検（全停電）', '【架空】一般財団法人 みずなら電気保安協会', bl, bl,
          ['0E-01', '0E-02', '0E-03', '0E-04', '0E-05', '0E-06', '0E-07'], 'INSPECT', f'E{y}')
        o(y, '消防用設備等点検（総合）', '【架空】ささゆり防災設備株式会社', D(y, 9, 14), D(y, 9, 15), [], 'INSPECT', f'F{y}')
        o(y, '非常用発電機 点検', '【架空】ほくと発動機株式会社', D(y, 11, 16), D(y, 11, 16), ['0E-12', '0E-13'], 'INSPECT', f'G{y}')
        o(y, '排ガス等測定（ばい煙・DXN）', '【架空】株式会社 かやの環境分析センター', D(y, 11, 17), D(y, 11, 18), [], 'INSPECT', f'A{y}')
    for y in (2021, 2023, 2025):
        o(y, '計量機 定期検査', '【架空】とまり計量器株式会社', D(y, 9, 7) if y != 2025 else D(2025, 9, 9), D(y, 9, 7) if y != 2025 else D(2025, 9, 9),
          ['0R-01', '0R-02'], 'INSPECT', f'S{y}')
    for y in (2022, 2024):
        o(y, 'クレーン性能検査（2基）', '【架空】一般社団法人 ひかり検査協会', STOPS[y]['all'][1] + dt.timedelta(3), STOPS[y]['all'][1] + dt.timedelta(3),
          ['0R-14', '0R-15'], 'INSPECT', f'C{y}')
    o(2021, '精密機能検査', '【架空】一般社団法人 くるみ廃棄物技術協会', D(2022, 1, 18), D(2022, 1, 20), [], 'INSPECT', 'P2021')
    o(2024, '精密機能検査', '【架空】一般社団法人 くるみ廃棄物技術協会', D(2025, 1, 21), D(2025, 1, 23), [], 'INSPECT', 'P2024')
    o(2024, '計量機（出口）ロードセル交換修理', '【架空】とまり計量器株式会社', D(2024, 8, 20), D(2024, 8, 20), ['0R-02'], 'OVERHAUL', 'L2024')
    o(2025, '区分開閉器（PAS）更新工事', '【架空】有限会社 こもれび電設', D(2025, 10, 19), D(2025, 10, 19), ['0E-01'], 'REPLACE', 'PAS2025')
    o(2025, '無停電電源装置 蓄電池交換', '【架空】さざなみ電機株式会社', D(2025, 11, 26), D(2025, 11, 26), ['0E-15'], 'REPLACE', 'UPS2025')
    o(2023, '2号ごみクレーン 巻上ブレーキライニング交換修理', '【架空】みさき起重機株式会社', D(2023, 9, 28), D(2023, 9, 28), ['0R-15'], 'OVERHAUL', 'CR2023')
    o(2022, '計量機 表示器修理', '【架空】とまり計量器株式会社', D(2022, 8, 3), D(2022, 8, 3), ['0R-03'], 'OVERHAUL', 'SC2022')


# どの資料にも書いていない作業（正解だけに持つ）
UNRECORDED = []


def build_unrecorded():
    for y in range(2021, 2026):
        UNRECORDED.append(dict(name='ごみクレーン 月次自主検査（1・2号）', machines=['0R-14', '0R-15'], wtype='INSPECT', year=y, by='運転業務委託（【架空】やまぶき環境サービス）',
                               reason='運転委託の仕様に含まれ、記録は運転会社の点検表（取り込み資料の外）'))
        UNRECORDED.append(dict(name='第二種圧力容器 定期自主検査（空気槽2基）', machines=['0K-04', '0K-05'], wtype='INSPECT', year=y, by='運転業務委託',
                               reason='運転委託で実施。y1 計画には無い（計画の欠け）、y2 R6 改訂から法定欄に記入'))
        UNRECORDED.append(dict(name='ホイスト 定期自主検査（4基）', machines=['0M-01', '0M-02', '0M-03', '0M-04'], wtype='INSPECT', year=y, by='運転業務委託',
                               reason='運転委託で実施、点検表は運転会社'))
        UNRECORDED.append(dict(name='DXN類 作業環境測定（年2回）', machines=[], wtype='INSPECT', year=y, by='【架空】かやの環境分析センター',
                               reason='支出負担行為の行はあるが、作業の帳票は測定報告書（取り込み資料の外）'))
        UNRECORDED.append(dict(name='受変電設備 月次点検', machines=['0E-02', '0E-03'], wtype='INSPECT', year=y, by='【架空】みずなら電気保安協会',
                               reason='電気保安業務委託の月次点検。点検記録は保安協会の様式（取り込み資料の外）'))
    for y, dd in ((2024, D(2024, 11, 12)), (2025, D(2025, 11, 11))):
        w = add_work(contract=f'DCS{y}', year=y, event='other', kubun='瑕疵担保点検', name='新DCS 年次点検（瑕疵担保期間中の無償点検）', team='－', persons=2,
                     plan=[(dd, dd)], act=[(dd, dd)], machines=[m['tag'] for m in MACHINES if m['cls'] == 'dcs'], wtype='INSPECT',
                     vendor='【架空】かもめシステム株式会社')
        w['unrecorded'] = True
    for y in range(2024, 2026):
        UNRECORDED.append(dict(name='新DCS 年次保守', machines=[m['tag'] for m in MACHINES if m['furnace'] == 0 and m['cls'] == 'dcs'], wtype='INSPECT', year=y,
                               by='【架空】かもめシステム（基幹改良の瑕疵担保期間中の無償点検）', reason='瑕疵担保期間中の無償点検で契約・帳票なし。y2 計画では毎年「点」'))


def build():
    _Y1V.clear(); NOTDONE_WHY.clear()
    WORKS.clear(); WBY.clear(); HAND_IID.clear(); OTHER.clear(); UNRECORDED.clear(); KIKAN_EV.clear()
    build_all_teiki()
    build_kikan()
    build_other()
    build_unrecorded()


build()

if __name__ == '__main__':
    from collections import Counter
    print(len(WORKS), Counter(w['contract'][:1] for w in WORKS))
    for y in TEIKI_EV:
        for ev, info in TEIKI_EV[y].items():
            print(y, ev, len(info['items']), info['w0'], info['w1'])
        print('  skip', TEIKI_SKIP[y])
