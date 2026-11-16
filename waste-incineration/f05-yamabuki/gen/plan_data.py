# -*- coding: utf-8 -*-
"""中長期整備計画の中身（正解の世界の状態・履歴 → 計画年度）。
y1 版（R5 年度版、2023-04 作成、R3〜R12）と y2 版（R6 改訂、2024-12 作成、R6〜R20）。
y2 版は 2024-12 までに分かったことだけで作る（R7 以降の結果は入れない）。
"""
import random, zlib
from common import SEED
from world import MACHINES, old_machines

Y2_YEARS = list(range(2024, 2039))   # R6〜R20
Y1_YEARS = list(range(2021, 2031))   # R3〜R12


def rng_for(m, salt):
    key = f"{m.get('tag') or m.get('old')}|{m['furnace']}|{salt}"
    return random.Random(SEED ^ zlib.crc32(key.encode()))


# 重要度（y2 が追加した列）：A=止まると炉停止・代替なし、B=炉の運転に影響／予備あり、C=事後保全でよい
def importance(m):
    c, n = m['cls'], m['name']
    if c in ('grate', 'feeder', 'hydraulic', 'furnace', 'idf', 'bf', 'tower', 'hv', 'dcs') and m['furnace'] in (1, 2):
        return 'A'
    if c in ('hv',) or n.startswith('受電') or n.startswith('区分開閉器') or n.startswith('共通コントローラ'):
        return 'A'
    if c == 'crane':
        return 'B'   # 予備機あり（F04 と違うところ）
    if c in ('ashpusher', 'fan', 'compressor', 'inst_anal', 'chem', 'blower', 'bf_aux', 'trans', 'gen', 'scale',
             'bucket', 'door', 'battery', 'panel', 'panel_plc', 'inverter'):
        return 'B'
    if c == 'pump' and m['standby'] == '':
        return 'B'
    if c == 'conveyor' and m['furnace'] in (1, 2):
        return 'B'
    if c == 'idle':
        return '－'
    return 'C'


# ---------------- y2 版 ----------------
# 記号：点=点検・測定 整=整備（分解・部品交換） 補=補修 更=更新
def y2_marks(m):
    r = rng_for(m, 'y2')
    c, n, k = m['cls'], m['name'], m['furnace']
    marks = {}
    def add(y, s, cost, note=None):
        if y in Y2_YEARS and y not in marks:
            marks[y] = (s, cost, note)
    renew = m.get('renew')
    if c == 'idle':
        return {}
    if c == 'grate':
        for y in Y2_YEARS:
            add(y, '点', 0)  # 全炉停止の主作業。費用は定期整備に含む
        # 部分取替は燃焼段が早い
        base = renew or 2022
        step = 3 if '燃焼段' in n and '後' not in n else 4
        y = base + step + (1 if k == 2 else 0)
        while y <= 2038:
            marks[y] = ('補', 1800 if step == 3 else 1100, '火格子片部分取替')
            y += step
        marks[base + 15] = ('更', 52000, None)  # 基幹改良から15年
    elif c == 'furnace':
        y = 2025 if k == 2 else 2026
        while y <= 2038:
            add(y, '補', 2600, '耐火物部分補修'); y += 2
    elif c == 'hydraulic':
        y = (renew or 2020) + 4 + r.randint(0, 1)
        while y <= 2038:
            add(y, '整', 900, '作動油・フィルタ・ポンプ'); y += 4
    elif c in ('hyd_dev', 'feeder'):
        y = 2025 + r.randint(0, 3)
        while y <= 2038:
            add(y, '整', 500); y += 5 + r.randint(0, 1)
    elif c == 'burner':
        y = 2025 + r.randint(0, 1)
        while y <= 2038:
            add(y, '点', 150); y += 2
        add(2030 + r.randint(0, 2), '更', 1600)
    elif c == 'ashpusher':
        y = 2025 if k == 2 else 2026  # 2号は R4 に摺動板取替（3年目）
        while y <= 2038:
            add(y, '補', 700, '摺動板'); y += 3
    elif c == 'conveyor':
        if r.random() < 0.85:
            y = 2025 + r.randint(0, 3)
            while y <= 2038:
                add(y, '整', 400, 'チェーン・スクリュー'); y += 4 + r.randint(0, 1)
        if r.random() < 0.4:
            add(2031 + r.randint(0, 5), '更', 2800)
    elif c == 'damper':
        if r.random() < 0.55:
            y = 2025 + r.randint(0, 5)
            while y <= 2038:
                add(y, '整', 250, '軸封・軸受'); y += 6
    elif c == 'fan':
        y = 2025 + r.randint(0, 4)
        while y <= 2038:
            add(y, '整', 650, '軸受交換'); y += 6
    elif c == 'idf':
        if k == 2:
            add(2025, '整', 1900, 'R6振動上昇 軸受交換・動バランス')
            nxt = 2031
        else:
            nxt = 2029  # 1号は H30 に軸受交換（資料外）
        for y in Y2_YEARS:
            add(y, '点', 0)  # 振動測定
        for y in (nxt, nxt + 6):
            marks[y] = ('整', 1700, '軸受交換')
        marks[2034 if k == 1 else 2035] = ('更', 9800, '羽根車・ケーシング')
    elif c == 'inverter':
        add(2026 + (k - 1), '更', 3800 if 'IDF' in n or '誘引' in n else 1500, '製造中止')
    elif c == 'static_hot':
        y = 2024 + r.randint(0, 2)
        while y <= 2038:
            add(y, '点', 120, '肉厚測定'); y += 3
        if '内筒' in n:
            add(2033 + k, '更', 6500)
        else:
            add(2028 + k, '補', 1400)
    elif c == 'tower':
        for y in range(2025, 2039, 2):
            add(y + (k - 1) % 2, '点', 0, '肉厚測定')
    elif c == 'pump':
        if m['standby'] == '予備':
            off = 2
        else:
            off = 0
        if r.random() < 0.9:
            y = 2025 + (r.randint(0, 2) + off) % 5
            while y <= 2038:
                add(y, '整', 300, 'メカシール・軸受'); y += 5
        if r.random() < 0.35 and not renew:
            add(2029 + r.randint(0, 6), '更', 900)
    elif c == 'pump_s' or c == 'fan_s':
        if r.random() < 0.6:
            y = 2026 + r.randint(0, 5)
            while y <= 2038:
                add(y, '整', 200); y += 6
    elif c == 'compressor':
        if '1号' in n:
            add(2026, '更', 6200, '部品供給終了の通知(R6)')
            add(2029, '整', 800); add(2032, '整', 800); add(2035, '整', 800); add(2038, '整', 800)
        elif '2号' in n:
            add(2027, '整', 800); add(2030, '整', 800); add(2033, '整', 800); add(2036, '整', 800)
        else:
            y = 2025 + k - 1
            while y <= 2038:
                add(y, '整', 700, '12,000h'); y += 3
    elif c == 'chem':
        y = 2025 + r.randint(0, 3)
        while y <= 2038:
            add(y, '整', 300); y += 4
    elif c == 'blower':
        y = 2025 + r.randint(0, 3)
        while y <= 2038:
            add(y, '整', 250); y += 4
    elif c == 'bf':
        add((renew or 2022) + 3, '点', 300, 'ろ布抜取試験')
        add((renew or 2022) + 6, '点', 300, 'ろ布抜取試験')
        add((renew or 2022) + 8, '更', 9600, 'ろ布全数')
        add((renew or 2022) + 11, '点', 300, 'ろ布抜取試験')
        add((renew or 2022) + 14, '点', 300, 'ろ布抜取試験')
    elif c == 'bf_aux':
        y = 2026 + r.randint(0, 3)
        while y <= 2038:
            add(y, '整', 350); y += 5
    elif c == 'inst_temp':
        y = 2025 + r.randint(0, 2)
        while y <= 2038:
            add(y, '更', 120, '熱電対'); y += 3
    elif c == 'inst_other':
        if r.random() < 0.75:
            y = 2024 + r.randint(0, 1)
            while y <= 2038:
                add(y, '点', 60, '校正'); y += 2
    elif c == 'inst_anal':
        for y in Y2_YEARS:
            add(y, '点', 450 if '分析計' in n else 180, '年次校正')
        if n == '排ガス分析計':
            marks[2027] = ('更', 14500, '15年経過')
        elif n == 'ばいじん計':
            marks[2029] = ('更', 3600, None)
    elif c == 'camera':
        add(2028 + r.randint(0, 4), '更', 450)
    elif c == 'panel':
        if r.random() < 0.5:
            add(2031 + r.randint(0, 6), '更', 3500)
    elif c == 'panel_plc':
        if not renew:
            add(2027 + r.randint(0, 3), '更', 1800, 'PLC製造中止')
    elif c == 'dcs':
        for y in Y2_YEARS:
            add(y, '点', 0)   # メーカー保守契約（定期整備外）
        add(2031 if m['furnace'] == 0 else 2032, '更', 2400 if m['furnace'] == 0 else 0, 'HMI・PC')
    elif c == 'static':
        if r.random() < 0.25:
            add(2027 + r.randint(0, 8), '点', 80)
    elif c == 'tank':
        if r.random() < 0.5:
            y = 2025 + r.randint(0, 4)
            while y <= 2038:
                add(y, '点', 150, '内部点検'); y += 5
        if r.random() < 0.2:
            add(2030 + r.randint(0, 6), '補', 900, '防食')
    elif c == 'scale':
        add(2028, '整', 1200, 'ロードセル') if '入口' in n else None
        add(2030 if '入口' in n else 2031, '更', 9800)
    elif c == 'pc':
        add(2027, '更', 1500)
    elif c == 'crane':
        for y in Y2_YEARS:
            add(y, '点', 0)  # 年次自主検査は法定欄にも
        if '1号' in n:
            marks[2024] = ('整', 700, 'ワイヤロープ')
            marks[2026] = ('整', 3500, '走行車輪・減速機・ワイヤ')
            marks[2028] = ('整', 700, 'ワイヤロープ')
            marks[2032] = ('整', 2800, '減速機'); marks[2030] = ('整', 700, 'ワイヤロープ')
            marks[2036] = ('更', 98000, None)
        else:
            marks[2025] = ('整', 700, 'ワイヤロープ')
            marks[2029] = ('整', 2600, '減速機'); marks[2031] = ('整', 700, 'ワイヤロープ')
            marks[2037] = ('更', 98000, None)
    elif c == 'bucket':
        y = 2025 + (0 if '1号' in n else 1)
        while y <= 2038:
            add(y, '整', 900, '油圧シリンダ・爪'); y += 3
    elif c == 'door':
        y = 2025 + r.randint(0, 3)
        while y <= 2038:
            add(y, '整', 350, 'シリンダパッキン' if '投入扉' in n else '駆動部'); y += 4
    elif c == 'hv':
        if '区分開閉器' in n:
            add(2025, '更', 1300, 'R6年次で推奨')
        elif n == '受電盤':
            add(2027, '更', 7800, 'VCB・継電器')
        elif n == '高圧配電盤':
            add(2028, '更', 6400, 'VCB')
        else:
            add(2029, '更', 2300)
    elif c == 'trans':
        y = 2025 + r.randint(0, 2)
        while y <= 2038:
            add(y, '点', 80, '絶縁油試験') if '動力' in n else None; y += 3
        add(2032 + r.randint(0, 3), '更', 5200 if '動力' in n else 2400)
    elif c == 'gen':
        add(2028, '整', 1500, '燃料噴射・冷却系')
        add(2034, '更', 18000)
    elif c == 'battery':
        if '直流' in n:
            add(2028, '更', 2600, '蓄電池')
        else:
            add(2025, '更', 900, '蓄電池'); add(2032, '更', 4200)
    elif c == 'hoist':
        if r.random() < 0.6:
            add(2027 + r.randint(0, 5), '整', 400)
    return marks


# R6（2024）の実績：y2 版は 2024-12 作成なので R6 列は実績で塗る
def r6_actual(m, marks):
    c = m['cls']
    s = marks.get(2024)
    if s is None:
        return None
    return True


# ---------------- y1 版（R5 年度版） ----------------
# 記号：○点検 ◎整備 ●更新 △補修 ★基幹改良
Y1_SYM = {'点': '○', '整': '◎', '更': '●', '補': '△'}


def y1_marks(m):
    """y1 は 2020 延命化計画を引き継いだ粗い計画。点検（○）は主要機器だけ。"""
    r = rng_for(m, 'y1')
    c, n, k = m['cls'], (m.get('name') or m.get('old')), m['furnace']
    marks = {}
    def add(y, s):
        if y in Y1_YEARS and y not in marks:
            marks[y] = s
    if m.get('kikan') or m.get('olddcs'):
        ky = 2022 if k == 1 else 2023
        if m.get('olddcs'):
            ky = 2023
        if m['furnace'] == 0 and m['cls'] == 'dcs':
            add(2022, '★'); add(2023, '★')
        else:
            add(ky, '★')
        if c == 'grate':
            if k == 2:
                add(2021, '○'); add(2022, '△')
            else:
                add(2021, '△')
        if c == 'bf' and k == 2:
            add(2021, '△'); add(2022, '△')
        return marks
    if c == 'idle':
        return {}
    if c in ('crane',):
        for y in Y1_YEARS:
            add(y, '○')
        if '1号' in (m.get('old') or ''):
            marks[2022] = '◎'; marks[2024] = '◎'; marks[2027] = '◎'
        else:
            marks[2025] = '◎'; marks[2028] = '◎'
        return marks
    if c in ('inst_anal',):
        for y in Y1_YEARS:
            add(y, '○')
        if m['old'] and m['old'].startswith('排ガス分析計'):
            marks[2027] = '●'
        return marks
    if c == 'idf':
        add(2023 if k == 1 else 2025, '◎')
        add(2029 if k == 1 else 2031, '◎')
        return marks
    if c == 'hv':
        add(2026 + r.randint(0, 2), '●'); return marks
    if c == 'compressor':
        if '1号' in (m.get('old') or ''):
            add(2021, '◎'); add(2024, '◎'); add(2027, '◎'); add(2030, '●')
        elif '2号' in (m.get('old') or ''):
            add(2022, '◎'); add(2025, '◎'); add(2028, '◎')
        else:
            add(2023 if k == 1 else 2024, '◎'); add(2027, '◎')
        return marks
    if c == 'furnace':
        for y in range(2021 + (k - 1), 2031, 2):
            add(y, '△')
        return marks
    if c == 'ashpusher':
        add(2023 if k == 1 else 2024, '△'); add(2027 if k == 1 else 2028, '△')
        return marks
    if c in ('fan', 'pump', 'conveyor', 'hydraulic', 'feeder', 'hyd_dev', 'chem', 'blower', 'door', 'bucket', 'bf_aux'):
        if r.random() < 0.8:
            cyc = 5 if c != 'bucket' else 3
            y = 2021 + r.randint(0, cyc - 1)
            while y <= 2030:
                add(y, '◎'); y += cyc
        if c in ('pump', 'conveyor') and r.random() < 0.25:
            add(2026 + r.randint(0, 4), '●')
        return marks
    if c in ('burner',):
        add(2026 + r.randint(0, 3), '●'); return marks
    if c == 'scale':
        add(2030, '●'); return marks
    if c == 'gen':
        add(2025, '◎'); add(2030, '●'); return marks
    if c == 'battery':
        add(2024 if 'UPS' in (m.get('old') or '') else 2025, '●'); return marks
    if c in ('inverter', 'panel_plc', 'camera', 'pc'):
        if r.random() < 0.6:
            add(2025 + r.randint(0, 5), '●')
        return marks
    if c in ('static_hot',):
        add(2021 + r.randint(0, 2), '○'); add(2026 + r.randint(0, 2), '○')
        return marks
    if c == 'inst_temp':
        add(2022 + r.randint(0, 2), '●'); add(2026 + r.randint(0, 2), '●'); return marks
    if c == 'trans':
        add(2028 + r.randint(0, 2), '●'); return marks
    return marks


# y1 の「済」：R3・R4 の計画のうち実際に行ったもの。行わなかったもの（当初計画は残す）
Y1_NOT_DONE = set()   # (tag/old, furnace, year)


def y1_done(m, y, s):
    """R3・R4 の計画が実施されたか（正解）。基幹改良 R4 は 1号・DCS盤 が実施（出来高）。"""
    if y >= 2023:
        return None
    r = rng_for(m, f'done{y}')
    if s == '★' or m['cls'] in ('grate', 'bf', 'crane', 'inst_anal', 'furnace', 'hv'):
        return True
    # R3 の 2号空気予熱器 補修 等は基幹改良待ちで見送り。ほか一部は予算で先送り
    if r.random() < 0.12:
        return False
    return True


# 費用の目安（千円）y1 は設備ごとの年額にまとめる
Y1_COST = {'○': 120, '◎': 600, '●': 2500, '△': 900, '★': 0}


def y1_cost(m, s):
    c = m['cls']
    base = Y1_COST[s]
    if s == '●':
        base = {'hv': 6800, 'compressor': 6000, 'scale': 9500, 'gen': 17000, 'inst_anal': 14000,
                'battery': 2400, 'trans': 4800}.get(c, base)
    if s == '◎' and c in ('crane',):
        base = 2400
    if s == '◎' and c in ('idf', 'compressor'):
        base = 1600
    if s == '△' and c == 'furnace':
        base = 2400
    return base
