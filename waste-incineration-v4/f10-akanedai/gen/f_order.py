# -*- coding: utf-8 -*-
"""発注台帳：財務会計 CSV（旧 H23〜H28：支出命令一覧／新 H29〜：契約・負担行為一覧）と副資材単価契約台帳（旧 H23〜R3：月 1 行／新 R4〜：納入 1 回 1 行）。
契約（正解）はここで年度ごとに組み立てる。"""
import csv, io, random, datetime as dt, calendar
from openpyxl import Workbook
from openpyxl.comments import Comment
import world, history as H, periods as PR
from xl import *
from rec import Rec

D = dt.date
VEND = {'千草環境システム株式会社': '0012035', '東雲築炉株式会社': '0018841', 'あかね電設株式会社': '0007712', '北辰計測株式会社': '0021190',
        'みずほ環境分析センター株式会社': '0021455', 'さくら計量器株式会社': '0009930', 'あかね防災設備株式会社': '0011208',
        'ときわ燃料株式会社': '0016601', '白石鉱業株式会社': '0016740', 'あかね酸素株式会社': '0014452', 'あかね台クレーン整備株式会社': '0013377',
        '東和電機サービス株式会社': '0019006', '日の出フィルター株式会社': '0020318', '関東電気保安協会あかね台支部': '0003301'}
SHORT = lambda v: v.replace('株式会社', '(株)')
FACN = world.FACN

HOSHU = {2011: 241_920_000, 2012: 243_600_000, 2013: 246_750_000, 2014: 249_480_000, 2015: 251_100_000, 2016: 252_720_000, 2017: 258_120_000,
         2018: 264_600_000, 2019: 271_700_000, 2020: 279_400_000, 2021: 286_000_000, 2022: 290_400_000, 2023: 296_120_000, 2024: 303_600_000,
         2025: 311_300_000, 2026: 319_000_000}
AKPAY = {2021: 475_200_000, 2022: 2_851_200_000, 2023: 3_009_600_000, 2024: 1_584_000_000}
COKE = {2011: 45_800, 2012: 46_500, 2013: 47_200, 2014: 48_000, 2015: 48_300, 2016: 48_600, 2017: 49_500, 2018: 50_800, 2019: 51_600, 2020: 52_000, 2021: 52_800,
        2022: 66_000, 2023: 76_000, 2024: 72_000, 2025: 71_500, 2026: 70_700}
LIME = {y: int(9_100 + (y - 2011) * 58) // 10 * 10 for y in range(2011, 2022)}; LIME.update({2022: 10_500, 2023: 10_720, 2024: 10_890, 2025: 10_890, 2026: 11_000})
LOX = {y: 52 if y < 2018 else 54 if y < 2021 else 55 for y in range(2011, 2022)}; LOX.update({2022: 60, 2023: 61, 2024: 61, 2025: 61, 2026: 62})
REVS = {2022: [(D(2022, 4, 1), 'コークス', 66_000, '契約単価'), (D(2022, 7, 1), 'コークス', 78_000, '指数 +18.2%'), (D(2022, 10, 1), 'コークス', 81_500, '指数 +4.5%'),
               (D(2023, 1, 1), 'コークス', 79_000, '指数 -3.1%')],
        2023: [(D(2023, 4, 1), 'コークス', 76_000, '契約単価'), (D(2023, 10, 1), 'コークス', 74_200, '指数 -3.2%')],
        2024: [(D(2024, 4, 1), 'コークス', 72_000, '契約単価')],
        2025: [(D(2025, 4, 1), 'コークス', 71_500, '契約単価'), (D(2025, 7, 1), 'コークス', 68_900, '指数 -3.6%'), (D(2026, 1, 1), 'コークス', 70_700, '指数 +2.6%（3%未満のため据置の申出あり→協議で改定）')],
        2026: [(D(2026, 4, 1), 'コークス', 70_700, '契約単価'), (D(2026, 7, 1), 'コークス', 69_800, '指数 -3.3%（協議中のため暫定）')]}

def tax(fy, d=None):
    if d is None: d = D(fy, 10, 1)
    return 1.05 if d < D(2014, 4, 1) else 1.08 if d < D(2019, 10, 1) else 1.10

# ------------------------------------------------------------------ 副資材の正解（月別）
STOPS = {}
def stopdays(P):
    STOPS.clear()
    for p in P:
        if p['kind'] in ('1号定修', '2号定修') or p['key'] in ('AK-2022', 'AK-2023-2'):
            f = 1 if ('1号' in p['kind'] or p['key'] == 'AK-2022') else 2
            d = p['start']
            while d <= p['end']:
                STOPS[(f, d.year, d.month)] = STOPS.get((f, d.year, d.month), 0) + 1; d += dt.timedelta(1)
        if p['kind'] == '全停':
            d = p['start']
            while d <= p['end']:
                for f in (1, 2): STOPS[(f, d.year, d.month)] = STOPS.get((f, d.year, d.month), 0) + 1
                d += dt.timedelta(1)
    for k in list(STOPS):
        STOPS[k] = min(STOPS[k], calendar.monthrange(k[1], k[2])[1])

def months(fy):
    return [(fy, m) for m in range(4, 13)] + [(fy + 1, m) for m in range(1, 4)]

def lox_rate(y, m, fy):
    if fy <= 2018: return 0.22
    if fy <= 2022: return {2019: 0.35, 2020: 0.5, 2021: 0.75, 2022: 0.9}[fy]
    if fy == 2023: return 0.9 if (y, m) < (2023, 6) else 0.0
    return 0.0

def consumption(fy):
    r = random.Random(f'sub-{fy}')
    out = []
    for (y, m) in months(fy):
        if D(y, m, 1) > D(2026, 11, 1): break
        dim = calendar.monthrange(y, m)[1]
        od = [dim - STOPS.get((f, y, m), 0) for f in (1, 2)]
        if D(y, m, 1) == D(2026, 11, 1): od = [max(0, 13 - STOPS.get((1, y, m), 0) * 13 // dim), 13]
        fdays = sum(od)
        starts = sum(1 for f in (1, 2) if STOPS.get((f, y, m), 0) and od[f - 1] > 0)
        ratio = 0.061 if fy < 2018 else 0.058 if fy < 2023 else 0.052
        coke = fdays * 120 * 0.92 * ratio * r.uniform(0.96, 1.04) + starts * 12
        lime = fdays * 120 * 0.92 * 0.036 * r.uniform(0.95, 1.05)
        lox = fdays * lox_rate(y, m, fy) * r.uniform(0.85, 1.15) * 1000
        if fy == 2023 and (y, m) in ((2023, 6), (2023, 7)):   # PSA 更新中（6/12〜7/28）は 1号炉を液酸で運転
            dd = 19 if m == 6 else 28
            lox += dd * 13.0 * 1000 * r.uniform(0.95, 1.05)
        if fy >= 2024 and m == 5: lox = {2024: 22_400, 2025: 31_400, 2026: 27_800}[fy]
        if fy == 2025 and m == 11: lox = 4_600
        out.append(dict(y=y, m=m, coke=round(coke, 1), lime=round(lime, 1), lox=int(round(lox, -2))))
    return out

def price(item, d, fy):
    if fy >= 2022:
        revs = [x for x in REVS.get(fy, []) if x[1] == item and x[0] <= d]
        if revs: return revs[-1][2]
    return {'コークス': COKE, '石灰石': LIME, '液体酸素': LOX}[item][fy]

def monthly_amount(fy, c, item):
    q = {'コークス': c['coke'], '石灰石': c['lime'], '液体酸素': c['lox']}[item]
    p = price(item, D(c['y'], c['m'], 15), fy)
    return int(q * p * tax(fy, D(c['y'], c['m'], 15)))

# ------------------------------------------------------------------ 契約（正解）
def contracts(fy, P, works_by_unit):
    """その年度の契約：dict(no, kind, name, method, vendor, cdate, s, e, amount, change, paid, setsu, pays[(date, amount, memo)], units, sub)"""
    C = []
    T = lambda d: tax(fy, d)
    def add(**k):
        k.setdefault('change', None); k.setdefault('pays', None); k.setdefault('units', []); k.setdefault('sub', None)
        if k['pays'] is None: k['pays'] = [(min(k['e'] + dt.timedelta(30), D(2027, 3, 31)), k['paid'], '')]
        C.append(k)
    a = HOSHU[fy]
    hoshu = dict(no=1, kind='委託', name=f'{PR.wa(fy)} {FACN} 年間保守点検整備業務委託', method='随意契約（特命）', vendor='千草環境システム株式会社',
                 cdate=D(fy, 4, 1), s=D(fy, 4, 1), e=D(fy + 1, 3, 31), amount=a, setsu='委託料', key=f'保守{fy}')
    hoshu['change'] = 289_520_000 if fy == 2021 else None
    paid = hoshu['change'] or a
    if fy == 2026:
        hoshu['paid'] = a // 4 * 2; hoshu['pays'] = [(D(2026, 7, 31), a // 4, '部分払 第1回'), (D(2026, 10, 30), a // 4, '部分払 第2回')]
    else:
        hoshu['paid'] = paid; q = paid // 4
        hoshu['pays'] = [(D(fy, 7, 31), q, '部分払 第1回'), (D(fy, 10, 31), q, '部分払 第2回'), (D(fy + 1, 1, 31), q, '部分払 第3回'), (D(fy + 1, 4, 28), paid - 3 * q, '完了払')]
    add(**hoshu)
    for c in PR.REPAIRS:
        if c['fy'] != fy: continue
        rp = [p for p in P if p.get('contract') == c['key']]
        s = min(p['start'] for p in rp); e = max(p['end'] for p in rp)
        cd = c.get('cs') or (s - dt.timedelta(days=24))
        add(no=10, kind='工事', name=c['name'], method='随意契約（特命）' if c['key'] != 'R2012' else '随意契約（緊急）', vendor='千草環境システム株式会社',
            cdate=cd, s=cd + dt.timedelta(days=3), e=c.get('ce') or (e + dt.timedelta(days=30)), amount=c['amount'], paid=c['amount'], setsu='工事請負費', key=c['key'])
    if fy in AKPAY:
        add(no=12, kind='工事', name=f'{FACN} 基幹的設備改良工事（継続費 令和3〜6年度）', method='随意契約（公募型プロポーザル）', vendor='千草環境システム株式会社',
            cdate=D(2021, 7, 5), s=D(2021, 7, 6), e=D(2025, 3, 25), amount=7_920_000_000, paid=AKPAY[fy], setsu='工事請負費', key='基幹改良',
            pays=[(D(fy + 1, 3, 25) if fy < 2024 else D(2025, 4, 25), AKPAY[fy], f'{PR.wa(fy)} 出来高' if fy < 2024 else '完成払')])
    # 市の物品購入
    for f in (1, 2):
        for u in world.UNITS:
            if u['furn'] == f and u['base'] == 'ろ過式集じん器本体' and H.truth(u).get(fy, ('', ''))[0] == '●':
                p0 = next(p for p in P if p['key'] == f'{fy}-{f}')
                am = int(432 * (24_500 if fy < 2020 else 26_400) * T(p0['start']) // 1000 * 1000)
                add(no=4, kind='物品', name=f'{FACN}{f}号ろ過式集じん器ろ布購入', method='指名競争入札', vendor='日の出フィルター株式会社',
                    cdate=p0['start'] - dt.timedelta(days=52), s=p0['start'] - dt.timedelta(days=49), e=p0['start'] - dt.timedelta(days=7), amount=am, paid=am, setsu='需用費',
                    units=[u['tag']], key=f'ろ布{fy}-{f}')
    if fy == 2016:
        add(no=5, kind='物品', name=f'{FACN}溶融炉羽口購入（予備品）', method='随意契約', vendor='千草環境システム株式会社', cdate=D(2016, 7, 25), s=D(2016, 7, 26),
            e=D(2016, 9, 20), amount=7_452_000, paid=7_452_000, setsu='備品購入費', key='羽口2016', memo='主羽口 3本')
    if fy == 2021:
        add(no=31, kind='物品', name=f'{FACN}溶融炉主羽口購入（予備品）', method='指名競争入札', vendor='千草環境システム株式会社', cdate=D(2021, 11, 19), s=D(2021, 11, 22),
            e=D(2022, 2, 28), amount=5_390_000, paid=5_390_000, setsu='備品購入費', key='羽口2021', memo='主羽口 2本（旧形）')
    # 他業者の作業
    groups = {}
    for u in world.UNITS:
        v = PR.other_vendor(u, fy)
        if v: groups.setdefault(v, []).append(u)
    for v, us in groups.items():
        if v == 'あかね台クレーン整備株式会社':
            nm = f'{FACN}ごみクレーン整備工事（巻上ワイヤ交換ほか）' if fy not in (2020, 2021) else f'{FACN}{"1" if fy == 2020 else "2"}号ごみクレーン走行減速機補修工事'
            am = 2_480_000 + 380_000 * len(us)
        elif v == 'さくら計量器株式会社': nm = f'{FACN}計量機整備工事（ロードセル交換）'; am = 2_200_000 * len(us)
        elif any('空冷復水器' in u['name'] for u in us): nm = f'{FACN}空冷復水器ファン減速機補修工事'; am = 3_850_000
        else: nm = f'{FACN}非常用発電機整備点検業務委託'; am = 4_950_000
        s0 = D(fy, 5, 20) if 'クレーン' in nm else D(fy, 8, 1)
        if fy == 2021 and 'クレーン' in nm: s0 = D(2021, 9, 13)
        add(no=21, kind='工事' if '工事' in nm else '委託', name=nm, method='指名競争入札' if am > 3_000_000 else '見積合わせ', vendor=v, cdate=s0 - dt.timedelta(3), s=s0,
            e=s0 + dt.timedelta(days=60 if fy != 2021 else 102), amount=am, paid=am, setsu='工事請負費' if '工事' in nm else '委託料', units=[u['tag'] for u in us], key=f'他{fy}-{VEND[v]}')
    # 委託（法定点検・測定）
    esc = 1 + (fy - 2011) * 0.012
    add(no=6, kind='委託', name=f'{FACN}排ガス等測定業務委託', method='指名競争入札', vendor='みずほ環境分析センター株式会社', cdate=D(fy, 4, 12), s=D(fy, 4, 12), e=D(fy + 1, 3, 31),
        amount=int(6_800_000 * esc) // 1000 * 1000, paid=int(6_800_000 * esc) // 1000 * 1000, setsu='委託料', key=f'測定{fy}')
    add(no=7, kind='委託', name=f'{FACN}ダイオキシン類作業環境測定業務委託', method='見積合わせ', vendor='北辰計測株式会社', cdate=D(fy, 4, 12), s=D(fy, 4, 12), e=D(fy + 1, 3, 31),
        amount=int(1_100_000 * esc) // 1000 * 1000, paid=int(1_100_000 * esc) // 1000 * 1000, setsu='委託料', key=f'作環{fy}')
    if fy % 2 == 1:
        add(no=9, kind='委託', name=f'{FACN}計量機定期検査業務委託', method='見積合わせ', vendor='さくら計量器株式会社', cdate=D(fy, 6, 1), s=D(fy, 6, 1), e=D(fy, 9, 30),
            amount=int(560_000 * esc) // 1000 * 1000, paid=int(560_000 * esc) // 1000 * 1000, setsu='委託料', key=f'計量{fy}')
    add(no=11, kind='委託', name=f'{FACN}消防用設備等点検業務委託', method='見積合わせ', vendor='あかね防災設備株式会社', cdate=D(fy, 5, 10), s=D(fy, 5, 10), e=D(fy + 1, 3, 31),
        amount=int(900_000 * esc) // 1000 * 1000, paid=int(900_000 * esc) // 1000 * 1000, setsu='委託料', key=f'消防{fy}')
    add(no=2, kind='委託', name=f'{FACN}受変電設備保守点検業務委託', method='随意契約', vendor='関東電気保安協会あかね台支部', cdate=D(fy, 4, 1), s=D(fy, 4, 1), e=D(fy + 1, 3, 31),
        amount=int(2_800_000 * esc) // 1000 * 1000, paid=int(2_800_000 * esc) // 1000 * 1000, setsu='委託料', key=f'受変電{fy}')
    if fy in (2013, 2016, 2019, 2023, 2026):
        add(no=8, kind='委託', name=f'{FACN}精密機能検査業務委託', method='指名競争入札', vendor='みずほ環境分析センター株式会社', cdate=D(fy, 9, 1), s=D(fy, 9, 1), e=D(fy + 1, 2, 28),
            amount=int(4_500_000 * esc) // 1000 * 1000, paid=int(4_500_000 * esc) // 1000 * 1000, setsu='委託料', key=f'精密{fy}')
    # 単価契約
    sub = consumption(fy)
    for item, ven, plan_q, unit_ in (('コークス', 'ときわ燃料株式会社', 4_300 if fy < 2023 else 3_900, 't'), ('石灰石', '白石鉱業株式会社', 2_700, 't'),
                                     ('液体酸素', 'あかね酸素株式会社', {2019: 160_000, 2020: 250_000, 2021: 420_000, 2022: 520_000, 2023: 900_000}.get(fy, 150_000 if fy < 2019 else 60_000), 'kg')):
        spent = sum(monthly_amount(fy, c, item) for c in sub)
        amt = int(plan_q * price(item, D(fy, 4, 1), fy) * T(D(fy, 4, 1)))
        chg = None
        total_q = sum({'コークス': c['coke'], '石灰石': c['lime'], '液体酸素': c['lox']}[item] for c in sub)
        if total_q > plan_q * 0.98 or spent > amt:
            chg = int(spent * 1.05) if item != '液体酸素' or fy != 2021 else int(560_000 * price(item, D(fy, 4, 1), fy) * 1.1)
        pays = [(D(c['y'], c['m'], calendar.monthrange(c['y'], c['m'])[1]) + dt.timedelta(25), monthly_amount(fy, c, item), f"{c['m']}月分")
                for c in sub if monthly_amount(fy, c, item)]
        add(no={'コークス': 3, '石灰石': 5, '液体酸素': 8}[item] + 100, kind='単価', name=f'{FACN}溶融炉用{item}購入（単価契約）' if item != '液体酸素' else f'{FACN}液体酸素購入（単価契約）',
            method='一般競争入札' if item != '液体酸素' else '指名競争入札', vendor=ven, cdate=D(fy, 3, 26) - dt.timedelta(365) + dt.timedelta(1) if False else D(fy, 3, 26),
            s=D(fy, 4, 1), e=D(fy + 1, 3, 31), amount=amt, change=chg, paid=spent, setsu='需用費', pays=pays, key=f'単価{fy}-{item}', sub=(item, unit_, plan_q))
    return C

# ------------------------------------------------------------------ CSV 旧（H23〜H28）
def csv_old(path, fy, C):
    rc = Rec(path.split('/')[-1])
    rows = []; r = random.Random(f'csvo{fy}'); seq = [1600 - fy % 7 * 37]
    for c in C:
        setsu = {'委託料': '委託料', '工事請負費': '工事請負費', '需用費': '需用費', '備品購入費': '備品購入費'}[c['setsu']]
        sai = {'委託料': '施設保守委託' if c['no'] == 1 else '測定委託' if c['no'] in (6, 7) else '点検委託', '工事請負費': '', '需用費': '燃料費' if 'コークス' in c['name'] else '薬品費' if c['kind'] == '単価' else '修繕料', '備品購入費': ''}[c['setsu']]
        nm = c['name'].replace(FACN, 'あかね台資源化センター').replace(f'{PR.wa(fy)} ', f'H{fy-1988} ' if c['no'] == 1 else '')
        if c['kind'] == '単価': nm = nm.replace('あかね台資源化センター', '')
        for (pd, am, memo) in c['pays']:
            seq[0] += r.randint(3, 40)
            kd = pd - dt.timedelta(days=r.randint(18, 28))
            nm2 = (nm + ' ' + memo) if c['kind'] == '単価' else nm
            tek = memo if c['kind'] != '単価' else ''
            if c['kind'] == '単価':
                cc = next(x for x in consumption(fy) if f"{x['m']}月分" == memo)
                q = {'コークス': cc['coke'], '石灰石': cc['lime'], '液体酸素': cc['lox']}[c['sub'][0]]
                tek = f'{q:.1f}t' if c['sub'][1] == 't' else f'{q:,}kg'
            if c.get('memo') and not tek: tek = c['memo']
            rows.append([f'{fy-1988:02d}{seq[0]:05d}', f'H{kd.year-1988}.{kd.month:02d}.{kd.day:02d}', setsu, sai, nm2, SHORT(c['vendor']), am,
                         f'H{pd.year-1988}.{pd.month:02d}.{pd.day:02d}', tek, c['key']])
    rows.sort(key=lambda x: x[0])
    with open(path, 'w', encoding='shift_jis', newline='') as f:
        f.write(f'【架空】あかね台市 財務会計システム 支出命令一覧表,所属:環境部あかね台資源化センター,年度:{PR.wa(fy)},出力日:H{fy+1-1988}.05.12\r\n')
        f.write('伝票番号,起票日,節,細節,件名,相手方,金額,支払日,摘要\r\n')
        for i, x in enumerate(rows, 3):
            f.write(','.join(str(v) for v in x[:9]) + '\r\n')
            rc.add('', i, 'item', dict(zip(['slipNo', 'issued', 'setsu', 'saisetsu', 'name', 'payee', 'amount', 'paid', 'memo'], x[:9])), [x[9]])
    return rc

# ------------------------------------------------------------------ CSV 新（H29〜）
def csv_new(path, fy, C):
    rc = Rec(path.split('/')[-1])
    R = []
    for c in C:
        moku = '04 施設整備費' if c['key'] == '基幹改良' else '03 施設管理費'
        setsu = {'委託料': '13 委託料', '工事請負費': '14 工事請負費', '需用費': '10 需用費', '備品購入費': '17 備品購入費'}[c['setsu']]
        kno = f"{str(fy-2018).rjust(2, '0') if fy >= 2019 else 'H' + str(fy-1988)}-{ {'委託': '委', '工事': '工', '物品': '物', '単価': '単'}[c['kind']] }-{c['no']:03d}"
        if c['key'] == '基幹改良': kno = 'R3-工-007'
        R.append([city(fy), str(fy), '単価契約' if c['kind'] == '単価' else '契約', f"{fy % 100:02d}-{c['no'] + 100:05d}", kno, c['name'].replace(FACN, 'あかね台資源化センター'),
                  c['method'], VEND[c['vendor']], c['vendor'], c['cdate'].strftime('%Y/%m/%d'), c['s'].strftime('%Y/%m/%d'), c['e'].strftime('%Y/%m/%d'),
                  str(c['amount']), str(c['change']) if c['change'] else '', str(c['paid']), '04 衛生費', '02 清掃費', moku, setsu, '環境部あかね台資源化センター', c['key']])
    R.sort(key=lambda x: x[3])
    hdr = ['担当', '年度', '区分', '負担行為番号', '契約番号', '件名', '契約方法', '業者コード', '業者名', '契約日', '履行期間自', '履行期間至',
           '契約金額', '変更後契約金額', '支出済額', '款', '項', '目', '節', '所属']
    buf = io.StringIO(); w = csv.writer(buf, quoting=csv.QUOTE_ALL, lineterminator='\r\n')
    out_d = D(fy + 1, 4, 8) if fy < 2026 else D(2026, 11, 13)
    w.writerow(['【架空】あかね台市財務会計システム 契約・支出負担行為一覧', '所属:環境部あかね台資源化センター', f'対象年度:{fy}', f'出力日:{out_d:%Y/%m/%d}'] + [''] * 16)
    w.writerow(hdr)
    for i, x in enumerate(R, 3):
        w.writerow(x[:20])
        rc.add('', i, 'item', dict(zip(['staff', 'fy', 'kubun', 'commitNo', 'contractNo', 'name', 'method', 'vendorCode', 'vendor', 'contractDate', 'from', 'to',
                                        'amount', 'changedAmount', 'paid', 'kan', 'kou', 'moku', 'setsu', 'section'], x[:20])), [x[20]])
    open(path, 'w', encoding='cp932', newline='').write(buf.getvalue())
    return rc

def city(fy): return PR.city_staff(fy)

# ------------------------------------------------------------------ 副資材台帳 旧（H23〜R3）
def sub_old(path, fy, C):
    sub = consumption(fy)
    rc = Rec(path.split('/')[-1])
    staff = PR.city_staff(fy)
    wb = Workbook(); ws = wb.active; ws.title = world.wa_short(fy)
    props(wb, f'【架空】あかね台資源化センター 副資材 単価契約 納入管理表 {PR.wa(fy)}', f'あかね台市 環境部 あかね台資源化センター {staff}')
    widths(ws, [8, 10, 9, 13, 10, 9, 13, 11, 7, 12, 14, 26])
    merge(ws, 1, 1, 1, 12, f'【架空】あかね台資源化センター　副資材 納入管理表（単価契約）　{PR.wa(fy)}', font(14, True), AL, b=None)
    put(ws, 2, 1, f'施設係 {staff}　金額は税込（円）', font(9), AL)
    cs = {c['sub'][0]: c for c in C if c['kind'] == '単価'}
    for i, item in enumerate(['コークス', '石灰石', '液体酸素']):
        c = cs[item]; un = c['sub'][1]
        put(ws, 3 + i, 1, item, font(9, True), AC, B, fill('EDEDED')); merge(ws, 3 + i, 2, 3 + i, 3, SHORT(c['vendor']).replace('(株)', '㈱'), font(9), AL, B)
        merge(ws, 3 + i, 4, 3 + i, 6, f"{price(item, D(fy, 4, 1), fy):,} 円/{un}（税抜）", font(9), AL, B)
        txt = f"予定 {c['sub'][2]:,}{un}" + (f" → 変更 {c['change']:,} 円（税込）" if c['change'] else '')
        merge(ws, 3 + i, 7, 3 + i, 10, txt, font(9), AL, B)
    r = 7
    merge(ws, r, 1, r + 1, 1, '月', font(9, True), AC, B, fill('DDEBF7'))
    for c0, t in ((2, 'コークス'), (5, '石灰石'), (8, '液体酸素')): merge(ws, r, c0, r, c0 + 2, t, font(9, True), AC, B, fill('DDEBF7'))
    merge(ws, r, 11, r + 1, 11, '合計', font(9, True), AC, B, fill('DDEBF7')); merge(ws, r, 12, r + 1, 12, '備考', font(9, True), AC, B, fill('DDEBF7'))
    for c0, t in zip(range(2, 11), ['数量 t', '回数', '金額', '数量 t', '回数', '金額', '数量 kg', '回数', '金額']): put(ws, r + 1, c0, t, font(8, True), AC, B, fill('DDEBF7'))
    r += 2; tot = [0] * 9
    for c in sub:
        if c['m'] == 10 and fy >= 2016: r += 2   # j2：上期小計の行を消した跡（空き 2 行）
        vals = [f"{c['m']}月", c['coke'], max(1, round(c['coke'] / 13.5)), monthly_amount(fy, c, 'コークス'), c['lime'], max(1, round(c['lime'] / 18)),
                monthly_amount(fy, c, '石灰石'), c['lox'], round(c['lox'] / 9800) if c['lox'] else 0, monthly_amount(fy, c, '液体酸素')]
        vals.append(vals[3] + vals[6] + vals[9])
        for j in range(9): tot[j] += vals[1 + j]
        for cc, v in enumerate(vals, 1): put(ws, r, cc, v, font(9), AC if cc == 1 else AR, B, nf='#,##0.0' if cc in (2, 5) else '#,##0')
        rmk = ''
        if STOPS.get((1, c['y'], c['m'])): rmk = '1号炉 停止 ' + str(STOPS[(1, c['y'], c['m'])]) + '日'
        if STOPS.get((2, c['y'], c['m'])): rmk = (rmk + '　' if rmk else '') + '2号炉 停止 ' + str(STOPS[(2, c['y'], c['m'])]) + '日'
        if fy >= 2019 and c['m'] == 5: rmk = (rmk + '　' if rmk else '') + 'PSA 製品酸素 濃度低下（液酸補給）'
        put(ws, r, 12, rmk, font(8), AL, B)
        rc.add(ws.title, r, 'item', dict(month=f"{c['y']}-{c['m']:02d}", cokeT=vals[1], cokeN=vals[2], cokeYen=vals[3], limeT=vals[4], limeN=vals[5], limeYen=vals[6],
                                         loxKg=vals[7], loxN=vals[8], loxYen=vals[9], totalYen=vals[10], memo=rmk), [cs['コークス']['key'], cs['石灰石']['key'], cs['液体酸素']['key']])
        r += 1
    put(ws, r, 1, '計', font(9, True), AC, B, fill('F2F2F2'))
    for j in range(9): put(ws, r, 2 + j, round(tot[j], 1) if j in (0, 3) else tot[j], font(9, True), AR, B, fill('F2F2F2'), '#,##0.0' if j in (0, 3) else '#,##0')
    put(ws, r, 11, tot[2] + tot[5] + tot[8], font(9, True), AR, B, fill('F2F2F2'), '#,##0')
    rc.add(ws.title, r, 'total', dict(cokeT=round(tot[0], 1), cokeYen=tot[2], limeT=round(tot[3], 1), limeYen=tot[5], loxKg=tot[6], loxYen=tot[8], totalYen=tot[2] + tot[5] + tot[8]))
    ws.cell(r, 11).comment = Comment(f'財務会計の支出済額と照合済（{staff}）', staff)
    put(ws, r + 2, 1, '※ 数量は納品書の計量値（場内トラックスケールで受入確認）。炉別の使用量は運転日報による。', font(8), AL)
    page(ws, land=True, paper=9, fitw=1, fith=1)
    wb.save(path)
    return rc

# ------------------------------------------------------------------ 副資材台帳 新（R4〜）
def sub_new(path, fy, C):
    sub = consumption(fy)
    rc = Rec(path.split('/')[-1])
    staff = 'j2' if fy <= 2023 else 'j3'
    cs = {c['sub'][0]: c for c in C if c['kind'] == '単価'}
    wb = Workbook(); ws = wb.active; ws.title = '納入記録'
    props(wb, f'【架空】あかね台資源化センター 副資材 単価契約台帳 {PR.wa(fy)}', f'あかね台市 環境部 あかね台資源化センター {staff}')
    widths(ws, [4, 11, 9, 12, 9, 6, 10, 13, 8, 8, 8, 9, 22])
    merge(ws, 1, 1, 1, 13, f'【架空】あかね台資源化センター　副資材 単価契約台帳（納入記録）　{PR.wa(fy)}', font(14, True), AL, b=None)
    put(ws, 2, 2, f'R4 改訂様式（納入 1 回 1 行・単価改定を反映）　記入：施設係 {staff}　金額：税込円', font(9), AL)
    for c, t in enumerate(['', 'No.', '納入日', '品目', '納品書No.', '数量', '単位', '適用単価\n(税抜)', '金額(税込)', '検収', '受入計量', '在庫(t)', '備考'], 1):
        if c > 1: put(ws, 4, c, t, font(9, True), AC, B, fill('E2EFDA'))
    ws.row_dimensions[4].height = 26
    r = 5; n = 0; rnd = random.Random(f'sn{fy}'); stock = {'コークス': 62.0, '石灰石': 95.0}; mon = {}
    for c in sub:
        y, m = c['y'], c['m']; dim = calendar.monthrange(y, m)[1]
        last_day = dim if D(y, m, dim) <= D(2026, 11, 13) else 13
        for item, qty, unit_, truck in (('コークス', c['coke'], 't', 13.5), ('石灰石', c['lime'], 't', 18.0), ('液体酸素', c['lox'], 'kg', 9800)):
            if not qty: continue
            k = max(1, round(qty / truck))
            wd = [d for d in range(1, last_day + 1) if D(y, m, d).weekday() < 5]
            days_ = sorted(rnd.sample(wd, min(k, len(wd))))
            while len(days_) < k: days_.append(days_[-1])
            acc = 0; target = monthly_amount(fy, c, item); amt_acc = 0
            for i in range(k):
                q = round(qty / k * rnd.uniform(0.97, 1.03), 2) if unit_ == 't' else int(round(qty / k, -2))
                if i == k - 1: q = round(qty - acc, 2) if unit_ == 't' else int(qty - acc)
                acc += q; d = D(y, m, days_[i]); n += 1
                p = price(item, D(y, m, 15), fy)
                amt = int(q * p * tax(fy, d)) if i < k - 1 else target - amt_acc
                amt_acc += amt
                mt = mon.setdefault((y, m, item), [0, 0]); mt[0] = round(mt[0] + q, 2); mt[1] += amt
                if unit_ == 't': stock[item] = round(stock[item] + q - qty / k, 1)
                memo = ''
                if item == '液体酸素' and m == 5: memo = 'PSA 点検'
                if fy == 2023 and item == '液体酸素' and m in (6, 7): memo = 'PSA 更新中（基幹改良）'
                vals = [None, n, f'{d.month}/{d.day}', item, f"{'TK' if item == 'コークス' else 'SH' if item == '石灰石' else 'AO'}{y % 100}{m:02d}-{i+1:02d}", q, unit_, p, amt,
                        staff if rnd.random() < 0.8 else 'k1', '済' if unit_ == 't' else '—', stock.get(item) if unit_ == 't' else None, memo]
                for cc, v in enumerate(vals, 1):
                    if cc == 1: continue
                    put(ws, r, cc, v, font(8), AL if cc in (4, 5, 13) else AC if cc in (2, 3, 7, 10, 11) else AR, box(hair, thin, thin, hair),
                        nf='#,##0.00' if cc == 6 and unit_ == 't' else '#,##0' if cc in (6, 8, 9) else '0.0' if cc == 12 else None)
                rc.add(ws.title, r, 'item', dict(no=n, date=d, item=item, slip=vals[4], qty=q, unit=unit_, price=p, amount=amt, checker=vals[9], stock=vals[11], memo=memo), [cs[item]['key']])
                r += 1
    ws.freeze_panes = 'C5'
    page(ws, land=False, paper=9, titles='4:4')
    ws2 = wb.create_sheet('単価改定')
    widths(ws2, [12, 10, 12, 44])
    merge(ws2, 1, 1, 1, 4, f'【架空】単価改定の記録（スライド条項 第12条）{PR.wa(fy)}', font(12, True), AL, b=None)
    for c, t in enumerate(['適用日', '品目', '単価(税抜)', '根拠'], 1): put(ws2, 3, c, t, font(9, True), AC, B, fill('E2EFDA'))
    revs = REVS.get(fy, []) + [(D(fy, 4, 1), '石灰石', LIME[fy], '契約単価'), (D(fy, 4, 1), '液体酸素', LOX[fy], '契約単価（円/kg）')]
    for i, (d, it, p, why) in enumerate(revs):
        for c, v in enumerate([d.strftime('%Y/%m/%d'), it, p, why], 1): put(ws2, 4 + i, c, v, font(9), AL if c == 4 else AC, B, nf='#,##0')
        rc.add(ws2.title, 4 + i, 'item', dict(date=d, item=it, price=p, reason=why))
    page(ws2, land=False)
    ws3 = wb.create_sheet('月計')
    widths(ws3, [8, 12, 14, 12, 14, 12, 14, 14])
    merge(ws3, 1, 1, 1, 8, f'【架空】月計（支出命令との照合用）{PR.wa(fy)}', font(12, True), AL, b=None)
    for c, t in enumerate(['月', 'コークス t', '金額', '石灰石 t', '金額', '液酸 kg', '金額', '計'], 1): put(ws3, 3, c, t, font(9, True), AC, B, fill('E2EFDA'))
    rr = 4; T = [0] * 7
    for c in sub:
        y, m = c['y'], c['m']
        a = mon.get((y, m, 'コークス'), [0, 0]); b = mon.get((y, m, '石灰石'), [0, 0]); l = mon.get((y, m, '液体酸素'), [0, 0])
        vals = [f'{m}月', a[0], a[1], b[0], b[1], l[0] or None, l[1] or None, a[1] + b[1] + l[1]]
        for j, v in enumerate(vals[1:]): T[j] += v or 0
        for cc, v in enumerate(vals, 1): put(ws3, rr, cc, v, font(9), AC if cc == 1 else AR, B, nf='#,##0.00' if cc in (2, 4) else '#,##0')
        rc.add(ws3.title, rr, 'subtotal', dict(month=f'{y}-{m:02d}', cokeT=a[0], cokeYen=a[1], limeT=b[0], limeYen=b[1], loxKg=l[0] or None, loxYen=l[1] or None, totalYen=vals[7]))
        rr += 1
    put(ws3, rr, 1, '計', font(9, True), AC, B, fill('F2F2F2'))
    for j in range(7): put(ws3, rr, 2 + j, round(T[j], 2), font(9, True), AR, B, fill('F2F2F2'), '#,##0.00' if j in (0, 2) else '#,##0')
    rc.add(ws3.title, rr, 'total', dict(cokeT=round(T[0], 2), cokeYen=T[1], limeT=round(T[2], 2), limeYen=T[3], loxKg=T[4], loxYen=T[5], totalYen=T[6]))
    page(ws3, land=True)
    wb.save(path)
    return rc
