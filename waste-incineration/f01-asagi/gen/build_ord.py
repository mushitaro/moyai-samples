# -*- coding: utf-8 -*-
"""発注台帳：H9〜H20 は契約係 c1 の Excel、H21〜R1 は旧財務会計システムの CSV、R2〜 は新システムの CSV。"""
import csv, os, datetime as dt
from openpyxl import Workbook
from xlutil import *
import ans, world2 as W
import f01_truth as TR
from f01_master import wareki, LEGAL, legal_years
from f01_history import BYTAG
from f01_contracts import MAKER, OPCO, teishu_amount, hoshu_amount
import build_sd as B
D = dt.date
KAKU = '【架空】'

def tax(fy):
    return 0.05 if fy < 2014 else (0.08 if fy < 2019 else 0.10)

def esc(fy):
    return 1.0 + 0.012 * (fy - 1997)

# 品目: (件名, 区分, 節キー, 相手方, 金額, 契約日, 履行開始, 履行終了, 支払済, 単価内容, 摘要, 結ぶ作業)
def items(fy):
    L = []
    tx = tax(fy)
    T = lambda a: int(round(a * (1 + tx), -3))
    s0 = D(fy, 4, 1); e0 = D(fy + 1, 3, 31)
    wins = W.windows(fy)
    pref = wareki(fy, False)
    def add(title, kind, sk, vendor, amt, cd, s, e, paid=None, tanka='', note='', works=()):
        if fy == 2026 and paid is None:
            paid = 0 if e > D(2026, 11, 13) and kind in ('業務委託', '工事') else amt
        L.append(dict(title=title, kind=kind, sk=sk, vendor=vendor, amt=amt, cd=cd, s=s, e=e, paid=amt if paid is None else paid, tanka=tanka, note=note, works=list(works)))
    if fy >= 2006:
        if fy <= 2012:
            add('%sあさぎ市清掃工場運転業務委託' % pref, '業務委託', 'itaku', OPCO, int(round(178000000 * (1 + 0.012 * (fy - 2006)) * (1 + tx) / 1.05, -3)), D(fy, 3, 25) - dt.timedelta(days=365), s0, e0)
        else:
            blk = 2013 if fy <= 2017 else (2018 if fy <= 2022 else 2023)
            per = {2013: 199500000, 2018: 231000000, 2023: 264000000}[blk]
            add('あさぎ市清掃工場運転業務委託（%s〜%s・%s分）' % (wareki(blk, False).replace('度', ''), wareki(blk + 4, False).replace('度', ''), pref), '業務委託', 'itaku', OPCO,
                per, D(blk - 1, 3, 22), s0, e0, note='長期継続契約')
    tw = [w['id'] for w in TR.WORKS if w['fy'] == fy and w['contract_type'] == 'T']
    paid = None if fy < 2026 else int(round(teishu_amount(fy) * 0.4, -3))
    add('%s あさぎ市清掃工場 定期整備業務委託' % pref, '業務委託', 'itaku', MAKER, teishu_amount(fy), *W.teishu_period(fy), paid=paid, note='竣工メーカー特命' + ('（前金払 40%）' if fy == 2026 else ''), works=tw)
    for c in W.hoshu_contracts(fy):
        a = hoshu_amount(fy, c['tag'])
        add(c['title'], '工事', 'koji', c['vendor'], a, c['cdate'], c['s'], c['e_'], works=TR.work_ids(fy, c['tag'], 'H'))
    for c in W.special_contracts(fy):
        amt = {('X', 2000): 1029000000, ('X', 2001): 1071000000, ('D', 2012): 399000000, ('N', 2010): 27825000}.get((c['kind'], fy))
        if c['kind'] == 'K':
            amt = int(round(W.R('k', fy).uniform(0.98, 1.0) * sum(cc for d in W.EQ for (k, t, cc, ct) in d['act'].get(fy, []) if ct == 'K') * 1380 * (1 + tx), -3))
        add(c['title'], '工事', 'koji', c['vendor'], amt, c['cdate'] if c['kind'] != 'K' else D(fy, 4, 1) if fy > 2017 else c['cdate'], c['s'], c['e_'],
            note='継続費' if c['kind'] in ('X', 'K') else '', works=[w['id'] for w in TR.WORKS if w['fy'] == fy and w['contract_type'] == c['kind']])
    sup = {2007: 'BF-101', 2013: 'BF-101', 2019: 'BF-101', 2025: 'BF-101', 2008: 'BF-201', 2014: 'BF-201', 2020: 'BF-201', 2026: 'BF-201'}.get(fy)
    if sup:
        up = int(round(23500 * esc(fy) / esc(2007), -2)) if fy != 2025 else 27500
        f = BYTAG[sup]['f']
        add('バグフィルタ用ろ布（φ165×6,000）購入（%d号炉用）' % f, '物品', 'juyo', '株式会社しおかぜフィルター', T(300 * up), D(fy, 4, 11), D(fy, 4, 14), wins[f][0] + dt.timedelta(days=10),
            tanka='%s円/本×300本' % '{:,}'.format(up), note='定期整備で支給。予備 12 本', works=[w['id'] for w in TR.WORKS if w['fy'] == fy and w['tag'] == sup and 'ろ布' in w['text']])
    if fy == 2011:
        add('ボイラ給水ポンプ用メカニカルシール（予備品）2組', '物品', 'juyo', MAKER, 882000, D(2011, 7, 29), D(2011, 7, 29), D(2011, 9, 30), note='予備品 倉庫保管（未使用）')
    if fy == 2025:
        add('燃焼火格子片（予備品）購入', '物品', 'juyo', MAKER, 1188000, D(2025, 4, 18), D(2025, 4, 21), D(2025, 5, 30), tanka='90,000円/枚×12枚',
            note='R7定修 1号炉で割れ片交換に使用', works=TR.work_ids(2025, 'ST-102'))
    # 単価契約（薬剤・燃料）
    chem = [('消石灰', 't', 780, 31000, '東浜ケミカル株式会社'), ('飛灰処理用キレート剤', 't', 68, 280000, '東浜ケミカル株式会社'),
            ('アンモニア水（25%）', 't', 95, 46000, 'あさぎ薬品商事株式会社'), ('灯油', 'kL', 120, 52000, 'あさぎ石油販売株式会社'),
            ('ボイラ用薬品', 'kg', 2400, 1150, 'あさぎ薬品商事株式会社')]
    if fy >= 2000:
        chem.insert(1, ('活性炭', 't', 46 if fy >= 2002 else 23, 250000, '株式会社みさと炭素'))
    for (nm, unit, q, up0, ven) in chem:
        up = int(round(up0 * esc(fy) * (1.6 if nm == '灯油' and fy >= 2008 else 1.0), -2 if up0 > 5000 else 0))
        used = round(q * W.R('use', fy, nm).uniform(0.88, 0.98), 1)
        add('%s %s単価契約' % (pref, nm), '物品（単価）', 'juyo', ven, T(q * up), D(fy, 4, 1),
            s0, e0, paid=T(used * up) if fy < 2026 else T(used * up * 0.6), tanka='%s円/%s（税抜）予定数量%s%s' % ('{:,}'.format(up), unit, q, unit))
    # 法令に基づく点検・測定の委託
    LG = {it[0]: it for it in LEGAL}
    legal = [('LG09', '清掃工場排ガス等測定分析業務委託', '株式会社あさぎ環境分析センター', 5200000),
             ('LG10', '清掃工場作業環境測定業務委託（ダイオキシン類）', '株式会社あさぎ環境分析センター', 800000),
             ('LG04', '清掃工場電気保安業務委託', '一般財団法人あさぎ電気保安協会', 1950000),
             ('LG06', '清掃工場クレーン定期自主検査業務委託', 'あさぎクレーン整備株式会社', 1500000),
             ('LG11', '清掃工場計量機定期点検業務委託', '株式会社みなみ計量器', 700000),
             ('LG12', '清掃工場消防用設備等点検業務委託', 'あさぎ防災設備株式会社', 560000),
             ('LG05', '清掃工場非常用発電機点検業務委託', 'あさぎ防災設備株式会社', 260000),
             ('LG13', 'あさぎ市清掃工場精密機能検査業務委託', '一般財団法人あさぎ環境技術協会', 4300000),
             ('LG08', '第一種圧力容器性能検査手数料（低圧蒸気だめ・給湯熱交換器）', '一般社団法人あさぎボイラ協会', 48000)]
    for (lid, title, ven, base) in legal:
        if fy not in legal_years(LG[lid], 1997, 2026) or (lid == 'LG10' and fy < 2001):
            continue
        t = title
        ws = ['%s-%d' % (lid, fy)]
        if lid == 'LG06' and fy in legal_years(LG['LG07'], 1997, 2026):
            t = title.replace('定期自主検査', '定期自主検査・性能検査受検'); ws.append('LG07-%d' % fy)
        if lid == 'LG11' and fy in legal_years(LG['LG11'], 1997, 2026):
            t = title + '（定期検査受検含む）'
        kind, sk = ('役務', 'ekimu') if lid == 'LG08' else ('業務委託', 'itaku')
        add('%s %s' % (pref, t), kind, sk, ven, int(round(base * esc(fy) * (1 + tx) / 1.05, -3)), D(fy, 4, 3), D(fy, 4, 6), D(fy + 1, 3, 20), works=ws)
    add('%s 飛灰処理物運搬業務委託' % pref, '業務委託', 'itaku', '有限会社あさぎ運輸', int(round(8700000 * esc(fy) * (1 + tx) / 1.05, -3)), D(fy, 4, 1) - dt.timedelta(days=5), s0, e0)
    if fy >= 2013:
        add('%s 清掃工場DCSソフトウェア保守業務委託（リモート監視・障害対応）' % pref, '業務委託', 'itaku', MAKER, int(round(4500000 * (1 + tx) / 1.08, -3)), D(fy, 4, 1), s0, e0)
    r = W.R('shuzen', fy)
    pool = ['ピット放水銃ノズル修繕', '計量データ処理装置プリンタ修繕', 'プラットホーム投入扉油圧ホース修繕', 'ごみクレーン操作卓スイッチ修繕', '中央制御室照明修繕',
            '洗車装置ポンプ修繕', '排水処理設備pH計修繕', '灰出し室集じん機ダクト修繕', '純水装置イオン交換樹脂補充']
    if fy < 2010:
        pool = [p for p in pool if '放水銃' not in p]
    for p in r.sample(pool, 2 + r.randint(0, 2)):
        cd = D(fy, r.randint(5, 12), r.randint(1, 26)) if True else None
        if fy == 2026 and cd > D(2026, 11, 12):
            continue
        add(p, '修繕', 'juyo_s', r.choice([MAKER, 'あさぎ防災設備株式会社', '株式会社みなみ計量器']), int(round(r.randint(9, 48) * 10000 * (1 + tx), -2)), cd, cd + dt.timedelta(days=1), cd + dt.timedelta(days=18))
    if fy == 2026:
        L = [x for x in L if x['cd'] <= D(2026, 11, 12)]
    return L

SETSU_OLD = {'itaku': ('13', '委託料'), 'koji': ('15', '工事請負費'), 'juyo': ('11', '需用費'), 'juyo_s': ('11', '需用費'), 'ekimu': ('12', '役務費')}
SETSU_NEW = {'itaku': ('12', '委託料', '委託料'), 'koji': ('14', '工事請負費', '工事請負費'), 'juyo': ('10', '需用費', '消耗品費'), 'juyo_s': ('10', '需用費', '修繕料'), 'ekimu': ('11', '役務費', '手数料')}
VCODE = {}
def vcode(v):
    if v not in VCODE:
        VCODE[v] = str(10000 + (W.h('v', v) % 89999))
    return VCODE[v]
def wa(x):
    return 'H%02d.%02d.%02d' % (x.year - 1988, x.month, x.day)

def method(x):
    if x['kind'] == '工事' and x['vendor'] != MAKER or ('工事' in x['kind'] and '指名' in x.get('note', '')):
        return '指名競争入札'
    if x['kind'] == '工事':
        return '指名競争入札' if '補修' in x['title'] or '更新工事' in x['title'] and '基幹' not in x['title'] else '随意契約'
    if x['kind'] == '物品（単価）':
        return '指名競争入札'
    if x['amt'] >= 5000000 and x['vendor'] not in (MAKER, OPCO) and 'kind' and '保安' not in x['title']:
        return '指名競争入札'
    return '随意契約'

def build(fy):
    L = items(fy)
    st = W.order_style(fy)
    pref = wareki(fy, False)
    if st == 'c1':
        name = '%sあさぎ市清掃工場_発注台帳_%s_契約係c1様式.xlsx' % (KAKU, pref)
        stem = ans.start(name)
        wb = Workbook(); ws = wb.active; ws.title = '契約台帳'
        fn = 'ＭＳ Ｐゴシック'
        widths(ws, {'A': 5, 'B': 9, 'C': 46, 'D': 12, 'E': 26, 'F': 14, 'G': 20, 'H': 10, 'I': 14, 'J': 24})
        merge(ws, 1, 1, 1, 10, KAKU + '%s　契約台帳（環境部清掃工場）' % pref, F(14, True, name=fn), AL, border=None)
        merge(ws, 2, 1, 2, 10, '契約係 c1　（金額は消費税込み、単価契約は予定総額）', F(9, name=fn), AL, border=None)
        hd = ['No', '契約日', '件　名', '契約方法', '相手方', '契約金額', '履行期間', '節', '支払済額', '備考']
        for i, x in enumerate(hd):
            put(ws, 4, i + 1, x, F(9, True, name=fn), AC, B_ALL, fill('D9D9D9'))
        r = 5
        tot = 0
        LS = sorted(L, key=lambda x: x['cd'])
        for i, x in enumerate(LS):
            setsu = SETSU_OLD[x['sk']]
            vals = [i + 1, wa(x['cd']).replace('H0', 'H'), x['title'], method(x), x['vendor'], x['amt'], '%s〜%s' % (wa(x['s']), wa(x['e'])), setsu[1], x['paid'],
                    ('〃' if i > 0 and x['note'] and x['note'] == LS[i - 1]['note'] else x['note']) + (('　' + x['tanka']) if x['tanka'] else '')]
            for j, v in enumerate(vals):
                put(ws, r, j + 1, v, F(9, name=fn), AC if j in (0, 1, 3, 7) else (AR if j in (5, 8) else AL), B_ALL, fmt='#,##0' if j in (5, 8) else None)
            ans.rec(stem, 'item', dict(no=i + 1, date=x['cd'].isoformat(), title=x['title'], method=vals[3], vendor=x['vendor'], amount=x['amt'], setsu=setsu[1], paid=x['paid']),
                    sheet='契約台帳', r=r, works=x['works'])
            tot += x['amt']
            r += 1
        merge(ws, r, 1, r, 5, '計', F(9, True, name=fn), AR)
        put(ws, r, 6, tot, F(9, True, name=fn), AR, box(t=dbl, b=thin, l=thin, r=thin), fmt='#,##0')
        ans.rec(stem, 'total', dict(amount=tot), sheet='契約台帳', r=r)
        page(ws, landscape=True, paper=9, area='A1:J%d' % r, title_rows='4:4')
        wb.save(ans.FILES + '/' + name)
        return
    if st == 'old':
        name = '%sあさぎ市清掃工場_発注台帳_%s_旧財務会計システム出力.csv' % (KAKU, pref)
        stem = ans.start(name)
        hdr = ['年度', '所属コード', '所属名', '契約番号', '件名', '契約区分', '契約方法', '款', '項', '目', '節', '節名称', '細節', '業者コード', '契約相手方',
               '契約日', '履行期限自', '履行期限至', '予定価格', '契約金額', 'うち消費税額', '変更後契約金額', '支払済額', '単価契約内容', '備考']
        title = [KAKU + '契約台帳一覧', '所属：0623 環境部清掃工場', '年度：%s' % pref, '出力日：%s' % wa(D(fy + 1, 6, 4)), '出力者：契約係']
        rows = []
        for i, x in enumerate(sorted(L, key=lambda x: x['cd'])):
            s_ = SETSU_OLD[x['sk']]
            rows.append([str(fy - 1988), '0623', '環境部清掃工場', '%02d-0623-%04d' % (fy - 1988, 100 + i * 3), x['title'].replace(' ', ''), x['kind'], method(x), '04', '02', '03',
                         s_[0], s_[1], '01', vcode(x['vendor']), x['vendor'], wa(x['cd']), wa(x['s']), wa(x['e']), int(round(x['amt'] * 1.03, -3)), x['amt'],
                         int(round(x['amt'] * tax(fy) / (1 + tax(fy)))), '', x['paid'], x['tanka'], x['note']])
        with open(ans.FILES + '/' + name, 'w', encoding='cp932', newline='') as f:
            w = csv.writer(f, lineterminator='\r\n'); w.writerow(title); w.writerow(hdr); w.writerows(rows)
        ans.rec(stem, 'parent', {'title': title[0]}, r=1)
        for i, (row, x) in enumerate(zip(rows, sorted(L, key=lambda x: x['cd']))):
            ans.rec(stem, 'item', dict(title=row[4], vendor=row[14], date=row[15], amount=row[19], setsu=row[11], paid=row[22]), r=i + 3, works=x['works'])
        return
    name = '%sあさぎ市清掃工場_発注台帳_%s_新財務会計システム出力.csv' % (KAKU, pref)
    stem = ans.start(name)
    hdr = ['会計年度', '契約管理番号', '担当所属', '件名', '契約種別', '契約方法', '会計区分', '款', '項', '目', '節', '節名称', '細節名称', '相手方番号', '相手方名称',
           '契約締結日', '履行期間開始', '履行期間終了', '契約金額（税込）', 'うち消費税等', '変更回数', '変更後契約金額', '支出済額', '単価／総価', '予定数量', '単位', '契約単価（税抜）', '摘要']
    out_d = D(2026, 11, 13) if fy == 2026 else D(fy + 1, 5, 31)
    title = [KAKU + '契約情報一覧（会計年度：%d　担当所属：環境部 清掃工場）' % fy, '出力日時：%s 09:12' % out_d.strftime('%Y/%m/%d')]
    rows = []
    LL = sorted(L, key=lambda x: x['cd'])
    for i, x in enumerate(LL):
        s_ = SETSU_NEW[x['sk']]
        q = unit = up = ''
        if x['tanka']:
            import re
            m = re.match(r'([\d,]+)円/(\S+?)（税抜）予定数量([\d.]+)', x['tanka']) or re.match(r'([\d,]+)円/(\S+?)×([\d.]+)', x['tanka'])
            if m:
                up, unit, q = m.group(1).replace(',', ''), m.group(2).replace('本', '本'), m.group(3)
        kind = {'工事': '工事請負', '物品（単価）': '物品購入', '物品': '物品購入', '修繕': '修繕・物品'}.get(x['kind'], x['kind'])
        rows.append([str(fy), 'R%02d-%05d' % (fy - 2018, 2300 + i * 7), '環境部 清掃工場', x['title'], kind, method(x), '一般会計', '04', '02', '03', s_[0], s_[1], s_[2],
                     '00' + vcode(x['vendor']), x['vendor'], x['cd'].strftime('%Y/%m/%d'), x['s'].strftime('%Y/%m/%d'), x['e'].strftime('%Y/%m/%d'), x['amt'],
                     int(round(x['amt'] * 10 / 110)), 0, '', x['paid'], '単価' if x['kind'] == '物品（単価）' else '総価', q, unit, up, x['note']])
    with open(ans.FILES + '/' + name, 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.writer(f, lineterminator='\r\n', quoting=csv.QUOTE_ALL); w.writerow(title); w.writerow(hdr); w.writerows(rows)
    ans.rec(stem, 'parent', {'title': title[0]}, r=1)
    for i, (row, x) in enumerate(zip(rows, LL)):
        ans.rec(stem, 'item', dict(title=row[3], vendor=row[14], date=row[15], amount=row[18], setsu=row[11], paid=row[22]), r=i + 3, works=x['works'])
