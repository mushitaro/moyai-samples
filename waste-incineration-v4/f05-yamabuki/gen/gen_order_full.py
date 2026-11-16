# -*- coding: utf-8 -*-
"""発注台帳：支出負担行為一覧表（組合 財務会計システムの Excel 出力）R3〜R7。
R5 から y1 が「工事件名」列を手で挿入（R6 から y2 が引き継ぎ、略し方が変わる）。金額は税込の契約額で原価ではない。"""
import os
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter as L
from common import *
import works as W
import gen_order as GO
import rec

COMMON = lambda y: [
    ('委託料', '委託料', '第二工場電気保安業務委託', ELEC, '随意契約', 1452000, 0, None, [('毎月', 1452000)], '', '', [f'E{y}']),
    ('委託料', '委託料', '第二工場消防用設備等点検業務委託', FIRE, '随意契約', 385000, 0, None, [(f'{y+1}-03-22', 385000)], '', '', [f'F{y}']),
    ('委託料', '委託料', '主灰・飛灰資源化処理業務委託（単価契約）', CEMENT, '随意契約', 24640000 + 330000 * (y - 2022), 0, None, [('毎月', 23918400 + 300000 * (y - 2022))], '', '', []),
    ('需用費', '消耗品費', '消石灰購入（単価契約）', SHOP, '指名競争', 9350000 + 110000 * (y - 2022), 0, None, [('毎月', 8512300 + 98000 * (y - 2022))], '', '', []),
    ('需用費', '消耗品費', '活性炭購入（単価契約）', SHOP, '指名競争', 4180000 + 55000 * (y - 2022), 0, None, [('毎月', 3689200 + 51000 * (y - 2022))], '', '', []),
    ('需用費', '燃料費', '灯油購入（単価契約）', GO.KOHARU, '見積合わせ', 3960000 + 110000 * (y - 2022), 0, None, [('毎月', 3804900 + 90000 * (y - 2022))], '', '', []),
    ('委託料', '委託料', '第二工場排ガス等測定業務委託', ANALYS, '指名競争', 2970000, 0, None, [(f'{y+1}-03-22', 2970000)], '', '', [f'A{y}']),
    ('委託料', '委託料', '第二工場ダイオキシン類作業環境測定業務委託', ANALYS, '随意契約', 660000, 0, None, [(f'{y+1}-03-22', 660000)], '', '', []),
    ('委託料', '委託料', '第二工場非常用発電機点検業務委託', GO.GEN, '随意契約', 253000, 0, None, [(f'{y}-12-22', 253000)], '', '', [f'G{y}']),
    ('需用費', '消耗品費', '排水処理薬品購入（凝集剤・苛性ソーダ）', SHOP, '見積合わせ', 429000 + 11000 * (y - 2022), 0, None, [(f'{y}-11-30', 429000 + 11000 * (y - 2022))], '', '', []),
]


def teiki_row(y, kouji):
    c = W.TEIKI[y]
    chg = 1320000 if y == 2022 else 0
    note = '定修' + ('　変更：2号灰押出装置摺動板取替' if y == 2022 else '')
    return ('委託料', '委託料', f'令和{y-2018}年度やまぶき広域処理組合第二工場定期整備業務委託', MAKER, '随意契約（特命）', c['amount'], chg,
            '2022-11-10' if chg else None, [(c['period'][1].replace(day=31 if c['period'][1].month == 3 else c['period'][1].day).isoformat() if False else f'{y+1}-03-29', c['amount'] + chg)],
            kouji, note, [f'T{y}'])


def rows_for(y):
    R = []
    if y in (2021, 2022, 2023):
        chg = {2021: 0, 2022: 2640000, 2023: 2860000}[y]
        R.append(('委託料', '委託料', '第二工場運転業務委託（長期継続 令和3〜5年度）', OPCO, '長期継続', 118800000, chg, f'{y}-12-20' if chg else None,
                  [('毎月', 118800000 + chg)], '', ('1号停止中 単独運転延長分 変更' if y == 2022 else ('2号停止中 単独運転延長分 変更' if y == 2023 else '')), []))
    else:
        R.append(('委託料', '委託料', '第二工場運転業務委託（長期継続 令和6〜8年度）', OPCO, '長期継続', 124300000, 0, None, [('毎月', 124300000)], '', '', []))
    R += COMMON(y)
    kj = {2021: '', 2022: '', 2023: 'R5定期整備', 2024: '定期整備(R6)', 2025: '定期整備(R7)'}[y]
    R.append(teiki_row(y, kj))
    if y == 2021:
        R += [('委託料', '委託料', '第二工場精密機能検査業務委託', PREC, '随意契約', 1980000, 0, None, [('2022-02-25', 1980000)], '', '', ['P2021']),
              ('委託料', '委託料', '計量機定期検査立会・調整業務委託', SCALE, '見積合わせ', 198000, 0, None, [('2021-10-29', 198000)], '', '', ['S2021']),
              ('役務費', '手数料', '計量機定期検査手数料（2台）', GO.TSUBAME, '－', 30800, 0, None, [('2021-09-30', 30800)], '', '', ['S2021']),
              ('需用費', '修繕料', 'ごみ投入扉No.2リミットスイッチ交換修理', GO.LOCAL, '見積合わせ', 121000, 0, None, [('2021-08-31', 121000)], '', '', [])]
    if y == 2022:
        R += [('工事請負費', '工事請負費', 'やまぶき広域処理組合第二工場基幹的設備改良工事（令和4〜5年度）', MAKER, '一般競争（総合評価）', 803000000, 0, None,
               [('2022-07-22', 321200000), ('2023-03-24', 481800000)], '', '継続費 R4年割額（総額 1,947,000,000）　前払金・出来高', ['KIKAN']),
              ('需用費', '修繕料', '計量機表示器修理', SCALE, '見積合わせ', 231000, 0, None, [('2022-08-31', 231000)], '', '', ['SC2022']),
              ('委託料', '委託料', '可燃ごみ処理業務委託（基幹的設備改良工事期間中・単価契約）', GO.AOSAGI, '随意契約', 18700000, 0, None, [('毎月', 17061500)], '', '基幹改良中', []),
              ('需用費', '消耗品費', 'ごみクレーン用ワイヤロープ購入', SHOP, '見積合わせ', 495000, 0, None, [('2022-10-31', 495000)], '', '1号クレーン用 2本　定修で支給', []),
              ('役務費', '手数料', 'クレーン性能検査手数料（2基）', GO.HIKARI, '－', 66000, 0, None, [('2022-10-20', 66000)], '', '', ['C2022']),
              ('需用費', '修繕料', 'プラットホーム出口扉センサ修理', GO.LOCAL, '見積合わせ', 88000, 0, None, [('2022-11-15', 88000)], '', '', []),
              ('需用費', '修繕料', 'プラットホーム入口扉リモコン受信機修理', GO.LOCAL, '見積合わせ', 143000, 0, None, [('2023-01-20', 143000)], '', '', [])]
    if y == 2023:
        R += [('工事請負費', '工事請負費', 'やまぶき広域処理組合第二工場基幹的設備改良工事（令和4〜5年度）', MAKER, '一般競争（総合評価）', 1144000000, 7700000, '2024-01-12',
               [('2023-12-22', 572000000), ('2024-03-28', 579700000)], '基幹改良(R4-5)', '継続費 R5年割額　変更：2号煙道減肉部補修追加（総額 1,954,700,000）', ['KIKAN']),
              ('委託料', '委託料', '計量機定期検査立会・調整業務委託', SCALE, '見積合わせ', 198000, 0, None, [('2023-10-31', 198000)], '計量機定期検査', '', ['S2023']),
              ('役務費', '手数料', '計量機定期検査手数料（2台）', GO.TSUBAME, '－', 30800, 0, None, [('2023-09-29', 30800)], '計量機定期検査', '', ['S2023']),
              ('委託料', '委託料', '可燃ごみ処理業務委託（基幹的設備改良工事期間中・単価契約）', GO.AOSAGI, '随意契約', 20350000, 0, None, [('毎月', 19432600)], '基幹改良', '基幹改良中', []),
              ('需用費', '修繕料', '2号ごみクレーン巻上ブレーキライニング交換修理', GO.CRANE, '見積合わせ', 341000, 0, None, [('2023-10-31', 341000)], 'クレーン修繕', '', ['CR2023']),
              ('委託料', '委託料', 'DCS切替期間運転支援業務委託（夜間立会）', OPCO, '随意契約', 1100000, 0, None, [('2024-03-15', 1100000)], '基幹改良 DCS切替', '', []),
              ('需用費', '修繕料', 'プラットホーム入口扉チェーン交換修理', GO.LOCAL, '見積合わせ', 165000, 0, None, [('2024-03-08', 165000)], '', '', [])]
    if y == 2024:
        R += [('委託料', '委託料', '第二工場精密機能検査業務委託', PREC, '随意契約', 2090000, 0, None, [('2025-02-28', 2090000)], '精密機能検査(R6)', '', ['P2024']),
              ('役務費', '手数料', 'クレーン性能検査手数料（2基）', GO.HIKARI, '－', 66000, 0, None, [('2024-10-18', 66000)], 'クレーン性能検査', '', ['C2024']),
              ('需用費', '消耗品費', 'ごみクレーン用ワイヤロープ購入', SHOP, '見積合わせ', 528000, 0, None, [('2024-10-31', 528000)], '定期整備(R6) 支給品', '1号クレーン用 2本', []),
              ('需用費', '修繕料', '計量機（出口）ロードセル交換修理', SCALE, '見積合わせ', 385000, 0, None, [('2024-09-30', 385000)], '計量機修繕', '器差不良', ['L2024'])]
    if y == 2025:
        R += [('工事請負費', '工事請負費', '第二工場区分開閉器（PAS）更新工事', '【架空】有限会社 こもれび電設', '見積合わせ', 1430000, 0, None, [('2025-11-28', 1430000)], 'PAS更新(R7)', '全停電日に施工', ['PAS2025']),
              ('需用費', '修繕料', '無停電電源装置蓄電池交換修理', '【架空】さざなみ電機株式会社', '見積合わせ', 990000, 0, None, [('2025-12-26', 990000)], 'UPS蓄電池(R7)', '', ['UPS2025']),
              ('委託料', '委託料', '計量機定期検査立会・調整業務委託', SCALE, '見積合わせ', 209000, 0, None, [('2025-10-31', 209000)], '計量機定期検査(R7)', '', ['S2025']),
              ('役務費', '手数料', '計量機定期検査手数料（2台）', GO.TSUBAME, '－', 30800, 0, None, [('2025-09-30', 30800)], '計量機定期検査(R7)', '', ['S2025']),
              ('需用費', '消耗品費', 'ごみクレーン用ワイヤロープ購入', SHOP, '見積合わせ', 539000, 0, None, [('2025-10-10', 539000)], '定期整備(R7) 支給品', '2号クレーン用 2本', [])]
    return R


def build(year):
    with_kouji = year >= 2023
    who = 'y1' if year <= 2023 else 'y2'
    tag = 'y1' if year <= 2022 else ('工事件名列追加' if year == 2023 else 'y2')
    fname = f'{FN}_発注台帳_令和{year-2018}年度支出負担行為一覧_{tag}.xlsx'
    stem = rec.start(fname)
    wb = Workbook(); ws = wb.active; ws.title = f'R{year-2018}支出負担行為一覧'
    props(wb, f'【架空】やまぶき広域処理組合 令和{year-2018}年度 支出負担行為一覧表（第二工場費）', f'組合事務局 {who}（財務会計システム出力）')
    cols = [('負担行為番号', 11), ('負担行為日', 10), ('節', 9), ('細節', 9), ('件名', 36)]
    if with_kouji:
        cols.append(('工事件名', 16))
    cols += [('相手方', 30), ('契約方法', 13), ('負担行為額', 13), ('変更後額', 13), ('支出済額', 13), ('残額', 11), ('支出日', 10), ('備考', 34)]
    for i, (t, w) in enumerate(cols, 1):
        ws.column_dimensions[L(i)].width = w
    nc = len(cols)
    put(ws, 1, 1, '【架空】支出負担行為一覧表', f(12, True, name=FONT_M))
    put(ws, 2, 1, f'会計年度：令和{year-2018:02d}年度　　会計：一般会計　　款：衛生費　項：清掃費　目：第二工場費', f(9, name=FONT_M))
    put(ws, 3, 1, f'所属：{UNION} 事務局　　出力日：令和{year-2017:02d}年05月31日　　出力者：{who}', f(9, name=FONT_M))
    r = 5
    for i, (t, w) in enumerate(cols, 1):
        put(ws, r, i, t, f(9, True, name=FONT_M), AC, fill('FFFF99') if t == '工事件名' else fill('C0C0C0'))
    if with_kouji:
        ws.cell(r, 6).comment = Comment('R5年度分から追加（y1）。システムの件名が20字で切れて工事の区別がつかないため手入力。', 'y1')
    box(ws, r, 1, r, nc)
    r += 1
    sums = {}
    rows = sorted(rows_for(year), key=lambda x: (x[8][0][0] if False else 0))
    # 負担行為日の順：単価・年間契約は 4/1
    dated = []
    for x in rows:
        d0 = {'委託料': f'{year}-04-01'}.get(x[0], f'{year}-04-01')
        if x[11] and x[11][0].startswith('T'):
            d0 = W.TEIKI[year]['period'][0].isoformat()
        elif x[11] == ['KIKAN']:
            d0 = '2022-06-27' if year == 2022 else '2023-04-03'
        elif x[8][0][0] != '毎月' and x[10] == '' and x[2].startswith(('計量機', 'クレーン性能', 'ごみクレーン用', '2号ごみ', '第二工場区分', '無停電', 'プラット', 'ごみ投入', 'DCS', '可燃ごみ', '第二工場精密', '計量機表示')):
            pd_ = x[8][0][0]
            y_, m_, d_ = map(int, pd_.split('-'))
            import datetime as dt
            d0 = (dt.date(y_, m_, d_) - dt.timedelta(days=35)).isoformat()
            if d0 < f'{year}-04-01':
                d0 = f'{year}-04-05'
        elif x[2].startswith(('可燃ごみ', 'DCS', 'ごみクレーン用')) or x[10]:
            pd_ = x[8][0][0] if x[8][0][0] != '毎月' else f'{year}-09-01'
            import datetime as dt
            y_, m_, d_ = map(int, pd_.split('-'))
            d0 = max(f'{year}-04-05', (dt.date(y_, m_, d_) - dt.timedelta(days=40)).isoformat())
        if x[2].startswith('第二工場排ガス') or x[2].startswith('第二工場ダイオキシン'):
            d0 = f'{year}-04-15'
        if x[2].startswith('第二工場非常用'):
            d0 = f'{year}-05-20'
        if x[2].startswith('可燃ごみ'):
            d0 = f'{year}-08-22'
        dated.append((d0, x))
    dated.sort(key=lambda t: t[0])
    for n, (d0, x) in enumerate(dated, 1):
        (setsu, saisetsu, name, aite, how, amt, chg, chgd, pays, kouji, note, wkeys) = x
        no = f'{year - 2018 + 0:02d}{n:05d}'
        after = amt + chg
        paid = sum(a for _, a in pays)
        payd = '月次' if pays[0][0] == '毎月' else '・'.join(GO.rdate(p) for p, _ in pays)
        vals = [no, GO.rdate(d0), setsu, saisetsu, GO.trunc(name)]
        if with_kouji:
            vals.append(kouji)
        vals += [aite, how, amt, after if chg else None, paid, after - paid, payd, (note + (f'（変更 {GO.rdate(chgd)}）' if chg else '')).strip()]
        for i, v in enumerate(vals, 1):
            numeric = isinstance(v, int)
            put(ws, r, i, v, f(9, name=FONT_M), AR if numeric else AL, nf='#,##0' if numeric else None)
        sums.setdefault(setsu, [0, 0, 0])
        sums[setsu][0] += after; sums[setsu][1] += paid; sums[setsu][2] += after - paid
        wl = []
        for k in wkeys:
            if k == 'KIKAN':
                wl += [w['id'] for w in W.WORKS if w['contract'] == 'KIKAN' and w['year'] == year]
            else:
                wl += [w['id'] for w in W.WORKS if w['contract'] == k]
        flds = dict(zip([c for c, _ in cols], vals))
        flds['負担行為日'] = d0
        flds['件名（全文）'] = name
        rec.add(stem, ws.title, r, 'item', flds, works=wl)
        r += 1
    box(ws, 6, 1, r - 1, nc, thin)
    r += 1
    names = [c for c, _ in cols]
    ia = names.index('変更後額') + 1
    lc = names.index('件名') + 1
    put(ws, r, lc, '＜節別集計＞（変更後額）', f(9, True, name=FONT_M))
    r += 1
    tot = [0, 0, 0]
    st = r
    for s in ('委託料', '工事請負費', '需用費', '役務費'):
        if s not in sums:
            continue
        put(ws, r, lc, s, f(9, name=FONT_M), AL)
        for j, v in enumerate(sums[s]):
            put(ws, r, ia + j, v, f(9, name=FONT_M), AR, nf='#,##0')
            tot[j] += v
        rec.add(stem, ws.title, r, 'subtotal', dict(節=s, 変更後額=sums[s][0], 支出済額=sums[s][1], 残額=sums[s][2], unit='円'))
        r += 1
    put(ws, r, lc, '合計', f(9, True, name=FONT_M), AL)
    for j, v in enumerate(tot):
        put(ws, r, ia + j, v, f(9, True, name=FONT_M), AR, nf='#,##0')
    rec.add(stem, ws.title, r, 'total', dict(変更後額=tot[0], 支出済額=tot[1], 残額=tot[2], unit='円'))
    box(ws, st, lc, r, ia + 2)
    put(ws, r + 2, 1, '※ 負担行為額は税込の契約額（単価契約は予定額）。変更後額は変更契約のある行のみ表示。', f(8, name=FONT_M))
    ws.freeze_panes = 'C6'
    setup_print(ws, True, 8, 1, 0, f'A1:{L(nc)}{r+2}', '5:5')
    ws.oddFooter.center.text = '&P / &N'
    wb.save(os.path.join(OUT, fname))
    return fname


def build_all():
    return [build(y) for y in range(2021, 2026)]


if __name__ == '__main__':
    print(build_all())
