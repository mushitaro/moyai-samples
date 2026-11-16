# -*- coding: utf-8 -*-
"""F10【架空】あかね台資源化センター 本作り：全資料・正解・説明を作り直す。  python3 run_full.py [--png]"""
import os, sys, glob, json, datetime as dt
HERE = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, HERE)
import world, history as H, periods as PR, jobs, daily_data as DD
import f_sched, f_daily, f_report, f_plan, f_ledger, f_order, rec

FULL = os.path.dirname(HERE)
OUT = os.path.join(FULL, 'files')
PFX = '【架空】あかね台資源化センター_'
MAN = []   # (file, 資料, 年度, 様式, 説明)

def path(name): return os.path.join(OUT, PFX + name)
def reg(fn, kind, fy, style, desc): MAN.append(dict(file=os.path.basename(fn), 資料=kind, 年度=fy, 様式=style, 説明=desc))

def main():
    os.makedirs(OUT, exist_ok=True)
    for f in glob.glob(os.path.join(OUT, '*')): os.remove(f)
    P = PR.build(); AK = PR.ak_periods(); RP = PR.repair_periods(P)
    PR.WORKS.clear()
    PR.assign_works(P); PR.assign_works(AK); PR.assign_works(RP)
    allp = P + AK + RP
    f_order.stopdays(allp)
    by = {p['key']: p for p in allp}
    # ---------------------------------------------------------- 工程表・日報・報告書（年間保守）
    for fy in range(PR.FY0, PR.FY1 + 1):
        ps = [by[f'{fy}-0'], by[f'{fy}-2'], by[f'{fy}-1']]
        sv = PR.sv(fy); city = PR.city_staff(fy)
        title = f'{PR.wa(fy)} {PR.FACN} 年間保守点検整備業務委託'
        sty = {'p1': 'p1様式', 'p2a': 'p2H30様式', 'p2b': 'p2R3様式'}[f_sched.fmt(fy)]
        for p, sh in zip(ps, ['全炉停止', '2号炉定修', '1号炉定修']): p['sheet'] = sh
        def other_blocks(p, fy=fy):
            bl = []
            if p['kind'] == '全停': bl.append(('受変電設備保守点検業務委託（市の別契約）', [PR.mk(o[1], o[2], [PR.unit(0, o[3])['tag']], o[4], '', o[5], o[6], o[7], o[8], 'ref') for o in p['other']]))
            for q in allp:
                if q['kind'] in ('基幹改良', '補修工事') and q['fy'] == fy and q['start'] >= p['start'] and q['end'] <= p['end'] + dt.timedelta(10) and q['key'] != p['key']:
                    nm = PR.AKC if q['kind'] == '基幹改良' else q['title'].replace(PR.FACN, 'あかね台資源化センター')
                    bl.append((nm, q['items']))
            return bl
        fn = path(f'工程表_{PR.wa(fy)}年間保守点検整備業務委託_{sty}.xlsx')
        f_sched.write(fn, ps, title, '', f'{PR.wa(fy)}', f'SV {sv}', other_blocks=other_blocks)
        reg(fn, '工程表', fy, sty, f'年間保守：全炉停止・2号炉定修・1号炉定修の 3 シート')
        dsty = {'p1a': 'p1班別様式', 'p1b': 'H25様式', 'p2': 'R1様式'}[f_daily.fmt(fy)]
        for p in ps:
            hand = None
            if p['key'] == '2018-1': hand = {x['d']: x for x in DD.H30}
            if p['key'] == '2021-1': hand = {x['d']: x for x in DD.R3}
            fn = path(f"作業日報_{PR.wa(fy)}年間保守点検整備業務委託{p['sheet']}_{dsty}.xlsx")
            if f_daily.write(fn, p, title, sv, city, hand=hand): reg(fn, '作業日報', fy, dsty, f"{p['part']}")
            if p.get('status') == '施工中': continue
            book = {0: '第3分冊　全炉停止（共通設備）', 2: '第1分冊　2号炉 定期整備', 1: '第2分冊　1号炉 定期整備'}[int(p['key'].split('-')[1])]
            rsty = 'p1様式' if f_report.fmt(fy) == 'p1' else 'R2ツール様式'
            fn = path(f"工事報告書_{PR.wa(fy)}年間保守点検整備業務委託{book.split('　')[0]}_{rsty}.pdf")
            f_report.write(fn, p, title, sv, book, city)
            reg(fn, '工事報告書', fy, rsty, book)
    # ---------------------------------------------------------- 基幹改良
    for p, sh in zip(AK, ['R3 先行作業', 'R4 1号炉', 'R5 2号炉', 'R5 酸素発生装置', 'R6 DCS']): p['sheet'] = sh; p['fmt'] = 'p2b'
    fn = path('工程表_基幹的設備改良工事（令和3〜6年度）_p2R3様式.xlsx')
    f_sched.write(fn, AK, PR.AKC, '契約 令和3年7月5日〜令和7年3月25日（継続費）', '令和3〜6年度', '現場代理人 q2')
    reg(fn, '工程表', 2021, 'p2R3様式', '基幹改良：年度ごとの施工期間を 1 シートずつ')
    for p in AK:
        fn = path(f"作業日報_基幹的設備改良工事{p['sheet'].replace(' ', '')}_R1様式.xlsx")
        if f_daily.write(fn, dict(p, dfmt='p2'), PR.AKC, 'q2', PR.city_staff(p['fy'])): reg(fn, '作業日報', p['fy'], 'R1様式', p['part'])
        fn = path(f"工事報告書_基幹的設備改良工事{p['sheet'].replace(' ', '')}_R2ツール様式.pdf")
        f_report.write(fn, dict(p, rfmt='tool'), PR.AKC, '現場代理人 q2', '出来形報告' if p['fy'] < 2024 else '完成報告（DCS 更新・切替）', PR.city_staff(p['fy']))
        reg(fn, '工事報告書', p['fy'], 'R2ツール様式', p['part'])
    # ---------------------------------------------------------- 補修工事
    groups = {}
    for p in RP: groups.setdefault(p['contract'], []).append(p)
    for ck, ps in groups.items():
        c = next(x for x in PR.REPAIRS if x['key'] == ck); fy = c['fy']
        nm = c['name']
        for p in ps: p['sheet'] = (p['part'] or '工程')[:20]
        sty = {'p1': 'p1様式', 'p2a': 'p2H30様式', 'p2b': 'p2R3様式'}[f_sched.fmt(fy)]
        fn = path(f'工程表_{nm.replace(PR.FACN + " ", "").replace(" ", "")}_{sty}.xlsx')
        f_sched.write(fn, ps, nm, '', PR.wa(fy), f"現場代理人 {c.get('qa', 'q1')}")
        reg(fn, '工程表', fy, sty, '補修工事（請負）')
        dsty = {'p1a': 'p1班別様式', 'p1b': 'H25様式', 'p2': 'R1様式'}[f_daily.fmt(fy)]
        for p in ps:
            hand = {x['d']: x for x in DD.TAP} if ck == 'R2021' else None
            fn = path(f"作業日報_{nm.replace(PR.FACN + ' ', '').replace(' ', '')}{p['part'] if len(ps) > 1 else ''}_{dsty}.xlsx")
            if f_daily.write(fn, dict(p, part=p['part'] if len(ps) > 1 else ''), nm, c.get('qa', 'q1'), PR.city_staff(fy), hand=hand, start=p['start']):
                reg(fn, '作業日報', fy, dsty, '補修工事（請負）')
        comb = dict(ps[0]); comb['items'] = sum([p['items'] for p in ps], []); comb['start'] = min(p['start'] for p in ps); comb['end'] = max(p['end'] for p in ps)
        comb['part'] = '・'.join(p['part'] for p in ps if p['part']) or '補修工事'
        rsty = 'p1様式' if f_report.fmt(fy) == 'p1' else 'R2ツール様式'
        fn = path(f'工事報告書_{nm.replace(PR.FACN + " ", "").replace(" ", "")}_{rsty}.pdf')
        f_report.write(fn, comb, nm, f"現場代理人 {c.get('qa', 'q1')}", '工事完了報告書', PR.city_staff(fy))
        reg(fn, '工事報告書', fy, rsty, '補修工事（請負）')
    # ---------------------------------------------------------- 発注台帳
    CON = {}
    for fy in range(PR.FY0, PR.FY1 + 1):
        C = f_order.contracts(fy, allp, None); CON[fy] = C
        if fy <= 2016:
            fn = path(f'発注台帳_{PR.wa(fy)}_財務会計旧システムCSV.csv'); f_order.csv_old(fn, fy, C); reg(fn, '発注台帳', fy, '財務会計 旧（支出命令一覧）', '')
        else:
            fn = path(f'発注台帳_{PR.wa(fy)}_財務会計新システムCSV.csv'); f_order.csv_new(fn, fy, C); reg(fn, '発注台帳', fy, '財務会計 新（契約・負担行為一覧）', '')
        if fy <= 2021:
            fn = path(f'発注台帳_{PR.wa(fy)}_副資材単価契約台帳旧様式.xlsx'); f_order.sub_old(fn, fy, C); reg(fn, '発注台帳', fy, '副資材台帳 旧（月 1 行）', '')
        else:
            fn = path(f'発注台帳_{PR.wa(fy)}_副資材単価契約台帳新様式.xlsx'); f_order.sub_new(fn, fy, C); reg(fn, '発注台帳', fy, '副資材台帳 新（納入 1 回 1 行）', '')
    # ---------------------------------------------------------- 計画（作業のつなぎ用）
    wb_unit = {}
    for w in PR.WORKS: wb_unit.setdefault(w['unit'], []).append(w)
    def wmap(made, lo, hi):
        m = {}
        for t, ws_ in wb_unit.items():
            m[t] = [w['id'] for w in ws_ if w['start'] and lo <= fyof(w['start']) < made]
        return m
    for made in (2011, 2013):
        fn = path(f'中長期整備計画_{PR.wa(made)}{"版" if made == 2011 else "改訂"}_j1様式.xlsx')
        f_plan.CUR['wmap'] = wmap(made, 2011, 2021); f_plan.j1(fn, made); reg(fn, '中長期整備計画', made, 'j1様式', '')
    for made in (2016, 2017, 2018):
        fn = path(f'中長期整備計画_{PR.wa(made)}改訂_j2H28様式.xlsx')
        f_plan.CUR['wmap'] = wmap(made, 2016, 2026); f_plan.j1(fn, made); reg(fn, '中長期整備計画', made, 'j2H28様式', '')
    for made in range(2019, 2024):
        fn = path(f'中長期整備計画_{PR.wa(made)}{"版" if made == 2019 else "改訂版"}_j2延命化計画様式.xlsx')
        f_plan.CUR['wmap'] = wmap(made, 2019, 2034); f_plan.old(fn, made); reg(fn, '中長期整備計画', made, 'j2延命化計画様式', '')
    for made in (2024, 2025, 2026):
        fn = path(f'中長期整備計画_{PR.wa(made)}版_j3系統別様式.xlsx')
        f_plan.CUR['wmap'] = wmap(made + 1, 2024, 2039); f_plan.new(fn, made); reg(fn, '中長期整備計画', made, 'j3系統別様式', '')
    # ---------------------------------------------------------- 機器台帳
    fm, fe = path('機器台帳_平成23年4月現在_竣工図書機器表.xlsx'), path('機器台帳_平成23年4月現在_電動機台帳.xlsx')
    f_ledger.split2011(fm, fe); reg(fm, '機器台帳', 2011, '竣工図書機器表', '統合前（機器）'); reg(fe, '機器台帳', 2011, '電動機台帳', '統合前（電動機）')
    for asof in (2012, 2016, 2020):
        fn = path(f'機器台帳_{PR.wa(asof)[:-1]}4月現在_j1統合様式.xlsx'); f_ledger.old(fn, asof); reg(fn, '機器台帳', asof, 'j1統合様式', '')
    for asof in (2025, 2026):
        fn = path(f'機器台帳_{PR.wa(asof)[:-1]}4月現在_j3設備別様式.xlsx'); f_ledger.new(fn, asof); reg(fn, '機器台帳', asof, 'j3設備別様式', '')
    # ---------------------------------------------------------- 正解
    rec.dump(os.path.join(FULL, 'answer', 'files'))
    import answers
    answers.write(FULL, P, AK, RP, CON, MAN)
    json.dump(MAN, open(os.path.join(FULL, 'answer', 'manifest.json'), 'w'), ensure_ascii=False, indent=1)
    print('files', len(MAN))

def fyof(d): return d.year if d.month >= 4 else d.year - 1

if __name__ == '__main__':
    main()
