# -*- coding: utf-8 -*-
"""answer/world.json を書く。"""
import json, os, datetime as dt
import world, history as H, periods as PR, f_daily, rec

def iso(d): return d.isoformat() if isinstance(d, dt.date) else d

GAPS = [
    dict(file='中長期整備計画_平成23年度版_j1様式.xlsx', what='法定点検の行に計量機の定期検査（計量法、2年）が無い', kind='法令点検の抜け'),
    dict(file='中長期整備計画_令和4年度改訂版_j2延命化計画様式.xlsx', what='精密機能検査を R4 に計画したが基幹改良と重なり未実施、R5 実施（H28・R1 から 3 年を超えた）', kind='周期の超え'),
    dict(file='中長期整備計画_令和5年度改訂版_j2延命化計画様式.xlsx', what='ボイラ設備の R7（2025）費用小計が機器の費用の和と 4,500 千円違う（◎ を動かした後の直し忘れ）', kind='費用の小計の食い違い'),
    dict(file='中長期整備計画_令和元年度版〜令和5年度改訂版（j2延命化計画様式）', what='ダイオキシン類作業環境測定の行に回数（6か月ごと＝年2回）の記載が無い。j3 の版で「年2回」と直した', kind='周期の欠け'),
    dict(file='中長期整備計画（j1・j2 の各版）', what='液体酸素設備（CE）の定期自主検査の行が無い。j3 の版で追加（確度 低・要確認）', kind='法令点検の抜け'),
    dict(file='中長期整備計画 j3 の各版', what='費用の行（非表示）に R5 版の古い見込み額が残り、R16〜R20 の費用は空', kind='費用の食い違い（隠し行）'),
]

def write(FULL, P, AK, RP, CON, MAN):
    bytag = {u['tag']: u for u in world.UNITS}
    units = []
    for u in world.UNITS:
        al = {'機器台帳': u['name'], '中長期整備計画': (f"電動機（{u['tag']}）" if u['kind'] == '電動機' else u['name'])}
        if u['kind'] == '電動機': al['電動機台帳（H23）'] = f"{u['tag']}（負荷機器 {bytag[u['parent']]['name']}）"
        if u['furn'] in (1, 2): al['日報の場所'] = f"{u['furn']}号{u['subname']}"
        units.append(dict(tag=u['tag'], name=u['name'], aliases=al, furnace={0: '共通', 1: '1号炉', 2: '2号炉'}[u['furn']], system=u['sysname'], equipment=u['subname'],
                          kind=u['kind'], parent=u['parent'], duty=u['duty'] or None, importance=u['imp'], installed=u['year'],
                          renewedByUpgrade=(world.ak_scope(u)[0] if world.ak_scope(u) else None), addedByUpgrade=u.get('ident') == 'ak_new'))
    cname = {}
    for fy, C in CON.items():
        for c in C: cname[c['key']] = c['name']
    cname['基幹改良'] = PR.AKC
    works = []
    for w in PR.WORKS:
        works.append(dict(id=w['id'], unit=w['unit'], unitName=bytag[w['unit']]['name'], type=w['type'], symbol=w['sym'], start=iso(w['start']), end=iso(w['end']),
                          manHours=f_daily.MH.get(w['id']), vendor=w['vendor'], contract=cname.get(w['contract'], w['contract']), contractKey=w['contract'], period=w['period'],
                          scheduleItem=w['item'], memo=w['desc'] or None, status='施工中' if w['end'] is None else None))
    # 他業者（財務 CSV だけ）
    n = len(works)
    for fy, C in CON.items():
        for c in C:
            for t in c['units']:
                if c['kind'] in ('物品',): continue
                n += 1
                s = H.truth(bytag[t]).get(fy, ('◎', ''))[0]
                works.append(dict(id=f'W{n:05d}', unit=t, unitName=bytag[t]['name'], type=PR.SYMW.get(s, 'OVERHAUL'), symbol=s, start=iso(c['s']), end=iso(c['e']),
                                  manHours=None, vendor=c['vendor'], contract=c['name'], contractKey=c['key'], period=None, scheduleItem=None,
                                  memo='業者の作業報告書は紙で市が保管（電子化なし）。資料は財務 CSV の契約行だけ'))
    have = {(w['unit'], int(w['start'][:4]) if int(w['start'][5:7]) >= 4 else int(w['start'][:4]) - 1) for w in works if w['start']}
    unrec = []
    for u in world.UNITS:
        for y, (s, d) in H.truth(u).items():
            if y < 2011 or y > 2026: continue
            if (u['tag'], y) in have: continue
            reason = '月例点検・日常保全の中で行い、定修・全停の工程表と日報には載らない（点検記録表は紙）'
            if s == '★': reason = '基幹改良の範囲（工程表は基幹改良の該当年度のシート）'
            if y == 2026 and u['furn'] == 1: reason = '令和8年度 1号炉定修（施工中）の後半に予定。11/13 時点で未着手'
            unrec.append(dict(unit=u['tag'], unitName=u['name'], year=y, symbol=s, type=PR.SYMW.get(s), memo=d or None, reason=reason))
    links = rec.links()
    clinks = [l for l in links if l['work'] and not str(l['work']).startswith('W')]
    wlinks = [l for l in links if l['work'] and str(l['work']).startswith('W')]
    contracts = [dict(fy=fy, key=c['key'], name=c['name'], kind=c['kind'], vendor=c['vendor'], method=c['method'], contractDate=iso(c['cdate']), start=iso(c['s']), end=iso(c['e']),
                      amount=c['amount'], changedAmount=c['change'], paidInFY=c['paid'], note='金額は契約額。工事の実績原価（外注・部品・社内工数の内訳）は財務資料には無い')
                 for fy, C in CON.items() for c in C]
    profile = dict(id='f10', name='【架空】あかね台資源化センター（あかね台市）', completed='2006-03', capacity='120 t/24h × 2 炉 ＝ 240 t/日',
                   furnace='シャフト炉式ガス化溶融炉（コークスベッド、酸素富化送風）、全連続', boiler='4MPa・400℃・16 t/h、発電 5,000kW、誘引通風機 25,000 m³N/h',
                   upgrade='2021–2024 年度 基幹的設備改良（炉体耐火物・羽口・PSA・DCS）', operation='市直営運転',
                   maintenance='千草環境システム（架空）への単年度「年間保守点検整備業務委託」（全炉停止・2号炉定修・1号炉定修、報告書は 3 分冊）＋補修請負＋基幹改良（4 か年請負）',
                   period='2011–2026 年度（16 年）。2010 年度以前は未記録', asOf='2026-11-13（2026 年度は計画＋上期実績、1号炉定修は施工中）',
                   units='915 台（1号炉 260、2号炉 260、共通 395。機器＋電動機を各 1 台、電動・調節弁、主要計装、盤。手動弁は除く。2025 年度までの旧台帳は 912 台＝基幹改良の新設 3 台が無い）',
                   staff=dict(city='j1 2011–2015、j2 2016–2023、j3 2024–', maker='SV p1 2011–2017、p2 2018–、補修工事 現場代理人 q1、基幹改良 現場代理人 q2'))
    out = dict(profile=profile, units=units, contracts=contracts, works=works, links=wlinks, contractLinks=clinks, unrecordedWorks=unrec, planGaps=GAPS,
               notes=['見本（f10/out）の H30・R3 の工程表・日報は、本作りで見本に無かった機器の作業のまとまり（E 番号）を足した。手書きの日の行は見本のまま。',
                      '日曜は休工（日報なし）。市運転班の作業（立下げ・乾燥・立上げ）は日報に載らない。'])
    json.dump(out, open(os.path.join(FULL, 'answer', 'world.json'), 'w'), ensure_ascii=False, indent=0)
    return out
