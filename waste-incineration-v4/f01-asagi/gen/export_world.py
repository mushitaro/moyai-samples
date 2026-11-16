# -*- coding: utf-8 -*-
"""answer/world.json：機器・実際の作業・資料との結び・どこにも書いていない作業・計画の欠け・施設の profile。"""
import json, datetime as dt
import ans
import f01_truth as TR
from f01_master import EQ, SYS, LEGAL, legal_years, TODAY
import build_plan as PL

def conv(o):
    if isinstance(o, (dt.date, dt.datetime)):
        return o.isoformat()
    return str(o)

PROFILE = dict(id='F01', name='【架空】あさぎ市清掃工場（あさぎ市 環境部 清掃工場）', completed='1997-03', capacity='110 t/24h × 2 炉 = 220 t/日', furnace='並行揺動式ストーカ、全連続燃焼式',
               boiler='3 MPa・300 ℃・12 t/h ×2', power='抽気復水タービン 2,600 kW', idf='33,000 m³N/h', crane='ごみクレーン 4.0 m³ ×2（常用1・予備1）',
               gas_treatment='減温塔＋消石灰・活性炭吹込＋バグフィルタ、アンモニア水による無触媒脱硝',
               history=['2000–2001 ダイオキシン類恒久対策（EP→BF ほか）', '2010 ごみピット放水銃設置', '2012 DCS 更新', '2017–2019 基幹的設備改良（目標 2034 年度まで延命）'],
               contracts='運転：1997–2005 市直営、2006– 運転業務委託（あさぎ環境運転サービス）。整備：メーカー（汐見重工環境）へ単年度の定期整備業務委託。補修：指名競争の補修工事（請負）年 0〜4 件',
               period='1997–2026 年度（2026 年度は 11 月 13 日時点）', equipment_count=445, counting='主機 1 台（電動機は付属）、電動弁・空気作動弁は各 1 台、盤は 1 面 1 台。手動弁・計装は含めない',
               people=dict(city=['a1（〜2007）', 'a2（2008–2020）', 'a3（2021–）'], maker_site=['m1（1997–2009）', 'm2（2010–2018）', 'm3（2019–）'], maker_report=['r1（1997–2014）', 'r2（2015–）'], contract='c1（1997–2008）'),
               today='2026-11-13')

def main():
    names = {}
    for stem, d in ans.A.items():
        for rc in d['records']:
            f = rc['fields']
            pass
    for L in ans.LINKS:
        if L['code']:
            pass
    equip = {}
    for d in EQ:
        e = TR.EQUIP[d['code']]
        equip[d['code']] = dict(code=d['code'], ledger_name=d['name'], tag=d['tag'], old_code=d.get('old_code'), furnace=d['fname'], system=SYS[d['sys']], kind=d['kind'],
                                role=d['role'] or None, inst_fy=d['inst'], importance=d['imp'], method=d['method'], names={})
    # 資料での呼び名（台帳・計画・報告書の記録から）
    for stem, d in ans.A.items():
        for rc in d['records']:
            f = rc['fields']
            nm = f.get('name') or (f.get('path', '').split('/')[-1] if f.get('path') else None)
            code = f.get('code')
            link = None
            if code in equip:
                link = code
            else:
                for L in []:
                    pass
            if link and nm:
                equip[link]['names'].setdefault(d['file'], nm)
    for L in ans.LINKS:
        if L['code'] in equip:
            pass
    linked = {}
    for L in ans.LINKS:
        for w in L['works']:
            linked.setdefault(w, []).append({k: v for k, v in L.items() if k in ('file', 'sheet', 'r', 'seq', 'image') and v is not None})
    works = []
    unrecorded = []
    for w in TR.WORKS:
        x = {k: w[k] for k in ('id', 'fy', 'code', 'type', 'kind_ja', 'text', 'start', 'end', 'manhours', 'vendor', 'contract', 'contract_type', 'status')}
        x['manDays'] = round(w['manhours'] / 8.0, 2)
        x['docs'] = linked.get(w['id'], [])
        works.append(x)
        if not x['docs']:
            reason = {'R': '運転中の突発故障を運転委託先（2005 年度以前は市運転班）が直営で修理。部品は在庫品で、契約・帳票が残らない'}.get(w['contract_type'])
            if not reason and w['fy'] == 2026:
                reason = 'R8 2号炉定修は施工中（記録時点 2026-11-13）。報告書・発注の締めは未作成'
            if not reason:
                reason = '資料の行と結べなかった（要確認）'
            unrecorded.append(dict(id=w['id'], code=w['code'], fy=w['fy'], text=w['text'], reason=reason))
    legal = []
    for it in LEGAL:
        for y in legal_years(it, 1997, 2026):
            if it[0] == 'LG10' and y < 2001:
                continue
            i = '%s-%d' % (it[0], y)
            legal.append(dict(id=i, fy=y, name=it[1], basis=it[2], targets=it[3], times_per_year=it[5], type='INSPECT', certainty=it[7],
                              status='完了' if y <= 2025 else '一部実施（上期）', docs=linked.get(i, [])))
    world = dict(profile=PROFILE, equipment=equip, works=works, legal_inspections=legal, unrecorded_works=unrecorded, plan_gaps=PL.GAPS,
                 doc_links=ans.LINKS, generated=dict(today='2026-11-13', seed='md5(タグ・年度)由来の固定乱数'))
    json.dump(world, open(ans.BASE + '/answer/world.json', 'w'), ensure_ascii=False, indent=1, default=conv)
    print('world: equip', len(equip), 'works', len(works), 'unrecorded', len(unrecorded), 'gaps', len(PL.GAPS))
