# -*- coding: utf-8 -*-
"""中長期整備計画の各版：a1 H16（10 か年計画）、a2 H25（長寿命化総合計画）、a2 R1 改訂、a3 R3・R5・R8（星取表）。各版は作成時点の情報だけ。"""
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.chart import BarChart, Reference
from xlutil import *
import ans
from f01_master import EQ, SYS, SYS_ORDER, wareki, LEGAL, legal_years, OLD_DCS_2008
import f01_history as H
from make_plan import PRI, SYMB, KFILL, LIFE, main_kind
import make_ledger, build_sd as B
KAKU = '【架空】'
GAPS = []   # 入れた計画の欠け

def exists(d, fy):
    return d['inst'] <= fy or d['inst'] == 1996

def wy(y, era='H'):
    return ('H%d' % (y - 1988)) if era == 'H' else (('H%d' % (y - 1988)) if y < 2019 else ('R元' if y == 2019 else 'R%d' % (y - 2018)))

def path(d):
    return '%s/%s/%s' % (SYS[d['sys']], d['fname'], d['name'])

def legal_rows(asof, y0, y1, skip=()):
    out = []
    for it in LEGAL:
        if it[0] in skip:
            continue
        ys = [y for y in legal_years(it, y0, y1)]
        if it[0] == 'LG10' and asof < 2001:
            continue
        out.append((it, ys))
    return out

# ------------------------------------------------------------------ a1 H16 10 か年計画（万円）
def build_a1():
    asof, y1 = 2004, 2013
    years = list(range(asof, y1 + 1))
    name = KAKU + 'あさぎ市清掃工場_中長期整備計画_10か年整備計画_平成16年度_a1様式.xlsx'
    stem = ans.start(name, plan=True)
    wb = Workbook(); ws = wb.active; ws.title = '10か年計画'
    fn = 'ＭＳ Ｐゴシック'
    widths(ws, {'A': 14, 'B': 7, 'C': 30, 'D': 10, 'E': 6, 'F': 7})
    yc0 = 7
    for i in range(len(years)):
        ws.column_dimensions[CL(yc0 + i)].width = 7
    cm = yc0 + len(years); ws.column_dimensions[CL(cm)].width = 9; ws.column_dimensions[CL(cm + 1)].width = 30
    merge(ws, 1, 1, 1, 12, KAKU + 'あさぎ市清掃工場　10か年整備計画（平成16年度〜平成25年度）', F(14, True, name=fn), AL, border=None)
    merge(ws, 2, 1, 2, 16, '平成16年9月　施設係 a1　　金額：万円（税込・概算）　　周期：標準的な整備周期（年）　前回：直近の整備年度', F(9, name=fn), AL, border=None)
    hd = ['設備', '号炉', '機器名', '機器番号', '周期', '前回'] + [wy(y) for y in years] + ['計', '備考']
    for i, x in enumerate(hd):
        put(ws, 4, i + 1, x, F(9, True, name=fn), AC, B_ALL, fill('D9D9D9'))
    r = 5
    for s in SYS_ORDER:
        merge(ws, r, 1, r, cm + 1, '【%s】' % SYS[s], F(10, True, name=fn), AL, fl=fill('F2F2F2'))
        ans.rec(stem, 'parent', {'path': SYS[s]}, sheet=ws.title, r=r)
        r += 1
        r0 = r
        rows = [d for d in EQ if d['sys'] == s and exists(d, 2004) and not d['tag'].startswith(('DCS', 'OPS', 'EWS', 'DL-', 'AHU-01', 'AHU-02', 'FW-'))]
        olds = [dict(sys=x[0], name=x[2], tag=x[4], fname='共通' if x[4] not in ('DCS-10', 'DCS-20') else ('1号炉' if x[4] == 'DCS-10' else '2号炉'), f=0) for x in OLD_DCS_2008] if s == 'M' else []
        for d in rows + olds:
            p = H.PLAN04.get(d.get('code'), {}) if d.get('code') else {2012: [('更新', 'DCS更新（検討）', 0)]}
            tr = H.tracks_for(d) if d.get('code') else []
            cyc = round((tr[-1][1] + tr[-1][2]) / 2) if tr and d.get('method') != 'BM' else None
            last = max([y for y, v in d.get('act', {}).items() if y < asof and any(e[0] in ('整備', '補修', '更新') for e in v)] or [None]) if d.get('code') else None
            vals = [SYS[s].replace('設備', ''), d['fname'], d['name'], d['tag'], cyc, wy(last) if last else None]
            for j, v in enumerate(vals):
                put(ws, r, 1 + j, v, F(8, name=fn), AC if j in (1, 3, 4, 5) else AL, B_ALL)
            marks = {}
            tot = 0
            for i, y in enumerate(years):
                v = p.get(y)
                val = None
                if v:
                    val = int(round(sum(e[2] for e in v) / 10.0)) if d.get('code') else '更新'
                    marks[str(y)] = val
                    tot += val if isinstance(val, int) else 0
                c = put(ws, r, yc0 + i, val, F(8, name=fn), AR, Border(left=hair, right=hair, top=hair, bottom=hair), fmt='#,##0')
                if v:
                    c.fill = fill(KFILL[main_kind(v)])
            put(ws, r, cm, tot or None, F(8, name=fn), AR, B_ALL, fmt='#,##0')
            work = '、'.join(sorted(set(e[1] for v in p.values() for e in v)))[:60]
            put(ws, r, cm + 1, work or ('事後保全' if d.get('method') == 'BM' else None), F(7, name=fn), Alignment(horizontal='left', vertical='center', shrink_to_fit=True), B_ALL)
            ans.rec(stem, 'item', dict(path='%s/%s/%s' % (SYS[s], d['fname'], d['name']), code=d['tag'], cycle=cyc, cost=tot or None, unit='万円', marks=marks, work=work or None),
                    sheet=ws.title, r=r, code=d.get('code'))
            r += 1
        # 小計
        merge(ws, r, 1, r, 6, '小計', F(8, True, name=fn), AR, fl=fill('F2F2F2'))
        sub = {}
        for i, y in enumerate(years):
            v = sum(int(round(sum(e[2] for e in H.PLAN04.get(d['code'], {}).get(y, [])) / 10.0)) for d in rows)
            sub[str(y)] = v
            put(ws, r, yc0 + i, v or None, F(8, True, name=fn), AR, B_ALL, fill('F2F2F2'), fmt='#,##0')
        put(ws, r, cm, sum(sub.values()), F(8, True, name=fn), AR, B_ALL, fill('F2F2F2'), fmt='#,##0')
        ans.rec(stem, 'subtotal', dict(path=SYS[s] + '/小計', cost=sum(sub.values()), unit='万円', marks={k: v for k, v in sub.items() if v}), sheet=ws.title, r=r)
        r += 1
    # 法令点検（a1 は一部しか載せていない：クレーン性能検査・作業環境測定・精密機能検査が抜けている）
    r += 1
    merge(ws, r, 1, r, cm + 1, '【法定点検・測定】', F(10, True, name=fn), AL, fl=fill('FFF2CC'))
    ans.rec(stem, 'parent', {'path': '法定点検・測定'}, sheet=ws.title, r=r)
    r += 1
    skip = ('LG07', 'LG10', 'LG13')
    GAPS.append(dict(file=name, kind='法令点検の抜け', where='法定点検・測定の欄', what='クレーン性能検査（LG07）・ダイオキシン類作業環境測定（LG10）・精密機能検査（LG13）の行がない',
                     why='a1 が当時の委託契約の一覧から行を起こしたため、別契約・受検だけのものを落とした'))
    for it, ys in legal_rows(asof, asof, y1, skip):
        put(ws, r, 1, '法定', F(8, name=fn), AC, B_ALL); put(ws, r, 2, None, None, None, B_ALL)
        put(ws, r, 3, it[1], F(8, name=fn), Alignment(horizontal='left', vertical='center', shrink_to_fit=True), B_ALL)
        put(ws, r, 4, it[2].split('（')[0], F(7, name=fn), Alignment(horizontal='left', vertical='center', shrink_to_fit=True), B_ALL)
        put(ws, r, 5, {'odd': 2, 'even': 2, 'every': 1, 'tb': 4, '3y': 3}[it[4]], F(8, name=fn), AC, B_ALL); put(ws, r, 6, None, None, None, B_ALL)
        marks = {}
        for i, y in enumerate(years):
            v = int(round(it[5] * it[6] / 10.0)) if y in ys else None
            if y in ys:
                marks[str(y)] = v if v else '○'
            put(ws, r, yc0 + i, v if v else ('○' if y in ys else None), F(8, name=fn), AR, Border(left=hair, right=hair, top=hair, bottom=hair), fmt='#,##0')
        put(ws, r, cm, sum(x for x in marks.values() if isinstance(x, int)) or None, F(8, name=fn), AR, B_ALL)
        put(ws, r, cm + 1, it[8], F(7, name=fn), Alignment(horizontal='left', vertical='center', shrink_to_fit=True), B_ALL)
        ans.rec(stem, 'item', dict(path='法定点検・測定/%s' % it[1], cycle={'odd': 2, 'even': 2, 'every': 1, 'tb': 4, '3y': 3}[it[4]], unit='万円', marks=marks, memo=it[8]),
                sheet=ws.title, r=r, works=['%s-%s' % (it[0], y) for y in ys])
        r += 1
    page(ws, landscape=True, paper=8, area='A1:%s%d' % (CL(cm + 1), r), title_rows='4:4')
    ws.freeze_panes = 'G5'
    wb.save(ans.FILES + '/' + name)

# ------------------------------------------------------------------ a2 様式（H25 版・R1 改訂版）
def build_a2(asof):
    P = H.PLAN13 if asof == 2013 else H.PLAN19
    HL = H.H13 if asof == 2013 else H.H19
    era = 'H'
    if asof == 2013:
        years = list(range(2013, 2035)); act_years = []
        name = KAKU + 'あさぎ市清掃工場_中長期整備計画_長寿命化総合計画_平成25年度策定_a2様式.xlsx'
        made = '平成25年6月'
    else:
        years = list(range(2013, 2035)); act_years = list(range(2013, 2019))
        name = KAKU + 'あさぎ市清掃工場_中長期整備計画_長寿命化総合計画_令和元年度改訂_a2様式.xlsx'
        made = '令和元年10月'
    stem = ans.start(name, plan=True)
    make_ledger.build_2008(nosave=True)
    wb = Workbook(); ws0 = wb.active; ws0.title = '表紙'
    fn = 'ＭＳ Ｐゴシック'
    for i, t in enumerate([KAKU + 'あさぎ市清掃工場', '長寿命化総合計画（施設保全計画）' + ('' if asof == 2013 else '　令和元年度改訂'), '別表1 機器別整備計画', '', made, 'あさぎ市 環境部 清掃工場（施設係 a2）']):
        merge(ws0, 4 + i * 2, 2, 4 + i * 2, 10, t, F(20 if i < 2 else 12, i < 2, name=fn), AC, border=None)
    merge(ws0, 17, 2, 20, 10, ('計画期間：平成25年度〜平成46年度（延命化目標年度 平成46年度）\n基幹的設備改良工事：平成29〜31年度（継続費）\n' if asof == 2013 else
                               '基幹的設備改良工事（H29〜R1）の完了に伴い改訂。H25〜H30 は実績額（灰色）。計画期間 令和元年度〜令和16年度。\n') +
          '金額は概算（千円、消費税込み）。定期整備業務委託の定額部分は含まない。', F(10, name=fn), AL, border=None)
    page(ws0, landscape=True, paper=9, fith=1, area='A1:K21')
    ws = wb.create_sheet('別表1')
    yc0 = 12
    for k, v in {1: 4, 2: 5, 3: 4, 4: 28, 5: 8, 6: 6, 7: 5, 8: 4.5, 9: 5, 10: 4.5, 11: 5}.items():
        ws.column_dimensions[CL(k)].width = v
    for i in range(len(years)):
        ws.column_dimensions[CL(yc0 + i)].width = 7.2
    ctot = yc0 + len(years); crem = ctot + 1
    ws.column_dimensions[CL(ctot)].width = 9; ws.column_dimensions[CL(crem)].width = 30
    merge(ws, 1, 1, 1, 20, KAKU + '別表1　機器別整備計画（あさぎ市清掃工場）' + ('' if asof == 2013 else '　令和元年度改訂'), F(14, True, name=fn), AL, border=None)
    merge(ws, 2, 1, 2, 24, '単位：千円（概算・税込）　%s策定　健全度 4：健全 3：軽微 2：要注意 1：要対策' % made + ('' if asof == 2013 else '　灰色の列は実績額'), F(9, name=fn), AL, border=None)
    hr = 4
    hd = {1: '設備', 2: '号炉', 3: 'No', 4: '機　器　名', 5: '管理No', 6: '設置\n年度', 7: '経過\n年数', 8: '重要\n度', 9: '保全\n方式', 10: '健全\n度', 11: '目標\n耐用'}
    for c, hh in hd.items():
        merge(ws, hr, c, hr + 1, c, hh, F(8, True, name=fn), AC, fl=fill('D9D9D9'))
    merge(ws, hr, yc0, hr, yc0 + len(years) - 1, '年　度　別　整　備　計　画　（千円）', F(9, True, name=fn), AC, fl=fill('D9D9D9'))
    for i, y in enumerate(years):
        put(ws, hr + 1, yc0 + i, wy(y) if asof == 2013 else wy(y, 'R'), F(8, True, name=fn), AC, B_ALL, fill('BFBFBF' if y in act_years else ('FFD966' if y in (2017, 2018, 2019) else 'D9D9D9')))
    merge(ws, hr, ctot, hr + 1, ctot, '計', F(8, True, name=fn), AC, fl=fill('D9D9D9'))
    merge(ws, hr, crem, hr + 1, crem, '主な整備内容・備考', F(8, True, name=fn), AC, fl=fill('D9D9D9'))
    r = hr + 2
    subs = []
    grand = {y: 0 for y in years}
    for s in SYS_ORDER:
        sr0 = r
        secsum = {y: 0 for y in years}
        for f in (1, 2, 0):
            grp = [d for d in EQ if d['sys'] == s and d['f'] == f]
            if not grp:
                continue
            fr0 = r
            for j, d in enumerate(grp, 1):
                new = d['inst'] > asof and d['inst'] != 1996
                put(ws, r, 3, j, F(8, name=fn), AC, B_ALL)
                nm = d['name'] + ('（新設予定）' if new else '') + ('　' if (s == 'E' and f in (1, 2) and not new) else '')
                put(ws, r, 4, nm, F(8, name=fn, color='C00000' if new else '000000'), AL, B_ALL)
                put(ws, r, 5, d.get('old_code', '（H24新設）' if d['inst'] == 2012 else ('新設' if new else '')), F(8, name=fn), AC, B_ALL)
                put(ws, r, 6, '—' if new else ('H8' if d['inst'] == 1996 else 'H%d' % (d['inst'] - 1988)), F(8, name=fn), AC, B_ALL)
                put(ws, r, 7, None if new else asof - d['inst'], F(8, name=fn), AC, B_ALL)
                put(ws, r, 8, d['imp'], F(8, name=fn), AC, B_ALL); put(ws, r, 9, d['method'], F(8, name=fn), AC, B_ALL)
                hv = None if new else (HL[d['code']] if d['method'] != 'BM' else '—')
                put(ws, r, 10, hv, F(8, name=fn, bold=hv in (1, 2)), AC, B_ALL, fill('FFF2CC') if hv == 2 else (fill('F8CBAD') if hv == 1 else None))
                put(ws, r, 11, LIFE.get(d['kind'], 20), F(8, name=fn), AC, B_ALL)
                marks = {}
                tot = 0
                for i, y in enumerate(years):
                    if y in act_years:
                        v = [(k, t, c) for (k, t, c, ct) in d['act'].get(y, []) if ct in ('T', 'H', 'K')]
                    else:
                        v = P[d['code']].get(y, []) if y >= asof else []
                    val = sum(e[2] for e in v) if v else None
                    c = put(ws, r, yc0 + i, val or None, F(8, name=fn), AR, Border(left=hair, right=hair, top=hair, bottom=hair), fmt='#,##0;-#,##0;;')
                    if v:
                        c.fill = fill('D9D9D9') if y in act_years else fill(KFILL[main_kind(v)])
                        if val:
                            marks[str(y)] = val; tot += val; secsum[y] += val
                put(ws, r, ctot, tot or None, F(8, name=fn), AR, B_ALL, fmt='#,##0;-#,##0;;')
                txt = '、'.join(sorted(set(e[1] for y, v in P[d['code']].items() if y >= asof for e in v), key=lambda t: ('基幹' not in t, t)))
                if d['method'] == 'BM' and not P[d['code']]:
                    txt = '事後保全（故障時修理）'
                put(ws, r, crem, txt, F(7, name=fn), Alignment(horizontal='left', vertical='center', shrink_to_fit=True), B_ALL)
                cm = {}
                if asof == 2013 and d['tag'] == 'F-103':
                    ws.cell(row=r, column=yc0 + 4).comment = Comment('インバータ化はCO2削減対策として交付金対象（3%以上削減の見込み）', 'a2')
                    cm['%s%d' % (CL(yc0 + 4), r)] = 'インバータ化はCO2削減対策として交付金対象（3%以上削減の見込み）'
                if asof == 2019 and d['tag'] == 'F-201':
                    t_ = '基幹改良の対象外。振動監視（CBM）とし R2 の軸受交換は状態を見て判断'
                    ws.cell(row=r, column=yc0 + 7).comment = Comment(t_, 'a2'); cm['%s%d' % (CL(yc0 + 7), r)] = t_
                ans.rec(stem, 'item', dict(path=path(d), code=d.get('old_code'), work=txt or None, cost=tot or None, unit='千円', marks=marks, comments=cm or None,
                                           memo='新設予定' if new else None), sheet='別表1', r=r, code=d['code'])
                r += 1
            merge(ws, fr0, 2, r - 1, 2, d['fname'], F(8, name=fn), ACV)
        merge(ws, sr0, 1, r - 1, 1, SYS[s], F(9, True, name=fn), ACV)
        merge(ws, r, 1, r, 11, '%s　小計' % SYS[s], F(8, True, name=fn), AR, fl=fill('F2F2F2'))
        shown = dict(secsum)
        if asof == 2019 and s == 'J':
            shown[2020] = secsum[2020] + 35000   # R2 列：基幹改良前の値のまま（小計の直し忘れ）
            GAPS.append(dict(file=name, kind='費用の小計の食い違い', where='別表1 余熱利用設備 小計の R2 列', what='小計が機器行の合計より 35,000 千円多い',
                             why='H25 版から写した時の値（タービン開放点検をR2に置いていた）を直し忘れた'))
        for i, y in enumerate(years):
            put(ws, r, yc0 + i, shown[y] or None, F(8, True, name=fn), AR, box(t=thin, b=thin, l=hair, r=hair), fill('F2F2F2'), fmt='#,##0;-#,##0;;')
            grand[y] += shown[y]
        put(ws, r, ctot, sum(shown.values()), F(8, True, name=fn), AR, B_ALL, fill('F2F2F2'), fmt='#,##0')
        ans.rec(stem, 'subtotal', dict(path=SYS[s], cost=sum(shown.values()), unit='千円', marks={str(y): v for y, v in shown.items() if v}), sheet='別表1', r=r)
        r += 1
        if s == 'L':
            r += 2
            merge(ws, r - 1, 4, r - 1, 20, '※計装制御設備は H24 年度 DCS 更新済み（旧CRT・ベンチボード撤去）', F(8, name=fn, color='595959'), AL, border=None)
    merge(ws, r, 1, r, 11, '合　　計', F(9, True, name=fn), AC, fl=fill('D9D9D9'))
    for i, y in enumerate(years):
        put(ws, r, yc0 + i, grand[y] or None, F(8, True, name=fn), AR, box(t=dbl, b=med, l=hair, r=hair), fill('D9D9D9'), fmt='#,##0')
    put(ws, r, ctot, sum(grand.values()), F(8, True, name=fn), AR, B_ALL, fill('D9D9D9'), fmt='#,##0')
    ans.rec(stem, 'total', dict(path='合計', cost=sum(grand.values()), unit='千円', marks={str(y): v for y, v in grand.items() if v}), sheet='別表1', r=r)
    r += 2
    merge(ws, r, 1, r, crem, '法令に基づく点検・検査・測定（委託・受検。上の合計には含めない）', F(9, True, name=fn), AL, fl=fill('FFF2CC'))
    ans.rec(stem, 'parent', {'path': '法令に基づく点検・検査・測定'}, sheet='別表1', r=r)
    r += 1
    for j, (it, ys) in enumerate(legal_rows(asof, years[0], years[-1]), 1):
        if asof == 2013 and it[0] == 'LG02':
            ys = [y for y in ys if y != 2024]
            GAPS.append(dict(file=name, kind='法令点検の抜け', where='別表1 法令の欄 2号炉ボイラー定期事業者検査の H36 列', what='H36（R6）の印がない（2 年周期の 1 回分の抜け）',
                             why='行をコピーして年を入れ直した時に 1 か所消した。実際には R6 に検査を行っている'))
        put(ws, r, 3, j, F(8, name=fn), AC, B_ALL)
        merge(ws, r, 4, r, 9, it[1], F(8, name=fn), Alignment(horizontal='left', vertical='center', shrink_to_fit=True))
        merge(ws, r, 10, r, 11, it[2].split('（')[0], F(7, name=fn), Alignment(horizontal='center', vertical='center', shrink_to_fit=True))
        marks = {}
        for i, y in enumerate(years):
            v = (it[5] * it[6]) if y in ys else None
            if y in ys:
                marks[str(y)] = v
            c = put(ws, r, yc0 + i, v, F(8, name=fn), AR, Border(left=hair, right=hair, top=hair, bottom=hair), fmt='#,##0;-#,##0;"(機器側)";')
            if y in ys:
                c.fill = fill('E2EFDA')
        put(ws, r, ctot, sum(x for x in marks.values() if isinstance(x, int)) or None, F(8, name=fn), AR, B_ALL, fmt='#,##0')
        put(ws, r, crem, it[8], F(7, name=fn), Alignment(horizontal='left', vertical='center', shrink_to_fit=True), B_ALL)
        ans.rec(stem, 'item', dict(path='法令に基づく点検・検査・測定/%s' % it[1], unit='千円', marks=marks, memo=it[8]), sheet='別表1', r=r,
                works=['%s-%d' % (it[0], y) for y in ys])
        r += 1
    page(ws, landscape=True, paper=8, area='A1:%s%d' % (CL(crem), r), title_rows='%d:%d' % (hr, hr + 1))
    ws.freeze_panes = '%s%d' % (CL(yc0), hr + 2)
    wb.save(ans.FILES + '/' + name)

# ------------------------------------------------------------------ a3 星取表（R3・R5・R8）
def plan_in_effect(y):
    return H.PLAN13 if y < 2019 else (H.PLAN19 if y < 2021 else (H.PLAN21 if y < 2023 else (H.PLAN23 if y < 2026 else H.PLAN26)))

def build_a3(asof):
    VER = {2021: H.PLAN21, 2023: H.PLAN23, 2026: H.PLAN26}[asof]
    HL = {2021: H.H21, 2023: H.H23v, 2026: H.H26}[asof]
    lab = {2021: '令和3年度版', 2023: '令和5年度版', 2026: '令和8年度版（R8.10 上期実績追記）'}[asof]
    name = KAKU + 'あさぎ市清掃工場_中長期整備計画_年度別星取表_%s_a3様式.xlsx' % lab.split('（')[0]
    stem = ans.start(name, plan=True)
    wb = Workbook(); ws = wb.active; ws.title = '星取表（R%d）' % (asof - 2018)
    fn = 'Meiryo UI'
    years = list(range(2017, 2035)); yc0 = 11
    for k, v in {1: 4, 2: 5, 3: 4, 4: 28, 5: 8, 6: 4.5, 7: 5, 8: 5, 9: 5, 10: 4.5}.items():
        ws.column_dimensions[CL(k)].width = v
    for i in range(len(years)):
        ws.column_dimensions[CL(yc0 + i)].width = 4.2
    crem = yc0 + len(years); ws.column_dimensions[CL(crem)].width = 34
    merge(ws, 1, 1, 1, 20, KAKU + '別表1　機器別整備計画（あさぎ市清掃工場）', F(10, name=fn, color='808080', strike=True), AL, border=None)
    merge(ws, 2, 1, 2, 30, KAKU + 'あさぎ市清掃工場　年度別整備計画（星取表）　' + lab, F(14, True, name=fn), AL, border=None)
    merge(ws, 3, 1, 3, 30, '○点検（○2＝年2回）　◇計測　◎整備（分解整備・部品交換）　△補修　●更新　／　上段：計画（R%d以前は当初計画）　下段：実績（灰色）%s' % (asof - 2019,
          '　▷ 施工中（R8 2号炉定修 R8.10.5〜）' if asof == 2026 else ''), F(9, name=fn), AL, border=None)
    hr = 5
    for c, hh in {1: '設備', 2: '号炉', 3: 'No', 4: '機　器　名', 5: '機器コード', 6: '重要\n度', 7: '保全\n方式', 8: '健全度\n(R%d)' % (asof - 2018), 9: '目標耐用\n(旧)', 10: '区分'}.items():
        merge(ws, hr, c, hr + 1, c, hh, F(8, True, name=fn), AC, fl=fill('D9E1F2') if c != 9 else fill('EDEDED'))
    merge(ws, hr, yc0, hr, yc0 + 2, '平成', F(8, True, name=fn), AC, fl=fill('D9E1F2'))
    merge(ws, hr, yc0 + 3, hr, yc0 + len(years) - 1, '令和', F(8, True, name=fn), AC, fl=fill('D9E1F2'))
    for i, y in enumerate(years):
        put(ws, hr + 1, yc0 + i, ('%d' % (y - 1988)) if y < 2019 else ('元' if y == 2019 else '%d' % (y - 2018)), F(8, True, name=fn), AC, B_ALL, fill('FFD966' if y == asof else 'D9E1F2'))
    merge(ws, hr, crem, hr + 1, crem, '備考', F(8, True, name=fn), AC, fl=fill('D9E1F2'))
    r = hr + 2
    grayf = fill('E7E6E6')
    notes26 = {'F-201': 'R6下期から振動上昇（R7.5 6.3mm/s）→R7 補修工事で軸受・軸受箱交換。以後CBM', 'GC-101': 'R7 下部コーン腐食（最小3.1mm）当て板。R9 部分更新',
               'BF-101': 'R7 ろ布全数交換（予備12本 市倉庫）。次回 R13（廃止まで使用）', 'BF-201': 'R8 2号炉定修でろ布全数交換（施工中）', 'TB-01': '開放点検 4年ごと（R1基幹で分解整備）',
               'ST-102': 'R7 最大摩耗4.2mm。R10 部分取替（約20%）', 'SH-201': 'H30更新。減肉 年0.04mm、廃止まで交換なし', 'OPS-01': 'R9 OS・ハード保守終了 → HMI更新',
               'CR-01': '常用機。ワイヤ 2〜3年', 'CR-02': '予備機。ワイヤ 5年程度', 'HV-01': '精密点検 3年ごと（全停時）'}
    for s in SYS_ORDER:
        sr0 = r
        for f in (1, 2, 0):
            grp = [d for d in EQ if d['sys'] == s and d['f'] == f]
            if not grp:
                continue
            fr0 = r
            for j, d in enumerate(grp, 1):
                merge(ws, r, 3, r + 1, 3, j, F(8, name=fn), AC); merge(ws, r, 4, r + 1, 4, d['name'], F(8, name=fn), AL)
                merge(ws, r, 5, r + 1, 5, d['code'], F(8, name=fn), AC); merge(ws, r, 6, r + 1, 6, d['imp'], F(8, name=fn), AC)
                merge(ws, r, 7, r + 1, 7, d['method'], F(8, name=fn), AC)
                hv = HL[d['code']] if d['method'] != 'BM' else '—'
                merge(ws, r, 8, r + 1, 8, hv, F(8, name=fn, bold=hv in (1, 2)), AC, fl=fill('FFF2CC') if hv == 2 else (fill('F8CBAD') if hv == 1 else None))
                merge(ws, r, 9, r + 1, 9, LIFE.get(d['kind'], 20) if d['inst'] <= 2013 or d['inst'] == 1996 else None, F(8, name=fn, color='A6A6A6'), AC)
                put(ws, r, 10, '計画', F(7, name=fn), AC, B_ALL); put(ws, r + 1, 10, '実績', F(7, name=fn), AC, B_ALL, grayf)
                pm, am = {}, {}
                for i, y in enumerate(years):
                    src = plan_in_effect(y) if y < asof else VER
                    pe = src[d['code']].get(y, [])
                    ae = [(k, t, c) for (k, t, c, ct) in d['act'].get(y, []) if ct != 'R'] if (y < asof or (y == 2026 and asof == 2026)) else []
                    c1 = put(ws, r, yc0 + i, SYMB[main_kind(pe)] if pe else None, F(9, name=fn), AC, Border(left=hair, right=hair, top=hair, bottom=dot))
                    c2 = put(ws, r + 1, yc0 + i, SYMB[main_kind(ae)] if ae else None, F(9, True, name=fn), AC, Border(left=hair, right=hair, top=dot, bottom=thin), grayf if y < asof or (asof == 2026 and y == 2026) else None)
                    if pe: pm[str(y)] = SYMB[main_kind(pe)]
                    if ae: am[str(y)] = SYMB[main_kind(ae)]
                    if asof == 2026 and y == 2026 and d.get('pending2026'):
                        c2.value = '▷'; c2.font = F(8, name=fn, color='0070C0'); am[str(y)] = '▷'
                    if y == asof:
                        c1.border = Border(left=med, right=med, top=hair, bottom=dot); c2.border = Border(left=med, right=med, top=dot, bottom=thin)
                rem = notes26.get(d['tag'], '') if asof == 2026 else ''
                if d['method'] == 'BM' and not any(VER[d['code']].values()):
                    rem = rem or '事後保全'
                merge(ws, r, crem, r + 1, crem, rem, F(7, name=fn), AL)
                ws.row_dimensions[r].height = 12.5; ws.row_dimensions[r + 1].height = 12.5
                cm = {}
                if d['tag'] == 'F-201' and asof == 2026:
                    a1 = '%s%d' % (CL(yc0 + years.index(2025)), r + 1); a2_ = '%s%d' % (CL(yc0 + years.index(2020)), r)
                    ws[a1].comment = Comment('R7.10 補修工事（請負）。定修とは別契約', 'a3'); ws[a2_].comment = Comment('H25版計画の軸受交換。R2 は振動2.0mm/sで良好のため見送り', 'a3')
                    cm = {a1: 'R7.10 補修工事（請負）。定修とは別契約', a2_: 'H25版計画の軸受交換。R2 は振動2.0mm/sで良好のため見送り'}
                wk = '、'.join(sorted(set(e[1] for y, v in VER[d['code']].items() if y >= asof for e in v)))
                ans.rec(stem, 'item', dict(path=path(d), code=d['code'], 区分='計画', marks=pm, memo=rem or None, comments={k: v for k, v in cm.items() if int(''.join(ch for ch in k if ch.isdigit())) == r} or None),
                        sheet=ws.title, r=r, code=d['code'])
                ans.rec(stem, 'item', dict(path=path(d), code=d['code'], 区分='実績', marks=am, comments={k: v for k, v in cm.items() if int(''.join(ch for ch in k if ch.isdigit())) == r + 1} or None),
                        sheet=ws.title, r=r + 1, code=d['code'], works=[w for y in am for w in B.TR.work_ids(int(y), d['tag'])])
                r += 2
            merge(ws, fr0, 2, r - 1, 2, d['fname'], F(8, name=fn), ACV)
        merge(ws, sr0, 1, r - 1, 1, SYS[s], F(9, True, name=fn), ACV)
    lr0 = r
    for j, (it, ys) in enumerate(legal_rows(asof, 2017, 2034), 1):
        merge(ws, r, 3, r + 1, 3, j, F(8, name=fn), AC)
        merge(ws, r, 4, r + 1, 5, it[1], F(7, name=fn), Alignment(horizontal='left', vertical='center', wrap_text=True))
        merge(ws, r, 6, r + 1, 9, it[2], F(7, name=fn), Alignment(horizontal='center', vertical='center', wrap_text=True))
        put(ws, r, 10, '計画', F(7, name=fn), AC, B_ALL); put(ws, r + 1, 10, '実績', F(7, name=fn), AC, B_ALL, grayf)
        pm, am = {}, {}
        for i, y in enumerate(years):
            mark = ('○' if it[5] == 1 else '○%d' % it[5]) if y in ys else None
            put(ws, r, yc0 + i, mark, F(8, name=fn), AC, Border(left=hair, right=hair, top=hair, bottom=dot))
            a = mark if (y < asof and y in ys) else ('▷' if (asof == 2026 and y == 2026 and y in ys) else None)
            put(ws, r + 1, yc0 + i, a, F(8, True, name=fn), AC, Border(left=hair, right=hair, top=dot, bottom=thin), grayf if y <= asof else None)
            if mark: pm[str(y)] = mark
            if a: am[str(y)] = a
        merge(ws, r, crem, r + 1, crem, it[8], F(7, name=fn), AL)
        ans.rec(stem, 'item', dict(path='法令に基づく点検・検査・測定/%s' % it[1], 区分='計画', marks=pm, memo=it[8]), sheet=ws.title, r=r)
        ans.rec(stem, 'item', dict(path='法令に基づく点検・検査・測定/%s' % it[1], 区分='実績', marks=am), sheet=ws.title, r=r + 1, works=['%s-%s' % (it[0], y) for y in am])
        r += 2
    merge(ws, lr0, 1, r - 1, 2, '法令に基づく点検・検査・測定', F(8, True, name=fn), ACV, fl=fill('FFF2CC'))
    page(ws, landscape=True, paper=8, area='A1:%s%d' % (CL(crem), r - 1), title_rows='%d:%d' % (hr, hr + 1))
    ws.freeze_panes = '%s%d' % (CL(yc0), hr + 2)
    # 年度別事業費
    w2 = wb.create_sheet('年度別事業費')
    merge(w2, 1, 1, 1, 10, KAKU + '年度別事業費（千円・税込）　計画：星取表の計画値の積上げ（過去は当初計画）　実績：契約金額', F(12, True, name=fn), AL, border=None)
    hd2 = ['年度', '計画（機器別積上げ）', '実績 定期整備業務委託', '実績 補修工事（請負）', '実績 基幹的設備改良', '実績 計', '備考']
    for i, x in enumerate(hd2):
        put(w2, 3, i + 1, x, F(9, True, name=fn), AC, B_ALL, fill('D9E1F2'))
    widths(w2, {'A': 8, 'B': 16, 'C': 16, 'D': 16, 'E': 16, 'F': 14, 'G': 40})
    for i, y in enumerate(years):
        rr = 4 + i
        put(w2, rr, 1, wareki(y), F(9, name=fn), AC, B_ALL)
        src = plan_in_effect(y) if y < asof else VER
        pl = sum(e[2] for d in EQ for e in src[d['code']].get(y, []))
        put(w2, rr, 2, pl, F(9, name=fn), AR, B_ALL, fmt='#,##0')
        f = dict(year=y, plan=pl)
        if y < asof:
            a = H.actual_amounts(y)
            tot = round(a['teishu'] + a['hoshu'] + a['kikan'])
            for c, v in ((3, a['teishu']), (4, a['hoshu']), (5, a['kikan'] or None), (6, tot)):
                put(w2, rr, c, v, F(9, name=fn), AR, B_ALL, fmt='#,##0')
            f.update(teishu=a['teishu'], hoshu=a['hoshu'], kikan=a['kikan'] or None, actual=tot)
        else:
            for c in (3, 4, 5, 6):
                put(w2, rr, c, None, None, None, B_ALL)
        put(w2, rr, 7, None, None, None, B_ALL)
        ans.rec(stem, 'item', f, sheet='年度別事業費', r=rr)
    re_ = 4 + len(years)
    merge(w2, re_ + 1, 1, re_ + 1, 7, '※計画は機器別の概算額で、定期整備委託の定額部分を含まない。実績は契約金額のため両者は直接比較できない。物品・運転委託・薬剤は含まない。', F(8, name=fn), AL, border=None)
    ch = BarChart(); ch.type = 'col'; ch.title = '年度別事業費（千円）'; ch.height = 8; ch.width = 22
    ch.add_data(Reference(w2, min_col=2, min_row=3, max_row=re_ - 1), titles_from_data=True)
    ch.add_data(Reference(w2, min_col=6, min_row=3, max_row=re_ - 1), titles_from_data=True)
    ch.set_categories(Reference(w2, min_col=1, min_row=4, max_row=re_ - 1))
    w2.add_chart(ch, 'A%d' % (re_ + 3))
    page(w2, landscape=True, paper=9, fith=1, area='A1:G%d' % (re_ + 20))
    wb.save(ans.FILES + '/' + name)
    if asof == 2023:
        GAPS.append(dict(file=name, kind='周期の超え', where='星取表 蒸気タービン（CJ-013）計画行', what='R5 の開放点検の次を R10 に置いた（5 年。定期事業者検査の 4 年を超える）',
                         why='R9 は減温塔下部コーンの補修と重なるため送ったが、法定周期の確認が漏れた。R8 版で R9 に直した'))
    if asof == 2021:
        GAPS.append(dict(file=name, kind='周期の超え', where='星取表 2号炉バグフィルタ本体（2D-020）計画行', what='R2 交換の次のろ布全数交換を R9（7 年後）に置いた',
                         why='予算をならすため 1 年送った。R5 の抜取検査（強度保持率 74%）を受け、R8 版で R8 に戻した'))
