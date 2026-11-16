# -*- coding: utf-8 -*-
"""工程表（全年度）。定期整備：n1 様式 R3〜R5、n2 様式 R6〜R7。基幹改良：n3 詳細工程 R4・R5（DCS切替含む）。"""
import os, datetime as dt
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter as L
from common import *
import works as W
import schedules as S
import rec

D = dt.date
is_off = W.is_off


def stamp_box(ws, r, c, labels, names):
    for i, (t, n) in enumerate(zip(labels, names)):
        put(ws, r, c + i, t, f(8), AC)
        put(ws, r + 1, c + i, n, f(9, color='C00000'), AC)
    box(ws, r, c, r + 1, c + len(labels) - 1)
    ws.row_dimensions[r + 1].height = 26


def fmt_spans(sp):
    return [[s.isoformat(), e.isoformat()] for s, e in sp]


def stop_bands(year, ev, days, info):
    """停止の帯（区分, [(s,e,文字,色)]）"""
    st = W.STOPS[year]
    bands = []
    if ev.startswith('mid'):
        k = int(ev[-1]); o = 3 - k
        down, w0, w1, up = st[ev]
        bands.append((f'{k}号炉', [(days[0], down, '運転', 'E2EFDA'), (down + dt.timedelta(1), up - dt.timedelta(1), '停止', 'D9D9D9'), (up, up, '立上', 'FFE699')]))
        other = '運転（単独運転・時間延長）'
        if (year == 2022 and o == 1) or (year == 2023 and o == 2):
            other = '運転'
        bands.append((f'{o}号炉', [(days[0], days[-1], other, 'E2EFDA')]))
    else:
        down, w0, w1, up = st['all']
        for k in (1, 2):
            if year == 2022 and k == 1:
                bands.append(('1号炉', [(days[0], days[-1], '基幹的設備改良工事のため停止中（別契約）', 'BFBFBF')]))
            elif year == 2023 and k == 2:
                bands.append(('2号炉', [(days[0], days[-1], '基幹的設備改良工事のため停止中（別契約）', 'BFBFBF')]))
            else:
                upk = up if k == 2 or year in (2021, 2022, 2023) else up + dt.timedelta(1)
                bands.append((f'{k}号炉', [(days[0], down, '運転', 'E2EFDA'), (down + dt.timedelta(1), upk - dt.timedelta(1), '停止', 'D9D9D9'), (upk, upk, '立上', 'FFE699')]))
        bands.append(('全停電', [(st['blackout'], st['blackout'], '停電', 'FFC000')]))
    return bands


# ------------------------------------------------------------------ n1（R3〜R5）
def n1_sheet(stem, ws, title, days, items, bands, revisions, sv='n1'):
    c0 = 5
    for c, w in zip('ABCD', (6, 3.5, 36, 12)):
        ws.column_dimensions[c].width = w
    for i in range(len(days)):
        ws.column_dimensions[L(c0 + i)].width = 3.6
    last = c0 + len(days) - 1
    merge(ws, 1, 1, 1, last - 4, '【架空】' + title, f(13, True), AL)
    merge(ws, 2, 1, 2, last - 4, f'施設：{FAC}　　施工：{MAKER}', f(9), AL)
    stamp_box(ws, 1, last - 2, ('組合', '照査', '作成'), ('', '', sv))
    r = 4
    for c, t in ((1, '区分'), (2, 'No'), (3, '作　業　項　目'), (4, '担当')):
        merge(ws, r, c, r + 1, c, t, f(9, True), AC, fill('DDEBF7'))
    for i, d in enumerate(days):
        put(ws, r, c0 + i, f'{d.month}/{d.day}' if (i == 0 or d.day == 1) else d.day, f(7), AC, fill('DDEBF7'))
        put(ws, r + 1, c0 + i, wd(d), f(7, color='C00000' if is_off(d) else '000000'), AC, fill('FCE4D6') if is_off(d) else fill('DDEBF7'))
    box(ws, r, 1, r + 1, last)
    r += 2
    for lab, spans in bands:
        merge(ws, r, 1, r, 4, lab, f(8), AL, fill('F2F2F2'))
        sp = []
        for i, d in enumerate(days):
            for (s, e, txt, col) in spans:
                if s <= d <= e:
                    put(ws, r, c0 + i, txt if d == s or (i == 0 and s < d) else None, f(7, True), ALN, fill(col))
        rec.add(stem, ws.title, r, 'item', dict(区分=lab, 状態=[[max(s, days[0]).isoformat(), min(e, days[-1]).isoformat(), t] for s, e, t, c in spans]))
        r += 1
    box(ws, r - len(bands), 1, r - 1, last)
    st = r
    prev = None
    pers = {}
    for no, it in enumerate(items, 1):
        put(ws, r, 1, it['kubun'] if it['kubun'] != prev else None, f(9), AC)
        put(ws, r, 2, no, f(8), AC)
        cc = put(ws, r, 3, it['name'], f(9, color='C00000' if it.get('rev') else '000000'), AL)
        if it.get('rev'):
            cc.comment = Comment(f"改{it['rev']}で追加", 'n1')
        put(ws, r, 4, it['team'], f(8), AC)
        for i, d in enumerate(days):
            if is_off(d):
                put(ws, r, c0 + i, None, fl=fill('FCE4D6'))
            for s, e in it['plan']:
                if s <= d <= e:
                    put(ws, r, c0 + i, None, fl=fill('FF7C80' if it.get('rev') else '5B9BD5'))
                    if not it.get('rev') and not is_off(d) or d == it.get('blackout_day'):
                        pers.setdefault(d, set()).add(it['team'])
        rec.add(stem, ws.title, r, 'item', dict(区分=it['kubun'] if it['kubun'] != prev else None, 区分の値=it['kubun'], No=no, 作業=it['name'], 担当=it['team'],
                                                plan=fmt_spans(it['plan']), 追加=bool(it.get('rev'))), works=[it.get('wid')])
        prev = it['kubun']
        r += 1
    box(ws, st, 1, r - 1, last)
    merge(ws, r, 1, r, 4, '予定人員（人）', f(9, True), AC, fill('F2F2F2'))
    TEAMN = {'機械1班': 5, '機械2班': 4, '機械3班（応援）': 5, '電計班': 3, '外注(足場)': 3, '外注(クレーン)': 3, '外注(計器)': 2, '全班': 7}
    pv = {}
    for i, d in enumerate(days):
        if d in pers:
            v = 1 + sum(TEAMN.get(t, 2) for t in pers[d])
            put(ws, r, c0 + i, v, f(7), AC, fill('F2F2F2'))
            pv[d.isoformat()] = v
        else:
            put(ws, r, c0 + i, None, fl=fill('F2F2F2'))
    rec.add(stem, ws.title, r, 'total', dict(予定人員=pv, unit='人'))
    box(ws, r, 1, r, last)
    r += 2
    put(ws, r, 1, '改訂', f(8, True), AL)
    for t in revisions:
        merge(ws, r, 2, r, last, t, f(8, color='C00000' if '改1' in t else '000000'), ALN)
        r += 1
    put(ws, r, 2, '凡例　■ 作業予定　■ 追加作業（赤）　網掛け 日曜・祝日', f(8), ALN)
    setup_print(ws, True, 9, 1, 1, f'A1:{L(last)}{r}')


def ev_days(year, ev):
    down, w0, w1, up = W.STOPS[year][ev]
    return list(daterange(down - dt.timedelta(1), up + dt.timedelta(1)))


def items_of(info):
    res = []
    for it in info['items']:
        if info.get('hand'):
            res.append(dict(kubun=it['kubun'], name=it['name'], team=it['team'], plan=it['plan'], act=it['act'], rev=it.get('rev', 0),
                            wid=it['wid'], persons=it.get('persons')))
        else:
            w = W.WBY[it['wid']]
            res.append(dict(kubun=it['kubun'], name=it['name'], team=it['team'], plan=w['plan'], act=w['act'], rev=0, wid=it['wid'], persons=it['persons']))
    return res


def build_n1_year(year):
    sv = 'n1'
    fname = f'{FN}_工程表_令和{year-2018}年度定期整備業務委託_n1様式.xlsx'
    stem = rec.start(fname)
    wb = Workbook()
    props(wb, f'【架空】令和{year-2018}年度 やまぶき広域処理組合第二工場 定期整備業務委託 工程表', f'{MAKER} {sv}')
    first = True
    for ev in ('mid1', 'mid2', 'all'):
        info = W.TEIKI_EV[year].get(ev)
        if not info:
            continue
        ws = wb.active if first else wb.create_sheet()
        first = False
        ws.title = W.EVLABEL[ev]
        days = ev_days(year, ev) if ev != 'all' else list(daterange(W.STOPS[year]['all'][0] - dt.timedelta(1), W.STOPS[year]['all'][3] + dt.timedelta(1)))
        items = items_of(info)
        sub = '令和{}年{}月{}日'.format(year - 2018, *(W.STOPS[year][ev][0] - dt.timedelta(14)).timetuple()[1:3])
        revs = [f'改0　R{year-2018}.{(W.STOPS[year][ev][0]-dt.timedelta(14)).month}.{(W.STOPS[year][ev][0]-dt.timedelta(14)).day}　提出']
        if year == 2022 and ev == 'all':
            revs = ['改0　R4.9.30　提出', '改1　R4.10.20　No.8 灰押出装置 摺動板取替 追加（10/19 点検で摩耗大、組合了承）']
        n1_sheet(stem, ws, f'令和{year-2018}年度 やまぶき広域処理組合第二工場 定期整備業務委託　{W.EVLABEL[ev]} 工程表', days, items,
                 stop_bands(year, ev, days, info), revs)
    wb.save(os.path.join(OUT, fname))
    return fname


# ------------------------------------------------------------------ n2（R6〜R7）
def n2_event_sheet(stem, ws, year, ev, info, sv='n2'):
    if ev == 'all':
        days = list(daterange(W.STOPS[year]['all'][0] - dt.timedelta(1), W.STOPS[year]['all'][3] + dt.timedelta(2)))
    else:
        days = ev_days(year, ev)
    c0 = 7
    for c, w in zip('ABCDEF', (6, 3.5, 34, 12, 4.5, 4)):
        ws.column_dimensions[c].width = w
    for i in range(len(days)):
        ws.column_dimensions[L(c0 + i)].width = 3.6
    last = c0 + len(days) - 1
    merge(ws, 1, 1, 1, max(8, last - 4), f'【架空】令和{year-2018}年度 やまぶき広域処理組合第二工場 定期整備業務委託　{W.EVLABEL[ev]} 工程表（予定・実績）', f(13, True), AL)
    pd_ = W.STOPS[year][ev][0] - dt.timedelta(14)
    merge(ws, 2, 1, 2, max(8, last - 4), f'施設：{FAC}　　施工：{MAKER}　　予定 R{year-2018}.{pd_.month}.{pd_.day} 提出 ／ 実績 R{year-2018}.{(W.STOPS[year][ev][3]+dt.timedelta(5)).month}.{(W.STOPS[year][ev][3]+dt.timedelta(5)).day} 記入', f(9), AL)
    stamp_box(ws, 1, last - 2, ('組合', '照査', '作成'), ('', '', sv))
    prevy = year - 1
    merge(ws, 3, 1, 3, 12, f'改0　R{prevy-2018}.{pd_.month}.{pd_.day}　提出', f(7, color='A6A6A6', strike=True), AL)
    r = 4
    for c, t in ((1, '区分'), (2, 'No'), (3, '作　業　項　目'), (4, '担当'), (5, '人数'), (6, '予実')):
        merge(ws, r, c, r + 1, c, t, f(9 if c < 5 else 8, True), AC, fill('DDEBF7'))
    for i, d in enumerate(days):
        put(ws, r, c0 + i, f'{d.month}/{d.day}' if i == 0 or d.day == 1 else d.day, f(7), AC, fill('DDEBF7'))
        put(ws, r + 1, c0 + i, wd(d), f(7, color='C00000' if is_off(d) else '000000'), AC, fill('FCE4D6') if is_off(d) else fill('DDEBF7'))
    box(ws, r, 1, r + 1, last)
    r += 2
    bands = stop_bands(year, ev, days, info)
    for lab, spans in bands:
        merge(ws, r, 1, r, 6, lab, f(8), AL, fill('F2F2F2'))
        for i, d in enumerate(days):
            for (s, e, txt, col) in spans:
                if s <= d <= e:
                    put(ws, r, c0 + i, txt if d == s or (i == 0 and s < d) else None, f(7, True), ALN, fill(col))
        rec.add(stem, ws.title, r, 'item', dict(区分=lab, 状態=[[max(s, days[0]).isoformat(), min(e, days[-1]).isoformat(), t] for s, e, t, c in spans]))
        r += 1
    box(ws, r - len(bands), 1, r - 1, last)
    prev = None
    delays = []
    for no, it in enumerate(items_of(info), 1):
        merge(ws, r, 1, r + 1, 1, it['kubun'] if it['kubun'] != prev else None, f(9), AC)
        merge(ws, r, 2, r + 1, 2, no, f(8), AC)
        merge(ws, r, 3, r + 1, 3, it['name'], f(9), AL)
        merge(ws, r, 4, r + 1, 4, it['team'], f(8), AC)
        merge(ws, r, 5, r + 1, 5, it['persons'], f(8), AC)
        put(ws, r, 6, '予', f(7), AC); put(ws, r + 1, 6, '実', f(7), AC)
        delayed = it['act'] != it['plan']
        pend = max(e for s, e in it['plan'])
        for i, d in enumerate(days):
            for s, e in it['plan']:
                if s <= d <= e:
                    put(ws, r, c0 + i, None, fl=fill('9BC2E6'))
            for s, e in it['act']:
                if s <= d <= e:
                    put(ws, r + 1, c0 + i, None, fl=fill('C00000' if (delayed and d > pend) else '1F4E78'))
        if delayed:
            ws.cell(r + 1, 6).comment = Comment('実績が予定より延びた（週報・日報参照）', sv)
            delays.append(no)
        for rr in (r, r + 1):
            for c in range(1, last + 1):
                ws.cell(rr, c).border = Border(left=thin, right=thin, top=thin if rr == r else dot, bottom=dot if rr == r else thin)
            ws.row_dimensions[rr].height = 11
        rec.add(stem, ws.title, r, 'item', dict(区分=it['kubun'] if it['kubun'] != prev else None, 区分の値=it['kubun'], No=no, 作業=it['name'], 担当=it['team'],
                                                人数=it['persons'], plan=fmt_spans(it['plan']), actual=fmt_spans(it['act'])), works=[it['wid']])
        prev = it['kubun']
        r += 2
    r += 1
    put(ws, r, 3, '凡例　予定 ■（淡）　実績 ■（濃）　実績の遅れ ■（赤）', f(8), ALN)
    if year == 2025 and ev == 'all':
        r += 1
        put(ws, r, 3, '※ No.12 2号IDF：10/20 試運転 振動 4.8mm/s → 再バランス、10/24 完了（2日延長）。No.23 ポンプ(A)：1日延長。', f(8), ALN)
    setup_print(ws, True, 9, 1, 1, f'A1:{L(last)}{r}')


def n2_annual(stem, ws, year, sv='n2'):
    c0 = 4
    ws.column_dimensions['A'].width = 4; ws.column_dimensions['B'].width = 30; ws.column_dimensions['C'].width = 5
    months = [(year, m) for m in range(4, 13)] + [(year + 1, m) for m in (1, 2, 3)]
    for i in range(36):
        ws.column_dimensions[L(c0 + i)].width = 2.6
    last = c0 + 35
    merge(ws, 1, 1, 1, last - 6, f'【架空】令和{year-2018}年度 やまぶき広域処理組合第二工場 定期整備業務委託　年間工程表', f(13, True), AL)
    lastday = W.STOPS[year]['all'][3] + dt.timedelta(5)
    merge(ws, 2, 1, 2, last - 6, f'{FAC}　／　{MAKER}　SV {sv}　　（R{year-2018}.{lastday.month}.{lastday.day} 実績記入）', f(9), AL)
    stamp_box(ws, 1, last - 2, ('組合', '照査', '作成'), ('', '', sv))
    r = 4
    merge(ws, r, 1, r + 1, 1, 'No', f(9, True), AC, fill('E2EFDA'))
    merge(ws, r, 2, r + 1, 2, '項　目', f(9, True), AC, fill('E2EFDA'))
    merge(ws, r, 3, r + 1, 3, '予/実', f(8, True), AC, fill('E2EFDA'))
    for j, (y, mth) in enumerate(months):
        merge(ws, r, c0 + 3 * j, r, c0 + 3 * j + 2, f'{mth}月', f(8, True), AC, fill('E2EFDA'))
        for k, t in enumerate('上中下'):
            put(ws, r + 1, c0 + 3 * j + k, t, f(7), AC, fill('E2EFDA'))
    box(ws, r, 1, r + 1, last)

    def ji(d):
        mi = months.index((d.year, d.month))
        return mi * 3 + (0 if d.day <= 10 else (1 if d.day <= 20 else 2))
    r = 6
    no = 0
    evs = []
    for ev in ('mid1', 'mid2', 'all'):
        if ev in W.TEIKI_EV[year]:
            down, w0, w1, up = W.STOPS[year][ev]
            evs.append((W.EVLABEL[ev] + ('（定期整備）' if ev == 'all' else ''), (down, up), (down, up), ev))
    if year == 2025:
        evs.append(('〃 2号IDF 軸受交換（重点）', (D(2025, 10, 14), D(2025, 10, 22)), (D(2025, 10, 14), D(2025, 10, 24)), 'all'))
    if year == 2024:
        evs.append(('〃 2号IDF 振動測定（重点監視）', (D(2024, 10, 16), D(2024, 10, 16)), (D(2024, 10, 16), D(2024, 10, 16)), 'all'))
    for name, plan, act, ev in evs:
        no += 1
        merge(ws, r, 1, r + 1, 1, no, f(9), AC)
        merge(ws, r, 2, r + 1, 2, name, f(9), AL)
        put(ws, r, 3, '予', f(8), AC); put(ws, r + 1, 3, '実', f(8), AC)
        for i in range(ji(plan[0]), ji(plan[1]) + 1):
            put(ws, r, c0 + i, None, fl=fill('9BC2E6'))
        for i in range(ji(act[0]), ji(act[1]) + 1):
            put(ws, r + 1, c0 + i, None, fl=fill('1F4E78'))
        rec.add(stem, ws.title, r, 'item', dict(No=no, 項目=name, plan=[[plan[0].isoformat(), plan[1].isoformat()]], actual=[[act[0].isoformat(), act[1].isoformat()]]),
                works=[w['id'] for w in W.WORKS if w['contract'] == f'T{year}' and w['event'] == ev][:1])
        r += 2
    merge(ws, r, 1, r, last, '他契約（組合発注）・法定点検', f(9, True), AL, fill('F2F2F2'))
    r += 1
    for w in sorted([w for w in W.OTHER if w['year'] == year], key=lambda w: w['plan'][0][0]):
        no += 1
        s, e = w['plan'][0]
        if not (months[0] <= (s.year, s.month) <= months[-1]):
            continue
        put(ws, r, 1, no, f(9), AC)
        put(ws, r, 2, w['name'], f(9), ALN)
        put(ws, r, 3, '予', f(8), AC)
        for i in range(ji(s), ji(e) + 1):
            put(ws, r, c0 + i, '●', f(8, color='7F6000'), AC)
        rec.add(stem, ws.title, r, 'item', dict(No=no, 項目=w['name'], plan=[[s.isoformat(), e.isoformat()]]), works=[w['id']])
        r += 1
    box(ws, 6, 1, r - 1, last)
    r += 1
    put(ws, r, 2, '凡例　予定 ■（淡）　実績 ■（濃）　● 他契約の実施日', f(8), ALN)
    r += 1
    put(ws, r, 2, '様式：R6.4 改定（n2）旧様式は停止ごとの工程表のみ', f(7, color='808080'), ALN)
    setup_print(ws, True, 9, 1, 1, f'A1:{L(last)}{r}')


def build_n2_year(year):
    fname = f'{FN}_工程表_令和{year-2018}年度定期整備業務委託_n2様式.xlsx'
    stem = rec.start(fname)
    wb = Workbook()
    props(wb, f'【架空】令和{year-2018}年度 やまぶき広域処理組合第二工場 定期整備業務委託 工程表', f'{MAKER} n2')
    ws = wb.active; ws.title = '年間工程'
    n2_annual(stem, ws, year)
    for ev in ('mid1', 'mid2', 'all'):
        if ev in W.TEIKI_EV[year]:
            ws = wb.create_sheet(W.EVLABEL[ev])
            n2_event_sheet(stem, ws, year, ev, W.TEIKI_EV[year][ev])
    wb.save(os.path.join(OUT, fname))
    return fname


# ------------------------------------------------------------------ n3（基幹改良）
def n3_overall(stem, ws, year):
    months = [(2022, m) for m in range(6, 13)] + [(2023, m) for m in range(1, 13)] + [(2024, m) for m in (1, 2, 3)]
    c0 = 3
    ws.column_dimensions['A'].width = 5; ws.column_dimensions['B'].width = 34
    for i in range(len(months)):
        ws.column_dimensions[L(c0 + i)].width = 4.2
    last = c0 + len(months) - 1
    merge(ws, 1, 1, 1, last, '【架空】基幹的設備改良工事（令和4〜5年度）　全体工程表', f(13, True), AL)
    merge(ws, 2, 1, 2, last, f'{FAC}　／　受注者 {MAKER}　現場代理人 n3　（' + ('第3回変更 R4.8.31' if year == 2022 else '第5回変更 R5.8.25') + '）', f(9), AL)
    r = 4
    put(ws, r, 1, 'No', f(9, True), AC, fill('FFF2CC')); put(ws, r, 2, '工種', f(9, True), AC, fill('FFF2CC'))
    for i, (y, m) in enumerate(months):
        put(ws, r, c0 + i, f"{'R' + str(y - 2018) + '.' if (i == 0 or m == 1) else ''}{m}", f(7, True), AC, fill('FFF2CC'))
    box(ws, r, 1, r, last)
    items = [('実施設計・承諾図', (2022, 6), (2022, 9)), ('機器製作（1号炉分・DCS）', (2022, 7), (2022, 10)),
             ('現地工事 1号炉（火格子・減温塔・BF）', (2022, 9), (2022, 12)), ('新DCS 盤据付（共通）', (2022, 11), (2022, 11)),
             ('R4 年度出来形確認', (2022, 12), (2022, 12)), ('機器製作（2号炉分）', (2023, 4), (2023, 8)),
             ('現地工事 2号炉（火格子・減温塔・BF）', (2023, 9), (2023, 12)), ('DCS 切替（全炉停止）', (2024, 2), (2024, 2)),
             ('完成検査・引渡し', (2024, 3), (2024, 3))]
    r += 1
    for no, (nm, a, b) in enumerate(items, 1):
        put(ws, r, 1, no, f(9), AC); put(ws, r, 2, nm, f(9), AL)
        ia, ib = months.index(a), months.index(b)
        ms = ia == ib and ('確認' in nm or '検査' in nm)
        for i in range(ia, ib + 1):
            put(ws, r, c0 + i, '◆' if ms else None, f(9, color='C00000'), AC,
                fill('F4B084') if '現地' in nm or '切替' in nm else (None if ms else fill('A9D08E')))
        rec.add(stem, ws.title, r, 'item', dict(No=no, 工種=nm, 期間=[f'{a[0]}-{a[1]:02d}', f'{b[0]}-{b[1]:02d}']))
        r += 1
    box(ws, 5, 1, r - 1, last)
    r += 1
    put(ws, r, 2, '※ 各年度の全炉停止（定期整備業務委託：別契約）は現地工事期間中に行う。R4：10/14〜10/30、R5：10/13〜10/29', f(8), ALN)
    setup_print(ws, True, 9, 1, 1, f'A1:{L(last)}{r}')


def n3_detail(stem, ws, title, sub, wbs_list, rng, bands, notes, works_by_wbs):
    d0, d1 = rng
    days = list(daterange(d0, d1))
    c0 = 6
    for c, w in zip('ABCDE', (4.5, 30, 8, 4, 4)):
        ws.column_dimensions[c].width = w
    for i in range(len(days)):
        ws.column_dimensions[L(c0 + i)].width = 1.9 if len(days) > 60 else 3.2
    last = c0 + len(days) - 1
    merge(ws, 1, 1, 1, max(10, last - 20), '【架空】' + title, f(13, True), AL)
    merge(ws, 2, 1, 2, max(10, last - 20), f'{FAC}　受注者 {MAKER}　現場代理人 n3　　{sub}', f(9), AL)
    step = 6 if len(days) > 60 else 3
    for k, (t, nm) in enumerate((('監督員', ''), ('現場代理人', 'n3'), ('作成', 'n3'))):
        cc = last - 3 * step + 1 + k * step
        merge(ws, 1, cc, 1, cc + step - 1, t, f(8), AC)
        merge(ws, 2, cc, 2, cc + step - 1, nm, f(9, color='C00000'), AC)
    box(ws, 1, last - 3 * step + 1, 2, last)
    ws.row_dimensions[2].height = 24
    r = 4
    for c, t in ((1, 'WBS'), (2, '工種・作業'), (3, '班'), (4, '人'), (5, '日')):
        merge(ws, r, c, r + 2, c, t, f(8, True), AC, fill('FFF2CC'))
    i = 0
    while i < len(days):
        j = i
        while j + 1 < len(days) and days[j + 1].month == days[i].month:
            j += 1
        merge(ws, r, c0 + i, r, c0 + j, f'{days[i].month}月', f(8, True), AC, fill('FFF2CC'))
        i = j + 1
    for i, d in enumerate(days):
        put(ws, r + 1, c0 + i, d.day, f(6), AC, fill('FCE4D6') if is_off(d) else fill('FFF2CC'))
        put(ws, r + 2, c0 + i, wd(d), f(6, color='C00000' if is_off(d) else '000000'), AC, fill('FCE4D6') if is_off(d) else fill('FFF2CC'))
    box(ws, r, 1, r + 2, last)
    r += 3
    for lab, spans in bands:
        merge(ws, r, 1, r, 5, lab, f(8, True), AL, fill('F2F2F2'))
        for i, d in enumerate(days):
            for (s, e, txt, col) in spans:
                if s <= d <= e:
                    put(ws, r, c0 + i, txt if d == s or (i == 0 and s < d) else None, f(6, True), ALN, fill(col))
        rec.add(stem, ws.title, r, 'item', dict(区分=lab, 状態=[[max(s, d0).isoformat(), min(e, d1).isoformat(), t] for s, e, t, c in spans]))
        r += 1
    box(ws, r - len(bands), 1, r - 1, last, med)
    st = r
    pers = {d: 0 for d in days}
    blk = [s for lab, sp in bands if lab.startswith('全停電') for s, e, t, c in sp]
    for (wbs, nm, team, n, spans, kind) in wbs_list:
        lvl = wbs.count('.')
        put(ws, r, 1, wbs, f(8, kind == 'H'), AL)
        put(ws, r, 2, ('　' * lvl) + nm, f(8, kind == 'H', 'C00000' if '追加' in nm else '000000'), AL, fill('EDEDED') if kind == 'H' else None)
        put(ws, r, 3, team, f(7), AC)
        put(ws, r, 4, n, f(7), AC)
        nd = sum(1 for s, e in spans for d in daterange(s, e) if not is_off(d) and d not in blk)
        put(ws, r, 5, nd if nd else None, f(7), AC)
        for i, d in enumerate(days):
            if is_off(d):
                put(ws, r, c0 + i, None, fl=fill('FCE4D6'))
            if d in blk:
                put(ws, r, c0 + i, None, fl=fill('FFC000'))
            for s, e in spans:
                if s <= d <= e:
                    if kind == 'M':
                        put(ws, r, c0 + i, '◆', f(8, color='C00000'), AC)
                    elif kind == 'K':
                        put(ws, r, c0 + i, '▲', f(8, color='0070C0'), AC)
                    elif not is_off(d) and d not in blk:
                        put(ws, r, c0 + i, None, fl=fill('E06666' if kind == 'C' else '70AD47'))
                        if n:
                            pers[d] += n
        if kind == 'H':
            rec.add(stem, ws.title, r, 'parent', dict(WBS=wbs, 工種=nm))
        else:
            rec.add(stem, ws.title, r, 'item', dict(WBS=wbs, 工種=nm, 班=team, 人=n, 日=nd or None, plan=fmt_spans(spans), 種別={'C': 'クリティカル', 'N': '一般', 'M': 'マイルストン', 'K': '立会検査'}[kind]),
                    works=[works_by_wbs.get(wbs)])
        r += 1
    box(ws, st, 1, r - 1, last)
    merge(ws, r, 1, r, 5, '日別 作業員（人）', f(8, True), AC, fill('F2F2F2'))
    pv = {}
    for i, d in enumerate(days):
        v = pers[d]
        put(ws, r, c0 + i, v if v else None, Font(name=FONT, size=5), AC, fill('F2F2F2'))
        if v:
            pv[d.isoformat()] = v
    rec.add(stem, ws.title, r, 'total', dict(日別作業員=pv, unit='人'))
    box(ws, r, 1, r, last)
    r += 2
    for t in notes:
        merge(ws, r, 2, r, last, t, f(8), ALN)
        r += 1
    setup_print(ws, True, 8, 1, 1, f'A1:{L(last)}{r}')


def build_n3():
    out = []
    wbw = {(w['event'], w.get('wbs')): w['id'] for w in W.WORKS if w['contract'] == 'KIKAN'}
    # R4
    fname = f'{FN}_工程表_基幹的設備改良工事_令和4年度現地工事_n3詳細工程.xlsx'
    stem = rec.start(fname)
    wb = Workbook()
    props(wb, '【架空】やまぶき広域処理組合第二工場 基幹的設備改良工事（令和4〜5年度）詳細工程表 令和4年度現地工事', f'{MAKER} 現場代理人 n3')
    ws = wb.active; ws.title = '全体工程'
    n3_overall(stem, ws, 2022)
    ws = wb.create_sheet('現地工事詳細(R4)')
    d0, d1 = S.KIKAN_R4_RANGE
    bands = [
        ('1号炉 運転状況', [(d0, D(2022, 9, 20), '運転', 'E2EFDA'), (D(2022, 9, 21), D(2022, 12, 4), '停止（本工事）', 'BFBFBF'), (D(2022, 12, 5), d1, '試運転→運転', 'FFE699')]),
        ('2号炉 運転状況', [(d0, D(2022, 10, 14), '単独運転（時間延長）', 'E2EFDA'), (D(2022, 10, 15), D(2022, 10, 29), '全炉停止(定期整備)', 'F8CBAD'), (D(2022, 10, 30), d1, '単独運転', 'E2EFDA')]),
        ('全停電（受変電年次）', [(D(2022, 10, 23), D(2022, 10, 23), '停', 'FFC000')]),
        ('炉室ホイスト 使用', [(D(2022, 10, 20), D(2022, 10, 22), '定整優先', 'F8CBAD'), (D(2022, 10, 24), D(2022, 10, 26), '本工事', 'F4B084')]),
    ]
    notes = ['凡例　■ クリティカル作業　■ 一般作業　◆ マイルストン　▲ 立会検査　網掛け 日曜・祝日（休工）　橙 全停電（全作業中止）',
             '調整事項　① 10/15〜10/29 は定期整備（別契約）の全炉停止。2号炉側の作業と動線・荷揚げを分ける（2号炉側立入は定期整備 SV n1 の許可）。',
             '　　　　　② 10/20〜10/22 炉室ホイストは定期整備（2号火格子片の荷揚げ）優先。新火格子搬入は 10/24〜26 に変更（第1版 10/19〜21）。',
             '　　　　　③ 10/23 全停電（受変電年次点検）は全作業中止。仮設電源も停止。']
    n3_detail(stem, ws, '基幹的設備改良工事　令和4年度　現地工事詳細工程表（1号炉・共通DCS盤）',
              '版：第2版 R4.10.21（定期整備 全炉停止・全停電・炉室ホイスト調整を反映）', S.KIKAN_R4, (d0, d1), bands, notes,
              {k[1]: v for k, v in wbw.items() if k[0] == 'kikanR4'})
    wb.active = 1
    wb.save(os.path.join(OUT, fname)); out.append(fname)
    # R5
    fname = f'{FN}_工程表_基幹的設備改良工事_令和5年度現地工事_n3詳細工程.xlsx'
    stem = rec.start(fname)
    wb = Workbook()
    props(wb, '【架空】やまぶき広域処理組合第二工場 基幹的設備改良工事（令和4〜5年度）詳細工程表 令和5年度現地工事', f'{MAKER} 現場代理人 n3')
    ws = wb.active; ws.title = '全体工程'
    n3_overall(stem, ws, 2023)
    ws = wb.create_sheet('現地工事詳細(R5)')
    d0, d1 = W.KIKAN_R5_RANGE
    r5 = [x for x in W.KIKAN_R5 if not x[0].startswith('6')]
    bands = [
        ('2号炉 運転状況', [(d0, D(2023, 9, 19), '運転', 'E2EFDA'), (D(2023, 9, 20), D(2023, 12, 3), '停止（本工事）', 'BFBFBF'), (D(2023, 12, 4), d1, '試運転→運転', 'FFE699')]),
        ('1号炉 運転状況', [(d0, D(2023, 10, 13), '単独運転（時間延長）', 'E2EFDA'), (D(2023, 10, 14), D(2023, 10, 28), '全炉停止(定期整備)', 'F8CBAD'), (D(2023, 10, 29), d1, '単独運転', 'E2EFDA')]),
        ('全停電（受変電年次）', [(D(2023, 10, 22), D(2023, 10, 22), '停', 'FFC000')]),
        ('炉室ホイスト 使用', [(D(2023, 10, 18), D(2023, 10, 20), '定整優先', 'F8CBAD'), (D(2023, 10, 23), D(2023, 10, 25), '本工事', 'F4B084')]),
    ]
    notes = ['凡例　■ クリティカル作業　■ 一般作業　◆ マイルストン　▲ 立会検査　網掛け 日曜・祝日（休工）　橙 全停電（全作業中止）',
             '調整事項　① 10/14〜10/28 は定期整備（別契約）の全炉停止。1号炉側の立入は定期整備 SV n1 の許可。',
             '　　　　　② 10/18〜10/20 炉室ホイストは定期整備（1号灰押出装置 摺動板）優先。新火格子搬入は 10/23〜25。',
             '　　　　　③ 2.6 2号煙道 減肉部補修 は 10/31 調査で追加（第6回変更 R6.1.12）。',
             '　　　　　④ DCS 切替は 1・2号炉とも停止する。R6.2.2 立下げ、別シート参照。']
    n3_detail(stem, ws, '基幹的設備改良工事　令和5年度　現地工事詳細工程表（2号炉）', '版：第3版 R5.11.6（2号煙道補修 追加を反映）', r5, (d0, d1), bands, notes,
              {k[1]: v for k, v in wbw.items() if k[0] == 'kikanR5'})
    ws = wb.create_sheet('DCS切替(R6.2)')
    d0, d1 = W.KIKAN_DCS_RANGE
    r6 = [x for x in W.KIKAN_R5 if x[0].startswith('6')]
    bands = [('1号炉 運転状況', [(d0, D(2024, 2, 2), '運転', 'E2EFDA'), (D(2024, 2, 3), D(2024, 2, 15), '停止（DCS切替）', 'BFBFBF'), (D(2024, 2, 16), d1, '立上げ→運転', 'FFE699')]),
             ('2号炉 運転状況', [(d0, D(2024, 2, 2), '運転', 'E2EFDA'), (D(2024, 2, 3), D(2024, 2, 17), '停止（DCS切替）', 'BFBFBF'), (D(2024, 2, 18), d1, '立上げ→運転', 'FFE699')])]
    notes = ['凡例　■ クリティカル作業　◆ マイルストン　網掛け 日曜・祝日',
             '調整事項　① 切替期間中は運転委託会社が夜間立会（組合別契約）。ごみはピット貯留。② 旧DCS（CRT・グラフィックパネル）は撤去し産廃処分。']
    n3_detail(stem, ws, '基幹的設備改良工事　DCS 切替 工程表（全炉停止）', '版：第1版 R6.1.15', r6, (d0, d1), bands, notes,
              {k[1]: v for k, v in wbw.items() if k[0] == 'kikanR5'})
    wb.active = 1
    wb.save(os.path.join(OUT, fname)); out.append(fname)
    return out


def build_all():
    out = []
    for y in (2021, 2022, 2023):
        out.append(build_n1_year(y))
    for y in (2024, 2025):
        out.append(build_n2_year(y))
    out += build_n3()
    return out


if __name__ == '__main__':
    print(build_all())
