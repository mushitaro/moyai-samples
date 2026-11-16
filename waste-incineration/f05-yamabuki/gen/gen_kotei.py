# -*- coding: utf-8 -*-
"""工程表 3 種
 (1) n1 様式：定期整備（簡易）。1 停止 1 シート、日単位の塗りつぶし、計画のみ＋改訂行
 (2) n2 様式（R6〜）：年間 1 枚＋全炉停止の日程（予定・実績の 2 段）
 (3) n3 様式：基幹改良 現地工事 詳細工程（WBS、班・人数、クリティカル、停止の分け合い）
"""
import os, datetime as dt
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter as L
from common import *
from schedules import *

HOL = {dt.date(2022, 9, 19), dt.date(2022, 9, 23), dt.date(2022, 10, 10), dt.date(2022, 11, 3), dt.date(2022, 11, 23),
       dt.date(2025, 10, 13), dt.date(2025, 11, 3), dt.date(2025, 11, 24)}


def is_off(d):
    return d.weekday() == 6 or d in HOL


def stamp_box(ws, r, c, labels=('確認', '照査', '作成'), names=('', '', '')):
    for i, (t, n) in enumerate(zip(labels, names)):
        put(ws, r, c + i, t, f(8), AC)
        put(ws, r + 1, c + i, n, f(9, color='C00000'), AC)
    box(ws, r, c, r + 1, c + len(labels) - 1)
    ws.row_dimensions[r + 1].height = 26


# ------------------------------------------------------------------ n1
def n1_sheet(ws, title, days, rows, stop_rows, revisions, persons_by_day):
    c0 = 5
    ws.column_dimensions['A'].width = 6
    ws.column_dimensions['B'].width = 3.5
    ws.column_dimensions['C'].width = 34
    ws.column_dimensions['D'].width = 10
    for i in range(len(days)):
        ws.column_dimensions[L(c0 + i)].width = 3.6
    last = c0 + len(days) - 1
    merge(ws, 1, 1, 1, last - 4, '【架空】' + title, f(13, True), AL)
    merge(ws, 2, 1, 2, last - 4, f'施設：{FAC}　　施工：{MAKER}', f(9), AL)
    stamp_box(ws, 1, last - 2, ('組合', '照査', '作成'), ('', '', 'n1'))
    r = 4
    merge(ws, r, 1, r + 1, 1, '区分', f(9, True), AC, fill('DDEBF7'))
    merge(ws, r, 2, r + 1, 2, 'No', f(9, True), AC, fill('DDEBF7'))
    merge(ws, r, 3, r + 1, 3, '作　業　項　目', f(9, True), AC, fill('DDEBF7'))
    merge(ws, r, 4, r + 1, 4, '担当', f(9, True), AC, fill('DDEBF7'))
    for i, d in enumerate(days):
        put(ws, r, c0 + i, f'{d.month}/{d.day}' if (i == 0 or d.day == 1) else d.day, f(7), AC, fill('DDEBF7'))
        put(ws, r + 1, c0 + i, wd(d), f(7, color='C00000' if is_off(d) else '000000'), AC, fill('FCE4D6') if is_off(d) else fill('DDEBF7'))
    box(ws, r, 1, r + 1, last)
    r += 2
    # 停止の帯
    for lab, spans in stop_rows:
        merge(ws, r, 1, r, 4, lab, f(8), AL, fill('F2F2F2'))
        for i, d in enumerate(days):
            for (s, e, txt, col) in spans:
                if s <= d <= e:
                    put(ws, r, c0 + i, txt if d == s else None, f(7, True), ALN, fill(col))
        r += 1
    box(ws, r - len(stop_rows), 1, r - 1, last)
    st = r
    prev = None
    no = 0
    for row in rows:
        kubun, name, team, spans, rev = row
        no += 1
        put(ws, r, 1, kubun if kubun != prev else None, f(9), AC)
        put(ws, r, 2, no, f(8), AC)
        put(ws, r, 3, name, f(9, color='C00000' if rev else '000000'), AL)
        put(ws, r, 4, team, f(8), AC)
        for i, d in enumerate(days):
            if is_off(d):
                put(ws, r, c0 + i, None, fl=fill('FCE4D6'))
            for s, e in spans:
                if s <= d <= e:
                    put(ws, r, c0 + i, None, fl=fill('FF7C80' if rev else '5B9BD5'))
        if rev:
            ws.cell(r, last).value = None
            put(ws, r, 3, name, f(9, color='C00000'), AL).comment = Comment('改1（R4.10.20）で追加', 'n1')
        prev = kubun
        r += 1
    box(ws, st, 1, r - 1, last)
    merge(ws, r, 1, r, 4, '予定人員（人）', f(9, True), AC, fill('F2F2F2'))
    for i, d in enumerate(days):
        v = persons_by_day.get(d)
        put(ws, r, c0 + i, v if v else None, f(7), AC, fill('F2F2F2'))
    box(ws, r, 1, r, last)
    r += 2
    put(ws, r, 1, '改訂', f(8, True), AL)
    for t in revisions:
        merge(ws, r, 2, r, last, t, f(8, color='C00000' if '改1' in t else '000000'), AL)
        r += 1
    put(ws, r, 2, '凡例　■ 作業予定　■ 追加作業（赤）　網掛け 日曜・祝日', f(8), ALN)
    setup_print(ws, True, 9, 1, 1, f'A1:{L(last)}{r}')
    return ws


def build_n1(path):
    wb = Workbook()
    props(wb, '【架空】令和4年度 やまぶき広域処理組合第二工場 定期整備業務委託 工程表', f'{MAKER} n1')
    # 2号炉中間点検
    ws = wb.active; ws.title = '2号炉中間点検'
    m = R4_STOPS['mid2']
    days = list(daterange(dt.date(2022, 6, 9), dt.date(2022, 6, 19)))
    rows = [(k, n, t, [sp], 0) for (k, n, t, sp, _) in R4_MID2]
    stop_rows = [('2号炉', [(dt.date(2022, 6, 9), m['down'], '運転', 'E2EFDA'), (m['down'] + dt.timedelta(1), m['up'] - dt.timedelta(1), '停止', 'D9D9D9'), (m['up'], m['up'], '立上', 'FFE699')]),
                 ('1号炉', [(days[0], days[-1], '運転（単独運転・時間延長）', 'E2EFDA')])]
    pers = {}
    for d in days:
        if m['work'][0] <= d <= m['work'][1]:
            pers[d] = 1 + 5 + 4 + 3
    n1_sheet(ws, '令和4年度 やまぶき広域処理組合第二工場 定期整備業務委託　2号炉中間点検 工程表', days, rows, stop_rows,
             ['改0　R4.5.27　提出'], pers)
    # 全炉停止
    ws = wb.create_sheet('全炉停止')
    days = list(daterange(dt.date(2022, 10, 13), dt.date(2022, 10, 31)))
    rows = []
    for (iid, kubun, name, team, plan, act, ms, rev) in R4_ALL:
        rows.append((kubun, name, team, plan, rev))
    a = R4_STOPS['all']
    stop_rows = [('2号炉', [(days[0], a['down'], '運転', 'E2EFDA'), (a['down'] + dt.timedelta(1), a['up'] - dt.timedelta(1), '停止（冷却 10/15・16）', 'D9D9D9'), (a['up'], a['up'], '立上', 'FFE699')]),
                 ('1号炉', [(days[0], days[-1], '基幹的設備改良工事のため停止中（別契約）', 'BFBFBF')]),
                 ('全停電', [(R4_STOPS['blackout'], R4_STOPS['blackout'], '停電', 'FFC000')])]
    pers = {}
    for d in days:
        if a['work'][0] <= d <= a['work'][1] and not is_off(d):
            n = 1
            teams = set()
            for (iid, kubun, name, team, plan, act, ms, rev) in R4_ALL:
                if rev:
                    continue
                for s, e in plan:
                    if s <= d <= e:
                        teams.add(team)
            n += sum({'機械1班': 5, '機械2班': 4, '電計班': 3, '外注(足場)': 3, '外注(クレーン)': 3, '外注(計器)': 2, '全班': 7}.get(t, 0) for t in teams)
            pers[d] = n
    pers[R4_STOPS['blackout']] = 2
    n1_sheet(ws, '令和4年度 やまぶき広域処理組合第二工場 定期整備業務委託　全炉停止 工程表', days, rows, stop_rows,
             ['改0　R4.9.30　提出', '改1　R4.10.20　No.8 灰押出装置 摺動板取替 追加（10/19 点検で摩耗大、組合了承）'], pers)
    wb.active = 1
    wb.save(path)


# ------------------------------------------------------------------ n2
def build_n2(path):
    wb = Workbook()
    props(wb, '【架空】令和7年度 やまぶき広域処理組合第二工場 定期整備業務委託 工程表', f'{MAKER} n2')
    ws = wb.active; ws.title = '年間工程'
    # 旬単位 36 列
    c0 = 4
    ws.column_dimensions['A'].width = 4
    ws.column_dimensions['B'].width = 30
    ws.column_dimensions['C'].width = 5
    months = [(2025, m) for m in range(4, 13)] + [(2026, m) for m in (1, 2, 3)]
    for i in range(36):
        ws.column_dimensions[L(c0 + i)].width = 2.6
    last = c0 + 35
    merge(ws, 1, 1, 1, last - 6, '【架空】令和7年度 やまぶき広域処理組合第二工場 定期整備業務委託　年間工程表', f(13, True), AL)
    merge(ws, 2, 1, 2, last - 6, f'{FAC}　／　{MAKER}　SV n2　　（R7.10.31 実績記入）', f(9), AL)
    stamp_box(ws, 1, last - 2, ('組合', '照査', '作成'), ('', '', 'n2'))
    r = 4
    merge(ws, r, 1, r + 1, 1, 'No', f(9, True), AC, fill('E2EFDA'))
    merge(ws, r, 2, r + 1, 2, '項　目', f(9, True), AC, fill('E2EFDA'))
    merge(ws, r, 3, r + 1, 3, '予/実', f(8, True), AC, fill('E2EFDA'))
    for j, (y, mth) in enumerate(months):
        merge(ws, r, c0 + 3 * j, r, c0 + 3 * j + 2, f'{mth}月', f(8, True), AC, fill('E2EFDA'))
        for k, t in enumerate('上中下'):
            put(ws, r + 1, c0 + 3 * j + k, t, f(7), AC, fill('E2EFDA'))
    box(ws, r, 1, r + 1, last)

    def jun_idx(d):
        mi = months.index((d.year, d.month))
        k = 0 if d.day <= 10 else (1 if d.day <= 20 else 2)
        return mi * 3 + k
    r = 6
    no = 0
    for name, plan, act in R7_EVENTS:
        no += 1
        merge(ws, r, 1, r + 1, 1, no, f(9), AC)
        merge(ws, r, 2, r + 1, 2, name, f(9), AL)
        put(ws, r, 3, '予', f(8), AC); put(ws, r + 1, 3, '実', f(8), AC)
        for i in range(jun_idx(plan[0]), jun_idx(plan[1]) + 1):
            put(ws, r, c0 + i, None, fl=fill('9BC2E6'))
        for i in range(jun_idx(act[0]), jun_idx(act[1]) + 1):
            put(ws, r + 1, c0 + i, None, fl=fill('1F4E78'))
        r += 2
    merge(ws, r, 1, r, last, '他契約（組合発注）・法定点検', f(9, True), AL, fill('F2F2F2'))
    r += 1
    for name, sp in R7_OTHER:
        no += 1
        put(ws, r, 1, no, f(9), AC)
        put(ws, r, 2, name.strip(), f(9), ALN)
        put(ws, r, 3, '予', f(8), AC)
        for i in range(jun_idx(sp[0]), jun_idx(sp[1]) + 1):
            put(ws, r, c0 + i, '●', f(8, color='7F6000'), AC)
        r += 1
    box(ws, 6, 1, r - 1, last)
    r += 1
    put(ws, r, 2, '凡例　予定 ■（淡）　実績 ■（濃）　● 他契約の実施日　　※10月以降の他契約は予定', f(8), ALN)
    r += 1
    put(ws, r, 2, '様式：R6.4 改定（n2）旧様式は停止ごとの工程表のみ', f(7, color='808080'), ALN)
    setup_print(ws, True, 9, 1, 1, f'A1:{L(last)}{r}')

    # 全炉停止 日程（n1 様式をコピーして 予/実 の 2 段に直したもの）
    ws = wb.create_sheet('全炉停止')
    days = list(daterange(dt.date(2025, 10, 9), dt.date(2025, 10, 28)))
    c0 = 7
    for c, w in zip('ABCDEF', (6, 3.5, 33, 11, 4.5, 4)):
        ws.column_dimensions[c].width = w
    for i in range(len(days)):
        ws.column_dimensions[L(c0 + i)].width = 3.6
    last = c0 + len(days) - 1
    merge(ws, 1, 1, 1, last - 4, '【架空】令和7年度 やまぶき広域処理組合第二工場 定期整備業務委託　全炉停止 工程表（予定・実績）', f(13, True), AL)
    merge(ws, 2, 1, 2, last - 4, f'施設：{FAC}　　施工：{MAKER}　　予定 R7.9.26 提出 ／ 実績 R7.10.31 記入', f(9), AL)
    stamp_box(ws, 1, last - 2, ('組合', '照査', '作成'), ('', '', 'n2'))
    # n1 様式の改訂履歴の行（前年度ファイルの残り、取消線）
    merge(ws, 3, 1, 3, 12, '改0　R6.9.27　提出', f(7, color='A6A6A6', strike=True), AL)
    r = 4
    for c, t in ((1, '区分'), (2, 'No'), (3, '作　業　項　目'), (4, '担当'), (5, '人数'), (6, '予実')):
        merge(ws, r, c, r + 1, c, t, f(9 if c < 5 else 8, True), AC, fill('DDEBF7'))
    for i, d in enumerate(days):
        put(ws, r, c0 + i, f'{d.month}/{d.day}' if i == 0 else d.day, f(7), AC, fill('DDEBF7'))
        put(ws, r + 1, c0 + i, wd(d), f(7, color='C00000' if is_off(d) else '000000'), AC, fill('FCE4D6') if is_off(d) else fill('DDEBF7'))
    box(ws, r, 1, r + 1, last)
    r += 2
    a0, a1 = R7_ALL_RANGE
    stops = [('1号炉', [(days[0], a0, '運転', 'E2EFDA'), (a0 + dt.timedelta(1), dt.date(2025, 10, 26), '停止', 'D9D9D9'), (dt.date(2025, 10, 27), dt.date(2025, 10, 27), '立上', 'FFE699')]),
             ('2号炉', [(days[0], a0, '運転', 'E2EFDA'), (a0 + dt.timedelta(1), dt.date(2025, 10, 25), '停止', 'D9D9D9'), (dt.date(2025, 10, 26), dt.date(2025, 10, 26), '立上', 'FFE699')]),
             ('全停電', [(dt.date(2025, 10, 19), dt.date(2025, 10, 19), '停電', 'FFC000')])]
    for lab, spans in stops:
        merge(ws, r, 1, r, 6, lab, f(8), AL, fill('F2F2F2'))
        for i, d in enumerate(days):
            for (s, e, txt, col) in spans:
                if s <= d <= e:
                    put(ws, r, c0 + i, txt if d == s else None, f(7, True), ALN, fill(col))
        r += 1
    box(ws, r - 3, 1, r - 1, last)
    st = r
    prev = None
    for no, (kubun, name, team, n, plan, act, ms) in enumerate(R7_ALL, 1):
        merge(ws, r, 1, r + 1, 1, kubun if kubun != prev else None, f(9), AC)
        merge(ws, r, 2, r + 1, 2, no, f(8), AC)
        merge(ws, r, 3, r + 1, 3, name, f(9), AL)
        merge(ws, r, 4, r + 1, 4, team, f(8), AC)
        merge(ws, r, 5, r + 1, 5, n, f(8), AC)
        put(ws, r, 6, '予', f(7), AC); put(ws, r + 1, 6, '実', f(7), AC)
        delayed = act != plan
        for i, d in enumerate(days):
            for s, e in plan:
                if s <= d <= e:
                    put(ws, r, c0 + i, None, fl=fill('9BC2E6'))
            for s, e in act:
                if s <= d <= e:
                    late = delayed and d > max(e2 for s2, e2 in plan)
                    put(ws, r + 1, c0 + i, None, fl=fill('C00000' if late else '1F4E78'))
        if delayed:
            ws.cell(r + 1, 6).comment = Comment('実績が予定より延びた（日報の週報 10/20 週 参照）', 'n2')
        for rr in (r, r + 1):
            for c in range(1, last + 1):
                ws.cell(rr, c).border = Border(left=thin, right=thin, top=thin if rr == r else dot, bottom=dot if rr == r else thin)
        ws.row_dimensions[r].height = 11; ws.row_dimensions[r + 1].height = 11
        prev = kubun
        r += 2
    r += 1
    put(ws, r, 3, '凡例　予定 ■（淡）　実績 ■（濃）　実績の遅れ ■（赤）', f(8), ALN)
    r += 1
    put(ws, r, 3, '※ No.12 2号IDF：10/20 試運転 振動 4.8mm/s → 再バランス、10/24 完了（2日延長）。No.23 ポンプ(A)：1日延長。', f(8), ALN)
    setup_print(ws, True, 9, 1, 1, f'A1:{L(last)}{r}')
    wb.save(path)


# ------------------------------------------------------------------ n3
def build_n3(path):
    wb = Workbook()
    props(wb, '【架空】やまぶき広域処理組合第二工場 基幹的設備改良工事（令和4〜5年度）詳細工程表 令和4年度現地工事', f'{MAKER} 現場代理人 n3')
    ws = wb.active; ws.title = '全体工程'
    months = [(2022, m) for m in range(6, 13)] + [(2023, m) for m in range(1, 13)] + [(2024, m) for m in (1, 2, 3)]
    c0 = 3
    ws.column_dimensions['A'].width = 5; ws.column_dimensions['B'].width = 34
    for i in range(len(months)):
        ws.column_dimensions[L(c0 + i)].width = 4.2
    last = c0 + len(months) - 1
    merge(ws, 1, 1, 1, last, '【架空】基幹的設備改良工事（令和4〜5年度）　全体工程表', f(13, True), AL)
    merge(ws, 2, 1, 2, last, f'{FAC}　／　受注者 {MAKER}　現場代理人 n3　（第3回変更 R4.8.31）', f(9), AL)
    r = 4
    put(ws, r, 1, 'No', f(9, True), AC, fill('FFF2CC')); put(ws, r, 2, '工種', f(9, True), AC, fill('FFF2CC'))
    for i, (y, m) in enumerate(months):
        put(ws, r, c0 + i, f"{'R' + str(y - 2018) + '.' if m in (1, 6) and (i == 0 or m == 1) else ''}{m}", f(7, True), AC, fill('FFF2CC'))
    box(ws, r, 1, r, last)
    items = [
        ('実施設計・承諾図', (2022, 6), (2022, 9)),
        ('機器製作（1号炉分・DCS）', (2022, 7), (2022, 10)),
        ('現地工事 1号炉（火格子・減温塔・BF）', (2022, 9), (2022, 12)),
        ('新DCS 盤据付（共通）', (2022, 11), (2022, 11)),
        ('R4 年度出来形確認', (2022, 12), (2022, 12)),
        ('機器製作（2号炉分）', (2023, 4), (2023, 8)),
        ('現地工事 2号炉（火格子・減温塔・BF）', (2023, 9), (2023, 12)),
        ('DCS 切替（全炉停止）', (2024, 2), (2024, 2)),
        ('完成検査・引渡し', (2024, 3), (2024, 3)),
    ]
    r += 1
    for no, (nm, a, b) in enumerate(items, 1):
        put(ws, r, 1, no, f(9), AC); put(ws, r, 2, nm, f(9), AL)
        ia, ib = months.index(a), months.index(b)
        for i in range(ia, ib + 1):
            put(ws, r, c0 + i, '◆' if ia == ib and ('確認' in nm or '検査' in nm) else None, f(9, color='C00000'), AC,
                fill('F4B084') if '現地' in nm or '切替' in nm else (None if ia == ib and ('確認' in nm or '検査' in nm) else fill('A9D08E')))
        r += 1
    box(ws, 5, 1, r - 1, last)
    r += 1
    put(ws, r, 2, '※ 各年度の全炉停止（定期整備業務委託：別契約）は現地工事期間中に行う。R4：10/14〜10/30、R5：10/13〜10/29（予定）', f(8), ALN)
    setup_print(ws, True, 9, 1, 1, f'A1:{L(last)}{r}')

    ws = wb.create_sheet('現地工事詳細(R4)')
    d0, d1 = KIKAN_R4_RANGE
    days = list(daterange(d0, d1))
    c0 = 6
    for c, w in zip('ABCDE', (4.5, 30, 8, 4, 4)):
        ws.column_dimensions[c].width = w
    for i in range(len(days)):
        ws.column_dimensions[L(c0 + i)].width = 1.9
    last = c0 + len(days) - 1
    merge(ws, 1, 1, 1, 50, '【架空】基幹的設備改良工事　令和4年度　現地工事詳細工程表（1号炉・共通DCS盤）', f(13, True), AL)
    merge(ws, 2, 1, 2, 50, f'{FAC}　受注者 {MAKER}　現場代理人 n3　　版：第2版 R4.10.21（定期整備 全炉停止・全停電・炉室ホイスト調整を反映）', f(9), AL)
    for k, (t, nm) in enumerate((('監督員', ''), ('現場代理人', 'n3'), ('作成', 'n3'))):
        cc = last - 17 + k * 6
        merge(ws, 1, cc, 1, cc + 5, t, f(8), AC)
        merge(ws, 2, cc, 2, cc + 5, nm, f(9, color='C00000'), AC)
    box(ws, 1, last - 17, 2, last)
    ws.row_dimensions[2].height = 24
    r = 4
    for c, t in ((1, 'WBS'), (2, '工種・作業'), (3, '班'), (4, '人'), (5, '日')):
        merge(ws, r, c, r + 2, c, t, f(8, True), AC, fill('FFF2CC'))
    # 月・日・曜日
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
    # 停止を分け合う帯
    bands = [
        ('1号炉 運転状況', [(d0, dt.date(2022, 9, 20), '運転', 'E2EFDA'), (dt.date(2022, 9, 21), dt.date(2022, 12, 4), '停止（本工事）', 'BFBFBF'), (dt.date(2022, 12, 5), d1, '試運転→運転', 'FFE699')]),
        ('2号炉 運転状況', [(d0, dt.date(2022, 10, 14), '単独運転（時間延長）', 'E2EFDA'), (dt.date(2022, 10, 15), dt.date(2022, 10, 29), '全炉停止(定期整備)', 'F8CBAD'), (dt.date(2022, 10, 30), d1, '単独運転', 'E2EFDA')]),
        ('全停電（受変電年次）', [(dt.date(2022, 10, 23), dt.date(2022, 10, 23), '停', 'FFC000')]),
        ('炉室ホイスト 使用', [(dt.date(2022, 10, 20), dt.date(2022, 10, 22), '定整優先', 'F8CBAD'), (dt.date(2022, 10, 24), dt.date(2022, 10, 26), '本工事', 'F4B084')]),
    ]
    for lab, spans in bands:
        merge(ws, r, 1, r, 5, lab, f(8, True), AL, fill('F2F2F2'))
        for i, d in enumerate(days):
            for (s, e, txt, col) in spans:
                if s <= d <= e:
                    put(ws, r, c0 + i, txt if d == s else None, f(6, True), ALN, fill(col))
        r += 1
    box(ws, r - len(bands), 1, r - 1, last, med)
    st = r
    pers = {d: 0 for d in days}
    for (wbs, nm, team, n, spans, kind) in KIKAN_R4:
        lvl = wbs.count('.')
        put(ws, r, 1, wbs, f(8, kind == 'H'), AL)
        put(ws, r, 2, ('　' * lvl) + nm, f(8, kind == 'H'), AL, fill('EDEDED') if kind == 'H' else None)
        put(ws, r, 3, team, f(7), AC)
        put(ws, r, 4, n, f(7), AC)
        nd = sum(1 for s, e in spans for d in daterange(s, e) if not is_off(d))
        put(ws, r, 5, nd if nd else None, f(7), AC)
        for i, d in enumerate(days):
            if is_off(d):
                put(ws, r, c0 + i, None, fl=fill('FCE4D6'))
            if d == dt.date(2022, 10, 23):
                put(ws, r, c0 + i, None, fl=fill('FFC000'))
            for s, e in spans:
                if s <= d <= e:
                    if kind == 'M':
                        put(ws, r, c0 + i, '◆', f(8, color='C00000'), AC)
                    elif kind == 'K':
                        put(ws, r, c0 + i, '▲', f(8, color='0070C0'), AC)
                    elif not is_off(d) and d != dt.date(2022, 10, 23):
                        put(ws, r, c0 + i, None, fl=fill('E06666' if kind == 'C' else '70AD47'))
                        if n:
                            pers[d] += n
        r += 1
    box(ws, st, 1, r - 1, last)
    merge(ws, r, 1, r, 5, '日別 作業員（人）', f(8, True), AC, fill('F2F2F2'))
    for i, d in enumerate(days):
        v = pers[d]
        put(ws, r, c0 + i, v if v else None, Font(name=FONT, size=5), AC, fill('F2F2F2'))
    box(ws, r, 1, r, last)
    r += 2
    for t in ('凡例　■ クリティカル作業　■ 一般作業　◆ マイルストン　▲ 立会検査　網掛け 日曜・祝日（休工）　橙 全停電（全作業中止）',
              '調整事項　① 10/15〜10/29 は定期整備（別契約）の全炉停止。2号炉側の作業と動線・荷揚げを分ける（2号炉側立入は定期整備 SV n1 の許可）。',
              '　　　　　② 10/20〜10/22 炉室ホイストは定期整備（2号火格子片の荷揚げ）優先。新火格子搬入は 10/24〜26 に変更（第1版 10/19〜21）。',
              '　　　　　③ 10/23 全停電（受変電年次点検）は全作業中止。仮設電源も停止。'):
        merge(ws, r, 2, r, last, t, f(8), AL)
        r += 1
    setup_print(ws, True, 8, 1, 1, f'A1:{L(last)}{r}')
    wb.active = 1
    wb.save(path)


if __name__ == '__main__':
    build_n1(os.path.join(OUT, f'{FN}_工程表_令和4年度定期整備業務委託_n1様式.xlsx'))
    build_n2(os.path.join(OUT, f'{FN}_工程表_令和7年度定期整備業務委託_n2様式.xlsx'))
    build_n3(os.path.join(OUT, f'{FN}_工程表_基幹的設備改良工事_令和4年度現地工事_n3詳細工程.xlsx'))
