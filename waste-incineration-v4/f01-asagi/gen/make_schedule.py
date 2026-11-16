# -*- coding: utf-8 -*-
import datetime as dt
from openpyxl import Workbook
from openpyxl.comments import Comment
from xlutil import *
from f01_contracts import *
import f01_truth as TR

WD = '月火水木金土日'
META = dict(made='H23.6.20', agent='m2', agent3='m3', op='運転委託先（あさぎ環境運転サービス）')
HOOK = [None]
def _hk(*a):
    if HOOK[0]:
        HOOK[0](*a)

def daterange(a, b):
    x = a
    while x <= b:
        yield x
        x += dt.timedelta(days=1)

def gantt_sheet(wb, name, title, sub, tasks, rng, zentei, style, changes=None, revs=None):
    ws = wb.create_sheet(name)
    days = list(daterange(*rng))
    m2 = style == 'm2'
    c0 = 4 if m2 else 3            # 日付の開始列（m2 は「班」列を追加）
    cend = c0 + len(days) - 1
    crem = cend + 1                # 備考列
    fnt = 'ＭＳ Ｐゴシック' if not m2 else 'ＭＳ ゴシック'
    f9 = F(9, name=fnt); f8 = F(8, name=fnt); f10b = F(11, True, name=fnt)
    widths(ws, {1: 5, 2: 46 if not m2 else 50})
    if m2:
        widths(ws, {3: 5})
    for i in range(len(days)):
        ws.column_dimensions[CL(c0 + i)].width = 3.0
    ws.column_dimensions[CL(crem)].width = 26
    # ---- 見出し
    if not m2:
        merge(ws, 1, 1, 1, cend - 14, title, F(14, True, name=fnt), AL, border=None)
        merge(ws, 2, 1, 2, cend - 14, sub, F(10, name=fnt), AL, border=None)
        merge(ws, 1, cend - 12, 1, cend - 10, '作成', f9, AC)
        merge(ws, 1, cend - 9, 1, cend - 5, META['made'], f9, AC)
        merge(ws, 2, cend - 12, 2, cend - 10, '受注者', f9, AC)
        merge(ws, 2, cend - 9, 2, crem, '汐見重工環境（株）　現場代理人　' + META['agent'], f9, AL)
        for i, s in enumerate(['担当', '係長', '課長']):
            merge(ws, 1, crem - 0 if False else cend - 4 + i * 0, 1, cend - 4 + i * 0, None, border=None)
        merge(ws, 1, cend - 4, 1, crem, '確認（あさぎ市）：　　　　　　　　　', f8, AL)
        r = 4
    else:
        merge(ws, 1, 1, 1, cend - 14, title, F(14, True, name=fnt), AL, border=None)
        # m2 で作った表形式の見出し（m3 も継続使用）
        put(ws, 2, 1, '工期', f9, AC, B_ALL); merge(ws, 2, 2, 2, 2, sub[0], f9, AL)
        put(ws, 3, 1, '施工', f9, AC, B_ALL); merge(ws, 3, 2, 3, 2, sub[1], f9, AL)
        merge(ws, 2, 3, 2, 12, '作成', f9, AC); merge(ws, 2, 13, 2, 26, sub[2], f9, AL)
        merge(ws, 3, 3, 3, 12, '改訂', f9, AC); merge(ws, 3, 13, 3, crem, sub[3], F(8, name=fnt, color='C00000'), AL)
        merge(ws, 2, 27, 2, cend - 6, '受注者　汐見重工環境（株）環境プラント事業部　現場代理人 ' + META['agent3'], f9, AL)
        merge(ws, 2, cend - 5, 2, crem, '発注者確認　あさぎ市 環境部 清掃工場', f8, AL)
        r = 5
    # ---- 日付ヘッダ（月・日・曜日）
    hr = r
    put(ws, hr, 1, 'No', f9, AC, B_ALL); ws.merge_cells(start_row=hr, start_column=1, end_row=hr + 2, end_column=1)
    put(ws, hr, 2, '作　業　項　目', f9, AC, B_ALL); ws.merge_cells(start_row=hr, start_column=2, end_row=hr + 2, end_column=2)
    if m2:
        put(ws, hr, 3, '班', f9, AC, B_ALL); ws.merge_cells(start_row=hr, start_column=3, end_row=hr + 2, end_column=3)
    put(ws, hr, crem, '備　考', f9, AC, B_ALL); ws.merge_cells(start_row=hr, start_column=crem, end_row=hr + 2, end_column=crem)
    grid(ws, hr, 1, hr + 2, crem)
    i = 0
    while i < len(days):
        j = i
        while j + 1 < len(days) and days[j + 1].month == days[i].month:
            j += 1
        merge(ws, hr, c0 + i, hr, c0 + j, '%d月' % days[i].month, f9, AC)
        i = j + 1
    for i, x in enumerate(days):
        sun = x.weekday() == 6
        sat = x.weekday() == 5
        put(ws, hr + 1, c0 + i, x.day, F(8, name=fnt, color='C00000' if sun else ('0000C0' if sat else '000000')), AC, B_ALL)
        put(ws, hr + 2, c0 + i, WD[x.weekday()], F(8, name=fnt, color='C00000' if sun else ('0000C0' if sat else '000000')), AC, B_ALL)
    r = hr + 3
    zfill = fill('FFF2CC')
    sunfill = fill('FDE9E7') if not m2 else fill('EDEDED')
    bar = fill('7F7F7F') if not m2 else fill('4F81BD')
    sepbar = pfill('darkUp', '595959') if not m2 else pfill('lightDown', 'C0504D')
    chg = pfill('darkUp', 'C00000')
    groups = []
    for t in tasks:
        if not groups or groups[-1][0] != t['grp']:
            groups.append((t['grp'], []))
        groups[-1][1].append(t)
    gname = {'1号炉': '■ 1号炉', '2号炉': '■ 2号炉', '共通': '■ 共通設備', '全停': '■ 全停（全炉停止）', '別契約': '■ 別契約（参考：本委託の範囲外）'}
    first = True
    for g, ts in groups:
        if m2 and g == '共通' and not first:
            r += 2   # m2 がコピー編集で残した空き 2 行（m3 もそのまま使用）
        first = False
        _hk(ws.title, r, 'parent', {'group': gname.get(g, '■ ' + g)})
        merge(ws, r, 1, r, crem, gname.get(g, '■ ' + g), F(10, True, name=fnt), AL, border=box(t=thin, b=thin, l=thin, r=thin), fl=fill('DDEBF7') if m2 else fill('E7E6E6'))
        r += 1
        for t in ts:
            fy_ = days[0].year if days[0].month >= 4 else days[0].year - 1
            TR.r(name, r, 'schedule_task:' + t['no'], [i for g in t['tags'] for i in TR.work_ids(fy_, g)])
            put(ws, r, 1, t['no'], f8, AC, B_ALL)
            put(ws, r, 2, t['item'], f9 if len(t['item']) < 30 else (f8 if len(t['item']) < 40 else F(7, name=fnt)), Alignment(horizontal='left', vertical='center', shrink_to_fit=True), B_ALL)
            if m2:
                put(ws, r, 3, '' if t['team'] in ('別',) else ('運' if t['team'] == '運転' else t['team']), f9, AC, B_ALL)
            ch = (changes or {}).get(t['no'])
            for i, x in enumerate(days):
                c = c0 + i
                cell = ws.cell(row=r, column=c)
                cell.border = Border(left=hair, right=hair, top=thin, bottom=thin)
                if zentei[0] <= x <= zentei[1]:
                    cell.fill = zfill
                elif x.weekday() == 6:
                    cell.fill = sunfill
                if t['s'] <= x <= t['e']:
                    if t['kind'] == 'stop':
                        cell.value = '▼'; cell.font = F(9, True, name=fnt, color='C00000'); cell.alignment = AC
                    elif t['kind'] == 'start' and x == t['s']:
                        cell.value = '▲'; cell.font = F(9, True, name=fnt, color='0070C0'); cell.alignment = AC
                        if t['e'] > t['s']:
                            pass
                    elif t['kind'] == 'sep':
                        cell.fill = sepbar
                    else:
                        if x.weekday() == 6 and t['team'] not in ('運転',) and g != '全停':
                            cell.fill = pfill('gray125', '7F7F7F', 'FFFFFF') if not m2 else fill('B4C6E7')
                        else:
                            cell.fill = bar
                    if t['kind'] == 'start' and x > t['s']:
                        cell.fill = bar
                if ch and t['e'] < x <= ch['actual_e']:
                    cell.fill = chg
            rem = ''
            if t['kind'] == 'sep':
                rem = '別途契約（工程調整のみ）'
            if t['team'] == '運転':
                rem = META['op'] if t['kind'] in ('stop', 'start') else ''
            if ch:
                rem = ch['note']
                ws.cell(row=r, column=c0 + (ch['actual_e'] - days[0]).days).comment = Comment('実績追記 m3：' + ch['note'], 'm3')
            put(ws, r, crem, rem, F(7, name=fnt), AL, B_ALL)
            _hk(ws.title, r, 'item', dict(no=t['no'], item=t['item'], team=(None if not m2 or t['team'] == '別' else ('運' if t['team'] == '運転' else t['team'])), start=t['s'].isoformat(), end=t['e'].isoformat(), remark=rem or None), t)
            ws.row_dimensions[r].height = 15
            r += 1
    # ---- 班別人員（m2 で追加）
    if m2:
        r += 1
        merge(ws, r, 1, r, crem, '■ 班別人員（予定・人）　A班：炉・築炉　B班：ボイラ・排ガス　C班：回転機・機械　D班：電気計装', F(10, True, name=fnt), AL, fl=fill('DDEBF7'))
        r += 1
        tot = [0] * len(days)
        for team in 'ABCD':
            put(ws, r, 2, '%s班' % team, f9, AR, B_ALL); put(ws, r, 3, team, f9, AC, B_ALL)
            for i, x in enumerate(days):
                n = sum(sum(t['n']) for t in tasks if t['team'] == team and x in work_days(t))
                tot[i] += n
                put(ws, r, c0 + i, n if n else None, F(7, name=fnt), AC, Border(left=hair, right=hair, top=thin, bottom=thin))
            put(ws, r, crem, None, f8, AL, B_ALL)
            r += 1
        _hk(ws.title, r, 'total', {'label': '合計', 'persons': {str(days[i]): tot[i] for i in range(len(days)) if tot[i]}})
        put(ws, r, 2, '合　計（運転委託先・別契約を除く）', f9, AR, B_ALL); put(ws, r, 3, '', f9, AC, B_ALL)
        for i in range(len(days)):
            put(ws, r, c0 + i, tot[i] if tot[i] else None, F(7, True, name=fnt), AC, Border(left=hair, right=hair, top=med, bottom=med))
        put(ws, r, crem, None, f8, AL, B_ALL)
        r += 1
    # ---- 凡例・注記
    r += 1
    notes = ['凡例：' + ('■ 施工（灰色）　▼ 焼却停止　▲ 立上げ　斜線 別契約　黄色 全停期間' if not m2 else '■ 施工（青）　▼ 焼却停止　▲ 立上げ　斜線（赤） 別契約　赤網 実績による延長　黄色 全停期間'),
             '※日曜日は休工とする。ただし冷却・全停・立上げ操作期間は除く。',
             '※焼却停止・立上げ・試運転時の運転操作は%sが行い、受注者は立会・調整を行う。' % META['op'].split('（')[0],
             '※受変電設備の年次点検（停電）は電気保安業務委託により別途実施。本工程では停電作業の範囲のみ示す。']
    if m2:
        notes.append('※別契約の補修工事は工程調整のため参考表示。作業区域・火気使用は日々の朝礼で調整する。')
    for n_ in notes:
        merge(ws, r, 1, r, crem, n_, F(8, name=fnt), AL, border=None)
        r += 1
    page(ws, landscape=True, paper=8, fith=1, area='A1:%s%d' % (CL(crem), r))
    ws.freeze_panes = "%s%d" % (CL(c0), hr + 3)
    return ws

def build_h23():
    TR.set_file('【架空】あさぎ市清掃工場_工程表_平成23年度定期整備業務委託_m1様式.xlsx')
    wb = Workbook(); wb.remove(wb.active)
    c = H23
    gantt_sheet(wb, '1号炉', '平成23年度　あさぎ市清掃工場　定期整備業務委託　工程表（１号炉・全停）',
                '工期：平成23年6月1日〜平成24年1月31日　　施工期間：平成23年7月4日〜8月5日（１号炉）　全停：7月20日〜22日',
                H23_F1, (D(2011, 7, 1), D(2011, 8, 8)), c['zentei'], 'm1')
    gantt_sheet(wb, '2号炉', '平成23年度　あさぎ市清掃工場　定期整備業務委託　工程表（２号炉）',
                '工期：平成23年6月1日〜平成24年1月31日　　施工期間：平成23年10月17日〜11月18日（２号炉）',
                H23_F2, (D(2011, 10, 14), D(2011, 11, 21)), (D(2000, 1, 1), D(2000, 1, 1)), 'm1')
    p = OUT + '/【架空】あさぎ市清掃工場_工程表_平成23年度定期整備業務委託_m1様式.xlsx'
    wb.save(p)
    return p

def build_r7():
    TR.set_file('【架空】あさぎ市清掃工場_工程表_令和7年度定期整備業務委託_m2様式.xlsx')
    wb = Workbook(); wb.remove(wb.active)
    c = R7
    # 年間総括（m3 が追加）
    ws = wb.create_sheet('年間総括')
    fnt = 'ＭＳ ゴシック'
    merge(ws, 1, 1, 1, 26, '令和7年度　あさぎ市清掃工場　定期整備業務委託　年間工程（総括）', F(14, True, name=fnt), AL, border=None)
    merge(ws, 2, 1, 2, 26, '工期：令和7年5月12日〜令和8年1月30日　　作成 R7.5.16 m3　改1 R7.6.19', F(9, name=fnt), AL, border=None)
    months = [(2025, m) for m in range(4, 13)] + [(2026, m) for m in range(1, 4)]
    widths(ws, {1: 5, 2: 44})
    put(ws, 4, 1, 'No', F(9, name=fnt), AC, B_ALL); put(ws, 4, 2, '区分・件名', F(9, name=fnt), AC, B_ALL)
    for i, (y, m) in enumerate(months):
        merge(ws, 4, 3 + i * 2, 4, 4 + i * 2, '%d月' % m, F(9, name=fnt), AC)
        ws.column_dimensions[CL(3 + i * 2)].width = 5; ws.column_dimensions[CL(4 + i * 2)].width = 5
    put(ws, 4, 27, '備考', F(9, name=fnt), AC, B_ALL); ws.column_dimensions['AA'].width = 30
    rows = [
        ('1', '1号炉 定期整備（焼却停止〜試運転）', D(2025, 6, 2), D(2025, 7, 4), '4F81BD', '6/2停止 7/1立上げ'),
        ('2', '全停（受変電停電・タービン・共通蒸気系）', D(2025, 6, 17), D(2025, 6, 19), 'FFC000', '2号炉 6/16停止 6/20立上げ'),
        ('3', '2号炉 定期整備（焼却停止〜試運転）', D(2025, 9, 29), D(2025, 10, 31), '4F81BD', '9/29停止 10/28立上げ'),
        ('別', '【別契約】1号炉ボイラ蒸発管補修工事', D(2025, 6, 12), D(2025, 6, 20), 'C0504D', '指名競争 R7.5.23契約'),
        ('別', '【別契約】2号炉押込送風機軸受補修工事', D(2025, 10, 6), D(2025, 10, 10), 'C0504D', '指名競争 R7.8.8契約'),
        ('別', '【別契約】2号炉後燃焼火格子補修工事', D(2025, 10, 6), D(2025, 10, 15), 'C0504D', '指名競争 R7.7.25契約'),
        ('別', '【別契約】灰ピット壁面補修工事', D(2025, 10, 14), D(2025, 10, 24), 'C0504D', '指名競争 R7.8.22契約（あさぎ建設）'),
        ('他', '（参考）電気保安 年次点検（停電）', D(2025, 6, 17), D(2025, 6, 18), 'A6A6A6', '電気保安業務委託'),
        ('他', '（参考）ごみクレーン年次自主検査', D(2025, 11, 10), D(2025, 11, 14), 'A6A6A6', '別途点検業務委託'),
    ]
    for k, (no, nm, s, e_, col, rem) in enumerate(rows):
        r = 5 + k
        put(ws, r, 1, no, F(9, name=fnt), AC, B_ALL); put(ws, r, 2, nm, F(9, name=fnt), AL, B_ALL)
        for i, (y, m) in enumerate(months):
            for half in (0, 1):
                cs = D(y, m, 1 if half == 0 else 16)
                ce = (D(y, m, 15) if half == 0 else (D(y + (m == 12), m % 12 + 1, 1) - dt.timedelta(days=1)))
                cell = put(ws, r, 3 + i * 2 + half, None, None, None, Border(left=hair if half else thin, right=thin if half else hair, top=thin, bottom=thin))
                if s <= ce and e_ >= cs:
                    cell.fill = fill(col)
        put(ws, r, 27, rem, F(8, name=fnt), AL, B_ALL)
        ws.row_dimensions[r].height = 18
    merge(ws, 15, 1, 15, 27, '※半月単位の表示。詳細は各号炉の工程表による。', F(8, name=fnt), AL, border=None)
    page(ws, landscape=True, paper=9, area='A1:AA15')
    gantt_sheet(wb, '1号炉・全停', '令和7年度　あさぎ市清掃工場　定期整備業務委託　工程表（１号炉・全停）',
                ['令和7年5月12日〜令和8年1月30日', '1号炉 令和7年6月2日〜7月4日　全停 6月17日〜19日', 'R7.5.16　m3', '改1 R7.6.19 BFろ布交換：リテーナ腐食によりリテーナ交換を追加、完了6/24→6/26（復旧・立上げは予備日で吸収）'],
                R7_F1, (D(2025, 5, 30), D(2025, 7, 7)), c['zentei'], 'm2', changes=R7_CHANGES)
    gantt_sheet(wb, '2号炉', '令和7年度　あさぎ市清掃工場　定期整備業務委託　工程表（２号炉）',
                ['令和7年5月12日〜令和8年1月30日', '2号炉 令和7年9月29日〜10月31日', 'R7.9.12　m3', '—'],
                R7_F2, (D(2025, 9, 26), D(2025, 11, 3)), (D(2000, 1, 1), D(2000, 1, 1)), 'm2')
    p = OUT + '/【架空】あさぎ市清掃工場_工程表_令和7年度定期整備業務委託_m2様式.xlsx'
    wb.save(p)
    return p

if __name__ == '__main__':
    p = build_h23(); print(render(p, pages=(1, 2)))
    p = build_r7(); print(render(p, pages=(1, 2, 3)))
