# -*- coding: utf-8 -*-
"""工程表・作業日報（定修・補修工事・改造工事）。工程表を先に作り、日報はその工程から作る。"""
import os, datetime as dt, glob, random
from openpyxl import Workbook
from openpyxl.comments import Comment
from xlutil import *
import make_schedule as MS, make_daily as MD_, ans, photo
import world2 as W, daily2 as DY
import f01_truth as TR
from f01_master import wareki, EQ
from f01_history import BYTAG
from f01_contracts import H23, R7, R7_CHANGES, MAKER
from daily_data import H23_DAYS, R7_DAYS

FILES = ans.FILES
HANDDIR = TMP + '_hand'
KAKU = '【架空】'
WDJ = '月火水木金土日'
PENDING_PHOTOS = []   # (xlsx, [(sheet, stem_out, records)], person)

def jshort(x):
    return ('H%d.%d.%d' % (x.year - 1988, x.month, x.day)) if (x.year, x.month) < (2019, 5) else ('R%d.%d.%d' % (x.year - 2018, x.month, x.day))

def wid(fy, tag, idx):
    d = BYTAG[tag]
    na = len(d['act'].get(fy, []))
    if idx < na:
        return 'W%d-%s-%d' % (fy, d['code'], idx + 1)
    return 'W%d-%s-p%d' % (fy, d['code'], idx - na + 1)

def task_works(fy, t):
    if t.get('evs'):
        return [wid(fy, tag, i) for (tag, i) in t['evs']]
    return [w for g in t['tags'] for w in TR.work_ids(fy, g)]

def fyname(fy):
    return wareki(fy, False)

def hook_for(stem, fy, photo_sheets=None):
    def hk(sheet, r, kind, fields, t=None):
        works = task_works(fy, t) if t else []
        if photo_sheets is not None:
            photo_sheets.setdefault(sheet, []).append((kind, fields, works))
        else:
            ans.rec(stem, kind, fields, sheet=sheet, r=r, works=works)
    return hk

# ------------------------------------------------------------------ 工程表（定修）
def sched_teishu(fy):
    P = W.teishu(fy)
    st = W.sched_style(fy)
    wins = P['wins']
    title = '%s　あさぎ市清掃工場　定期整備業務委託' % fyname(fy)
    MS.META.update(made=jshort(W.teishu_period(fy)[0] + dt.timedelta(days=4)), agent=W.agent(fy), agent3=W.agent(fy),
                   op='%s（%s）' % (W.operator_word(fy), 'あさぎ環境運転サービス' if fy >= 2006 else '清掃工場'))
    name = '%sあさぎ市清掃工場_工程表_%s定期整備業務委託_%s様式' % (KAKU, fyname(fy), {'hand': '手書き', 'm1': 'm1', 'm2': 'm2'}[st])
    wb = Workbook(); wb.remove(wb.active)
    psheets = {} if st == 'hand' else None
    stem = None if st == 'hand' else ans.start(name + '.xlsx')
    MS.HOOK[0] = hook_for(stem, fy, psheets)
    TR.set_file(name + '.xlsx')
    for key, f in (('f1', 1), ('f2', 2)):
        ts = P[key]
        s = min(t['s'] for t in ts) - dt.timedelta(days=2)
        e = max(t['e'] for t in ts if t['kind'] != 'start' or True)
        e = min(e, s + dt.timedelta(days=52)) + dt.timedelta(days=2)
        z = wins['Z'] if any(t['grp'] == '全停' for t in ts) else (D(2000, 1, 1), D(2000, 1, 1))
        lab = '%d号炉' % f + ('・全停' if z[0].year > 2000 else '')
        sub = '工期：%s〜%s　　施工期間：%s〜%s（%d号炉）' % (jshort(W.teishu_period(fy)[1]), jshort(W.teishu_period(fy)[2]),
                                                     jshort(wins[f][0]), jshort(wins[f][1]), f)
        if st == 'm2':
            sub = [sub.split('　　')[0].replace('工期：', ''), '%d号炉 %s〜%s' % (f, jshort(wins[f][0]), jshort(wins[f][1])), '%s　%s' % (MS.META['made'], W.agent(fy)),
                   ('改1 ' + R7_CHANGES['1-18']['note']) if (fy == 2025 and f == 1) else '—']
        MS.gantt_sheet(wb, lab, KAKU + '%s　工程表（%s）' % (title, lab), sub, ts, (s, e), z, 'm2' if st == 'm2' else 'm1',
                       changes=R7_CHANGES if fy == 2025 else None)
    return finish(wb, name, st, psheets, W.agent(fy))

def finish(wb, name, st, psheets, person):
    if st == 'hand':
        os.makedirs(HANDDIR, exist_ok=True)
        p = HANDDIR + '/' + name + '.xlsx'
        wb.save(p)
        PENDING_PHOTOS.append((p, name, [(ws.title, psheets.get(ws.title, [])) for ws in wb.worksheets], person))
        return p
    p = FILES + '/' + name + '.xlsx'
    wb.save(p)
    return p

D = dt.date

# ------------------------------------------------------------------ 日報（m1 様式 2004〜2010・手書き 1997〜2003 も同じ書式）
TRADES = ['社員', '機械', '溶接', '築炉', '電工', '足場', '雑']
def split_trades(team, s, k, key):
    v = [s, 0, 0, 0, 0, 0, 0]
    order = {'A': [3, 5, 6], 'B': [2, 1], 'C': [1, 2], 'D': [4], '運': [6]}.get(team, [6])
    r = random.Random(str(key))
    for i in range(k):
        v[order[i % len(order) if r.random() < 0.7 else r.randrange(len(order))]] += 1
    return v

def m1d_sheet(wb, x, dd, meta, hk):
    ws = wb.create_sheet('%d.%d' % (x.month, x.day) if meta['style'] == 'm1d' else '%02d%02d' % (x.month, x.day))
    fn = 'ＭＳ 明朝' if meta['style'] == 'hand' else 'ＭＳ Ｐゴシック'
    widths(ws, {'A': 4, 'B': 7, 'C': 22, 'D': 44, 'E': 4.5, 'F': 4.5, 'G': 4.5, 'H': 4.5, 'I': 4.5, 'J': 4.5, 'K': 4.5, 'L': 5, 'M': 10, 'N': 14})
    merge(ws, 1, 1, 1, 11, KAKU + '作　業　日　報', F(16, True, name=fn), AC, border=None)
    put(ws, 1, 12, 'No.', F(9, name=fn), AR); merge(ws, 1, 13, 1, 14, dd['no'], F(11, name=fn), AC, border=box(b=thin, t=NONE, l=NONE, r=NONE))
    put(ws, 2, 1, '工事名', F(9, name=fn), AC, B_ALL); merge(ws, 2, 2, 2, 9, meta['title'], F(10, name=fn), AL)
    merge(ws, 2, 10, 2, 11, '日付', F(9, name=fn), AC); merge(ws, 2, 12, 2, 14, '%s（%s）' % (jshort(x), WDJ[x.weekday()]), F(10, name=fn), AC)
    put(ws, 3, 1, '天候', F(9, name=fn), AC, B_ALL); merge(ws, 3, 2, 3, 3, '%s　%s℃' % (dd['weather'], dd['temp']), F(10, name=fn), AL)
    put(ws, 3, 4, '作業時間 ' + dd['time'], F(9, name=fn), AL, B_ALL)
    merge(ws, 3, 5, 3, 11, '責任者', F(9, name=fn), AC); merge(ws, 3, 12, 3, 14, meta['agent'], F(10, name=fn), AC)
    r = 5
    hd = ['No', '区分', '作業項目（箇所）', '作業内容'] + TRADES + ['計', '工数', '備考']
    for i, hh in enumerate(hd):
        put(ws, r, i + 1, hh, F(8, True, name=fn), AC, B_ALL, fill('E7E6E6'))
    ws.row_dimensions[r].height = 22
    r += 1
    tot = [0] * 7
    tmh = 0
    for (no, team, f, place, text, s, k, hrs, prog) in dd['rows']:
        v = split_trades(team, s, k, (no, x.isoformat()))
        put(ws, r, 1, no, F(8, name=fn), AC, B_ALL); put(ws, r, 2, f, F(8, name=fn), AC, B_ALL)
        put(ws, r, 3, place, F(8, name=fn), AL, B_ALL); put(ws, r, 4, text, F(8, name=fn), AL, B_ALL)
        for i, n in enumerate(v):
            put(ws, r, 5 + i, n or None, F(8, name=fn), AC, B_ALL)
            tot[i] += n
        put(ws, r, 12, s + k, F(8, name=fn), AC, B_ALL)
        mtxt = '%d人×%gh' % (s + k, hrs) if (s + k) else '—'
        put(ws, r, 13, mtxt, F(8, name=fn), AC, B_ALL)
        put(ws, r, 14, '完了' if prog >= 100 else '%d%%' % prog, F(8, name=fn), AC, B_ALL)
        tmh += (s + k) * hrs
        hk(ws.title, r, 'item', dict(date=x.isoformat(), no=no, furnace=f, place=place, content=text, persons=s + k,
                                     trades={TRADES[i]: n for i, n in enumerate(v) if n}, manDays=round((s + k) * hrs / 8.0, 2), manDaysText=mtxt,
                                     progress='完了' if prog >= 100 else '%d%%' % prog), no)
        ws.row_dimensions[r].height = 12 * max(1, (len(text) + 27) // 28) + 3
        r += 1
    while r < 6 + 14:
        grid(ws, r, 1, r, 14)
        r += 1
    merge(ws, r, 1, r, 4, '人　員　計', F(9, True, name=fn), AC)
    for i, n in enumerate(tot):
        put(ws, r, 5 + i, n, F(9, True, name=fn), AC, box(t=dbl, b=thin, l=thin, r=thin))
    put(ws, r, 12, sum(tot), F(9, True, name=fn), AC, box(t=dbl, b=thin, l=thin, r=thin))
    put(ws, r, 13, '%g人工' % round(tmh / 8.0, 1), F(9, True, name=fn), AC, box(t=dbl, b=thin, l=thin, r=thin))
    put(ws, r, 14, None, None, None, box(t=dbl, b=thin, l=thin, r=thin))
    hk(ws.title, r, 'total', dict(date=x.isoformat(), persons=sum(tot), manDays=round(tmh / 8.0, 2), manDaysText='%g人工' % round(tmh / 8.0, 1)), None)
    r += 2
    for lab, v in (('安全', dd['ky']), ('記事', (dd.get('note') or '') + ('　立会：' + dd['witness'] if dd.get('witness') not in (None, 'なし') else ''))):
        put(ws, r, 1, lab, F(9, name=fn), AC, B_ALL); merge(ws, r, 2, r, 11, v, F(9, name=fn), AL)
        r += 1
    merge(ws, r - 2, 12, r - 2, 14, '市確認（%s）' % meta['city'], F(8, name=fn), AC)
    merge(ws, r - 1, 12, r - 1, 14, '', None, AC)
    merge(ws, r + 1, 1, r + 1, 14, '汐見重工環境株式会社　あさぎ市清掃工場作業所', F(8, name=fn), AL, border=None)
    page(ws, landscape=True, paper=9, fith=1, area='A1:N%d' % (r + 1))
    return ws

def daily_book(fy, tasks, title, name, style, sample=None, upto=None, numbering=1, changes=None):
    """日報のブック（1 日 1 シート）。style: hand/m1d/m2/m3"""
    days = DY.report_days(tasks, upto)
    wb = Workbook(); wb.remove(wb.active)
    psheets = {} if style == 'hand' else None
    stem = None if style == 'hand' else ans.start(name + '.xlsx')
    TR.set_file(name + '.xlsx')
    meta = dict(title=title, agent=W.agent(fy), city=W.city_person(fy), style=style)
    hk = hook_for(stem, fy, psheets)
    def hk2(sheet, r, kind, fields, no):
        t = next((t for t in tasks if t['no'] == no), None) if no else None
        hk(sheet, r, kind, fields, t)
    for i, x in enumerate(days):
        dd = dict(DY.dd_for(fy, tasks, x, sample))
        dd['no'] = numbering + i
        if not dd['rows']:
            continue
        if style in ('hand', 'm1d'):
            m1d_sheet(wb, x, dd, meta, hk2)
        elif style == 'm2':
            MD_.MD.update(title=title, agent=meta['agent'], city=meta['city'], fy=fy)
            MD_.HOOK[0] = hk2
            MD_.m2_sheet(wb, x, dd)
        else:
            MD_.MD.update(title=title, agent=meta['agent'], city=meta['city'], fy=fy, changes=changes)
            MD_.HOOK[0] = hk2
            MD_.m3_sheet(wb, x, dd, tasks)
    if not wb.worksheets:
        ans.drop(stem) if stem else None
        return None
    return finish(wb, name, 'hand' if style == 'hand' else style, psheets, meta['agent'])

def daily_teishu(fy):
    P = W.teishu(fy)
    style = W.daily_style(fy)
    sample = H23_DAYS if fy == 2011 else (R7_DAYS if fy == 2025 else None)
    tasks = P['f1'] + P['f2']
    title = '%s あさぎ市清掃工場 定期整備業務委託' % fyname(fy)
    name = '%sあさぎ市清掃工場_作業日報_%s定期整備業務委託_%s様式' % (KAKU, fyname(fy), {'hand': '手書き', 'm1d': 'm1', 'm2': 'm2', 'm3': 'm3'}[style])
    upto = D(2026, 11, 12) if fy == 2026 else None
    return daily_book(fy, tasks, title, name, style, sample, upto, changes=R7_CHANGES if fy == 2025 else None)

# ------------------------------------------------------------------ 週単位の工程表（改造工事・長い工事）と業者様式
def gantt_week(wb, sheet, title, sub, tasks, s, e, fy, stem, psheets, vendor_style=False):
    ws = wb.create_sheet(sheet)
    fn = 'ＭＳ 明朝' if vendor_style else 'ＭＳ Ｐゴシック'
    wk0 = s - dt.timedelta(days=s.weekday())
    weeks = []
    x = wk0
    while x <= e:
        weeks.append(x)
        x += dt.timedelta(days=7)
    widths(ws, {'A': 5, 'B': 54})
    for i in range(len(weeks)):
        ws.column_dimensions[CL(3 + i)].width = 3.6
    cr = 3 + len(weeks)
    ws.column_dimensions[CL(cr)].width = 18
    merge(ws, 1, 1, 1, min(cr, 30), KAKU + title, F(14, True, name=fn), AL if not vendor_style else AC, border=None)
    merge(ws, 2, 1, 2, min(cr, 30), sub, F(9, name=fn), AL, border=None)
    hr = 4
    put(ws, hr, 1, 'No', F(8, name=fn), AC, B_ALL); put(ws, hr, 2, '工　種' if vendor_style else '作業項目', F(8, name=fn), AC, B_ALL)
    i = 0
    while i < len(weeks):
        j = i
        while j + 1 < len(weeks) and weeks[j + 1].month == weeks[i].month:
            j += 1
        merge(ws, hr, 3 + i, hr, 3 + j, '%d年%d月' % (weeks[i].year, weeks[i].month) if vendor_style else '%d月' % weeks[i].month, F(8, name=fn), AC)
        i = j + 1
    put(ws, hr, cr, '備考', F(8, name=fn), AC, B_ALL)
    put(ws, hr + 1, 1, None, None, None, B_ALL); put(ws, hr + 1, 2, '（週の初日）', F(7, name=fn), AC, B_ALL)
    for i, wk in enumerate(weeks):
        put(ws, hr + 1, 3 + i, wk.day, F(7, name=fn), AC, B_ALL)
    put(ws, hr + 1, cr, None, None, None, B_ALL)
    r = hr + 2
    barf = fill('404040') if vendor_style else fill('7F7F7F')
    for t in tasks:
        put(ws, r, 1, t['no'], F(8, name=fn), AC, B_ALL)
        put(ws, r, 2, t['item'], F(8, name=fn), Alignment(horizontal='left', vertical='center', shrink_to_fit=True), B_ALL)
        for i, wk in enumerate(weeks):
            c = put(ws, r, 3 + i, None, None, None, Border(left=hair, right=hair, top=thin, bottom=thin))
            if t['s'] <= wk + dt.timedelta(days=6) and t['e'] >= wk:
                c.fill = barf
        put(ws, r, cr, None, None, None, B_ALL)
        fields = dict(no=t['no'], item=t['item'], start=t['s'].isoformat(), end=t['e'].isoformat())
        if psheets is not None:
            psheets.setdefault(sheet, []).append(('item', fields, task_works(fy, t)))
        else:
            ans.rec(stem, 'item', fields, sheet=sheet, r=r, works=task_works(fy, t))
        r += 1
    merge(ws, r + 1, 1, r + 1, min(cr, 30), '※日曜日・祝日は原則休工。工程は進捗により変更することがある。', F(8, name=fn), AL, border=None)
    page(ws, landscape=True, paper=9 if len(weeks) < 30 else 8, fith=1, area='A1:%s%d' % (CL(cr), r + 1))
    return ws

def vendor_daily_sheet(wb, x, dd, meta, hk):
    """業者（築炉・建設・防災）の工事日報"""
    ws = wb.create_sheet('%02d%02d' % (x.month, x.day))
    fn = 'ＭＳ 明朝'
    widths(ws, {'A': 12, 'B': 50, 'C': 8, 'D': 8, 'E': 8, 'F': 12})
    merge(ws, 1, 1, 1, 6, KAKU + '工　事　日　報', F(16, True, name=fn), AC, border=None)
    merge(ws, 2, 1, 2, 6, meta['vendor'], F(9, name=fn), Alignment(horizontal='right'), border=None)
    rows = [('工事名', meta['title']), ('日　付', '%s年%d月%d日（%s）　天候 %s' % (('平成%d' % (x.year - 1988)) if (x.year, x.month) < (2019, 5) else ('令和%d' % (x.year - 2018)), x.month, x.day, WDJ[x.weekday()], dd['weather']))]
    r = 3
    for k, v in rows:
        put(ws, r, 1, k, F(10, name=fn), AC, B_ALL); merge(ws, r, 2, r, 6, v, F(10, name=fn), AL)
        r += 1
    r += 1
    for i, hh in enumerate(['工　種', '作業内容', '職長', '作業員', '計', '備考']):
        put(ws, r, 1 + i, hh, F(9, True, name=fn), AC, B_ALL)
    r += 1
    tp = 0; tmh = 0
    for (no, team, f, place, text, s, k, hrs, prog) in dd['rows']:
        put(ws, r, 1, place, F(9, name=fn), AL, B_ALL); put(ws, r, 2, text, F(9, name=fn), AL, B_ALL)
        put(ws, r, 3, s, F(9, name=fn), AC, B_ALL); put(ws, r, 4, k, F(9, name=fn), AC, B_ALL); put(ws, r, 5, s + k, F(9, name=fn), AC, B_ALL)
        put(ws, r, 6, '%g時間' % hrs, F(9, name=fn), AC, B_ALL)
        hk(ws.title, r, 'item', dict(date=x.isoformat(), place=place, content=text, persons=s + k, manDays=round((s + k) * hrs / 8.0, 2),
                                     manDaysText='%g時間' % hrs), no)
        tp += s + k; tmh += (s + k) * hrs
        r += 1
    put(ws, r, 1, '出面計', F(9, True, name=fn), AC, B_ALL); merge(ws, r, 2, r, 4, None); put(ws, r, 5, tp, F(9, True, name=fn), AC, B_ALL); put(ws, r, 6, None, None, None, B_ALL)
    hk(ws.title, r, 'total', dict(date=x.isoformat(), persons=tp, manDays=round(tmh / 8.0, 2)), None)
    r += 2
    put(ws, r, 1, '安全', F(9, name=fn), AC, B_ALL); merge(ws, r, 2, r, 6, dd['ky'], F(9, name=fn), AL)
    put(ws, r + 1, 1, '使用機械', F(9, name=fn), AC, B_ALL); merge(ws, r + 1, 2, r + 1, 6, random.Random(str(x)).choice(['電動ピック、ミキサ', 'ハンドブレーカ', '溶接機 2台', 'ミキサ、高圧洗浄機']), F(9, name=fn), AL)
    page(ws, landscape=False, paper=9, fith=1, area='A1:F%d' % (r + 1))
    return ws

def contract_docs(fy, c, L, kind):
    """補修工事・改造工事の工程表と日報"""
    vendor = c['vendor']
    vstyle = vendor != MAKER
    hand = fy <= 2003
    short = c['title'].split('清掃工場')[-1].replace(' ', '')
    tag = '業者様式' if vstyle else None
    sst = 'hand' if hand else ('v' if vstyle else W.sched_style(fy))
    name = '%sあさぎ市清掃工場_工程表_%s%s_%s' % (KAKU, fyname(fy), short, '手書き' if hand else ('業者様式' if vstyle else W.sched_style(fy) + '様式'))
    wb = Workbook(); wb.remove(wb.active)
    psheets = {} if hand else None
    stem = None if hand else ans.start(name + '.xlsx')
    TR.set_file(name + '.xlsx')
    s, e = min(t['s'] for t in L), max(t['e'] for t in L)
    long = (e - s).days > 40
    sub = '工期：%s〜%s　　請負者：%s　　%s' % (jshort(c['s']), jshort(c['e_']), vendor, '現場代理人 ' + W.agent(fy) if not vstyle else '現場代理人 ' + {'港東築炉工業株式会社': 'k1', 'あさぎ建設株式会社': 'b1', 'あさぎ防災設備株式会社': 'f1'}[vendor])
    if long or vstyle:
        gantt_week(wb, '工程表', c['title'] + '　工程表', sub, L, s, e, fy, stem, psheets, vendor_style=vstyle)
    else:
        MS.META.update(made=jshort(c['s']), agent=W.agent(fy), agent3=W.agent(fy), op=W.operator_word(fy))
        MS.HOOK[0] = hook_for(stem, fy, psheets)
        MS.gantt_sheet(wb, '工程表', KAKU + c['title'] + '　工程表', sub if W.sched_style(fy) != 'm2' else [sub.split('　　')[0].replace('工期：', ''), '%s〜%s' % (jshort(s), jshort(e)), jshort(c['s']), '—'],
                       [t for t in L if t['team'] != '事務'] , (s - dt.timedelta(days=2), max(e + dt.timedelta(days=2), s + dt.timedelta(days=32))), (D(2000, 1, 1), D(2000, 1, 1)),
                       'm2' if W.sched_style(fy) == 'm2' else 'm1')
    finish(wb, name, 'hand' if hand else 'x', psheets, W.agent(fy))
    # 日報
    dst = 'hand' if hand else ('v' if vstyle else W.daily_style(fy))
    name = '%sあさぎ市清掃工場_作業日報_%s%s_%s' % (KAKU, fyname(fy), short, '手書き' if hand else ('業者様式' if vstyle else dst.replace('m1d', 'm1') + '様式'))
    if vstyle and not hand:
        days = DY.report_days(L)
        wb = Workbook(); wb.remove(wb.active)
        stem = ans.start(name + '.xlsx')
        TR.set_file(name + '.xlsx')
        meta = dict(title=c['title'], vendor=vendor)
        def hk2(sheet, r, kind_, fields, no):
            t = next((t for t in L if t['no'] == no), None) if no else None
            ans.rec(stem, kind_, fields, sheet=sheet, r=r, works=task_works(fy, t) if t else [])
        for x in days:
            dd = DY.dd_for(fy, L, x)
            if dd['rows']:
                vendor_daily_sheet(wb, x, dd, meta, hk2)
        wb.save(FILES + '/' + name + '.xlsx')
    else:
        daily_book(fy, L, c['title'], name, dst)

def process_photos():
    """手書き帳票のブックを PDF にして、1 頁 1 枚の写真にする"""
    if not PENDING_PHOTOS:
        return
    pdfdir = TMP + '_handpdf'
    os.makedirs(pdfdir, exist_ok=True)
    for f in glob.glob(pdfdir + '/*'):
        os.remove(f)
    photo.to_pdf([p for (p, *_r) in PENDING_PHOTOS], pdfdir)
    import subprocess
    for (p, name, sheets, person) in PENDING_PHOTOS:
        pdf = os.path.join(pdfdir, os.path.splitext(os.path.basename(p))[0] + '.pdf')
        info = subprocess.run(['pdfinfo', pdf], capture_output=True, text=True).stdout
        npg = int([l for l in info.splitlines() if l.startswith('Pages')][0].split()[1])
        assert npg == len(sheets), (p, npg, len(sheets))
        for i, (sheet, recs) in enumerate(sheets):
            rnd = random.Random(name + sheet)
            out = '%s_%s.jpg' % (name, sheet)
            Wd, Hd, words = photo.page_words(pdf, i + 1)
            img = photo.raster(pdf, i + 1, 1800)
            scale = img.width / Wd
            # 手書きの丸：数字・日付の欄を何か所か
            targets = [(a, b, c, d) for (a, b, c, d, w) in words if any(ch.isdigit() for ch in w) and len(w) <= 6]
            rnd.shuffle(targets)
            img = photo.handwrite(img, rnd, targets, scale, person)
            # しみ・ぼけで読めない欄（3 枚に 1 枚程度、作業内容の欄 1 つ）
            unread = None
            items = [k for k, (kind, f, w_) in enumerate(recs) if kind == 'item' and f.get('content')]
            if items and rnd.random() < 0.33:
                k = rnd.choice(items)
                key = recs[k][1]['content'][:4]
                hit = next(((a, b, c, d) for (a, b, c, d, w) in words if w.startswith(key)), None)
                if hit:
                    img = photo.stain(img, (hit[0] - 2, hit[1] - 2, min(Wd, hit[2] + 160), hit[3] + 2), scale, rnd)
                    unread = k
            img = photo.augment(img, rnd)
            size, nbytes = photo.save_jpeg(img, FILES + '/' + out, '【架空】試験用サンプル あさぎ市清掃工場 F01（手書き帳票のスキャン）')
            stem = ans.start(out)
            ans.A[stem]['image_size'] = list(size)
            for k, (kind, f, works) in enumerate(recs):
                f = dict(f)
                if k == unread:
                    f['content'] = '読めない'
                ans.rec(stem, kind, f, seq=k, image=out, works=works)
    PENDING_PHOTOS.clear()
