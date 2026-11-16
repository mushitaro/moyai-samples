# -*- coding: utf-8 -*-
"""工程表（千草）。様式：p1（H23〜H29：1 作業 1 行、予定＝灰色塗り、実績＝●、班ごと）→ p2 H30 様式（2 段、予定 ━・実績 緑塗り、節目＝黄色の日付）
→ p2 R3 様式（H30 を複写し、週・停止条件の帯（◆節目）・群見出し・別契約の行）。1 期間 1 シート。"""
import datetime as dt
from openpyxl import Workbook
from openpyxl.comments import Comment
import world
from xl import *
from rec import Rec

WD = '月火水木金土日'
GREEN = '2E7D32'; ACT = fill('2E7D32'); OTH = 'C55A11'; OTH_A = fill('ED7D31'); WEEKEND = fill('F2F2F2'); MS = fill('FFE699'); GRAYP = fill('BFBFBF')
BYTAG = {}

def days(a, b):
    return [a + dt.timedelta(i) for i in range((b - a).days + 1)]

def unitname(t):
    if not BYTAG: BYTAG.update({u['tag']: u for u in world.UNITS})
    return BYTAG[t]['name']

def fmt(fy):
    return 'p1' if fy <= 2017 else 'p2a' if fy <= 2020 else 'p2b'

def rng_txt(r):
    return '' if not r else f'{r[0].month}/{r[0].day}' + ('' if r[0] == r[1] else f'〜{r[1].month}/{r[1].day}')

def mandays(ppl, r):
    if not r or not ppl: return None
    return ppl * ((r[1] - r[0]).days + 1)

def unit_disp(it):
    n = unitname(it['units'][0]) if it['units'] else ''
    if len(it['units']) > 1 and it['tmpl'].startswith('ev'): n += f' ほか{len(it["units"]) - 1}'
    return n

def day_header(ws, r0, C0, dl, week_from=None):
    r = r0
    if week_from:
        for i, d in enumerate(dl):
            if i == 0 or d.weekday() == 0:
                j = i
                while j + 1 < len(dl) and dl[j + 1].weekday() != 0: j += 1
                merge(ws, r, C0 + i, r, C0 + j, f'第{(d - week_from).days // 7 + 1}週', font(7), AC, B, fill('EDEDED'))
        r += 1
    m0 = None
    for i, d in enumerate(dl):
        if d.month != m0:
            j = i
            while j + 1 < len(dl) and dl[j + 1].month == d.month: j += 1
            merge(ws, r, C0 + i, r, C0 + j, f'{d.month}月', font(8, True), AC, B, fill('DDEBF7')); m0 = d.month
        we = WEEKEND if d.weekday() >= 5 else None
        put(ws, r + 1, C0 + i, d.day, font(7, color='C00000' if d.weekday() == 6 else None), AC, B, we)
        put(ws, r + 2, C0 + i, WD[d.weekday()], font(7, color='C00000' if d.weekday() == 6 else None), AC, B, we)
    return r + 3

WK = {'on': False}
def bar(ws, r, C0, dl, rr, how, color=GREEN):
    if not rr: return
    for i, d in enumerate(dl):
        if (rr[0] <= d <= rr[1]) if not WK['on'] else (rr[0] <= d + dt.timedelta(6) and rr[1] >= d):
            x = ws.cell(r, C0 + i)
            if how == 'line': x.value = '━'; x.font = font(8, True, color); x.alignment = AC
            elif how == 'dot': x.value = '●'; x.font = font(7, color='C00000'); x.alignment = AC
            else: x.fill = how

def write(path, P, title, header_info, made, staff, other_blocks=None, extra_sheets=None, contract_name=None):
    """P：期間（1 シート）またはそのリスト。"""
    Ps = P if isinstance(P, list) else [P]
    wb = Workbook(); wb.remove(wb.active)
    props(wb, f'【架空】{title}', f'千草環境システム株式会社 {staff}')
    rec = Rec(path.split('/')[-1])
    for p in Ps:
        ws = wb.create_sheet(p.get('sheet', p['part'][:28]))
        sheet(ws, p, title, header_info, made, staff, rec, other_blocks)
    wb.save(path)
    return rec

def sheet(ws, p, title, header_info, made, staff, rec, other_blocks):
    f = p.get('fmt') or fmt(p['fy'])
    dl = days(p['start'], p['end'])
    items = [it for it in p['items']]
    if f == 'p1':
        C0 = 5
        widths(ws, [7, 5, 34, 5] + [2.6] * len(dl) + [20])
        last = C0 + len(dl)
        merge(ws, 1, 1, 1, last, f"【架空】{title}　{p['part']}　工程表", font(13, True), AL, b=None)
        merge(ws, 2, 1, 2, last, f"期間　{p['start']:%Y.%m.%d} 〜 {p['end']:%Y.%m.%d}　　{header_info}　　作成 {made}　千草環境システム㈱ SV {staff}", font(9), AL, b=None)
        put(ws, 3, 1, '凡例：灰色＝予定　●＝実施', font(8), AL)
        for c, t in enumerate(['班', 'No.', '作　業　項　目', '人員'], 1): merge(ws, 4, c, 6, c, t, font(9, True), AC, B, fill('DDEBF7'))
        merge(ws, 4, last, 6, last, '備　考', font(9, True), AC, B, fill('DDEBF7'))
        r = day_header(ws, 4, C0, dl)
        put(ws, r, 3, '主要工程', font(8, True), AL, B, MS)
        for d, t in p['milestones']:
            if d in dl: put(ws, r, C0 + dl.index(d), t[:2], font(6, True, 'C00000'), AC)
        r += 1
        order = sorted(items, key=lambda it: (it['team'] or 'zz', items.index(it)))
        prev = None
        for it in order:
            team = it['team'] if it['co'] != '市運転班' else '市'
            tw = team if team != prev else '〃'
            put(ws, r, 1, tw, font(8), AC, B); prev = team
            put(ws, r, 2, it['no'], font(8), AC, B)
            nm = it['name'] + ('' if it['co'] in ('千草環境', '市運転班') else f"（{it['co']}）")
            put(ws, r, 3, nm, font(8), AL, B)
            put(ws, r, 4, it['ppl'] or '', font(8), AC, B)
            for i in range(len(dl)): ws.cell(r, C0 + i).border = box(thin, hair, hair, thin)
            bar(ws, r, C0, dl, it['plan'], GRAYP); bar(ws, r, C0, dl, it['act'], 'dot')
            put(ws, r, last, it['note'], font(7, color='C00000' if it['note'] else None), AL, B)
            rec.add(ws.title, r, 'item', dict(team=tw, no=it['no'], name=nm, ppl=it['ppl'] or None, planStart=it['plan'][0] if it['plan'] else None,
                    planEnd=it['plan'][1] if it['plan'] else None, actStart=it['act'][0] if it['act'] else None, actEnd=it['act'][1] if it['act'] else None,
                    note=it['note']), it.get('wids'))
            r += 1
        page(ws, land=True, paper=8, fitw=1, fith=1)
        return
    new = f == 'p2b'
    WK['on'] = len(dl) > 70
    if WK['on']:
        d0 = p['start'] - dt.timedelta(p['start'].weekday())
        dl = [d0 + dt.timedelta(7 * i) for i in range(((p['end'] - d0).days) // 7 + 1)]
    C0 = 7
    widths(ws, [5, 20, 32, 11, 5, 5] + [2.7 if not WK['on'] else 3.6] * len(dl) + [22])
    last = C0 + len(dl)
    merge(ws, 1, 1, 1, last, f"【架空】{title}　{p['part']}　工程表", font(14, True), AL, b=None)
    merge(ws, 2, 1, 2, last, f"工期　{p['start']:%Y年%m月%d日} 〜 {p['end']:%Y年%m月%d日}　　{header_info}　　作成　{made}　　千草環境システム株式会社　{staff}", font(9), AL, b=None)
    put(ws, 3, 1, '凡例：上段 予定（━）／下段 実績（塗り）' + ('　別契約は橙　◆＝節目（停止条件の帯の中）' if new else '　黄色の日付＝節目（' + '・'.join(t for d, t in p['milestones']) + '）'), font(8), AL)
    r0 = 5; hdr = 4 if new else 3
    for c, t in enumerate(['No.', '対象機器', '作　業　項　目', '担当会社', '延人日', '区分'], 1):
        merge(ws, r0, c, r0 + hdr - 1, c, t, font(9, True), AC, B, fill('DDEBF7'))
    merge(ws, r0, last, r0 + hdr - 1, last, '備　考', font(9, True), AC, B, fill('DDEBF7'))
    if WK['on']:
        m0 = None; r = r0
        for i, d in enumerate(dl):
            if d.month != m0:
                j = i
                while j + 1 < len(dl) and dl[j + 1].month == d.month: j += 1
                merge(ws, r, C0 + i, r, C0 + j, f'{d.year}年{d.month}月', font(8, True), AC, B, fill('DDEBF7')); m0 = d.month
            put(ws, r + 1, C0 + i, f'{d.day}', font(7), AC, B); put(ws, r + 2, C0 + i, '週', font(6), AC, B)
            put(ws, r + 3, C0 + i, f'{i+1}', font(6), AC, B, fill('EDEDED'))
        r = r0 + 4
        put(ws, 3, 12, '（期間が長いため週単位：各列は月曜はじまりの 1 週間）', font(8, color='C00000'), AL)
    else:
        r = day_header(ws, r0, C0, dl, week_from=p['start'] if new else None)
    if not new:
        for d, t in p['milestones']:
            if d in dl:
                ws.cell(r - 2, C0 + dl.index(d)).fill = MS; ws.cell(r - 2, C0 + dl.index(d)).comment = Comment(t, 'p2')
    else:
        put(ws, r, 3, '炉の状態（停止条件）', font(8, True), AL, B, MS)
        for c in (1, 2, 4, 5, 6): put(ws, r, c, None, None, None, B, MS)
        used = [-1]
        for a, b, t in p['stop']:
            if WK['on']:
                ia = max(0, (a - dl[0]).days // 7); ib = min(len(dl) - 1, (b - dl[0]).days // 7)
            else:
                a = max(a, dl[0]); b = min(b, dl[-1])
                if a > b: continue
                ia, ib = dl.index(a), dl.index(b)
            ms = ' '.join(f"◆{tt}{d.month}/{d.day}" for d, tt in p['milestones'] if a <= d <= b)
            ia = max(ia, used[0] + 1)
            if ib < ia: continue
            used[0] = ib
            merge(ws, r, C0 + ia, r, C0 + ib, f'{t} {ms}'.strip(), font(6.5, True), AC, B, fill('FCE4D6') if t != '運転' else fill('E2EFDA'))
        ws.row_dimensions[r].height = 22; r += 1
    seq = []
    if new and p.get('groups'):
        byno = {it['no']: it for it in items}; used = set()
        for g, ks in p['groups']:
            xs = [byno[k] for k in ks if k in byno]
            if not xs: continue
            seq.append(('G', g)); seq += [('I', x) for x in xs]; used |= {x['no'] for x in xs}
        seq += [('I', it) for it in items if it['no'] not in used]
    else:
        seq = [('I', it) for it in items]
    for kind, it in seq:
        if kind == 'G':
            merge(ws, r, 1, r, 6, f'■ {it}', font(9, True), AL, B, fill('EDEDED'))
            for i in range(len(dl)): ws.cell(r, C0 + i).border = box(thin, hair, hair, thin)
            put(ws, r, last, None, None, None, B); r += 1; continue
        r = item_rows(ws, r, C0, dl, last, it, rec, GREEN, ACT)
    blocks = other_blocks(p) if other_blocks else []
    if new and blocks:
        r += 1
        merge(ws, r, 1, r, 6, '■ 別契約（参考）　同じ停止期間に施工', font(9, True, 'C55A11'), AL, B, fill('FBE5D6'))
        for i in range(len(dl)): ws.cell(r, C0 + i).border = box(thin, hair, hair, thin)
        r += 1
        for cname, its in blocks:
            merge(ws, r, 1, r, 6, cname, font(8, True), AL, B); r += 1
            for it in its: r = item_rows(ws, r, C0, dl, last, it, rec, OTH, OTH_A, ref=True)
    put(ws, r + 1, 2, '※ 延人日は標準人員×実績日数（予定のみの作業は予定日数）。日々の人員・工数は作業日報による。', font(7), AL)
    ws.freeze_panes = ws.cell(r0 + hdr, C0)
    tall = len(seq) + sum(len(x[1]) for x in blocks) > 32
    page(ws, land=not tall, paper=8, fitw=1, fith=1)

def item_rows(ws, r, C0, dl, last, it, rec, color, actfill, ref=False):
    ud = unit_disp(it)
    merge(ws, r, 1, r + 1, 1, it['no'], font(8), AC, B)
    merge(ws, r, 2, r + 1, 2, ud, font(7), ALW, B)
    merge(ws, r, 3, r + 1, 3, it['name'], font(8), ALW, B)
    merge(ws, r, 4, r + 1, 4, it['co'], font(7), AC, B)
    md = mandays(it['ppl'], it['act'] or it['plan'])
    merge(ws, r, 5, r + 1, 5, md or '—', font(8), AC, B)
    put(ws, r, 6, '予定', font(7), AC, B); put(ws, r + 1, 6, '実績', font(7), AC, B)
    for i in range(len(dl)):
        ws.cell(r, C0 + i).border = box(thin, hair, hair, dot); ws.cell(r + 1, C0 + i).border = box(dot, hair, hair, thin)
    bar(ws, r, C0, dl, it['plan'], 'line', color); bar(ws, r + 1, C0, dl, it['act'], actfill)
    note = it['note']
    merge(ws, r, last, r + 1, last, note, font(7, color='C00000' if (it['plan'] != it['act']) else None), ALW, B)
    rec.add(ws.title, r, 'item', dict(no=it['no'], unit=ud, name=it['name'], company=it['co'], manDays=md, planStart=it['plan'][0] if it['plan'] else None,
            planEnd=it['plan'][1] if it['plan'] else None, actStart=it['act'][0] if it['act'] else None, actEnd=it['act'][1] if it['act'] else None,
            note=note, ref='別契約' if ref else None), None if ref else it.get('wids'))
    return r + 2
