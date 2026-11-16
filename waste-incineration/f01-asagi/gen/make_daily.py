# -*- coding: utf-8 -*-
from openpyxl import Workbook
from openpyxl.comments import Comment
from xlutil import *
from f01_contracts import *
from daily_data import H23_DAYS, R7_DAYS
import f01_truth as TR
def _ids(fy, tasks, no):
    t = next((t for t in tasks if t['no'] == no), None)
    return [i for g in (t['tags'] if t else ()) for i in TR.work_ids(fy, g)]

WD = '月火水木金土日'
MD = dict(title=None, agent='m2', city='a2', fy=2011, changes=None)
HOOK = [None]
def _hk(*a):
    if HOOK[0]:
        HOOK[0](*a)
def _wa(x):
    return ('平成%d年%d月%d日' % (x.year - 1988, x.month, x.day)) if (x.year, x.month) < (2019, 5) else ('令和%d年%d月%d日' % (x.year - 2018, x.month, x.day))

def m2_sheet(wb, x, dd):
    """m2 様式（2011〜）A4 縦。1 日 1 シート。"""
    ws = wb.create_sheet('%d.%d' % (x.month, x.day))
    fn = 'ＭＳ Ｐゴシック'
    widths(ws, {'A': 5, 'B': 6, 'C': 17, 'D': 42, 'E': 5, 'F': 5, 'G': 5, 'H': 6, 'I': 6})
    merge(ws, 1, 1, 1, 7, '作　業　日　報', F(16, True, name=fn), AC, border=None)
    put(ws, 1, 8, 'No.', F(9, name=fn), AR); put(ws, 1, 9, dd['no'], F(11, name=fn), AC, box(b=thin, l=NONE, r=NONE, t=NONE))
    put(ws, 2, 1, '工事名', F(9, name=fn), AC, B_ALL); ws.merge_cells('A2:B2')
    merge(ws, 2, 3, 2, 9, MD['title'] or H23['title'], F(10, name=fn), AL)
    put(ws, 3, 1, '日付', F(9, name=fn), AC, B_ALL); ws.merge_cells('A3:B3')
    merge(ws, 3, 3, 3, 4, '%s（%s）' % (_wa(x), WD[x.weekday()]), F(10, name=fn), AL)
    put(ws, 3, 5, '天候', F(9, name=fn), AC, B_ALL); merge(ws, 3, 6, 3, 7, dd['weather'], F(10, name=fn), AC)
    put(ws, 3, 8, '気温', F(9, name=fn), AC, B_ALL); put(ws, 3, 9, dd['temp'] + '℃', F(10, name=fn), AC, B_ALL)
    put(ws, 4, 1, '時間', F(9, name=fn), AC, B_ALL); ws.merge_cells('A4:B4')
    merge(ws, 4, 3, 4, 4, dd['time'] + '（休憩 12:00〜13:00）', F(10, name=fn), AL)
    merge(ws, 4, 5, 4, 6, '代理人', F(9, name=fn), AC); merge(ws, 4, 7, 4, 9, MD['agent'], F(10, name=fn), AC)
    grid(ws, 2, 1, 4, 9)
    r = 6
    hd = ['No', '号炉', '作業箇所・設備', '作　業　内　容', '社員', '協力', '計', '工数\n(人h)', '進捗\n(%)']
    for i, h in enumerate(hd):
        put(ws, r, i + 1, h, F(9, True, name=fn), AC, B_ALL, fill('F2F2F2'))
    ws.row_dimensions[r].height = 26
    r += 1
    ts = tn = th = 0
    for (no, team, f, place, text, s, k, h, prog) in dd['rows']:
        TR.r(ws.title, r, 'daily_row:%s:%s' % (x.isoformat(), no), _ids(2011, H23_F1, no))
        put(ws, r, 1, no, F(9, name=fn), AC, B_ALL)
        put(ws, r, 2, f, F(9, name=fn), AC, B_ALL)
        put(ws, r, 3, place, F(9, name=fn), AL, B_ALL)
        put(ws, r, 4, text, F(9, name=fn), AL, B_ALL)
        put(ws, r, 5, s, F(9, name=fn), AC, B_ALL); put(ws, r, 6, k, F(9, name=fn), AC, B_ALL)
        put(ws, r, 7, '=E%d+F%d' % (r, r), F(9, name=fn), AC, B_ALL)
        put(ws, r, 8, '=G%d*%s' % (r, h), F(9, name=fn), AC, B_ALL, fmt='0.0')
        put(ws, r, 9, prog, F(9, name=fn), AC, B_ALL)
        _hk(ws.title, r, 'item', dict(date=x.isoformat(), no=no, furnace=f, place=place, content=text, staff=s, partner=k, persons=s + k, manDays=round((s + k) * h / 8.0, 2), manDaysText='%.1f' % ((s + k) * h), progress=prog), no)
        lines = max(1, (len(text) + 23) // 24, (len(place) + 9) // 10)
        ws.row_dimensions[r].height = 14 * lines + 2
        r += 1
    # 空き行（m2 は 12 行固定の枠で使う）
    while r < 7 + 12:
        grid(ws, r, 1, r, 9)
        r += 1
    _tp = sum(x_[5] + x_[6] for x_ in dd['rows']); _tm = sum((x_[5] + x_[6]) * x_[7] for x_ in dd['rows'])
    _hk(ws.title, r, 'total', dict(date=x.isoformat(), persons=_tp, manDays=round(_tm / 8.0, 2), manDaysText='%.1f' % _tm), None)
    put(ws, r, 1, '合　計', F(9, True, name=fn), AC, B_ALL); ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    for c, col in ((5, 'E'), (6, 'F'), (7, 'G'), (8, 'H')):
        put(ws, r, c, '=SUM(%s7:%s%d)' % (col, col, r - 1), F(9, True, name=fn), AC, box(t=med, b=med, l=thin, r=thin), fmt='0.0' if c == 8 else '0')
    put(ws, r, 9, '', None, None, B_ALL)
    r += 2
    for lab, v in (('安全（KY）', dd['ky']), ('ヒヤリハット', dd['hiyari']), ('立会・検査', dd['witness']), ('記　事', dd['note']), ('明日の予定', dd['tomorrow'])):
        merge(ws, r, 1, r, 2, lab, F(9, name=fn), AC)
        merge(ws, r, 3, r, 9, v, F(9, name=fn), AL)
        ws.row_dimensions[r].height = 30 if len(v) > 40 else 18
        r += 1
    r += 1
    merge(ws, r, 6, r, 7, '市確認', F(9, name=fn), AC); merge(ws, r, 8, r, 9, '代理人', F(9, name=fn), AC)
    merge(ws, r + 1, 6, r + 3, 7, MD['city'], F(9, name=fn, color='808080'), AC); merge(ws, r + 1, 8, r + 3, 9, MD['agent'], F(9, name=fn, color='808080'), AC)
    merge(ws, r + 4, 1, r + 4, 9, '汐見重工環境株式会社　環境プラント事業部　あさぎ市清掃工場 定期整備作業所', F(8, name=fn), AL, border=None)
    page(ws, landscape=False, paper=9, fith=1, area='A1:I%d' % (r + 4))
    return ws

def build_h23():
    TR.set_file('【架空】あさぎ市清掃工場_作業日報_平成23年度定期整備業務委託_1号炉_m2様式.xlsx')
    wb = Workbook(); wb.remove(wb.active)
    for x in sorted(H23_DAYS):
        m2_sheet(wb, x, H23_DAYS[x])
    p = OUT + '/【架空】あさぎ市清掃工場_作業日報_平成23年度定期整備業務委託_1号炉_m2様式.xlsx'
    wb.save(p)
    return p

TEAMS = {'A': 'A班（炉・築炉）', 'B': 'B班（ボイラ・排ガス）', 'C': 'C班（回転機・機械）', 'D': 'D班（電気計装）', '運': '運転操作（運転委託先）'}

def m3_sheet(wb, x, dd, sched):
    """m3 様式（2021〜）A4 横。m2 様式をコピーして作り替え：班別ブロック・工程No・予定/実績進捗・発見事項欄。"""
    ws = wb.create_sheet('%02d%02d' % (x.month, x.day))
    fn = 'ＭＳ ゴシック'
    widths(ws, {'A': 13, 'B': 6, 'C': 7, 'D': 20, 'E': 52, 'F': 5, 'G': 5, 'H': 5, 'I': 7, 'J': 6, 'K': 6, 'L': 7, 'M': 2})
    # m2 様式から残った見出しセル（No. 欄）と日報名
    merge(ws, 1, 1, 1, 9, '定期整備　作業日報', F(15, True, name=fn), AL, border=None)
    put(ws, 1, 10, 'No.', F(9, name=fn), AR); merge(ws, 1, 11, 1, 12, dd['no'], F(11, name=fn), AC, border=box(b=thin, t=NONE, l=NONE, r=NONE))
    put(ws, 2, 1, '件名', F(9, name=fn), AC, B_ALL); merge(ws, 2, 2, 2, 6, MD['title'] or R7['title'], F(10, name=fn), AL)
    merge(ws, 2, 7, 2, 8, '日付', F(9, name=fn), AC); merge(ws, 2, 9, 2, 12, 'R%d.%d.%d（%s）' % (x.year - 2018, x.month, x.day, WD[x.weekday()]), F(10, True, name=fn), AC)
    put(ws, 3, 1, '天候／気温', F(9, name=fn), AC, B_ALL); merge(ws, 3, 2, 3, 4, '%s　%s℃' % (dd['weather'], dd['temp']), F(10, name=fn), AL)
    put(ws, 3, 5, '作業時間 %s　（休憩 10:00〜10:15・12:00〜13:00・15:00〜15:15）' % dd['time'], F(9, name=fn), AL, B_ALL)
    merge(ws, 3, 6, 3, 8, '代理人', F(9, name=fn), AC); merge(ws, 3, 9, 3, 12, MD['agent'], F(10, name=fn), AC)
    r = 5
    hd = ['班', '工程No', '号炉', '作業箇所', '作業内容・結果', '社員', '協力', '時間', '工数(人h)', '予定%', '実績%', '状況']
    for i, h in enumerate(hd):
        put(ws, r, i + 1, h, F(9, True, name=fn), AC, B_ALL, fill('DDEBF7'))
    r += 1
    first_row = r
    sub_rows = []
    for team in ['運', 'A', 'B', 'C', 'D']:
        rows = [row for row in dd['rows'] if row[1] == team]
        if not rows:
            continue
        r0 = r
        for (no, tm, f, place, text, s, k, h, prog) in rows:
            t = next(t for t in sched if t['no'] == no)
            wd = work_days(t)
            plan = int(round(100.0 * (len([y for y in wd if y <= x]) / len(wd)))) if wd else 100
            ch = (MD['changes'] or {}).get(no)
            TR.r(ws.title, r, 'daily_row:%s:%s' % (x.isoformat(), no), _ids(2025, sched, no))
            put(ws, r, 2, no, F(9, name=fn), AC, B_ALL); put(ws, r, 3, f, F(9, name=fn), AC, B_ALL)
            put(ws, r, 4, place, F(9, name=fn), AL, B_ALL); put(ws, r, 5, text, F(9, name=fn), AL, B_ALL)
            put(ws, r, 6, s, F(9, name=fn), AC, B_ALL); put(ws, r, 7, k, F(9, name=fn), AC, B_ALL)
            put(ws, r, 8, h, F(9, name=fn), AC, B_ALL, fmt='0.0')
            put(ws, r, 9, '=(F%d+G%d)*H%d' % (r, r, r), F(9, name=fn), AC, B_ALL, fmt='0.0')
            if x > t['e']:
                plan = 100
            put(ws, r, 10, plan, F(9, name=fn), AC, B_ALL); put(ws, r, 11, prog, F(9, name=fn), AC, B_ALL)
            st = '完了' if prog >= 100 else ('遅れ' if prog < plan - 10 else '継続')
            c = put(ws, r, 12, st, F(9, True, name=fn, color='C00000' if st == '遅れ' else '000000'), AC, B_ALL)
            if st == '遅れ' and ch:
                c.comment = Comment(ch['note'], 'm3')
            _hk(ws.title, r, 'item', dict(date=x.isoformat(), team=TEAMS[team], no=no, furnace=f, place=place, content=text, staff=s, partner=k, manDays=round((s + k) * h / 8.0, 2), manDaysText='%.1f' % ((s + k) * h), plan=plan, progress=prog, status=st), no)
            ws.row_dimensions[r].height = 13 * max(1, (len(text) + 31) // 32, (len(place) + 10) // 11) + 3
            r += 1
        merge(ws, r0, 1, r - 1, 1, TEAMS[team], F(9, True, name=fn), AC)
        # 班小計
        put(ws, r, 1, '', None, None, B_ALL); merge(ws, r, 2, r, 5, '小計', F(8, name=fn), AR, fl=fill('F2F2F2'))
        for c, col in ((6, 'F'), (7, 'G'), (9, 'I')):
            put(ws, r, c, '=SUM(%s%d:%s%d)' % (col, r0, col, r - 1), F(8, name=fn), AC, B_ALL, fill('F2F2F2'), fmt='0.0' if c == 9 else '0')
        for c in (8, 10, 11, 12):
            put(ws, r, c, None, None, None, B_ALL, fill('F2F2F2'))
        _sp = sum(x_[5] + x_[6] for x_ in rows); _sm = sum((x_[5] + x_[6]) * x_[7] for x_ in rows)
        _hk(ws.title, r, 'subtotal', dict(label='小計', staff=sum(x_[5] for x_ in rows), partner=sum(x_[6] for x_ in rows), manDays=round(_sm / 8.0, 2), manDaysText='%.1f' % _sm), None)
        sub_rows.append(r)
        r += 1
    _tp = sum(x_[5] + x_[6] for x_ in dd['rows']); _tm = sum((x_[5] + x_[6]) * x_[7] for x_ in dd['rows'])
    _hk(ws.title, r, 'total', dict(date=x.isoformat(), staff=sum(x_[5] for x_ in dd['rows']), partner=sum(x_[6] for x_ in dd['rows']), manDays=round(_tm / 8.0, 2), manDaysText='%.1f' % _tm), None)
    merge(ws, r, 1, r, 5, '合　計（別契約の補修工事は含まない）', F(9, True, name=fn), AR)
    for c, col in ((6, 'F'), (7, 'G'), (9, 'I')):
        put(ws, r, c, '=' + '+'.join('%s%d' % (col, sr) for sr in sub_rows), F(9, True, name=fn), AC, box(t=med, b=med, l=thin, r=thin), fmt='0.0' if c == 9 else '0')
    for c in (8, 10, 11, 12):
        put(ws, r, c, None, None, None, box(t=med, b=med, l=thin, r=thin))
    r += 2
    merge(ws, r, 1, r, 12, '■ 発見事項・処置', F(10, True, name=fn), AL, border=None)
    r += 1
    for i, h in enumerate(['機器', '内容', '処置・方針', '報告']):
        cs = [(1, 3), (4, 5), (6, 10), (11, 12)][i]
        merge(ws, r, cs[0], r, cs[1], h, F(9, True, name=fn), AC, fl=fill('DDEBF7'))
    r += 1
    finds = dd['finds'] or [('—', 'なし', '', '')]
    for fd in finds:
        for i, v in enumerate(fd):
            cs = [(1, 3), (4, 5), (6, 10), (11, 12)][i]
            merge(ws, r, cs[0], r, cs[1], v, F(9, name=fn), AL)
        ws.row_dimensions[r].height = 13 * max(1, (len(fd[2]) + 19) // 20, (len(fd[1]) + 29) // 30, (len(fd[3]) + 7) // 8) + 4
        r += 1
    r += 1
    for lab, v in (('作業許可', dd['permits']), ('KY', dd['ky']), ('立会', dd['witness']), ('翌日（工程表）', dd['tomorrow'])):
        put(ws, r, 1, lab, F(9, name=fn), AC, B_ALL)
        merge(ws, r, 2, r, 9, v, F(9, name=fn), AL)
        r += 1
    merge(ws, r - 4, 10, r - 4, 12, '市確認（%s）' % MD['city'], F(8, name=fn), AC)
    merge(ws, r - 3, 10, r - 1, 12, '', None, AC)
    page(ws, landscape=True, paper=9, fith=1, area='A1:L%d' % (r - 1))
    return ws

def build_r7():
    TR.set_file('【架空】あさぎ市清掃工場_作業日報_令和7年度定期整備業務委託_1号炉・全停_m3様式.xlsx')
    wb = Workbook(); wb.remove(wb.active)
    for x in sorted(R7_DAYS):
        m3_sheet(wb, x, R7_DAYS[x], R7_F1)
    p = OUT + '/【架空】あさぎ市清掃工場_作業日報_令和7年度定期整備業務委託_1号炉・全停_m3様式.xlsx'
    wb.save(p)
    return p

if __name__ == '__main__':
    p = build_h23(); print(render(p, pages=(3, 5)))
    p = build_r7(); print(render(p, pages=(5, 6)))
