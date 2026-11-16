# -*- coding: utf-8 -*-
"""作業日報：n1 様式（1 日 1 枚・縦型、R4 全炉停止）と n2 様式（1 週 1 枚・横型の週報、R7 全炉停止）"""
import os, datetime as dt
from openpyxl import Workbook
from openpyxl.utils import get_column_letter as L
from common import *
from schedules import *

KY = {
    D(2022,10,17): '足場上の工具落下防止', D(2022,10,18): 'ワイヤ交換時の巻込まれ', D(2022,10,19): '炉内の粉じん（防じんマスク着用）',
    D(2022,10,20): '重量物（火格子片）の荷揚げ', D(2022,10,21): 'マンホール閉止時の残留者確認', D(2022,10,22): '油圧の残圧',
    D(2022,10,23): '停電・復電操作の確認呼称', D(2022,10,24): 'BF内 酸欠・粉じん', D(2022,10,25): 'ろ布交換時の墜落',
    D(2022,10,26): '旧摺動板の吊り作業', D(2022,10,27): '足場解体時の立入禁止', D(2022,10,28): '試運転時の回転体接近',
    D(2022,10,29): '試運転時の連絡合図',
}


def hours(t):
    if not t:
        return 0
    a, b = t.split('～')
    h1, m1 = map(int, a.split(':')); h2, m2 = map(int, b.split(':'))
    v = (h2 * 60 + m2 - h1 * 60 - m1) / 60
    if v > 6:
        v -= 1.0   # 昼休み
    return v


def build_n1(path):
    wb = Workbook()
    props(wb, '【架空】令和4年度 やまぶき広域処理組合第二工場 定期整備業務委託 作業日報（全炉停止）', f'{MAKER} n1')
    first = True
    for no, (d, rec) in enumerate(sorted(R4_DAILY.items()), 1):
        ws = wb.active if first else wb.create_sheet()
        first = False
        ws.title = f'{d.month}.{d.day}'
        for c, w in zip('ABCDEFGH', (4, 34, 13, 10, 6, 12, 8, 2)):
            ws.column_dimensions[c].width = w
        merge(ws, 1, 1, 1, 7, '【架空】作　業　日　報', f(16, True, u='single'), AC)
        put(ws, 2, 7, f'No.{no}', f(9), AR)
        merge(ws, 3, 1, 3, 7, f'工事件名　{R4_CONTRACT["name"]}（全炉停止）　　{FAC}', f(9), AL)
        merge(ws, 4, 1, 4, 2, f'日　付　令和4年{d.month}月{d.day}日（{wd(d)}）', f(10, True), AL)
        put(ws, 4, 3, f'天候　{rec["w"]}', f(10), AL)
        merge(ws, 4, 4, 4, 7, f'施工　{MAKER}', f(9), AL)
        merge(ws, 5, 1, 5, 2, '記入者　n1（SV）', f(9), AL)
        put(ws, 5, 4, '組合確認', f(8), AC)
        merge(ws, 5, 5, 5, 6, None)
        box(ws, 3, 1, 5, 7)
        ws.row_dimensions[5].height = 22
        r = 7
        for i, t in enumerate(['No', '作　業　内　容', '場　所', '班', '人員', '時　間', '進捗'], 1):
            put(ws, r, i, t, f(9, True), AC, fill('EDEDED'))
        r += 1
        rows = rec['rows']
        nline = max(12, len(rows))
        team_p = {}
        team_h = {}
        for i in range(nline):
            put(ws, r, 1, i + 1, f(9), AC)
            if i < len(rows):
                wk, place, team, n, tm, prog = rows[i]
                put(ws, r, 2, wk, f(9, color='C00000' if '追加' in wk or '★' in wk else '000000'), AL)
                put(ws, r, 3, place, f(9), AC if place == '〃' else AL)
                put(ws, r, 4, team, f(8), AC)
                put(ws, r, 5, n if n else '', f(9), AC)
                put(ws, r, 6, tm, f(8), AC)
                put(ws, r, 7, prog, f(9), AC)
                if n:
                    team_p.setdefault(team, []).append(n)
                    team_h[team] = team_h.get(team, 0) + n * hours(tm)
            ws.row_dimensions[r].height = 18
            r += 1
        box(ws, 7, 1, r - 1, 7)
        r += 1
        merge(ws, r, 1, r, 7, '人員・工数', f(9, True), AL)
        r += 1
        hdr = ['区分', '', '人数', '工数（人・h）', '', '', '']
        put(ws, r, 1, None, fl=fill('EDEDED'))
        merge(ws, r, 1, r, 2, '職　種', f(9, True), AC, fill('EDEDED'))
        put(ws, r, 3, '人　数', f(9, True), AC, fill('EDEDED'))
        merge(ws, r, 4, r, 5, '工　数（人・h）', f(9, True), AC, fill('EDEDED'))
        merge(ws, r, 6, r, 7, '備　考', f(9, True), AC, fill('EDEDED'))
        r += 1
        tr = r
        sv = 0 if d == R4_STOPS['blackout'] else 1
        cats = [
            ('SV（監督）', sv, sv * 7.5, ''),
            ('機械工（機械1班・2班）', sum(team_p.get('機械1班', [])) + sum(team_p.get('機械2班', [])),
             team_h.get('機械1班', 0) + team_h.get('機械2班', 0), ''),
            ('電工・計装工（電計班）', sum(team_p.get('電計班', [])), team_h.get('電計班', 0), ''),
            ('外注', sum(sum(v) for k, v in team_p.items() if k.startswith('外注')),
             sum(v for k, v in team_h.items() if k.startswith('外注')),
             '・'.join(sorted({k[3:-1] for k in team_p if k.startswith('外注')}))),
        ]
        # 同じ班で 〃 行の重複は人数 1 回だけ（同一班員が続けて行う作業）
        if d == D(2022, 10, 29):
            cats[1] = ('機械工（機械1班・2班）', 5, 5 * 5.5, '')
        tp = th = 0
        for nm, p, h, nt in cats:
            merge(ws, r, 1, r, 2, nm, f(9), AL)
            put(ws, r, 3, p, f(9), AC)
            merge(ws, r, 4, r, 5, h, f(9), AC, nf='0.0')
            merge(ws, r, 6, r, 7, nt, f(8), AL)
            tp += p; th += h
            r += 1
        merge(ws, r, 1, r, 2, '合　計', f(9, True), AC)
        put(ws, r, 3, tp, f(9, True), AC)
        merge(ws, r, 4, r, 5, th, f(9, True), AC, nf='0.0')
        merge(ws, r, 6, r, 7, None)
        box(ws, tr - 1, 1, r, 7)
        r += 2
        merge(ws, r, 1, r, 7, '特記事項（発見事項・指示事項）', f(9, True), AL, fill('EDEDED'))
        r += 1
        merge(ws, r, 1, r + 2, 7, rec['note'] or '特になし', f(9, color='C00000' if '★' in rec['note'] else '000000'), ALT)
        box(ws, r - 1, 1, r + 2, 7)
        r += 4
        merge(ws, r, 1, r, 7, '明日の予定', f(9, True), AL, fill('EDEDED'))
        r += 1
        merge(ws, r, 1, r + 1, 7, rec['tomorrow'], f(9), ALT)
        box(ws, r - 1, 1, r + 1, 7)
        r += 3
        merge(ws, r, 1, r, 2, f'安全　KY：{KY.get(d, "")}', f(9), AL)
        merge(ws, r, 3, r, 5, '災害・ヒヤリ：無', f(9), AL)
        merge(ws, r, 6, r, 7, 'TBM 8:20', f(9), AL)
        box(ws, r, 1, r, 7)
        setup_print(ws, False, 9, 1, 1, f'A1:G{r}')
    wb.save(path)


def build_n2(path):
    wb = Workbook()
    props(wb, '【架空】令和7年度 やまぶき広域処理組合第二工場 定期整備業務委託 作業日報（週報様式・全炉停止）', f'{MAKER} n2')
    first = True
    for wk, teams in sorted(R7_WEEKLY.items()):
        ws = wb.active if first else wb.create_sheet()
        first = False
        days = sorted({d for t in teams.values() for d in t})
        ws.title = f'{days[0].month}.{days[0].day}-{days[-1].month}.{days[-1].day}'
        c0 = 3
        ws.column_dimensions['A'].width = 12
        ws.column_dimensions['B'].width = 7
        for i in range(len(days)):
            ws.column_dimensions[L(c0 + i)].width = 22 if len(days) <= 5 else 19
        last = c0 + len(days) - 1
        put(ws, 1, last + 1, None)
        merge(ws, 1, 1, 1, last - 2, '【架空】作業日報（週報）', f(15, True), AL)
        put(ws, 2, 1, '工事件名', f(9), AL)
        merge(ws, 2, 2, 2, last - 2, f'{R7_CONTRACT["name"]}（全炉停止）　{FAC}', f(9), AL)
        put(ws, 3, 1, '期　間', f(9), AL)
        merge(ws, 3, 2, 3, last - 2, f'令和7年{days[0].month}月{days[0].day}日（{wd(days[0])}）～{days[-1].month}月{days[-1].day}日（{wd(days[-1])}）　施工 {MAKER}　記入 n2（SV）', f(9), AL)
        put(ws, 2, last - 1, '組合確認', f(8), AC); put(ws, 2, last, 'SV', f(8), AC)
        put(ws, 3, last, 'n2', f(9, color='C00000'), AC)
        box(ws, 2, last - 1, 3, last)
        ws.row_dimensions[3].height = 22
        r = 5
        merge(ws, r, 1, r, 2, '班　／　日', f(9, True), AC, fill('DDEBF7'))
        for i, d in enumerate(days):
            put(ws, r, c0 + i, f'{d.month}/{d.day}（{wd(d)}）', f(9, True, 'C00000' if d.weekday() == 6 else '000000'), AC, fill('FCE4D6') if d.weekday() == 6 else fill('DDEBF7'))
        r += 1
        merge(ws, r, 1, r, 2, '天候', f(8), AC)
        wth = {D(2025,10,14): '晴', D(2025,10,15): '晴', D(2025,10,16): '曇', D(2025,10,17): '雨', D(2025,10,18): '曇',
               D(2025,10,19): '晴', D(2025,10,20): '晴', D(2025,10,21): '曇', D(2025,10,22): '晴', D(2025,10,23): '晴',
               D(2025,10,24): '雨', D(2025,10,25): '晴'}
        for i, d in enumerate(days):
            put(ws, r, c0 + i, wth.get(d, ''), f(8), AC)
        box(ws, 5, 1, r, last)
        r += 1
        totp = {d: 0 for d in days}
        toth = {d: 0 for d in days}
        st = r
        for team, cells in teams.items():
            merge(ws, r, 1, r + 2, 1, team, f(9, True), AC, fill('F2F2F2'))
            put(ws, r, 2, '作業', f(8), AC)
            put(ws, r + 1, 2, '人数', f(8), AC)
            put(ws, r + 2, 2, '工数h', f(8), AC)
            for i, d in enumerate(days):
                txt, n, h = cells.get(d, ('', 0, 0))
                put(ws, r, c0 + i, txt, f(8.5, color='C00000' if ('再バランス' in txt or '4.8' in txt) else '000000'), ALT,
                    fill('FCE4D6') if d.weekday() == 6 else None)
                put(ws, r + 1, c0 + i, n if n else None, f(8), AC)
                put(ws, r + 2, c0 + i, n * h if n else None, f(8), AC, nf='0.0')
                totp[d] += n; toth[d] += n * h
            ws.row_dimensions[r].height = 40
            for rr in (r + 1, r + 2):
                ws.row_dimensions[rr].height = 13
            r += 3
        # SV
        merge(ws, r, 1, r, 2, 'SV（n2）', f(8), AC)
        for i, d in enumerate(days):
            sv = 0 if d.weekday() == 6 else 1
            if d == D(2025, 10, 19):
                sv = 1
            put(ws, r, c0 + i, sv if sv else None, f(8), AC)
            totp[d] += sv; toth[d] += sv * (9.0 if d == D(2025, 10, 19) else 7.5 if d.weekday() < 5 else 6.0 if d == D(2025,10,25) else 7.5)
        r += 1
        merge(ws, r, 1, r, 2, '日計 人数／工数h', f(8, True), AC, fill('DDEBF7'))
        for i, d in enumerate(days):
            put(ws, r, c0 + i, f'{totp[d]}人　／　{toth[d]:.1f}h' if totp[d] else '－', f(8, True), AC, fill('DDEBF7'))
        box(ws, st, 1, r, last)
        r += 1
        wk_p = sum(totp.values()); wk_h = sum(toth.values())
        merge(ws, r, 1, r, 2, '週計', f(8, True), AC)
        merge(ws, r, c0, r, last, f'延べ {wk_p} 人　／　{wk_h:.1f} 人・h（外注含む）', f(9, True), AL)
        box(ws, r, 1, r, last)
        r += 2
        merge(ws, r, 1, r, last, '発見事項・特記事項', f(9, True), AL, fill('EDEDED'))
        r += 1
        for t in R7_WEEK_NOTES[wk]:
            merge(ws, r, 1, r, last, '・' + t, f(9, color='C00000' if '4.8' in t else '000000'), AL)
            ws.row_dimensions[r].height = 16
            r += 1
        box(ws, r - len(R7_WEEK_NOTES[wk]) - 1, 1, r - 1, last)
        r += 1
        merge(ws, r, 1, r, 2, '安全', f(9, True), AC)
        merge(ws, r, c0, r, last, '毎朝 TBM・KY 実施。今週の重点：' + ('回転機の開放作業・重量物の吊り' if wk == D(2025, 10, 13) else '足場解体・試運転時の回転体接近') + '　災害・ヒヤリ：無', f(9), AL)
        box(ws, r, 1, r, last)
        setup_print(ws, True, 8, 1, 1, f'A1:{L(last)}{r}')
    wb.save(path)


if __name__ == '__main__':
    build_n1(os.path.join(OUT, f'{FN}_作業日報_令和4年度定期整備業務委託_全炉停止_n1様式.xlsx'))
    build_n2(os.path.join(OUT, f'{FN}_作業日報_令和7年度定期整備業務委託_全炉停止_n2週報様式.xlsx'))
