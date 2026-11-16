# -*- coding: utf-8 -*-
"""作業日報（全年度）。定期整備：n1 様式（1日1枚）R3〜R6（R6 は n2 が n1 様式を使用）、n2 週報 R7。
基幹改良：n3 工事日報（1日1枚・横型）R4・R5（DCS切替を含む）。工程（works）から作る。原価は書かない。"""
import os, datetime as dt
from openpyxl import Workbook
from openpyxl.utils import get_column_letter as L
from common import *
import works as W
import schedules as S
import rec

D = dt.date
KYS = ['足場上の工具落下防止', '回転体の巻込まれ', '炉内の粉じん（防じんマスク）', '重量物の荷揚げ合図', 'マンホール閉止前の残留者確認', '油圧の残圧開放',
       '停電・復電の確認呼称', 'BF内の酸欠・粉じん', '高所作業の安全帯', '吊り荷の下に入らない', '立入禁止区画の確認', '試運転時の連絡合図', '溶接火花の養生',
       '段差・開口部の転落', '熱中症（水分補給）', '感電防止（検電）']


def hours(t):
    if not t:
        return 0
    a, b = t.split('～')
    h1, m1 = map(int, a.split(':')); h2, m2 = map(int, b.split(':'))
    v = (h2 * 60 + m2 - h1 * 60 - m1) / 60
    return v - 1.0 if v > 6 else v


# R4 全炉停止（見本の手書き日報）の行 → 工程の作業
R4KEY = [('性能検査', 'q2'), ('年次自主', 'q'), ('ワイヤ', 'p'), ('足場', 'a'), ('クリンカ', 'b'), ('火格子 計測', 'c'), ('火格子 点検', 'c'), ('火格子片', 'd'),
         ('耐火物', 'e'), ('給じん', 'f'), ('油圧ユニット', 'f'), ('落じん', 'f'), ('摺動板', 'g2'), ('水張り', 'g2'), ('灰押出装置 点検', 'g'),
         ('IDF', 'h'), ('送風機', 'i'), ('ダンパ', 'i'), ('減温塔噴霧制御盤', 'o'), ('減温塔', 'j'), ('BF制御盤', 'o'), ('BF', 'k'), ('供給機', 'l'),
         ('分析計', 'n'), ('温度計', 'm'), ('差圧計', 'm'), ('流量計', 'm'), ('計装 校正', 'm'), ('MCC', 'o'), ('炉前', 'o'), ('カメラ', 'o'),
         ('投入扉', 'r'), ('冷却水ポンプ', 's'), ('空気圧縮機', 't'), ('全停電', 'u'), ('単体試運転', 'v'), ('インターロック', 'v'), ('ループ', 'v'),
         ('各機器', 'v'), ('総合試運転', 'w'), ('組合確認', 'w'), ('炉内 最終', 'v'), ('〃', None)]


def r4_wid(text, prev):
    for k, iid in R4KEY:
        if k in text:
            if iid is None:
                return prev
            return W.HAND_IID[('R4', iid)]
    return prev


def daily_sheet(stem, ws, no, d, rec_, cname, sv, ev_label, hand=False):
    for c, w in zip('ABCDEFGH', (4, 36, 13, 11, 6, 12, 8, 2)):
        ws.column_dimensions[c].width = w
    merge(ws, 1, 1, 1, 7, '【架空】作　業　日　報', f(16, True, u='single'), AC)
    put(ws, 2, 7, f'No.{no}', f(9), AR)
    merge(ws, 3, 1, 3, 7, f'工事件名　{cname}（{ev_label}）　　{FAC}', f(9), AL)
    merge(ws, 4, 1, 4, 2, f'日　付　令和{d.year-2018}年{d.month}月{d.day}日（{wd(d)}）', f(10, True), AL)
    put(ws, 4, 3, f'天候　{rec_["w"]}', f(10), AL)
    merge(ws, 4, 4, 4, 7, f'施工　{MAKER}', f(9), AL)
    merge(ws, 5, 1, 5, 2, f'記入者　{sv}（SV）', f(9), AL)
    put(ws, 5, 4, '組合確認', f(8), AC)
    merge(ws, 5, 5, 5, 6, None)
    box(ws, 3, 1, 5, 7)
    ws.row_dimensions[5].height = 22
    r = 7
    for i, t in enumerate(['No', '作　業　内　容', '場　所', '班', '人員', '時　間', '進捗'], 1):
        put(ws, r, i, t, f(9, True), AC, fill('EDEDED'))
    r += 1
    rows = rec_['rows']
    nline = max(12, len(rows))
    team_p, team_h = {}, {}
    prevw = None
    prevplace = None
    head = dict(日付=d.isoformat(), 天候=rec_['w'], 番号=f'No.{no}', 記入者=sv)
    for i in range(nline):
        put(ws, r, 1, i + 1, f(9), AC)
        if i < len(rows):
            if hand:
                wk, place, team, n, tm, prog = rows[i]
                wid = r4_wid(wk, prevw)
            else:
                x = rows[i]
                wk, place, team, n, tm, prog, wid = x['text'], x['place'], x['team'], x['persons'], x['time'], x['prog'], x['wid']
                if place == prevplace and i > 0:
                    place = '〃'
            put(ws, r, 2, wk, f(9, color='C00000' if '追加' in wk or '★' in wk else '000000'), AL)
            put(ws, r, 3, place, f(9), AC if place == '〃' else AL)
            put(ws, r, 4, team, f(8), AC)
            put(ws, r, 5, n if n else '', f(9), AC)
            put(ws, r, 6, tm, f(8), AC)
            put(ws, r, 7, prog, f(9), AC)
            if n:
                team_p.setdefault(team, []).append(n)
                team_h[team] = team_h.get(team, 0) + n * hours(tm)
            mh = n * hours(tm) if n else 0
            rec.add(stem, ws.title, r, 'item', dict(head, No=i + 1, 作業内容=wk, 場所=place, 班=team, 人員=n or None, 時間=tm or None, 進捗=prog,
                                                    manDays=round(mh / 7.5, 2) if mh else None, manDaysText=f'{n}人' if n else None), works=[wid])
            prevw = wid
            prevplace = x['place'] if not hand else place
        ws.row_dimensions[r].height = 18
        r += 1
    box(ws, 7, 1, r - 1, 7)
    r += 1
    merge(ws, r, 1, r, 7, '人員・工数', f(9, True), AL)
    r += 1
    merge(ws, r, 1, r, 2, '職　種', f(9, True), AC, fill('EDEDED'))
    put(ws, r, 3, '人　数', f(9, True), AC, fill('EDEDED'))
    merge(ws, r, 4, r, 5, '工　数（人・h）', f(9, True), AC, fill('EDEDED'))
    merge(ws, r, 6, r, 7, '備　考', f(9, True), AC, fill('EDEDED'))
    r += 1
    tr = r
    sv_n = 0 if (rows and all(('全停電' in (x[0] if hand else x['text'])) for x in rows)) else 1
    mech = [t for t in team_p if t.startswith('機械')]
    cats = [
        ('SV（監督）', sv_n, sv_n * 7.5, ''),
        ('機械工（' + '・'.join(t.replace('（応援）', '') for t in sorted(mech)) + '）' if mech else '機械工', sum(sum(team_p[t]) for t in mech), sum(team_h[t] for t in mech),
         '応援含む' if any('応援' in t for t in mech) else ''),
        ('電工・計装工（電計班）', sum(team_p.get('電計班', [])), team_h.get('電計班', 0), ''),
        ('外注', sum(sum(v) for k, v in team_p.items() if k.startswith('外注')), sum(v for k, v in team_h.items() if k.startswith('外注')),
         '・'.join(sorted({k[3:-1] for k in team_p if k.startswith('外注')}))),
    ]
    if hand and d == D(2022, 10, 29):
        cats[1] = ('機械工（機械1班・2班）', 5, 5 * 5.5, '')
    if not hand:
        cats[1] = (cats[1][0], cats[1][1], cats[1][2], cats[1][3])
    tp = th = 0
    for nm, p, h, nt in cats:
        merge(ws, r, 1, r, 2, nm, f(9), AL)
        put(ws, r, 3, p, f(9), AC)
        merge(ws, r, 4, r, 5, h, f(9), AC, nf='0.0')
        merge(ws, r, 6, r, 7, nt, f(8), AL)
        rec.add(stem, ws.title, r, 'subtotal', dict(head, 職種=nm, 人数=p, 工数=h, manDays=round(h / 7.5, 2), manDaysText=f'{h:.1f}', 備考=nt))
        tp += p; th += h
        r += 1
    merge(ws, r, 1, r, 2, '合　計', f(9, True), AC)
    put(ws, r, 3, tp, f(9, True), AC)
    merge(ws, r, 4, r, 5, th, f(9, True), AC, nf='0.0')
    merge(ws, r, 6, r, 7, None)
    rec.add(stem, ws.title, r, 'total', dict(head, 人数=tp, 工数=th, manDays=round(th / 7.5, 2), manDaysText=f'{th:.1f}'))
    box(ws, tr - 1, 1, r, 7)
    r += 2
    merge(ws, r, 1, r, 7, '特記事項（発見事項・指示事項）', f(9, True), AL, fill('EDEDED'))
    r += 1
    note = rec_['note'] or '特になし'
    merge(ws, r, 1, r + 2, 7, note, f(9, color='C00000' if '★' in note else '000000'), ALT)
    rec.add(stem, ws.title, r, 'item', dict(head, 特記事項=note))
    box(ws, r - 1, 1, r + 2, 7)
    r += 4
    merge(ws, r, 1, r, 7, '明日の予定', f(9, True), AL, fill('EDEDED'))
    r += 1
    merge(ws, r, 1, r + 1, 7, rec_['tomorrow'], f(9), ALT)
    rec.add(stem, ws.title, r, 'item', dict(head, 明日の予定=rec_['tomorrow']))
    box(ws, r - 1, 1, r + 1, 7)
    r += 3
    ky = KYS[(d.toordinal() * 7) % len(KYS)] if not hand else S_KY.get(d, KYS[d.day % len(KYS)])
    merge(ws, r, 1, r, 2, f'安全　KY：{ky}', f(9), AL)
    merge(ws, r, 3, r, 5, '災害・ヒヤリ：無', f(9), AL)
    merge(ws, r, 6, r, 7, 'TBM 8:20', f(9), AL)
    rec.add(stem, ws.title, r, 'item', dict(head, KY=ky, 災害='無'))
    box(ws, r, 1, r, 7)
    setup_print(ws, False, 9, 1, 1, f'A1:G{r}')


S_KY = {
    D(2022,10,17): '足場上の工具落下防止', D(2022,10,18): 'ワイヤ交換時の巻込まれ', D(2022,10,19): '炉内の粉じん（防じんマスク着用）',
    D(2022,10,20): '重量物（火格子片）の荷揚げ', D(2022,10,21): 'マンホール閉止時の残留者確認', D(2022,10,22): '油圧の残圧',
    D(2022,10,23): '停電・復電操作の確認呼称', D(2022,10,24): 'BF内 酸欠・粉じん', D(2022,10,25): 'ろ布交換時の墜落',
    D(2022,10,26): '旧摺動板の吊り作業', D(2022,10,27): '足場解体時の立入禁止', D(2022,10,28): '試運転時の回転体接近', D(2022,10,29): '試運転時の連絡合図'}


def build_n1_year(year):
    sv = W.TEIKI[year]['sv']
    fname = f'{FN}_作業日報_令和{year-2018}年度定期整備業務委託_n1様式.xlsx'
    stem = rec.start(fname)
    wb = Workbook()
    props(wb, f'【架空】令和{year-2018}年度 やまぶき広域処理組合第二工場 定期整備業務委託 作業日報', f'{MAKER} {sv}')
    first = True
    no = 0
    for ev in ('mid1', 'mid2', 'all'):
        info = W.TEIKI_EV[year].get(ev)
        if not info:
            continue
        hand = info.get('hand', False)
        for d, rr in sorted(info['daily'].items()):
            no += 1
            ws = wb.active if first else wb.create_sheet()
            first = False
            ws.title = f'{d.month}.{d.day}'
            daily_sheet(stem, ws, no, d, rr, W.TEIKI[year]['name'], sv, info['label'], hand=hand)
    wb.save(os.path.join(OUT, fname))
    return fname


# ------------------------------------------------------------------ n2 週報（R7）
R7KEY = [('足場', [1, 25]), ('火格子片', [3]), ('火格子点検', [2, 8]), ('火格子', [2, 8]), ('BF ろ布', [4]), ('BF ろ布抜取', [4]), ('1号IDF', [5]),
         ('熱電対', [6, 14]), ('BFダスト', [7]), ('給じん', [9]), ('耐火物', [10]), ('灰押出', [11]), ('2号IDF', [12]), ('入口ダンパ', [13]),
         ('ノズル', [15]), ('計装校正', [16]), ('分析計', [17]), ('ワイヤ', [18]), ('年次', [19]), ('バケット', [20]), ('投入扉', [21]),
         ('油圧ユニット', [22]), ('ポンプ', [23]), ('全停電', [24]), ('インターロック', [26]), ('単体試運転', [26]), ('盤点検', [16]),
         ('総合試運転', [27]), ('炉内清掃', [2, 8]), ('炉内最終', [26]), ('回転機', [12]), ('計器', [17]), ('クレーン', [18, 19, 20])]


def r7_wids(text):
    res = []
    for k, ids in R7KEY:
        if k in text:
            for i in ids:
                w = W.HAND_IID[('R7', i)]
                if w not in res:
                    res.append(w)
    return res


def weekly_sheet(stem, ws, year, ev_label, teams, notes, sv='n2', weather=None, hand=False):
    days = sorted({d for t in teams.values() for d in t})
    c0 = 3
    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 7
    for i in range(len(days)):
        ws.column_dimensions[L(c0 + i)].width = 22 if len(days) <= 5 else 19
    last = c0 + len(days) - 1
    merge(ws, 1, 1, 1, last - 2, '【架空】作業日報（週報）', f(15, True), AL)
    put(ws, 2, 1, '工事件名', f(9), AL)
    merge(ws, 2, 2, 2, last - 2, f'{W.TEIKI[year]["name"]}（{ev_label}）　{FAC}', f(9), AL)
    put(ws, 3, 1, '期　間', f(9), AL)
    merge(ws, 3, 2, 3, last - 2, f'令和{year-2018}年{days[0].month}月{days[0].day}日（{wd(days[0])}）～{days[-1].month}月{days[-1].day}日（{wd(days[-1])}）　施工 {MAKER}　記入 {sv}（SV）', f(9), AL)
    put(ws, 2, last - 1, '組合確認', f(8), AC); put(ws, 2, last, 'SV', f(8), AC)
    put(ws, 3, last, sv, f(9, color='C00000'), AC)
    box(ws, 2, last - 1, 3, last)
    ws.row_dimensions[3].height = 22
    r = 5
    merge(ws, r, 1, r, 2, '班　／　日', f(9, True), AC, fill('DDEBF7'))
    for i, d in enumerate(days):
        put(ws, r, c0 + i, f'{d.month}/{d.day}（{wd(d)}）', f(9, True, 'C00000' if d.weekday() == 6 else '000000'), AC, fill('FCE4D6') if d.weekday() == 6 else fill('DDEBF7'))
    r += 1
    merge(ws, r, 1, r, 2, '天候', f(8), AC)
    for i, d in enumerate(days):
        put(ws, r, c0 + i, weather.get(d, ''), f(8), AC)
    rec.add(stem, ws.title, r, 'item', dict(期間=[days[0].isoformat(), days[-1].isoformat()], 天候={d.isoformat(): weather.get(d) for d in days if weather.get(d)}))
    box(ws, 5, 1, r, last)
    r += 1
    totp = {d: 0 for d in days}; toth = {d: 0 for d in days}
    st = r
    for team, cells in teams.items():
        merge(ws, r, 1, r + 2, 1, team, f(9, True), AC, fill('F2F2F2'))
        put(ws, r, 2, '作業', f(8), AC); put(ws, r + 1, 2, '人数', f(8), AC); put(ws, r + 2, 2, '工数h', f(8), AC)
        for i, d in enumerate(days):
            v = cells.get(d, ('', 0, 0))
            if len(v) == 4:
                txt, n, h, wids = v
            else:
                txt, n, h = v
                wids = r7_wids(txt) if hand else []
            put(ws, r, c0 + i, txt, f(8.5, color='C00000' if ('再バランス' in txt or '4.8' in txt) else '000000'), ALT,
                fill('FCE4D6') if d.weekday() == 6 else None)
            put(ws, r + 1, c0 + i, n if n else None, f(8), AC)
            put(ws, r + 2, c0 + i, n * h if n else None, f(8), AC, nf='0.0')
            totp[d] += n; toth[d] += n * h
            if txt:
                rec.add(stem, ws.title, r, 'item', dict(日付=d.isoformat(), 班=team, 列=L(c0 + i), 作業=txt, 人数=n or None, 工数=n * h if n else None,
                                                        manDays=round(n * h / 7.5, 2) if n else None, manDaysText=f'{n * h:.1f}' if n else None), works=wids)
        ws.row_dimensions[r].height = 40
        ws.row_dimensions[r + 1].height = 13; ws.row_dimensions[r + 2].height = 13
        r += 3
    merge(ws, r, 1, r, 2, f'SV（{sv}）', f(8), AC)
    for i, d in enumerate(days):
        svn = 0 if d.weekday() == 6 and not any(teams[t].get(d, ('', 0))[1] for t in teams) else 1
        if d.weekday() == 6 and any(teams[t].get(d, ('', 0))[1] for t in teams):
            svn = 1
        put(ws, r, c0 + i, svn if svn else None, f(8), AC)
        totp[d] += svn; toth[d] += svn * 7.5
    r += 1
    merge(ws, r, 1, r, 2, '日計 人数／工数h', f(8, True), AC, fill('DDEBF7'))
    for i, d in enumerate(days):
        put(ws, r, c0 + i, f'{totp[d]}人　／　{toth[d]:.1f}h' if totp[d] else '－', f(8, True), AC, fill('DDEBF7'))
        if totp[d]:
            rec.add(stem, ws.title, r, 'subtotal', dict(日付=d.isoformat(), 列=L(c0 + i), 人数=totp[d], 工数=toth[d], manDays=round(toth[d] / 7.5, 2), manDaysText=f'{toth[d]:.1f}h'))
    box(ws, st, 1, r, last)
    r += 1
    wp, wh = sum(totp.values()), sum(toth.values())
    merge(ws, r, 1, r, 2, '週計', f(8, True), AC)
    merge(ws, r, c0, r, last, f'延べ {wp} 人　／　{wh:.1f} 人・h（外注含む）', f(9, True), AL)
    rec.add(stem, ws.title, r, 'total', dict(延べ人数=wp, 工数=wh, manDays=round(wh / 7.5, 2), manDaysText=f'{wh:.1f}'))
    box(ws, r, 1, r, last)
    r += 2
    merge(ws, r, 1, r, last, '発見事項・特記事項', f(9, True), AL, fill('EDEDED'))
    r += 1
    n0 = r
    for t in notes or ['特になし']:
        merge(ws, r, 1, r, last, '・' + t, f(9, color='C00000' if '4.8' in t else '000000'), AL)
        rec.add(stem, ws.title, r, 'item', dict(特記事項=t))
        ws.row_dimensions[r].height = 16
        r += 1
    box(ws, n0 - 1, 1, r - 1, last)
    r += 1
    merge(ws, r, 1, r, 2, '安全', f(9, True), AC)
    sft = '毎朝 TBM・KY 実施。今週の重点：' + ('回転機の開放作業・重量物の吊り' if days[0].day < 18 else '足場解体・試運転時の回転体接近') + '　災害・ヒヤリ：無'
    merge(ws, r, c0, r, last, sft, f(9), AL)
    rec.add(stem, ws.title, r, 'item', dict(安全=sft))
    box(ws, r, 1, r, last)
    setup_print(ws, True, 8, 1, 1, f'A1:{L(last)}{r}')


def gen_weekly(info):
    """生成した停止 → 班×日 の週報セル"""
    teams = {}
    order = ['機械1班', '機械2班', '機械3班（応援）', '電計班', '外注']
    TN = {'機械1班': '機械1班（5）', '機械2班': '機械2班（4）', '機械3班（応援）': '機械3班（応援）', '電計班': '電計班（3）', '外注': '外注'}
    cells = {}
    for d, rr in info['daily'].items():
        for x in rr['rows']:
            tk = '外注' if x['team'].startswith('外注') else x['team']
            c = cells.setdefault((tk, d), ['', 0, 0, []])
            txt = x['text'] if tk != '外注' else f"{x['team'][3:-1]} {x['persons']}（{x['text']}）"
            c[0] = (c[0] + ' ／ ' if c[0] else '') + txt
            c[1] += x['persons']
            c[2] = max(c[2], hours(x['time']))
            c[3].append(x['wid'])
    for t in order:
        tc = {}
        for (tk, d), v in cells.items():
            if tk == t:
                tc[d] = (v[0], min(v[1], {'機械1班': 5, '機械2班': 4, '電計班': 3}.get(t, 99)), v[2], v[3])
        if tc:
            for d in info['daily']:
                tc.setdefault(d, ('', 0, 0, []))
            teams[TN[t]] = dict(sorted(tc.items()))
    return teams


def build_n2_year(year):
    sv = 'n2'
    fname = f'{FN}_作業日報_令和{year-2018}年度定期整備業務委託_n2週報様式.xlsx'
    stem = rec.start(fname)
    wb = Workbook()
    props(wb, f'【架空】令和{year-2018}年度 やまぶき広域処理組合第二工場 定期整備業務委託 作業日報（週報様式）', f'{MAKER} {sv}')
    first = True
    for ev in ('mid1', 'mid2'):
        info = W.TEIKI_EV[year][ev]
        teams = gen_weekly(info)
        notes = [f"{d.month}/{d.day} {r['note'].rstrip('。')}" for d, r in sorted(info['daily'].items()) if r['note']]
        ws = wb.active if first else wb.create_sheet()
        first = False
        days = sorted(info['daily'])
        ws.title = f'{days[0].month}.{days[0].day}-{days[-1].month}.{days[-1].day}'
        weather = {d: r['w'] for d, r in info['daily'].items()}
        weekly_sheet(stem, ws, year, info['label'], teams, notes[:5], sv, weather)
    wth = {D(2025,10,14): '晴', D(2025,10,15): '晴', D(2025,10,16): '曇', D(2025,10,17): '雨', D(2025,10,18): '曇',
           D(2025,10,19): '晴', D(2025,10,20): '晴', D(2025,10,21): '曇', D(2025,10,22): '晴', D(2025,10,23): '晴',
           D(2025,10,24): '雨', D(2025,10,25): '晴'}
    for wk, teams in sorted(S.R7_WEEKLY.items()):
        ws = wb.create_sheet()
        days = sorted({d for t in teams.values() for d in t})
        ws.title = f'{days[0].month}.{days[0].day}-{days[-1].month}.{days[-1].day}'
        weekly_sheet(stem, ws, year, '全炉停止', teams, S.R7_WEEK_NOTES[wk], sv, wth, hand=True)
    wb.save(os.path.join(OUT, fname))
    return fname


# ------------------------------------------------------------------ n3 工事日報
INSP = {D(2022, 11, 17): '1号炉 火格子 据付検査（組合 y1 立会）合格', D(2022, 12, 9): '年度出来形確認（組合 y1・監督員）',
        D(2022, 10, 28): 'BF ケーシング 気密試験（自主）', D(2022, 11, 4): '減温塔 下部胴 溶接部 非破壊検査（外注）',
        D(2023, 11, 16): '2号炉 火格子 据付検査（組合 y2 立会）合格', D(2023, 12, 8): '出来形確認（組合）', D(2023, 11, 1): '2号BF ケーシング 気密試験',
        D(2024, 3, 8): '完成検査（組合 y2・監督員）合格', D(2024, 2, 15): 'DCS 連動試験 組合確認'}
DELIV = {D(2022, 10, 24): '新火格子 フレーム（トレーラ 2台）', D(2022, 10, 25): '火格子片 一式（乾燥・燃焼・後燃焼）', D(2022, 10, 11): '炉油圧ユニット（1基）',
         D(2022, 10, 21): '減温塔 新下部胴（25t ラフター 手配）', D(2022, 11, 18): 'ろ布 192本＋予備 20本', D(2022, 10, 24): '噴射水ポンプ(A)(B)',
         D(2022, 11, 7): '新DCS 盤（共通FCS・ネットワーク盤・OPS）',
         D(2023, 10, 23): '新火格子 フレーム・火格子片 一式', D(2023, 10, 10): '炉油圧ユニット（1基）', D(2023, 10, 20): '減温塔 新下部胴', D(2023, 11, 17): 'ろ布 192本',
         D(2023, 11, 20): '2号 FCS 盤・BF盤・減温塔盤'}
KNOTE = {D(2022, 9, 20): '1号炉 16:00 ごみ投入停止、立下げ（運転委託）。', D(2022, 10, 14): '2号炉 立下げ（定期整備 全炉停止）。2号炉側への立入禁止区画を設定。',
         D(2022, 10, 20): '炉室ホイストは定期整備優先（10/20〜22）。新火格子搬入を 10/24 へ変更（工程表 第2版）。',
         D(2022, 10, 22): '明日 10/23 全停電のため仮設電源も停止。', D(2022, 12, 5): '1号炉 ごみ投入 開始 9:00。',
         D(2023, 9, 19): '2号炉 立下げ（運転委託）。', D(2023, 10, 31): '★2号 煙道（減温塔入口側）に減肉部を確認（最小 2.9mm／原厚 4.5mm）。監督員へ報告、補修方法を協議。',
         D(2023, 11, 6): '2号煙道 減肉部補修を追加施工（監督員指示）。', D(2023, 12, 4): '2号炉 ごみ投入 開始 9:00。',
         D(2024, 2, 2): '1・2号炉 立下げ（DCS 切替）。運転委託会社 夜間立会。', D(2024, 2, 16): '1号炉 新DCS で立上げ。', D(2024, 2, 18): '2号炉 新DCS で立上げ。'}


def n3_rows(wbs_list, d, key):
    """その日の WBS ごとの行"""
    rows = []
    for (wbs, nm, team, n, spans, kind) in wbs_list:
        if kind in ('H',) or not spans:
            continue
        for s, e in spans:
            if s <= d <= e:
                if kind in ('M', 'K'):
                    continue
                wdays = W.workdays(s, e)
                if d not in wdays:
                    continue
                k = wdays.index(d)
                steps = W.KSTEP.get(wbs, [nm])
                txt = steps[min(len(steps) - 1, int(k * len(steps) / len(wdays)))]
                prog = int(round(100 * (k + 1) / len(wdays)))
                wid = next((w['id'] for w in W.KIKAN_EV.get(key, []) if w.get('wbs') == wbs), None)
                rows.append((wbs, nm, txt, team, n or 0, '完了' if prog == 100 else f'{prog}%', wid))
    return rows


def n3_sheet(stem, ws, no, d, rows, r_):
    for c, w in zip('ABCDEFGH', (6, 26, 34, 26, 6, 8, 18, 2)):
        ws.column_dimensions[c].width = w
    merge(ws, 1, 1, 1, 5, '【架空】工　事　日　報（基幹的設備改良工事）', f(15, True), AL)
    merge(ws, 1, 6, 1, 7, f'No.{no}', f(9), AR)
    merge(ws, 2, 1, 2, 4, f'工事件名　{W.KIKAN["name"]}　　{FAC}', f(9), AL)
    merge(ws, 3, 1, 3, 2, f'日付　令和{d.year-2018}年{d.month}月{d.day}日（{wd(d)}）', f(10, True), AL)
    weather = r_.choice(W.WEATHER)
    put(ws, 3, 3, f'天候　{weather}　　作業時間 8:00～17:00', f(9), AL)
    put(ws, 3, 4, f'受注者　{MAKER}　現場代理人 n3', f(9), AL)
    put(ws, 2, 6, '監督員', f(8), AC); put(ws, 2, 7, '現場代理人', f(8), AC)
    put(ws, 3, 7, 'n3', f(9, color='C00000'), AC)
    box(ws, 2, 6, 3, 7)
    box(ws, 2, 1, 3, 4)
    ws.row_dimensions[3].height = 22
    head = dict(日付=d.isoformat(), 天候=weather, 番号=f'No.{no}')
    r = 5
    for i, t in enumerate(['WBS', '工　種', '本日の作業内容', '協力会社', '人員', '進捗', '備　考'], 1):
        put(ws, r, i, t, f(9, True), AC, fill('FFF2CC'))
    r += 1
    tot = 0
    nline = max(10, len(rows))
    for i in range(nline):
        if i < len(rows):
            wbs, nm, txt, team, n, prog, wid = rows[i]
            sub = W.SUBCON.get(team or '－', '－').replace('【架空】', '')
            put(ws, r, 1, wbs, f(8), AC)
            put(ws, r, 2, nm, f(8.5, color='C00000' if '追加' in nm else '000000'), AL)
            put(ws, r, 3, txt, f(9), AL)
            put(ws, r, 4, sub, f(8), AL)
            put(ws, r, 5, n or None, f(9), AC)
            put(ws, r, 6, prog, f(9), AC)
            put(ws, r, 7, team, f(8), AC)
            tot += n
            rec.add(stem, ws.title, r, 'item', dict(head, WBS=wbs, 工種=nm, 作業内容=txt, 協力会社=sub, 人員=n or None, 進捗=prog, 班=team,
                                                    manDays=n or None, manDaysText=f'{n}人' if n else None), works=[wid])
        ws.row_dimensions[r].height = 17
        r += 1
    put(ws, r, 4, '作業員 計', f(9, True), AR)
    put(ws, r, 5, tot, f(9, True), AC)
    put(ws, r, 7, '＋ 職員 3（n3 ほか）', f(8), AL)
    rec.add(stem, ws.title, r, 'total', dict(head, 作業員計=tot, manDays=tot, manDaysText=f'{tot}人'))
    box(ws, 5, 1, r, 7)
    r += 2
    for lab, txt in (('検査・立会', INSP.get(d, 'なし')), ('搬入資材', DELIV.get(d, 'なし')), ('記事', KNOTE.get(d, '特になし'))):
        put(ws, r, 1, lab, f(9, True), AC, fill('EDEDED'))
        merge(ws, r, 2, r, 7, txt, f(9, color='C00000' if '★' in txt else '000000'), AL)
        rec.add(stem, ws.title, r, 'item', dict(head, **{lab: txt}))
        box(ws, r, 1, r, 7)
        r += 1
    r += 1
    ky = KYS[(d.toordinal() * 5) % len(KYS)]
    put(ws, r, 1, '安全', f(9, True), AC, fill('EDEDED'))
    merge(ws, r, 2, r, 7, f'KY：{ky}　新規入場者：{r_.choice([0, 0, 0, 1, 2])}名　災害・ヒヤリ：無', f(9), AL)
    rec.add(stem, ws.title, r, 'item', dict(head, KY=ky, 災害='無'))
    box(ws, r, 1, r, 7)
    setup_print(ws, True, 9, 1, 1, f'A1:G{r}')


def build_n3():
    out = []
    for key, wbs_list, rngs, nm in (('R4', S.KIKAN_R4, [S.KIKAN_R4_RANGE], '令和4年度'), ('R5', W.KIKAN_R5, [W.KIKAN_R5_RANGE, W.KIKAN_DCS_RANGE], '令和5年度')):
        fname = f'{FN}_作業日報_基幹的設備改良工事_{nm}現地工事_n3様式.xlsx'
        stem = rec.start(fname)
        wb = Workbook()
        props(wb, f'【架空】やまぶき広域処理組合第二工場 基幹的設備改良工事 工事日報 {nm}', f'{MAKER} 現場代理人 n3')
        first = True
        no = 0
        for d0, d1 in rngs:
            for d in W.workdays(d0, d1):
                rows = n3_rows(wbs_list, d, key)
                if not rows and d not in INSP:
                    continue
                no += 1
                ws = wb.active if first else wb.create_sheet()
                first = False
                ws.title = f'{d.month}.{d.day}' + ('' if d.year == d0.year else '')
                n3_sheet(stem, ws, no, d, rows, W.R(('n3', d)))
        wb.save(os.path.join(OUT, fname)); out.append(fname)
    return out


def build_all():
    out = []
    for y in (2021, 2022, 2023, 2024):
        out.append(build_n1_year(y))
    out.append(build_n2_year(2025))
    out += build_n3()
    return out


if __name__ == '__main__':
    print(build_all())
