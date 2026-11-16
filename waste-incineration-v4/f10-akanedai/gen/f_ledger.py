# -*- coding: utf-8 -*-
"""機器台帳 旧（j1 2012 統合様式：機器 1 行の右に旧電動機台帳の列を貼った横長 1 シート）と 新（j3 2025、設備別シート＋旧番号・新番号併記＋追加シート）"""
import random, os
from openpyxl import Workbook
from openpyxl.comments import Comment
import world, history as H
from xl import *
from rec import Rec
CFG = {'asof': 2020}

REPL_CATS = {'rot', 'motor', 'valve', 'instr', 'analyzer', 'panel', 'hv', 'scale', 'mech', 'dcs', 'psa', 'crane', 'furnace', 'tuyere', 'taphole', 'egen'}
LEGAL = {'crane': 'クレーン則', 'crane2': 'クレーン則', 'vessel': '一圧', 'boiler': '電事法', 'turbine': '電事法', 'gen': '電事法',
         'scale': '計量法', 'egen': '電事法・消防法', 'hv': '電事法', 'safety': '電事法'}

def model(u, ver):
    r = random.Random(f"{u['tag']}-{ver}")
    k = u['kind']
    if k == '電動機': return f"TW-{r.choice(['IK','FK','HK'])}{int((u['kw'] or 1)*10):04d}"
    if k in ('電動弁', '調節弁'): return f"{r.choice(['TV','HB'])}-{r.randint(100,899)}{r.choice(['A','E','S'])}"
    if k == '計装': return f"{r.choice(['HK','KY'])}{r.randint(10,99)}-{r.randint(100,999)}"
    if k == '盤': return f"{r.choice(['AD','TW'])}-P{r.randint(1000,9999)}"
    return f"CK-{u['sys']}{r.randint(10,99)}{r.choice(['','R','S'])}"

def replaced(u, h, upto):
    ys = [y for y, (s, d) in h.items() if y <= upto and (s == '★' or (s == '●' and u['cat'] in REPL_CATS))]
    return max(ys) if ys else None

def wa_ym(y, m=3):
    if y >= 2019 and (y > 2019 or m >= 5): return f"R{y-2018}.{m}"
    return f"H{y-1988}.{m}"

SEC_OLD = {1: '【1号炉】', 2: '【2号炉】', 0: '【共通】'}

def old_note(u, h):
    A = CFG['asof']
    ry = replaced(u, h, A - 1)
    y = ry or u['year']; note = ''
    if ry:
        ry_txt = wa_ym(ry, 10 if ry != 2019 else 11)
        if A >= 2017 and ry >= 2016 and random.Random(u['tag']).random() < 0.3:
            y = u['year']; note = f'{ry_txt.replace(".", "年")}月 更新（台帳未修正 j2）'
        else:
            note = f'{ry_txt} 更新 ' + ('j1' if ry <= 2015 else 'j2')
    fixed = {(1, '主羽口装置'): '予備羽口 H28購入 3本（H30 2本使用、残1本 市保管）', (1, '出湯口'): 'R1 水冷スリーブ交換',
             (1, 'ろ過式集じん器本体'): 'ろ布 H27・R2 全数交換', (2, 'ろ過式集じん器本体'): 'ろ布 H28 全数交換', (1, '触媒反応塔'): 'H28 上層触媒交換',
             (0, '液体酸素貯槽（CE）'): '高圧ガス 第二種製造者届（PSA予備）', (0, 'PSA吸着塔'): '製品酸素濃度低下傾向（R1〜）'}
    if A == 2020: note = fixed.get((u['furn'], u['base']), note)
    elif A == 2016 and (u['furn'], u['base']) == (1, '主羽口装置'): note = '予備羽口 H28購入 3本 市保管（9月納入予定）'
    return y, note

def old(path, ASOF=2020):
    """j1 H24 統合台帳：機器 1 行。右側に旧電動機台帳（H18〜H23）の列を貼り付けた横長の形。
    電動機は行にせず右の列で持つ（台数は電動機も 1 台ずつ数える）。"""
    CFG['asof'] = ASOF
    us = [u for u in world.UNITS if u.get('ident') != 'ak_new']
    staff = 'j1' if ASOF <= 2015 else 'j2'
    rc = Rec(path.split('/')[-1])
    motors = {}
    for u in us:
        if u['kind'] == '電動機': motors.setdefault(u['parent'], []).append(u)
    rows = [u for u in us if u['kind'] != '電動機']
    wb = Workbook(); ws = wb.active; ws.title = '機器台帳'
    props(wb, '【架空】あかね台資源化センター 機器台帳（機械・電動機統合）', f'あかね台市 環境部 あかね台資源化センター {staff}')
    widths(ws, [15, 4, 13, 30, 5, 30, 10, 11, 7, 2, 8, 4, 7, 6, 4, 9, 12, 7, 26])
    merge(ws, 1, 1, 1, 19, '【架空】あかね台資源化センター　機器台帳（機械・電動機統合）', font(14, True), AL, b=None)
    merge(ws, 2, 1, 2, 19, f'{world.wareki(ASOF)}年4月1日現在　　環境部 あかね台資源化センター 施設係　（H24 j1：機器台帳の右に旧電動機台帳の列を貼り付けて統合）', font(9), AL, b=None)
    merge(ws, 4, 1, 4, 9, '機　器　台　帳（竣工図書 機器表より）', font(9, True), AC, B, fill('D9D9D9'))
    merge(ws, 4, 11, 4, 18, '電動機台帳（旧 H18〜H23 の台帳をそのまま転記）', font(9, True), AC, B, fill('FCE4D6'))
    put(ws, 4, 19, None, None, None, B, fill('D9D9D9'))
    hdr = ['系統図番号', '区分', '設備', '名　　称', '常用\n予備', '形式・仕様', 'メーカー', '製造番号', '設置\n年月', '',
           '電動機No.', '台数', '出力\n(kW)', '電圧\n(V)', '極', '電動機\nメーカー', '電動機\n製造番号', '設置\n年月', '備　　考']
    for c, t in enumerate(hdr, 1):
        if c == 10: continue
        put(ws, 5, c, t, font(8, True), AC, B, fill('FCE4D6') if 11 <= c <= 18 else fill('D9D9D9'))
    ws.row_dimensions[5].height = 26
    ws.cell(5, 2).comment = Comment('機＝機器　弁＝電動弁・調節弁　計＝計装　盤＝盤（H24 統合時 j1）\n電動機は右の電動機台帳欄で 1 台ずつ数える', 'j1')
    kmap = {'機器': '機', '電動弁': '弁', '調節弁': '弁', '計装': '計', '盤': '盤'}
    r = 6; cur = (None, None)
    for u in rows:
        h = H.truth(u)
        if u['furn'] != cur[0]:
            merge(ws, r, 1, r, 19, SEC_OLD[u['furn']], font(10, True), AL, b=box(med, thin, thin, thin), fl=fill('BFBFBF')); r += 1
            cur = (u['furn'], None)
        if u['sys'] != cur[1]:
            merge(ws, r, 1, r, 9, f"{u['furn']}-{u['sys']}　{u['sysname']}", font(9, True), AL, fl=fill('EDEDED'))
            merge(ws, r, 11, r, 19, None, None, AL, fl=fill('EDEDED'))
            r += 1; cur = (cur[0], u['sys'])
        y, note = old_note(u, h)
        nm = ('　' * u['depth']) + u['name']
        typ = model(u, 'old')
        vals = [u['tag'], kmap[u['kind']], u['subname'], nm, u['duty'] or '', f"{typ}　{u['spec']}".strip(),
                strip(u['maker']).replace('千草環境システム', '千草環境'), u['serial'], wa_ym(y, 3 if y == 2006 else 10), None]
        ms = motors.get(u['tag'], [])
        if ms:
            m = ms[0]
            my, mnote = old_note(m, H.truth(m))
            vals += [m['tag'] if len(ms) == 1 else f"{ms[0]['tag']}〜{ms[-1]['tag'][-2:]}", len(ms), m['kw'], m['volt'], m['pole'],
                     strip(m['maker']).replace('【架空】', ''), '／'.join(x['serial'].replace('M', 'TW', 1) for x in ms), wa_ym(my, 3 if my == 2006 else 10)]
            if mnote and not note: note = '電動機 ' + mnote
        else:
            vals += [None] * 8
        vals.append(note)
        for c, v in enumerate(vals, 1):
            if c == 10: continue
            a = AL if c in (1, 3, 4, 6, 7, 8, 11, 16, 17, 19) else AC
            put(ws, r, c, v, font(7 if c in (6, 17, 19) else 8), a, box(hair, thin, thin, hair), fill('FFF8F3') if 11 <= c <= 18 and ms else None)
        keys = ['code', 'kind', 'equipment', 'name', 'duty', 'spec', 'maker', 'serial', 'installed', None, 'motorNo', 'motorQty', 'kW', 'V', 'poles', 'motorMaker', 'motorSerial', 'motorInstalled', 'memo']
        fl = {k: (v.strip('　') if isinstance(v, str) else v) for k, v in zip(keys, vals) if k}
        rc.add(ws.title, r, 'item', fl)
        r += 1
    nm_ = sum(len(v) for v in motors.values())
    put(ws, r + 1, 1, f'計　機器等 {len(rows)} 行 ＋ 電動機 {nm_} 台 ＝ {len(rows) + nm_} 台', font(9, True), AL)
    rc.add(ws.title, r + 1, 'total', dict(rows=len(rows), motors=nm_, units=len(rows) + nm_))
    put(ws, r + 2, 1, '※ 電動機No. は旧電動機台帳の番号（M-）。同じ機器に電動機が 2 台あるものは台数 2 として 1 行に書いた。', font(8), AL)
    ws.freeze_panes = 'E6'
    page(ws, land=True, paper=8, titles='4:5', header='【架空】あかね台資源化センター 機器台帳')
    wb.save(path)
    return len(rows) + nm_

NEW_SHEETS = [('1号炉', lambda u: u['furn'] == 1, '1F'), ('2号炉', lambda u: u['furn'] == 2, '2F'),
              ('共通(受入・副資材)', lambda u: u['furn'] == 0 and u['sys'] in ('01', '02'), 'C1'),
              ('共通(酸素)', lambda u: u['furn'] == 0 and u['sys'] == '03', 'C2'),
              ('共通(スラグ・灰・薬剤)', lambda u: u['furn'] == 0 and u['sys'] in ('04', '06', '07'), 'C3'),
              ('共通(余熱・発電)', lambda u: u['furn'] == 0 and u['sys'] == '05', 'C4'),
              ('共通(用水・排水・空気・付帯)', lambda u: u['furn'] == 0 and u['sys'] in ('08', '09', '10', '13', '14'), 'C5'),
              ('共通(電気)', lambda u: u['furn'] == 0 and u['sys'] == '11', 'C6'),
              ('共通(計装制御)', lambda u: u['furn'] == 0 and u['sys'] == '12', 'C7')]

def new(path, ASOF=2025):
    CFG['asof'] = ASOF
    rc = Rec(path.split('/')[-1])
    wb = Workbook(); ws0 = wb.active; ws0.title = '表紙'
    props(wb, f'【架空】あかね台資源化センター 機器台帳（{world.wareki(ASOF)}年度' + (' 改訂）' if ASOF == 2025 else ' 更新）'), 'あかね台市 環境部 あかね台資源化センター j3')
    widths(ws0, [4, 24, 60])
    merge(ws0, 2, 2, 2, 3, '【架空】あかね台資源化センター　機器台帳', font(18, True), AL, b=None)
    merge(ws0, 3, 2, 3, 3, f'{world.wareki(ASOF)}年4月1日現在（基幹的設備改良工事 完了後の全面改訂' + ('）' if ASOF == 2025 else '、R8 更新：追加シートを各シートへ組込）'), font(11), AL, b=None)
    rows = [('作成', '環境部 あかね台資源化センター 施設係 j3（R7.5 作成）'), ('元台帳', 'H24 j1 統合台帳（機械・電動機）を設備別シートに分割'),
            ('管理No.', 'R7 新番号（シート記号-連番）。旧台帳の系統図番号は「旧番号」列に残す'),
            ('図番', '基幹改良で更新した機器は千草環境の改良工事完成図書の番号（AK-…）'),
            ('数え方', '機器・電動機・電動弁・調節弁・主要計装・盤を各 1 台。手動弁は載せない'),
            ('区分', '機器／電動機／電動弁／調節弁／計装／盤'), ('重要度', 'A：停止で炉停止・法令影響あり（予備なし）　B：予備あり・切替可　C：その他'),
            ('追加分', '基幹改良で新たに設けた機器は「基幹改良追加(R6)」シート（各シートへの組込は未了）' if ASOF == 2025 else 'R6 追加の 3 台は R8 に各シートの末尾へ組込（管理No. は末尾の続き番号）'), ('集計', '「集計」シート')]
    for i, (a, b) in enumerate(rows):
        put(ws0, 5 + i, 2, a, font(10, True), AL, B, fill('EDEDED')); put(ws0, 5 + i, 3, b, font(10), ALW, B)
        ws0.row_dimensions[5 + i].height = 22
    page(ws0, land=False)
    hdr = ['No.', '管理No.', '旧番号\n(系統図/電動機No.)', '改良後図番', '機器名称', '区分', '親機器\n(管理No.)', '常用\n予備', '形式', '主な仕様',
           '出力\nkW', 'メーカー', '製造番号', '設置\n年度', '経過\n年数', '重要度', '法定', '更新履歴', '備考（旧台帳より）']
    W = [4, 8, 14, 13, 30, 6, 8, 5, 12, 26, 6, 11, 12, 6, 5, 5, 8, 20, 22]
    counts = {}
    newmap = {}
    allsheets = NEW_SHEETS + ([('基幹改良追加(R6)', lambda u: u.get('ident') == 'ak_new', 'AK')] if ASOF == 2025 else [])
    for sn, pred, pre in allsheets:
        ws = wb.create_sheet(sn)
        widths(ws, W)
        merge(ws, 1, 1, 1, 12, f'【架空】あかね台資源化センター　機器台帳　{sn}', font(13, True), AL, b=None)
        merge(ws, 2, 1, 2, 12, f'{world.wareki(ASOF)}年4月1日現在　（旧：機器台帳（電動機台帳統合） H24）', font(9), AL, b=None)
        for c, t in enumerate(hdr, 1): put(ws, 4, c, t, font(9, True), AC, B, fill('DDEBF7'))
        ws.row_dimensions[4].height = 28
        r = 5; n = 0; lastsub = None
        us = [u for u in world.UNITS if pred(u) and (sn.startswith('基幹') or u.get('ident') != 'ak_new')]
        if ASOF >= 2026: us += [u for u in world.UNITS if pred(u) and u.get('ident') == 'ak_new' and u not in us]
        for u in us:
            n += 1
            newmap[u['tag']] = f'{pre}-{n:03d}'
        for u in us:
            h = H.truth(u)
            if not sn.startswith('基幹') and (u['sys'], u['sub']) != lastsub:
                merge(ws, r, 1, r, 19, f"{u['sysname']}／{u['subname']}", font(9, True), AL, fl=fill('F2F2F2'))
                lastsub = (u['sys'], u['sub']); r += 1
            ry = replaced(u, h, ASOF - 1)
            y = ry or u['year']
            ak = world.ak_scope(u)
            fig = ''
            if ak and ak[0] <= 2024 or u.get('ident') == 'ak_new':
                fig = f"AK-{u['furn']}{u['sys']}-{random.Random(u['tag']).randint(100, 999)}"
            hist = '、'.join(f"{y2}{'更新' if s in '●★' else '分解' if s == '◎' else '補修'}" for y2, (s, d) in sorted(h.items())
                            if y2 <= ASOF - 1 and s in ('●', '★') and (u['cat'] in REPL_CATS or s == '★'))
            note = ''
            if u['base'] == '主羽口装置' and u['furn'] == 1: note = '予備羽口（旧形）R3購入 2本 保管。基幹改良後の羽口と互換なし・未使用'
            if u['base'] == '2次過熱器' and u['furn'] == 1: note = 'R6 下部ベンド管4本 部分更新'
            if u['base'] == '2次過熱器' and u['furn'] == 2: note = '減肉進行 R7 最小3.2mm（R9 更新予定）'
            if u['base'] == '液体酸素貯槽（CE）': note = '高圧ガス 第二種製造者届'
            if u['base'] == 'ろ過式集じん器本体': note = 'ろ布 H27・R2 全数交換' if u['furn'] == 1 else 'ろ布 H28・R3 全数交換'
            if u.get('ident') == 'ak_new': note = 'R6.3 基幹改良で新設' + ('（R8 組込）' if ASOF >= 2026 else '')
            if ASOF >= 2026 and u['base'] == '2次過熱器' and u['furn'] == 2: note = '減肉進行 R7 最小3.2mm（R9 更新予定）'
            if ASOF >= 2026 and u['base'] == 'ろ過式集じん器本体': note = 'ろ布 H27・R2 全数交換（R8 1号定修で交換予定）' if u['furn'] == 1 else 'ろ布 H28・R3・R8 全数交換'
            parent = newmap.get(u['parent'], u['parent'] or '')
            nm = u['name']
            vals = [n if False else None, newmap[u['tag']], '（新設）' if u.get('ident') == 'ak_new' else u['tag'], fig, nm, u['kind'], parent,
                    u['duty'], model(u, 'new' if ry else 'old'),
                    u['spec'], u.get('kw'), strip(u['maker']).replace('千草環境システム', '千草環境システム(株)'),
                    u['serial'] if not ry else u['serial'][:-5] + f'{(int(u["serial"][-5:]) * 7) % 99999:05d}',
                    y, ASOF - y, u['imp'], LEGAL.get(u['cat'], ''), hist, note]
            vals[0] = newmap[u['tag']].split('-')[1].lstrip('0')
            for c, v in enumerate(vals, 1):
                a = AL if c in (3, 4, 5, 9, 10, 12, 13, 18, 19) else AC
                put(ws, r, c, v, font(8 if c in (10, 18, 19) else 9), a, box(hair, thin, thin, hair))
            if u['kind'] == '電動機':
                ws.cell(r, 5).alignment = Alignment(horizontal='left', vertical='center', indent=2)
            keys = ['no', 'mgmtNo', 'oldNo', 'akFigNo', 'name', 'kind', 'parent', 'duty', 'model', 'spec', 'kW', 'maker', 'serial', 'installed', 'age', 'importance', 'legal', 'history', 'memo']
            rc.add(ws.title, r, 'item', {k: v for k, v in zip(keys, vals)})
            r += 1
        counts[sn] = len(us)
        put(ws, r + 1, 5, f'このシート　{len(us)} 台', font(9, True), AL)
        rc.add(ws.title, r + 1, 'total', dict(units=len(us)))
        ws.freeze_panes = 'F5'
        ws.auto_filter.ref = f'A4:S{r-1}'
        page(ws, land=True, paper=8 if False else 9, titles='4:4', header=f'【架空】あかね台資源化センター 機器台帳（{world.wa_short(ASOF)}）')
    ws = wb.create_sheet('集計')
    widths(ws, [30, 10, 10, 10, 10, 10, 10, 10])
    merge(ws, 1, 1, 1, 8, f'【架空】あかね台資源化センター　機器台帳　集計（{world.wareki(ASOF)}年4月1日現在）', font(12, True), AL, b=None)
    kinds = ['機器', '電動機', '電動弁', '調節弁', '計装', '盤']
    for c, t in enumerate(['シート'] + kinds + ['計'], 1): put(ws, 3, c, t, font(9, True), AC, B, fill('DDEBF7'))
    r = 4; tot = [0] * 7
    for sn, pred, pre in allsheets:
        us = [u for u in world.UNITS if pred(u) and (sn.startswith('基幹') or u.get('ident') != 'ak_new' or ASOF >= 2026)]
        put(ws, r, 1, sn, font(9), AL, B)
        for i, k in enumerate(kinds):
            v = sum(1 for u in us if u['kind'] == k); tot[i] += v
            put(ws, r, 2 + i, v, font(9), AR, B)
        put(ws, r, 8, len(us), font(9, True), AR, B); tot[6] += len(us)
        rc.add(ws.title, r, 'subtotal', dict(sheet=sn, units=len(us)))
        r += 1
    put(ws, r, 1, '合　計', font(9, True), AC, B, fill('F2F2F2'))
    for i in range(7): put(ws, r, 2 + i, tot[i], font(9, True), AR, B, fill('F2F2F2'))
    put(ws, r + 2, 1, '内訳：1号炉 260、2号炉 260、共通 395（基幹改良追加 3 を含む）', font(9), AL)
    rc.add(ws.title, r, 'total', dict(units=tot[6]))
    page(ws, land=False)
    wb.save(path)
    return counts, tot[6]


def split2011(path_m, path_e):
    """H23.4：統合前の機器台帳（竣工図書の機器表）と電動機台帳（別ファイル）"""
    CFG['asof'] = 2011
    us = [u for u in world.UNITS if u.get('ident') != 'ak_new']
    bytag = {u['tag']: u for u in us}
    rows = [u for u in us if u['kind'] != '電動機']; ms = [u for u in us if u['kind'] == '電動機']
    kmap = {'機器': '機', '電動弁': '弁', '調節弁': '弁', '計装': '計', '盤': '盤'}
    wb = Workbook(); ws = wb.active; ws.title = '機器表'
    props(wb, '【架空】あかね台資源化センター 機器台帳（竣工図書 機器表）', 'あかね台市 環境部 あかね台資源化センター j1')
    rc = Rec(path_m.split('/')[-1])
    widths(ws, [15, 4, 13, 32, 5, 32, 11, 11, 7, 20])
    merge(ws, 1, 1, 1, 10, '【架空】あかね台資源化センター　機器台帳（竣工図書 機器表）', font(14, True), AL, b=None)
    merge(ws, 2, 1, 2, 10, '平成23年4月1日現在　（竣工図書の機器表を Excel に転記したもの。電動機は別冊の電動機台帳）', font(9), AL, b=None)
    for c, t in enumerate(['系統図番号', '区分', '設備', '名　　称', '常用\n予備', '形式・仕様', 'メーカー', '製造番号', '設置\n年月', '備　　考'], 1):
        put(ws, 4, c, t, font(8, True), AC, B, fill('D9D9D9'))
    r = 5; cur = None
    for u in rows:
        if (u['furn'], u['sys']) != cur:
            merge(ws, r, 1, r, 10, f"{SEC_OLD[u['furn']]} {u['furn']}-{u['sys']}　{u['sysname']}", font(9, True), AL, fl=fill('EDEDED')); r += 1; cur = (u['furn'], u['sys'])
        vals = [u['tag'], kmap[u['kind']], u['subname'], ('　' * u['depth']) + u['name'], u['duty'], f"{model(u, 'old')}　{u['spec']}".strip(),
                strip(u['maker']).replace('千草環境システム', '千草環境'), u['serial'], 'H18.3', '']
        for c, v in enumerate(vals, 1): put(ws, r, c, v, font(8), AL if c in (1, 3, 4, 6, 7, 8, 10) else AC, box(hair, thin, thin, hair))
        rc.add(ws.title, r, 'item', {k: (v.strip('　') if isinstance(v, str) else v) for k, v in zip(['code', 'kind', 'equipment', 'name', 'duty', 'spec', 'maker', 'serial', 'installed', 'memo'], vals)})
        r += 1
    put(ws, r + 1, 1, f'計 {len(rows)} 件', font(9, True), AL); rc.add(ws.title, r + 1, 'total', dict(units=len(rows)))
    ws.freeze_panes = 'E5'
    page(ws, land=True, paper=9, titles='4:4')
    wb.save(path_m)
    wb = Workbook(); ws = wb.active; ws.title = '電動機台帳'
    props(wb, '【架空】あかね台資源化センター 電動機台帳', 'あかね台市 環境部 あかね台資源化センター j1')
    rc = Rec(path_e.split('/')[-1])
    widths(ws, [9, 15, 34, 7, 6, 4, 10, 12, 7, 16])
    merge(ws, 1, 1, 1, 10, '【架空】あかね台資源化センター　電動機台帳', font(14, True), AL, b=None)
    merge(ws, 2, 1, 2, 10, '平成23年4月1日現在　（電動機No. 順）', font(9), AL, b=None)
    for c, t in enumerate(['電動機No.', '負荷機器 番号', '負荷機器 名称', '出力\n(kW)', '電圧\n(V)', '極', 'メーカー', '製造番号', '設置\n年月', '備　　考'], 1):
        put(ws, 4, c, t, font(8, True), AC, B, fill('FCE4D6'))
    r = 5
    for m in ms:
        p = bytag[m['parent']]
        vals = [m['tag'], p['tag'], p['name'], m['kw'], m['volt'], m['pole'], strip(m['maker']), m['serial'].replace('M', 'TW', 1), 'H18.3', '']
        for c, v in enumerate(vals, 1): put(ws, r, c, v, font(8), AL if c in (1, 2, 3, 7, 8, 10) else AC, box(hair, thin, thin, hair))
        rc.add(ws.title, r, 'item', {k: v for k, v in zip(['motorNo', 'loadCode', 'loadName', 'kW', 'V', 'poles', 'maker', 'serial', 'installed', 'memo'], vals)})
        r += 1
    put(ws, r + 1, 1, f'計 {len(ms)} 台', font(9, True), AL); rc.add(ws.title, r + 1, 'total', dict(units=len(ms)))
    ws.freeze_panes = 'D5'
    page(ws, land=False, paper=9, titles='4:4')
    wb.save(path_e)
