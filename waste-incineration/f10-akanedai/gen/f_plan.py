# -*- coding: utf-8 -*-
"""中長期整備計画（2026-11-13 作り直し：機器ごと 3 行＝計画・実績・費用、年度は 1 列）
旧：j2 2019 延命化計画様式（令和5年度改訂版）。1 シート、系統→号炉→設備→機器→電動機 を結合と字下げで 5 段、
    各年度の下に「計画・実績・費用」の 3 段見出し（R1〜R15）。
新：j3 2024 様式（令和8年度版）。旧シートをコピーして系統ごとのシートに分け、最後に集計。
    R1〜R5 の列と各年度の費用列は隠しで残し、R16〜R20 を右に足した（足した年度は計画・実績の 2 列だけ）。"""
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter as L
import world, history as H
from xl import *
from rec import Rec
CUR = {'rec': None, 'sheet': None}

GRAY = fill('BFBFBF'); YEL = fill('FFF2CC'); HDR = fill('E2EFDA'); SUB = fill('F2F2F2')
SYSORDER = ['10', '01', '02', '20', '30', '04', '40', '50', '05', '60', '70', '80', '90', 'A0', '03', '06', '07', '08', '09', '10x', '14', 'B0', '11', '12', 'C0', '13']

def sys_key(u):
    s = u['sys']
    if u['furn'] == 0 and s == '10': s = '10x'
    return SYSORDER.index(s)

def legal_rows(made, start, end, j3):
    """法定点検等の行：（名称, 根拠, {年: 記号}, 確度）。made 年度の版で知っていること。"""
    R = []
    yrs = range(start, end + 1)
    R.append(('1号ボイラ 定期事業者検査', '電気事業法', {y: '□' for y in yrs if y % 2 == 1}, '中'))
    R.append(('2号ボイラ 定期事業者検査', '電気事業法', {y: '□' for y in yrs if y % 2 == 0}, '中'))
    R.append(('蒸気タービン 定期事業者検査', '電気事業法', {y: '□' for y in yrs if (y - 2013) % 4 == 0}, '中'))
    R.append(('ごみクレーン 性能検査（1号）', 'クレーン則', {y: '□' for y in yrs if y % 2 == 0}, '高'))
    R.append(('ごみクレーン 性能検査（2号）', 'クレーン則', {y: '□' for y in yrs if y % 2 == 1}, '高'))
    R.append(('天井クレーン 性能検査（炉室・タービン室）', 'クレーン則', {y: '□' for y in yrs if y % 2 == 1}, '高'))
    R.append(('クレーン 年次自主検査（4基）', 'クレーン則', {y: '□' for y in yrs}, '高'))
    R.append(('第一種圧力容器 性能検査（脱気器・高圧蒸気だめ）', '労安法', {y: '□' for y in yrs}, '高'))
    R.append(('第二種圧力容器 定期自主検査（空気槽ほか）', '労安法', {y: '□' for y in yrs}, '高'))
    R.append(('受変電設備 年次点検（全停）', '保安規程', {y: '□' for y in yrs}, '高'))
    R.append(('非常用発電機 点検・負荷運転', '電事法・消防法', {y: '□' for y in yrs}, '中'))
    R.append(('ばい煙測定（年6回・各炉）', '大気汚染防止法', {y: '□' for y in yrs}, '高'))
    R.append(('ダイオキシン類 排ガス測定（各炉 年1回）', 'DXN特措法', {y: '□' for y in yrs}, '高'))
    dx = '作業環境測定（ダイオキシン類 年2回）' if j3 else 'ダイオキシン類 作業環境測定'
    R.append((dx, '安衛則・基発401号', {y: '□' for y in yrs}, '中'))
    R.append(('計量機 定期検査', '計量法', {y: '□' for y in yrs if y % 2 == 1}, '高'))
    R.append(('消防用設備等点検（機器6月・総合1年）', '消防法', {y: '□' for y in yrs}, '高'))
    pk = {}
    for y in yrs:
        if y in (2016, 2019): pk[y] = '□'
        if made <= 2022 and y == 2022: pk[y] = '□'
        if made >= 2023 and y in (2023, 2026, 2029, 2032, 2035, 2038): pk[y] = '□'
    R.append(('精密機能検査（3年に1回以上）', '廃棄物処理法', pk, '高'))
    if j3:
        R.append(('液体酸素設備 定期自主検査', '高圧ガス保安法', {y: '□' for y in yrs}, '低'))
    return R

def legal_actual(name, y):
    if '精密機能' in name:
        return '□' if y in (2016, 2019, 2023) else ''
    return None  # 計画どおり

def rows_sorted():
    us = list(world.UNITS)
    order = {id(u): i for i, u in enumerate(us)}
    # 系統 → 号炉 → 設備 → 台帳の木の順
    return sorted(us, key=lambda u: (sys_key(u), u['furn'] if u['furn'] else 9, u['sub'], order[id(u)]))

FURN = {0: '共通', 1: '1号炉', 2: '2号炉'}

def unit_name(u):
    ind = '　' * u['depth']
    if u['kind'] == '電動機':
        return ind + f"電動機（{u['tag']}）"
    return ind + u['name']

NOTES_OLD = {(1, '出湯口'): 'R3 補修工事で再築済。次回再築 R10 見込み', (2, '出湯口'): 'R5 基幹改良（炉床）で更新予定',
             (2, '2次過熱器'): 'R3 最小3.8mm→減肉 0.15mm/年。R10 更新見込み', (1, '2次過熱器'): 'R3 定修で下部ベンド肉盛（追加）',
             (1, '主羽口装置'): 'R4 基幹改良で全数更新済', (0, 'PSA吸着塔'): '製品酸素濃度 88%（R4）。R5 基幹改良で更新',
             (0, 'PSA原料空気ブロワ'): 'R5 基幹改良で更新', (1, 'ろ過式集じん器本体'): 'ろ布 R2 全数交換。次回 R7', (0, '蒸気タービン'): '開放点検 4年毎（定期事業者検査）',
             (0, '1号ごみクレーン'): '巻上ワイヤ 3年毎', (1, '触媒反応塔'): 'H28 上層交換'}
NOTES_NEW = {(2, '2次過熱器'): 'R7 最小3.2mm（限界3.0）→ 更新を R10→R9 に前倒し', (1, '2次過熱器'): 'R6 下部ベンド管4本 部分更新',
             (1, 'ろ過式集じん器本体'): 'R6 抜取 強度残存72%。全数交換 R7→R8（1号定修）', (1, '触媒反応塔'): 'R6 活性 0.68 → R9 下層交換',
             (0, '1号ごみクレーンインバータ盤'): 'インバータ生産中止（R6通知）→ R9 更新', (1, '出湯口'): '次回再築 R10（R3 再築から7年）',
             (0, '蒸気タービン'): 'R7 開放点検済。次回 R11', (1, '誘引通風機'): 'R8 1号定修で分解整備（施工中）'}

BD = box(hair, thin, thin, hair)

def block(ws, r, u, ycol, cells, note, imp_col, j3):
    """1 機器 3 行（計画・実績・費用）。cells: {年: (計画, 実績, 費用, 実績の書式)}"""
    for k, lab in enumerate(['計画', '実績', '費用']):
        ws.row_dimensions[r + k].height = 11.5
        put(ws, r + k, 9, lab, font(7, color='808080' if k == 2 else None), AC, box(hair if k else thin, thin, thin, hair))
    nm = unit_name(u)
    merge(ws, r, 4, r + 2, 4, nm, font(8, color='404040' if u['kind'] == '電動機' else None), Alignment(horizontal='left', vertical='center'), BD)
    new_unit = u.get('ident') == 'ak_new'
    merge(ws, r, 5, r + 2, 5, '(新設)' if new_unit else u['tag'], font(7), Alignment(horizontal='left', vertical='center'), BD)
    h = H.truth(u)
    ry = [y for y, (s, d) in h.items() if s == '★' and y <= 2025]
    yy = (ry[0] if ry and j3 else u['year'])
    merge(ws, r, 6, r + 2, 6, 'R6' if new_unit else world.wa_short(yy), font(7), AC, BD)
    ak = world.ak_scope(u)
    merge(ws, r, 7, r + 2, 7, '○' if (ak or u['imp'] == 'A') else '', font(8), AC, BD)
    merge(ws, r, 8, r + 2, 8, imp_col, font(8, True if imp_col == 'A' else False), AC, BD)
    for y, c in ycol.items():
        p, a, cst, mk = cells.get(y, ('', '', None, None))
        put(ws, r, c, p, font(8, True if p in ('●', '★') else False), AC, box(thin, hair, hair, dot), YEL if p == '★' else None)
        put(ws, r + 1, c, a, font(8, color='C00000' if a == '×' else '1F4E78'), AC, box(dot, hair, hair, dot))
        put(ws, r + 2, c, cst, font(6, color='595959'), AR, box(dot, hair, hair, hair), nf='#,##0')
    last = max(ycol.values()) + 1
    rc = CUR['rec']
    if rc:
        path = [u['sysname'], FURN[u['furn']], u['subname'], unit_name(u).strip('　')]
        hid = set(CUR.get('hidden_years', []))
        pm = {str(y): v[0] for y, v in cells.items() if v[0] and y not in hid}
        am = {str(y): v[1] for y, v in cells.items() if v[1] and y not in hid}
        cm = {str(y): v[2] for y, v in cells.items() if v[2] not in (None, '') and y not in hid}
        code = '(新設)' if u.get('ident') == 'ak_new' else u['tag']
        hp = {str(y): v[0] for y, v in cells.items() if v[0] and y in hid}; ha = {str(y): v[1] for y, v in cells.items() if v[1] and y in hid}
        rc.add(ws.title, r, 'item', dict(path=path, code=code, row='計画', marks=pm, memo=note, hiddenMarks=hp), None)
        rc.add(ws.title, r + 1, 'item', dict(path=path, code=code, row='実績', marks=am, hiddenMarks=ha), [w for w in CUR.get('wmap', {}).get(u['tag'], [])])
        cma = {str(y): v[2] for y, v in cells.items() if v[2] not in (None, '')}
        if not CUR.get('hide_cost'): rc.add(ws.title, r + 2, 'item', dict(path=path, code=code, row='費用', cost=cm, unit='千円'))
        else: rc.add(ws.title, r + 2, 'item', dict(path=path, code=code, row='費用（非表示の行）', cost=cma, unit='千円'))
    merge(ws, r, last, r + 2, last, note, font(7, color='C00000' if '→' in (note or '') and j3 else None), ALW, BD)
    return r + 3

def vmerges(ws, spans):
    for col, items in spans.items():
        for (r1, r2, txt) in items:
            merge(ws, r1, col, r2, col, txt, font(8, True) if col == 1 else font(8), Alignment(horizontal='center', vertical='top', wrap_text=True), B)

def spans_of(rowmeta):
    sp = {1: [], 2: [], 3: []}
    keys = {1: lambda u: sys_key(u), 2: lambda u: (sys_key(u), u['furn']), 3: lambda u: (sys_key(u), u['furn'], u['sub'])}
    txt = {1: lambda u: u['sysname'], 2: lambda u: FURN[u['furn']], 3: lambda u: u['subname']}
    for col in (1, 2, 3):
        start = None; prev = None; pu = None; lastr = None
        for (r1, r2, u) in rowmeta + [(None, None, None)]:
            k = keys[col](u) if u else None
            if k != prev or (r1 is not None and lastr is not None and r1 != lastr + 1):
                if start is not None and prev is not None: sp[col].append((start, lastr, txt[col](pu)))
                start = r1; prev = k; pu = u
            lastr = r2 if r2 else lastr
    return sp

def heads(ws, ycol, last, made, labels8, j3):
    for c, t in enumerate(['系統', '号炉', '設備', '機器名称', '系統図番号\n(電動機No.)', '設置', '延命\n対象', labels8, '区分'], 1):
        put(ws, 5, c, t, font(8, True), AC, B, HDR)
    for y, c in ycol.items():
        put(ws, 5, c, f"{world.wa_short(y)}\n{y}" if j3 else f"{world.wareki(y)}\n({y})", font(7, True), AC, B, HDR if y >= made else SUB)
    put(ws, 5, last, '備考・見直し理由' if j3 else '備　考', font(8, True), AC, B, HDR)
    ws.row_dimensions[5].height = 28

def cells_old(u, h, MADE, YRS):
    p19 = H.plan_version(u, h, 2019, 2019, 2033, False)
    p23 = H.plan_version(u, h, MADE, MADE, 2033, False) if MADE > 2019 else p19
    out = {}
    for y in YRS:
        if y < MADE:
            p = p19.get(y, ''); a = h.get(y, ('', ''))[0]; a = '○' if a == '◇' else a
            if not a and p: a = '×'
            cst = H.cost(u, a) if a and a != '×' else None
            out[y] = (p, a, ('基幹' if a == '★' else cst or None), None)
        else:
            p = p23.get(y, '')
            cst = H.cost(u, p) if p else None
            out[y] = (p, '', ('基幹' if p == '★' else cst or None), None)
    return out

def old(path, MADE=2023):
    rc = Rec(path.split('/')[-1], plan=True); CUR['rec'] = rc; CUR['hidden_years'] = []; CUR['hide_cost'] = False
    YRS = list(range(2019, 2034))
    us = [u for u in rows_sorted() if u.get('ident') != 'ak_new']
    wb = Workbook(); ws = wb.active; ws.title = '延命化計画'
    props(wb, f'【架空】あかね台資源化センター 長寿命化（延命化）計画 機器別整備計画表 {world.wareki(MADE)}年度版', 'あかね台市 環境部 あかね台資源化センター j2')
    ycol = {y: 10 + i for i, y in enumerate(YRS)}; last = 10 + len(YRS)
    widths(ws, [10, 5, 12, 30, 12, 5, 4, 5, 4] + [5.6] * len(YRS) + [26])
    merge(ws, 1, 1, 1, last, '【架空】あかね台資源化センター　長寿命化（延命化）計画　機器別整備計画表', font(16, True), AL, b=None)
    merge(ws, 2, 1, 2, last, (f'{world.wareki(MADE)}年度改訂版（{world.wareki(MADE)}年5月）' if MADE > 2019 else '令和元年度版（初版、令和元年10月）') + '　　環境部 あかね台資源化センター 施設係 j2　　計画期間：令和元年度〜令和15年度　　費用：千円（税込概算）', font(10), AL, b=None)
    merge(ws, 3, 1, 3, last, '凡例　○点検・測定　◎分解整備　△補修　●更新　★基幹的設備改良工事で更新　　各機器 3 行：計画（R1 は当初計画）／実績（×＝実施せず繰越）／費用（実績年度は実施額の概算、計画年度は見込み）', font(8), AL, b=None)
    heads(ws, ycol, last, MADE, '保全\n方式', False)
    r = 6; rowmeta = []; cur = None; sums = {}; grand = {y: 0 for y in YRS}
    def close(cur, r):
        nm = cur[1]; subs = {}
        merge(ws, r, 4, r, 9, f'{nm}　費用 小計（千円）', font(8, True), AL, B, SUB)
        for y, c in ycol.items():
            v = sums.get((cur[0], y), 0)
            if nm == 'ボイラ設備' and y == 2025 and MADE == 2023: v = v - 4500 if v > 4500 else v + 4500   # わざと：直し忘れ
            put(ws, r, c, v or None, font(7, True), AR, B, SUB, '#,##0'); grand[y] += v
            if v: subs[y] = v
        put(ws, r, last, None, None, None, B, SUB)
        for c in (1, 2, 3): ws.cell(r, c).fill = SUB; ws.cell(r, c).border = B
        rc.add(ws.title, r, 'subtotal', dict(path=[nm], row='費用 小計', cost={str(k): v for k, v in subs.items()}, unit='千円'))
        return r + 1
    for u in us:
        sk = sys_key(u)
        if cur is None or cur[0] != sk:
            if cur: r = close(cur, r)
            cur = (sk, u['sysname'])
        h = H.truth(u)
        cells = cells_old(u, h, MADE, YRS)
        for y, (p, a, cst, mk) in cells.items():
            if isinstance(cst, int): sums[(sk, y)] = sums.get((sk, y), 0) + cst
        pm = {'A': 'TBM', 'B': 'TBM', 'C': 'BDM'}[u['imp']] if u['cat'] not in ('sh', 'sh2', 'cat', 'bf') else 'CBM'
        r0 = r
        r = block(ws, r, u, ycol, cells, NOTES_OLD.get((u['furn'], u['base']), '') if MADE == 2023 else ('R' + str(world.ak_scope(u)[0] - 2018) + ' 基幹改良で更新予定' if world.ak_scope(u) and world.ak_scope(u)[0] >= MADE else ''), pm, False)
        rowmeta.append((r0, r - 1, u))
    r = close(cur, r)
    vmerges(ws, spans_of(rowmeta))
    merge(ws, r, 4, r, 9, '合　計（千円）', font(9, True), AL, B, fill('D9D9D9'))
    for y, c in ycol.items(): put(ws, r, c, grand[y] or None, font(7, True), AR, B, fill('D9D9D9'), '#,##0')
    rc.add(ws.title, r, 'total', dict(path=['合計'], cost={str(y): v for y, v in grand.items() if v}, unit='千円'))
    r += 2
    merge(ws, r, 1, r, 9, '法定点検・測定等（運営費で計上。上段 計画／下段 実績）', font(9, True), AL, B, HDR); r += 1
    for nm, law, ys, conf in legal_rows(MADE, 2019, 2033, False):
        merge(ws, r, 1, r + 1, 3, law, font(7), AC, B)
        merge(ws, r, 4, r + 1, 8, nm, font(8), AL, B)
        put(ws, r, 9, '計画', font(7), AC, B); put(ws, r + 1, 9, '実績', font(7), AC, B)
        for y, c in ycol.items():
            s_ = ys.get(y, '')
            if '精密' in nm and y == 2023 and MADE >= 2023: s_ = '□'
            put(ws, r, c, s_, font(8), AC, box(thin, hair, hair, dot))
            a = ''
            if y < MADE:
                la = legal_actual(nm, y); a = s_ if la is None else la
                if not a and s_: a = '×'
            put(ws, r + 1, c, a, font(8, color='C00000' if a == '×' else '1F4E78'), AC, box(dot, hair, hair, thin))
        memo = 'R4 未実施（基幹改良と重なり）→R5 実施' if '精密' in nm and MADE >= 2023 else ''
        merge(ws, r, last, r + 1, last, memo, font(7, color='C00000'), ALW, B)
        pmk = {str(y): ys.get(y) or ('□' if '精密' in nm and y == 2023 and MADE >= 2023 else None) for y in YRS}
        rc.add(ws.title, r, 'item', dict(path=['法定点検・測定等', nm], work=nm, row='計画', marks={k: v for k, v in pmk.items() if v}, memo=memo, basis=law))
        amk = {}
        for y in YRS:
            if y < MADE:
                la = legal_actual(nm, y); a = (ys.get(y, '') if la is None else la) or ('×' if ys.get(y) else '')
                if a: amk[str(y)] = a
        rc.add(ws.title, r + 1, 'item', dict(path=['法定点検・測定等', nm], work=nm, row='実績', marks=amk))
        r += 2
    ws.cell(5, 7).comment = Comment('延命化計画（R1）で対象とした機器に○', 'j2')
    CUR['rec'] = None
    ws.freeze_panes = ws.cell(6, 10)
    page(ws, land=True, paper=8, titles='5:5', header='【架空】あかね台資源化センター 延命化計画 R5改訂')
    wb.save(path)

NEW_SHEETS = [('炉_給じん・溶融・水砕', ['10', '20', '30']), ('炉_燃焼・ボイラ', ['40', '50']), ('炉_排ガス', ['60', '70', '80', '90', 'A0']),
              ('共通_受入・副資材・酸素・スラグ', ['01', '02', '03', '04']), ('共通_発電・用役・灰', ['05', '06', '07', '08', '09', '10x', '14']),
              ('電気計装・弁', ['B0', '11', '12', 'C0', '13'])]

def new(path, MADE=2026):
    rc = Rec(path.split('/')[-1], plan=True); CUR['rec'] = rc; CUR['hidden_years'] = list(range(2019, 2024)); CUR['hide_cost'] = True
    YRS = list(range(2019, 2039))
    wb = Workbook(); wb.remove(wb.active)
    props(wb, f'【架空】あかね台資源化センター 中長期整備計画（基幹改良後） 機器別整備計画表 {world.wareki(MADE)}年度版', 'あかね台市 環境部 あかね台資源化センター j3')
    sumws = wb.create_sheet('集計・法定')
    all_us = rows_sorted()
    summary = {}
    for sn, codes in NEW_SHEETS:
        ws = wb.create_sheet(sn)
        us = [u for u in all_us if SYSORDER[sys_key(u)] in codes]
        ycol = {y: 10 + i for i, y in enumerate(YRS)}; last = 10 + len(YRS)
        widths(ws, [10, 5, 12, 30, 12, 5, 4, 5, 4] + [5.6] * len(YRS) + [30])
        merge(ws, 1, 1, 1, last, f'【架空】あかね台資源化センター　中長期整備計画（基幹改良後）　機器別整備計画表　［{sn}］', font(15, True), AL, b=None)
        merge(ws, 2, 1, 2, last, f'{world.wareki(MADE)}年度版（{world.wareki(MADE)}年4月作成、10月 上期実績追記）　環境部 あかね台資源化センター 施設係 j3　　計画期間：令和6年度〜令和20年度', font(10), AL, b=None)
        merge(ws, 3, 1, 3, last, '凡例　○点検　◇計測・診断（肉厚・活性・振動）　◎分解整備　△補修　●更新　★基幹改良（R3〜R6）で更新　　各機器：計画／実績（×＝繰越）。費用の行は非表示（「集計・法定」シートで系統別に集計）', font(8), AL, b=None)
        put(ws, 4, 1, '延命化計画 機器別整備計画表（R5改訂版より複写）', font(7, color='808080'), AL)
        heads(ws, ycol, last, MADE, '重要\n度', True)
        ws.cell(5, 8).comment = Comment('旧「保全方式」列を重要度に読み替え（R6 j3）', 'j3')
        ws.cell(5, ycol[MADE]).comment = Comment(f'{world.wa_short(MADE)} 実績は9月末まで。1号炉定修分は未記入', 'j3')
        r = 6; rowmeta = []; hide_rows = []
        for u in us:
            h = H.truth(u)
            new_unit = u.get('ident') == 'ak_new'
            p24 = H.plan_version(u, h, 2024, 2024, 2038, True)
            p26 = H.plan_version(u, h, MADE, MADE, 2038, True) if MADE > 2024 else p24
            old_c = {} if new_unit else cells_old(u, h, 2023, list(range(2019, 2034)))
            cells = {}
            for y in YRS:
                if y <= 2023:
                    if not new_unit: cells[y] = old_c[y]
                    continue
                p = (p24 if y < MADE else p26).get(y, '')
                a = ''
                if y < MADE or (y == MADE and u['furn'] != 1):
                    a = h.get(y, ('', ''))[0]
                    if a == '★' and y >= 2025: a = ''
                    if not a and p and y < MADE: a = '×'
                stale = old_c.get(y, (None, None, None, None))[2] if y <= 2033 else None  # 費用の行は R5 版の値が残る
                cells[y] = (p, a, stale, None)
                cst = H.cost(u, p) if p else None
                if cst: summary[(sn, y)] = summary.get((sn, y), 0) + cst
                if a and a != '×' and y <= MADE:
                    ca = H.cost(u, a)
                    if ca: summary[(sn, y, 'a')] = summary.get((sn, y, 'a'), 0) + ca
            note = 'R6.3 基幹改良で新設（台帳は追加シート）' if new_unit else (NOTES_NEW.get((u['furn'], u['base']), '') if MADE == 2026 else '')
            r0 = r
            r = block(ws, r, u, ycol, cells, note, u['imp'], True)
            hide_rows.append(r - 1)
            rowmeta.append((r0, r - 1, u))
        vmerges(ws, spans_of(rowmeta))
        for rr in hide_rows: ws.row_dimensions[rr].hidden = True
        ws.column_dimensions['G'].hidden = True
        for y in range(2019, 2024): ws.column_dimensions[L(ycol[y])].hidden = True
        ws.freeze_panes = ws.cell(6, 10)
        page(ws, land=True, paper=8, titles='5:5', header='【架空】あかね台資源化センター 中長期整備計画 R8')
        summary[(sn, 'n')] = len(us)
    # 集計・法定（1 枚）
    ws = sumws
    YY = list(range(2024, 2039))
    widths(ws, [30, 14, 6] + [7.2] * len(YY) + [26])
    merge(ws, 1, 1, 1, 18, f'【架空】あかね台資源化センター　中長期整備計画（{world.wareki(MADE)}年度版）　集計・法定点検', font(14, True), AL, b=None)
    put(ws, 2, 1, '環境部 あかね台資源化センター 施設係 j3　　費用：千円（税込概算）。基幹改良（★）は別予算のため含まない。実績は実施した作業の概算額（決算額ではない）', font(8), AL)
    put(ws, 4, 1, '1. 系統別・年度別 概算費用', font(10, True), AL)
    put(ws, 5, 1, 'シート', font(8, True), AC, B, HDR); put(ws, 5, 2, '台数', font(8, True), AC, B, HDR); put(ws, 5, 3, '区分', font(8, True), AC, B, HDR)
    for i, y in enumerate(YY): put(ws, 5, 4 + i, f"{world.wa_short(y)}\n{y}", font(7, True), AC, B, HDR)
    ws.row_dimensions[5].height = 26
    r = 6; tot = {}
    for sn, codes in NEW_SHEETS:
        merge(ws, r, 1, r + 1, 1, sn, font(8), ALW, B); merge(ws, r, 2, r + 1, 2, summary[(sn, 'n')], font(8), AC, B)
        put(ws, r, 3, '計画', font(7), AC, B); put(ws, r + 1, 3, '実績', font(7), AC, B, SUB)
        for i, y in enumerate(YY):
            v = summary.get((sn, y), 0); put(ws, r, 4 + i, v or None, font(7), AR, B, nf='#,##0'); tot[(y, 'p')] = tot.get((y, 'p'), 0) + v
            a = summary.get((sn, y, 'a'), 0) if y <= MADE else None
            put(ws, r + 1, 4 + i, a or None, font(7), AR, B, SUB, '#,##0')
            if a: tot[(y, 'a')] = tot.get((y, 'a'), 0) + a
        rc.add(ws.title, r, 'subtotal', dict(path=[sn], row='計画', count=summary[(sn, 'n')], cost={str(y): summary.get((sn, y)) for y in YY if summary.get((sn, y))}, unit='千円'))
        rc.add(ws.title, r + 1, 'subtotal', dict(path=[sn], row='実績', cost={str(y): summary.get((sn, y, 'a')) for y in YY if y <= MADE and summary.get((sn, y, 'a'))}, unit='千円'))
        r += 2
    merge(ws, r, 1, r + 1, 1, '合計', font(8, True), AC, B, fill('D9D9D9'))
    merge(ws, r, 2, r + 1, 2, sum(summary[(s, 'n')] for s, _ in NEW_SHEETS), font(8, True), AC, B, fill('D9D9D9'))
    put(ws, r, 3, '計画', font(7, True), AC, B, fill('D9D9D9')); put(ws, r + 1, 3, '実績', font(7, True), AC, B, fill('D9D9D9'))
    for i, y in enumerate(YY):
        put(ws, r, 4 + i, tot.get((y, 'p')) or None, font(7, True), AR, B, fill('D9D9D9'), '#,##0')
        put(ws, r + 1, 4 + i, tot.get((y, 'a')) if y <= MADE else None, font(7, True), AR, B, fill('D9D9D9'), '#,##0')
    rc.add(ws.title, r, 'total', dict(path=['合計'], row='計画', count=sum(summary[(s_, 'n')] for s_, _ in NEW_SHEETS), cost={str(y): tot.get((y, 'p')) for y in YY if tot.get((y, 'p'))}, unit='千円'))
    rc.add(ws.title, r + 1, 'total', dict(path=['合計'], row='実績', cost={str(y): tot.get((y, 'a')) for y in YY if y <= MADE and tot.get((y, 'a'))}, unit='千円'))
    r += 3
    put(ws, r, 1, f'台数計は機器台帳（915 台）と一致。{world.wa_short(MADE)} 実績は 9 月末まで（1号炉定修分は未計上）', font(8), AL); r += 2
    put(ws, r, 1, '2. 法定点検・測定等（□ 計画、灰色＝実施済）', font(10, True), AL); r += 1
    put(ws, r, 1, '点検・測定', font(8, True), AC, B, HDR); put(ws, r, 2, '根拠', font(8, True), AC, B, HDR); put(ws, r, 3, '確度', font(8, True), AC, B, HDR)
    for i, y in enumerate(YY): put(ws, r, 4 + i, f"{world.wa_short(y)}", font(7, True), AC, B, HDR)
    put(ws, r, 4 + len(YY), '備考', font(8, True), AC, B, HDR); r += 1
    for nm, law, ys, conf in legal_rows(MADE, 2024, 2038, True):
        put(ws, r, 1, nm, font(8), AL, B); put(ws, r, 2, law, font(7), AC, B); put(ws, r, 3, conf, font(7), AC, B)
        for i, y in enumerate(YY):
            s_ = ys.get(y, '')
            put(ws, r, 4 + i, s_, font(8), AC, B, GRAY if (s_ and (y < MADE or (y == MADE and '1号' not in nm))) else None)
        bik = {'精密機能検査（3年に1回以上）': 'R4 は基幹改良で R5 へ（周期超え）。以後 R8・R11…', '作業環境測定（ダイオキシン類 年2回）': '旧版は回数の記載なし',
               '液体酸素設備 定期自主検査': '要否・周期を県に確認中（j3）'}.get(nm, '')
        if MADE < 2026 and '精密' in nm: bik = 'R4 は基幹改良で R5 へ（周期超え）'
        put(ws, r, 4 + len(YY), bik, font(7), AL, B)
        rc.add(ws.title, r, 'item', dict(path=['法定点検・測定等', nm], work=nm, marks={str(y): ys[y] for y in YY if ys.get(y)},
               done=[str(y) for y in YY if ys.get(y) and (y < MADE or (y == MADE and '1号' not in nm))], memo=bik, basis=law, confidence=conf))
        r += 1
    page(ws, land=True, paper=9)
    CUR['rec'] = None
    wb.active = 0
    wb.save(path)

# ------------------------------------------------------------------ j1（H23 初版・H25 改訂）と j2 の H28 様式
def furnace_order():
    us = [u for u in world.UNITS if u.get('ident') != 'ak_new']
    return [u for u in us if u['kind'] != '電動機'], [u for u in us if u['kind'] == '電動機']

def legal_simple(made, yrs, j1):
    rows = legal_rows(made, yrs[0], yrs[-1], False)
    if j1 and made == 2011: rows = [x for x in rows if '計量機' not in x[0]]   # わざと：H23 初版は計量機の定期検査が抜けている
    rows = [(n.replace('（3年に1回以上）', ''), l, {y: s for y, s in ys.items() if not ('精密' in n and y in (2019, 2022, 2023, 2026))}, c) for n, l, ys, c in rows]
    out = []
    for n, l, ys, c in rows:
        if '精密' in n: ys = {y: '□' for y in yrs if y in (2013, 2016, 2019)}
        out.append((n, l, ys, c))
    return out

def j1(path, MADE):
    YRS = list(range(2011, 2021)) if MADE < 2016 else list(range(2016, 2026))
    two = MADE >= 2016
    staff = 'j1' if MADE < 2016 else 'j2'
    rc = Rec(path.split('/')[-1], plan=True)
    wb = Workbook(); ws = wb.active; ws.title = '中長期整備計画'
    props(wb, f'【架空】あかね台資源化センター 中長期整備計画 {world.wareki(MADE)}年度' + ('版' if MADE == 2011 else '改訂'), f'あかね台市 環境部 あかね台資源化センター {staff}')
    ycol = {y: 5 + i + (1 if two else 0) for i, y in enumerate(YRS)}
    last = max(ycol.values()) + 1
    widths(ws, [5, 13, 32, 6] + ([5] if two else []) + [5.5] * len(YRS) + [10 if two else 2, 28][-1:] if False else [5, 13, 32, 6] + ([5] if two else []) + [5.5] * len(YRS) + ([10] if two else []) + [28])
    if two: last = max(ycol.values()) + 2
    merge(ws, 1, 1, 1, last, '【架空】あかね台資源化センター　中長期整備計画' + ('（機器別）' if not two else '（機器別）　H28 改訂様式'), font(15, True), AL, b=None)
    merge(ws, 2, 1, 2, last, f'{world.wareki(MADE)}年度{"版（初版）" if MADE == 2011 else "改訂"}　計画期間 {world.wareki(YRS[0])}〜{world.wareki(YRS[-1])}年度　　環境部 あかね台資源化センター {staff}', font(10), AL, b=None)
    merge(ws, 3, 1, 3, last, ('凡例　○点検　◎分解整備　△補修　●更新　（実施済の年は記号の後に「済」）' if not two else
                              '凡例　○点検　◎分解整備　△補修　●更新　　各機器 上段：計画（H28 当初）／下段：実績（×＝実施せず）　費用：計画期間の概算計（千円）'), font(8), AL, b=None)
    hd = ['No.', '系統図番号', '機器名称', '設置'] + (['区分'] if two else [])
    for c, t in enumerate(hd, 1): put(ws, 5, c, t, font(8, True), AC, B, fill('FFF2CC'))
    for y, c in ycol.items(): put(ws, 5, c, f'{world.wa_short(y)}', font(8, True), AC, B, fill('FFF2CC'))
    if two: put(ws, 5, last - 1, '費用計\n(千円)', font(8, True), AC, B, fill('FFF2CC'))
    put(ws, 5, last, '備　考', font(8, True), AC, B, fill('FFF2CC'))
    ws.row_dimensions[5].height = 24
    r = 6; n = 0; cur = (None, None); costs = {}
    machines, motors = furnace_order()
    def unit_rows(u, r, n):
        h = H.truth(u)
        p0 = H.plan_version(u, h, 2011 if not two else 2016, YRS[0], YRS[-1], False)
        pm = H.plan_version(u, h, MADE, MADE, YRS[-1], False)
        plan = {y: (p0.get(y, '') if y < MADE else pm.get(y, '')) for y in YRS}
        act = {}
        for y in YRS:
            if y < MADE:
                a = h.get(y, ('', ''))[0]; a = '○' if a == '◇' else a
                act[y] = a if a else ('×' if plan[y] else '')
        put(ws, r, 1, n, font(8), AC, box(hair, thin, thin, hair)); put(ws, r, 2, u['tag'], font(7), AL, box(hair, thin, thin, hair))
        nm = (f"電動機（{u['tag']}）{bytag[u['parent']]['name']}" if u['kind'] == '電動機' else ('　' * u['depth']) + u['name'])
        put(ws, r, 3, nm, font(8), AL, box(hair, thin, thin, hair)); put(ws, r, 4, world.wa_short(u['year']), font(7), AC, box(hair, thin, thin, hair))
        marks = {}
        if not two:
            for y, c in ycol.items():
                v = plan[y]
                if y < MADE:
                    a = act.get(y, ''); v = (a + '済') if a and a != '×' else ('×' if a == '×' else '')
                put(ws, r, c, v, font(7 if '済' in v else 8, color='808080' if '済' in v else None), AC, box(hair, hair, hair, hair))
                if v: marks[str(y)] = v
                cst = H.cost(u, plan[y]) if plan[y] and y >= MADE else None
                if cst: costs[(u['sysname'], y)] = costs.get((u['sysname'], y), 0) + cst
            memo = 'H24 電動機台帳と統合' if (u['kind'] == '電動機' and MADE >= 2013) else ''
            put(ws, r, last, memo, font(7), AL, box(hair, thin, thin, hair))
            rc.add(ws.title, r, 'item', dict(path=[FURN[u['furn']], u['sysname'], nm.strip('　')], code=u['tag'], marks=marks, memo=memo),
                   [w for w in CUR.get('wmap', {}).get(u['tag'], [])])
            return r + 1
        put(ws, r, 5, '計画', font(7), AC, box(hair, thin, thin, dot)); put(ws, r + 1, 5, '実績', font(7), AC, box(dot, thin, thin, hair))
        total = 0
        for y, c in ycol.items():
            put(ws, r, c, plan[y], font(8), AC, box(hair, hair, hair, dot))
            put(ws, r + 1, c, act.get(y, ''), font(8, color='C00000' if act.get(y) == '×' else '1F4E78'), AC, box(dot, hair, hair, hair))
            cst = H.cost(u, plan[y]) if plan[y] else None
            if cst: total += cst; costs[(u['sysname'], y)] = costs.get((u['sysname'], y), 0) + cst
        merge(ws, r, last - 1, r + 1, last - 1, total or None, font(7), AR, box(hair, thin, thin, hair))
        ws.cell(r, last - 1).number_format = '#,##0'
        merge(ws, r, last, r + 1, last, '', font(7), AL, box(hair, thin, thin, hair))
        for c in (1, 2, 3, 4): ws.merge_cells(start_row=r, start_column=c, end_row=r + 1, end_column=c)
        rc.add(ws.title, r, 'item', dict(path=[FURN[u['furn']], u['sysname'], nm.strip('　')], code=u['tag'], row='計画', marks={str(y): v for y, v in plan.items() if v}, cost=total or None, unit='千円'))
        rc.add(ws.title, r + 1, 'item', dict(path=[FURN[u['furn']], u['sysname'], nm.strip('　')], code=u['tag'], row='実績', marks={str(y): v for y, v in act.items() if v}),
               [w for w in CUR.get('wmap', {}).get(u['tag'], [])])
        return r + 2
    bytag = {u['tag']: u for u in world.UNITS}
    for u in machines:
        if (u['furn'], u['sys']) != cur:
            if u['furn'] != cur[0]:
                merge(ws, r, 1, r, last, {0: '【共通】', 1: '【1号炉】', 2: '【2号炉】'}[u['furn']], font(10, True), AL, B, fill('D9D9D9')); r += 1
            merge(ws, r, 1, r, last, f"{u['furn']}-{u['sys']}　{u['sysname']}", font(9, True), AL, B, fill('F2F2F2')); r += 1
            cur = (u['furn'], u['sys'])
        n += 1; r = unit_rows(u, r, n)
    merge(ws, r, 1, r, last, '【電動機】（電動機台帳より。番号は電動機No.）', font(10, True), AL, B, fill('D9D9D9')); r += 1
    for u in motors:
        n += 1; r = unit_rows(u, r, n)
    r += 1
    merge(ws, r, 1, r, 4, '法定点検・測定等', font(9, True), AL, B, fill('FFF2CC')); r += 1
    for nm, law, ys, conf in legal_simple(MADE, YRS, True):
        put(ws, r, 3, nm, font(8), AL, B); put(ws, r, 2, law, font(7), AC, B)
        mk_ = {}
        for y, c in ycol.items():
            v = ys.get(y, '')
            if y < MADE and v: v = v + '済' if not two else v
            put(ws, r, c, v, font(8), AC, B)
            if v: mk_[str(y)] = v
        rc.add(ws.title, r, 'item', dict(path=['法定点検・測定等', nm], work=nm, marks=mk_, basis=law)); r += 1
    ws.freeze_panes = 'F6' if two else 'E6'
    page(ws, land=True, paper=8, titles='5:5', header='【架空】あかね台資源化センター 中長期整備計画')
    # 年度別概算費用
    ws2 = wb.create_sheet('年度別概算費用')
    sysn = []
    for u in machines + motors:
        if u['sysname'] not in sysn: sysn.append(u['sysname'])
    widths(ws2, [24] + [9] * len(YRS) + [11])
    merge(ws2, 1, 1, 1, 8, f'【架空】年度別概算費用（千円）　{world.wareki(MADE)}年度', font(12, True), AL, b=None)
    put(ws2, 3, 1, '系統', font(8, True), AC, B, fill('FFF2CC'))
    for i, y in enumerate(YRS): put(ws2, 3, 2 + i, world.wa_short(y), font(8, True), AC, B, fill('FFF2CC'))
    put(ws2, 3, 2 + len(YRS), '計', font(8, True), AC, B, fill('FFF2CC'))
    rr = 4; tot = {}
    for sname in sysn:
        put(ws2, rr, 1, sname, font(8), AL, B); s_ = 0
        for i, y in enumerate(YRS):
            v = costs.get((sname, y), 0) if (two or y >= MADE) else None
            put(ws2, rr, 2 + i, v or None, font(8), AR, B, nf='#,##0'); s_ += v or 0; tot[y] = tot.get(y, 0) + (v or 0)
        put(ws2, rr, 2 + len(YRS), s_ or None, font(8), AR, B, nf='#,##0')
        rc.add(ws2.title, rr, 'subtotal', dict(path=[sname], cost={str(y): costs.get((sname, y)) for y in YRS if costs.get((sname, y)) and (two or y >= MADE)}, unit='千円'))
        rr += 1
    put(ws2, rr, 1, '合計', font(8, True), AC, B, fill('F2F2F2'))
    for i, y in enumerate(YRS): put(ws2, rr, 2 + i, tot.get(y) or None, font(8, True), AR, B, fill('F2F2F2'), '#,##0')
    rc.add(ws2.title, rr, 'total', dict(path=['合計'], cost={str(y): tot[y] for y in YRS if tot.get(y)}, unit='千円'))
    put(ws2, rr + 2, 1, '※ 実施済の年度は空欄（決算は財務会計による）' if not two else '※ 計画（当初）の概算。実施済の年度も当初の額', font(8), AL)
    page(ws2, land=True)
    wb.save(path)
