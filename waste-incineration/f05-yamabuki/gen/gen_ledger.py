# -*- coding: utf-8 -*-
"""機器台帳：旧版（2021、組合が竣工図書から作り直した 1 シート）と、
2023 刷新版（基幹改良時にメーカーが提出した機器表を組合が台帳にしたもの。設備別シート）"""
import os
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter as L
from common import *
from world import MACHINES, old_machines, SYS, SYSNAME

SYS_OLD = {'R': '受入供給', 'B': '燃焼', 'F': '通風', 'C': 'ガス冷却', 'G': '排ガス処理', 'A': '灰出し',
           'O': '燃料', 'W': '給水', 'D': '排水処理', 'K': '空気圧縮', 'E': '電気', 'I': '計装', 'S': '計装',
           'M': 'その他'}


def h_date(ym):
    y, m = int(ym[:4]), int(ym[5:])
    if y < 2019 or (y == 2019 and m < 5):
        return f'H{y-1988}.{m}'
    return f'R{y-2018}.{m}'


def build_old(path):
    wb = Workbook(); ws = wb.active; ws.title = '機器一覧'
    props(wb, '【架空】やまぶき広域処理組合 第二工場 主要機器一覧（機器台帳）令和3年度', '組合事務局 y1')
    cols = [('No', 5), ('区分', 9), ('機器名称', 26), ('号炉', 5), ('仕様', 34), ('数量', 5), ('電動機(kW)', 8),
            ('メーカー', 13), ('設置', 7), ('備考', 24)]
    for i, (t, w) in enumerate(cols, 1):
        ws.column_dimensions[L(i)].width = w
    merge(ws, 1, 1, 1, 10, '【架空】やまぶき広域処理組合 第二工場　主要機器一覧', f(14, True), AL)
    merge(ws, 2, 1, 2, 10, '令和3年5月作成（竣工図書 H13.3「機器仕様書」より転記）　組合事務局 施設係　※事務局移転により旧データ無し', f(9), AL)
    r = 4
    for i, (t, w) in enumerate(cols, 1):
        put(ws, r, i, t, f(9, True), AC, fill('D9D9D9'))
    box(ws, r, 1, r, 10)
    r += 1
    prev = None
    for no, m in enumerate(old_machines(), 1):
        nm = m['old']
        kub = SYS_OLD[m['sys']]
        put(ws, r, 1, no, f(9), AC)
        put(ws, r, 2, kub if (prev is None or SYS_OLD[prev['sys']] != kub) else '〃', f(9), AC)
        put(ws, r, 3, nm, f(9), AL)
        put(ws, r, 4, {0: '-', 1: '1号', 2: '2号'}[m['furnace']], f(9), AC)
        spec = m['spec']
        if m.get('kikan') and m['furnace'] == 1:
            spec = spec  # R4 更新後だが旧版は竣工時のまま
        # 竣工時の仕様（基幹改良で変わった所は旧仕様）
        spec = (spec.replace('ASR-30', 'ASR-30R').replace(' 【R5新設】', '').replace('【2018休止】', '').strip())
        if m['cls'] == 'bf':
            spec = 'パルスジェット式 4室 ろ布192本'
        if m['name'] == '減温塔':
            spec = '水噴射式 一流体噴霧'
        put(ws, r, 5, spec, f(8), AL)
        put(ws, r, 6, 1, f(9), AC)
        put(ws, r, 7, m['kw'], f(9), AC, nf='0.0#')
        mk = m['maker'] if m['maker'] != '-' else ''
        if m['cls'] in ('damper', 'static', 'tank', 'conveyor', 'static_hot') and mk:
            mk = ''   # 竣工図書の仕様書にメーカー記載なし
        put(ws, r, 8, mk, f(8), AL)
        inst = m['inst'] if m.get('renew') is None or m['renew'] < 2021 else '2001-03'
        put(ws, r, 9, h_date(inst), f(8), AC)
        note = ''
        if m['status'] == '休止':
            note = 'H30.4〜休止'
        elif m.get('renew') and m['renew'] < 2021:
            note = f"H{m['renew']-1988}更新"
        elif m.get('olddcs'):
            note = ''
        if m['cls'] == 'crane':
            note = '1号常用・2号予備（月替り運転なし）'
        if m['cls'] == 'pv2':
            note = '第二種圧力容器'
        put(ws, r, 10, note, f(8), AL)
        prev = m
        r += 1
    box(ws, 4, 1, r - 1, 10)
    put(ws, r + 1, 3, f'計　{len(old_machines())}台', f(9, True), AL)
    ws.freeze_panes = 'A5'
    setup_print(ws, True, 9, 1, 0, f'A1:J{r+1}', '4:4')
    ws.oddFooter.center.text = '- &P -'
    wb.save(path)


SHEETS = [('受入供給', ['R']), ('燃焼・燃料', ['B', 'O']), ('通風', ['F']), ('ガス冷却・排ガス処理', ['C', 'G']),
          ('灰出し', ['A']), ('給排水', ['W', 'D']), ('圧縮空気・雑', ['K', 'M']), ('電気', ['E']), ('計装・制御', ['I', 'S'])]


def build_new(path):
    wb = Workbook()
    props(wb, '【架空】やまぶき広域処理組合 第二工場 機器台帳（基幹的設備改良工事 完成時 機器表より）', '組合事務局 y1（機器表：朝凪炉工 n3 提出）')
    cv = wb.active; cv.title = '表紙・集計'
    for c, w in zip('ABCDEFG', (4, 24, 10, 10, 10, 10, 30)):
        cv.column_dimensions[c].width = w
    merge(cv, 2, 2, 2, 7, '【架空】やまぶき広域処理組合 第二工場', f(16, True), AC)
    merge(cv, 3, 2, 3, 7, '機　器　台　帳', f(20, True), AC)
    merge(cv, 5, 2, 5, 7, '令和6年3月　（基幹的設備改良工事（令和4〜5年度）完成図書「機器表」を基に作成）', f(10), AC)
    merge(cv, 6, 2, 6, 7, '機器表提出：【架空】朝凪炉工株式会社（現場代理人 n3）　台帳整理：組合事務局 施設係 y1', f(9), AC)
    r = 8
    for i, t in enumerate(['シート', '1号炉', '2号炉', '共通', '計', '備考'], 2):
        put(cv, r, i, t, f(10, True), AC, fill('BDD7EE'))
    tot = [0, 0, 0]
    for sname, codes in SHEETS:
        r += 1
        cnt = [sum(1 for m in MACHINES if m['sys'] in codes and m['furnace'] == k) for k in (1, 2, 0)]
        put(cv, r, 2, sname, f(10), AL)
        for j, v in enumerate(cnt):
            put(cv, r, 3 + j, v, f(10), AC); tot[j] += v
        put(cv, r, 6, sum(cnt), f(10), AC)
        put(cv, r, 7, '休止設備を含む' if sname == '灰出し' else ('新DCS 8台（旧5台撤去）' if sname == '計装・制御' else ''), f(9), AL)
    r += 1
    put(cv, r, 2, '合計', f(10, True), AC)
    for j, v in enumerate(tot):
        put(cv, r, 3 + j, v, f(10, True), AC)
    put(cv, r, 6, sum(tot), f(10, True), AC)
    box(cv, 8, 2, r, 7)
    merge(cv, r + 2, 2, r + 4, 7, '数え方：機器1台を1行（電動機は機器に付属し kW を記載）。主要計装（温度計・流量計・分析計等）は1台として計上。弁類・建築設備は対象外。\n機器番号＝メーカー機器表の番号（炉番号＋設備記号＋連番、共通は 0）', f(9), ALT)
    setup_print(cv, False, 9, 1, 1)
    cols = [('機器番号', 9), ('機器名称', 24), ('号炉', 5), ('形式・仕様', 36), ('電動機\nkW', 6.5), ('製作者', 13),
            ('製造番号', 9), ('設置年月', 8.5), ('改良区分', 12), ('常用/予備', 6.5), ('状態', 6), ('旧台帳名称（R3）', 20), ('備考', 22)]
    oldno = {}
    for no, m in enumerate(old_machines(), 1):
        oldno[id(m)] = no
    for sname, codes in SHEETS:
        ws = wb.create_sheet(sname)
        for i, (t, w) in enumerate(cols, 1):
            ws.column_dimensions[L(i)].width = w
        merge(ws, 1, 1, 1, 10, f'【架空】やまぶき広域処理組合 第二工場　機器台帳　（{sname}）', f(12, True), AL)
        put(ws, 1, 13, 'R6.3 作成', f(9), AR)
        r = 3
        for i, (t, w) in enumerate(cols, 1):
            put(ws, r, i, t, f(9, True), AC, fill('BDD7EE' if i <= 11 else 'FFF2CC'))
        ws.cell(r, 12).comment = Comment('組合で追記（旧台帳との突合せ用）', 'y1')
        ws.row_dimensions[r].height = 28
        r += 1
        for code in codes:
            merge(ws, r, 1, r, 13, f'{code}　{SYSNAME[code]}', f(9, True), AL, fill('EDEDED'))
            r += 1
            for k in (1, 2, 0):
                grp = [m for m in MACHINES if m['sys'] == code and m['furnace'] == k]
                for m in grp:
                    put(ws, r, 1, m['tag'], f(8), AC)
                    put(ws, r, 2, m['name'], f(9), AL)
                    put(ws, r, 3, {0: '共通', 1: '1号', 2: '2号'}[k], f(8), AC)
                    put(ws, r, 4, m['spec'].replace('【2018休止】', '').replace('【R5新設】', '').strip(), f(8), AL)
                    put(ws, r, 5, m['kw'], f(8), AC, nf='0.0#')
                    put(ws, r, 6, m['maker'] if m['maker'] != '-' else '（不明）', f(8), AL)
                    put(ws, r, 7, m['serial'] if (m.get('kikan') or (m['renew'] or 0) >= 2012) else '-', f(8), AC)
                    put(ws, r, 8, m['inst'].replace('-', '/'), f(8), AC)
                    kb = '既設'
                    if m.get('kikan'):
                        kb = 'R4基幹' if m['renew'] == 2022 else 'R5基幹'
                        if m['old'] is None:
                            kb = 'R5基幹(新設)'
                    elif m.get('renew'):
                        kb = f"既設(H{m['renew']-1988}更新)" if m['renew'] < 2019 else '既設'
                    put(ws, r, 9, kb, f(8, kb.startswith('R')), AC, fill('FFF2CC') if kb.startswith('R') else None)
                    put(ws, r, 10, m['standby'], f(8), AC)
                    put(ws, r, 11, m['status'], f(8, color='C00000' if m['status'] == '休止' else '000000'), AC)
                    on = m['old'] if (m['old'] and not m.get('old_same')) else ('（同左）' if m.get('old_same') else '（新設）')
                    put(ws, r, 12, on, f(8, color='595959'), AL)
                    note = ''
                    if m['status'] == '休止':
                        note = 'H30.4〜休止・残置'
                    if m['cls'] == 'pv2':
                        note = '第二種圧力容器（定期自主検査 年1）'
                    if m['cls'] == 'crane':
                        note = 'クレーン検査証あり（性能検査2年）'
                    if m['cls'] == 'hoist':
                        note = '定期自主検査 年1'
                    if m['cls'] == 'scale':
                        note = '特定計量器（定期検査2年）'
                    if m['name'] == '非常用発電機':
                        note = '消防用非常電源兼用'
                    if m['name'] == 'ピット火災検知装置':
                        note = 'H25追加（竣工図書外）'
                    put(ws, r, 13, note, f(8), AL)
                    r += 1
        box(ws, 3, 1, r - 1, 13)
        ws.freeze_panes = 'C4'
        setup_print(ws, True, 9, 1, 0, f'A1:M{r-1}', '3:3')
        ws.oddFooter.center.text = '&A　&P/&N'
    wb.save(path)


if __name__ == '__main__':
    build_old(os.path.join(OUT, f'{FN}_機器台帳_令和3年度旧版_組合1シート.xlsx'))
    build_new(os.path.join(OUT, f'{FN}_機器台帳_令和5年度刷新_メーカー機器表様式.xlsx'))
