# -*- coding: utf-8 -*-
"""中長期整備計画：y1 様式（R5 年度版）と y2 様式（R6 改訂版、y1 のファイルをコピーして直したもの）"""
import os, copy
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter as L
from common import *
from world import MACHINES, old_machines, SYS, SYSNAME, PER_FURNACE
from plan_data import *

FUR = {0: '共', 1: '1', 2: '2'}


def ordered(ms):
    """設備ごと。炉別機器は 1号・2号を並べ、その後に共通"""
    res = []
    for code, sname in SYS:
        g = [m for m in ms if m['sys'] == code]
        if not g:
            continue
        f1 = [m for m in g if m['furnace'] == 1]
        f2 = [m for m in g if m['furnace'] == 2]
        c0 = [m for m in g if m['furnace'] == 0]
        res.append((code, sname, f1, f2, c0))
    return res


# ---------------------------------------------------------------- y1
def build_y1(path):
    wb = Workbook(); ws = wb.active; ws.title = '整備計画(R5)'
    props(wb, '【架空】やまぶき広域処理組合 第二工場 延命化整備計画 年次スケジュール 令和5年度版', '組合事務局 y1')
    ys = Y1_YEARS
    c0 = 6
    ncol = c0 + len(ys)
    widths = [5, 3, 26, 4.5, 7] + [6.2] * len(ys) + [30]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[L(i)].width = w
    merge(ws, 1, 1, 1, ncol, '【架空】やまぶき広域処理組合 第二工場　延命化整備計画　年次スケジュール（令和5年度版）', f(14, True), AL)
    merge(ws, 2, 1, 2, 8, '令和5年4月　組合事務局 施設係（y1）　※令和2年度策定 延命化計画を基に毎年度見直し', f(9), AL)
    merge(ws, 3, 1, 3, ncol, '凡例　○点検　◎整備（分解・部品交換）　●更新　△補修　★基幹的設備改良工事（令和4〜5年度）　「済」＝実施済　　金額：千円（税込概算、設備ごとの年額）', f(9), AL)
    r = 5
    merge(ws, r, 1, r + 1, 1, 'No', f(9, True), AC, fill('D9E1F2'))
    merge(ws, r, 2, r + 1, 2, '設備', f(9, True), AC, fill('D9E1F2'))
    merge(ws, r, 3, r + 1, 3, '機器名', f(9, True), AC, fill('D9E1F2'))
    merge(ws, r, 4, r + 1, 4, '炉', f(9, True), AC, fill('D9E1F2'))
    merge(ws, r, 5, r + 1, 5, '設置', f(9, True), AC, fill('D9E1F2'))
    merge(ws, r, c0, r, ncol - 1, '整　備　年　度', f(9, True), AC, fill('D9E1F2'))
    for i, y in enumerate(ys):
        put(ws, r + 1, c0 + i, f'R{wareki(y)}', f(9, True), AC, fill('FCE4D6' if y == 2023 else 'D9E1F2'))
    merge(ws, r, ncol, r + 1, ncol, '備考', f(9, True), AC, fill('D9E1F2'))
    box(ws, r, 1, r + 1, ncol)
    r = 7
    no = 0
    total = {y: 0 for y in ys}
    rows_meta = []
    for code, sname, f1, f2, cm in ordered(old_machines()):
        gr = r
        merge(ws, r, 2, r, 5, f'■ {sname}', f(10, True), AL, fill('EDEDED'))
        put(ws, r, 1, None, fl=fill('EDEDED'))
        put(ws, r, ncol, '（千円）', f(8), AL, fill('EDEDED'))
        r += 1
        gsum = {y: 0 for y in ys}
        seq = []
        for a, b in zip(f1, f2):
            seq += [a, b]
        seq += cm
        prev = None
        for m in seq:
            no += 1
            put(ws, r, 1, no, f(9), AC)
            nm = m['old'] or m['name']
            put(ws, r, 2, None)
            # 1号・2号 が続くとき、2号行の機器名は 〃（y1 の癖）
            if prev is not None and prev['old'] == m['old'] and prev['furnace'] == 1 and m['furnace'] == 2:
                put(ws, r, 3, '　〃', f(9), AL)
            else:
                put(ws, r, 3, nm, f(9), AL)
            put(ws, r, 4, FUR[m['furnace']], f(9), AC)
            inst = m['inst'] if m['inst'] <= '2023-03' else '2001-03'
            iy, im = inst[:4], inst[5:]
            ins = 'H13' if inst == '2001-03' else f"H{int(iy)-1988}" if int(iy) < 2019 else f"R{int(iy)-2018}"
            put(ws, r, 5, ins, f(8), AC)
            mk = y1_marks(m)
            for i, y in enumerate(ys):
                s = mk.get(y)
                if not s:
                    continue
                txt = s
                done = y1_done(m, y, s)
                if done is True:
                    txt = s + '済'
                cell = put(ws, r, c0 + i, txt, f(10 if s != '★' else 10, s == '★', 'C00000' if s == '★' else '000000'), AC,
                           fill('FFF2CC') if s == '★' else None)
                if done is False:
                    cell.comment = Comment('R4全炉停止で点検、状態良好のため見送り' if y == 2022 else '予算減額のため先送り', 'y1')
                    Y1_NOT_DONE.add((m['old'], m['furnace'], y))
                gsum[y] += y1_cost(m, s)
            note = ''
            if m.get('kikan') or m.get('olddcs'):
                note = '基幹改良対象' + ('（撤去・更新）' if m.get('olddcs') else '')
            if m['status'] == '休止':
                note = 'H30より休止（灰資源化委託）'
            if m['cls'] == 'crane':
                note = '年次自主検査　' + ('常用' if '1号' in nm else '予備')
            put(ws, r, ncol, note, f(8), AL)
            rows_meta.append(r)
            prev = m
            r += 1
        for i, y in enumerate(ys):
            v = gsum[y]
            put(ws, gr, c0 + i, v if v else None, f(8, False, '1F3864', italic=True), AR, fill('EDEDED'), '#,##0')
            total[y] += v
        box(ws, gr, 1, r - 1, ncol)
    # 合計
    merge(ws, r, 1, r, 5, '合　計（千円）', f(10, True), AC, fill('D9E1F2'))
    for i, y in enumerate(ys):
        put(ws, r, c0 + i, total[y], f(9, True), AR, fill('D9E1F2'), '#,##0')
    put(ws, r, ncol, '基幹改良は別枠（継続費）', f(8), AL, fill('D9E1F2'))
    box(ws, r, 1, r, ncol, med)
    r += 2
    merge(ws, r, 1, r, 5, '■ 法定点検等', f(10, True), AL, fill('EDEDED'))
    r += 1
    legal = [
        ('ごみクレーン性能検査', {y: '○' for y in (2022, 2024, 2026, 2028, 2030)}, '2年（検査証）'),
        ('受変電設備年次点検', {y: '○' for y in ys}, '電気保安協会委託'),
        ('計量機定期検査', {y: '○' for y in (2021, 2023, 2025, 2027, 2029)}, '2年'),
        ('精密機能検査', {y: '○' for y in (2021, 2024, 2027, 2030)}, '3年'),
        ('排ガス測定', {y: '○' for y in ys}, '年2回（DXN年1回）'),
    ]
    st = r
    for nm, mk, note in legal:
        merge(ws, r, 2, r, 5, nm, f(9), AL)
        for i, y in enumerate(ys):
            if y in mk:
                put(ws, r, c0 + i, mk[y] + ('済' if y < 2023 else ''), f(10), AC)
        put(ws, r, ncol, note, f(8), AL)
        r += 1
    box(ws, st - 1, 1, r - 1, ncol)
    ws.freeze_panes = 'F7'
    setup_print(ws, True, 8, 1, 0, f'A1:{L(ncol)}{r-1}', '5:6')
    ws.oddFooter.center.text = '&P / &N'
    wb.save(path)
    return wb


# ---------------------------------------------------------------- y2
def build_y2(path):
    """y1 のファイルをコピーして直した：旧No 列を残す、記号を漢字1字へ、15年・西暦併記、号炉の中間見出し、重要度・予備列"""
    wb = Workbook(); ws = wb.active; ws.title = '整備計画(R6改訂)'
    props(wb, '【架空】やまぶき広域処理組合 第二工場 長寿命化整備計画 年次スケジュール 令和6年度改訂版', '組合事務局 y2（y1 R5版を改訂）')
    ys = Y2_YEARS
    c0 = 10
    ncol = c0 + len(ys)
    widths = [4.2, 8.5, 2.5, 22, 3.5, 4.5, 4.5, 7.5, 14] + [6.6] * len(ys) + [30]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[L(i)].width = w
    merge(ws, 1, 1, 1, ncol - 2, '【架空】やまぶき広域処理組合 第二工場　長寿命化整備計画　年次スケジュール（令和6年度改訂版）', f(14, True), AL)
    put(ws, 1, ncol, '令和5年度版', f(8, color='A6A6A6'), AR)   # y1 版の表題の残り
    merge(ws, 2, 1, 2, 12, '令和6年12月改訂　組合事務局 施設係（y2）　※基幹的設備改良工事（R4〜R5）完了に伴い全面見直し', f(9), AL)
    merge(ws, 3, 1, 3, ncol, '凡例　点＝点検・測定　整＝整備（分解・部品交換）　補＝補修　更＝更新　／　網掛け＝R6実績　／　金額：号炉ごとの小計（千円・税込概算）、定期整備に含むものは計上しない', f(9), AL)
    merge(ws, 4, 1, 4, ncol, '（旧）○点検　◎整備（分解・部品交換）　●更新　△補修　★基幹的設備改良工事（令和4〜5年度）　「済」＝実施済', f(8, color='808080', strike=True), AL)
    r = 5
    hdr = fill('DDEBF7')
    for c, t in ((1, '旧No'), (2, '機器番号'), (3, '設備・号炉・機器'), (5, '炉'), (6, '重要度'), (7, '予備'), (8, '設置'), (9, 'R3〜R5の主な実績')):
        if c == 3:
            merge(ws, r, 3, r + 2, 4, t, f(9, True), AC, hdr)
        else:
            merge(ws, r, c, r + 2, c, t, f(9 if c != 1 else 8, True, '808080' if c == 1 else '000000'), AC, hdr)
    ws.cell(r, 6).comment = Comment('R6改訂で追加。A：停止で炉停止・代替なし　B：運転に影響／予備あり　C：事後保全', 'y2')
    merge(ws, r, c0, r, ncol - 1, '整　備　年　度', f(9, True), AC, hdr)
    for i, y in enumerate(ys):
        put(ws, r + 1, c0 + i, y, f(8, True), AC, hdr)
        put(ws, r + 2, c0 + i, f'R{wareki(y)}', f(8, True), AC, fill('FCE4D6') if y == 2024 else hdr)
    merge(ws, r, ncol, r + 2, ncol, '備考', f(9, True), AC, hdr)
    box(ws, r, 1, r + 2, ncol)
    # y1 の基幹改良の帯の 2 行を消さずに空けたまま
    ws.row_dimensions[8].height = 6; ws.row_dimensions[9].height = 6
    r = 10
    # 旧 No の対応（y1 版の並びの No）
    oldno = {}
    n = 0
    for code, sname, f1, f2, cm in ordered(old_machines()):
        seq = []
        for a, b in zip(f1, f2):
            seq += [a, b]
        seq += cm
        for m in seq:
            n += 1
            oldno[(m['old'], m['furnace'])] = n
    total = {y: 0 for y in ys}
    subtotal_rows = []
    for code, sname, f1, f2, cm in ordered(MACHINES):
        merge(ws, r, 3, r, 9, f'■ {sname}', f(10, True), AL, fill('E2EFDA'))
        for c in (1, 2):
            put(ws, r, c, None, fl=fill('E2EFDA'))
        for i in range(len(ys)):
            put(ws, r, c0 + i, None, fl=fill('E2EFDA'))
        put(ws, r, ncol, None, fl=fill('E2EFDA'))
        gtop = r
        r += 1
        for label, grp in (('1号炉', f1), ('2号炉', f2), ('共通', cm)):
            if not grp:
                continue
            sr = r
            put(ws, r, 3, None)
            merge(ws, r, 4, r, 9, f'【{label}】', f(9, True), AL, fill('F2F2F2'))
            for c in (1, 2, 3):
                put(ws, r, c, None, fl=fill('F2F2F2'))
            put(ws, r, ncol, '小計（千円）', f(8), AL, fill('F2F2F2'))
            r += 1
            ssum = {y: 0 for y in ys}
            for m in grp:
                on = oldno.get((m['old'], m['furnace']))
                put(ws, r, 1, on, f(8, color='808080'), AC)
                put(ws, r, 2, m['tag'], f(8), AC)
                put(ws, r, 3, None)
                put(ws, r, 4, m['name'], f(9), AL)
                put(ws, r, 5, FUR[m['furnace']], f(8), AC)
                put(ws, r, 6, importance(m), f(9, importance(m) == 'A'), AC)
                put(ws, r, 7, '有' if m['standby'] else '', f(8), AC)
                put(ws, r, 8, m['inst'].replace('-', '.'), f(8), AC)
                past = ''
                if m.get('kikan'):
                    past = f"R{wareki(m['renew'])} 基幹改良で更新"
                elif m['name'] == '灰押出装置' and m['furnace'] == 2:
                    past = 'R4 摺動板取替（追加）'
                elif m['name'] == 'ごみクレーン（1号）':
                    past = 'R4 ワイヤロープ交換'
                elif m['name'] == 'バグフィルタ' and False:
                    past = ''
                elif m['name'] == '空気圧縮機（1号）':
                    past = 'R3 オーバーホール'
                elif m['name'] == '計量機（入口）':
                    past = 'R3・R5 定期検査'
                put(ws, r, 9, past, f(7.5), AL)
                mk = y2_marks(m)
                # R6 実績で計画外のもの（2024 年に起きたこと）
                if m['name'] == '計量機（出口）':
                    mk[2024] = ('補', 0, 'R6 ロードセル1点交換（器差）')
                for i, y in enumerate(ys):
                    if y not in mk:
                        continue
                    s, cost, note = mk[y]
                    rr = rng_for(m, 'r6')
                    fl = None
                    cmt = None
                    if y == 2024:
                        if rr.random() < 0.08 and s != '点':
                            cmt = 'R6未実施→R7へ'
                            if 2025 not in mk:
                                mk[2025] = (s, cost, note)
                        else:
                            fl = fill('BFBFBF')
                    cell = put(ws, r, c0 + i, s, f(9, s == '更', 'C00000' if s == '更' else '000000'), AC, fl)
                    if cmt:
                        cell.comment = Comment(cmt, 'y2')
                    ssum[y] += cost
                notes = []
                if m['status'] == '休止':
                    notes.append('休止・残置（H30〜灰資源化委託）')
                for y in ys:
                    if y in mk and mk[y][2] and y >= 2025 and mk[y][0] in ('更', '整', '補'):
                        notes.append(f"R{wareki(y)}{mk[y][2]}")
                        break
                if m['name'] == '計量機（出口）':
                    notes.insert(0, 'R6 ロードセル1点交換（器差）')
                if m['cls'] == 'crane':
                    notes.append('年次自主検査・性能検査は下段')
                put(ws, r, ncol, '　'.join(notes)[:40], f(7.5), AL)
                r += 1
            for i, y in enumerate(ys):
                v = ssum[y]
                total[y] += v
                put(ws, sr, c0 + i, v if v else None, f(7.5, color='1F3864', italic=True), AR, fill('F2F2F2'), '#,##0')
            subtotal_rows.append((code, label, sr))
        box(ws, gtop, 1, r - 1, ncol)
    # わざと残した小計の食い違い：通風設備 1号炉 R9 は修正前の値のまま（+600）
    for code, label, sr in subtotal_rows:
        if code == 'F' and label == '1号炉':
            cidx = c0 + ys.index(2027)
            cell = ws.cell(sr, cidx)
            cell.value = (cell.value or 0) + 600
            NOTES_MISMATCH.append(f'通風設備【1号炉】小計 R9(2027) セル {L(cidx)}{sr} は {cell.value} 千円（明細の合計より +600、押込送風機の整備を R9 から R10 へ移した時の直し忘れ）。合計行は明細から計算した値')
    merge(ws, r, 1, r, 9, '合　計（千円）', f(10, True), AC, fill('DDEBF7'))
    for i, y in enumerate(ys):
        put(ws, r, c0 + i, total[y], f(8, True), AR, fill('DDEBF7'), '#,##0')
    put(ws, r, ncol, '基幹改良（R4〜R5）は含まない', f(8), AL, fill('DDEBF7'))
    box(ws, r, 1, r, ncol, med)
    r += 2
    # 法定点検等：y1 の 5 行をコピーして貼り直し、行を追加。項目名は 1 列右（E 列から）に入ってしまっている
    merge(ws, r, 3, r, 9, '■ 法定点検・測定等', f(10, True), AL, fill('E2EFDA'))
    r += 1
    def yrs(*a):
        return set(a)
    allys = set(ys)
    legal = [
        ('ごみクレーン年次自主検査（1・2号）', allys, '年1回　月次は運転委託'),
        ('ごみクレーン性能検査（1・2号）', yrs(2024, 2026, 2028, 2030, 2032, 2034, 2036, 2038), '2年（検査証の有効期間）'),
        ('ホイスト定期自主検査（4基）', allys, '年1回　R6追加'),
        ('受変電設備年次点検', allys, '電気保安協会委託　停電を伴う→全炉停止中'),
        ('非常用発電機点検（消防・保安）', allys, '機器6月・総合1年　R6追加'),
        ('第二種圧力容器定期自主検査（空気槽2基）', allys, '年1回　R6追加'),
        ('排ガス測定（ばい煙・DXN）', allys, 'ばい煙 年2回・DXN 年1回'),
        ('DXN類作業環境測定', allys, '6月以内ごと　R6追加'),
        ('計量機定期検査', yrs(2025, 2027, 2029, 2031, 2033, 2035, 2037), '2年'),
        ('消防用設備等点検', allys, '機器6月・総合1年　R6追加'),
        ('精密機能検査', yrs(2024, 2027, 2030, 2033, 2039), '3年'),   # R18(2036) が抜けている（コピー時の欠け）
    ]
    st = r
    for nm, ss, note in legal:
        merge(ws, r, 5, r, 9, nm, f(9), AL)
        for i, y in enumerate(ys):
            if y in ss:
                put(ws, r, c0 + i, '点', f(9), AC, fill('BFBFBF') if y == 2024 else None)
        put(ws, r, ncol, note, f(8), AL)
        r += 1
    box(ws, st - 1, 1, r - 1, ncol)
    put(ws, r, 4, '※第一種圧力容器：該当なし（ボイラ・蒸気設備なし）', f(8), ALN)
    ws.freeze_panes = ws.cell(10, c0)
    setup_print(ws, True, 8, 1, 0, f'A1:{L(ncol)}{r}', '5:7')
    ws.oddFooter.center.text = '&P / &N'
    ws.oddFooter.right.text = '長寿命化整備計画（R6改訂）'
    # y1 の R5 版シートを非表示で残す（コピー元）
    wb2 = build_y1(os.path.join(WORK, '_y1_tmp.xlsx'))
    old = wb.create_sheet('R5版(y1)')
    src = wb2.active
    for row in src.iter_rows():
        for c in row:
            if c.value is not None:
                nc = old.cell(c.row, c.column, c.value)
                nc.font = copy.copy(c.font); nc.alignment = copy.copy(c.alignment)
                nc.fill = copy.copy(c.fill); nc.border = copy.copy(c.border); nc.number_format = c.number_format
    for rng in src.merged_cells.ranges:
        old.merge_cells(str(rng))
    for k, d in src.column_dimensions.items():
        old.column_dimensions[k].width = d.width
    old.sheet_state = 'hidden'
    wb.save(path)


NOTES_MISMATCH = []

if __name__ == '__main__':
    p1 = os.path.join(OUT, f'{FN}_中長期整備計画_令和5年度版_y1様式.xlsx')
    p2 = os.path.join(OUT, f'{FN}_中長期整備計画_令和6年度改訂版_y2様式.xlsx')
    build_y1(p1)
    build_y2(p2)
    print(NOTES_MISMATCH)
    print(sorted(Y1_NOT_DONE, key=str)[:40])
