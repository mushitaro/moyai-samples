# -*- coding: utf-8 -*-
"""中長期整備計画（全版）：y1 R3・R4・R5 年度版、y2 R6 改訂・R7 年度版。様式は見本（承認済み）のまま。
過去の年度の列は当初計画の記号を残し、実施は y1「済」／y2 網掛け、見送りはセルのコメント。"""
import os, copy
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter as L
from common import *
from world import MACHINES, old_machines, SYS, SYSNAME
import plan_data as PD
import works as W
import rec

FUR = {0: '共', 1: '1', 2: '2'}
Y2SYM = {'○': '点', '◎': '整', '●': '更', '△': '補', '★': '更'}
GAPS = []   # 計画の欠け（world.json へ）


def ordered(ms):
    res = []
    for code, sname in SYS:
        g = [m for m in ms if m['sys'] == code]
        if not g:
            continue
        res.append((code, sname, [m for m in g if m['furnace'] == 1], [m for m in g if m['furnace'] == 2], [m for m in g if m['furnace'] == 0]))
    return res


def seq_of(f1, f2, cm):
    s = []
    for a, b in zip(f1, f2):
        s += [a, b]
    return s + cm


def actual_sym(m, y):
    """その年度の実績（定期整備・別契約・基幹改良）から記号"""
    wts = W.wtypes_in(m['tag'], y) if m.get('tag') else set()
    if not wts:
        return None
    if 'UPGRADE' in wts:
        return '★'
    if 'REPLACE' in wts:
        return '更'
    if 'OVERHAUL' in wts:
        return '補' if m['cls'] in ('grate', 'furnace', 'ashpusher', 'bf', 'tower') else '整'
    return '点'


# ======================================================================= y1
def build_y1(path, V):
    fname = os.path.basename(path)
    stem = rec.start(fname, plan=True)
    wb = Workbook(); ws = wb.active; ws.title = f'整備計画(R{V-2018})'
    props(wb, f'【架空】やまぶき広域処理組合 第二工場 延命化整備計画 年次スケジュール 令和{V-2018}年度版', '組合事務局 y1')
    ys = PD.Y1_YEARS
    c0 = 6
    ncol = c0 + len(ys)
    for i, w in enumerate([5, 3, 26, 4.5, 7] + [6.2] * len(ys) + [30], 1):
        ws.column_dimensions[L(i)].width = w
    merge(ws, 1, 1, 1, ncol, f'【架空】やまぶき広域処理組合 第二工場　延命化整備計画　年次スケジュール（令和{V-2018}年度版）', f(14, True), AL)
    merge(ws, 2, 1, 2, 8, f'令和{V-2018}年4月　組合事務局 施設係（y1）　※令和2年度策定 延命化計画を基に毎年度見直し', f(9), AL)
    merge(ws, 3, 1, 3, ncol, '凡例　○点検　◎整備（分解・部品交換）　●更新　△補修　★基幹的設備改良工事（令和4〜5年度）　「済」＝実施済　　金額：千円（税込概算、設備ごとの年額）', f(9), AL)
    r = 5
    for c, t in ((1, 'No'), (2, '設備'), (3, '機器名'), (4, '炉'), (5, '設置')):
        merge(ws, r, c, r + 1, c, t, f(9, True), AC, fill('D9E1F2'))
    merge(ws, r, c0, r, ncol - 1, '整　備　年　度', f(9, True), AC, fill('D9E1F2'))
    for i, y in enumerate(ys):
        put(ws, r + 1, c0 + i, f'R{wareki(y)}', f(9, True), AC, fill('FCE4D6' if y == V else 'D9E1F2'))
    merge(ws, r, ncol, r + 1, ncol, '備考', f(9, True), AC, fill('D9E1F2'))
    box(ws, r, 1, r + 1, ncol)
    r = 7
    no = 0
    total = {y: 0 for y in ys}
    for code, sname, f1, f2, cm in ordered(old_machines()):
        gr = r
        merge(ws, r, 2, r, 5, f'■ {sname}', f(10, True), AL, fill('EDEDED'))
        put(ws, r, 1, None, fl=fill('EDEDED'))
        put(ws, r, ncol, '（千円）', f(8), AL, fill('EDEDED'))
        r += 1
        gsum = {y: 0 for y in ys}
        prev = None
        for m in seq_of(f1, f2, cm):
            no += 1
            put(ws, r, 1, no, f(9), AC)
            nm = m['old'] or m['name']
            written = nm
            if prev is not None and prev['old'] == m['old'] and prev['furnace'] == 1 and m['furnace'] == 2:
                written = '　〃'
            put(ws, r, 3, written, f(9), AL)
            put(ws, r, 4, FUR[m['furnace']], f(9), AC)
            inst = m['inst'] if m['inst'] <= f'{V}-03' else '2001-03'
            iy = int(inst[:4])
            ins = 'H13' if inst == '2001-03' else (f"H{iy-1988}" if iy < 2019 else f"R{iy-2018}")
            put(ws, r, 5, ins, f(8), AC)
            mk = W.y1v(m, V) if m.get('tag') else PD.y1_marks(m)
            carry = mk.get('_carry', {})
            marks, comments = {}, {}
            for i, y in enumerate(ys):
                s = mk.get(y)
                act = actual_sym(m, y) if y < V else None
                if m.get('olddcs') and y < V:
                    act = None
                cell = None
                if s:
                    txt = s
                    done = None
                    if y < V:
                        done = W.compat(s, W.wtypes_in(m['tag'], y)) if m.get('tag') else (s == '★' and y == 2022 and False)
                        if s == '★' and m.get('olddcs'):
                            done = False
                    if done:
                        txt = s + '済'
                    cell = put(ws, r, c0 + i, txt, f(10, s == '★', 'C00000' if s == '★' else '000000'), AC, fill('FFF2CC') if s == '★' else None)
                    marks[str(y)] = txt
                    if y < V and done is False and m.get('tag'):
                        why = W.DEFER.get((m['tag'], y)) or W.NOTDONE_WHY.get((m['tag'], y)) or '未実施'
                        if why == '別契約':
                            why = '予算化見送り'
                        cell.comment = Comment(why, 'y1'); comments[f'{L(c0+i)}{r}'] = why
                    elif y in carry and y >= V - 0:
                        cell.comment = Comment(carry[y], 'y1'); comments[f'{L(c0+i)}{r}'] = carry[y]
                    gsum[y] += PD.y1_cost(m, s)
                elif act and act in ('整', '補', '更') and y < V:
                    sym = {'整': '◎', '補': '△', '更': '●'}[act]
                    cell = put(ws, r, c0 + i, sym + '済', f(10, color='0070C0'), AC)
                    cell.comment = Comment('計画外（定期整備等で実施）', 'y1')
                    marks[str(y)] = sym + '済'; comments[f'{L(c0+i)}{r}'] = '計画外（定期整備等で実施）'
            note = ''
            if m.get('kikan') or m.get('olddcs'):
                note = '基幹改良対象' + ('（撤去・更新）' if m.get('olddcs') else '')
            if m['status'] == '休止':
                note = 'H30より休止（灰資源化委託）'
            if m['cls'] == 'crane':
                note = '年次自主検査　' + ('常用' if '1号' in nm else '予備')
            put(ws, r, ncol, note, f(8), AL)
            works = [w['id'] for w in W.WORKS if m.get('tag') and m['tag'] in w['machines'] and w['year'] < V]
            rec.add(stem, ws.title, r, 'item', dict(path=[sname, nm], nameAsWritten=written if written != nm else None, no=no, furnace=FUR[m['furnace']],
                                                    install=ins, marks=marks, memo=note, comments=comments, code=m.get('tag')), works=works)
            prev = m
            r += 1
        for i, y in enumerate(ys):
            v = gsum[y]
            put(ws, gr, c0 + i, v if v else None, f(8, False, '1F3864', italic=True), AR, fill('EDEDED'), '#,##0')
            total[y] += v
        rec.add(stem, ws.title, gr, 'subtotal', dict(path=[sname], cost={str(y): v for y, v in gsum.items() if v}, unit='千円'))
        box(ws, gr, 1, r - 1, ncol)
    merge(ws, r, 1, r, 5, '合　計（千円）', f(10, True), AC, fill('D9E1F2'))
    for i, y in enumerate(ys):
        put(ws, r, c0 + i, total[y], f(9, True), AR, fill('D9E1F2'), '#,##0')
    put(ws, r, ncol, '基幹改良は別枠（継続費）', f(8), AL, fill('D9E1F2'))
    rec.add(stem, ws.title, r, 'total', dict(path=['合計'], cost={str(y): v for y, v in total.items()}, unit='千円', memo='基幹改良は別枠（継続費）'))
    box(ws, r, 1, r, ncol, med)
    r += 2
    merge(ws, r, 1, r, 5, '■ 法定点検等', f(10, True), AL, fill('EDEDED'))
    rec.add(stem, ws.title, r, 'parent', dict(path=['法定点検等']))
    r += 1
    legal = [
        ('ごみクレーン性能検査', (2022, 2024, 2026, 2028, 2030), '2年（検査証）'),
        ('受変電設備年次点検', tuple(ys), '電気保安協会委託'),
        ('計量機定期検査', (2021, 2023, 2025, 2027, 2029), '2年'),
        ('精密機能検査', (2021, 2024, 2027, 2030), '3年'),
        ('排ガス測定', tuple(ys), '年2回（DXN年1回）'),
    ]
    st = r
    for nm, yy, note in legal:
        merge(ws, r, 2, r, 5, nm, f(9), AL)
        marks = {}
        for i, y in enumerate(ys):
            if y in yy:
                t = '○' + ('済' if y < V else '')
                put(ws, r, c0 + i, t, f(10), AC)
                marks[str(y)] = t
        put(ws, r, ncol, note, f(8), AL)
        rec.add(stem, ws.title, r, 'item', dict(path=['法定点検等', nm], marks=marks, memo=note))
        r += 1
    box(ws, st - 1, 1, r - 1, ncol)
    ws.freeze_panes = 'F7'
    setup_print(ws, True, 8, 1, 0, f'A1:{L(ncol)}{r-1}', '5:6')
    ws.oddFooter.center.text = '&P / &N'
    wb.save(path)
    if V == 2021:
        GAPS.append(dict(file=fname, gap='法定点検等が 5 行だけ（DXN作業環境測定・第二種圧力容器・非常用発電機・消防用設備・ホイストが無い）。R3〜R5 版に継承'))
    return wb


# ======================================================================= y2
def y2_marks_v(m, V):
    """y2 の V 年度版の記号。R6 列は y1 R5 版の R6 計画（記号を読み替え）、R7 列は R6 改訂の計画"""
    base = PD.y2_marks(m)
    res = {}
    old = W.y1v(m, 2023).get(2024) if m.get('tag') else None
    if old and old != '★':
        res[2024] = (Y2SYM[old], PD.y1_cost(m, old), None)
    for y, v in base.items():
        if y >= 2025:
            res[y] = v
    if m['name'] == 'ごみクレーン（1号）':
        res[2024] = ('整', 700, 'ワイヤロープ')
    if V >= 2025:
        # R7 版の見直し：R7 の結果を受けて R8 以降を直す
        if m['tag'] == '2B-12':
            res[2026] = ('点', 0, '摺動板計測（R7残厚13mm）')
        if m['tag'] == '2F-09':
            res[2031] = ('整', 1900, '軸受交換（R7交換から6年）')
    return res


def build_y2(path, V):
    fname = os.path.basename(path)
    stem = rec.start(fname, plan=True)
    wb = Workbook(); ws = wb.active
    ws.title = '整備計画(R6改訂)' if V == 2024 else '整備計画(R7)'
    ttl = '令和6年度改訂版' if V == 2024 else '令和7年度版'
    props(wb, f'【架空】やまぶき広域処理組合 第二工場 長寿命化整備計画 年次スケジュール {ttl}', '組合事務局 y2' + ('（y1 R5版を改訂）' if V == 2024 else ''))
    ys = PD.Y2_YEARS
    c0 = 10
    ncol = c0 + len(ys)
    for i, w in enumerate([4.2, 8.5, 2.5, 22, 3.5, 4.5, 4.5, 7.5, 14] + [6.6] * len(ys) + [30], 1):
        ws.column_dimensions[L(i)].width = w
    merge(ws, 1, 1, 1, ncol - 2, f'【架空】やまぶき広域処理組合 第二工場　長寿命化整備計画　年次スケジュール（{ttl}）', f(14, True), AL)
    put(ws, 1, ncol, '令和5年度版', f(8, color='A6A6A6'), AR)
    if V == 2024:
        merge(ws, 2, 1, 2, 12, '令和6年12月改訂　組合事務局 施設係（y2）　※基幹的設備改良工事（R4〜R5）完了に伴い全面見直し', f(9), AL)
    else:
        merge(ws, 2, 1, 2, 12, '令和7年12月見直し　組合事務局 施設係（y2）　※R7 定期整備の結果を反映（R6 改訂版を更新）', f(9), AL)
    shade_txt = 'R6実績' if V == 2024 else 'R6・R7実績'
    merge(ws, 3, 1, 3, ncol, f'凡例　点＝点検・測定　整＝整備（分解・部品交換）　補＝補修　更＝更新　／　網掛け＝{shade_txt}　／　金額：号炉ごとの小計（千円・税込概算）、定期整備に含むものは計上しない', f(9), AL)
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
        put(ws, r + 2, c0 + i, f'R{wareki(y)}', f(8, True), AC, fill('FCE4D6') if y == V else hdr)
    merge(ws, r, ncol, r + 2, ncol, '備考', f(9, True), AC, hdr)
    box(ws, r, 1, r + 2, ncol)
    ws.row_dimensions[8].height = 6; ws.row_dimensions[9].height = 6
    r = 10
    oldno = {}
    n = 0
    for code, sname, f1, f2, cm in ordered(old_machines()):
        for m in seq_of(f1, f2, cm):
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
        rec.add(stem, ws.title, r, 'parent', dict(path=[sname]))
        gtop = r
        r += 1
        for label, grp in (('1号炉', f1), ('2号炉', f2), ('共通', cm)):
            if not grp:
                continue
            sr = r
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
                put(ws, r, 4, m['name'], f(9), AL)
                put(ws, r, 5, FUR[m['furnace']], f(8), AC)
                imp = PD.importance(m)
                put(ws, r, 6, imp, f(9, imp == 'A'), AC)
                put(ws, r, 7, '有' if m['standby'] else '', f(8), AC)
                put(ws, r, 8, m['inst'].replace('-', '.'), f(8), AC)
                past = ''
                if m.get('kikan'):
                    past = f"R{wareki(m['renew'])} 基幹改良で更新"
                else:
                    for y in (2023, 2022, 2021):
                        a = actual_sym(m, y)
                        if a in ('整', '補', '更'):
                            past = f'R{y-2018} ' + {'整': '整備', '補': '補修', '更': '更新'}[a]
                            break
                put(ws, r, 9, past, f(7.5), AL)
                mk = y2_marks_v(m, V)
                marks, comments = {}, {}
                memo_extra = []
                for i, y in enumerate(ys):
                    plan = mk.get(y)
                    if y < 2024 + (V - 2024) + 1 and y <= V:
                        act = actual_sym(m, y)
                        if act == '★':
                            act = '更'
                        if plan:
                            s, cost, note = plan
                            done = W.compat(s, W.wtypes_in(m['tag'], y))
                            cell = put(ws, r, c0 + i, s, f(9, s == '更', 'C00000' if s == '更' else '000000'), AC, fill('BFBFBF') if done else None)
                            marks[str(y)] = s
                            if not done:
                                why = W.DEFER.get((m['tag'], y)) or W.NOTDONE_WHY.get((m['tag'], y)) or '未実施'
                                if why == '別契約':
                                    why = '予算化を見送り、計画を見直し'
                                cell.comment = Comment(why, 'y2'); comments[f'{L(c0+i)}{r}'] = why
                            elif act and act != s and act in ('整', '補', '更') and s == '点':
                                pass
                            ssum[y] += cost
                        elif act:
                            cell = put(ws, r, c0 + i, act, f(9, act == '更', 'C00000' if act == '更' else '000000'), AC, fill('BFBFBF'))
                            marks[str(y)] = act
                            if act != '点':
                                cell.comment = Comment('計画外（実績）', 'y2'); comments[f'{L(c0+i)}{r}'] = '計画外（実績）'
                        continue
                    if plan:
                        s, cost, note = plan
                        put(ws, r, c0 + i, s, f(9, s == '更', 'C00000' if s == '更' else '000000'), AC)
                        marks[str(y)] = s
                        ssum[y] += cost
                notes = []
                if m['status'] == '休止':
                    notes.append('休止・残置（H30〜灰資源化委託）')
                for y in ys:
                    if y in mk and mk[y][2] and y > V and mk[y][0] in ('更', '整', '補', '点'):
                        if mk[y][0] == '点' and 'R7' not in (mk[y][2] or ''):
                            continue
                        notes.append(f"R{wareki(y)}{mk[y][2]}")
                        break
                if m['name'] == '計量機（出口）':
                    notes.insert(0, 'R6 ロードセル1点交換（器差）')
                if m['cls'] == 'crane':
                    notes.append('年次自主検査・性能検査は下段')
                memo = '　'.join(notes)[:40]
                put(ws, r, ncol, memo, f(7.5), AL)
                works = [w['id'] for w in W.WORKS if m['tag'] in w['machines'] and 2024 <= w['year'] <= V]
                rec.add(stem, ws.title, r, 'item', dict(path=[sname, label, m['name']], code=m['tag'], oldNo=on, furnace=FUR[m['furnace']],
                                                        importance=imp, standby='有' if m['standby'] else None, install=m['inst'].replace('-', '.'),
                                                        pastWork=past, marks=marks, memo=memo, comments=comments), works=works)
                r += 1
            for i, y in enumerate(ys):
                v = ssum[y]
                total[y] += v
                put(ws, sr, c0 + i, v if v else None, f(7.5, color='1F3864', italic=True), AR, fill('F2F2F2'), '#,##0')
            subtotal_rows.append((code, label, sr, ssum, sname))
        box(ws, gtop, 1, r - 1, ncol)
    for code, label, sr, ssum, sname in subtotal_rows:
        cost = {str(y): v for y, v in ssum.items() if v}
        if code == 'F' and label == '1号炉':
            cidx = c0 + ys.index(2027)
            cell = ws.cell(sr, cidx)
            cell.value = (cell.value or 0) + 600
            cost['2027'] = cell.value
            if V == 2024:
                GAPS.append(dict(file=fname, gap=f'通風設備【1号炉】小計 R9(2027) {L(cidx)}{sr} が明細の合計より +600（押込送風機の整備を R9→R10 へ移した時の直し忘れ）。R7 版にもそのまま継承'))
        rec.add(stem, ws.title, sr, 'subtotal', dict(path=[sname, label], cost=cost, unit='千円'))
    merge(ws, r, 1, r, 9, '合　計（千円）', f(10, True), AC, fill('DDEBF7'))
    for i, y in enumerate(ys):
        put(ws, r, c0 + i, total[y], f(8, True), AR, fill('DDEBF7'), '#,##0')
    put(ws, r, ncol, '基幹改良（R4〜R5）は含まない', f(8), AL, fill('DDEBF7'))
    rec.add(stem, ws.title, r, 'total', dict(path=['合計'], cost={str(y): v for y, v in total.items()}, unit='千円', memo='基幹改良（R4〜R5）は含まない'))
    box(ws, r, 1, r, ncol, med)
    r += 2
    merge(ws, r, 3, r, 9, '■ 法定点検・測定等', f(10, True), AL, fill('E2EFDA'))
    rec.add(stem, ws.title, r, 'parent', dict(path=['法定点検・測定等']))
    r += 1
    allys = set(ys)
    legal = [
        ('ごみクレーン年次自主検査（1・2号）', allys, '年1回　月次は運転委託'),
        ('ごみクレーン性能検査（1・2号）', {2024, 2026, 2028, 2030, 2032, 2034, 2036, 2038}, '2年（検査証の有効期間）'),
        ('ホイスト定期自主検査（4基）', allys, '年1回　R6追加'),
        ('受変電設備年次点検', allys, '電気保安協会委託　停電を伴う→全炉停止中'),
        ('非常用発電機点検（消防・保安）', allys, '機器6月・総合1年　R6追加'),
        ('第二種圧力容器定期自主検査（空気槽2基）', allys, '年1回　R6追加'),
        ('排ガス測定（ばい煙・DXN）', allys, 'ばい煙 年2回・DXN 年1回'),
        ('DXN類作業環境測定', allys, '6月以内ごと　R6追加'),
        ('計量機定期検査', {2025, 2027, 2029, 2031, 2033, 2035, 2037}, '2年'),
        ('消防用設備等点検', allys, '機器6月・総合1年　R6追加'),
        ('精密機能検査', {2024, 2027, 2030, 2033}, '3年'),
    ]
    st = r
    for nm, ss, note in legal:
        merge(ws, r, 5, r, 9, nm, f(9), AL)
        marks = {}
        for i, y in enumerate(ys):
            if y in ss:
                put(ws, r, c0 + i, '点', f(9), AC, fill('BFBFBF') if y <= V else None)
                marks[str(y)] = '点'
        put(ws, r, ncol, note, f(8), AL)
        rec.add(stem, ws.title, r, 'item', dict(path=['法定点検・測定等', nm], marks=marks, memo=note))
        r += 1
    box(ws, st - 1, 1, r - 1, ncol)
    put(ws, r, 4, '※第一種圧力容器：該当なし（ボイラ・蒸気設備なし）', f(8), ALN)
    if V == 2024:
        GAPS.append(dict(file=fname, gap='精密機能検査 R18(2036) が抜けている（R6・R9・R12・R15 の次）。R7 版にも継承'))
    ws.freeze_panes = ws.cell(10, c0)
    setup_print(ws, True, 8, 1, 0, f'A1:{L(ncol)}{r}', '5:7')
    ws.oddFooter.center.text = '&P / &N'
    ws.oddFooter.right.text = f'長寿命化整備計画（{ttl}）'
    # コピー元を非表示で残す
    srcs = [('R5版(y1)', build_y1(os.path.join(WORK, '_y1_R5.xlsx'), 2023, record=False))]
    if V == 2025:
        srcs.append(('R6改訂(y2)', _load(os.path.join(WORK, '_y2_R6.xlsx'))))
    for title, wb2 in srcs:
        old = wb.create_sheet(title)
        src = wb2.worksheets[0]
        for row in src.iter_rows():
            for c in row:
                if c.value is not None or c.has_style:
                    nc = old.cell(c.row, c.column, c.value)
                    nc.font = copy.copy(c.font); nc.alignment = copy.copy(c.alignment)
                    nc.fill = copy.copy(c.fill); nc.border = copy.copy(c.border); nc.number_format = c.number_format
        for rng in src.merged_cells.ranges:
            old.merge_cells(str(rng))
        for k, d in src.column_dimensions.items():
            old.column_dimensions[k].width = d.width
        old.sheet_state = 'hidden'
    wb.save(path)
    if V == 2024:
        import shutil
        shutil.copy(path, os.path.join(WORK, '_y2_R6.xlsx'))


def _load(p):
    from openpyxl import load_workbook
    return load_workbook(p)


_orig_build_y1 = build_y1


def build_y1(path, V, record=True):
    if record:
        return _orig_build_y1(path, V)
    saved = dict(rec.FILES), list(rec.LINKS), list(GAPS)
    wb = _orig_build_y1(path, V)
    rec.FILES.clear(); rec.FILES.update(saved[0]); rec.LINKS[:] = saved[1]; GAPS[:] = saved[2]
    return wb


def build_all():
    out = []
    for V in (2021, 2022, 2023):
        p = os.path.join(OUT, f'{FN}_中長期整備計画_令和{V-2018}年度版_y1様式.xlsx')
        build_y1(p, V); out.append(p)
    for V, nm in ((2024, '令和6年度改訂版'), (2025, '令和7年度版')):
        p = os.path.join(OUT, f'{FN}_中長期整備計画_{nm}_y2様式.xlsx')
        build_y2(p, V); out.append(p)
    return out


if __name__ == '__main__':
    os.makedirs(OUT, exist_ok=True); os.makedirs(WORK, exist_ok=True)
    print(build_all())
    print(GAPS)
