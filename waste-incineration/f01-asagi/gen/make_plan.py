# -*- coding: utf-8 -*-
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.chart import BarChart, Reference
from xlutil import *
from f01_master import EQ, SYS, SYS_ORDER, wareki, LEGAL, legal_years
from f01_history import PLAN13, PLAN21, PLAN26, H13, H26, tracks_for, actual_amounts, BYTAG, health_at
import make_ledger
import f01_truth as TR

PRI = ['更新', '補修', '整備', '計測', '点検']
SYMB = {'点検': '○', '計測': '◇', '整備': '◎', '補修': '△', '更新': '●'}
KFILL = {'点検': 'E2EFDA', '計測': 'E2EFDA', '整備': 'DDEBF7', '補修': 'FCE4D6', '更新': 'F4B084'}
LIFE = {'送風機': 25, 'ポンプ': 15, '電動弁': 20, '空気作動弁': 20, '盤': 20, 'コンベヤ': 15, '火格子': 15, 'ボイラ': 25, 'クレーン': 25, 'タービン': 30,
        '発電機': 30, 'ロータリバルブ': 15, 'スートブロワ': 15, 'バグフィルタ': 25, '減温塔': 25, '変圧器': 30, '空気圧縮機': 15}

def main_kind(evs):
    ks = [e[0] for e in evs]
    for k in PRI:
        if k in ks:
            return k

def build_a2():
    make_ledger.build_2008()
    TR.set_file('【架空】あさぎ市清掃工場_中長期整備計画_長寿命化総合計画_平成25年度策定_a2様式.xlsx')   # 管理No（H20）を付ける
    wb = Workbook()
    ws0 = wb.active; ws0.title = '表紙'
    fn = 'ＭＳ Ｐゴシック'
    for i, t in enumerate(['あさぎ市清掃工場', '長寿命化総合計画（施設保全計画）', '別表1 機器別整備計画　別表2 年度別整備費', '', '平成25年6月', 'あさぎ市 環境部 清掃工場（施設係 a2）']):
        merge(ws0, 4 + i * 2, 2, 4 + i * 2, 10, t, F(20 if i < 2 else 12, i < 2, name=fn), AC, border=None)
    merge(ws0, 17, 2, 20, 10, '計画期間：平成25年度〜平成46年度（延命化目標年度 平成46年度）\n基幹的設備改良工事：平成29〜31年度（継続費、循環型社会形成推進交付金を予定）\n金額は概算（千円、消費税込み）。定期整備業務委託の定額部分（共通仮設・清掃・法定点検等）は含まない。', F(10, name=fn), AL, border=None)
    widths(ws0, {'A': 3})
    page(ws0, landscape=True, paper=9, fith=1, area='A1:K21')
    ws = wb.create_sheet('別表1')
    years = list(range(2013, 2035))
    yc0 = 12
    W = {1: 4, 2: 5, 3: 4, 4: 28, 5: 8, 6: 6, 7: 5, 8: 4.5, 9: 5, 10: 4.5, 11: 5, }
    for k, v in W.items(): ws.column_dimensions[CL(k)].width = v
    for i in range(len(years)): ws.column_dimensions[CL(yc0 + i)].width = 7.2
    ctot = yc0 + len(years); crem = ctot + 1
    ws.column_dimensions[CL(ctot)].width = 8; ws.column_dimensions[CL(crem)].width = 30
    merge(ws, 1, 1, 1, 20, '別表1　機器別整備計画（あさぎ市清掃工場）', F(14, True, name=fn), AL, border=None)
    merge(ws, 2, 1, 2, 20, '単位：千円（概算・税込）　　平成25年6月策定　　健全度　4：健全　3：軽微な劣化　2：劣化進行（要注意）　1：要対策', F(9, name=fn), AL, border=None)
    lg = 2
    for k in ['点検', '整備', '補修', '更新']:
        put(ws, lg, 22 + (lg - 2) * 0, None)
    for i, k in enumerate(['点検・計測', '整備', '補修', '更新']):
        put(ws, 2, yc0 + 10 + i * 3, k, F(8, name=fn), AC, B_ALL, fill(KFILL[k.split('・')[0]]))
        ws.merge_cells(start_row=2, start_column=yc0 + 10 + i * 3, end_row=2, end_column=yc0 + 11 + i * 3)
    hr = 4
    hd = {1: '設備', 2: '号炉', 3: 'No', 4: '機　器　名', 5: '管理No', 6: '設置\n年度', 7: '経過\n年数', 8: '重要\n度', 9: '保全\n方式', 10: '健全\n度', 11: '目標\n耐用'}
    for c, h in hd.items():
        merge(ws, hr, c, hr + 1, c, h, F(8, True, name=fn), AC, fl=fill('D9D9D9'))
    merge(ws, hr, yc0, hr, yc0 + len(years) - 1, '年　度　別　整　備　計　画　（千円）', F(9, True, name=fn), AC, fl=fill('D9D9D9'))
    for i, y in enumerate(years):
        put(ws, hr + 1, yc0 + i, 'H%d' % (y - 1988), F(8, True, name=fn), AC, B_ALL, fill('D9D9D9' if y not in (2017, 2018, 2019) else 'FFD966'))
    merge(ws, hr, ctot, hr + 1, ctot, '計', F(8, True, name=fn), AC, fl=fill('D9D9D9'))
    merge(ws, hr, crem, hr + 1, crem, '主な整備内容・備考', F(8, True, name=fn), AC, fl=fill('D9D9D9'))
    r = hr + 2
    sub_rows = []
    for s in SYS_ORDER:
        sr0 = r
        for f in (1, 2, 0):
            grp = [d for d in EQ if d['sys'] == s and d['f'] == f]
            if not grp: continue
            fr0 = r
            for j, d in enumerate(grp, 1):
                new = d['inst'] > 2013 and d['inst'] != 1996
                put(ws, r, 3, j, F(8, name=fn), AC, B_ALL)
                nm = d['name'] + ('（新設予定）' if new else '')
                if s == 'E' and f in (1, 2) and not new:
                    nm = nm + '　'   # a2 が名称の後ろに全角スペースを入れる癖（台帳からの貼付け）
                TR.r('別表1', r, 'plan_row', code=d['code'], name=nm.strip())
                put(ws, r, 4, nm, F(8, name=fn, color='C00000' if new else '000000'), AL, B_ALL)
                put(ws, r, 5, d.get('old_code', '（H24新設）' if d['inst'] == 2012 else ('新設' if new else '')), F(8, name=fn), AC, B_ALL)
                put(ws, r, 6, '—' if new else ('H8' if d['inst'] == 1996 else 'H%d' % (d['inst'] - 1988)), F(8, name=fn), AC, B_ALL)
                put(ws, r, 7, None if new else 2013 - d['inst'], F(8, name=fn), AC, B_ALL)
                put(ws, r, 8, d['imp'], F(8, name=fn), AC, B_ALL)
                put(ws, r, 9, d['method'], F(8, name=fn), AC, B_ALL)
                h = None if new else (H13[d['code']] if d['method'] != 'BM' else '—')
                put(ws, r, 10, h, F(8, name=fn, bold=h in (1, 2), color='C00000' if h == 1 else '000000'), AC, B_ALL, fill('FFF2CC') if h == 2 else (fill('F8CBAD') if h == 1 else None))
                put(ws, r, 11, LIFE.get(d['kind'], 20), F(8, name=fn), AC, B_ALL)
                p = PLAN13[d['code']]
                for i, y in enumerate(years):
                    c = put(ws, r, yc0 + i, None, F(8, name=fn), AR, Border(left=hair, right=hair, top=hair, bottom=hair), fmt='#,##0;-#,##0;;')
                    if y in p:
                        c.value = sum(e[2] for e in p[y])
                        c.fill = fill(KFILL[main_kind(p[y])])
                        if '基幹' in ''.join(e[1] for e in p[y]):
                            c.font = F(8, True, name=fn)
                put(ws, r, ctot, '=SUM(%s%d:%s%d)' % (CL(yc0), r, CL(yc0 + len(years) - 1), r), F(8, name=fn), AR, B_ALL, fmt='#,##0;-#,##0;;')
                txt = '、'.join(sorted(set(e[1] for v in p.values() for e in v), key=lambda t: ('基幹' not in t, t)))
                if d['method'] == 'BM' and not p:
                    txt = '事後保全（故障時修理）'
                put(ws, r, crem, txt, F(7, name=fn), Alignment(horizontal='left', vertical='center', shrink_to_fit=True), B_ALL)
                if d['tag'] == 'F-103':
                    ws.cell(row=r, column=yc0 + 4).comment = Comment('インバータ化はCO2削減対策として交付金対象（3%以上削減の見込み）', 'a2')
                if d['tag'] == 'SH-201':
                    ws.cell(row=r, column=yc0 + 1).comment = Comment('H23定修報告書の提案（2次過熱器下段 減肉）による', 'a2')
                if d['tag'] == 'BF-101':
                    ws.cell(row=r, column=yc0).comment = Comment('ろ布は市で購入し支給（納期約4か月）', 'a2')
                r += 1
            merge(ws, fr0, 2, r - 1, 2, {1: '1号炉', 2: '2号炉', 0: '共通'}[f], F(8, name=fn), ACV)
        merge(ws, sr0, 1, r - 1, 1, SYS[s], F(9, True, name=fn), ACV)
        # 小計
        merge(ws, r, 1, r, 11, '%s　小計' % SYS[s], F(8, True, name=fn), AR, fl=fill('F2F2F2'))
        for i in range(len(years) + 1):
            col = CL(yc0 + i)
            put(ws, r, yc0 + i, '=SUM(%s%d:%s%d)' % (col, sr0, col, r - 1), F(8, True, name=fn), AR, box(t=thin, b=thin, l=hair, r=hair), fill('F2F2F2'), fmt='#,##0;-#,##0;;')
        put(ws, r, crem, None, None, None, B_ALL, fill('F2F2F2'))
        sub_rows.append(r)
        r += 1
        if s == 'L':
            r += 2   # 計装制御の前に空いた 2 行（a2 が DCS 更新の注記を入れるつもりだった）
            merge(ws, r - 1, 4, r - 1, 20, '※計装制御設備は H24 年度 DCS 更新済み（旧CRT・ベンチボード撤去）', F(8, name=fn, color='595959'), AL, border=None)
    merge(ws, r, 1, r, 11, '合　　計', F(9, True, name=fn), AC, fl=fill('D9D9D9'))
    for i in range(len(years) + 1):
        col = CL(yc0 + i)
        put(ws, r, yc0 + i, '=' + '+'.join('%s%d' % (col, x) for x in sub_rows), F(8, True, name=fn), AR, box(t=dbl, b=med, l=hair, r=hair), fill('D9D9D9'), fmt='#,##0')
    put(ws, r, crem, None, None, None, B_ALL, fill('D9D9D9'))
    total_row = r
    # 法令に基づく点検・検査・測定（別枠。合計には含めない）
    r += 2
    merge(ws, r, 1, r, crem, '法令に基づく点検・検査・測定（委託・受検。上の合計には含めない）', F(9, True, name=fn), AL, fl=fill('FFF2CC'))
    r += 1
    lr0 = r
    for j, it in enumerate(LEGAL, 1):
        put(ws, r, 3, j, F(8, name=fn), AC, B_ALL)
        merge(ws, r, 4, r, 9, it[1], F(8, name=fn), Alignment(horizontal='left', vertical='center', shrink_to_fit=True))
        merge(ws, r, 10, r, 11, it[2].split('（')[0].replace('労働安全衛生規則', '安衛則').replace('クレーン等安全規則', 'クレーン則'), F(7, name=fn), Alignment(horizontal='center', vertical='center', shrink_to_fit=True))
        ys = legal_years(it, 2013, 2034)
        for i, y in enumerate(years):
            c = put(ws, r, yc0 + i, (it[5] * it[6]) if y in ys else None, F(8, name=fn), AR, Border(left=hair, right=hair, top=hair, bottom=hair), fmt='#,##0;-#,##0;"(機器側)";')
            if y in ys:
                c.fill = fill('E2EFDA')
        put(ws, r, ctot, '=SUM(%s%d:%s%d)' % (CL(yc0), r, CL(yc0 + len(years) - 1), r), F(8, name=fn), AR, B_ALL, fmt='#,##0;-#,##0;;')
        put(ws, r, crem, it[8], F(7, name=fn), Alignment(horizontal='left', vertical='center', shrink_to_fit=True), B_ALL)
        r += 1
    merge(ws, lr0, 1, r - 1, 2, '法定', F(8, True, name=fn), ACV)
    merge(ws, r, 1, r, 11, '法定点検等　計', F(8, True, name=fn), AR, fl=fill('F2F2F2'))
    for i in range(len(years) + 1):
        col = CL(yc0 + i)
        put(ws, r, yc0 + i, '=SUM(%s%d:%s%d)' % (col, lr0, col, r - 1), F(8, True, name=fn), AR, box(t=thin, b=thin, l=hair, r=hair), fill('F2F2F2'), fmt='#,##0;-#,##0;;')
    put(ws, r, crem, None, None, None, B_ALL, fill('F2F2F2'))
    page(ws, landscape=True, paper=8, area='A1:%s%d' % (CL(crem), r), title_rows='%d:%d' % (hr, hr + 1))
    ws.freeze_panes = '%s%d' % (CL(yc0), hr + 2)
    # 別表2
    w2 = wb.create_sheet('別表2')
    merge(w2, 1, 1, 1, 12, '別表2　年度別整備費（千円・税込）', F(14, True, name=fn), AL, border=None)
    put(w2, 3, 1, '年度', F(9, True, name=fn), AC, B_ALL, fill('D9D9D9')); put(w2, 3, 2, '設備整備費（別表1）', F(9, True, name=fn), AC, B_ALL, fill('D9D9D9'))
    put(w2, 3, 3, 'うち基幹的設備改良', F(9, True, name=fn), AC, B_ALL, fill('D9D9D9')); put(w2, 3, 4, '定期整備委託 定額部分（見込）', F(9, True, name=fn), AC, B_ALL, fill('D9D9D9'))
    put(w2, 3, 5, '合計', F(9, True, name=fn), AC, B_ALL, fill('D9D9D9'))
    widths(w2, {'A': 8, 'B': 18, 'C': 18, 'D': 24, 'E': 14})
    for i, y in enumerate(years):
        rr = 4 + i
        put(w2, rr, 1, 'H%d' % (y - 1988), F(9, name=fn), AC, B_ALL)
        put(w2, rr, 2, "='別表1'!%s%d" % (CL(yc0 + i), total_row), F(9, name=fn), AR, B_ALL, fmt='#,##0')
        k = sum(e[2] for d in EQ for e in PLAN13[d['code']].get(y, []) if '基幹' in e[1])
        put(w2, rr, 3, int(k * 1.38) if k else None, F(9, name=fn), AR, B_ALL, fmt='#,##0')
        put(w2, rr, 4, 92000 + (y - 2013) * 1500, F(9, name=fn), AR, B_ALL, fmt='#,##0')
        put(w2, rr, 5, '=B%d+D%d+C%d*0.38/1.38' % (rr, rr, rr), F(9, name=fn), AR, B_ALL, fmt='#,##0')
    merge(w2, 4 + len(years) + 1, 1, 4 + len(years) + 1, 8, '※基幹的設備改良は機器費に共通仮設・諸経費（38%）を加えた額。', F(8, name=fn), AL, border=None)
    page(w2, landscape=False, paper=9, fith=1)
    p = OUT + '/【架空】あさぎ市清掃工場_中長期整備計画_長寿命化総合計画_平成25年度策定_a2様式.xlsx'
    wb.save(p)
    return p

def build_a3():
    make_ledger.build_2008()
    TR.set_file('【架空】あさぎ市清掃工場_中長期整備計画_年度別星取表_令和8年度版_a3様式.xlsx')
    wb = Workbook(); ws = wb.active; ws.title = '星取表（R8）'
    fn = 'Meiryo UI'
    years = list(range(2017, 2035))
    yc0 = 11
    W = {1: 4, 2: 5, 3: 4, 4: 28, 5: 8, 6: 4.5, 7: 5, 8: 5, 9: 5, 10: 4.5}
    for k, v in W.items(): ws.column_dimensions[CL(k)].width = v
    for i in range(len(years)): ws.column_dimensions[CL(yc0 + i)].width = 4.2
    crem = yc0 + len(years); ws.column_dimensions[CL(crem)].width = 34
    # a2 の別表1 をコピーして作り替えた名残：旧表題が 1 行目に残り、取消線で消してある
    merge(ws, 1, 1, 1, 20, '別表1　機器別整備計画（あさぎ市清掃工場）', F(10, name=fn, color='808080', strike=True), AL, border=None)
    merge(ws, 2, 1, 2, 30, 'あさぎ市清掃工場　年度別整備計画（星取表）　令和8年度版（R8.10 上期実績追記）', F(14, True, name=fn), AL, border=None)
    merge(ws, 3, 1, 3, 30, '○点検（○2＝年2回）　◇計測　◎整備（分解整備・部品交換）　△補修　●更新　／　上段：計画（R7以前は当初計画）　下段：実績（灰色）　▷ 施工中（R8 2号炉定修 R8.10.5〜）', F(9, name=fn), AL, border=None)
    hr = 5
    hd = {1: '設備', 2: '号炉', 3: 'No', 4: '機　器　名', 5: '機器コード', 6: '重要\n度', 7: '保全\n方式', 8: '健全度\n(R8)', 9: '目標耐用\n(旧)', 10: '区分'}
    for c, h in hd.items():
        merge(ws, hr, c, hr + 1, c, h, F(8, True, name=fn), AC, fl=fill('D9E1F2') if c != 9 else fill('EDEDED'))
    merge(ws, hr, yc0, hr, yc0 + 2, '平成', F(8, True, name=fn), AC, fl=fill('D9E1F2'))
    merge(ws, hr, yc0 + 3, hr, yc0 + len(years) - 1, '令和', F(8, True, name=fn), AC, fl=fill('D9E1F2'))
    for i, y in enumerate(years):
        lab = ('%d' % (y - 1988)) if y < 2019 else ('元' if y == 2019 else '%d' % (y - 2018))
        put(ws, hr + 1, yc0 + i, lab, F(8, True, name=fn), AC, B_ALL, fill('FFD966' if y == 2026 else 'D9E1F2'))
    merge(ws, hr, crem, hr + 1, crem, '備考', F(8, True, name=fn), AC, fl=fill('D9E1F2'))
    r = hr + 2
    grayf = fill('E7E6E6')
    for s in SYS_ORDER:
        sr0 = r
        for f in (1, 2, 0):
            grp = [d for d in EQ if d['sys'] == s and d['f'] == f]
            if not grp: continue
            fr0 = r
            for j, d in enumerate(grp, 1):
                merge(ws, r, 3, r + 1, 3, j, F(8, name=fn), AC)
                TR.r('星取表（R8）', r, 'plan_row', code=d['code'], name=d['name'])
                merge(ws, r, 4, r + 1, 4, d['name'], F(8, name=fn), AL)
                merge(ws, r, 5, r + 1, 5, d['code'], F(8, name=fn), AC)
                merge(ws, r, 6, r + 1, 6, d['imp'], F(8, name=fn), AC)
                merge(ws, r, 7, r + 1, 7, d['method'], F(8, name=fn), AC)
                h = H26[d['code']] if d['method'] != 'BM' else '—'
                merge(ws, r, 8, r + 1, 8, h, F(8, name=fn, bold=h in (1, 2)), AC, fl=fill('FFF2CC') if h == 2 else (fill('F8CBAD') if h == 1 else None))
                # 旧様式の目標耐用年数は a3 は更新していない（a2 の値が残る）
                merge(ws, r, 9, r + 1, 9, LIFE.get(d['kind'], 20) if d['inst'] <= 2013 or d['inst'] == 1996 else None, F(8, name=fn, color='A6A6A6'), AC)
                put(ws, r, 10, '計画', F(7, name=fn), AC, B_ALL); put(ws, r + 1, 10, '実績', F(7, name=fn), AC, B_ALL, grayf)
                for i, y in enumerate(years):
                    src = PLAN13 if y <= 2020 else (PLAN21 if y <= 2025 else PLAN26)
                    pe = src[d['code']].get(y, [])
                    ae = [(k, t, c) for (k, t, c, ct) in d['act'].get(y, [])]
                    c1 = put(ws, r, yc0 + i, SYMB[main_kind(pe)] if pe else None, F(9, name=fn), AC, Border(left=hair, right=hair, top=hair, bottom=dot))
                    c2 = put(ws, r + 1, yc0 + i, SYMB[main_kind(ae)] if ae else None, F(9, True, name=fn), AC, Border(left=hair, right=hair, top=dot, bottom=thin), grayf if y <= 2026 else None)
                    if ae:
                        TR.r('星取表（R8）', r + 1, 'plan_actual_cell:%d' % y, TR.work_ids(y, d['tag']), d['code'], col=yc0 + i)
                    if y == 2026 and d.get('pending2026'):
                        c2.value = '▷'; c2.font = F(8, name=fn, color='0070C0')
                    if y == 2026:
                        c1.border = Border(left=med, right=med, top=hair, bottom=dot); c2.border = Border(left=med, right=med, top=dot, bottom=thin)
                rem = ''
                notes = {'F-201': 'R6下期から振動上昇（R7.5 6.3mm/s）→R7 補修工事で軸受・軸受箱交換。以後CBM',
                         'GC-101': 'R7 下部コーン腐食（最小3.1mm）当て板。R9 部分更新',
                         'BF-101': 'R7 ろ布全数交換（予備12本 市倉庫）。次回 R13（廃止まで使用）',
                         'BF-201': 'R8 2号炉定修でろ布全数交換（施工中）',
                         'TB-01': '開放点検 4年ごと（R1基幹で分解整備）',
                         'ST-102': 'R7 最大摩耗4.2mm。R10 部分取替（約20%）',
                         'SH-201': 'H30更新。減肉 年0.04mm、廃止まで交換なし',
                         'SH-101': 'H29更新。R8 肉厚測定済',
                         'OPS-01': 'R9 OS・ハード保守終了 → HMI更新', 'DCS-01': 'R10 CPU・I/O 予備品化',
                         'F-103': 'H29基幹で更新・インバータ化。R5軸受交換', 'INV-101': 'R11 平滑コンデンサ交換（製造後12年）',
                         'CR-01': '常用機。ワイヤ 2〜3年', 'CR-02': '予備機。ワイヤ 5年程度', 'HV-01': '精密点検 3年ごと（全停時）'}
                rem = notes.get(d['tag'], '')
                if d['method'] == 'BM' and not any(PLAN26[d['code']].values()):
                    rem = rem or '事後保全'
                merge(ws, r, crem, r + 1, crem, rem, F(7, name=fn), AL)
                ws.row_dimensions[r].height = 12.5; ws.row_dimensions[r + 1].height = 12.5
                if d['tag'] == 'F-201':
                    ws.cell(row=r + 1, column=yc0 + years.index(2025)).comment = Comment('R7.10 補修工事（請負）。定修とは別契約', 'a3')
                    ws.cell(row=r, column=yc0 + years.index(2020)).comment = Comment('H25版計画の軸受交換。R2 は振動2.0mm/sで良好のため見送り', 'a3')
                r += 2
            merge(ws, fr0, 2, r - 1, 2, {1: '1号炉', 2: '2号炉', 0: '共通'}[f], F(8, name=fn), ACV)
        merge(ws, sr0, 1, r - 1, 1, SYS[s], F(9, True, name=fn), ACV)
    # 法令に基づく点検・検査・測定
    lr0 = r
    for j, it in enumerate(LEGAL, 1):
        merge(ws, r, 3, r + 1, 3, j, F(8, name=fn), AC)
        merge(ws, r, 4, r + 1, 5, it[1], F(7, name=fn), Alignment(horizontal='left', vertical='center', wrap_text=True))
        merge(ws, r, 6, r + 1, 9, it[2], F(7, name=fn), Alignment(horizontal='center', vertical='center', wrap_text=True, shrink_to_fit=False))
        put(ws, r, 10, '計画', F(7, name=fn), AC, B_ALL); put(ws, r + 1, 10, '実績', F(7, name=fn), AC, B_ALL, grayf)
        ys = legal_years(it, 2017, 2034)
        for i, y in enumerate(years):
            mark = ('○' if it[5] == 1 else '○%d' % it[5]) if y in ys else None
            c1 = put(ws, r, yc0 + i, mark, F(8, name=fn), AC, Border(left=hair, right=hair, top=hair, bottom=dot))
            am = mark if (y <= 2025 and y in ys) else None
            if y == 2026 and y in ys:
                am = '▷'
            c2 = put(ws, r + 1, yc0 + i, am, F(8, True, name=fn, color='0070C0' if am == '▷' else '000000'), AC, Border(left=hair, right=hair, top=dot, bottom=thin), grayf if y <= 2026 else None)
            if y == 2026:
                c1.border = Border(left=med, right=med, top=hair, bottom=dot); c2.border = Border(left=med, right=med, top=dot, bottom=thin)
        merge(ws, r, crem, r + 1, crem, it[8], F(7, name=fn), AL)
        ws.row_dimensions[r].height = 12.5; ws.row_dimensions[r + 1].height = 12.5
        r += 2
    merge(ws, lr0, 1, r - 1, 2, '法令に基づく点検・検査・測定', F(8, True, name=fn), ACV, fl=fill('FFF2CC'))
    page(ws, landscape=True, paper=8, area='A1:%s%d' % (CL(crem), r - 1), title_rows='%d:%d' % (hr, hr + 1))
    ws.freeze_panes = '%s%d' % (CL(yc0), hr + 2)
    # 年度別事業費（計画／実績）
    w2 = wb.create_sheet('年度別事業費')
    merge(w2, 1, 1, 1, 10, '年度別事業費（千円・税込）　計画：星取表の計画値の積上げ（R7以前は当初計画）　実績：契約金額', F(12, True, name=fn), AL, border=None)
    hd2 = ['年度', '計画（機器別積上げ）', '実績 定期整備業務委託', '実績 補修工事（請負）', '実績 基幹的設備改良', '実績 計', '備考']
    for i, h in enumerate(hd2):
        put(w2, 3, i + 1, h, F(9, True, name=fn), AC, B_ALL, fill('D9E1F2'))
    widths(w2, {'A': 8, 'B': 16, 'C': 16, 'D': 16, 'E': 16, 'F': 14, 'G': 40})
    for i, y in enumerate(years):
        rr = 4 + i
        lab = wareki(y)
        put(w2, rr, 1, lab, F(9, name=fn), AC, B_ALL)
        src = PLAN13 if y <= 2020 else (PLAN21 if y <= 2025 else PLAN26)
        pl = sum(e[2] for d in EQ for e in src[d['code']].get(y, []))
        put(w2, rr, 2, pl, F(9, name=fn), AR, B_ALL, fmt='#,##0')
        if y <= 2025:
            a = actual_amounts(y)
            put(w2, rr, 3, a['teishu'], F(9, name=fn), AR, B_ALL, fmt='#,##0')
            put(w2, rr, 4, a['hoshu'], F(9, name=fn), AR, B_ALL, fmt='#,##0')
            put(w2, rr, 5, a['kikan'] or None, F(9, name=fn), AR, B_ALL, fmt='#,##0')
            put(w2, rr, 6, '=SUM(C%d:E%d)' % (rr, rr), F(9, True, name=fn), AR, B_ALL, fmt='#,##0')
        else:
            for c in (3, 4, 5, 6):
                put(w2, rr, c, None, None, None, B_ALL)
        put(w2, rr, 7, {2017: '基幹改良 1年目（1号炉）', 2018: '基幹改良 2年目（2号炉）', 2019: '基幹改良 3年目（共通）', 2025: '補修工事4件（2号炉FDF軸受ほか）', 2026: 'R8：実績は年度末に記入'}.get(y, ''), F(8, name=fn), AL, B_ALL)
    re_ = 4 + len(years)
    merge(w2, re_ + 1, 1, re_ + 1, 7, '※計画は機器別の概算額で、定期整備委託の定額部分（共通仮設・清掃・法定点検対応等）を含まない。実績は契約金額のため両者は直接比較できない。物品（ろ布等の支給品）・運転委託・薬剤は含まない。', F(8, name=fn), AL, border=None)
    w2.row_dimensions[re_ + 1].height = 30
    ch = BarChart(); ch.type = 'col'; ch.grouping = 'clustered'; ch.title = '年度別事業費（千円）'; ch.height = 8; ch.width = 22
    ch.add_data(Reference(w2, min_col=2, min_row=3, max_row=re_ - 1), titles_from_data=True)
    ch.add_data(Reference(w2, min_col=6, min_row=3, max_row=re_ - 1), titles_from_data=True)
    ch.set_categories(Reference(w2, min_col=1, min_row=4, max_row=re_ - 1))
    w2.add_chart(ch, 'A%d' % (re_ + 3))
    page(w2, landscape=True, paper=9, fith=1, area='A1:G%d' % (re_ + 20))
    p = OUT + '/【架空】あさぎ市清掃工場_中長期整備計画_年度別星取表_令和8年度版_a3様式.xlsx'
    wb.save(p)
    return p

if __name__ == '__main__':
    p = build_a2(); print(render(p, pages=(2, 6, 18), dpi=60))
    p = build_a3(); print(render(p, pages=(1, 6), dpi=60))
