# -*- coding: utf-8 -*-
from openpyxl import Workbook
from openpyxl.comments import Comment
from xlutil import *
from f01_master import EQ, SYS, SYS_ORDER, OLD_DCS_2008, wareki
from f01_history import BYTAG
import f01_truth as TR
import ans
LSTEM = [None]

PFX = dict(A='受', B='燃', C='冷', D='排', E='通', F='灰', G='給', H='水', J='余', K='雑', L='電', M='計')
ZEN = str.maketrans('0123456789', '０１２３４５６７８９')

def inst_ym(d):
    y = d['inst']
    if y == 1996: return 'H9.3'
    m = {2000: 'H13.3', 2001: 'H14.2', 2010: 'H22.3', 2012: 'H25.3', 2014: 'H27.3', 2017: 'H30.3', 2018: 'H31.3'}
    return m.get(y, wareki(y))

def legal(d):
    p = d['prof']
    if d['sys'] == 'C' and d['kind'] == 'ボイラ': return '法定（ボイラ）'
    if p in ('crane_main', 'crane_sub', 'crane_ash', 'crane_small'): return 'クレーン則 年次'
    if p == 'vessel_small': return '第二種圧力容器'
    if p == 'wb': return '計量法'
    if p in ('pump_fire', 'firegun'): return '消防法'
    if d['sys'] == 'L' and d['kind'] in ('盤', '変圧器', '発電機'): return '保安規程'
    if d['prof'] in ('turbine', 'generator'): return '保安規程'
    return ''

# 2008 年時点の構成
def eq2008():
    rows = [d for d in EQ if d['inst'] <= 2008 or d['inst'] == 1996]
    rows = [d for d in rows if not (d['tag'].startswith('DCS') or d['tag'].startswith('OPS') or d['tag'] in ('EWS-01', 'DL-01'))]
    old = []
    for (s, k, n, sp, tg) in OLD_DCS_2008:
        f = 1 if tg == 'DCS-10' else (2 if tg == 'DCS-20' else 0)
        old.append(dict(sys=s, kind=k, name=n, spec=sp, kw=None, role='', tag=tg, prof='panel_dcs_old', imp='A', method='TBM', inst=1996, f=f, old=True))
    return rows, old

def history_note(d, upto):
    ev = d.get('act', {})
    notes = []
    for y in sorted(ev):
        if y > upto: continue
        for (k, t, c, ct) in ev[y]:
            if k in ('更新', '補修') and c >= 2000 or (k == '整備' and c >= 8000):
                notes.append('%s %s' % (wareki(y), t.replace('（基幹改良）', '')))
    return notes

def build_2008(nosave=False):
    TR.set_file('【架空】あさぎ市清掃工場_機器台帳_平成20年Excel化_a2様式.xlsx')
    LSTEM[0] = None if nosave else ans.start('【架空】あさぎ市清掃工場_機器台帳_平成20年Excel化_a2様式.xlsx')
    rows, old = eq2008()
    wb = Workbook(); ws0 = wb.active; ws0.title = '表紙'
    fn = 'ＭＳ Ｐゴシック'
    merge(ws0, 2, 2, 2, 8, 'あさぎ市清掃工場　機器台帳', F(18, True, name=fn), AC, border=None)
    merge(ws0, 4, 2, 4, 8, '平成20年10月　施設係（a2）', F(11, name=fn), AC, border=None)
    merge(ws0, 5, 2, 5, 8, '竣工図書（H9.3）機器表、DXN対策工事完成図書（H13.3・H14.2）より入力', F(10, name=fn), AC, border=None)
    put(ws0, 7, 3, '設備', F(10, True, name=fn), AC, B_ALL, fill('D9D9D9')); put(ws0, 7, 4, '1号炉', F(10, True, name=fn), AC, B_ALL, fill('D9D9D9'))
    put(ws0, 7, 5, '2号炉', F(10, True, name=fn), AC, B_ALL, fill('D9D9D9')); put(ws0, 7, 6, '共通', F(10, True, name=fn), AC, B_ALL, fill('D9D9D9')); put(ws0, 7, 7, '計', F(10, True, name=fn), AC, B_ALL, fill('D9D9D9'))
    widths(ws0, {'B': 2, 'C': 22, 'D': 9, 'E': 9, 'F': 9, 'G': 9, 'H': 2})
    allrows = rows + old
    r0 = 8
    for i, s in enumerate(SYS_ORDER):
        put(ws0, r0 + i, 3, SYS[s], F(10, name=fn), AL, B_ALL)
        for j, f in enumerate((1, 2, 0)):
            put(ws0, r0 + i, 4 + j, len([d for d in allrows if d['sys'] == s and d['f'] == f]) or None, F(10, name=fn), AC, B_ALL)
        put(ws0, r0 + i, 7, '=SUM(D%d:F%d)' % (r0 + i, r0 + i), F(10, name=fn), AC, B_ALL)
    re_ = r0 + len(SYS_ORDER)
    put(ws0, re_, 3, '合　計', F(10, True, name=fn), AC, B_ALL)
    for j, col in enumerate('DEFG'):
        put(ws0, re_, 4 + j, '=SUM(%s%d:%s%d)' % (col, r0, col, re_ - 1), F(10, True, name=fn), AC, box(t=dbl, b=thin, l=thin, r=thin))
    merge(ws0, re_ + 2, 3, re_ + 2, 8, '※台数は主機1台を1台とする（電動機は主機に含む）。電動弁・空気作動弁は1台ずつ、盤は1面1台。手動弁・計装品は載せない。', F(9, name=fn), AL, border=None)
    merge(ws0, re_ + 3, 3, re_ + 3, 8, '※設置年月は竣工・改造の完成月。', F(9, name=fn), AL, border=None)
    page(ws0, landscape=False, paper=9, fith=1, area='B2:H%d' % (re_ + 3))
    hdr = ['管理No', '機器番号', '名　　称', '数量', '形　式・仕　様', '', '電動機\n(kW)', '設置場所', '設置年月', '備　考']
    for s in SYS_ORDER:
        ws = wb.create_sheet(SYS[s].replace('設備', ''))
        cs = 2 if s == 'L' else 1   # 電気設備シートだけ 1 列右にずれている（別ファイルから貼付け）
        widths(ws, {CL(cs): 8, CL(cs + 1): 10, CL(cs + 2): 30, CL(cs + 3): 8.5, CL(cs + 4): 16, CL(cs + 5): 26, CL(cs + 6): 7, CL(cs + 7): 12, CL(cs + 8): 8, CL(cs + 9): 34})
        if cs == 2: widths(ws, {'A': 3})
        merge(ws, 1, cs, 1, cs + 9, 'あさぎ市清掃工場　機器台帳　【%s】' % SYS[s], F(14, True, name=fn), AL, border=None)
        merge(ws, 2, cs, 2, cs + 9, 'H20.10 作成（竣工図書 機器表より）　施設係 a2', F(9, name=fn, color='595959'), AL, border=None)
        r = 4
        for i, h in enumerate(hdr):
            put(ws, r, cs + i, h, F(10, True, name=fn), AC, B_ALL, fill('C0C0C0'))
        ws.merge_cells(start_row=r, start_column=cs + 4, end_row=r, end_column=cs + 5)
        ws.row_dimensions[r].height = 28
        r += 1
        cnt = 0
        for f in (1, 2, 0):
            grp = [d for d in allrows if d['sys'] == s and d['f'] == f]
            if not grp: continue
            merge(ws, r, cs, r, cs + 9, {1: '１号炉', 2: '２号炉', 0: '共　通'}[f], F(10, True, name=fn), AL, fl=fill('F2F2F2'))
            r += 1
            k = 0
            i = 0
            while i < len(grp):
                d = grp[i]
                k += 1; cnt += 1
                code = '%s-%d%02d' % (PFX[s], f, k)
                d['old_code'] = code
                nm = d['name']
                if s in ('C', 'D') and f in (1, 2) and d['kind'] == 'ポンプ':
                    nm = 'No.%d %s' % (f, nm)       # 竣工図書の書き方のまま
                elif s == 'E' and f in (1, 2):
                    nm = ('%d号' % f).translate(ZEN) + nm
                nm = nm.replace('（', ' （') if s == 'J' else nm
                if not d.get('old'):
                    TR.r(ws.title, r, 'ledger_row', code=d['code'], name=nm)
                if LSTEM[0]:
                    ans.rec(LSTEM[0], 'item', dict(kanri=code, tag=d['tag'], name=nm, qty='1台' if d['role'] != '予備' else '1台(予備)', inst=inst_ym(d)), sheet=ws.title, r=r, code=d.get('code'))
                put(ws, r, cs, code, F(9, name=fn), AC, B_ALL)
                put(ws, r, cs + 1, d['tag'], F(9, name=fn), AC, B_ALL)
                put(ws, r, cs + 2, nm + ('（撤去予定）' if False else ''), F(9, name=fn), AL, B_ALL)
                put(ws, r, cs + 3, '1台' if d['role'] != '予備' else '1台(予備)', F(9, name=fn), AC, B_ALL)
                sp = d['spec']
                kind_txt = sp.split(' ')[0]
                rest = ' '.join(sp.split(' ')[1:]) or '—'
                # A/B の対（常用・予備）は形式・仕様を結合して 1 回だけ書く（a2 の癖）
                pair = i + 1 < len(grp) and d['role'] == '常用' and grp[i + 1]['role'] == '予備' and grp[i + 1]['spec'] == d['spec']
                if pair:
                    merge(ws, r, cs + 4, r + 1, cs + 4, kind_txt, F(9, name=fn), AL)
                    merge(ws, r, cs + 5, r + 1, cs + 5, rest + '\n（A・B 同仕様）', F(9, name=fn), AL)
                    merge(ws, r, cs + 6, r + 1, cs + 6, d['kw'], F(9, name=fn), AC)
                elif not (i > 0 and d['role'] == '予備' and grp[i - 1]['role'] == '常用' and grp[i - 1]['spec'] == d['spec']):
                    put(ws, r, cs + 4, kind_txt, F(9, name=fn), AL, B_ALL)
                    put(ws, r, cs + 5, rest, F(9, name=fn), AL, B_ALL)
                    put(ws, r, cs + 6, d['kw'], F(9, name=fn), AC, B_ALL)
                put(ws, r, cs + 7, {'A': 'プラットホーム', 'B': '炉室', 'C': 'ボイラ室', 'D': '排ガス処理室', 'E': '炉室・送風機室', 'F': '灰出し室', 'G': 'ポンプ室',
                                    'H': '排水処理室', 'J': 'タービン室', 'K': '機械室', 'L': '電気室', 'M': '中央制御室'}[s], F(9, name=fn), AL, B_ALL)
                put(ws, r, cs + 8, inst_ym(d), F(9, name=fn), AC, B_ALL)
                notes = history_note(d, 2008) if not d.get('old') else []
                bik = '、'.join(notes[-2:])
                if d.get('old'):
                    bik = '旧DCS（竣工時）'
                if d['inst'] in (2000, 2001):
                    bik = ('DXN対策で新設　' + bik).strip()
                c = put(ws, r, cs + 9, bik, F(8, name=fn), AL, B_ALL)
                if len(notes) > 2:
                    c.comment = Comment('\n'.join(notes), 'a2')
                if d['tag'] in ('F-201', 'F-103', 'SH-201', 'BF-101'):
                    ws.cell(row=r, column=cs + 2).comment = Comment({'F-201': 'H18 軸受交換（定修）', 'F-103': 'H15 羽根車肉盛、要注意', 'SH-201': 'H16 管8本取替（補修工事）', 'BF-101': 'H19 ろ布全数交換（支給品）'}[d['tag']], 'a2')
                ws.row_dimensions[r].height = 24 if len(rest) > 22 else 15
                r += 1
                i += 1
            if s == 'D' and f in (1, 2):
                merge(ws, r, cs, r, cs + 9, '　（%d号炉 電気集じん器：H%s DXN対策工事で撤去。バグフィルタへ更新）' % (f, '13.3' if f == 1 else '14.2'), F(8, name=fn, color='808080'), AL)
                r += 1
            if s == 'J' and f == 0:
                r += 2   # 余分な空き 2 行（後で追記するつもりで空けた行）
        put(ws, r, cs + 2, '計', F(10, True, name=fn), AR, box(t=dbl, b=thin, l=thin, r=thin))
        put(ws, r, cs + 3, '%d台' % cnt, F(10, True, name=fn), AC, box(t=dbl, b=thin, l=thin, r=thin))
        page(ws, landscape=True, paper=9, area='%s1:%s%d' % (CL(cs), CL(cs + 9), r), title_rows='4:4')
        ws.freeze_panes = '%s5' % CL(cs)
    p = OUT + '/【架空】あさぎ市清掃工場_機器台帳_平成20年Excel化_a2様式.xlsx'
    if not nosave:
        wb.save(p)
    return p, len(allrows)

def build_2019():
    build_2008(nosave=True)  # old_code を付与
    LST = ans.start('【架空】あさぎ市清掃工場_機器台帳_令和元年度全面改訂_令和8年4月更新_a3様式.xlsx')
    TR.set_file('【架空】あさぎ市清掃工場_機器台帳_令和元年度全面改訂_令和8年4月更新_a3様式.xlsx')
    wb = Workbook(); ws = wb.active; ws.title = '機器台帳'
    fn = 'Meiryo UI'
    hdr = ['機器コード', '旧管理No\n(H20)', '設備', '機器種別', '号炉', '機器名称', '常用\n予備', '形式・仕様', '電動機\nkW', 'メーカー\n機器番号', '設置\n年度', '直近更新\n年度', '直近更新の内容', '重要度', '保全\n方式', '法定点検', '備考']
    W = [10, 8, 13, 11, 6, 30, 5, 44, 6, 10, 6, 7, 28, 5, 5, 15, 26]
    for i, w in enumerate(W):
        ws.column_dimensions[CL(i + 1)].width = w
    # a2 の受入供給シートをコピーして作った名残：1 行目の結合範囲（A:J）と 2 行目の注記
    merge(ws, 1, 1, 1, 10, 'あさぎ市清掃工場　機器台帳（全設備）', F(14, True, name=fn), AL, border=None)
    merge(ws, 2, 1, 2, 10, 'H20.10 作成（竣工図書 機器表より）　施設係 a2　→ R1.12 全面改訂（基幹改良反映・機器コード付番）a3　→ R8.4 更新 a3', F(9, name=fn, color='595959'), AL, border=None)
    put(ws, 2, 13, '台数', F(9, True, name=fn), AR); put(ws, 2, 14, '=COUNTA(A5:A600)-12', F(10, True, name=fn), AC, B_ALL)
    r = 4
    for i, h in enumerate(hdr):
        put(ws, r, i + 1, h, F(9, True, name=fn, color='FFFFFF'), AC, B_ALL, fill('305496'))
    ws.row_dimensions[r].height = 30
    r += 1
    for s in SYS_ORDER:
        merge(ws, r, 1, r, len(hdr), '■' + SYS[s], F(10, True, name=fn), AL, fl=fill('D9E1F2'))  # a2 の設備区切りを行で残した
        r += 1
        for d in [e for e in EQ if e['sys'] == s]:
            lastupd = [(y, t) for y, v in d['act'].items() for (k, t, c, ct) in v if k == '更新' and y <= 2025]
            lu = max(lastupd) if lastupd else None
            vals = [d['code'], d.get('old_code', '（新設）'), SYS[s], d['kind'], d['fname'], d['name'], d['role'], d['spec'], d['kw'], d['tag'],
                    1996 if d['inst'] == 1996 else d['inst'], lu[0] if lu else None, lu[1].replace('（基幹改良）', '（基幹）') if lu else None, d['imp'], d['method'], legal(d), '']
            if d['tag'].startswith('INV'):
                vals[16] = 'H29/H30 基幹改良で新設'
            if d['tag'] in ('OPS-01', 'OPS-02', 'OPS-03', 'EWS-01'):
                vals[16] = 'R9 保守終了予定（メーカー通知 R6.10）'
            if d['tag'] == 'F-201':
                vals[16] = 'R7 補修工事で軸受・軸受箱交換'
            if d['tag'] in ('FW-01', 'FW-02'):
                vals[16] = 'H21 補正予算で新設'
            TR.r('機器台帳', r, 'ledger_row', code=d['code'], name=d['name'])
            ans.rec(LST, 'item', dict(code=d['code'], kanri=vals[1], name=d['name'], role=d['role'] or None, spec=d['spec'], kw=d['kw'], tag=d['tag'], inst=vals[10], upd=vals[11], imp=d['imp'], method=d['method']), sheet='機器台帳', r=r, code=d['code'])
            for i, v in enumerate(vals):
                c = put(ws, r, i + 1, v, F(9, name=fn), AC if i in (0, 1, 4, 6, 8, 10, 11, 13, 14) else AL, Border(left=hair, right=hair, top=hair, bottom=hair))
                if i == 10:
                    c.number_format = '0'
            if d['inst'] == 1996:
                ws.cell(row=r, column=11).value = 1996
                ws.cell(row=r, column=11).comment = None
            if d['tag'].startswith('DCS') or d['tag'] in ('OPS-01', 'EWS-01'):
                ws.cell(row=r, column=2).comment = Comment('H24 DCS更新で旧CRT・ベンチボード・旧制御盤を撤去。旧管理No 計-0xx は欠番', 'a3')
            ws.row_dimensions[r].height = 13.5 if len(d['spec']) < 30 and len(d['name']) < 20 else 24
            r += 1
    ws.auto_filter.ref = 'A4:Q%d' % (r - 1)
    ws.freeze_panes = 'G5'
    page(ws, landscape=True, paper=8, area='A1:Q%d' % (r - 1), title_rows='4:4')
    # 凡例・数え方
    w2 = wb.create_sheet('凡例・数え方')
    widths(w2, {'A': 3, 'B': 18, 'C': 90})
    lines = [('数え方', '主機1台を1台（電動機は付属扱い）。電動弁・空気作動弁は各1台、盤は1面1台。手動弁・計装品は台帳に載せない（計装は計装リストで別管理）。'),
             ('常用／予備', '常用機と予備機は別の台として登録する。'),
             ('機器コード', '号炉（1／2／C=共通）＋設備記号＋連番。例 1E-013＝1号炉 通風設備 13番。設備記号 A受入供給 B燃焼 C燃焼ガス冷却 D排ガス処理 E通風 F灰出し G給水 H排水処理 J余熱利用 K雑 L電気 M計装制御'),
             ('重要度', 'A：故障で炉停止・処理能力低下・環境基準に直結（予備なし）　B：予備あり・切替可能、または停止しても短期は運転継続可能　C：上記以外'),
             ('保全方式', 'TBM：時間基準　CBM：状態基準（計測値で判断）　BM：事後保全'),
             ('設置年度', '竣工時の機器は 1996（平成8年度、H9.3 竣工）。'),
             ('旧管理No', 'H20 版（a2）の管理番号。H24 以降の新設は「（新設）」。')]
    for i, (k, v) in enumerate(lines):
        put(w2, 2 + i, 2, k, F(10, True, name=fn), AL, B_ALL, fill('D9E1F2')); put(w2, 2 + i, 3, v, F(10, name=fn), AL, B_ALL)
        w2.row_dimensions[2 + i].height = 30
    page(w2, landscape=True, paper=9, fith=1)
    # 変更履歴
    w3 = wb.create_sheet('変更履歴')
    widths(w3, {'A': 10, 'B': 8, 'C': 80})
    for i, h in enumerate(['日付', '担当', '内容']):
        put(w3, 1, i + 1, h, F(10, True, name=fn), AC, B_ALL, fill('D9E1F2'))
    hist = [('R1.12.20', 'a3', 'H20版（設備別12シート）を1シートに統合し全面改訂。機器コード付番、基幹改良（H29〜R1）の更新・新設を反映。445台。'),
            ('R2.6.1', 'a3', '直近更新年度の列を追加。'), ('R3.4.12', 'a3', '中長期整備計画を星取表に改めたのに合わせ、重要度・保全方式を見直し。'),
            ('R5.4.10', 'a3', 'R4 実績を反映。'), ('R6.4.8', 'a3', 'R5 実績を反映（タービン開放点検ほか）。'),
            ('R7.4.7', 'a3', 'R6 実績を反映。'), ('R8.4.13', 'a3', 'R7 実績を反映（電気室空調機No.1 更新、2号炉炉内監視カメラ更新ほか）。台数変更なし。')]
    for i, row in enumerate(hist):
        for j, v in enumerate(row):
            put(w3, 2 + i, j + 1, v, F(10, name=fn), AL, B_ALL)
    page(w3, landscape=True, paper=9, fith=1)
    p = OUT + '/【架空】あさぎ市清掃工場_機器台帳_令和元年度全面改訂_令和8年4月更新_a3様式.xlsx'
    wb.save(p)
    return p

if __name__ == '__main__':
    p, n = build_2008(); print('2008 total', n)
    print(render(p, pages=(1, 4)))
    p = build_2019()
    print(render(p, pages=(1, 3)))
