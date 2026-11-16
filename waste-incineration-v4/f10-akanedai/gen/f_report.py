# -*- coding: utf-8 -*-
"""工事報告書（文字のある PDF、reportlab）。
p1 様式（H23〜R1）：文章型。表紙→目次・概要→工程（表）→作業内容と点検結果（作業ごとの節）→使用部品→所見。
千草 報告書作成ツール R2 様式（R2〜）：1 枚目 系統別件数の集計表・総括・追加作業、2 枚目 機器シート索引、以降 1 機器 1 枚、最後に使用部品。"""
import datetime as dt
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.platypus import Paragraph, Spacer, Table, PageBreak
import world, history as H
from gen_report_base import Doc, P, S, tbl, photo, form_head
from rec import Rec

BYTAG = {}
def U(t):
    if not BYTAG: BYTAG.update({u['tag']: u for u in world.UNITS})
    return BYTAG[t]

def fmt(fy): return 'p1' if fy <= 2019 else 'tool'

def jd(d):
    return (f'平成{d.year-1988}' if d < dt.date(2019, 5, 1) else f'令和{d.year-2018}') + f'年{d.month}月{d.day}日'

def rtxt(r):
    return '—' if not r else f'{r[0].month}/{r[0].day}' + ('' if r[0] == r[1] else f'〜{r[1].month}/{r[1].day}')

KUBUN = {'○': '点検', '◇': '計測', '◎': '整備', '△': '補修', '●': '交換', '★': '交換'}
HANTEI = ['良', '要注意', '補修済', '要補修']

def unit_result(u, fy, it):
    s, d = H.truth(u).get(fy, ('○', ''))
    if it.get('sym') == '★': s = '★'; d = ''
    if d and d != '計画（R8）': res = d
    else: res = {'○': '異常なし', '◇': '前回と同程度', '◎': '整備後 試運転 良', '△': '補修後 良', '●': '更新後 試運転 良', '★': '更新（据付・試験）完了'}.get(s, '異常なし')
    hj = '補修済' if s in ('△', '◎') and ('補修' in res or '交換' in res or '漏水' in res) else ('要注意' if ('mm' in res and u['cat'] == 'sh2') or 'ベンド' in res else '良')
    if it['note'] and '施工中' in it['note']: hj = '—'
    return s, res, hj

def parts_of(p):
    out = []
    for it in p['items']:
        if not it['act']: continue
        if '噴霧ノズル交換' in it['name']: out.append(['減温塔噴霧ノズル', '二流体ノズル', '6 本', '委託範囲', it['name']])
        if 'ろ布全数交換' in it['name']: out.append(['ろ布', 'PTFE', '432 本', '市支給（物品購入）', it['name']])
        if '誘引通風機 分解' in it['name']: out.append(['軸受', '自動調心ころ軸受', '2 個', '委託範囲', it['name']])
        if it['tmpl'] == 'tap' and it.get('sym') == '△': out.append(['SiC れんが', '出湯口用', '20〜40 個', '委託範囲', it['name']])
        if it['tmpl'] == 'refr' or it['no'] == '6': out.append(['吹付材', 'アルミナ質', '1〜2 t', '委託範囲', it['name']])
        if '羽口交換' in it['name']: out.append(['主羽口', '水冷銅羽口', '1〜2 本', '市支給（予備品）', it['name']])
    out.append(['パッキン・ガスケット類', '—', '一式', '委託範囲', '各マンホール'])
    return out

def write(path, p, title, staff, book_label, city, extra_items=None, info_extra=None):
    f = p.get('rfmt') or fmt(p['fy'])
    name = path.split('/')[-1]
    rc = Rec(name)
    doc = Doc(path, f'【架空】{title} {"業務完了報告書" if "委託" in title else "工事完了報告書"}（{book_label}）', f'千草環境システム株式会社 {staff}',
              f'【架空】{title}　{book_label}', None if f == 'p1' else f'報告 {p["key"]}')
    seq = [0]
    def R(kind, fields, works=None):
        rc.add('', None, kind, fields, works, seq=seq[0]); seq[0] += 1
    items = [it for it in p['items'] if it['co'] != '市運転班']
    st = []
    if f == 'p1':
        st += [Spacer(1, 40 * mm), P('【架空】', S(12, al=1)), P(title, S(18, 28, al=1)), Spacer(1, 8 * mm),
               P('業 務 完 了 報 告 書' if '委託' in title else '工 事 完 了 報 告 書', S(22, 30, al=1)), Spacer(1, 4 * mm), P(book_label, S(14, al=1)),
               Spacer(1, 50 * mm), P(jd(p['end'] + dt.timedelta(days=21)), S(12, al=1)), Spacer(1, 6 * mm), P('あかね台市長　様', S(12, al=1)),
               Spacer(1, 16 * mm), P(f'千草環境システム株式会社<br/>あかね台事業所　SV {staff}', S(12, 20, al=1)), PageBreak()]
        st += [P('1. 業務概要', S(13)), Spacer(1, 2 * mm)]
        info = [['項目', '内容'], ['件名', title], ['本分冊の範囲', f"{p['part']}（{jd(p['start'])} 〜 {jd(p['end'])}）"], ['受注者', '千草環境システム株式会社'],
                ['現場責任者', f'SV {staff}']] + (info_extra or [])
        st.append(tbl(info, [35, 140]))
        st += [Spacer(1, 5 * mm), P('2. 実施工程', S(13)), Spacer(1, 2 * mm)]
        rows = [['No.', '作業項目', '予定', '実績', '備考']]
        for it in p['items']:
            rows.append([it['no'], it['name'], rtxt(it['plan']), rtxt(it['act']), it['note']])
            R('item', dict(section='実施工程', no=it['no'], name=it['name'], plan=rtxt(it['plan']), act=rtxt(it['act']), note=it['note']), it.get('wids') if it['co'] != '市運転班' else None)
        st.append(tbl(rows, [12, 70, 24, 24, 46]))
        st += [P('※ 立下げ・乾燥・立上げは市運転班による（当社立会）。', S(8)), PageBreak(), P('3. 作業内容および点検結果', S(13)), Spacer(1, 3 * mm)]
        for i, it in enumerate(items, 1):
            us = [U(t) for t in it['units']]
            parts = []
            for u in us[:12]:
                s, res, hj = unit_result(u, p['fy'], it)
                parts.append(f"{u['name']}は{res}")
            body = f"{rtxt(it['act'] or it['plan'])}に{it['name']}を行った（{it['co']}）。" + '。'.join(parts) + '。' + (f"なお、{it['note']}。" if it['note'] else '')
            st += [P(f'3.{i} {it["name"]}', S(10.5)), P(body, S(9.5, 16)), Spacer(1, 2 * mm)]
            R('item', dict(section=f'3.{i}', name=it['name'], date=rtxt(it['act'] or it['plan']), text=body), it.get('wids'))
        st += [PageBreak(), P('4. 使用部品・材料', S(13)), Spacer(1, 2 * mm)]
        pr = parts_of(p)
        st.append(tbl([['品名', '仕様', '数量', '区分', '使用箇所']] + pr, [34, 36, 18, 36, 54]))
        for x in pr: R('item', dict(section='使用部品', name=x[0], spec=x[1], qty=x[2], kind=x[3], place=x[4]))
        st += [Spacer(1, 6 * mm), P('5. 所見', S(13))]
        fnds = [f"{U(t)['name']}：{H.truth(U(t))[p['fy']][1]}" for it in items for t in it['units'] if H.truth(U(t)).get(p['fy'], ('', ''))[1] not in ('', '計画（R8）')]
        for x in (fnds[:8] or ['特記すべき不具合はなかった。']):
            st.append(P('・' + x, S(9.5, 16))); R('item', dict(section='所見', text=x))
        doc.build(st)
        return rc
    # ---- ツール様式
    machines = []
    for it in items:
        for t, w in zip(it['units'], it.get('wids', [None] * len(it['units']))):
            u = U(t); s, res, hj = unit_result(u, p['fy'], it)
            machines.append(dict(u=u, it=it, s=s, res=res, hj=hj, w=w))
    form_head(st, '業務完了報告書　集計表' if '委託' in title else '工事完了報告書　集計表', '千草環境システム 報告書作成ツール R2 ／ 様式 CK-R-01')
    info = [('件名', title), ('範囲', f"{p['part']}（{book_label}）"), ('実施期間', f"{jd(p['start'])} 〜 {jd(p['end'])}" + ('（施工中、11月13日まで）' if p.get('status') == '施工中' else '')),
            ('受注者', f'千草環境システム株式会社　{staff}'), ('報告日', jd(min(p['end'] + dt.timedelta(days=21), dt.date(2026, 11, 13))))] + [tuple(x) for x in (info_extra or [])]
    st.append(tbl([[k, v] for k, v in info], [32, 146], head=0, sz=8.5, extra=[('BACKGROUND', (0, 0), (0, -1), colors.HexColor('#E8F5E9'))]))
    st += [Spacer(1, 3 * mm), P('1. 系統別 実施件数（作業区分別・判定別）', S(10))]
    systems = []
    for m in machines:
        if m['u']['sysname'] not in systems: systems.append(m['u']['sysname'])
    K = ['点検', '計測', '整備', '補修', '交換']
    data = [['系統', '機器数'] + K + HANTEI]; tot = [0] * (1 + len(K) + len(HANTEI))
    for sy in systems:
        xs = [m for m in machines if m['u']['sysname'] == sy]
        row = [len(xs)] + [sum(1 for m in xs if KUBUN.get(m['s']) == k) for k in K] + [sum(1 for m in xs if m['hj'] == h) for h in HANTEI]
        tot = [a + b for a, b in zip(tot, row)]
        data.append([sy] + [v or '' for v in row])
        R('item', dict(section='系統別件数', system=sy, **{'機器数': row[0]}, **{k: v for k, v in zip(K + HANTEI, row[1:]) if v}))
    data.append(['計'] + tot); R('total', dict(section='系統別件数', **{'機器数': tot[0]}))
    st.append(tbl(data, [40, 14] + [12] * len(K) + [15] * len(HANTEI), sz=8, extra=[('BACKGROUND', (0, len(data) - 1), (-1, len(data) - 1), colors.HexColor('#F2F2F2'))]))
    st.append(P('※ 1 機器 1 件。判定は機器ごとに 1 つ。', S(7.5)))
    adds = [it for it in p['items'] if it.get('added') or (it['note'] and it['plan'] != it['act'])]
    st += [Spacer(1, 4 * mm), P('2. 総括', S(10)),
           P(f"{p['part']}を{'予定の期間内に終えた' if not adds else '一部の作業を延長して終えた'}。" + ('施工中のため、本報告は 11月13日までの実施分である。' if p.get('status') == '施工中' else ''), S(9, 15))]
    R('item', dict(section='総括', text=f"{p['part']}"))
    if adds or extra_items:
        rows = [['No.', '内容', '扱い']]
        for i, it in enumerate(adds, 1):
            rows.append([f'A{i}', f"{it['name']}（{rtxt(it['act'])}）{it['note']}", '追加' if it.get('added') else '工程内で吸収'])
            R('item', dict(section='追加作業', no=f'A{i}', text=rows[-1][1], kind=rows[-1][2]), it.get('wids'))
        for x in (extra_items or []):
            rows.append(x); R('item', dict(section='追加作業', no=x[0], text=x[1], kind=x[2]))
        st += [Spacer(1, 3 * mm), P('3. 追加作業・特記事項', S(10)), tbl(rows, [10, 128, 40], sz=7.5)]
    st.append(PageBreak())
    form_head(st, '機器シート索引', '様式 CK-R-01 続紙')
    half = (len(machines) + 1) // 2
    rows = [['シート', '機器名称', '判定', 'シート', '機器名称', '判定']]
    for i in range(half):
        a = machines[i]; b = machines[i + half] if i + half < len(machines) else None
        rows.append([f'M-{i+1:02d}', a['u']['name'], a['hj']] + ([f'M-{i+half+1:02d}', b['u']['name'], b['hj']] if b else ['', '', '']))
    st.append(tbl(rows, [14, 60, 15, 14, 60, 15], sz=7.5))
    st.append(PageBreak())
    hd = f'【架空】{title}'
    for n, m in enumerate(machines, 1):
        u, it = m['u'], m['it']
        form_head(st, f'機器シート　M-{n:02d}', '様式 CK-R-03（1 機器 1 枚）')
        rec = [['機器名称', u['name']], ['系統図番号', u['tag']], ['系統', u['sysname']], ['実施日', rtxt(it['act'] or it['plan'])],
               ['作業区分', KUBUN.get(m['s'], '点検')], ['作業内容', it['name']], ['結果', m['res']], ['判定', m['hj']], ['担当', it['co']]]
        left = [tbl(rec, [22, 83], head=0, sz=8, extra=[('BACKGROUND', (0, 0), (0, -1), colors.HexColor('#E8F5E9'))])]
        kind = 'pipe' if u['cat'] in ('sh', 'sh2', 'boiler') else 'tuyere' if u['cat'] == 'tuyere' else 'brick' if u['cat'] in ('furnace', 'taphole', 'furnace2') else 'box'
        right = [photo(f"{u['name']} 作業前", kind, w=64, h=58), Spacer(1, 3 * mm), photo(f"{u['name']} 作業後", 'box', w=64, h=58)]
        st += [Table([[left, right]], colWidths=[108 * mm, 70 * mm], style=[('VALIGN', (0, 0), (-1, -1), 'TOP')]), Spacer(1, 3 * mm), P(hd, S(7, textColor=colors.grey)), PageBreak()]
        R('item', dict(section='機器シート', sheet=f'M-{n:02d}', unit=u['name'], code=u['tag'], date=rtxt(it['act'] or it['plan']), kubun=KUBUN.get(m['s'], '点検'),
                       work=it['name'], result=m['res'], judge=m['hj'], company=it['co']), [m['w']] if m['w'] else None)
        R('item', dict(section='写真', sheet=f'M-{n:02d}', image='写真代替図（線画）', caption=f"{u['name']} 作業前"))
        R('item', dict(section='写真', sheet=f'M-{n:02d}', image='写真代替図（線画）', caption=f"{u['name']} 作業後"))
    form_head(st, '使用部品・材料', '様式 CK-R-04')
    pr = parts_of(p)
    st.append(tbl([['品名', '仕様', '数量', '区分', '使用箇所']] + pr, [34, 36, 18, 36, 54]))
    for x in pr: R('item', dict(section='使用部品', name=x[0], spec=x[1], qty=x[2], kind=x[3], place=x[4]))
    doc.build(st)
    return rc
