# -*- coding: utf-8 -*-
"""工事報告書（文字のある PDF。Word で書いて LibreOffice で PDF）。
定期整備：R3 は旧章立て、R4〜R7 は基幹改良の報告書に合わせた新章立て（R6・R7 は n2）。基幹改良：完成報告書（n3）。
ページをまたぐ表は見出し行をくり返す。原価は書かない。"""
import os, datetime as dt
from collections import OrderedDict
from docx.shared import Mm
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from common import *
from world import MACHINES, SYSNAME
import works as W
import schedules as S
import gen_report as G
import render
import rec

MBY = W.MBY
D = dt.date


class Doc:
    """段落・表の行を書きながら、正解の seq を振る"""

    def __init__(self, stem, d):
        self.stem, self.d = stem, d
        self.sec = ''

    def head(self, t, lvl=1):
        G.head(self.d, t, lvl)
        if lvl == 1:
            self.sec = t

    def p(self, t, works=None, **kw):
        G.para(self.d, t, **kw)
        rec.add(self.stem, None, None, 'item', dict(章=self.sec, 文=t), seq=rec.next_seq(self.stem), works=works)

    def note(self, t, **kw):
        G.para(self.d, t, **kw)   # 注記：数えない

    def table(self, hdr, rows, widths, size=9, works=None):
        G.table(self.d, hdr, rows, widths, size)
        for i, row in enumerate(rows):
            rec.add(self.stem, None, None, 'item', dict(章=self.sec, **{h: v for h, v in zip(hdr, row) if v not in (None, '')}),
                    seq=rec.next_seq(self.stem), works=(works[i] if works else None))


def fmt(sp):
    return '・'.join(f'{s.month}/{s.day}' + (f'～{e.month}/{e.day}' if e != s else '') for s, e in sp)


def works_of(year, ev=None):
    return [w for w in W.WORKS if w['contract'] == f'T{year}' and (ev is None or w['event'] == ev)]


def cover(dc, title, name, period, sub, date_txt):
    d = dc.d
    G.para(d, '【架空】試験用サンプル', 8, align='r', color='808080')
    for _ in range(5):
        G.para(d, '')
    G.para(d, title, 22, True, 'c', after=18)
    G.para(d, name, 13, True, 'c', after=6)
    G.para(d, sub, 11, align='c', after=40)
    t = d.add_table(rows=4, cols=2); t.style = 'Table Grid'; t.alignment = WD_TABLE_ALIGNMENT.CENTER
    vals = (('施 設 名', FAC), ('工　　期', period), ('発 注 者', f'{UNION}　管理者'), ('受 注 者', MAKER))
    for i, (a, b) in enumerate(vals):
        t.cell(i, 0).text = ''; t.cell(i, 1).text = ''
        G.setfont(t.cell(i, 0).paragraphs[0].add_run(a), 11, True)
        G.setfont(t.cell(i, 1).paragraphs[0].add_run(b), 11)
    t.autofit = False; t.columns[0].width = Mm(35); t.columns[1].width = Mm(105)
    for i in range(4):
        t.cell(i, 0).width = Mm(35); t.cell(i, 1).width = Mm(105)
    for a, b in vals:
        rec.add(dc.stem, None, None, 'item', dict(章='表紙', 項目=a.replace(' ', '').replace('　', ''), 内容=b), seq=rec.next_seq(dc.stem))
    for _ in range(8):
        G.para(d, '')
    G.para(d, date_txt, 12, align='c')
    G.para(d, MAKER, 12, True, 'c')
    d.add_page_break()


def ev_period(year, ev):
    down, w0, w1, up = W.STOPS[year][ev]
    return f'令和{year-2018}年{w0.month}月{w0.day}日～{w1.month}月{w1.day}日（停止 {down.month}/{down.day} 立下げ～{up.month}/{up.day} 立上げ）'


def sys_of(w):
    if not w['machines']:
        return None
    return MBY[w['machines'][0]]['sys']


def wtext(w):
    t = w['name']
    if w['finding']:
        t += f"。{w['finding']}"
    return t


def measures(year):
    rows, ws = [], []
    for w in works_of(year):
        if w['measure']:
            m = MBY[w['machines'][0]] if w['machines'] else None
            a, b, c, dd = w['measure']
            rows.append((m['tag'] if m else '', W.short(m) if m else w['name'], a, b, c, dd))
            ws.append([w['id']])
    return rows, ws


def parts(year):
    rows, ws = [], []
    for w in works_of(year):
        for p in w['parts']:
            rows.append(p + ('',) if len(p) == 5 else p)
            ws.append([w['id']])
    return rows, ws


# ---------------------------------------------------------------- R3（旧章立て）
def build_r3():
    year = 2021
    name = W.TEIKI[year]['name']
    fname = f'{FN}_工事報告書_令和3年度定期整備業務委託_n1旧章立て.pdf'
    stem = rec.start(fname)
    d = G.newdoc(f'【架空】{name} 工事報告書', f'{MAKER} n1')
    dc = Doc(stem, d)
    cover(dc, '工 事 報 告 書', name, '令和3年4月26日 ～ 令和4年3月18日', '（1号炉中間点検・2号炉中間点検・全炉停止整備）', '令和3年12月10日')
    dc.head('1．工事概要')
    dc.table(['項目', '内容'], [('工事件名', name), ('施設', f'{FAC}（30t/16h×2炉、並行揺動式ストーカ、准連続）'), ('工期', '令和3年4月26日～令和4年3月18日'),
                              ('1号炉中間点検', ev_period(year, 'mid1')), ('2号炉中間点検', ev_period(year, 'mid2')), ('全炉停止整備', ev_period(year, 'all')),
                              ('現場責任者', 'SV n1')], [40, 126])
    dc.head('2．作業内容')
    for i, ev in enumerate(('mid1', 'mid2', 'all'), 1):
        dc.head(f'2.{i}　{W.EVLABEL[ev]}', 2)
        for w in works_of(year, ev):
            if w['kubun'] in ('全体', '準備') and '試運転' not in w['name']:
                continue
            dc.p('・' + wtext(w) + '。' if not wtext(w).endswith('。') else '・' + wtext(w), works=[w['id']], indent=4)
    dc.head('3．点検・測定結果')
    rows, ws = measures(year)
    dc.table(['機器番号', '機器', '項目', '基準・前回', '結果', '判定'], rows, [14, 34, 32, 30, 32, 24], 8.5, works=ws)
    dc.head('4．所見')
    for t in ['1号炉・2号炉とも燃焼段の火格子片の摩耗が進んでいる。令和4〜5年度の基幹的設備改良工事で更新が予定されているため、本年度は脱落・割れ片の取替えにとどめた。',
              '1号炉減温塔の下部胴は限界厚さを下回る箇所があり当て板で補修したが、令和4年度の更新までの応急処置である。運転中は外面の点検を強化されたい。',
              '2号炉バグフィルタの差圧は上昇傾向にある。ろ布は平成28年度の全数交換から5年を経過しており、更新（令和5年度予定）までの間、破損ろ布の交換で対応する。予備ろ布の在庫は 22 本（本年度 3 本使用）。',
              '1号 BF ダスト排出装置の整備は予算の都合で次年度へ送られた。',
              'ごみクレーン1号の走行車輪にわずかな摩耗が見られる。次年度以降も計測を継続する。']:
        dc.p(t, indent=4, after=6)
    dc.head('5．使用材料')
    rows, ws = parts(year)
    dc.table(['品名', '仕様', '数量', '手配', '使用箇所', '備考'], rows, [32, 42, 16, 22, 38, 16], 8.5, works=ws)
    dc.note('※ 写真は別冊「工事写真帳」による。', size=9)
    p = os.path.join(WORK, fname.replace('.pdf', '.docx'))
    d.save(p)
    return p, fname


# ---------------------------------------------------------------- R4〜R7（新章立て）
TEXT = {
    2022: dict(
        sum='本年度は、基幹的設備改良工事（別契約）で 1号炉が 9 月から 12 月まで停止しているため、1号炉の中間点検は行わず、2号炉中間点検（6月）と全炉停止整備（10月）を実施した。全炉停止整備は改良工事の現地工事期間中に行い、停止期間・炉室ホイスト・全停電日を改良工事と調整した。主な結果は下表のとおりである。',
        coord='基幹的設備改良工事（現場代理人 n3）：2号炉側の立入は SV の許可制。炉室ホイストは 10/20～22 本工事優先。10/23 全停電（受変電年次点検）は両工事とも作業中止。',
        special=[('(1) 2号炉 灰押出装置 摺動板取替（追加）', '10/19 の点検で摺動板の著しい摩耗と穴あきを確認し、10/20 に組合と協議して追加施工とした（工程表 改1）。摺動板はメーカー在庫品を手配し 10/26 に着荷、10/26～27 に取替え、10/28 に水張り・作動確認を行った。本件は変更契約（令和4年11月）の対象である。'),
                 ('(2) 1号炉中間点検', '1号炉は基幹的設備改良工事のため 9/20 から停止しており、火格子・減温塔・バグフィルタは同工事で更新される。このため本業務の 1号炉中間点検は行っていない。1号炉の改良範囲外の整備（ブリッジ解除装置・押込送風機ほか）は令和5年度へ送った。'),
                 ('(3) 改良工事との取合い', '炉室ホイストは 10/20～22 に 2号炉火格子片の荷揚げに使用し、改良工事の新火格子搬入は 10/24 以降とした。全停電日（10/23）は両工事とも作業を中止した。')],
        advice=['2号炉の火格子・減温塔・バグフィルタは令和5年度の基幹的設備改良工事で更新されるため、令和5年度の全炉停止整備では更新範囲外の機器（灰押出装置、給じん装置、通風機類、耐火物）を中心とする。',
                '灰押出装置の摺動板は今回取替えたため、3 年後（令和7年度）を目安に残厚を計測し、取替の要否を判断されたい。1号炉の摺動板は令和5年度に計測を推奨する。',
                'ごみクレーン 1号の走行車輪は摩耗が進んでおり（R3 から 1.5mm 増）、令和7〜8年度の車輪交換・減速機整備の計画化を推奨する。',
                '2号炉 誘引通風機の振動は 2.1mm/s で基準内であるが、毎年の測定で傾向を確認されたい。'],
        photos=['2号炉 燃焼段 火格子片 取替前', '2号炉 燃焼段 火格子片 取替後', '2号炉 灰押出装置 摺動板 穴あき部', '2号炉 灰押出装置 摺動板 取替後', '2号炉 側壁 耐火物 補修後', 'ごみクレーン1号 ワイヤロープ 交換'],
        date='令和4年12月9日', author='n1', period='令和4年4月28日 ～ 令和5年3月17日'),
    2023: dict(
        sum='本年度は、基幹的設備改良工事（別契約）で 2号炉が 9 月から 12 月まで停止しているため、2号炉の中間点検は行わず、1号炉中間点検（6月）と全炉停止整備（10月、1号炉と共通設備）を実施した。1号炉中間点検では、令和4年度に改良工事のため送った 1号炉の整備を行った。',
        coord='基幹的設備改良工事（現場代理人 n3）：1号炉側の立入は SV の許可制。炉室ホイストは 10/18～20 本業務優先。10/22 全停電は両工事とも作業中止。',
        special=[('(1) 2号炉中間点検', '2号炉は基幹的設備改良工事のため 9/19 から停止しており、本業務の 2号炉中間点検は行っていない。2号炉の改良範囲外の整備（落じんコンベヤ・空気予熱器・煙突内筒ほか）は令和6年度へ送った。'),
                 ('(2) 1号炉 誘引通風機', '計画では本年度に軸受交換としていたが、振動 1.4mm/s 前後で安定しているため見送った（組合了承）。'),
                 ('(3) 改良後の 1号炉', '1号炉は改良後 1 年目である。火格子・減温塔・バグフィルタとも異常は無く、初期の残厚・差圧を記録した。')],
        advice=['1号炉 灰押出装置の摺動板は端部 2 枚を取替えた。残りは残厚 11mm で、令和7〜8年度に再計測されたい。',
                '令和6年度は、2号炉の改良範囲外の整備（送った分）と、改良後初年度の 2号炉点検を合わせて行う。',
                '空気圧縮機 1号は竣工時の機器で、部品供給の見通しをメーカーに確認されたい。'],
        photos=['1号炉 火格子（改良後1年）', '1号炉 灰押出装置 摺動板 取替', 'ブリッジ解除装置 分解整備', '押込送風機 軸受交換', 'ごみクレーン 年次自主検査', '全停電 立会'],
        date='令和5年12月8日', author='n1', period='令和5年4月27日 ～ 令和6年3月15日'),
    2024: dict(
        sum='本年度から SV が n2 に交代した。基幹的設備改良工事の完了後、初めて 1号炉・2号炉の両方の中間点検と全炉停止整備を行った。令和5年度に送った 2号炉の整備を 2号炉中間点検で行った。全炉停止整備の振動測定で、2号炉 誘引通風機の振動上昇を確認した。',
        coord='新DCS の年次点検はメーカー（瑕疵担保期間中）が別途実施。10/20 全停電（受変電年次点検：別契約）に立会。',
        special=[('(1) 2号炉 誘引通風機 振動上昇', '10/16 の測定で 3.9mm/s（R4：2.1mm/s）、反負荷側軸受に異音があった。判定値 4.5mm/s 以下で運転は継続できるが、軸受の手配に約 3 か月を要するため、令和7年度全炉停止での軸受交換・動バランスを組合へ提案した。'),
                 ('(2) 計量機（出口）', 'ロードセル 1 点の交換（器差）は組合の別発注（計量器メーカー）で 8 月に行われた。')],
        advice=['2号炉 誘引通風機：令和7年度全炉停止で軸受交換・動バランスを行う（軸受は早期手配）。それまで月 1 回の振動測定を運転委託へ依頼されたい。',
                'ごみクレーン 1号：本年度ワイヤロープを交換した。走行車輪・減速機の整備を令和8年度に計画されたい。',
                '1号炉 BF：改良後 3 年目の令和7年度にろ布抜取試験を行う。'],
        photos=['2号IDF 振動測定', '2号炉 落じんコンベヤ 整備', 'ごみクレーン1号 ワイヤロープ交換', '1号炉 火格子 点検', '2号炉 火格子 点検', '全停電 立会'],
        date='令和6年12月6日', author='n2', period='令和6年4月25日 ～ 令和7年3月14日'),
    2025: dict(
        sum='本年度の重点は 2号炉 誘引通風機の軸受交換（令和6年度の振動上昇への対応）である。交換後の単体試運転で振動が判定値を超えたため再バランスを行い、工程が 2 日延びたが、立上げには影響しなかった。',
        coord='10/19 全停電：受変電年次点検（電気保安協会）と区分開閉器（PAS）更新（組合の別発注）。',
        special=[('(1) 2号炉 誘引通風機', '旧軸受（反負荷側）の外輪にはく離があり、令和6年度の振動上昇の原因と判断した。交換後 10/20 の単体試運転で 4.8mm/s（判定 4.5mm/s 以下）となったため 10/22〜23 に再バランスを行い、10/24 に 2.2mm/s で合格した。'),
                 ('(2) 2号炉 灰押出装置 摺動板', '計画では本年度に補修としていたが、残厚 13mm（R4 取替品）で取替不要と判断し、点検のみとした。'),
                 ('(3) ポンプ(A) 整備', '噴射水供給ポンプ(A) の軸スリーブに摩耗があり予備品で対応したため、1 日延びた。')],
        advice=['2号炉 灰押出装置：令和8年度に摺動板を再計測されたい。',
                'ごみクレーン 1号：令和8年度に走行車輪・減速機の整備を行う（部品手配 約 4 か月）。',
                '空気圧縮機 1号：部品供給終了の通知に基づき令和8年度の更新を推奨する。',
                '1号炉 BF：ろ布引張強さは新品比 88% で良好。次回抜取試験は令和10年度。'],
        photos=['2号IDF 旧軸受 外輪はく離', '2号IDF 新軸受 組込', '1号炉 燃焼段 火格子片 取替', '1号BF ろ布 抜取', 'ごみクレーン2号 ワイヤロープ交換', '2号炉 耐火物 補修後'],
        date='令和7年12月5日', author='n2', period='令和7年4月25日 ～ 令和8年3月13日'),
}


def build_new(year):
    T = TEXT[year]
    name = W.TEIKI[year]['name']
    sv = T['author']
    fname = f'{FN}_工事報告書_令和{year-2018}年度定期整備業務委託_{sv}新章立て.pdf'
    stem = rec.start(fname)
    d = G.newdoc(f'【架空】{name} 工事報告書', f'{MAKER} {sv}')
    dc = Doc(stem, d)
    evs = [e for e in ('mid1', 'mid2', 'all') if e in W.TEIKI_EV[year]]
    cover(dc, '工 事 報 告 書', name, T['period'], '（' + '・'.join(W.EVLABEL[e] + ('整備' if e == 'all' else '') for e in evs) + '）', T['date'])
    G.para(d, '目　次', 13, True, 'c', after=8)
    for t in ['第1章　総括', '第2章　工事概要', '第3章　工程実績', '第4章　設備別 施工内容', '第5章　点検・測定記録',
              '第6章　使用材料・支給品', '第7章　追加工事・特記事項', '第8章　今後の整備に関する提言', '付録　工事写真（抜粋）']:
        G.para(d, t, 11, indent=30, after=5)
    if year == 2022:
        G.para(d, '※ 本年度より、基幹的設備改良工事の報告書と章立てをそろえた。', 9, color='595959', before=12, indent=30)
    d.add_page_break()
    dc.head('第1章　総括')
    dc.p(T['sum'], after=6)
    # サマリー：整備・補修・更新と、要注意の測定
    rows, ws = [], []
    no = 0
    import plan_data as PD
    def score(w):
        m = MBY[w['machines'][0]]
        sc = 0
        if w['measure'] and w['measure'][3] not in ('良',):
            sc += 10
        if w['wtype'] != 'INSPECT':
            sc += 3
        sc += {'A': 4, 'B': 2}.get(PD.importance(m), 0)
        if w['rev']:
            sc += 8
        if m['cls'] in ('crane', 'idf', 'grate'):
            sc += 2
        return -sc
    cand = [w for w in works_of(year) if w['machines'] and (w['wtype'] != 'INSPECT' or w['measure'])]
    cand = sorted(sorted(cand, key=lambda w: w['id'])[:200], key=score)[:12]
    cand.sort(key=lambda w: w['id'])
    for w in cand:
        m = MBY[w['machines'][0]]
        no += 1
        tags = '・'.join(w['machines'][:3])
        res = w['measure'][2] if w['measure'] else ('完了' if not w['finding'] else w['finding'][:24])
        nxt = '－'
        if w['measure'] and w['measure'][3] in ('要注意', '不良（当て板）'):
            nxt = '傾向監視・次年度計画'
        rows.append((no, f"{W.short(m)}（{tags}）", w['name'].replace(W.short(m), '').strip() or w['name'], res, nxt))
        ws.append([w['id']])
    dc.table(['No', '対象機器（機器番号）', '実施内容', '結果', '次回への申し送り'], rows, [8, 46, 44, 38, 30], 8.5, works=ws)
    dc.head('第2章　工事概要')
    rows2 = [('工事件名', name), ('施設', f'{FAC}（30t/16h×2炉）'), ('工期', T['period'].replace(' ', ''))]
    for e in ('mid1', 'mid2', 'all'):
        if e in W.TEIKI_EV[year]:
            rows2.append((W.EVLABEL[e] + ('整備' if e == 'all' else ''), ev_period(year, e)))
        else:
            rows2.append((W.EVLABEL[e], '実施せず（基幹的設備改良工事の範囲で開放・更新するため。仕様書 第3条ただし書）'))
    teams = sorted({w['team'] for w in works_of(year)} - {'全班', '－'})
    rows2.append(('体制', f'SV {sv}、' + '、'.join(teams)))
    rows2.append(('他工事との調整', T['coord']))
    dc.table(['項目', '内容'], rows2, [36, 130])
    dc.head('第3章　工程実績')
    rows, ws = [], []
    for e in evs:
        for w in works_of(year, e):
            diff = '追加' if w['rev'] else ('延長' if w['plan'] != w['act'] else '')
            rows.append((W.EVLABEL[e], w['kubun'], w['name'].replace('【追加】', ''), ('（改1）' if w['rev'] else '') + fmt(w['plan']), fmt(w['act']), diff))
            ws.append([w['id']])
    dc.table(['停止', '区分', '作業', '計画', '実績', '差'], rows, [26, 12, 64, 26, 26, 12], 8, works=ws)
    dc.head('第4章　設備別 施工内容')
    groups = OrderedDict()
    for w in works_of(year):
        sy = sys_of(w)
        if sy is None:
            continue
        groups.setdefault(sy, []).append(w)
    for i, (sy, lst) in enumerate(groups.items(), 1):
        dc.head(f'4.{i}　{SYSNAME[sy]}', 2)
        for w in lst:
            when = fmt(w['act'])
            t = f"{w['name']}（{W.EVLABEL.get(w['event'], '')} {when}）" + (f"：{w['finding']}。" if w['finding'] else '。')
            dc.p(t, works=[w['id']], indent=4, after=3)
    dc.head('第5章　点検・測定記録')
    rows, ws = measures(year)
    dc.table(['機器番号', '機器', '項目', '基準・前回', '結果', '判定'], rows, [14, 34, 30, 30, 32, 26], 8.5, works=ws)
    dc.head('第6章　使用材料・支給品')
    rows, ws = parts(year)
    dc.table(['品名', '仕様', '数量', '区分', '使用箇所', '備考'], rows, [32, 40, 14, 26, 36, 18], 8.5, works=ws)
    dc.head('第7章　追加工事・特記事項')
    for h, t in T['special']:
        G.para(d, h, 10.5, True)
        dc.p(t, indent=4, after=6)
    skipped = [x for x in W.TEIKI_SKIP[year]]
    if skipped:
        G.para(d, f'({len(T["special"]) + 1}) 計画のうち本年度に行わなかったもの', 10.5, True)
        dc.table(['機器番号', '機器', '計画', '理由'], [(t, W.short(MBY[t]), W.sym_of(s), why) for t, s, why in skipped], [18, 50, 14, 84], 8.5)
    dc.head('第8章　今後の整備に関する提言')
    for t in T['advice']:
        dc.p('・' + t, indent=4, after=5)
    d.add_page_break()
    G.head(d, '付録　工事写真（抜粋）')
    photo_table(dc, T['photos'])
    p = os.path.join(WORK, fname.replace('.pdf', '.docx'))
    d.save(p)
    return p, fname


def photo_table(dc, caps):
    d = dc.d
    t = d.add_table(rows=6, cols=2); t.style = 'Table Grid'; t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, cap in enumerate(caps[:6]):
        rr, cc = (i // 2) * 2, i % 2
        c = t.cell(rr, cc); c.text = ''
        G.shade(c, 'D9D9D9')
        p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        G.setfont(p.add_run('\n\n\n（写真）\n\n\n'), 10, color='808080')
        c2 = t.cell(rr + 1, cc); c2.text = ''
        G.setfont(c2.paragraphs[0].add_run(f'写真{i+1}　{cap}'), 9)
    t.autofit = False; t.columns[0].width = Mm(82); t.columns[1].width = Mm(82)
    for i, cap in enumerate(caps[:6]):
        rec.add(dc.stem, None, None, 'item', dict(章='付録　工事写真（抜粋）', 写真=f'写真{i+1}', 説明=cap, image='（写真の枠のみ。画像なし）'), seq=rec.next_seq(dc.stem))


# ---------------------------------------------------------------- 基幹改良 完成報告書（n3）
def build_kikan():
    name = W.KIKAN['name']
    fname = f'{FN}_工事報告書_基幹的設備改良工事_完成報告書_n3.pdf'
    stem = rec.start(fname)
    d = G.newdoc(f'【架空】{name} 完成報告書', f'{MAKER} 現場代理人 n3')
    dc = Doc(stem, d)
    cover(dc, '完 成 報 告 書', name, '令和4年6月27日 ～ 令和6年3月15日', '（1号炉・2号炉 火格子・減温塔・バグフィルタ更新、DCS更新）', '令和6年3月15日')
    G.para(d, '目　次', 13, True, 'c', after=8)
    for t in ['第1章　総括', '第2章　工事概要', '第3章　工程実績', '第4章　設備別 施工内容', '第5章　試運転・性能確認',
              '第6章　主要機器・材料', '第7章　変更・追加工事', '第8章　保守に関する申し送り', '付録　工事写真（抜粋）']:
        G.para(d, t, 11, indent=30, after=5)
    d.add_page_break()
    dc.head('第1章　総括')
    dc.p('本工事は、竣工後 21〜23 年を経過した第二工場の主要設備（火格子、水噴射式減温塔、バグフィルタ）を炉ごとに更新し、分散制御システム（DCS）を更新したものである。'
         '令和4年度に 1号炉、令和5年度に 2号炉を施工し、各年度の定期整備業務委託（別契約）の全炉停止と停止期間を分け合った。DCS は令和6年2月の全炉停止で切替えた。更新した機器は下表のとおりである。')
    rows, ws = [], []
    for k in (1, 2):
        for m in MACHINES:
            if m['furnace'] == k and m.get('kikan'):
                w = [x['id'] for x in W.WORKS if x['contract'] == 'KIKAN' and m['tag'] in x['machines']]
                rows.append((m['tag'], m['name'], f'{k}号炉', m['spec'], m['inst'].replace('-', '.')))
                ws.append(w)
    for m in MACHINES:
        if m['furnace'] == 0 and m.get('kikan'):
            w = [x['id'] for x in W.WORKS if x['contract'] == 'KIKAN' and m['tag'] in x['machines']]
            rows.append((m['tag'], m['name'], '共通', m['spec'].replace('【R5新設】', '（新設）'), m['inst'].replace('-', '.')))
            ws.append(w)
    dc.table(['機器番号', '機器', '号炉', '新仕様', '引渡'], rows, [16, 40, 14, 74, 20], 8, works=ws)
    dc.head('第2章　工事概要')
    dc.table(['項目', '内容'], [('工事件名', name), ('施設', f'{FAC}（30t/16h×2炉）'), ('工期', '令和4年6月27日～令和6年3月15日（継続費 2 か年）'),
                              ('令和4年度 現地工事', '令和4年9月12日～12月9日（1号炉、共通DCS盤 据付）'), ('令和5年度 現地工事', '令和5年9月11日～12月8日（2号炉）'),
                              ('DCS 切替', '令和6年2月2日～2月22日（全炉停止）'), ('現場代理人', 'n3'),
                              ('協力会社', '、'.join(v.replace('【架空】', '') for k, v in W.SUBCON.items() if k not in ('全班', '－')))], [40, 126])
    dc.head('第3章　工程実績')
    rows, ws = [], []
    for key, lst in (('R4', S.KIKAN_R4), ('R5', W.KIKAN_R5)):
        for (wbs, nm, team, n, spans, kind) in lst:
            if kind == 'H' or not spans:
                continue
            wid = next((w['id'] for w in W.KIKAN_EV[key] if w.get('wbs') == wbs), None)
            rows.append(('令和4年度' if key == 'R4' else ('DCS切替' if wbs.startswith('6') else '令和5年度'), wbs, nm.replace('【追加】', '（追加）'), team or '－', fmt(spans)))
            ws.append([wid])
    dc.table(['年度', 'WBS', '作業', '班', '実績'], rows, [20, 12, 76, 22, 36], 8, works=ws)
    dc.head('第4章　設備別 施工内容')
    secs = [
        ('4.1　火格子', ['各炉とも既設の火格子（乾燥段・燃焼段・後燃焼段）、駆動フレーム、油圧配管を撤去し、新しい並行揺動式火格子（ASR-30）を据付けた。炉油圧ユニットも更新した。据付後の芯出し・ストローク調整を行い、組合立会の据付検査に合格した（1号炉 R4.11.17、2号炉 R5.11.16）。',
                       '令和4年度は炉室ホイストを定期整備業務委託と共用したため、新火格子の搬入を 10/24〜26 に変更した。']),
        ('4.2　減温塔', ['下部胴・ホッパを更新し、噴霧方式を一流体から二流体ノズルへ改めた。噴射水ポンプ(A)(B) と減温塔噴霧制御盤を更新した。溶接部は非破壊検査を行った。']),
        ('4.3　バグフィルタ', ['本体ケーシング・ホッパ、逆洗装置（電磁弁 48 個）、加温装置を更新し、ろ布 192 本を装着した。予備ろ布 20 本を引渡した（組合保管）。旧ろ布の予備品（組合保管 14 本）は同寸法のため新設 BF に使用できる。']),
        ('4.4　計装・制御', ['新DCS（オペレータステーション 3、共通・炉コントローラ、エンジニアリングステーション、帳票サーバ、ネットワーク盤、大型表示装置）を据付け、令和6年2月の全炉停止で旧DCS から切替えた。旧 CRT・グラフィックパネルは撤去した。減温塔出口温度計・BF 入口温度計・BF 差圧計を取替えた。']),
    ]
    kw = {'4.1': ['1.1', '1.2', '1.3', '1.4', '1.5'], '4.2': ['2.1', '2.2', '2.3', '2.4', '2.5'], '4.3': ['3.1', '3.2', '3.3', '3.4'], '4.4': ['4.1', '4.2', '4.3', '4.4', '6.2', '6.3']}
    for h, ts in secs:
        G.head(d, h, 2)
        wids = [w['id'] for w in W.WORKS if w['contract'] == 'KIKAN' and w.get('wbs') in kw[h[:3]]]
        for t in ts:
            dc.p(t, works=wids, indent=4, after=5)
    dc.head('第5章　試運転・性能確認')
    dc.table(['項目', '保証値・基準', '1号炉（R4.12）', '2号炉（R5.12）'], [
        ('処理能力', '30t/16h', '30.8t/16h', '31.2t/16h'), ('減温塔出口温度', '165℃±10℃', '163℃', '166℃'), ('BF 差圧', '1.5kPa 以下', '1.1kPa', '1.0kPa'),
        ('煙突 ばいじん', '0.01g/m³N 以下', '0.002g/m³N', '0.001g/m³N未満'), ('煙突 ダイオキシン類', '0.1ng-TEQ/m³N 以下', '0.011ng-TEQ/m³N', '0.008ng-TEQ/m³N'),
        ('誘引通風機 振動（参考）', '4.5mm/s 以下', '1.6mm/s', '2.0mm/s')], [44, 44, 40, 40], 8.5)
    dc.head('第6章　主要機器・材料')
    dc.table(['品名', '仕様', '数量', '区分', '使用箇所'], [
        ('火格子片', 'ASR-30 乾燥・燃焼・後燃焼段', '各炉 1式', '本工事', '1号炉・2号炉'), ('炉油圧ユニット', '油圧ポンプ2台内蔵 400L', '2基', '本工事', '1号炉・2号炉'),
        ('減温塔 下部胴', 'SS400 t6（内面耐酸塗装）', '2基', '本工事', '1号炉・2号炉'), ('二流体噴霧ノズル', '噴霧量 0.6m³/h', '各炉 6本＋予備 2本', '本工事', '減温塔'),
        ('ろ布', 'φ150×5,000 PTFE', '各炉 192本＋予備 20本', '本工事', 'バグフィルタ'), ('キャスタブル', '高アルミナ質', '各炉 4,200kg', '本工事', '炉床まわり'),
        ('当て板（追加）', 'SS400 t6', '3枚', '本工事（変更）', '2号 煙道')], [30, 52, 30, 26, 28], 8.5)
    dc.head('第7章　変更・追加工事')
    w26 = [w['id'] for w in W.WORKS if w['contract'] == 'KIKAN' and w.get('wbs') == '2.6']
    dc.p('(1) 2号煙道 減肉部補修（追加）：R5.10.31、減温塔入口側の煙道に減肉部（最小 2.9mm／原厚 4.5mm）を確認し、監督員の指示で当て板補修 3 か所を追加施工した（第6回変更 R6.1.12）。', works=w26)
    dc.p('(2) 令和4年度の新火格子搬入日の変更：定期整備業務委託との炉室ホイスト共用のため 10/19〜21 → 10/24〜26 に変更した（工程表 第2版）。')
    dc.head('第8章　保守に関する申し送り')
    for t in ['火格子：燃焼段は 3 年、乾燥段・後燃焼段は 4 年を目安に残厚を計測し、部分取替を計画されたい。',
              'バグフィルタ：ろ布は 3 年ごとに抜取試験を行い、8 年を目安に全数交換を計画されたい。',
              '減温塔：二流体ノズルは毎回の停止で点検し、先端摩耗のものは予備品と交換されたい。',
              'DCS：瑕疵担保期間（2 年）中は年 1 回の無償点検を行う。HMI・PC は 8 年を目安に更新を計画されたい。']:
        dc.p('・' + t, indent=4, after=5)
    d.add_page_break()
    G.head(d, '付録　工事写真（抜粋）')
    photo_table(dc, ['1号炉 既設火格子 撤去', '1号炉 新火格子 据付', '1号減温塔 新下部胴 吊込み', '1号BF ろ布 装着', '2号煙道 減肉部 当て板', '新DCS オペレータステーション'])
    p = os.path.join(WORK, fname.replace('.pdf', '.docx'))
    d.save(p)
    return p, fname


def to_pdf(docx_path, pdf_name):
    pdf = render.to_pdf(docx_path, WORK)
    dst = os.path.join(OUT, pdf_name)
    os.replace(pdf, dst)
    return dst


def build_all():
    out = []
    for p, fn in [build_r3()] + [build_new(y) for y in (2022, 2023, 2024, 2025)] + [build_kikan()]:
        to_pdf(p, fn)
        out.append(fn)
    return out


if __name__ == '__main__':
    print(build_all())
