# -*- coding: utf-8 -*-
"""工事報告書（PDF）。docx を作って LibreOffice で文字のある PDF にする。表は頁ごとに見出し行をくり返す。"""
import os, glob, datetime as dt, random, subprocess, shutil
from docx import Document
from docx.shared import Pt, Mm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import ans, world2 as W
from xlutil import TMP, normalize_zip
from f01_master import SYS, SYS_ORDER, wareki, EQ
from f01_history import BYTAG, PLAN26
from f01_contracts import MAKER
from make_report import RES, result, judge, setfont, shade
import build_sd as B

DOCX = TMP + '_docx'
KAKU = '【架空】'
PENDING = []   # (docx, pdfname)

class Rep:
    def __init__(self, name, font, plan=False):
        self.name = name
        self.stem = ans.start(name + '.pdf')
        self.doc = Document()
        self.font = font
        self.seq = 0
        s = self.doc.sections[0]
        s.page_width, s.page_height = Mm(210), Mm(297)
        s.left_margin = s.right_margin = Mm(18); s.top_margin = Mm(18); s.bottom_margin = Mm(16)
    def para(self, text, size=10.5, bold=False, align=None, before=0, after=2, kind='item', fields=None, works=None, record=True, color=None):
        p = self.doc.add_paragraph()
        r = p.add_run(text); setfont(r, self.font, size, bold, color)
        if align == 'c': p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if align == 'r': p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.paragraph_format.space_before = Pt(before); p.paragraph_format.space_after = Pt(after)
        if record and text.strip():
            ans.rec(self.stem, kind, fields or {'text': text}, seq=self.seq, works=works)
            self.seq += 1
        return p
    def table(self, rows, widths, size=8.5, hdr='D9D9D9', keys=None, works=None, kinds=None):
        t = self.doc.add_table(rows=len(rows), cols=len(rows[0]))
        t.style = 'Table Grid'; t.alignment = WD_TABLE_ALIGNMENT.CENTER
        t.autofit = False
        for i, row in enumerate(rows):
            for j, v in enumerate(row):
                c = t.cell(i, j); c.text = ''
                p = c.paragraphs[0]; p.paragraph_format.space_after = Pt(0)
                setfont(p.add_run('' if v is None else str(v)), self.font, size, bold=(i == 0))
                c.width = Mm(widths[j])
                if i == 0:
                    shade(c, hdr); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if i == 0:   # 見出し行を頁ごとにくり返す
                trPr = t.rows[0]._tr.get_or_add_trPr()
                e = OxmlElement('w:tblHeader'); e.set(qn('w:val'), 'true'); trPr.append(e)
            keys_ = keys or [str(x) for x in rows[0]]
            f = {k: (row[j] if row[j] not in ('', None) else None) for j, k in enumerate(keys_)}
            kind = 'parent' if i == 0 else ((kinds or {}).get(i, 'item'))
            ans.rec(self.stem, kind, f, seq=self.seq, works=(works or {}).get(i))
            self.seq += 1
        for j, w in enumerate(widths):
            t.columns[j].width = Mm(w)
        return t
    def kv(self, items, w=(32, 140), size=9.5):
        rows = [['項目', '内容']] + [list(x) for x in items]
        return self.table(rows, w, size, hdr='F2F2F2', keys=['項目', '内容'])
    def save(self, author):
        cp = self.doc.core_properties
        cp.title = KAKU + self.name.replace(KAKU, ''); cp.author = KAKU + author; cp.last_modified_by = cp.author
        cp.comments = KAKU + '試験用サンプル（あさぎ市清掃工場 F01）。実在の施設・会社・人とは関係ありません。'
        cp.created = dt.datetime(2026, 11, 13, 9, 0, 0); cp.modified = cp.created; cp.revision = 1
        os.makedirs(DOCX, exist_ok=True)
        p = DOCX + '/' + self.name + '.docx'
        self.doc.save(p)
        normalize_zip(p)
        PENDING.append(p)

def jd(x):
    return ('平成%d年%d月%d日' % (x.year - 1988, x.month, x.day)) if (x.year, x.month) < (2019, 5) else ('令和%d年%d月%d日' % (x.year - 2018, x.month, x.day))

def evs_T(fy, cts=('T',)):
    out = []
    for (d, k, t, c, ct, i) in W.events_for(fy, None, cts):
        if fy == 2026:
            continue
        out.append((d, k, t, c, ct, i))
    return out

def meas_text(fy, d, k, t):
    if (fy, d['tag']) in RES:
        return RES[(fy, d['tag'])]
    v = W.h('m', fy, d['tag']) % 1000
    if k == '計測':
        return '最小 %.1fmm（管理値内）。' % (2.8 + v / 600.0) if d['prof'] in ('tube', 'sh') else '最大摩耗 %.1fmm（限度内）。' % (1.5 + v / 250.0)
    if k == '点検' and d['prof'].startswith('fan'):
        return '振動 %.1fmm/s、異常なし。' % (0.9 + v / 600.0)
    return result(fy, d, k, t)

SUPPLIED = {2007: 'BF-101', 2013: 'BF-101', 2019: 'BF-101', 2025: 'BF-101', 2008: 'BF-201', 2014: 'BF-201', 2020: 'BF-201'}

def teishu_report(fy):
    st = W.report_style(fy)
    P = W.teishu(fy)
    wins = P['wins']
    name = '%sあさぎ市清掃工場_工事報告書_%s定期整備業務委託_%s様式' % (KAKU, wareki(fy, False), st)
    R_ = Rep(name, 'ＭＳ 明朝' if st == 'r1' else 'ＭＳ Ｐゴシック')
    evs = evs_T(fy)
    code = lambda d: d['code'] if fy >= 2019 else d['tag']
    writer = W.report_writer(fy)
    title = '%s あさぎ市清掃工場 定期整備業務委託' % wareki(fy, False)
    if st == 'r1':
        for _ in range(4): R_.para('', record=False)
        R_.para(wareki(fy, False), 16, align='c')
        R_.para(KAKU + 'あさぎ市清掃工場', 20, True, align='c', after=6)
        R_.para('定期整備業務委託', 20, True, align='c', after=24)
        R_.para('工　事　報　告　書', 24, True, align='c', after=60)
        R_.para(jd(wins[2][1] + dt.timedelta(days=30)).replace('日', '').rsplit('月', 1)[0] + '月', 14, align='c', after=30)
        R_.para('汐見重工環境株式会社　環境プラント事業部', 13, align='c', after=10)
        R_.para('作成 %s' % writer, 10, align='r')
        R_.doc.add_page_break()
        R_.para('１．工事概要', 12, True, kind='parent')
    else:
        hp = R_.doc.sections[0].header.paragraphs[0]
        setfont(hp.add_run(KAKU + 'あさぎ市清掃工場　%s 定期整備業務委託　工事報告書' % wareki(fy, False)), R_.font, 8, color='595959')
        R_.para(KAKU + '工事報告書（サマリー）', 16, True, after=4, kind='parent')
    op = W.operator_word(fy)
    R_.kv([('件名', title), ('契約工期', '%s〜%s' % (jd(W.teishu_period(fy)[1]), jd(W.teishu_period(fy)[2]))),
           ('施工期間', '1号炉 %s〜%s／全停 %s〜%s／2号炉 %s〜%s' % (jd(wins[1][0]), jd(wins[1][1]), jd(wins['Z'][0]), jd(wins['Z'][1]), jd(wins[2][0]), jd(wins[2][1]))),
           ('受注者', '汐見重工環境株式会社 環境プラント事業部（現場代理人 %s）' % W.agent(fy)), ('炉の運転操作', op)])
    n = {f: len([e for e in evs if e[0]['f'] == f]) for f in (1, 2, 0)}
    R_.para('', record=False)
    R_.para(('２．実施概要' if st == 'r1' else '1. 実施概要'), 11, True, kind='parent')
    hs = W.hoshu_contracts(fy)
    R_.para('定期整備仕様書に基づき、1号炉・2号炉の順に炉を停止して点検・整備を行い、%s期間中に全停を行った。対象機器は1号炉 %d 件、2号炉 %d 件、共通 %d 件の計 %d 件である。%s'
            % ('1号炉' if (fy, 1) not in W.SPECIAL_LONG else '2号炉', n[1], n[2], n[0], len(evs),
               ('別契約の補修工事（%s）は本報告の対象外である。' % '、'.join(c['title'].split('清掃工場')[-1] for c in hs)) if hs else ''), 9.5)
    if fy == 1997:
        R_.para('竣工後 1 年目の定期整備であり、瑕疵担保期間中のため、各機器の点検と消耗品の確認を主とした。交換部品は保証で対応した。', 9.5)
    # 対象機器と実施内容
    R_.para(('３．整備対象機器と実施内容' if st == 'r1' else '2. 対象機器と実施内容'), 11, True, kind='parent', before=4)
    order = {1: 0, 0: 1, 2: 2}
    evs.sort(key=lambda e: (order[e[0]['f']], e[0]['code']))
    rows = [['機器コード' if fy >= 2019 else '機器番号', '号炉', '機器名', '区分', '実施内容', '結果' if st == 'r1' else '判定']]
    wk = {}
    for (d, k, t, c, ct, i) in evs:
        rows.append([code(d), d['fname'], d['name'], k, t, meas_text(fy, d, k, t) if st == 'r1' else judge(fy, d, k)])
        wk[len(rows) - 1] = [B.wid(fy, d['tag'], i)]
    if len(rows) > 1:
        R_.table(rows, (17, 12, 40, 10, 50, 49) if st == 'r1' else (17, 12, 44, 10, 80, 11), 7.5,
                 keys=['code', 'furnace', 'name', 'kind', 'work', 'result'], works=wk)
    else:
        R_.para('（整備対象機器なし。各設備の外観点検・運転状態の確認のみ）', 9)
    # 主な確認結果
    R_.para(('４．主な確認結果' if st == 'r1' else '3. 主な確認結果'), 11, True, kind='parent', before=6)
    keyev = [e for e in evs if e[1] in ('計測', '補修', '更新') or (fy, e[0]['tag']) in RES][:10]
    for (d, k, t, c, ct, i) in keyev:
        R_.para('■ %s %s（%s）：%s' % (d['fname'], d['name'], code(d), meas_text(fy, d, k, t)), 9.5, works=[B.wid(fy, d['tag'], i)])
    # 支給品
    R_.para(('５．使用部品・支給品' if st == 'r1' else '4. 支給品・部品の扱い'), 11, True, kind='parent', before=6)
    sup = SUPPLIED.get(fy)
    rows = [['品名', '数量', '使用箇所', '区分']]
    if sup:
        f = BYTAG[sup]['f']
        rows.append(['ろ布（φ165×6,000）', '300本受入・288本使用', '%d号炉 バグフィルタ' % f, '市支給品。残 12 本は予備として市へ返納'])
    for (d, k, t, c, ct, i) in [e for e in evs if e[1] in ('整備', '補修', '更新')][:8]:
        rows.append([B.DY.PARTS[B.DY.part_group(d)][W.h('pt', fy, d['tag']) % 3], '1式', '%s %s' % (d['fname'], d['name']), '受注者手配'])
    R_.table(rows, (45, 30, 55, 48), 8, keys=['part', 'qty', 'where', 'kind'])
    # 所見・提案（次年度の補修工事・計画の主な作業）
    R_.para(('６．所見および今後の整備について' if st == 'r1' else '5. 今後の整備に関する提案'), 11, True, kind='parent', before=6)
    nxt = W.hoshu_contracts(fy + 1) if fy < 2026 else []
    if nxt:
        for c in nxt[:3]:
            d = BYTAG[c['tag']]
            R_.para('・%s %s：劣化が進んでおり、%sの補修（%s）を提案する。' % (d['fname'], d['name'], wareki(fy + 1, False), c['title'].split('清掃工場')[-1]), 9.5)
    else:
        R_.para('・特記事項なし。次年度も定期整備の範囲で点検・整備を継続する。', 9.5)
    R_.save('汐見重工環境 環境プラント事業部 ' + writer)

def contract_report(fy, c, L, kind):
    vend = c['vendor']
    vst = vend != MAKER
    st = 'v' if vst else W.report_style(fy)
    short = c['title'].split('清掃工場')[-1].replace(' ', '')
    name = '%sあさぎ市清掃工場_工事報告書_%s%s_%s' % (KAKU, wareki(fy, False), short, '業者様式' if vst else st + '様式')
    R_ = Rep(name, 'ＭＳ 明朝' if st in ('r1', 'v') else 'ＭＳ Ｐゴシック')
    if vst:
        R_.para(KAKU + '工事完成報告書', 20, True, align='c', after=20, kind='parent')
        R_.para('あさぎ市長　様', 11, after=10)
        R_.para(vend, 11, align='r', after=16)
    else:
        R_.para(KAKU + ('工　事　報　告　書' if st == 'r1' else '工事報告書'), 18, True, align='c', after=12, kind='parent')
    evs = [(d, k, t, cc, ct, i) for (d, k, t, cc, ct, i) in W.events_for(fy, None, (kind,)) if kind != 'H' or d['tag'] == c['tag']]
    s, e = min(t['s'] for t in L), max(t['e'] for t in L)
    R_.kv([('工事件名', c['title']), ('工事場所', 'あさぎ市清掃工場'), ('契約工期', '%s〜%s' % (jd(c['s']), jd(c['e_']))), ('現場施工', '%s〜%s' % (jd(s), jd(e))),
           ('請負者', vend)])
    R_.para('', record=False)
    R_.para('1. 施工内容' if not vst else '一、施工内容', 11, True, kind='parent')
    for (d, k, t, cc, ct, i) in evs[:40]:
        R_.para('%s %s：%s。%s' % (d['fname'], d['name'], t.replace('（基幹改良）', ''), meas_text(fy, d, k, t) if kind == 'H' else '据付・試運転を完了した。'),
                9.5, works=[B.wid(fy, d['tag'], i)])
    if len(evs) > 40:
        R_.para('ほか %d 件（付表のとおり）。' % (len(evs) - 40), 9.5)
    R_.para('2. 工程' if not vst else '二、工程', 11, True, kind='parent', before=4)
    rows = [['No', '作業', '開始', '終了']] + [[t['no'], t['item'], t['s'].isoformat(), t['e'].isoformat()] for t in L]
    R_.table(rows, (10, 110, 25, 25), 8, keys=['no', 'item', 'start', 'end'])
    R_.para('3. 検査' if not vst else '三、検査', 11, True, kind='parent', before=4)
    R_.para('出来形・寸法検査を市（%s）の立会で行い、合格した。' % W.city_person(fy), 9.5)
    R_.save(vend + ('' if vst else ' ' + W.report_writer(fy)))

def to_pdfs():
    from photo import to_pdf
    out = TMP + '_pdf'
    os.makedirs(out, exist_ok=True)
    for f in glob.glob(out + '/*'):
        os.remove(f)
    to_pdf(PENDING, out)
    for p in PENDING:
        b = os.path.splitext(os.path.basename(p))[0] + '.pdf'
        shutil.copy(os.path.join(out, b), os.path.join(ans.FILES, b))
    PENDING.clear()
