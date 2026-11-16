# -*- coding: utf-8 -*-
"""工事報告書（メーカー n1、Word で書いて PDF で提出）。
R3：旧章立て（工事概要→作業内容→点検結果→所見→使用材料。サマリーなし）
R4：基幹改良（n3）の報告書に合わせた新章立て（第1章 総括のサマリー表から始まる）
"""
import os, datetime as dt
from docx import Document
from docx.shared import Pt, Mm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from common import *
from schedules import *
import render

JF = 'IPAGothic'


def setfont(run, size=10.5, bold=False, color=None):
    run.font.name = JF
    run._element.rPr.rFonts.set(qn('w:eastAsia'), JF)
    run.font.size = Pt(size)
    run.bold = bold
    if color:
        run.font.color.rgb = RGBColor.from_string(color)


def newdoc(title, author):
    d = Document()
    sec = d.sections[0]
    sec.page_width, sec.page_height = Mm(210), Mm(297)
    sec.left_margin = sec.right_margin = Mm(22)
    sec.top_margin = Mm(20); sec.bottom_margin = Mm(18)
    st = d.styles['Normal']
    st.font.name = JF; st.element.rPr.rFonts.set(qn('w:eastAsia'), JF); st.font.size = Pt(10.5)
    d.core_properties.title = title
    d.core_properties.author = author
    d.core_properties.subject = '【架空】試験用サンプル'
    d.core_properties.keywords = '【架空】'
    d.core_properties.created = dt.datetime(2026, 11, 13, 9, 0, 0)
    # ページ番号
    p = sec.footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(); setfont(r, 9)
    for t, txt in (('begin', None), (None, 'PAGE'), ('end', None)):
        if t:
            e = OxmlElement('w:fldChar'); e.set(qn('w:fldCharType'), t); r._element.append(e)
        else:
            e = OxmlElement('w:instrText'); e.set(qn('xml:space'), 'preserve'); e.text = txt; r._element.append(e)
    return d


def para(d, text, size=10.5, bold=False, align=None, before=0, after=4, color=None, indent=None):
    p = d.add_paragraph()
    if align == 'c':
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == 'r':
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.space_before = Pt(before); p.paragraph_format.space_after = Pt(after)
    if indent:
        p.paragraph_format.left_indent = Mm(indent)
    r = p.add_run(text); setfont(r, size, bold, color)
    return p


def head(d, text, lvl=1):
    return para(d, text, 13 if lvl == 1 else 11, True, before=10 if lvl == 1 else 6, after=4)


def shade(cell, hexc):
    tcPr = cell._element.get_or_add_tcPr()
    s = OxmlElement('w:shd'); s.set(qn('w:val'), 'clear'); s.set(qn('w:color'), 'auto'); s.set(qn('w:fill'), hexc)
    tcPr.append(s)


def table(d, hdr, rows, widths, size=9, hdr_fill='D9D9D9'):
    t = d.add_table(rows=1, cols=len(hdr))
    t.style = 'Table Grid'
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    # 見出し行をページごとにくり返す
    trPr = t.rows[0]._tr.get_or_add_trPr()
    e = OxmlElement('w:tblHeader'); e.set(qn('w:val'), 'true'); trPr.append(e)
    for i, h in enumerate(hdr):
        c = t.rows[0].cells[i]
        c.text = ''
        r = c.paragraphs[0].add_run(h); setfont(r, size, True)
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        shade(c, hdr_fill)
    for row in rows:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = ''
            r = cells[i].paragraphs[0].add_run('' if v is None else str(v)); setfont(r, size)
    t.autofit = False
    for i, w in enumerate(widths):
        t.columns[i].width = Mm(w)
    for row in t.rows:
        # 行が途中で割れないように
        trPr = row._tr.get_or_add_trPr()
        e = OxmlElement('w:cantSplit'); trPr.append(e)
        for i, w in enumerate(widths):
            row.cells[i].width = Mm(w)
    d.add_paragraph().paragraph_format.space_after = Pt(2)
    return t


def cover(d, title, name, period, sub, date_txt):
    para(d, '【架空】試験用サンプル', 8, align='r', color='808080')
    for _ in range(5):
        para(d, '')
    para(d, title, 22, True, 'c', after=18)
    para(d, name, 13, True, 'c', after=6)
    para(d, sub, 11, align='c', after=40)
    t = d.add_table(rows=4, cols=2); t.style = 'Table Grid'; t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, (a, b) in enumerate((('施 設 名', FAC), ('工　　期', period), ('発 注 者', f'{UNION}　管理者'), ('受 注 者', MAKER))):
        t.cell(i, 0).text = ''; t.cell(i, 1).text = ''
        setfont(t.cell(i, 0).paragraphs[0].add_run(a), 11, True)
        setfont(t.cell(i, 1).paragraphs[0].add_run(b), 11)
        t.cell(i, 0).width = Mm(35); t.cell(i, 1).width = Mm(105)
    t.autofit = False; t.columns[0].width = Mm(35); t.columns[1].width = Mm(105)
    for _ in range(8):
        para(d, '')
    para(d, date_txt, 12, align='c')
    para(d, MAKER, 12, True, 'c')
    d.add_page_break()


# ======================================================================= R3（旧章立て）
def build_r3(path_docx):
    name = '令和3年度 やまぶき広域処理組合第二工場 定期整備業務委託'
    d = newdoc(f'【架空】{name} 工事報告書', f'{MAKER} n1')
    cover(d, '工 事 報 告 書', name, '令和3年4月26日 ～ 令和4年3月18日', '（1号炉中間点検・2号炉中間点検・全炉停止整備）', '令和3年12月10日')
    head(d, '1．工事概要')
    table(d, ['項目', '内容'], [
        ('工事件名', name), ('施設', f'{FAC}（30t/16h×2炉、並行揺動式ストーカ、准連続）'),
        ('工期', '令和3年4月26日～令和4年3月18日'),
        ('1号炉中間点検', '令和3年6月14日～6月18日（停止 6/11～6/19）'),
        ('2号炉中間点検', '令和3年7月12日～7月16日（停止 7/9～7/17）'),
        ('全炉停止整備', '令和3年10月18日～10月29日（停止 10/15～10/31）'),
        ('現場責任者', 'SV n1'),
    ], [40, 126])
    head(d, '2．作業内容')
    head(d, '2.1　1号炉中間点検', 2)
    for t in ['炉内清掃、火格子の点検・計測。燃焼段に脱落した火格子片 3 片があり予備品と取替えた。',
              '減温塔の噴霧ノズル清掃（4本）、下部胴の肉厚測定（超音波厚さ計、24点）。',
              'バグフィルタ ろ布点検。破損 1 本を組合予備品と交換した。',
              '誘引通風機 振動測定、軸受給脂。']:
        para(d, '・' + t, indent=4)
    head(d, '2.2　2号炉中間点検', 2)
    for t in ['炉内清掃、火格子の点検・計測。燃焼段 割れ片 8 片を取替えた。',
              'バグフィルタ 差圧の上昇（1.8kPa）が続くため、ろ布を全数目視点検し破損 2 本を交換した。逆洗電磁弁 2 個の動作不良を調整した。',
              '減温塔 噴霧ノズル清掃、空気予熱器 伝熱管の肉厚測定。']:
        para(d, '・' + t, indent=4)
    head(d, '2.3　全炉停止整備', 2)
    for t in ['1号炉 炉内耐火物（側壁）部分補修 約0.5m³。',
              '1号炉 減温塔 下部胴の減肉部（最小 4.1mm）に当て板補修 1 か所（300×300、t6）。',
              'ごみクレーン 1号・2号 年次自主検査（外注）。',
              '空気圧縮機 1号 オーバーホール（運転時間 12,300h）。',
              '1号・2号 計装 校正、排ガス分析計 年次点検（外注）。',
              '投入扉 油圧ユニット 作動油交換、機器冷却水ポンプ(B) 分解整備。']:
        para(d, '・' + t, indent=4)
    head(d, '3．点検・測定結果')
    table(d, ['号炉', '機器', '項目', '基準・前回', '結果', '判定'], [
        ('1', '火格子（燃焼段）', '火格子片 残厚（平均／最小）', '原厚比 70%以上', '71% ／ 58%', '要注意'),
        ('1', '火格子（燃焼段）', '脱落', '－', '3片', '取替'),
        ('1', '減温塔', '下部胴 肉厚（最小）', '原厚 6.0mm／限界 4.5mm', '4.1mm', '不良（当て板）'),
        ('1', 'バグフィルタ', '差圧', '1.5kPa 以下', '1.3kPa', '良'),
        ('1', '誘引通風機', '軸受振動', '4.5mm/s 以下', '1.4mm/s', '良'),
        ('2', '火格子（燃焼段）', '火格子片 残厚（平均／最小）', '原厚比 70%以上', '68% ／ 55%', '要注意'),
        ('2', 'バグフィルタ', '差圧', '1.5kPa 以下', '1.8kPa', '要注意'),
        ('2', 'バグフィルタ', 'ろ布破損', '－', '2本', '交換'),
        ('2', '減温塔', '下部胴 肉厚（最小）', '原厚 6.0mm', '4.9mm', '要注意'),
        ('2', '空気予熱器', '伝熱管 肉厚（最小）', '原厚 3.2mm', '2.8mm', '良'),
        ('2', '誘引通風機', '軸受振動', '4.5mm/s 以下', '1.9mm/s', '良'),
        ('共', 'ごみクレーン 1号', '年次自主検査', '－', '異常なし（走行車輪 摩耗 軽微）', '良'),
        ('共', 'ごみクレーン 2号', '年次自主検査', '－', '異常なし', '良'),
        ('共', '空気圧縮機 1号', '吐出圧力・温度（OH後）', '0.7MPa／90℃以下', '0.70MPa／82℃', '良'),
    ], [9, 30, 38, 34, 38, 17], 8.5)
    head(d, '4．所見')
    for t in ['1号炉・2号炉とも燃焼段の火格子片の摩耗が進んでいる。令和4〜5年度の基幹的設備改良工事で更新が予定されているため、本年度は脱落・割れ片の取替えにとどめた。',
              '1号炉減温塔の下部胴は限界厚さを下回る箇所があり当て板で補修したが、令和4年度の更新までの応急処置である。運転中は外面の点検を強化されたい。',
              '2号炉バグフィルタの差圧は上昇傾向にある。ろ布は平成28年度の全数交換から5年を経過しており、更新（令和5年度予定）までの間、破損ろ布の交換で対応する。予備ろ布の在庫は 22 本（本年度 3 本使用）。',
              'ごみクレーン1号の走行車輪にわずかな摩耗が見られる。次年度以降も計測を継続する。']:
        para(d, t, indent=4, after=6)
    head(d, '5．使用材料')
    table(d, ['品名', '仕様', '数量', '手配', '使用箇所'], [
        ('火格子片', 'ASR-30R 燃焼段用', '11片', '当社', '1号 3片・2号 8片'),
        ('ろ布', 'φ150×5,000 PTFE', '3本', '組合予備品', '1号 1本・2号 2本'),
        ('当て板', 'SS400 t6 300×300', '1枚', '当社', '1号 減温塔 下部胴'),
        ('キャスタブル', '高アルミナ質', '1,100kg', '当社', '1号 側壁'),
        ('作動油', 'ISO VG46', '200L', '当社', '投入扉 油圧ユニット'),
        ('OH部品', '空気圧縮機 12,000h 整備キット', '1式', '当社', '空気圧縮機 1号'),
        ('メカニカルシール・軸受', '冷却水ポンプ用', '1式', '当社', '機器冷却水ポンプ(B)'),
    ], [30, 46, 18, 24, 48], 9)
    para(d, '※ 写真は別冊「工事写真帳」による。', 9)
    d.save(path_docx)


# ======================================================================= R4（新章立て）
def build_r4(path_docx):
    name = R4_CONTRACT['name']
    d = newdoc(f'【架空】{name} 工事報告書', f'{MAKER} n1')
    cover(d, '工 事 報 告 書', name, '令和4年4月28日 ～ 令和5年3月17日', '（2号炉中間点検・全炉停止整備）', '令和4年12月9日')
    para(d, '目　次', 13, True, 'c', after=8)
    for t in ['第1章　総括', '第2章　工事概要', '第3章　工程実績', '第4章　設備別 施工内容', '第5章　点検・測定記録',
              '第6章　使用材料・支給品', '第7章　追加工事・特記事項', '第8章　今後の整備に関する提言', '付録　工事写真（抜粋）']:
        para(d, t, 11, indent=30, after=5)
    para(d, '※ 本年度より、基幹的設備改良工事の報告書と章立てをそろえた。', 9, color='595959', before=12, indent=30)
    d.add_page_break()
    head(d, '第1章　総括')
    para(d, '本年度は、基幹的設備改良工事（別契約）で 1号炉が 9 月から 12 月まで停止しているため、1号炉の中間点検は行わず、2号炉中間点検（6月）と全炉停止整備（10月）を実施した。全炉停止整備は改良工事の現地工事期間中に行い、停止期間・炉室ホイスト・全停電日を改良工事と調整した。主な結果は下表のとおりである。', after=6)
    table(d, ['No', '対象機器（機器番号）', '実施内容', '結果', '次回への申し送り'], [
        (1, '2号炉 火格子 燃焼段（2B-06）', '計測、割れ・摩耗片 24片 取替', '残厚 平均64%・最小52%', 'R5 基幹改良で更新'),
        (2, '2号炉 灰押出装置（2B-12）', '点検、摺動板 4枚 取替（追加）', '残厚 9mm（原厚16mm）、穴あき1か所', '3年後に摺動板 点検・計測'),
        (3, '2号炉 焼却炉本体（2B-09）', '側壁 耐火物 部分補修 0.8m³', '剥離部 補修完了', '次回 全炉停止で確認'),
        (4, '2号炉 バグフィルタ（2G-05）', 'ろ布点検、破損 8本 交換（6月2・10月6）', '差圧 1.9kPa（6月）', 'R5 基幹改良で更新'),
        (5, '2号炉 減温塔（2C-01）', '内部点検、ノズル清掃', '付着物 中程度', 'R5 基幹改良で更新'),
        (6, '2号炉 誘引通風機（2F-09）', '軸受給脂、振動測定', '2.1mm/s', '傾向監視'),
        (7, 'ごみクレーン 1号（0R-14）', 'ワイヤロープ交換（組合支給品）、年次自主検査、性能検査', '合格', '走行車輪 摩耗 計測継続'),
        (8, 'ごみクレーン 2号（0R-15）', '年次自主検査、性能検査', '合格', '－'),
        (9, '機器冷却水ポンプ(A)（0W-07）', '分解整備（メカシール・軸受交換）', '良', '－'),
        (10, '2号炉 計装・排ガス分析計', '校正、年次点検', '良', '－'),
    ], [8, 40, 46, 38, 34], 8.5)
    head(d, '第2章　工事概要')
    table(d, ['項目', '内容'], [
        ('工事件名', name), ('施設', f'{FAC}（30t/16h×2炉）'), ('工期', '令和4年4月28日～令和5年3月17日'),
        ('2号炉中間点検', '令和4年6月13日～6月17日（停止 6/10 立下げ～6/18 立上げ）'),
        ('全炉停止整備', '令和4年10月17日～10月29日（2号炉 10/14 立下げ～10/30 立上げ。1号炉は改良工事のため停止中）'),
        ('1号炉中間点検', '実施せず（基幹的設備改良工事の範囲で開放・更新するため。仕様書 第3条ただし書）'),
        ('体制', 'SV n1、機械1班 5名、機械2班 4名、電計班 3名、外注（足場・クレーン検査・分析計）'),
        ('他工事との調整', '基幹的設備改良工事（現場代理人 n3）：2号炉側の立入は SV の許可制。炉室ホイストは 10/20～22 本工事優先。10/23 全停電（受変電年次点検）は両工事とも作業中止。'),
    ], [36, 130])
    head(d, '第3章　工程実績')
    rows = []
    for (iid, kubun, nm, team, plan, act, ms, rev) in R4_ALL:
        fmt = lambda sp: '・'.join(f'{s.month}/{s.day}' + (f'～{e.month}/{e.day}' if e != s else '') for s, e in sp)
        diff = '追加' if rev else ('+1日' if plan != act else '')
        rows.append((kubun, nm.replace('【追加】', ''), fmt(plan) if not rev else '（改1）' + fmt(plan), fmt(act), diff))
    table(d, ['区分', '作業', '計画', '実績', '差'], rows, [14, 74, 32, 32, 14], 8.5)
    para(d, '耐火物補修は養生のため 1 日延びたが、全体工程（10/29 総合試運転、10/30 立上げ）への影響はなかった。', 9.5)
    head(d, '第4章　設備別 施工内容')
    sec = [
        ('4.1　燃焼設備（2号炉）', [
            '炉内のクリンカを除去したのち、火格子全段を計測した。燃焼段は残厚 平均 64%、最小 52% で、割れ・摩耗の大きい 24 片を取替えた（10/20・10/21）。乾燥段・後燃焼段は残厚 75% 以上で良好であった。令和5年度に基幹的設備改良工事で全段更新が予定されているため、取替は割れ・摩耗片に限った。',
            '灰押出装置を点検した結果、摺動板の摩耗が大きく（残厚 9mm／原厚 16mm）、端部に穴あき 1 か所があった。このまま運転を続けると灰汚水の漏れ・押出不良による炉停止のおそれがあるため、組合と協議のうえ摺動板 4 枚の取替を追加した（第7章）。',
            '側壁（燃焼段上部 左側）のキャスタブル剥離 約 0.8m³ をはつり、型枠施工・養生した。給じん装置のライナは摩耗 1mm 程度で良好。炉油圧ユニットはフィルタを交換し、作動油を採取して分析に出した（結果 良好）。']),
        ('4.2　通風設備（2号炉）', [
            '誘引通風機のケーシングを開放し羽根車の付着ダストを除去、軸受のグリスを入替えた。手入れ後の振動は 2.1mm/s（R3：1.9mm/s）。押込・二次・炉壁冷却送風機は異常なし。']),
        ('4.3　燃焼ガス冷却設備（2号炉）', [
            '減温塔を開放し内面の付着物を除去、噴霧ノズル 6 本を清掃した。令和5年度に更新予定のため、肉厚測定は前年度測定点（最小 4.9mm）の 4 点のみとし、最小 4.8mm であった。']),
        ('4.4　排ガス処理設備（2号炉）', [
            'バグフィルタのろ布を全数目視点検し、破損 6 本（第3室 4本、第4室 2本）を組合予備品と交換した。蛍光粉によるリークテストで漏れのないことを確認した。中間点検（6月）の交換 2 本と合わせ、本年度の使用は 8 本、予備品の残りは 14 本である。消石灰・活性炭 定量供給機は異常なし。']),
        ('4.5　受入供給設備（共通）', [
            'ごみクレーン 1号のワイヤロープ 2 本を組合支給品と交換した（前回 R2）。1号・2号の年次自主検査を行い、10/20 に登録性能検査機関の性能検査を受け、いずれも合格した。投入扉 油圧ユニットに作動油 20L を補充した。']),
        ('4.6　給水・圧縮空気設備（共通）', [
            '機器冷却水ポンプ(A) を分解し、メカニカルシールと軸受を交換した。空気圧縮機 1号・2号は点検のみで異常なし。']),
        ('4.7　電気・計装設備（2号炉）', [
            '2号炉の温度・圧力・流量計を校正し、排ガス分析計は外注で年次点検・校正を行った。2号 MCC・炉前盤・BF 盤を点検した。10/23 の全停電（電気保安協会の年次点検、別契約）では停電・復電の操作に立会った。']),
    ]
    for h, ts in sec:
        head(d, h, 2)
        for t in ts:
            para(d, t, indent=4, after=5)
    head(d, '第5章　点検・測定記録')
    table(d, ['機器番号', '機器', '項目', '基準・前回', '結果', '判定'], [
        ('2B-06', '火格子（燃焼段）', '残厚 平均／最小', 'R3：68%／55%', '64%／52%', '取替（部分）'),
        ('2B-05', '火格子（乾燥段）', '残厚 平均', '原厚比 70%以上', '78%', '良'),
        ('2B-07', '火格子（後燃焼段）', '残厚 平均', '原厚比 70%以上', '76%', '良'),
        ('2B-12', '灰押出装置', '摺動板 残厚', '原厚 16mm', '9mm（穴あき1）', '取替（追加）'),
        ('2F-09', '誘引通風機', '軸受振動', '4.5mm/s 以下／R3 1.9', '2.1mm/s', '良'),
        ('2C-01', '減温塔', '下部胴 肉厚（最小）', 'R3：4.9mm', '4.8mm', '要注意'),
        ('2G-05', 'バグフィルタ', '差圧（6月）', '1.5kPa 以下', '1.9kPa', '要注意'),
        ('2I-12', '排ガス分析計', '校正', '－', 'スパン調整 CO・HCl', '良'),
        ('0R-14', 'ごみクレーン 1号', '性能検査', '－', '合格', '良'),
        ('0R-15', 'ごみクレーン 2号', '性能検査', '－', '合格', '良'),
    ], [16, 30, 30, 32, 34, 22], 8.5)
    head(d, '第6章　使用材料・支給品')
    table(d, ['品名', '仕様', '数量', '区分', '使用箇所', '備考'], [
        ('火格子片', 'ASR-30R 燃焼段用', '24片', '当社手配', '2号炉 燃焼段', ''),
        ('摺動板', '灰押出装置用 SS400 t16', '4枚', '当社手配（追加）', '2号炉 灰押出装置', 'メーカー在庫品'),
        ('キャスタブル', '高アルミナ質', '1,800kg', '当社手配', '2号炉 側壁', ''),
        ('ろ布', 'φ150×5,000 PTFE', '8本', '組合予備品', '2号炉 BF（6月2本・10月6本）', '残 14本'),
        ('ワイヤロープ', 'IWRC 6×Fi(29) φ22 85m', '2本', '組合支給品', 'ごみクレーン 1号', '旧ロープは組合で処分'),
        ('作動油', 'ISO VG46', '20L', '当社手配', '投入扉 油圧ユニット', '補充'),
        ('油圧フィルタ', '炉油圧ユニット用', '1組', '当社手配', '2号炉', ''),
        ('メカニカルシール・軸受', '冷却水ポンプ用', '1式', '当社手配', '機器冷却水ポンプ(A)', ''),
    ], [24, 40, 14, 26, 36, 26], 8.5)
    head(d, '第7章　追加工事・特記事項')
    para(d, '(1) 2号炉 灰押出装置 摺動板取替（追加）', 10.5, True)
    para(d, '10/19 の点検で摺動板の著しい摩耗と穴あきを確認し、10/20 に組合と協議して追加施工とした（工程表 改1）。摺動板はメーカー在庫品を手配し 10/26 に着荷、10/26～27 に取替え、10/28 に水張り・作動確認を行った。本件は変更契約（令和4年11月）の対象である。', indent=4, after=6)
    para(d, '(2) 1号炉中間点検', 10.5, True)
    para(d, '1号炉は基幹的設備改良工事のため 9/20 から停止しており、火格子・減温塔・バグフィルタは同工事で更新される。このため本業務の 1号炉中間点検は行っていない。', indent=4, after=6)
    para(d, '(3) 改良工事との取合い', 10.5, True)
    para(d, '炉室ホイストは 10/20～22 に 2号炉火格子片の荷揚げに使用し、改良工事の新火格子搬入は 10/24 以降とした。全停電日（10/23）は両工事とも作業を中止した。', indent=4, after=6)
    head(d, '第8章　今後の整備に関する提言')
    for t in ['2号炉の火格子・減温塔・バグフィルタは令和5年度の基幹的設備改良工事で更新されるため、令和5年度の全炉停止整備では更新範囲外の機器（灰押出装置、給じん装置、通風機類、耐火物）を中心とする。',
              '灰押出装置の摺動板は今回取替えたため、3 年後（令和7年度）を目安に残厚を計測し、取替の要否を判断されたい。1号炉の摺動板は令和5年度に計測を推奨する。',
              'ごみクレーン 1号の走行車輪は摩耗が進んでおり（R3 から 1.5mm 増）、令和7〜8年度の車輪交換・減速機整備の計画化を推奨する。',
              '2号炉 誘引通風機の振動は 2.1mm/s で基準内であるが、毎年の測定で傾向を確認されたい。']:
        para(d, '・' + t, indent=4, after=5)
    d.add_page_break()
    head(d, '付録　工事写真（抜粋）')
    caps = ['写真1　2号炉 燃焼段 火格子片 取替前', '写真2　2号炉 燃焼段 火格子片 取替後', '写真3　2号炉 灰押出装置 摺動板 穴あき部',
            '写真4　2号炉 灰押出装置 摺動板 取替後', '写真5　2号炉 側壁 耐火物 補修後', '写真6　ごみクレーン1号 ワイヤロープ 交換']
    t = d.add_table(rows=6, cols=2); t.style = 'Table Grid'; t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, cap in enumerate(caps):
        rr, cc = (i // 2) * 2, i % 2
        c = t.cell(rr, cc); c.text = ''
        shade(c, 'D9D9D9')
        p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        setfont(p.add_run('\n\n\n（写真）\n\n\n'), 10, color='808080')
        c2 = t.cell(rr + 1, cc); c2.text = ''
        setfont(c2.paragraphs[0].add_run(cap), 9)
        c.width = c2.width = Mm(82)
    t.autofit = False; t.columns[0].width = Mm(82); t.columns[1].width = Mm(82)
    d.save(path_docx)


def to_pdf(docx_path, pdf_name):
    pdf = render.to_pdf(docx_path, WORK)
    dst = os.path.join(OUT, pdf_name)
    os.replace(pdf, dst)
    return dst


if __name__ == '__main__':
    a = os.path.join(WORK, f'{FN}_工事報告書_令和3年度定期整備業務委託_n1旧章立て.docx')
    b = os.path.join(WORK, f'{FN}_工事報告書_令和4年度定期整備業務委託_n1新章立て.docx')
    build_r3(a); build_r4(b)
    print(to_pdf(a, os.path.basename(a).replace('.docx', '.pdf')))
    print(to_pdf(b, os.path.basename(b).replace('.docx', '.pdf')))
