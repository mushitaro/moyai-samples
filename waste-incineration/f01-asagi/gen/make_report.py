# -*- coding: utf-8 -*-
from docx import Document
from docx.shared import Pt, Mm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from xlutil import OUT, render
from f01_master import SYS, SYS_ORDER
from f01_history import EQ, BYTAG, events_in
from f01_contracts import *
import f01_truth as TR

def setfont(run, name, size=None, bold=None, color=None):
    run.font.name = name
    rpr = run._element.get_or_add_rPr()
    rf = rpr.find(qn('w:rFonts'))
    if rf is None:
        rf = OxmlElement('w:rFonts'); rpr.append(rf)
    rf.set(qn('w:eastAsia'), name); rf.set(qn('w:ascii'), name); rf.set(qn('w:hAnsi'), name)
    if size: run.font.size = Pt(size)
    if bold is not None: run.font.bold = bold
    if color: run.font.color.rgb = RGBColor.from_string(color)

def para(doc, text, font, size=10.5, bold=False, align=None, before=0, after=2, indent=None, color=None):
    p = doc.add_paragraph()
    r = p.add_run(text); setfont(r, font, size, bold, color)
    if align == 'c': p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if align == 'r': p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    pf = p.paragraph_format
    pf.space_before = Pt(before); pf.space_after = Pt(after)
    if indent: pf.first_line_indent = Pt(indent)
    return p

def shade(cell, color):
    tcPr = cell._element.get_or_add_tcPr()
    sh = OxmlElement('w:shd'); sh.set(qn('w:val'), 'clear'); sh.set(qn('w:color'), 'auto'); sh.set(qn('w:fill'), color)
    tcPr.append(sh)

def table(doc, rows, font, size=9, widths=None, header=True, hdr_color='D9D9D9', style='Table Grid'):
    t = doc.add_table(rows=len(rows), cols=len(rows[0]))
    t.style = style
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, row in enumerate(rows):
        for j, v in enumerate(row):
            c = t.cell(i, j)
            c.text = ''
            p = c.paragraphs[0]
            r = p.add_run('' if v is None else str(v)); setfont(r, font, size, bold=(header and i == 0))
            p.paragraph_format.space_after = Pt(0)
            if header and i == 0:
                shade(c, hdr_color); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if widths:
                c.width = Mm(widths[j])
    if widths:
        t.autofit = False
        for j, w in enumerate(widths):
            t.columns[j].width = Mm(w)
    return t

def kv_table(doc, items, font, size=10, w=(35, 125)):
    t = doc.add_table(rows=len(items), cols=2); t.style = 'Table Grid'
    for i, (k, v) in enumerate(items):
        a, b = t.cell(i, 0), t.cell(i, 1)
        a.text = ''; b.text = ''
        setfont(a.paragraphs[0].add_run(k), font, size); setfont(b.paragraphs[0].add_run(v), font, size)
        shade(a, 'F2F2F2'); a.width = Mm(w[0]); b.width = Mm(w[1])
    t.autofit = False
    t.columns[0].width = Mm(w[0]); t.columns[1].width = Mm(w[1])
    return t

def jdate(x):
    return '平成%d年%d月%d日' % (x.year - 1988, x.month, x.day) if x.year < 2019 or (x.year == 2019 and x.month < 5) else '令和%d年%d月%d日' % (x.year - 2018, x.month, x.day)

# ---------------------------------------------------------------- 結果の記述（代表は個別、その他は区分に応じて）
RES = {
 (2011, 'ST-101'): '割れ片18枚を取替（市予備品使用）。摩耗は最大3.2mmで限度内。',
 (2011, 'ST-102'): '抜取66枚を計測。最大摩耗6.8mm（限度12mm）。H22部分取替範囲以外で摩耗が進んでおり、H25〜26年度の部分取替を提案する。',
 (2011, 'ST-103'): '摩耗計測を実施。部分取替は別契約（後燃焼火格子補修工事）で施工。',
 (2011, 'SH-101'): '1次過熱器上段 最小4.0mm（新管4.5mm）。減肉速度は小さい。',
 (2011, 'SH-201'): '1次・2次全段を測定。2次過熱器 出口側下段で最小3.6mm（新管4.5mm、計算必要厚3.1mm）。H26年度頃の部分取替が必要と判断する。',
 (2011, 'BF-101'): 'ろ布4本の引張強度保持率 平均82%。交換時期はH25年度と想定。',
 (2011, 'BF-201'): 'ろ布4本の引張強度保持率 平均87%。',
 (2011, 'F-103'): '羽根車付着ダスト除去後、振動 2.0mm/s（運転時）。前縁肉盛部（H15）は健全。',
 (2011, 'F-101'): '反負荷側軸受外輪に圧こんを認め、両側軸受を交換。復旧後 振動1.4mm/s。',
 (2011, 'F-201'): '開放点検の結果、異常なし。振動1.6mm/s。',
 (2011, 'GC-101'): '下部コーン付着物 約0.3m³を除去。H22当て板部に異常なし。',
 (2011, 'TB-01'): '非常調速機・低油圧トリップ等の保安装置作動試験良好。',
 (2011, 'P-201B'): 'メカニカルシール・軸受を交換（受注者手配品）。市予備品（2組）は緊急時用として未使用。',
 (2011, 'B-203'): '最小2.9mm（計算必要厚3.1mm未満）の管12本を確認。取替は別契約（ボイラ蒸発管補修工事）で施工。',
 (2011, 'CR-01'): '巻上・横行ワイヤロープ（φ22）を交換。',
 (2025, 'ST-101'): '最大摩耗2.4mm。限度内。',
 (2025, 'ST-102'): '最大摩耗4.2mm（限度12mm）。割れ片12枚を交換（市予備品使用）。R10年度に部分取替（約20%）を提案。',
 (2025, 'F-101'): '軸受2個を交換。復旧後 振動1.1mm/s。',
 (2025, 'F-103'): '羽根車付着ダスト除去。振動1.6mm/s（インバータ定格周波数）。R5交換の軸受に異常なし。',
 (2025, 'INV-101'): '冷却ファン2台を交換。平滑コンデンサ容量は管理値内（製造後8年）。',
 (2025, 'GC-101'): '下部コーン腐食部2箇所（最小3.1mm／原厚6.0mm）に当て板（t6、計0.8m²）。腐食範囲が広がっており、R9年度に下部コーン部分更新を提案。',
 (2025, 'BF-101'): '市支給ろ布288本を全数交換。リテーナ腐食24本を交換（受注者手配）。各室漏れ確認良好。差圧は交換前2.3kPa→交換後1.1kPa（運転時）。',
 (2025, 'SH-201'): '1次・2次全段を測定。最小4.2mm（新管4.5mm、H30更新）。減肉速度は年0.04mm程度で、廃止目標（R16）まで交換不要と判断。',
 (2025, 'EC-201'): '最小4.1mm（新管4.5mm）。',
 (2025, 'P-201A'): 'メカシール・軸受・ライナリングを交換。',
 (2025, 'P-101B'): 'メカシール・軸受を交換。ライナリングすき間0.32mmで継続使用。',
 (2025, 'TB-01'): '保安装置作動試験良好。軸受振動 最大 18μm（p-p）。',
 (2025, 'ACC-01'): '6セルのフィン洗浄。復水器真空度 運転時 −88kPa→−91kPa に回復。',
 (2025, 'CR-01'): '巻上ワイヤロープ（φ22×2本）を交換。',
 (2025, 'B-101'): 'ボイラ開放点検（法定）。ドラム内部 スケール付着軽微。',
 (2025, 'TV-201'): 'カメラ・モニタを更新（製造中止品のため）。',
 (2025, 'AHU-01'): 'H24設置品を更新（冷媒漏れ）。',
}
def result(fy, d, k, t):
    if (fy, d['tag']) in RES:
        return RES[(fy, d['tag'])]
    if k == '点検':
        return '異常なし。'
    if k == '計測':
        return '測定値は管理値内。'
    if k == '整備':
        return '整備後の試運転良好。'
    if k == '補修':
        return '補修後の作動確認良好。'
    return '更新後の試運転良好。'

def judge(fy, d, k):
    if (fy, d['tag']) in (('2011', ''),):
        return 'A'
    b_list = {2011: {'ST-102', 'SH-201', 'BF-101'}, 2025: {'ST-102', 'GC-101', 'P-101B'}}
    return 'B' if d['tag'] in b_list.get(fy, set()) else 'A'

def name_of(d, fy):
    if fy < 2019:
        n = d['name']
        return ('%s %s' % ({1: '1号炉', 2: '2号炉', 0: ''}[d['f']], n)).strip()
    return d['name']

# ---------------------------------------------------------------- r1 様式（〜2015）
def build_r1():
    c = H23
    fn = 'ＭＳ 明朝'; fg = 'ＭＳ ゴシック'
    doc = Document()
    s = doc.sections[0]; s.page_width, s.page_height = Mm(210), Mm(297)
    s.left_margin = s.right_margin = Mm(22); s.top_margin = Mm(20); s.bottom_margin = Mm(18)
    # 表紙
    for _ in range(5): para(doc, '', fn)
    para(doc, '平成23年度', fn, 16, align='c')
    para(doc, '【架空】あさぎ市清掃工場', fn, 20, True, align='c', after=6)
    para(doc, '定期整備業務委託', fn, 20, True, align='c', after=30)
    para(doc, '工　事　報　告　書', fn, 26, True, align='c', after=80)
    para(doc, '平成24年1月', fn, 14, align='c', after=40)
    para(doc, '汐見重工環境株式会社', fn, 14, align='c')
    para(doc, '環境プラント事業部', fn, 12, align='c', after=30)
    t = table(doc, [['承認', '照査', '作成'], ['', '', 'r1']], fn, 9, widths=(18, 18, 18))
    t.alignment = WD_TABLE_ALIGNMENT.RIGHT
    doc.add_page_break()
    para(doc, '目　　次', fn, 14, True, align='c', after=12)
    for l in ['１．工事概要', '２．実施工程', '３．実施内容', '　３．１　１号炉', '　３．２　全停（共通設備）', '　３．３　２号炉', '　３．４　共通設備',
              '４．点検・計測結果', '５．使用部品・材料', '６．所見および今後の整備について', '付表　整備対象機器一覧']:
        para(doc, l, fn, 11, after=4)
    doc.add_page_break()
    para(doc, '１．工事概要', fg, 12, True, after=6)
    kv_table(doc, [('件　名', c['title']), ('施工場所', 'あさぎ市清掃工場（あさぎ市 環境部）'),
                   ('契約工期', '平成23年6月1日　〜　平成24年1月31日'),
                   ('施工期間', '１号炉　平成23年7月4日〜8月5日\n全　停　平成23年7月20日〜7月22日（２号炉 7月19日停止、7月23日立上げ）\n２号炉　平成23年10月17日〜11月18日'),
                   ('発注者', 'あさぎ市'), ('受注者', '汐見重工環境株式会社　環境プラント事業部'), ('現場代理人', 'm2'),
                   ('対象設備', '焼却施設（110t/日×2炉）、余熱利用設備（蒸気タービン発電機 2,600kW）、受変電設備ほか共通設備'),
                   ('委託範囲', '定期整備仕様書による点検・整備。補修工事（別契約）4件、受変電設備の年次点検（電気保安業務委託）は含まない')], fn, 10, w=(30, 136))
    para(doc, '', fn)
    para(doc, '２．実施工程', fg, 12, True, after=6)
    table(doc, [['区分', '計画', '実績', '備考'],
                ['１号炉', '7/4 停止 〜 8/5 試運転完了', '7/4 停止 〜 8/5 試運転完了', '炉開放は温度低下待ちで 7/7 10:00（1時間遅れ）'],
                ['全停', '7/20 〜 7/22', '7/20 〜 7/22', '２号炉 7/19 21:00 停止、7/23 立上げ'],
                ['２号炉', '10/17 停止 〜 11/18 試運転完了', '10/17 停止 〜 11/18 試運転完了', '別契約2件と炉内作業区域を分割']], fn, 9, widths=(20, 42, 42, 62))
    para(doc, '', fn)
    para(doc, '３．実施内容', fg, 12, True, after=6)
    texts = {
     '１号炉': [
      '焼却停止後3日間の自然冷却ののち、7月7日に酸素・一酸化炭素濃度を確認して炉およびボイラを開放した。炉内のクリンカは側壁の燃焼段から後燃焼段に多く、約3.5tを除去した。',
      '火格子は乾燥・燃焼・後燃焼の各段で摩耗計測を行った。燃焼火格子は抜取66枚の最大摩耗が6.8mmで、平成22年度に取り替えた範囲以外で摩耗が進んでいる。乾燥火格子は割れ18枚を確認し、市の予備品で取り替えた。後燃焼火格子の部分取替は別契約の補修工事で施工されたため、本委託では計測と工程調整のみを行った。',
      '耐火物は炉側壁（左）燃焼段上部1.2m²と天井部0.4m²をはつり、アンカーを溶接したうえでキャスタブルを打設した。乾燥は立上げ時の昇温計画に組み込んだ。',
      'ボイラはドラムを開放して内部を点検し、スケール付着は軽微であった。過熱器・蒸発管の肉厚を測定した結果は４．に示す。スートブロワNo.3の噴射管を交換した。',
      '減温塔は下部コーンの付着物を除去し、平成22年度に当て板補修した箇所に異常のないことを確認した。バグフィルタは各室1本のろ布を抜き取って強度試験を行い、出口ダンパのシール類を交換した。',
      '押込送風機は開放点検で反負荷側軸受外輪に圧こんを認めたため、両側の軸受を交換した。誘引通風機は羽根車の付着ダストを除去し、平成15年度に肉盛した前縁部が健全であることを確認した。煙道（BF出口〜煙突）底部の腐食2箇所に当て板を溶接した。主灰搬出コンベヤはチェーンとフライトを交換した。'],
     '全停（共通設備）': [
      '7月20日8時30分に受電を停止し、6.6kV高圧配電盤5面と低圧主幹盤2面の精密点検、共通動力制御盤3面・進相コンデンサ盤・励磁装置盤の点検を行った。停電中は非常用発電機で中央制御室と保安照明を確保した。受電は7月21日16時に復帰した。',
      '蒸気タービンは保安装置の作動試験を行い、いずれも良好であった。入口主蒸気止め弁のアクチュエータ、空冷復水器ファンNo.1の減速機、脱気器圧力調節弁・ホットウェル水位調節弁を整備した。'],
     '２号炉': [
      '10月20日に開放した。給じん装置は摺動板とシリンダパッキンを整備した。ボイラは蒸発管と過熱器全段の肉厚を測定し、蒸発管の減肉部は別契約（ボイラ蒸発管補修工事）で取り替えられた。2次過熱器の出口側下段で減肉が進んでおり、所見を６．に記す。',
      'ボイラ給水ポンプBはメカニカルシールと軸受を交換し、清缶剤注入ポンプを更新した。バグフィルタはろ布の抜取検査と、入口・出口・バイパス・第2室出口ダンパの整備を行った。耐火物補修は別契約で施工されたため、炉内の作業区域を分けて調整した。'],
     '共通設備': [
      'ごみクレーン1号機のワイヤロープを交換し、この間は2号機で運用した。プラットホーム投入扉No.1、計量機（搬出）のロードセル、ピット脱臭ファンを整備・点検した。',
      '排水処理設備は排水移送ポンプAを整備し、凝集剤・pH調整剤（アルカリ）注入ポンプを更新、キレート吸着塔のろ材を交換した。中央制御室空調機は冷房能力の低下のため更新した。'],
    }
    for i, (k, ps) in enumerate(texts.items(), 1):
        para(doc, '３．%d　%s' % (i, k), fg, 11, True, before=4, after=3)
        for ptxt in ps:
            para(doc, ptxt, fn, 10.5, indent=10.5, after=3)
    para(doc, '４．点検・計測結果', fg, 12, True, before=8, after=6)
    para(doc, '（１）火格子摩耗（限度 12mm）', fn, 10.5, after=2)
    table(doc, [['号炉', '段', '計測枚数', '平均摩耗', '最大摩耗', '処置'],
                ['１号炉', '乾燥', '42', '1.9mm', '3.2mm', '割れ18枚取替'], ['１号炉', '燃焼', '66', '4.1mm', '6.8mm', '経過観察'], ['１号炉', '後燃焼', '30', '5.6mm', '9.4mm', '別契約で部分取替'],
                ['２号炉', '乾燥', '42', '2.2mm', '3.8mm', '—'], ['２号炉', '後燃焼', '30', '3.9mm', '6.1mm', '—']], fn, 9)
    para(doc, '', fn)
    para(doc, '（２）ボイラ管肉厚（新管 4.5mm、計算必要厚 3.1mm）', fn, 10.5, after=2)
    table(doc, [['号炉', '部位', '測定点数', '最小値', '前回最小値', '判定'],
                ['１号炉', '1次過熱器', '96', '4.0mm', '4.1mm（H21）', '良'], ['１号炉', '2次過熱器', '52', '3.9mm', '4.0mm（H21）', '良'],
                ['２号炉', '1次過熱器', '96', '4.0mm', '4.1mm（H18）', '良'], ['２号炉', '2次過熱器', '104', '3.6mm', '3.9mm（H18）', '要注意'],
                ['２号炉', '蒸発管群', '80', '2.9mm', '3.5mm（H20）', '別契約で取替']], fn, 9)
    para(doc, '', fn)
    para(doc, '（３）バグフィルタろ布 引張強度保持率', fn, 10.5, after=2)
    table(doc, [['号炉', '抜取本数', '保持率（平均）', '前回', '備考'], ['１号炉', '4', '82%', '91%（H19交換後）', '交換時期 H25年度想定'], ['２号炉', '4', '87%', '—', '']], fn, 9)
    para(doc, '５．使用部品・材料', fg, 12, True, before=8, after=6)
    table(doc, [['品名', '数量', '使用箇所', '区分'],
                ['乾燥火格子片', '18枚', '1号炉 乾燥火格子', '市予備品（使用）'],
                ['軸受（押込送風機用）', '2個', '1号炉 押込送風機', '受注者手配'],
                ['メカニカルシール（給水ポンプ用）', '1組', '2号炉 ボイラ給水ポンプB', '受注者手配'],
                ['メカニカルシール（給水ポンプ用）', '2組', '—', '市予備品（H23.9納入）。緊急時用として未使用、倉庫保管'],
                ['キャスタブル（高アルミナ質）', '1.9t', '1号炉 炉側壁・天井', '受注者手配'],
                ['ワイヤロープ φ22', '2本', 'ごみクレーン1号機', '受注者手配'],
                ['チェーン・フライト一式', '1式', '1号炉 主灰搬出コンベヤ', '受注者手配']], fn, 9)
    para(doc, '６．所見および今後の整備について', fg, 12, True, before=8, after=6)
    for ptxt in ['（１）２号炉2次過熱器の出口側下段は、前回（H18）から5年で0.3mm減肉しており、このままでは平成26年度頃に計算必要厚に近づく。平成26年度の定期整備時に該当管の部分取替を行うことを提案する。',
                 '（２）１号炉燃焼火格子は、平成22年度の部分取替範囲以外の摩耗が6.8mmに達した。平成25〜26年度の部分取替を計画されたい。',
                 '（３）１号炉バグフィルタのろ布は強度保持率が82%に低下しており、平成25年度の全数交換が適当と考える。ろ布は納期が約4か月であるため、前年度中の手配が必要である。',
                 '（４）竣工から15年目を迎え、火格子・過熱器・誘引通風機・受変電設備など主要設備の劣化が並行して進んでいる。個別の補修に加え、延命化を前提とした計画的な更新の検討を提案する。']:
        para(doc, ptxt, fn, 10.5, after=4)
    doc.add_page_break()
    para(doc, '付表　整備対象機器一覧（本委託の範囲）', fg, 11, True, after=4)
    rows = [['号炉', '設備', '機器名', '機器番号', '実施内容', '結果']]
    evs = [e for e in events_in(2011) if e[4] == 'T']
    evs.sort(key=lambda e: ({1: 0, 0: 1, 2: 2}[e[0]['f']] if e[0]['sys'] not in ('J', 'L') else 1, SYS_ORDER.index(e[0]['sys'])))
    TR.set_file('【架空】あさぎ市清掃工場_工事報告書_平成23年度定期整備業務委託_r1様式.docx')
    for (d, k, t, cst, ct) in evs:
        rows.append([{1: '1号', 2: '2号', 0: '共通'}[d['f']], SYS[d['sys']].replace('設備', ''), d['name'], d['tag'], t, result(2011, d, k, t)])
        TR.r('付表', len(rows) - 1, 'report_table_row', [x['id'] for x in TR._BY[(2011, d['tag'])] if x['text'] == t], d['code'], d['name'])
    table(doc, rows, fn, 7.5, widths=(10, 18, 34, 16, 40, 50))
    para(doc, '※別契約の補修工事（1号炉後燃焼火格子、2号炉耐火物、2号炉ボイラ蒸発管、ごみピット壁面）は各工事の報告書による。', fn, 8, before=4)
    p = OUT + '/【架空】あさぎ市清掃工場_工事報告書_平成23年度定期整備業務委託_r1様式.docx'
    doc.save(p)
    return p

# ---------------------------------------------------------------- r2 様式（2016〜）
def build_r2():
    c = R7
    fg = 'ＭＳ ゴシック'; fn = 'ＭＳ Ｐゴシック'
    doc = Document()
    s = doc.sections[0]; s.page_width, s.page_height = Mm(210), Mm(297)
    s.left_margin = s.right_margin = Mm(15); s.top_margin = Mm(15); s.bottom_margin = Mm(15)
    hp = s.header.paragraphs[0]; r = hp.add_run('【架空】あさぎ市清掃工場　令和7年度 定期整備業務委託　工事報告書　　報告書番号 SKE-AS-R07-03'); setfont(r, fn, 8, color='595959')
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    fp = s.footer.paragraphs[0]; r = fp.add_run('汐見重工環境株式会社 環境プラント事業部　（作成 r2）'); setfont(r, fn, 8, color='595959')
    para(doc, '【架空】工事報告書（サマリー）', fg, 16, True, after=4)
    kv_table(doc, [('件名', c['title']), ('契約工期', '令和7年5月12日〜令和8年1月30日'),
                   ('施工期間', '1号炉 R7.6.2〜7.4／全停 R7.6.17〜6.19／2号炉 R7.9.29〜10.31'),
                   ('受注者', '汐見重工環境株式会社 環境プラント事業部（現場代理人 m3）'), ('報告日', '令和7年12月12日')], fn, 9, w=(28, 152))
    para(doc, '', fn, 4)
    para(doc, '1. 実施概要', fg, 11, True, after=2)
    evs = [e for e in events_in(2025) if e[4] == 'T']
    n1 = len([e for e in evs if e[0]['f'] == 1]); n2 = len([e for e in evs if e[0]['f'] == 2]); n0 = len([e for e in evs if e[0]['f'] == 0])
    para(doc, '定期整備仕様書に基づき、1号炉・2号炉の順に炉を停止して点検・整備を行い、1号炉の施工期間中に全停（受変電停電・蒸気タービン・共通蒸気系）を行った。対象機器は1号炉%d件、2号炉%d件、共通%d件の計%d件（一覧は2.）。'
         '主な作業は1号炉バグフィルタのろ布全数交換（市支給品）、両炉の火格子・ボイラ管の計測、蒸気タービン簡易点検と空冷復水器の洗浄である。'
         '1号炉バグフィルタでリテーナの腐食24本を発見して交換を追加したため、同作業の完了が2日延びたが、復旧・立上げの予備日で吸収し、全体工程は計画どおり完了した。' % (n1, n2, n0, len(evs)), fn, 9.5, after=2)
    para(doc, '別契約の補修工事（1号炉ボイラ蒸発管、2号炉押込送風機軸受、2号炉後燃焼火格子、灰ピット壁面）は本報告の対象外である（工程調整のみ実施）。', fn, 9, after=4, color='404040')
    para(doc, '2. 対象機器と実施内容', fg, 11, True, after=2)
    para(doc, '判定　A：良好　B：経過観察（次回以降の計画に反映）　C：要対策', fn, 8, after=2)
    rows = [['機器コード', '号炉', '機器名', '区分', '実施内容', '判定']]
    order = {1: 0, 0: 1, 2: 2}
    evs.sort(key=lambda e: (order[e[0]['f']], e[0]['code']))
    TR.set_file('【架空】あさぎ市清掃工場_工事報告書_令和7年度定期整備業務委託_r2様式.docx')
    for (d, k, t, cst, ct) in evs:
        rows.append([d['code'], {1: '1号炉', 2: '2号炉', 0: '共通'}[d['f']], d['name'], k, t, judge(2025, d, k)])
        TR.r('2.対象機器と実施内容', len(rows) - 1, 'report_table_row', [x['id'] for x in TR._BY[(2025, d['tag'])] if x['text'] == t], d['code'], d['name'])
    tb = table(doc, rows, fn, 7.5, widths=(17, 12, 44, 12, 86, 9), hdr_color='BDD7EE')
    for i, row in enumerate(rows[1:], 1):
        if row[5] == 'B':
            shade(tb.cell(i, 5), 'FFF2CC')
    doc.add_page_break()
    para(doc, '3. 主な確認結果', fg, 11, True, after=3)
    key = ['BF-101', 'ST-102', 'GC-101', 'SH-201', 'F-103', 'INV-101', 'TB-01', 'ACC-01', 'P-101B', 'B-101']
    for tg in key:
        d = BYTAG[tg]
        para(doc, '■ %s %s（%s）' % ({1: '1号炉', 2: '2号炉', 0: ''}[d['f']], d['name'], d['code']), fg, 9.5, True, before=3, after=1)
        para(doc, RES[(2025, tg)], fn, 9.5, after=1, indent=9.5)
    para(doc, '4. 工程の実績', fg, 11, True, before=6, after=3)
    table(doc, [['区分', '計画', '実績', '差異の理由'],
                ['1号炉', '6/2停止〜7/4完了', '6/2停止〜7/4完了', 'BFろ布交換 6/24→6/26完了（リテーナ交換追加）。復旧・立上げの予備日で吸収'],
                ['全停', '6/17〜6/19', '6/17〜6/19', '受電復帰 6/18 16:10'],
                ['2号炉', '9/29停止〜10/31完了', '9/29停止〜10/31完了', '別契約3件と炉内・灰ピットの作業区域を分割']], fn, 8.5, widths=(16, 34, 34, 96), hdr_color='BDD7EE')
    para(doc, '5. 支給品・部品の扱い', fg, 11, True, before=6, after=3)
    table(doc, [['品名', '受入数', '使用数', '残数の扱い', '区分'],
                ['ろ布（PTFE膜/ガラス φ165×6,000）', '300本', '288本', '12本を市倉庫へ返納（予備品）', '市支給品（R7.6.13 納入）'],
                ['燃焼火格子片', '12枚', '12枚', '—', '市予備品'],
                ['リテーナ（SUS304 φ160×6,000）', '24本', '24本', '—', '受注者手配（追加）'],
                ['インバータ冷却ファン', '2台', '2台', '—', '受注者手配'],
                ['ワイヤロープ φ22・φ18', '各2本', '各2本', '—', '受注者手配']], fn, 8.5, widths=(48, 16, 16, 55, 45), hdr_color='BDD7EE')
    para(doc, '6. 今後の整備に関する提案', fg, 11, True, before=6, after=3)
    for ptxt in ['(1) 1号炉減温塔 下部コーン：腐食範囲が広がっており、当て板では対応しきれなくなる。R9年度に下部コーンの部分更新（補修工事）を提案する。',
                 '(2) 1号炉燃焼火格子：R10年度に部分取替（約20%）を提案する。廃止目標（R16年度）までの残存期間から、全数更新は不要と考える。',
                 '(3) 2号炉バグフィルタ：ろ布はR2年度交換で、R8年度の全数交換時期にあたる。R8年度当初にろ布の手配（納期約2か月）が必要。',
                 '(4) DCS：オペレータステーションのOS・ハードウェアの保守がR9年度で終了する通知を受けている。R9年度のHMI更新を検討されたい。']:
        para(doc, ptxt, fn, 9.5, after=2)
    p = OUT + '/【架空】あさぎ市清掃工場_工事報告書_令和7年度定期整備業務委託_r2様式.docx'
    doc.save(p)
    return p

if __name__ == '__main__':
    p = build_r1(); print(render(p, pages=(3, 4, 7)))
    p = build_r2(); print(render(p, pages=(1, 2, 3)))
