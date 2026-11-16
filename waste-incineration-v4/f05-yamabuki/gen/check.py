# -*- coding: utf-8 -*-
"""F05 正解と資料の突合せ。
1) answer/files/<stem>.json の各行が、資料のその行（xlsx は sheet・r、PDF は seq の順）に書いてあるとおりか
2) どのファイルにも【架空】がある（Excel：表題セルと文書の表題、PDF：1 ページ目と文書情報）
3) 実在の名前に当たらない（社名・施設名の禁止語）
4) 写真の大きさ（画像ファイルがあれば 1,568〜2,000px・JPEG 約 100KB）
5) world.json の結び（links）が実在の作業・ファイル・行を指す
"""
import os, sys, json, glob, re, unicodedata
from openpyxl import load_workbook
import pdfplumber

HERE = os.path.dirname(os.path.abspath(__file__))
FULL = os.path.dirname(HERE)
FILES = os.path.join(FULL, 'files')
ANS = os.path.join(FULL, 'answer')

SKIP = {'plans', 'days', 'manDays', 'manDaysText', 'code', '区分の値', '件名（全文）', 'path', 'works', 'plan', 'actual', 'marks', 'comments', 'cost', 'unit', '状態',
        '章', 'nameAsWritten', 'image', '区分（値）', '期間', '天候', '予定人員', '日別作業員', '列', '日付', '番号', '記入者', 'head', '負担行為日',
        'pastWork', 'oldNo', 'install', 'furnace', 'no', 'importance', 'standby', 'memo', '追加', '班', 'No', '人', '日', '種別', '工数', '人数', '人員',
        '作業員計', '延べ人数', '変更後額', '支出済額', '残額', '1号炉', '2号炉', '共通', '計', '負担行為額', '数量', 'WBS', '写真', '説明'}
REAL = ['日立造船', 'カナデビア', 'JFEエンジニアリング', '荏原環境', 'タクマ', '川崎重工', '三菱重工', '神鋼環境', 'クボタ', '住友重機', 'IHI', '栗田工業', '横河電機',
        '富士電機', '明電舎', '東芝', '三菱電機', '日鉄エンジニアリング', '新明和', 'プランテック', '極東開発', '水ing', 'メタウォーター', '月島', '前澤工業',
        '関東電気保安協会', '中部電気保安協会', '日本環境衛生センター', '太平洋セメント', '住友大阪セメント', 'トキコ', '鎌長製衡', '大和製衡']


def norm(s):
    s = unicodedata.normalize('NFKC', str(s))
    return re.sub(r'\s+', '', s).replace('〜', '~').replace('～', '~').replace('－', '-').replace('−', '-')


errs = []
nrec = 0


def cell_texts(ws, r):
    out = []
    for c in range(1, ws.max_column + 1):
        v = ws.cell(r, c).value
        if v is not None:
            out.append(v)
    return out


def val_in_row(v, row):
    if isinstance(v, bool):
        return True
    if isinstance(v, (int, float)):
        for x in row:
            if isinstance(x, (int, float)) and abs(float(x) - float(v)) < 0.06:
                return True
            if isinstance(x, str) and (f'{v:,}' in x or str(v) in x):
                return True
        return False
    t = norm(v)
    if not t:
        return True
    joined = norm(''.join(str(x) for x in row))
    return t in joined


def check_xlsx(path, recs):
    global nrec
    wb = load_workbook(path)
    for rc in recs:
        nrec += 1
        sh = rc.get('sheet')
        if sh not in wb.sheetnames:
            errs.append(f'{os.path.basename(path)}: シート無し {sh}'); continue
        ws = wb[sh]
        row = cell_texts(ws, rc['r'])
        f = rc['fields']
        if not row:
            errs.append(f'{os.path.basename(path)} {sh}!{rc["r"]}: 空の行'); continue
        if 'path' in f:
            nm = f.get('nameAsWritten') or f['path'][-1]
            if not (val_in_row(nm, row) or val_in_row('■' + nm, row) or val_in_row(f'【{nm}】', row) or nm in ('合計',) and val_in_row('合　計', row)):
                errs.append(f'{os.path.basename(path)} {sh}!{rc["r"]}: 名前 {nm} が無い')
        for y, s in (f.get('marks') or {}).items():
            if not val_in_row(s, row):
                errs.append(f'{os.path.basename(path)} {sh}!{rc["r"]}: 記号 {y}={s} が無い')
        for k, v in f.items():
            if k in SKIP or v is None:
                continue
            if isinstance(v, (dict, list)):
                continue
            if not val_in_row(v, row):
                errs.append(f'{os.path.basename(path)} {sh}!{rc["r"]}: {k}={v} が無い')


def check_pdf(path, recs):
    global nrec
    with pdfplumber.open(path) as pdf:
        text = norm(''.join(p.extract_text() or '' for p in pdf.pages))
        first = pdf.pages[0].extract_text() or ''
        meta = pdf.metadata or {}
    if '【架空】' not in first or '【架空】' not in str(meta.get('Title', '')):
        errs.append(f'{os.path.basename(path)}: 【架空】が 1 ページ目か文書情報に無い')
    grams = set(text[i:i + 3] for i in range(len(text) - 2))
    for rc in sorted(recs, key=lambda x: x['seq']):
        nrec += 1
        f = rc['fields']
        for k, v in f.items():
            if k in ('章', 'image') or not isinstance(v, (str, int)):
                continue
            t = norm(v)
            if t in text:
                continue
            # 表のセル内の折返しで行が混ざるため、3 文字のかけらの 70% 以上があればよい
            gs = [t[i:i + 3] for i in range(max(1, len(t) - 2))]
            hit = sum(1 for g in gs if g in grams) / len(gs)
            if hit < 0.7:
                errs.append(f'{os.path.basename(path)} seq{rc["seq"]}: {k}={str(v)[:30]} が無い（{hit:.0%}）')


def check_marks():
    for p in glob.glob(os.path.join(FILES, '*.xlsx')):
        wb = load_workbook(p)
        if '【架空】' not in (wb.properties.title or ''):
            errs.append(f'{os.path.basename(p)}: 文書の表題に【架空】が無い')
        ws = wb.worksheets[0]
        ok = any('【架空】' in str(ws.cell(r, c).value or '') for r in range(1, 7) for c in range(1, 12))
        if not ok:
            errs.append(f'{os.path.basename(p)}: 表題セルに【架空】が無い')
        for ws in wb.worksheets:
            for row in ws.iter_rows(values_only=True):
                for v in row:
                    if isinstance(v, str):
                        for n in REAL:
                            if n in v:
                                errs.append(f'{os.path.basename(p)}: 実在名 {n}')
        if not os.path.basename(p).startswith('【架空】'):
            errs.append(f'{os.path.basename(p)}: ファイル名')
    for p in glob.glob(os.path.join(FILES, '*.pdf')):
        with pdfplumber.open(p) as pdf:
            t = ''.join(pg.extract_text() or '' for pg in pdf.pages)
        for n in REAL:
            if n in t:
                errs.append(f'{os.path.basename(p)}: 実在名 {n}')


def check_photos():
    from PIL import Image
    n = 0
    for p in glob.glob(os.path.join(FILES, '*.jpg')) + glob.glob(os.path.join(FILES, '*.png')):
        n += 1
        im = Image.open(p)
        if not (1568 <= max(im.size) <= 2000):
            errs.append(f'{os.path.basename(p)}: 大きさ {im.size}')
        kb = os.path.getsize(p) / 1024
        if not (50 <= kb <= 200):
            errs.append(f'{os.path.basename(p)}: {kb:.0f}KB')
    return n


def check_links(recs_by_file):
    w = json.load(open(os.path.join(ANS, 'world.json'), encoding='utf-8'))
    if w.get('schema') != 'moyai.sample-world/0':
        errs.append('world.json: schema')
    if w.get('facility', {}).get('id') != 'F05' or '【架空】' not in w.get('facility', {}).get('name', ''):
        errs.append('world.json: facility')
    if w.get('maker', {}).get('today') != '2026-11-13':
        errs.append('world.json: maker.today')
    ids = {x['id'] for x in w['works']}
    aids = {a['id'] for a in w['assets']}
    files = {os.path.basename(p) for p in glob.glob(os.path.join(FILES, '*'))}
    if {f['file'] for f in w['files']} != files:
        errs.append('world.json: files[] が files/ と合わない')
    for a in w['assets']:
        for k in ('id', 'name', 'tag', 'system', 'furnace', 'installed', 'class', 'part', 'names'):
            if k not in a:
                errs.append(f'asset {a.get("id")}: {k} が無い')
    for x in w['works']:
        for k in ('id', 'asset', 'fy', 'status', 'task', 'attr', 'start', 'end', 'manDays', 'days'):
            if k not in x:
                errs.append(f'work {x["id"]}: {k} が無い')
        if x['attr'] not in ('INSPECT', 'OVERHAUL', 'REPLACE', 'UPGRADE'):
            errs.append(f'work {x["id"]}: attr')
        if x['asset'] and x['asset'] not in aids:
            errs.append(f'work {x["id"]}: asset {x["asset"]} が無い')
        if x['manDays'] is not None and not isinstance(x['manDays'], (int, float)):
            errs.append(f'work {x["id"]}: manDays が数でない')
    for l in w['links']:
        if l['work'] not in ids:
            errs.append(f'link: 作業 {l["work"]} が無い')
        if l['file'] not in files:
            errs.append(f'link: ファイル {l["file"]} が無い')
        key = (l['file'], l.get('sheet'), l.get('r'), l.get('seq'))
        if key not in recs_by_file:
            errs.append(f'link: 記録が無い {key}')
        if 'how' not in l:
            errs.append('link: how が無い')
    for e in w['errata']:
        if e['type'] not in ('omitMark', 'lastYearUnreflected'):
            errs.append(f'errata: type {e["type"]}')
        if (e['file'], e['sheet'], e['r'], None) not in recs_by_file:
            errs.append(f'errata: 記録が無い {e["file"]} {e["r"]}')
    linked = {l['work'] for l in w['links']}
    nolink = [x['id'] for x in w['works'] if x['id'] not in linked and not x.get('unrecorded')]
    return len(w['links']), nolink


def check_hash():
    p = os.path.join(FULL, 'work', 'files_before.sha256')
    if not os.path.exists(p):
        return '控え無し'
    import hashlib
    bad = 0
    for line in open(p, encoding='utf-8'):
        h, name = line.strip().split('  ', 1)
        cur = hashlib.sha256(open(os.path.join(FILES, name), 'rb').read()).hexdigest()
        if cur != h:
            errs.append(f'files/ が変わった：{name}'); bad += 1
    return '一致' if not bad else f'{bad} 件不一致'


def main():
    for jp in sorted(glob.glob(os.path.join(ANS, 'files', '*.json'))):
        a = json.load(open(jp, encoding='utf-8'))
        p = os.path.join(FILES, a['file'])
        if not os.path.exists(p):
            errs.append(f'{a["file"]}: 資料が無い'); continue
        if p.endswith('.xlsx'):
            check_xlsx(p, a['records'])
        elif p.endswith('.pdf'):
            check_pdf(p, a['records'])
    recs_by_file = set()
    for jp in glob.glob(os.path.join(ANS, 'files', '*.json')):
        a = json.load(open(jp, encoding='utf-8'))
        for rc in a['records']:
            recs_by_file.add((a['file'], rc.get('sheet'), rc.get('r'), rc.get('seq') if a['file'].endswith('.pdf') else None))
            if a['plan'] and rc['kind'] in ('item', 'parent'):
                for k in ('path', 'marks', 'plans'):
                    if k not in rc['fields']:
                        errs.append(f'{a["file"]} r{rc["r"]}: {k} が無い')
            md = rc['fields'].get('manDays')
            if md is not None and not isinstance(md, (int, float)):
                errs.append(f'{a["file"]} r{rc.get("r")}: manDays が数でない')
    answered = {json.load(open(j, encoding='utf-8'))['file'] for j in glob.glob(os.path.join(ANS, 'files', '*.json'))}
    for p in glob.glob(os.path.join(FILES, '*')):
        if os.path.basename(p) not in answered:
            errs.append(f'{os.path.basename(p)}: 正解ファイルが無い')
    check_marks()
    nph = check_photos()
    nl, nolink = check_links(recs_by_file)
    hs = check_hash()
    print(f'records {nrec} / links {nl} / 写真 {nph} / 資料に出ない実績 {len(nolink)} / files ハッシュ {hs}')
    if errs:
        print(f'NG {len(errs)}')
        for e in errs[:60]:
            print(' ', e)
        return 1
    print('OK')
    return 0


if __name__ == '__main__':
    sys.exit(main())
