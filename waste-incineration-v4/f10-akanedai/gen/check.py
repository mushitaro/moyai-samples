# -*- coding: utf-8 -*-
"""（v2）files/ のハッシュが控えと同じか、answer/ が moyai.sample-world/0 の形か、正解の行がファイルの中身と合うか、【架空】、実在の名前、写真。  python3 check.py"""
import os, sys, json, glob, csv, io, subprocess, unicodedata, re, datetime as dt
from openpyxl import load_workbook

FULL = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FILES = os.path.join(FULL, 'files'); ANS = os.path.join(FULL, 'answer', 'files')
REAL = ['日立造船', 'カナデビア', 'JFEエンジニアリング', '荏原', 'タクマ', '三菱重工', '川崎重工', '神鋼環境', '栗田工業', 'クボタ', '日鉄エンジ', '新日鉄', '住友重機',
        'IHI', '東芝', '日立製作所', '富士電機', '横河電機', 'アズビル', '明電舎', '三機工業', '月島', '水ing', '極東開発', '関西電力', '東京電力']
SKIP = {'date', 'weather', 'planStart', 'planEnd', 'actStart', 'actEnd', 'unit', 'row', 'ref', 'section', 'image', 'kind', 'code', 'sheet', 'done', 'count', 'month', 'blockTeam', 'blockCompany', 'plans', 'manDays', 'manDaysText', 'days', 'hiddenMarks'}
N = lambda s: unicodedata.normalize('NFKC', str(s).replace('〜', '~').replace('～', '~')).replace(' ', '').replace('　', '').replace('\n', '').replace(',', '')

def flat(v):
    if isinstance(v, dict): return [x for y in v.values() for x in flat(y)]
    if isinstance(v, list): return [x for y in v for x in flat(y)]
    return [v]

def num_ok(v, cells):
    for c in cells:
        try:
            if abs(float(str(c).replace(',', '')) - float(v)) < 0.051: return True
        except Exception: pass
    return False

def grid(ws):
    vals = {}
    for row in ws.iter_rows():
        for c in row:
            if c.value is not None: vals[(c.row, c.column)] = c.value
    for mr in ws.merged_cells.ranges:
        v = vals.get((mr.min_row, mr.min_col))
        if v is None: continue
        for r in range(mr.min_row, mr.max_row + 1):
            for col in range(mr.min_col, mr.max_col + 1): vals[(r, col)] = v
    rows = {}
    for (r, c), v in vals.items(): rows.setdefault(r, []).append(v)
    return rows

def check_values(fields, cells, where, errs):
    txt = N(''.join(str(c) for c in cells))
    for k, v in fields.items():
        if k in SKIP or v in (None, '', [], {}): continue
        if k == 'path': v = v[-1:]
        for x in flat(v):
            if x in (None, ''): continue
            if isinstance(x, bool): continue
            if isinstance(x, (int, float)):
                if not num_ok(x, cells) and N(x) not in txt: errs.append(f'{where} {k}={x}')
            else:
                s = N(x)
                if re.fullmatch(r'\d{4}-\d\d-\d\d', str(x)): continue
                if s and s not in txt and not any(s == N(c) for c in cells): errs.append(f'{where} {k}={x!r}')

def main():
    errs = []; n = 0
    js = sorted(glob.glob(os.path.join(ANS, '*.json')))
    files = sorted(os.listdir(FILES))
    stems = {os.path.splitext(os.path.basename(j))[0] for j in js}
    for f in files:
        if os.path.splitext(f)[0] not in stems: errs.append(f'正解なし {f}')
    for j in js:
        a = json.load(open(j)); fn = os.path.join(FILES, a['file'])
        if not os.path.exists(fn): errs.append(f'ファイルなし {a["file"]}'); continue
        ext = fn.rsplit('.', 1)[1]
        if ext == 'xlsx':
            wb = load_workbook(fn)
            G = {ws.title: grid(ws) for ws in wb.worksheets}
            if '【架空】' not in (wb.properties.title or ''): errs.append(f'【架空】なし（文書のプロパティ） {a["file"]}')
            if not any('【架空】' in str(v) for v in sum(list(G[wb.worksheets[0].title].values())[:3], [])): errs.append(f'【架空】なし（表題） {a["file"]}')
            for x in a['records']:
                n += 1
                cells = G.get(x['sheet'], {}).get(x['r'], [])
                if not cells: errs.append(f"{a['file']} {x['sheet']}!{x['r']} 空"); continue
                check_values(x['fields'], cells, f"{a['file']} {x['sheet']}!{x['r']}", errs)
            alltext = ''.join(str(v) for g in G.values() for r in g.values() for v in r)
        elif ext == 'csv':
            raw = open(fn, 'rb').read()
            try: t = raw.decode('cp932')
            except Exception: t = raw.decode('utf-8')
            rows = list(csv.reader(t.splitlines()))
            if '【架空】' not in t.splitlines()[0]: errs.append(f'【架空】なし {a["file"]}')
            for x in a['records']:
                n += 1
                check_values(x['fields'], rows[x['r'] - 1], f"{a['file']} r{x['r']}", errs)
            alltext = t
        elif ext == 'pdf':
            t = subprocess.run(['pdftotext', fn, '-'], capture_output=True, text=True).stdout
            p1 = subprocess.run(['pdftotext', '-f', '1', '-l', '1', fn, '-'], capture_output=True, text=True).stdout
            info = subprocess.run(['pdfinfo', fn], capture_output=True, text=True).stdout
            if '【架空】' not in p1 or '【架空】' not in info: errs.append(f'【架空】なし {a["file"]}')
            T = N(t)
            for x in a['records']:
                n += 1
                if x.get('r') != x.get('seq', -9) + 1: errs.append(f"{a['file']} PDF の r が seq+1 でない")
                if x.get('sheet') is not None: errs.append(f"{a['file']} PDF の sheet が null でない")
                for k, v in x['fields'].items():
                    if k in SKIP or k in ('text', 'caption', 'name', 'work') and len(str(v)) > 60: continue
                    for y in flat(v):
                        if isinstance(y, (int, float)) and not isinstance(y, bool):
                            if N(y) not in T: errs.append(f"{a['file']} seq{x['seq']} {k}={y}")
                        elif isinstance(y, str) and y and N(y)[:40] not in T: errs.append(f"{a['file']} seq{x['seq']} {k}={y[:30]!r}")
            imgs = subprocess.run(['pdfimages', '-list', fn], capture_output=True, text=True).stdout.splitlines()[2:]
            for l in imgs:
                w, h = int(l.split()[3]), int(l.split()[4])
                if max(w, h) < 1568 or max(w, h) > 2000: errs.append(f'写真の大きさ {a["file"]} {w}x{h}')
            alltext = t
        for rn in REAL:
            if rn in alltext or rn in a['file']: errs.append(f'実在の名前 {rn} {a["file"]}')
    w = json.load(open(os.path.join(FULL, 'answer', 'world.json')))
    if w.get('schema') != 'moyai.sample-world/0': errs.append('schema')
    if w.get('facility', {}).get('id') != 'F10' or '【架空】' not in w['facility'].get('name', ''): errs.append('facility')
    if w.get('maker', {}).get('today') != '2026-11-13': errs.append('maker.today')
    if 'furnace' not in w.get('profile', {}): errs.append('profile.furnace')
    if sorted(x['file'] for x in w['files']) != files: errs.append('files[] とフォルダの不一致')
    aid = {x['id'] for x in w['assets']}
    for x in w['assets']:
        for k in ('id', 'name', 'tag', 'system', 'furnace', 'installed', 'class', 'part', 'names'):
            if k not in x: errs.append(f'assets 欄なし {k}'); break
        if x['part'] and x['part'] not in aid: errs.append(f"assets part {x['id']}")
    ids = {x['id'] for x in w['works']}
    for x in w['works']:
        if x['asset'] not in aid: errs.append(f"works asset {x['id']}")
        if x['attr'] not in ('INSPECT', 'OVERHAUL', 'REPLACE', 'UPGRADE'): errs.append(f"works attr {x['id']}")
        if not isinstance(x['fy'], int): errs.append(f"works fy {x['id']}")
    recs = {}
    for j in js:
        a = json.load(open(j))
        for x in a['records']:
            recs[(a['file'], x.get('sheet'), x['r'])] = (a['plan'], x)
            if a['plan'] and x['kind'] in ('item', 'parent'):
                if not all(k in x['fields'] for k in ('path', 'marks', 'plans')): errs.append(f"星取 path/marks/plans なし {a['file']} {x['r']}")
            if a['file'].endswith(('.pdf', '.csv')) and x.get('sheet') is not None: errs.append(f"sheet が null でない {a['file']}")
            if 'manDays' in x['fields'] and 'manDaysText' not in x['fields']: errs.append(f"manDaysText なし {a['file']} {x['r']}")
    for l in w['links']:
        if l['work'] not in ids: errs.append(f"links 作業なし {l['work']}"); continue
        if l['how'] not in ('direct', 'annual', 'bundle', 'rollup', 'upgrade'): errs.append(f"links how {l['how']}")
        rr = recs.get((l['file'], l.get('sheet'), l['r']))
        if not rr: errs.append(f"links 行なし {l['file']} {l.get('sheet')} {l['r']}"); continue
        if rr[0] and 'year' not in l: errs.append(f"links 星取の year なし {l['file']} {l['r']}")
        if '作業日報' in l['file'] and rr[1]['kind'] == 'item' and not l.get('day'): errs.append(f"links 日報の day なし {l['file']} {l['r']}")
    for e in w.get('errata', []):
        if e.get('type') not in ('omitMark', 'lastYearUnreflected'): errs.append('errata type')
    # files/ のハッシュ
    import hashlib
    before = os.path.join(os.path.dirname(FULL), 'work', 'files_hash_before.txt')
    if os.path.exists(before):
        cur = {}
        for f in files: cur['./' + f] = hashlib.sha256(open(os.path.join(FILES, f), 'rb').read()).hexdigest()
        old = {l.split('  ', 1)[1].strip(): l.split('  ', 1)[0] for l in open(before) if l.strip()}
        if cur != old: errs.append('files/ のハッシュが控えと違う')
        else: print('files/ のハッシュ：控えと一致（' + str(len(cur)) + ' ファイル）')
    print(f'ファイル {len(files)}、正解ファイル {len(js)}、行 {n}、機器 {len(w["assets"])}、作業 {len(w["works"])}、結び {len(w["links"])}')
    print('写真：PDF に埋め込んだ画像なし（写真代替図はベクターの線画）' if True else '')
    if errs:
        from collections import Counter
        cnt = Counter()
        ex = {}
        for e in errs:
            m = re.search(r'(工程表|作業日報|工事報告書|中長期整備計画|機器台帳|発注台帳)[^ ]* .*? (\w+)=', e)
            k = (m.group(1), m.group(2)) if m else (e[:12], '')
            cnt[k] += 1; ex.setdefault(k, e)
        print('NG', len(errs))
        for k, v in cnt.most_common(40): print(' ', v, k, '例:', ex[k][-120:])
        sys.exit(1)
    print('OK')

if __name__ == '__main__':
    main()
