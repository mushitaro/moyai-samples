# -*- coding: utf-8 -*-
"""F01 本作りの検査：
1. 正解（answer/files/*.json）の行がファイルの中身と合う
2. どのファイルにも【架空】がある（ファイル名・中身・文書の情報）
3. 実在の名前に当たらない
4. 写真の大きさ（長辺 1,568〜2,000px、60〜160KB 程度）
python3 check.py  → 問題があれば一覧を出し、終了コード 1
"""
import os, sys, json, glob, csv, re, subprocess, unicodedata
from openpyxl import load_workbook
from PIL import Image

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FILES = BASE + '/files'
ADIR = BASE + '/answer/files'
KAKU = '【架空】'
SKIP_KEYS = {'date', 'start', 'end', 'year', 'manDays', 'kind', 'unit', 'plans', 'days'}
# 実在の会社・自治体・施設で、紛れ込みやすい名前（部分一致で検出）
REAL = ['日立造船', 'カナデビア', 'JFEエンジニアリング', 'タクマ', '荏原環境', '荏原製作所', '川崎重工', 'カワサキ', '三菱重工', '神鋼環境', 'IHI', '住友重機', 'クボタ',
        '新明和', '栗田工業', '日鉄エンジニアリング', '月島', '極東開発', 'プランテック', '三機工業', '東芝', '日立製作所', '富士電機', '横河電機', '明電舎', 'アズビル',
        '東京都', '大阪市', '横浜市', '名古屋市', '札幌市', '福岡市', '清掃工場（東京', '杉並清掃工場', '東京二十三区清掃']

def norm(s):
    s = unicodedata.normalize('NFKC', str(s)).replace('〜', '~').replace('～', '~')
    return re.sub(r'\s+', '', s)

def numeq(a, b):
    try:
        return abs(float(str(a).replace(',', '')) - float(str(b).replace(',', ''))) < 0.051
    except Exception:
        return False

def row_values(ws, r, merged):
    vals = []
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=r, column=c).value
        if v is None and (r, c) in merged:
            v = merged[(r, c)]
        if v is not None:
            vals.append(v)
    return vals

def merged_map(ws):
    m = {}
    for rg in ws.merged_cells.ranges:
        v = ws.cell(row=rg.min_row, column=rg.min_col).value
        for r in range(rg.min_row, rg.max_row + 1):
            for c in range(rg.min_col, rg.max_col + 1):
                m[(r, c)] = v
    return m

def scalar_hits(val, vals_n, vals_raw):
    if isinstance(val, (int, float)) and not isinstance(val, bool):
        return any(numeq(val, x) for x in vals_raw) or norm(val) in ''.join(vals_n)
    s = norm(val)
    if not s:
        return True
    joined = ''.join(vals_n)
    if s in joined:
        return True
    if re.fullmatch(r'[\d.,]+', s):
        return any(numeq(s, x) for x in vals_raw)
    return False

def check_fields(fields, vals_raw, where, errs, ws=None, r=None):
    vals_n = [norm(v) for v in vals_raw]
    for k, v in fields.items():
        if k in SKIP_KEYS or v is None:
            continue
        if k == 'path':
            v = v[-1] if isinstance(v, list) and v else (str(v).split('/')[-1] if v else None)
            if v is None:
                continue
        if k == 'comments' and ws is not None:
            for addr, text in v.items():
                cm = ws[addr].comment
                if not cm or norm(text) not in norm(cm.text):
                    errs.append('%s: コメント %s が合わない' % (where, addr))
            continue
        if isinstance(v, dict):
            for kk, vv in v.items():
                if vv is not None and not scalar_hits(vv, vals_n, vals_raw):
                    errs.append('%s: %s.%s=%r が行にない' % (where, k, kk, vv))
            continue
        if isinstance(v, list):
            continue
        if not scalar_hits(v, vals_n, vals_raw):
            errs.append('%s: %s=%r が行にない' % (where, k, v))

def pdf_text(p):
    return subprocess.run(['pdftotext', p, '-'], capture_output=True, text=True).stdout

def pdf_has(text_n, val):
    s = norm(val)
    if not s or s in text_n:
        return True
    # 表のセル内の折り返しで行が混ざる場合：8 文字ずつに分けてすべてあるか
    parts = [s[i:i + 3] for i in range(0, len(s), 3)]
    return sum(1 for p in parts if p in text_n) >= 0.8 * len(parts)

def main():
    errs = []
    answers = sorted(glob.glob(ADIR + '/*.json'))
    files = sorted(os.listdir(FILES))
    stems_with_answer = set(os.path.splitext(os.path.basename(a))[0] for a in answers)
    for f in files:
        if not f.startswith(KAKU):
            errs.append('ファイル名に【架空】がない: ' + f)
        if os.path.splitext(f)[0] not in stems_with_answer:
            errs.append('正解のないファイル: ' + f)
    wbcache = {}
    nrec = 0
    for a in answers:
        A = json.load(open(a))
        f = A['file']
        p = os.path.join(FILES, f)
        if not os.path.exists(p):
            errs.append('正解に対応するファイルがない: ' + f); continue
        ext = f.rsplit('.', 1)[-1]
        if ext == 'xlsx':
            wb = load_workbook(p)
            if KAKU not in (wb.properties.creator or '') or KAKU not in (wb.properties.title or ''):
                errs.append('文書のプロパティに【架空】がない: ' + f)
            if not any(isinstance(c.value, str) and c.value.startswith(KAKU) for ws in wb.worksheets for row in ws.iter_rows(max_row=6) for c in row):
                errs.append('表題のセルに【架空】がない: ' + f)
            mm = {ws.title: merged_map(ws) for ws in wb.worksheets}
            for i, rc in enumerate(A['records']):
                nrec += 1
                if rc.get('sheet') not in wb.sheetnames:
                    errs.append('%s: シート %s がない' % (f, rc.get('sheet'))); continue
                ws = wb[rc['sheet']]
                vals = row_values(ws, rc['r'], mm[ws.title])
                check_fields(rc['fields'], vals, '%s[%s!%d]' % (f, rc['sheet'], rc['r']), errs, ws, rc['r'])
            alltext = ' '.join(str(c.value) for ws in wb.worksheets for row in ws.iter_rows() for c in row if isinstance(c.value, str))
        elif ext == 'csv':
            raw = open(p, 'rb').read()
            enc = 'utf-8-sig' if raw.startswith(b'\xef\xbb\xbf') else 'cp932'
            rows = list(csv.reader(open(p, encoding=enc, newline='')))
            if KAKU not in ''.join(rows[0]):
                errs.append('CSV の 1 行目に【架空】がない: ' + f)
            for rc in A['records']:
                nrec += 1
                vals = rows[rc['r'] - 1] if rc['r'] - 1 < len(rows) else []
                check_fields(rc['fields'], vals, '%s[r%d]' % (f, rc['r']), errs)
            alltext = raw.decode(enc, errors='ignore')
        elif ext == 'pdf':
            t = pdf_text(p)
            tn = norm(t)
            first = norm(subprocess.run(['pdftotext', '-l', '1', p, '-'], capture_output=True, text=True).stdout)
            info = subprocess.run(['pdfinfo', p], capture_output=True, text=True).stdout
            if KAKU not in first or KAKU not in info:
                errs.append('PDF の 1 頁目か文書の情報に【架空】がない: ' + f)
            last = -1
            if any(rc.get('r') != rc.get('seq', -9) + 1 or rc.get('sheet') for rc in A['records']):
                errs.append('PDF の記録の r/sheet: ' + f)
            for rc in A['records']:
                nrec += 1
                for k, v in rc['fields'].items():
                    if k in SKIP_KEYS or v is None or isinstance(v, (dict, list)):
                        continue
                    if not pdf_has(tn, v):
                        errs.append('%s[seq%d]: %s=%r が本文にない' % (f, rc['seq'], k, str(v)[:30]))
                if rc['seq'] <= last:
                    errs.append('%s: seq の順が乱れている' % f)
                last = rc['seq']
            alltext = t
        elif ext == 'jpg':
            im = Image.open(p)
            long = max(im.size)
            size = os.path.getsize(p)
            if not (1568 <= long <= 2000):
                errs.append('写真の長辺 %dpx: %s' % (long, f))
            if not (60000 <= size <= 160000):
                errs.append('写真の容量 %d bytes: %s' % (size, f))
            com = im.info.get('comment', b'')
            if KAKU.encode('utf-8') not in (com if isinstance(com, bytes) else com.encode('utf-8')):
                errs.append('写真のコメントに【架空】がない: ' + f)
            for rc in A['records']:
                nrec += 1
                if rc.get('image') != f or 'seq' not in rc:
                    errs.append('写真の正解に image/seq がない: ' + f)
            alltext = ''
        else:
            errs.append('想定外の形式: ' + f); continue
        for nm in REAL:
            if nm in alltext:
                errs.append('実在の名前の疑い「%s」: %s' % (nm, f))
    # 正解・world.json・user.md・README にも実在の名前がないか
    for p in [BASE + '/answer/world.json', BASE + '/answer/world.extra.json', BASE + '/user.md', BASE + '/README.md'] + answers[:0]:
        if os.path.exists(p):
            t = open(p).read()
            for nm in REAL:
                if nm in t:
                    errs.append('実在の名前の疑い「%s」: %s' % (nm, p))
        else:
            errs.append('ない: ' + p)
    # world.json（moyai.sample-world/0）
    W = json.load(open(BASE + '/answer/world.json'))
    if W.get('schema') != 'moyai.sample-world/0':
        errs.append('world.json の schema が違う')
    if W.get('facility', {}).get('id') != 'F01' or KAKU not in W.get('facility', {}).get('name', ''):
        errs.append('world.json の facility が違う')
    if W.get('maker', {}).get('today') != '2026-11-13':
        errs.append('world.json の maker.today が違う')
    fset = set(files)
    if set(x['file'] for x in W.get('files', [])) != fset:
        errs.append('world.json の files[] が files/ と合わない')
    wids = set(x['id'] for x in W['works'])
    aids = set(x['id'] for x in W['assets'])
    for x in W['works']:
        if x.get('asset') is not None and x['asset'] not in aids:
            errs.append('作業の機器がない: %s' % x['id'])
        if x.get('attr') not in ('INSPECT', 'OVERHAUL', 'REPLACE', 'UPGRADE'):
            errs.append('作業の attr: %s' % x['id'])
    for L in W['links']:
        if L['file'].endswith('.pdf') and (L.get('r') != L.get('seq', -9) + 1 or L.get('sheet')):
            errs.append('PDF の結びの r/sheet: %s' % L['file'])
        if L.get('how') not in ('direct', 'annual', 'bundle', 'rollup', 'upgrade'):
            errs.append('how の値: %r' % L.get('how'))
        if L['file'] not in fset or L['work'] not in wids:
            errs.append('結びの先がない: %s %s' % (L['file'], L['work']))
    for e in W.get('errata', []):
        if e.get('type') not in ('omitMark', 'lastYearUnreflected') or e['file'] not in fset:
            errs.append('errata の形: %r' % e)
    # files/ が変わっていないか（answer をそろえる前に控えたハッシュ）
    hp = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'files_sha256_before_v2.txt')
    if os.path.exists(hp):
        import hashlib
        want = dict((l.split('  ', 1)[1].strip(), l.split('  ', 1)[0]) for l in open(hp, encoding='utf-8') if l.strip())
        for fn in files:
            if hashlib.sha256(open(os.path.join(FILES, fn), 'rb').read()).hexdigest() != want.get(fn):
                errs.append('files/ が変わった: ' + fn)
        if set(want) != fset:
            errs.append('files/ のファイルの数が変わった')
    print('ファイル %d、正解 %d、記録 %d 行、world: 機器 %d 作業 %d 結び %d' % (len(files), len(answers), nrec, len(aids), len(wids), len(W['links'])))
    if errs:
        print('問題 %d 件' % len(errs))
        for e in (errs if os.environ.get("ALL") else errs[:200]):
            print(' -', e)
        sys.exit(1)
    print('OK：問題なし')

if __name__ == '__main__':
    main()
