# -*- coding: utf-8 -*-
import os, subprocess, shutil, glob, tempfile
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.utils import get_column_letter as CL

FULL = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = FULL + '/files'
TMP = FULL + '/gen/_render'

thin = Side(style='thin', color='000000')
hair = Side(style='hair', color='000000')
med = Side(style='medium', color='000000')
dot = Side(style='dotted', color='000000')
dbl = Side(style='double', color='000000')
NONE = Side(style=None)

def box(l=thin, r=thin, t=thin, b=thin):
    return Border(left=l, right=r, top=t, bottom=b)
B_ALL = box()
def fill(c):
    return PatternFill('solid', fgColor=c)
def pfill(pat, fg, bg='FFFFFF'):
    return PatternFill(pat, fgColor=fg, bgColor=bg)
AC = Alignment(horizontal='center', vertical='center', wrap_text=True)
AL = Alignment(horizontal='left', vertical='center', wrap_text=True)
ALT = Alignment(horizontal='left', vertical='top', wrap_text=True)
AR = Alignment(horizontal='right', vertical='center')
ACV = Alignment(horizontal='center', vertical='center', text_rotation=255, wrap_text=True)

def F(size=10, bold=False, name='ＭＳ Ｐゴシック', color='000000', u=None, strike=False):
    return Font(name=name, size=size, bold=bold, color=color, underline=u, strike=strike)

def put(ws, r, c, v, font=None, al=None, border=None, fl=None, fmt=None):
    cell = ws.cell(row=r, column=c, value=v)
    if font: cell.font = font
    if al: cell.alignment = al
    if border: cell.border = border
    if fl: cell.fill = fl
    if fmt: cell.number_format = fmt
    return cell

def merge(ws, r1, c1, r2, c2, v=None, font=None, al=None, border=B_ALL, fl=None):
    if (r1, c1) != (r2, c2):
        ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
    cell = ws.cell(row=r1, column=c1, value=v)
    if font: cell.font = font
    if al: cell.alignment = al
    if fl: cell.fill = fl
    if border:
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                ws.cell(row=r, column=c).border = border
    return cell

def grid(ws, r1, c1, r2, c2, border=B_ALL):
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            ws.cell(row=r, column=c).border = border

def outline(ws, r1, c1, r2, c2, side=med):
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = ws.cell(row=r, column=c)
            b = cell.border
            cell.border = Border(left=side if c == c1 else b.left, right=side if c == c2 else b.right,
                                 top=side if r == r1 else b.top, bottom=side if r == r2 else b.bottom)

def page(ws, landscape=True, paper=9, fitw=1, fith=0, area=None, margins=(0.4, 0.4, 0.5, 0.5), title_rows=None, center=True):
    ws.page_setup.orientation = 'landscape' if landscape else 'portrait'
    ws.page_setup.paperSize = paper  # 9=A4, 8=A3
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth = fitw
    ws.page_setup.fitToHeight = fith
    ws.page_margins.left, ws.page_margins.right, ws.page_margins.top, ws.page_margins.bottom = margins
    ws.print_options.horizontalCentered = center
    if area:
        ws.print_area = area
    if title_rows:
        ws.print_title_rows = title_rows

def widths(ws, d):
    for k, v in d.items():
        ws.column_dimensions[k if isinstance(k, str) else CL(k)].width = v

def render(path, pages=(1,), dpi=70, outnames=None):
    """LibreOffice で PDF → pdftoppm で PNG。out/ に 同名.png（複数ページは _pN）。"""
    os.makedirs(TMP, exist_ok=True)
    for f in glob.glob(TMP + '/*'):
        os.remove(f)
    src = path
    subprocess.run(['/usr/bin/soffice', '-env:UserInstallation=file:///tmp/lo_F01', '--headless', '--convert-to', 'pdf', '--outdir', TMP, src],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=300)
    pdf = glob.glob(TMP + '/*.pdf')[0]
    info = subprocess.run(['pdfinfo', pdf], capture_output=True, text=True).stdout
    npages = int([l for l in info.splitlines() if l.startswith('Pages')][0].split()[1])
    res = []
    base = os.path.splitext(os.path.basename(path))[0]
    for i, p in enumerate(pages):
        if p > npages:
            continue
        subprocess.run(['pdftoppm', '-png', '-r', str(dpi), '-f', str(p), '-l', str(p), pdf, TMP + '/pg'], check=True)
        g = sorted(glob.glob(TMP + '/pg*.png'))[-1]
        name = (outnames[i] if outnames else (base + '.png' if len(pages) == 1 else base + '_p%d.png' % (i + 1)))
        shutil.move(g, os.path.join(OUT, name))
        res.append(name)
    return res, npages

# ------------------------------------------------------------------ 保存時の共通処理
import re as _re
import openpyxl as _opx
from openpyxl.utils import column_index_from_string as _ci, range_boundaries as _rb

KAKU = '【架空】'

def _creator(fn):
    b = os.path.basename(fn)
    for key, who in (('a2様式', 'あさぎ市 環境部清掃工場 施設係 a2'), ('a3様式', 'あさぎ市 環境部清掃工場 施設係 a3'),
                     ('m1様式', '汐見重工環境 現場代理人 m2（m1様式）'), ('m2様式', '汐見重工環境 現場代理人'), ('m3様式', '汐見重工環境 現場代理人 m3')):
        if key in b:
            if key == 'm2様式' and '令和7' in b:
                who = '汐見重工環境 現場代理人 m3（m2様式）'
            elif key == 'm2様式':
                who = '汐見重工環境 現場代理人 m2'
            return KAKU + who
    return KAKU + '試験用サンプル'

def resolve_formulas(wb):
    """openpyxl の数式は値を持たないので、保存前に計算済みの数値へ置き換える（小計・合計など）。"""
    memo = {}
    def val(ws, coord):
        key = (ws.title, coord)
        if key in memo:
            return memo[key]
        v = ws[coord].value
        if isinstance(v, str) and v.startswith('='):
            v = ev(ws, v[1:])
            memo[key] = v
            return v
        return v
    def num(x):
        return x if isinstance(x, (int, float)) and not isinstance(x, bool) else 0
    def rng_vals(ws, a, b):
        c1, r1, c2, r2 = _rb('%s:%s' % (a, b))
        out = []
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                out.append(val(ws, '%s%d' % (CL(c), r)))
        return out
    def ev(ws, f):
        def sheetref(m):
            return '_V(%r,%r)' % (m.group(1), m.group(2))
        f = _re.sub(r"'([^']+)'!(\$?[A-Z]{1,3}\$?\d+)", sheetref, f)
        f = _re.sub(r'SUM\((\$?[A-Z]{1,3}\$?\d+):(\$?[A-Z]{1,3}\$?\d+)\)', lambda m: '_S(%r,%r)' % (m.group(1), m.group(2)), f)
        f = _re.sub(r'COUNTA\((\$?[A-Z]{1,3}\$?\d+):(\$?[A-Z]{1,3}\$?\d+)\)', lambda m: '_C(%r,%r)' % (m.group(1), m.group(2)), f)
        f = _re.sub(r"(?<!['A-Za-z_])(\$?[A-Z]{1,3}\$?\d+)(?![\w'])", lambda m: '_N(%r)' % m.group(1), f)
        env = {'_V': lambda s, c: num(val(wb[s], c)), '_S': lambda a, b: sum(num(x) for x in rng_vals(ws, a, b)),
               '_C': lambda a, b: sum(1 for x in rng_vals(ws, a, b) if x not in (None, '')), '_N': lambda c: num(val(ws, c))}
        r = eval(f, {'__builtins__': {}}, env)
        return int(r) if isinstance(r, float) and abs(r - round(r)) < 1e-9 else (round(r, 1) if isinstance(r, float) else r)
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith('='):
                    cell.value = val(ws, cell.coordinate)

TITLE_KEYS = ('あさぎ市清掃工場', '日　報', '作業日報', '別表', '年度別事業費')
def _title_prefix(wb):
    for ws in wb.worksheets:
        first = True
        for row in ws.iter_rows(min_row=1, max_row=6):
            for cell in row:
                v = cell.value
                if not (isinstance(v, str) and any(k in v for k in TITLE_KEYS)):
                    continue
                big = cell.font is not None and cell.font.sz is not None and cell.font.sz >= 12
                if v.startswith(KAKU):
                    first = False
                    continue
                if first or big:
                    cell.value = KAKU + v
                    first = False

_orig_save = _opx.Workbook.save
def _save(self, filename):
    resolve_formulas(self)
    if '_render' not in str(filename) and not str(filename).startswith(TMP):
        _title_prefix(self)
    import datetime as _d
    self.properties.created = _d.datetime(2026, 11, 13, 9, 0, 0)
    self.properties.modified = _d.datetime(2026, 11, 13, 9, 0, 0)
    self.properties.creator = _creator(str(filename))
    self.properties.lastModifiedBy = self.properties.creator
    self.properties.title = KAKU + os.path.splitext(os.path.basename(str(filename)))[0].replace(KAKU, '')
    self.properties.description = KAKU + '試験用サンプル（あさぎ市清掃工場 F01）。実在の施設・会社・人とは関係ありません。'
    return _orig_save(self, filename)
_opx.Workbook.save = _save
from openpyxl.workbook import workbook as _wbm
_wbm.Workbook.save = _save

import zipfile as _zf
def normalize_zip(path):
    """再実行で同じバイト列になるよう、zip 内の時刻と core.xml の更新日時を固定する。"""
    with _zf.ZipFile(path) as z:
        items = [(n, z.read(n)) for n in z.namelist()]
    with _zf.ZipFile(path, 'w', _zf.ZIP_DEFLATED) as z:
        for n, b in items:
            if n == 'docProps/core.xml':
                b = _re.sub(rb'(<dcterms:modified[^>]*>)[^<]*(</dcterms:modified>)', rb'\g<1>2026-11-13T09:00:00Z\g<2>', b)
                b = _re.sub(rb'(<dcterms:created[^>]*>)[^<]*(</dcterms:created>)', rb'\g<1>2026-11-13T09:00:00Z\g<2>', b)
            zi = _zf.ZipInfo(n, date_time=(2026, 11, 13, 9, 0, 0))
            zi.compress_type = _zf.ZIP_DEFLATED
            z.writestr(zi, b)

def _save2(self, filename):
    r = _save(self, filename)
    normalize_zip(filename)
    return r
_opx.Workbook.save = _save2
_wbm.Workbook.save = _save2
