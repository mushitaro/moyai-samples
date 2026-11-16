# -*- coding: utf-8 -*-
import subprocess, os, shutil, glob
PROF = 'file:///tmp/lo_F10'
TMP = os.path.join(os.path.dirname(os.path.abspath(__file__)), '_work')
def to_pdf(path):
    os.makedirs(TMP, exist_ok=True)
    if path.lower().endswith('.pdf'): return path
    if path.lower().endswith('.csv'):
        # 表計算ソフトで開いた見た目として確認する（中身はそのまま、列幅と印刷の向きだけ設定）
        import csv as _csv
        from openpyxl import Workbook
        from xl import font, page
        raw = open(path, 'rb').read()
        for enc in ('cp932', 'utf-8'):
            try: txt = raw.decode(enc); break
            except Exception: pass
        rows = list(_csv.reader(txt.splitlines()))
        wb = Workbook(); ws = wb.active; ws.title = os.path.basename(path)[:28]
        wmax = {}
        for r, row in enumerate(rows, 1):
            for c, v in enumerate(row, 1):
                x = ws.cell(r, c, v); x.font = font(8)
                if r > 1: wmax[c] = max(wmax.get(c, 4), min(40, len(v) * 1.6 + 1))
        from openpyxl.utils import get_column_letter as L
        for c, w in wmax.items(): ws.column_dimensions[L(c)].width = w
        page(ws, land=True, paper=8, fitw=1, fith=0)
        tmpx = os.path.join(TMP, os.path.splitext(os.path.basename(path))[0] + '.xlsx')
        wb.save(tmpx)
        path = tmpx
    subprocess.run(['soffice', f'-env:UserInstallation={PROF}', '--headless', '--convert-to', 'pdf', '--outdir', TMP, path],
                   check=True, capture_output=True, timeout=600)
    return os.path.join(TMP, os.path.splitext(os.path.basename(path))[0] + '.pdf')
def npages(pdf):
    out = subprocess.run(['pdfinfo', pdf], capture_output=True, text=True).stdout
    return int([l for l in out.splitlines() if l.startswith('Pages:')][0].split()[1])
def png(path, pages, dpi=80):
    """pages: ページ番号のリスト。出力は path 名＋.png（複数なら _p1..）"""
    pdf = to_pdf(path)
    n = npages(pdf)
    pages = [p for p in pages if p <= n]
    outs = []
    base = os.path.basename(path)
    d = os.path.dirname(path)
    for i, p in enumerate(pages, 1):
        name = base + '.png' if len(pages) == 1 else f'{base}_p{i}.png'
        pre = os.path.join(TMP, 'pp')
        subprocess.run(['pdftoppm', '-f', str(p), '-l', str(p), '-r', str(dpi), '-png', '-singlefile', pdf, pre], check=True)
        shutil.move(pre + '.png', os.path.join(d, name))
        outs.append(name)
    return outs, n
