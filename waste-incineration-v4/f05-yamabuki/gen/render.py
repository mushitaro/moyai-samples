# -*- coding: utf-8 -*-
"""LibreOffice（自分専用プロファイル /tmp/lo_F05）で PDF → pdftoppm で PNG"""
import os, subprocess, shutil, sys, glob
from common import OUT, WORK, LO_PROFILE


def to_pdf(path, outdir=WORK):
    subprocess.run(['soffice', f'-env:UserInstallation={LO_PROFILE}', '--headless', '--convert-to', 'pdf',
                    '--outdir', outdir, path], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=240)
    return os.path.join(outdir, os.path.splitext(os.path.basename(path))[0] + '.pdf')


def npages(pdf):
    out = subprocess.run(['pdfinfo', pdf], capture_output=True, text=True).stdout
    for line in out.splitlines():
        if line.startswith('Pages:'):
            return int(line.split()[1])
    return 0


def png(path, pages=(1,), dpi=80, pdf=None):
    """pages の各ページを PNG に。1 ページなら 同名.png、複数なら _p1 …（p番号は選んだ順）"""
    pdf = pdf or to_pdf(path)
    os.makedirs(os.path.join(WORK, 'png'), exist_ok=True)
    base = os.path.join(WORK, 'png', os.path.basename(path))
    res = []
    n = npages(pdf)
    pages = [p for p in pages if p <= n]
    for i, p in enumerate(pages, 1):
        tmp = os.path.join(WORK, '_pg')
        subprocess.run(['pdftoppm', '-f', str(p), '-l', str(p), '-r', str(dpi), '-png', '-singlefile', pdf, tmp], check=True)
        dst = base + '.png' if len(pages) == 1 else f'{base}_p{i}.png'
        shutil.move(tmp + '.png', dst)
        res.append(os.path.basename(dst))
    return res, n


if __name__ == '__main__':
    for a in sys.argv[1:]:
        print(png(a))
