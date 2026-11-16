# -*- coding: utf-8 -*-
"""機器台帳 2 版（見本の様式のまま）と正解の行。"""
import os, re
from openpyxl import load_workbook
from common import *
from world import MACHINES, old_machines
import gen_ledger as GL
import rec


def build_all():
    out = []
    f1 = f'{FN}_機器台帳_令和3年度旧版_組合1シート.xlsx'
    f2 = f'{FN}_機器台帳_令和5年度刷新_メーカー機器表様式.xlsx'
    GL.build_old(os.path.join(OUT, f1))
    GL.build_new(os.path.join(OUT, f2))
    # 旧版
    stem = rec.start(f1)
    wb = load_workbook(os.path.join(OUT, f1))
    ws = wb.active
    om = old_machines()
    hdr = [ws.cell(4, c).value for c in range(1, 11)]
    prevk = None
    for r in range(5, ws.max_row + 1):
        no = ws.cell(r, 1).value
        if isinstance(no, int):
            m = om[no - 1]
            vals = {h: ws.cell(r, c).value for c, h in enumerate(hdr, 1) if ws.cell(r, c).value not in (None, '')}
            if vals.get('区分') == '〃':
                vals['区分（値）'] = prevk
            else:
                prevk = vals.get('区分')
            vals['code'] = m.get('tag')
            rec.add(stem, ws.title, r, 'item', vals)
        elif ws.cell(r, 3).value and str(ws.cell(r, 3).value).startswith('計'):
            rec.add(stem, ws.title, r, 'total', dict(計=ws.cell(r, 3).value))
    out.append(f1)
    # 刷新版
    stem = rec.start(f2)
    wb = load_workbook(os.path.join(OUT, f2))
    cv = wb['表紙・集計']
    for r in range(9, 30):
        v = cv.cell(r, 2).value
        if v and v != 'シート':
            kind = 'total' if v == '合計' else 'subtotal'
            rec.add(stem, cv.title, r, kind, dict(シート=v, **{'1号炉': cv.cell(r, 3).value, '2号炉': cv.cell(r, 4).value, '共通': cv.cell(r, 5).value, '計': cv.cell(r, 6).value,
                                                             '備考': cv.cell(r, 7).value}))
            if v == '合計':
                break
    for ws in wb.worksheets[1:]:
        hdr = [ws.cell(3, c).value for c in range(1, 14)]
        for r in range(4, ws.max_row + 1):
            a = ws.cell(r, 1).value
            if a and re.match(r'^[0-2][A-Z]-\d\d$', str(a)):
                vals = {h.replace('\n', ''): ws.cell(r, c).value for c, h in enumerate(hdr, 1) if ws.cell(r, c).value not in (None, '')}
                vals['code'] = a
                rec.add(stem, ws.title, r, 'item', vals)
            elif a and ws.cell(r, 2).value is None and '　' in str(a):
                rec.add(stem, ws.title, r, 'parent', dict(系統=a))
    out.append(f2)
    return out


if __name__ == '__main__':
    print(build_all())
