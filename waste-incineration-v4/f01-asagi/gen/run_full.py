# -*- coding: utf-8 -*-
"""F01 本作り：1997〜2026 年度の全資料と正解を作る。  python3 run_full.py"""
import os, glob, sys, time, datetime as dt
import ans
for d_ in (ans.FILES, ans.ADIR):
    os.makedirs(d_, exist_ok=True)
    for f in glob.glob(d_ + '/*'):
        os.remove(f)
import world2 as W, build_sd as B, build_rep as BR, build_ord as O, build_plan as PL, make_ledger as LG
from f01_master import TODAY
t0 = time.time()
years = [int(a) for a in sys.argv[1:]] or W.YEARS
for fy in years:
    B.sched_teishu(fy)
    B.daily_teishu(fy)
    if fy < 2026:
        BR.teishu_report(fy)
    O.build(fy)
    for c, L in W.hoshu_all(fy):
        B.contract_docs(fy, c, L, 'H')
        if c['e_'] <= TODAY:
            BR.contract_report(fy, c, L, 'H')
    for c in W.special_contracts(fy):
        L = W.special_tasks(fy, c)
        B.contract_docs(fy, c, L, c['kind'])
        BR.contract_report(fy, c, L, c['kind'])
    print(fy, 'ok', round(time.time() - t0))
if not sys.argv[1:]:
    PL.build_a1(); PL.build_a2(2013); PL.build_a2(2019)
    for y in (2021, 2023, 2026):
        PL.build_a3(y)
    LG.build_2008(); LG.build_2019()
print('xlsx/csv done', round(time.time() - t0))
B.process_photos()
print('photos done', round(time.time() - t0))
BR.to_pdfs()
print('pdf done', round(time.time() - t0))
ans.write_all()
import export_world
export_world.main()
print('all done', round(time.time() - t0))
