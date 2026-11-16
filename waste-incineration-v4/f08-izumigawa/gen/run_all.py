# -*- coding: utf-8 -*-
"""F08 本作り：python3 -I run_all.py で files/・answer/・user.md・README.md を全部作り直す。"""
import glob
import os
import shutil
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import ANS, OUT  # noqa: E402


def main():
    t0 = time.time()
    for d in (OUT, os.path.join(ANS, "files")):
        shutil.rmtree(d, ignore_errors=True)
        os.makedirs(d)
    import alljobs
    import w_daily
    import w_ledger
    import w_orders
    import w_plan
    import w_report
    import w_sched
    js = alljobs.jobs()
    n = 0
    for j in js:              # 工程表が先、日報はその工程から、報告書は日報から
        w_sched.write(j)
        w_daily.write(j)
        w_report.write(j)
        n += 3
    for v in (2009, 2012, 2015):
        w_plan.old_plan(v)
    for v in (2016, 2019, 2022):
        w_plan.new_plan(v)
    for y in (2005, 2010, 2015, 2017, 2021, 2024):
        w_ledger.ledger(y)
    w_orders.write_all(js)
    import answer
    answer.write(js)
    print("files", len(glob.glob(os.path.join(OUT, "*"))), "sec", round(time.time() - t0, 1))


if __name__ == "__main__":
    main()
