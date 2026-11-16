# -*- coding: utf-8 -*-
"""【架空】かざはやエコプラザ（F07）本作り：資料・正解を全部作り直す。
  cd f07/full/gen && python3 run_all.py   （約 3 分。LibreOffice は file:///tmp/lo_F07）
順序：工程表 → 日報（工程から）→ 報告書 → 発注台帳 → 機器台帳 → 中長期整備計画 → 正解（answer/）→ check.py"""
import glob
import json
import os
import shutil
import subprocess
import sys
import time

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from common import *  # noqa


def clean():
    for d in (OUT, ANS, WORK):
        os.makedirs(d, exist_ok=True)
        for f in glob.glob(os.path.join(d, "*")):
            if os.path.isfile(f):
                os.remove(f)


def main():
    t0 = time.time()
    clean()
    import gen_schedule_full, gen_daily_full, gen_report_full, gen_orders_full, gen_ledger_full, gen_plan_full
    s = gen_schedule_full.run(); print("工程表", len(s), round(time.time() - t0))
    d = gen_daily_full.run(); print("日報", len(d), round(time.time() - t0))
    r = gen_report_full.run(); print("報告書", len(r), round(time.time() - t0))
    o = gen_orders_full.run(); print("発注", len(o), round(time.time() - t0))
    l = gen_ledger_full.run(); print("台帳", len(l), round(time.time() - t0))
    p = gen_plan_full.run(); print("計画", len(p), round(time.time() - t0))
    save_answers()
    import world_json
    world_json.write(gen_schedule_full.LINKS + gen_daily_full.LINKS + gen_report_full.LINKS + gen_orders_full.LINKS,
                     gen_plan_full.ACTLINK)
    import gen_docs
    gen_docs.write()
    print("正解・README・user.md", round(time.time() - t0))
    subprocess.run([sys.executable, os.path.join(HERE, "check.py")], check=False)


if __name__ == "__main__":
    main()
