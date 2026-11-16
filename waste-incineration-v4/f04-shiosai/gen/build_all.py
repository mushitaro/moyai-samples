# -*- coding: utf-8 -*-
"""F04 本作り：全資料・正解・説明を作り直す。 python3 build_all.py
 LibreOffice は /tmp/lo_F04 の専用プロファイル（render.py）。"""
import os, sys, shutil, glob, json, time
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xl import OUT, BUILD
import world as W, cw, sg
import gen_docs as GD, gen_pdf as GP, gen_kanri as GK, gen_plan as PL, gen_ledger as LG
import answer

def main(skip_pdf=False):
    t0 = time.time()
    shutil.rmtree(OUT, ignore_errors=True); os.makedirs(OUT)
    shutil.rmtree(BUILD, ignore_errors=True); os.makedirs(BUILD)
    shutil.rmtree(answer.ANS, ignore_errors=True)
    made = []
    def add(p): made.append(os.path.basename(p))
    # 計画・台帳
    add(PL.build_old(original=True)); add(PL.build_old()); add(PL.build_new())
    add(LG.ledger_2015()); add(LG.ledger_2024())
    # 管理表
    for p in GK.build_all(): add(p)
    # 契約ごと
    for c in cw.CONTRACTS.values():
        if c.kind == "MK":
            s = sg.SCHED[c.cid]
            add(GD.koutei_v1_old(c, s) if c.fy <= 2014 else GD.koutei_v1(c, s))
            if c.nippo: add(GD.nippo_v1(c, s))
            if c.report and not skip_pdf: add(GP.mk_old(c, s) if c.fy <= 2015 else GP.mk_new(c, s))
        elif c.kind == "EN":
            for fy in (2015, 2016):
                s = sg.SCHED["EN%d" % fy]
                add(GD.koutei_v1(c, s, fy)); add(GD.nippo_v1(c, s, fy))
                if not skip_pdf: add(GP.mk_old(c, s, fy) if fy == 2015 else GP.mk_new(c, s, fy))
        elif c.kind == "CRUP":
            s = sg.SCHED[c.cid]
            add(GD.koutei_v1(c, s)); add(GD.nippo_v1(c, s))
            if not skip_pdf: add(GP.mk_new(c, s))
        elif c.kind == "LOCAL" and c.docs:
            s = sg.SCHED[c.cid]
            add(GD.koutei_g1(c, s) if c.fy <= 2018 else GD.koutei_g2(c, s))
            add(GD.nippo_g1(c, s) if c.fy <= 2019 else GD.nippo_g2(c, s))
            if not skip_pdf: add(GP.local_report(c, s))
    answer.write_all(made)
    print("files", len(made), "%.0fs" % (time.time() - t0))

if __name__ == "__main__":
    main(skip_pdf="--nopdf" in sys.argv)
