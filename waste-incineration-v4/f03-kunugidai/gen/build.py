# -*- coding: utf-8 -*-
"""F03【架空】くぬぎ台環境センター 本作り（2001〜2026 年度）。  python3 gen/build.py
files/ answer/ を作り直す。工程 → 日報 → 報告書 → 発注 の順。最後に check.py を流す。"""
import os, sys, json, shutil, subprocess, glob, time
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import xlhelp as X
import world as W
import model as M
import texts
import rec as R
import f_plan, f_ledger, f_sched, f_daily, f_scan, f_report, f_order
import rep_base as RB
from world import wareki

P = "【架空】くぬぎ台環境センター_"
T0 = time.time()


def log(*a):
    print("[%5.0fs]" % (time.time() - T0), *a, flush=True)


def short(c):
    if c.kind == "teishu":
        return "%d号炉定期整備業務委託" % c.unit
    if c.kind == "kyotsu":
        return "全炉停止時共通設備整備業務委託"
    return c.name.split("くぬぎ台環境センター")[-1].strip().replace(" ", "")


def main():
    for d in (X.OUT, os.path.join(X.BASE, "answer"), X.BUILD):
        if os.path.isdir(d):
            shutil.rmtree(d)
        os.makedirs(d)
    tmp = os.path.join(X.BUILD, "doc")
    os.makedirs(tmp)
    R.LINKS.clear()
    eqs = W.world()
    m = M.build(eqs)
    M.apply_direct_samples(m)
    log("model", len(m.contracts), "contracts", len(M.WORKS), "works")
    docs = [c for c in m.contracts.values() if c.docs and not getattr(c, "planned_only", False)]
    for c in docs:
        texts.gen_daily(c, m)
    # 直営の工数
    for it in m.direct:
        if it["wid"]:
            M.WORKS[it["wid"]]["manDays"] = round((M.WORKS[it["wid"]]["manDays"] or 0) + it["n"] * it["h"] / 8.0, 2)
    # 履歴を実際の作業（WORKS）にそろえる（計画・台帳はこれを読む）
    byt = {}
    for w in M.WORKS.values():
        byt.setdefault(w["tag"], []).append(dict(fy=w["fy"], kind=w["kind"], what=w["what"], contract=w["contract"]))
    for e in eqs:
        e.events = sorted(byt.get(e.tag, []), key=lambda x: x["fy"])
    o = lambda n: os.path.join(X.OUT, n)
    files = []

    # ---- 中長期整備計画（成立した版だけ）
    n = P + "中長期整備計画_H13_整備年次計画_k1様式.xlsx"; f_plan.k1_2001(eqs, o(n), n); files.append(("中長期整備計画", n, 2001))
    n = P + "中長期整備計画_H23_長寿命化計画原案_k2様式.xlsx"; f_plan.k2(eqs, 2011, o(n), n); files.append(("中長期整備計画", n, 2011))
    n = P + "中長期整備計画_H30版_k2様式.xlsx"; f_plan.k2(eqs, 2018, o(n), n); files.append(("中長期整備計画", n, 2018))
    for v in range(2019, 2027):
        n = P + "中長期整備計画_%s版_k3様式.xlsx" % wareki(v, True)
        f_plan.k3(eqs, v, o(n), n)
        files.append(("中長期整備計画", n, v))
    log("plans")
    # ---- 機器台帳
    n = P + "機器台帳_H15_k1仕様ブロック様式.xlsx"; f_ledger.make_k1(eqs, o(n), n); files.append(("機器台帳", n, 2003))
    n = P + "機器台帳_H26年4月_k2炉別シート様式.xlsx"; f_ledger.make_old(eqs, o(n), 2013, n); files.append(("機器台帳", n, 2014))
    n = P + "機器台帳_R3年度末_k2炉別シート様式.xlsx"; f_ledger.make_old(eqs, o(n), 2021, n); files.append(("機器台帳", n, 2021))
    n = P + "機器台帳_R8年10月_k3統合様式.xlsx"; f_ledger.make_new(eqs, o(n), n); files.append(("機器台帳", n, 2026))
    log("ledgers")
    # ---- 工程表（先）→ 日報 → 報告書
    for fy in range(2001, 2027):
        cs = [c for c in m.by_fy(fy) if c.kind in ("teishu", "kyotsu", "hoshu")]
        if fy <= 2014:
            for c in cs:
                if not c.docs:
                    continue
                n = P + "工程表_%s_%s_m1様式.xlsx" % (wareki(fy, True), short(c))
                f_sched.m1(c, o(n), n)
                files.append(("工程表", n, fy))
        else:
            n = P + "工程表_%s_定期整備（全契約）_m2様式.xlsx" % wareki(fy, True)
            f_sched.m2(fy, cs, o(n), n, m)
            files.append(("工程表", n, fy))
    log("schedules")
    for c in docs:
        if not c.daily:
            continue
        if c.fy <= 2013:
            n = P + "作業日報_%s_%s_m1様式.xlsx" % (wareki(c.fy, True), short(c))
            f_daily.m1(c, o(n), n)
        else:
            n = P + "作業日報_%s_%s_m2様式.xlsx" % (wareki(c.fy, True), short(c))
            f_daily.m2(c, o(n), n)
        files.append(("作業日報", n, c.fy))
    log("daily")
    for c in docs:
        st = "r1" if c.fy <= 2007 else ("r2旧" if c.fy <= 2019 else "r2新")
        n = P + "工事報告書_%s_%s_%s様式.pdf" % (wareki(c.fy, True), short(c), st)
        if c.sample:
            f_report.sample(c.sample, c, o(n), tmp)
        else:
            f_report.generic(c, m, o(n), tmp)
        files.append(("工事報告書", n, c.fy))
    log("report docx", len(RB.PENDING))
    _convert(RB.PENDING)
    log("report pdf")
    # ---- 発注（控 → CSV → 控の Excel）
    hik, pmap = f_order.build_hikae(m)
    for fy in range(2001, 2027):
        rows = f_order.year_rows(fy, m, hik)
        if fy <= 2013:
            n = P + "発注台帳_%s_旧財務システム契約台帳.csv" % wareki(fy, True)
            f_order.write_old(fy, rows, o(n), n)
        else:
            n = P + "発注台帳_%s_新財務システム契約一覧.csv" % wareki(fy, True)
            f_order.write_new(fy, rows, o(n), n)
        files.append(("発注台帳", n, fy))
    n = P + "発注台帳_H24-R8_直営部品発注控_k2様式.xlsx"
    f_order.write_hikae(hik, o(n), n)
    files.append(("発注台帳", n, 2012))
    log("orders")
    # ---- 直営
    for fy in range(2001, 2017):
        n = P + "作業日報_%s_直営補修班_d1手書きスキャン.pdf" % wareki(fy, True)
        k = f_scan.make_year(fy, m, o(n), n, pmap, tmp)
        if k:
            files.append(("作業日報", n, fy))
    for fy in range(2017, 2027):
        n = P + "作業日報_%s_直営補修班_d2週報様式.xlsx" % wareki(fy, True)
        f_daily.d2_weekly(fy, m, o(n), n, pmap)
        files.append(("作業日報", n, fy))
    log("direct")
    import out_world
    out_world.write(m, eqs, files)
    log("world")
    with open(os.path.join(X.BUILD, "files.json"), "w", encoding="utf-8") as fh:
        json.dump(files, fh, ensure_ascii=False, indent=0)
    log("done", len(files), "files")


def _convert(pending):
    if not pending:
        return
    outd = os.path.join(X.BUILD, "docpdf")
    os.makedirs(outd, exist_ok=True)
    dx = [p for p, _ in pending]
    for i in range(0, len(dx), 40):
        subprocess.run(["soffice", "-env:UserInstallation=" + X.LO_PROFILE, "--headless", "--convert-to", "pdf", "--outdir", outd] + dx[i:i + 40],
                       check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=1800)
    for d, p in pending:
        src = os.path.join(outd, os.path.splitext(os.path.basename(d))[0] + ".pdf")
        shutil.move(src, p)
    pending.clear()


if __name__ == "__main__":
    main()
