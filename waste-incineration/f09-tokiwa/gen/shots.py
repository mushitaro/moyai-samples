import os, sys, glob
HERE = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, HERE)
import render
render.WORK = os.path.join(HERE, "tmp", "render")
F = os.path.join(HERE, "..", "files"); OUT = os.path.join(HERE, "png")
picks = [("長期修繕計画2010年度版", [1]), ("長期修繕計画2021年度改訂", [1]), ("令和6年度策定", [2]), ("機器台帳_2009年度末", [1]),
         ("令和6年4月1日現在", [2]), ("工程表_2012年度1号炉", [1]), ("工程表_2020年度2号炉", [1]), ("工程表_2024年度2号炉", [1]),
         ("工程表_基幹的設備改良工事R8", [1]), ("作業日報_2011年度1号炉", [3]), ("作業日報_2019年度DCS", [5]), ("作業日報_2026年度2号炉", [4]),
         ("作業日報_基幹的設備改良工事R8", [10]), ("工事報告書_2012年度2号炉", [2, 3]), ("工事報告書_2021年度1号炉", [2, 3]), ("工事報告書_2019年度DCS", [2]),
         ("発注台帳_2012年度", [1]), ("発注台帳_2020年度", [1])]
sel = sys.argv[1:] or None
for key, pages in picks:
    if sel and not any(s in key for s in sel): continue
    fs = [x for x in glob.glob(os.path.join(F, "*")) if key in x]
    if not fs: print("missing", key); continue
    inf = "CSV:44,34,76" if "新購買" in fs[0] else None
    print(render.to_png(fs[0], pages, OUT, infilter=inf, dpi=70))
