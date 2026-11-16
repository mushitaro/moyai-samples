import sys, glob, os, render
from xl import OUT
pats = sys.argv[1:]
for pt in pats:
    pat, pg = pt.split("@") if "@" in pt else (pt, "1")
    fs = sorted(glob.glob(os.path.join(OUT, "*" + pat + "*")))
    if not fs: print("NONE", pat); continue
    print(render.png(fs[0], [int(x) for x in pg.split(",")]))
