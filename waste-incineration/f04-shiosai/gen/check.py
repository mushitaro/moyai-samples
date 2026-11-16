# -*- coding: utf-8 -*-
"""F04 本作りの検査：python3 check.py
 1) 正解の行がファイルの中身と合う  2) どのファイルにも【架空】がある  3) 実在の名前に当たらない  4) 写真の大きさ"""
import os, sys, json, glob, re, subprocess
from openpyxl import load_workbook

HERE = os.path.dirname(os.path.abspath(__file__))
FILES = os.path.abspath(os.path.join(HERE, "..", "files"))
ANS = os.path.abspath(os.path.join(HERE, "..", "answer", "files"))

SKIP = {"日付", "No", "天候", "作業時間", "manDays", "manDaysText", "unit", "cycle", "周期", "comments", "予定", "実績", "月日", "曜", "cost", "work", "path", "marks", "code", "健全度", "設置年度"}
# 実在の会社・自治体の名前（一部）。帳票に出てはいけない
REAL = ["日立造船", "カナデビア", "JFE", "タクマ", "荏原", "川崎重工", "三菱重工", "神鋼環境", "住友重機", "クボタ", "IHI", "日鉄エンジニアリング", "栗田工業",
        "月島", "水ing", "三機工業", "極東開発", "新明和", "ユニチカ", "東芝", "富士電機", "明電舎", "安川電機", "横河電機", "アズビル", "日本碍子", "ヤンマー",
        "北海道", "青森", "東京都", "大阪市", "横浜市", "名古屋市", "札幌市", "福岡市", "電気保安協会"]

def norm(s):
    return re.sub(r"\s+", "", str(s)).replace("　", "")

def wareki(y):
    return "H%d" % (y - 1988) if y <= 2018 else "R%d" % (y - 2018)

def row_values(ws, r, merged):
    vals = []
    for c in ws[r]:
        if c.value is not None: vals.append(c.value)
    for (r1, c1, r2, c2) in merged:
        if r1 < r <= r2:
            v = ws.cell(row=r1, column=c1).value
            if v is not None: vals.append(v)
    return vals

def num_ok(v, vals, text):
    for x in vals:
        if isinstance(x, (int, float)) and not isinstance(x, bool) and abs(float(x) - float(v)) < 1e-6: return True
    s = ("%g" % v) if isinstance(v, float) else str(v)
    return s in text or ("{:,}".format(v) if isinstance(v, int) else s) in text

def check_xlsx(path, recs, errs):
    wb = load_workbook(path)
    cache = {}
    for i, rc in enumerate(recs):
        sh = rc.get("sheet")
        if sh not in wb.sheetnames:
            errs.append("%s: sheet %s が無い" % (os.path.basename(path), sh)); continue
        ws = wb[sh]
        if sh not in cache:
            cache[sh] = [(m.min_row, m.min_col, m.max_row, m.max_col) for m in ws.merged_cells.ranges]
        vals = row_values(ws, rc["r"], cache[sh])
        text = norm("".join(str(v) for v in vals))
        f = rc["fields"]
        bad = []
        for k, v in f.items():
            if k in SKIP or v is None or isinstance(v, (list, dict)): continue
            if isinstance(v, bool): continue
            if isinstance(v, (int, float)):
                if not num_ok(v, vals, text): bad.append(k)
            elif norm(v) not in text:
                bad.append(k)
        if "path" in f and rc["kind"] == "item":
            if norm(f["path"][-1]) not in text and norm(f["path"][-1]).replace("No.1", "No.1") not in text.replace("No.1", "No.1"):
                if not (f["path"][-1] == "灰出しコンベヤNo.1" and "灰出しコンベヤNo.1" in text): bad.append("path")
        for mk in ("marks", "plans"):
          if mk not in f: continue
          for y, m in f[mk].items():
                if isinstance(m, (int, float)):
                    if not num_ok(m, vals, text): bad.append(mk + y)
                else:
                    for part in (m if isinstance(m, list) else str(m).split("/")):
                        if part and norm(part) not in text: bad.append(mk + y)
        if "設置年度" in f and isinstance(f["設置年度"], int):
            if wareki(f["設置年度"]) not in text: bad.append("設置年度")
        if "code" in f and norm(f["code"]) not in text: bad.append("code")
        if bad:
            errs.append("%s [%s r%s] 合わない: %s" % (os.path.basename(path)[:60], sh, rc["r"], ",".join(sorted(set(bad)))))
    return wb

def pdf_text(path):
    return subprocess.run(["pdftotext", "-layout", path, "-"], capture_output=True, text=True).stdout

def check_pdf(path, recs, errs):
    t = norm(pdf_text(path))
    t2 = norm(subprocess.run(["pdftotext", "-raw", path, "-"], capture_output=True, text=True).stdout)
    for rc in recs:
        for k, v in rc["fields"].items():
            if k in SKIP or v in (None, "", "－") or isinstance(v, (list, dict, bool)): continue
            s = norm(v)
            if isinstance(v, (int, float)):
                if s not in t and ("%g" % v) not in t: errs.append("%s seq%d %s=%s" % (os.path.basename(path)[:50], rc["seq"], k, v))
                continue
            chunks = [s[i:i + 5] for i in range(0, max(1, len(s) - 4), 7)] or [s]
            hit = sum(1 for ch in chunks if ch in t or ch in t2)
            if hit < max(1, int(len(chunks) * 0.7)):
                errs.append("%s seq%d %s=%s" % (os.path.basename(path)[:50], rc["seq"], k, s[:20]))

def check_v2(files):
    """物差しの形（moyai.sample-world/0）と files/ の不変"""
    import hashlib
    e = []
    base = os.path.abspath(os.path.join(HERE, "..", "answer"))
    w = json.load(open(os.path.join(base, "world.json"), encoding="utf-8"))
    if w.get("schema") != "moyai.sample-world/0": e.append("schema が違う")
    if w.get("facility", {}).get("id") != "F04" or "【架空】" not in w.get("facility", {}).get("name", ""): e.append("facility が違う")
    if w.get("maker", {}).get("today") != "2026-11-13": e.append("maker.today が違う")
    if not w.get("profile", {}).get("furnace"): e.append("profile.furnace が無い")
    if not os.path.exists(os.path.join(base, "world.extra.json")): e.append("world.extra.json が無い")
    if sorted(x["file"] for x in w["files"]) != sorted(files): e.append("files[] が files/ と合わない")
    aid = set()
    for a in w["assets"]:
        for k in ("id", "name", "tag", "system", "furnace", "installed", "class", "part", "names"):
            if k not in a: e.append("assets に %s が無い: %s" % (k, a.get("id")))
        aid.add(a["id"])
    wid = set()
    for x in w["works"]:
        wid.add(x["id"])
        if x["asset"] not in aid: e.append("works.asset が assets に無い: " + x["id"])
        if x["attr"] not in ("INSPECT", "OVERHAUL", "REPLACE", "UPGRADE"): e.append("attr: " + x["id"])
        if x["status"] not in ("done", "planned", "postponed", "not_done"): e.append("status: " + x["id"])
        if not isinstance(x["fy"], int): e.append("fy: " + x["id"])
        if x.get("manDays") is not None and not isinstance(x["manDays"], (int, float)): e.append("manDays: " + x["id"])
    recs = {}
    plans = set()
    for p in glob.glob(os.path.join(ANS, "*.json")):
        d = json.load(open(p, encoding="utf-8"))
        if d.get("plan"): plans.add(d["file"])
        for rc in d["records"]:
            if not isinstance(rc.get("r"), int): e.append("r が数でない: %s" % d["file"]); break
            if d["file"].endswith(".pdf") and rc.get("sheet") is not None: e.append("PDF に sheet: " + d["file"]); break
            recs[(d["file"], rc.get("sheet"), rc["r"])] = rc
            f = rc["fields"]
            if d.get("plan") and rc["kind"] in ("item", "parent"):
                for k in ("path", "marks", "plans"):
                    if k not in f: e.append("星取表に %s が無い: %s r%d" % (k, d["file"][:40], rc["r"]))
                for k in ("marks", "plans"):
                    for y, v in f.get(k, {}).items():
                        if not (len(y) == 4 and y.isdigit()) or len(str(v)) > 8: e.append("印の形: %s r%d %s" % (d["file"][:40], rc["r"], y))
            if "manDays" in f and not isinstance(f["manDays"], (int, float)): e.append("manDays が数でない: " + d["file"])
            if rc["kind"] == "total" and "manDays" in f: e.append("合計行に manDays: " + d["file"])
    for l in w["links"]:
        if l["work"] not in wid: e.append("links.work が無い: " + l["work"])
        if (l["file"], l.get("sheet"), l["r"]) not in recs: e.append("links の行が記録に無い: %s %s %s" % (l["file"][:40], l.get("sheet"), l["r"]))
        if l["file"] in plans and not isinstance(l.get("year"), int): e.append("星取表の結びに year が無い")
        if l.get("how") not in ("direct", "annual", "bundle", "rollup", "upgrade"): e.append("how: %s" % l.get("how"))
    for x in w.get("errata", []):
        if x["type"] not in ("omitMark", "lastYearUnreflected"): e.append("errata の種類")
        if (x["file"], x["sheet"], x["r"]) not in recs: e.append("errata の行が無い")
        for i in x["works"]:
            if i not in wid: e.append("errata.works が無い")
    # files/ のハッシュ
    before = os.path.abspath(os.path.join(HERE, "..", "..", "notes", "files_sha256_before.txt"))
    if os.path.exists(before):
        want = {}
        for line in open(before, encoding="utf-8"):
            h, n = line.strip().split("  ", 1); want[n] = h
        for n in files:
            h = hashlib.sha256(open(os.path.join(FILES, n), "rb").read()).hexdigest()
            if want.get(n) != h: e.append("files/ が変わった: " + n)
        if len(want) != len(files): e.append("files/ の数が変わった")
    return e

def main():
    errs = []
    js = sorted(glob.glob(os.path.join(ANS, "*.json")))
    files = sorted(os.listdir(FILES))
    stems = {os.path.splitext(os.path.basename(j))[0] for j in js}
    for fn in files:
        if os.path.splitext(fn)[0] not in stems: errs.append("正解が無い: " + fn)
        if not fn.startswith("【架空】"): errs.append("ファイル名に【架空】が無い: " + fn)
    nrec = 0
    for j in js:
        d = json.load(open(j, encoding="utf-8"))
        p = os.path.join(FILES, d["file"])
        if not os.path.exists(p): errs.append("ファイルが無い: " + d["file"]); continue
        nrec += len(d["records"])
        if p.endswith(".xlsx"):
            wb = check_xlsx(p, d["records"], errs)
            if "【架空】" not in (wb.properties.title or ""): errs.append("文書の表題に【架空】が無い: " + d["file"])
            ws0 = wb[wb.sheetnames[0]]
            head = "".join(str(c.value) for row in ws0.iter_rows(min_row=1, max_row=12) for c in row if c.value)
            if "【架空】" not in head: errs.append("表題のセルに【架空】が無い: " + d["file"])
            alltext = "".join(str(c.value) for ws in wb for row in ws.iter_rows() for c in row if c.value)
        else:
            check_pdf(p, d["records"], errs)
            info = subprocess.run(["pdfinfo", p], capture_output=True, text=True).stdout
            if "【架空】" not in info: errs.append("PDF の文書情報に【架空】が無い: " + d["file"])
            p1 = subprocess.run(["pdftotext", "-f", "1", "-l", "1", p, "-"], capture_output=True, text=True).stdout
            if "【架空】" not in p1: errs.append("PDF 1 ページ目に【架空】が無い: " + d["file"])
            alltext = pdf_text(p)
            # 写真の大きさ
            lst = subprocess.run(["pdfimages", "-list", p], capture_output=True, text=True).stdout.splitlines()[2:]
            for l in lst:
                a = l.split()
                if len(a) > 4:
                    w_, h_ = int(a[3]), int(a[4])
                    if not (300 <= w_ <= 2000 and 200 <= h_ <= 2000): errs.append("写真の大きさ %dx%d: %s" % (w_, h_, d["file"]))
            nimg = sum(1 for r in d["records"] if r.get("image"))
            if nimg and len(lst) < nimg: errs.append("写真の数が合わない %d<%d: %s" % (len(lst), nimg, d["file"]))
        for nm in REAL:
            if nm in alltext: errs.append("実在の名前らしき語「%s」: %s" % (nm, d["file"]))
    errs += check_v2(files)
    print("files %d / answers %d / records %d" % (len(files), len(js), nrec))
    if errs:
        print("NG %d" % len(errs))
        for e in errs[:400]: print("  " + e)
        sys.exit(1)
    print("OK")

if __name__ == "__main__":
    main()
