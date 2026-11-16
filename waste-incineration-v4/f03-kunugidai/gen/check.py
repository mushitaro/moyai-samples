# -*- coding: utf-8 -*-
"""F03 本作りの検査。  python3 gen/check.py
1. 正解の行（answer/files/*.json）がファイルの中身と合う
2. どのファイルにも【架空】がある（Excel：表題セルとプロパティ、CSV：1 行目、PDF：1 ページ目と文書情報）
3. 実在の社名・自治体名に当たらない
4. 写真・スキャン画像の大きさ
5. world.json の台数・作業の結び"""
import os, sys, json, csv, re, glob, subprocess, io
from openpyxl import load_workbook
from pypdf import PdfReader

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FILES = os.path.join(BASE, "files")
ANS = os.path.join(BASE, "answer", "files")
REAL = ["タクマ", "日立造船", "カナデビア", "荏原", "JFE", "三菱重工", "川崎重工", "神鋼環境", "クボタ", "栗田工業", "住友重機", "IHI", "東芝", "富士電機",
        "明電舎", "横河電機", "日鉄", "前澤工業", "月島", "鹿島建設", "大林組", "清水建設", "竹中工務店", "大成建設", "水ing", "メタウォーター", "東京都", "大阪市"]
SKIP = {"marks", "plans", "days", "comments", "values", "path", "manDays", "manDaysText", "page", "cycle", "cost", "unit", "旧No", "code", "kW", "No"}
DATEKEYS = {"予定開始", "予定終了", "実績開始", "実績終了", "工期開始", "工期終了", "日付", "停止期間", "月日"}
norm = lambda s: re.sub(r"\s+", "", str(s))
errs = []
stats = {"files": 0, "records": 0, "checked": 0}


def err(f, m):
    if len(errs) < 100000:
        errs.append("%s: %s" % (f, m))
    else:
        errs.append(None)


def _pdftext(p):
    a = subprocess.run(["pdftotext", "-layout", p, "-"], capture_output=True, text=True).stdout
    b = subprocess.run(["pdftotext", "-raw", p, "-"], capture_output=True, text=True).stdout
    return norm(a) + "||" + norm(b)


def check_xlsx(path, recs, fname):
    wb = load_workbook(path, read_only=False)
    cache = {}
    merged = {}
    for w in wb.worksheets:
        mm = {}
        for rg in w.merged_cells.ranges:
            v = w.cell(rg.min_row, rg.min_col).value
            if v is None:
                continue
            for rr in range(rg.min_row, rg.max_row + 1):
                mm.setdefault(rr, []).append(norm(v))
        merged[w.title] = mm
    for rc in recs:
        sh = rc.get("sheet")
        if sh not in wb.sheetnames:
            err(fname, "シートが無い %s" % sh)
            continue
        ws = wb[sh]
        r0, r1 = rc.get("r_start", rc["r"]), rc.get("r_end", rc["r"])
        key = (sh, r0, r1)
        if key not in cache:
            vals = []
            for rr in list(range(r0, r1 + 1)) + [1, 2, 3, 4, 5]:
                for c in ws[rr]:
                    if c.value is not None:
                        vals.append(norm(c.value))
                vals += merged[sh].get(rr, [])
            cache[key] = "|".join(vals)
        row = cache[key]
        for k, v in rc["fields"].items():
            if k in SKIP or k in DATEKEYS or isinstance(v, (dict, list)) or v in (None, ""):
                continue
            stats["checked"] += 1
            s = norm(v)
            if isinstance(v, float) and v == int(v):
                s = str(int(v))
            if s not in row and s[:15] not in row:
                err(fname, "%s!%d %s=%s が行に無い" % (sh, r0, k, str(v)[:30]))
        for y, m in (rc["fields"].get("marks") or {}).items():
            stats["checked"] += 1
            if norm(m) not in row:
                err(fname, "%s!%d marks %s=%s が無い" % (sh, r0, y, m))
    # 【架空】
    if not (wb.properties.title or "").startswith("【架空】"):
        err(fname, "文書プロパティの表題に【架空】が無い")
    ws = wb[wb.sheetnames[0]]
    if not any("【架空】" in str(c.value) for row in ws.iter_rows(min_row=1, max_row=2) for c in row if c.value):
        err(fname, "表題セルに【架空】が無い")
    txt = " ".join(str(c.value) for w in wb.worksheets for row in w.iter_rows() for c in row if c.value is not None)
    return txt


def check_csv(path, recs, fname):
    raw = open(path, "rb").read()
    txt = raw.decode("utf-8-sig") if raw.startswith(b"\xef\xbb\xbf") else raw.decode("cp932")
    lines = txt.split("\r\n")
    if "【架空】" not in lines[0]:
        err(fname, "1 行目に【架空】が無い")
    rows = list(csv.reader(io.StringIO("\r\n".join(lines[1:]))))
    hdr = rows[0]
    for rc in recs:
        r = rc["r"]
        if r - 2 >= len(rows):
            err(fname, "行が無い %d" % r)
            continue
        row = rows[r - 2]
        for k, v in rc["fields"].items():
            stats["checked"] += 1
            if k in hdr:
                if norm(row[hdr.index(k)]) != norm(v):
                    err(fname, "%d %s: %s≠%s" % (r, k, row[hdr.index(k)], v))
            elif norm(v) not in norm(",".join(row)):
                err(fname, "%d %s=%s が行に無い" % (r, k, v))
    return txt


def check_pdf(path, recs, fname, scan):
    rd = PdfReader(path)
    title = (rd.metadata or {}).get("/Title", "") or ""
    if "【架空】" not in title:
        err(fname, "文書情報の表題に【架空】が無い")
    if scan:
        n = len(rd.pages)
        for rc in recs:
            stats["checked"] += 1
            if rc["fields"].get("page", 1) > n:
                err(fname, "頁が無い")
        for i, pg in enumerate(rd.pages):
            for im in pg.images:
                w, h = im.image.size
                kb = len(im.data) / 1024.0
                if not (1568 <= max(w, h) <= 2000):
                    err(fname, "スキャン %d 頁の画像の大きさ %dx%d" % (i + 1, w, h))
                if not (60 <= kb <= 180):
                    err(fname, "スキャン %d 頁の JPEG %.0fKB" % (i + 1, kb))
        return ""
    t = _pdftext(path)
    p1 = norm(subprocess.run(["pdftotext", "-f", "1", "-l", "1", path, "-"], capture_output=True, text=True).stdout)
    if "【架空】" not in p1:
        err(fname, "1 ページ目に【架空】が無い")
    for rc in recs:
        for k, v in rc["fields"].items():
            if isinstance(v, (dict, list)) or not v:
                continue
            stats["checked"] += 1
            s = norm(v)
            if s not in t and s[:12] not in t and s[:6] not in t:
                err(fname, "seq %d %s=%s が本文に無い" % (rc["seq"], k, str(v)[:30]))
        if rc.get("image"):
            stats["checked"] += 1
    for i, pg in enumerate(rd.pages):
        for im in pg.images:
            w, h = im.image.size
            if not (400 <= max(w, h) <= 2000):
                err(fname, "写真 %d 頁 %dx%d" % (i + 1, w, h))
    return t


def check_v2(allf):
    """物差しの形（moyai.sample-world/0）と files/ の不変"""
    import hashlib
    w = json.load(open(os.path.join(BASE, "answer", "world.json"), encoding="utf-8"))
    stats["links"], stats["works"] = len(w.get("links", [])), len(w.get("works", []))
    if w.get("schema") != "moyai.sample-world/0":
        err("world.json", "schema")
    if w.get("facility", {}).get("id") != "F03" or "【架空】" not in w.get("facility", {}).get("name", ""):
        err("world.json", "facility")
    if w.get("maker", {}).get("today") != "2026-11-13":
        err("world.json", "maker.today")
    if "furnace" not in w.get("profile", {}):
        err("world.json", "profile.furnace")
    ids = [a["id"] for a in w["assets"]]
    if len(ids) != 725 or len(set(ids)) != 725:
        err("world.json", "assets %d" % len(ids))
    for a in w["assets"]:
        for k in ("id", "name", "tag", "system", "furnace", "installed", "class", "part", "names"):
            if k not in a:
                err("world.json", "asset %s に %s が無い" % (a.get("id"), k))
                break
    aset = set(ids)
    wid = set()
    for x in w["works"]:
        wid.add(x["id"])
        if x["asset"] not in aset:
            err("world.json", "work %s の asset %s" % (x["id"], x["asset"]))
        if x["attr"] not in ("INSPECT", "OVERHAUL", "REPLACE", "UPGRADE"):
            err("world.json", "work %s attr %s" % (x["id"], x["attr"]))
        if x["status"] not in ("done", "planned", "postponed", "not_done"):
            err("world.json", "work %s status" % x["id"])
        if not isinstance(x["fy"], int):
            err("world.json", "work %s fy" % x["id"])
        if x.get("manDays") is not None and not isinstance(x["manDays"], (int, float)):
            err("world.json", "work %s manDays" % x["id"])
    names = {x["file"] for x in w["files"]}
    if names != set(allf):
        err("world.json", "files[] とフォルダが合わない（%d / %d）" % (len(names), len(allf)))
    plans = {}
    for p in glob.glob(os.path.join(ANS, "*.json")):
        a = json.load(open(p, encoding="utf-8"))
        plans[a["file"]] = a.get("plan")
        for rc in a["records"]:
            if not isinstance(rc.get("r"), int):
                err(a["file"], "r が数でない記録")
                break
            if a.get("plan") and rc["kind"] in ("parent", "item"):
                f = rc["fields"]
                if not isinstance(f.get("path"), list) or not isinstance(f.get("marks"), dict) or not isinstance(f.get("plans"), dict):
                    err(a["file"], "%s!%d path・marks・plans" % (rc.get("sheet"), rc["r"]))
                for v in list(f["marks"].values()) + list(f["plans"].values()):
                    if len(str(v)) > 8:
                        err(a["file"], "8 字を越える印")
            if a["file"].endswith((".pdf", ".csv")) and rc.get("sheet"):
                err(a["file"], "PDF・CSV の sheet")
    for l in w["links"]:
        if l["file"] not in names or l["work"] not in wid or not isinstance(l.get("r"), int):
            err("world.json", "link %s" % l)
            break
        if plans.get(l["file"]) and not isinstance(l.get("year"), int):
            err("world.json", "星取表の結びに year が無い %s" % l)
            break
    for e in w.get("errata", []):
        if e.get("type") not in ("omitMark", "lastYearUnreflected"):
            err("world.json", "errata type")
    # files/ のハッシュ
    hb = os.path.join(BASE, "build", "hash_before.txt")
    if os.path.exists(hb):
        before = dict(reversed(line.rstrip("\n").split("  ", 1)) for line in open(hb, encoding="utf-8"))
        for fn in allf:
            h = hashlib.sha256(open(os.path.join(FILES, fn), "rb").read()).hexdigest()
            if before.get(fn) != h:
                err(fn, "files/ のハッシュが変わった")
        if set(before) != set(allf):
            err("files/", "ファイルの増減")
        stats["hash"] = "一致" if not any(e and "ハッシュ" in e for e in errs) else "不一致"


def main():
    allf = sorted(os.listdir(FILES))
    texts = {}
    for fname in allf:
        stats["files"] += 1
        if not fname.startswith("【架空】"):
            err(fname, "ファイル名が【架空】で始まらない")
        ap = os.path.join(ANS, os.path.splitext(fname)[0] + ".json")
        if not os.path.exists(ap):
            err(fname, "正解ファイルが無い")
            continue
        a = json.load(open(ap, encoding="utf-8"))
        recs = a["records"]
        stats["records"] += len(recs)
        p = os.path.join(FILES, fname)
        if fname.endswith(".xlsx"):
            texts[fname] = check_xlsx(p, recs, fname)
        elif fname.endswith(".csv"):
            texts[fname] = check_csv(p, recs, fname)
        elif fname.endswith(".pdf"):
            texts[fname] = check_pdf(p, recs, fname, "手書き" in fname)
    for fname, t in texts.items():
        for nm in REAL:
            if nm in t:
                err(fname, "実在の名前に当たる：%s" % nm)
    check_v2(allf)
    n_err = len([e for e in errs])
    print("files %d / records %d / checked values %d / links %d / works %d" % (stats["files"], stats["records"], stats["checked"], stats["links"], stats["works"]))
    for e in errs[:60]:
        if e:
            print("NG", e)
    print("files/ hash:", stats.get("hash", "未確認"))
    print("RESULT:", "OK" if not errs else "NG %d" % n_err)
    return 0 if not errs else 1


if __name__ == "__main__":
    sys.exit(main())
