# -*- coding: utf-8 -*-
"""F11 本作りの検査：python3 check.py
 1) 正解の行がファイルの中身と合う  2) どのファイルにも【架空】がある  3) 実在の名前に当たらない  4) 写真（スキャン）の大きさ"""
import os, re, json, glob, csv, subprocess, sys
import openpyxl
from PIL import Image

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FILES = os.path.join(BASE, "files")
ANS = os.path.join(BASE, "answer", "files")
REAL = ["日立", "三菱", "荏原", "JFE", "タクマ", "クボタ", "川崎重工", "住友", "神鋼", "IHI", "日立造船", "カナデビア", "栗田", "東芝", "富士電機", "明電舎",
        "横河", "SKF", "NTN", "日本精工", "東京都", "大阪市", "横浜市", "名古屋市", "札幌市", "福岡市", "神戸市", "京都市", "さいたま市", "千葉市"]

def norm(s):
    return re.sub(r"[\s　]", "", str(s))

errs = []
def err(f, m):
    errs.append("%s: %s" % (f, m))

PROBE = ["台帳No", "作業", "作業内容", "職種", "品名", "設備", "項目", "品名", "伝票番号", "内容", "結果", "説明", "測定", "主な作業"]

def probes(f):
    out = []
    for k in PROBE:
        v = f.get(k)
        if isinstance(v, str) and v and v != "読めない":
            out.append(v)
    if isinstance(f.get("path"), list) and f["path"]:
        out.append(str(f["path"][-1]))
    if isinstance(f.get("marks"), dict):
        out += [str(x) for x in f["marks"].values()]
    return out

texts = {}
n_rec = 0
for p in sorted(glob.glob(os.path.join(ANS, "*.json"))):
    d = json.load(open(p, encoding="utf-8"))
    fn = os.path.join(FILES, d["file"])
    if not os.path.exists(fn):
        err(d["file"], "ファイルが無い"); continue
    ext = fn.rsplit(".", 1)[1]
    if ext == "xlsx":
        wb = openpyxl.load_workbook(fn)
        if "【架空】" not in (wb.properties.title or ""):
            err(d["file"], "文書のプロパティ（表題）に【架空】が無い")
        first = wb.worksheets[0]
        if not any("【架空】" in str(c.value) for row in first.iter_rows(max_row=3) for c in row if c.value):
            err(d["file"], "表題のセルに【架空】が無い")
        alltxt = []
        for r in d["records"]:
            n_rec += 1
            ws = wb[r["sheet"]]
            rows = range(r["r"], r["r"] + (6 if d["file"].find("機器台帳") >= 0 and r["kind"] == "item" and r["sheet"] == "機器仕様台帳" else 1))
            t = norm("".join(str(c.value) for rr in rows for c in ws[rr] if c.value is not None))
            for pv in probes(r["fields"]):
                if norm(pv) not in t:
                    err(d["file"], "%s!%d に「%s」が無い" % (r["sheet"], r["r"], pv)); break
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for c in row:
                    if isinstance(c.value, str):
                        alltxt.append(c.value)
        texts[d["file"]] = "\n".join(alltxt)
    elif ext == "csv":
        raw = open(fn, "rb").read().decode("cp932")
        lines = raw.split("\r\n")
        if not lines[0].startswith("【架空】"):
            err(d["file"], "1 行目に【架空】が無い")
        rows = list(csv.reader(lines))
        for r in d["records"]:
            n_rec += 1
            row = rows[r["r"] - 1]
            f = r["fields"]
            if r["kind"] == "item":
                if f["伝票番号"] not in row or str(f["金額"]) not in row:
                    err(d["file"], "%d 行目が合わない" % r["r"])
            else:
                if str(f["金額"]) not in row:
                    err(d["file"], "%d 行目（小計）が合わない" % r["r"])
        texts[d["file"]] = raw
    elif ext == "pdf":
        info = subprocess.run(["pdfinfo", fn], capture_output=True, text=True).stdout
        if "【架空】" not in info:
            err(d["file"], "文書の情報に【架空】が無い")
        npages = int(re.search(r"Pages:\s+(\d+)", info).group(1))
        if d.get("images"):
            if npages != len(d["images"]):
                err(d["file"], "頁数 %d と画像の数 %d が違う" % (npages, len(d["images"])))
            for im in d["images"]:
                if not (1568 <= max(im["px"]) <= 2000):
                    err(d["file"], "画像 %s の大きさ %s" % (im["image"], im["px"]))
                if not (60 <= im["kb"] <= 140):
                    err(d["file"], "画像 %s の容量 %dKB" % (im["image"], im["kb"]))
            n_rec += len(d["records"])
            # スキャンは文字が無いので、原文と記号の数だけ確かめる
            for r in d["records"]:
                if r["fields"].get("作業内容") == "読めない" and not (r["fields"].get("原文") or r.get("_check", {}).get("原文")):
                    err(d["file"], "読めない行に原文が無い")
            texts[d["file"]] = ""
            continue
        t = subprocess.run(["pdftotext", "-raw", fn, "-"], capture_output=True, text=True).stdout
        p1 = subprocess.run(["pdftotext", "-f", "1", "-l", "1", fn, "-"], capture_output=True, text=True).stdout
        if "【架空】" not in p1:
            err(d["file"], "1 ページ目に【架空】が無い")
        tn = norm(t)
        for r in d["records"]:
            n_rec += 1
            for pv in probes(r["fields"]):
                if norm(pv)[:14] not in tn:
                    err(d["file"], "seq %d「%s」が本文に無い" % (r["seq"], pv[:20])); break
        texts[d["file"]] = t
# 実在の名前
for f, t in texts.items():
    for nm in REAL:
        pat = r"(?<![0-9０-９])日立(?!上)" if nm == "日立" else re.escape(nm)
        if re.search(pat, t):
            err(f, "実在の名前「%s」" % nm)
# ファイル名
for f in os.listdir(FILES):
    if not f.startswith("【架空】"):
        err(f, "ファイル名が【架空】で始まらない")
    if not os.path.exists(os.path.join(ANS, os.path.splitext(f)[0] + ".json")):
        err(f, "正解の json が無い")
# ---- 物差しの形（moyai.sample-world/0） ----
import hashlib
w = json.load(open(os.path.join(BASE, "answer", "world.json"), encoding="utf-8"))
if w.get("schema") != "moyai.sample-world/0": err("world.json", "schema")
if w["facility"].get("id") != "F11" or "【架空】" not in w["facility"].get("name", ""): err("world.json", "facility")
if w["maker"].get("today") != "2026-11-13": err("world.json", "maker.today")
if "furnace" not in w.get("profile", {}): err("world.json", "profile.furnace")
units = sum(a.get("units", 1) for a in w["assets"])
if len(w["assets"]) != 198 or units != 273:
    err("world.json", "機器 %d 行 %d 台" % (len(w["assets"]), units))
for a in w["assets"]:
    for k in ("id", "name", "tag", "system", "furnace", "installed", "class", "part", "names"):
        if k not in a: err("world.json", "assets %s に %s が無い" % (a.get("id"), k)); break
AID = {a["id"] for a in w["assets"]}
WID = {}
for x in w["works"]:
    WID[x["id"]] = x
    if x["attr"] not in ("INSPECT", "OVERHAUL", "REPLACE", "UPGRADE"): err("world.json", "works %s attr" % x["id"])
    if x["status"] not in ("done", "planned", "postponed", "not_done"): err("world.json", "works %s status" % x["id"])
    if x["asset"] is not None and x["asset"] not in AID: err("world.json", "works %s asset" % x["id"])
    if not isinstance(x["fy"], int): err("world.json", "works %s fy" % x["id"])
    if x["manDays"] is not None and not isinstance(x["manDays"], (int, float)): err("world.json", "works %s manDays" % x["id"])
FL = {f["file"] for f in w["files"]}
if FL != set(os.listdir(FILES)): err("world.json", "files[] と files/ が合わない")
REC = {}
for p in glob.glob(os.path.join(ANS, "*.json")):
    d = json.load(open(p, encoding="utf-8"))
    for r in d["records"]:
        REC[(d["file"], r.get("sheet"), r["r"])] = (d["plan"], r)
        if d["file"].endswith(".pdf"):
            if r.get("r") != r.get("seq", -9) + 1: err(d["file"], "r≠seq+1")
        if d["file"].endswith((".pdf", ".csv")) and r.get("sheet") is not None: err(d["file"], "sheet が null でない")
        f = r["fields"]
        if "manDays" in f and not isinstance(f["manDays"], (int, float)): err(d["file"], "manDays が数でない")
        if "days" in f and not isinstance(f["days"], (int, float)): err(d["file"], "days が数でない")
        if d["plan"] and r["kind"] in ("parent", "item"):
            for k in ("path", "marks", "plans"):
                if k not in f: err(d["file"], "%d に %s が無い" % (r["r"], k))
            for v in list(f.get("marks", {}).values()) + list(f.get("plans", {}).values()):
                if len(str(v)) > 8: err(d["file"], "印が 8 字を越える")
for l in w["links"]:
    if l["how"] not in ("direct", "annual", "bundle", "rollup", "upgrade"): err("links", "how %s" % l["how"])
    if l["work"] not in WID: err("links", "work %s が無い" % l["work"])
    if l["file"] not in FL: err("links", "file %s" % l["file"])
    k = (l["file"], l.get("sheet"), l["r"])
    if k not in REC: err("links", "行が無い %s" % (k,)); continue
    if REC[k][0] and "year" not in l: err("links", "星取表の結びに year が無い")
for e in w["errata"]:
    if e["type"] not in ("omitMark", "lastYearUnreflected"): err("errata", e["type"])
    if (e["file"], e["sheet"], e["r"]) not in REC: err("errata", "行が無い")
# files/ のハッシュ（変換の前に控えたもの）
hb = os.path.join(os.path.dirname(BASE), "files_sha256_before.txt")
if os.path.exists(hb):
    now = {f: hashlib.sha256(open(os.path.join(FILES, f), "rb").read()).hexdigest() for f in os.listdir(FILES)}
    for line in open(hb, encoding="utf-8"):
        hsh, f = line.strip().split("  ", 1)
        if now.get(f) != hsh: err(f, "ハッシュが変わった")
print("files %d, records %d, errors %d" % (len(os.listdir(FILES)), n_rec, len(errs)))
for e in errs[:60]:
    print(" -", e)
sys.exit(1 if errs else 0)
