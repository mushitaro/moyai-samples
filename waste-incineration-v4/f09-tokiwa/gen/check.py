# -*- coding: utf-8 -*-
"""F09 正解とファイルの照合：python3 -I check.py
 1. answer/files/*.json の各行がファイルの中身と合う（xlsx は行の値、CSV は行、PDF は本文）
 2. どのファイルにも【架空】（ファイル名・表題・文書の情報）
 3. 実在の名前に当たらない
 4. 写真（画像ファイル）の大きさ（本施設は手書きスキャン・写真ファイルなし。あれば 1,568〜2,000px・JPEG 100KB 前後を確かめる）
 5. world.json の結び（links）が works と資料にあること
"""
import os, sys, json, csv, glob, re, subprocess, zipfile
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
FULL = os.path.normpath(os.path.join(HERE, ".."))
FILES = os.path.join(FULL, "files"); ANS = os.path.join(FULL, "answer")
REAL = ["日立", "荏原", "川崎重工", "三菱", "タクマ", "クボタ", "JFE", "神鋼", "日鉄", "住友", "IHI", "栗田", "東芝", "富士電機", "横河", "明電",
        "東京都", "大阪市", "横浜市", "名古屋市", "札幌市", "福岡市", "新明和", "極東開発", "水ing", "日本製鉄", "三井", "月島", "前澤", "メタウォーター"]
ISO = re.compile(r"^\d{4}-\d{2}-\d{2}$")
errs = []; stats = {"files": 0, "records": 0, "checked": 0}

def norm(s):
    return re.sub(r"\s+", "", str(s)).replace("，", ",").replace("〜", "～")

def scalars(fields):
    out = []
    for k, v in fields.items():
        if isinstance(v, (dict, list)): continue
        if isinstance(v, float) and k in ("manDays",): continue
        if k in ("manDays", "days", "ledgerTarget", "level", "cycle", "progress", "unit"): continue
        s = str(v)
        if ISO.match(s) or len(s) < 1: continue
        out.append((k, s))
    return out

def xlsx_rows(path):
    wb = openpyxl.load_workbook(path)
    sheets = {}
    for ws in wb.worksheets:
        merged = {}
        for mr in ws.merged_cells.ranges:
            v = ws.cell(mr.min_row, mr.min_col).value
            for r in range(mr.min_row, mr.max_row + 1):
                merged.setdefault(r, []).append(v)
        sheets[ws.title] = (ws, merged)
    return wb, sheets

def check_xlsx(path, ans):
    wb, sheets = xlsx_rows(path)
    for rec in ans["records"]:
        stats["records"] += 1
        ws, merged = sheets[rec["sheet"]]
        r = rec["r"]
        vals = [c.value for c in ws[r]] + merged.get(r, [])
        text = norm("|".join("" if v is None else str(v) for v in vals))
        sc = scalars(rec["fields"])
        if "path" in rec["fields"]:
            sc.append(("path", rec["fields"]["path"][-1].replace("■ ", "")))
        top = norm("|".join(str(c.value) for row in ws.iter_rows(min_row=1, max_row=4) for c in row if c.value is not None))
        miss = []
        for k, s in sc:
            ns = norm(s)
            if ns in text or ns.replace(".0", "") in text: continue
            if k in ("weather",) and ns in top: continue
            if k == "week" and f"第{ns}週" in top: continue
            if k in ("place", "time", "team", "kou", "company") and ("〃" in text or ws.cell(r, 1).value in (None, "")): continue
            miss.append(k)
        if "marks" in rec["fields"]:
            cells = [str(c.value) for c in ws[r] if c.value is not None]
            for y, m in rec["fields"]["marks"].items():
                if m not in cells: miss.append(f"marks{y}")
            nmarks = sum(1 for c in ws[r] if c.value in ("○", "◎", "●", "△", "◇", "◆"))
            if nmarks and ans["file"].find("組合") < 0 and nmarks != len(rec["fields"]["marks"]): miss.append("marks数")
        if not sc and "marks" not in rec["fields"] and "daily" not in rec["fields"] and "people" not in rec["fields"] and "plan" not in rec["fields"] and "actual" not in rec["fields"]:
            continue
        if miss:
            errs.append(f"{ans['file']} {rec['sheet']}!{r}: 合わない {miss[:4]}")
        stats["checked"] += 1
    # 【架空】
    t = (wb.properties.title or "") + (wb.properties.creator or "")
    if "【架空】" not in (wb.properties.title or "") or "【架空】" not in (wb.properties.creator or ""):
        errs.append(f"{ans['file']}: 文書のプロパティに【架空】が無い")
    ws0 = wb.worksheets[0]
    top = "".join(str(c.value) for row in ws0.iter_rows(min_row=1, max_row=3) for c in row if c.value)
    if "【架空】" not in top:
        errs.append(f"{ans['file']}: 表題のセルに【架空】が無い")
    alltext = "".join(str(c.value) for ws in wb.worksheets for row in ws.iter_rows() for c in row if c.value is not None)
    return alltext

def check_csv(path, ans):
    raw = open(path, "rb").read()
    enc = "utf-8-sig" if raw.startswith(b"\xef\xbb\xbf") else "cp932"
    text = raw.decode(enc)
    lines = text.split("\r\n")
    rows = list(csv.reader(lines))
    if "【架空】" not in lines[0]:
        errs.append(f"{ans['file']}: 1 行目に【架空】が無い")
    for rec in ans["records"]:
        stats["records"] += 1
        row = rows[rec["r"] - 1]
        line = norm("|".join(row))
        miss = [k for k, s in scalars(rec["fields"]) if norm(s) not in line]
        if miss: errs.append(f"{ans['file']} r{rec['r']}: 合わない {miss[:4]}")
        stats["checked"] += 1
    return text

def check_pdf(path, ans):
    txt = subprocess.run(["pdftotext", path, "-"], capture_output=True, text=True).stdout
    t = norm(txt)
    p1 = subprocess.run(["pdftotext", "-l", "1", path, "-"], capture_output=True, text=True).stdout
    info = subprocess.run(["pdfinfo", path], capture_output=True, text=True).stdout
    if "【架空】" not in p1: errs.append(f"{ans['file']}: 1 ページ目に【架空】が無い")
    if "【架空】" not in info: errs.append(f"{ans['file']}: 文書の情報に【架空】が無い")
    last = -1
    for rec in ans["records"]:
        stats["records"] += 1
        for k, s in scalars(rec["fields"]):
            if k in ("image", "section"): continue
            ns = norm(s)
            probe = ns[:12] if len(ns) > 12 else ns
            if probe not in t and not all(norm(x)[:12] in t for x in re.split(r"[\s　]+", s) if len(x) >= 2):
                errs.append(f"{ans['file']} seq{rec['seq']}: 本文に無い {k}={s[:20]}"); break
        stats["checked"] += 1
    return txt

import hashlib
HOW = {"direct", "annual", "bundle", "rollup", "upgrade"}
ATTRS = {"INSPECT", "OVERHAUL", "REPLACE", "UPGRADE", None}

def check_v2(files):
    """物差しの形（moyai.sample-world/0）"""
    w = json.load(open(os.path.join(ANS, "world.json"), encoding="utf-8"))
    if w.get("schema") != "moyai.sample-world/0": errs.append("world.json: schema")
    if w.get("facility", {}).get("id") != "F09" or "【架空】" not in w["facility"].get("name", ""): errs.append("world.json: facility")
    if w.get("maker", {}).get("today") != "2026-11-13": errs.append("world.json: maker.today")
    if "furnace" not in w.get("profile", {}): errs.append("world.json: profile.furnace")
    if not os.path.exists(os.path.join(ANS, "world.extra.json")): errs.append("world.extra.json が無い")
    if sorted(x["file"] for x in w["files"]) != sorted(files): errs.append("world.json: files[] と files/ が合わない")
    ids = [a["id"] for a in w["assets"]]
    if len(ids) != len(set(ids)): errs.append("assets: id が重なる")
    for a in w["assets"]:
        if a["part"] and a["part"] not in set(ids): errs.append(f"assets: part 不明 {a['id']}")
    if sum(1 for a in w["assets"] if a["part"] is None) != 648: errs.append("assets: 機器（part なし）が 648 でない")
    wid = {}
    for x in w["works"]:
        if x["id"] in wid: errs.append(f"works: id が重なる {x['id']}")
        wid[x["id"]] = x
        if x["attr"] not in ATTRS: errs.append(f"works: attr {x['id']}")
        if not isinstance(x["fy"], int): errs.append(f"works: fy {x['id']}")
        if x["manDays"] is not None and not isinstance(x["manDays"], (int, float)): errs.append(f"works: manDays {x['id']}")
        if x["asset"] and x["asset"] not in set(ids): errs.append(f"works: asset 不明 {x['id']}")
    recs = {}
    for fn in files:
        a = json.load(open(os.path.join(ANS, "files", os.path.splitext(fn)[0] + ".json"), encoding="utf-8"))
        for r in a["records"]:
            recs[(fn, r.get("sheet"), r["r"])] = r
            if fn.endswith((".pdf", ".csv")) and r.get("sheet") is not None: errs.append(f"{fn}: sheet が null でない")
            if fn.endswith(".pdf") and "seq" in r and r["r"] != r["seq"] + 1: errs.append(f"{fn}: r≠seq+1")
            fl = r["fields"]
            if a["plan"] and r["kind"] in ("parent", "item"):
                if not isinstance(fl.get("path"), list) or not isinstance(fl.get("marks"), dict) or not isinstance(fl.get("plans"), dict):
                    errs.append(f"{fn} {r.get('sheet')}!{r['r']}: path/marks/plans")
                for v in list(fl.get("marks", {}).values()) + list(fl.get("plans", {}).values()):
                    if len(str(v)) > 8: errs.append(f"{fn}: 8 字を越える印")
            if "manDays" in fl and not isinstance(fl["manDays"], (int, float)): errs.append(f"{fn}: manDays が数でない")
            if "days" in fl and not isinstance(fl["days"], (int, float)): errs.append(f"{fn}: days が数でない")
    bad = 0
    for l in w["links"]:
        if l["how"] not in HOW: errs.append(f"links: how {l['how']}"); bad += 1
        if l["work"] not in wid: errs.append(f"links: 作業が無い {l['work']}"); bad += 1
        if (l["file"], l.get("sheet"), l["r"]) not in recs: errs.append(f"links: 行が無い {l['file']} {l.get('sheet')}!{l['r']}"); bad += 1
        pf = json.load(open(os.path.join(ANS, "files", os.path.splitext(l["file"])[0] + ".json"), encoding="utf-8"))["plan"] if False else None
        if "year" in l and not isinstance(l["year"], int): errs.append("links: year"); bad += 1
        if bad > 20: break
    for e in w["errata"]:
        if e.get("type") not in ("omitMark", "lastYearUnreflected"): errs.append("errata: type")
    hp = os.path.normpath(os.path.join(FULL, "..", "files_hash_before.txt"))
    if os.path.exists(hp):
        now = sorted(f"{hashlib.sha256(open(os.path.join(FILES, fn), 'rb').read()).hexdigest()}  {fn}" for fn in files)
        if now != [x.rstrip("\n") for x in sorted(open(hp, encoding="utf-8"))]: errs.append("files/ のハッシュが控えと合わない")
    return f"links {len(w['links'])}（不備 {bad}）・works {len(w['works'])}・assets {len(w['assets'])}"

def main():
    files = sorted(os.listdir(FILES))
    texts = {}
    for fn in files:
        stats["files"] += 1
        if not fn.startswith("【架空】"): errs.append(f"{fn}: ファイル名が【架空】で始まらない")
        stem = os.path.splitext(fn)[0]
        ap = os.path.join(ANS, "files", stem + ".json")
        if not os.path.exists(ap):
            errs.append(f"{fn}: 正解が無い"); continue
        ans = json.load(open(ap, encoding="utf-8"))
        p = os.path.join(FILES, fn)
        if fn.endswith(".xlsx"): texts[fn] = check_xlsx(p, ans)
        elif fn.endswith(".csv"): texts[fn] = check_csv(p, ans)
        elif fn.endswith(".pdf"): texts[fn] = check_pdf(p, ans)
        elif fn.lower().endswith((".jpg", ".jpeg", ".png")):
            from PIL import Image
            im = Image.open(p); w, h = im.size; kb = os.path.getsize(p) / 1024
            if not (1568 <= max(w, h) <= 2000) or not (60 <= kb <= 160): errs.append(f"{fn}: 写真の大きさ {w}x{h} {kb:.0f}KB")
    for fn, t in texts.items():
        for w in REAL:
            if w in t or w in fn: errs.append(f"{fn}: 実在の名前「{w}」")
    imgs = [f for f in files if f.lower().endswith((".jpg", ".jpeg", ".png"))]
    bad = check_v2(files)
    print(f"ファイル {stats['files']}、正解の行 {stats['records']}（照合 {stats['checked']}）、画像 {len(imgs)}、v2 の結び {bad}")
    if errs:
        print(f"NG {len(errs)} 件"); [print(" -", e) for e in errs[:5000]]
        sys.exit(1)
    print("OK")

if __name__ == "__main__":
    main()
