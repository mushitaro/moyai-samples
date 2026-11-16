# -*- coding: utf-8 -*-
"""F08 本作りの検査：python3 -I check.py
1. 正解の行（answer/files/*.json）がファイルの中身と合う
2. どのファイルにも【架空】がある（中身と文書の情報）
3. 実在の名前（大手のプラント・重電・計測メーカー等）に当たらない
4. 写真の大きさ（F08 は手書きスキャン・写真なし。画像ファイルが無いことを確かめる）
5. world.json（moyai.sample-world/0）の形と結び：作業・機器・ファイル・記録の行が実在する
6. 星取表の記録に path・marks・plans、ほかの資料の manDays・days が数"""
import csv
import glob
import io
import json
import os
import re
import subprocess
import sys

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FILES = os.path.join(BASE, "files")
ANS = os.path.join(BASE, "answer")
REAL = ["日立", "三菱", "荏原", "タクマ", "川崎重工", "JFE", "クボタ", "住友", "IHI", "神鋼", "栗田", "東芝", "富士電機",
        "横河", "明電", "安川", "新日鉄", "日鉄", "カナデビア", "メタウォーター", "月島", "三井", "クリタ", "オルガノ", "コベルコ",
        "鹿島", "大成", "清水建設", "大林", "竹中", "ダイキン", "パナソニック", "鴻池", "大同特殊鋼"]
PRIMARY = ["code", "name", "work", "件名", "市番号", "機器名称", "品名", "text", "工事名", "機器"]
fail = []


def norm(x):
    return re.sub(r"\s+", "", str(x))


def scalars(v):
    if isinstance(v, dict):
        for x in v.values():
            yield from scalars(x)
    elif isinstance(v, list):
        for x in v:
            yield from scalars(x)
    elif v is not None and not isinstance(v, bool):
        yield v


def xlsx_check(fn, recs):
    import openpyxl
    wb = openpyxl.load_workbook(os.path.join(FILES, fn))
    first = wb[wb.sheetnames[0]]
    if "【架空】" not in str(first["A1"].value) or "【架空】" not in (wb.properties.title or ""):
        fail.append(f"{fn}: 【架空】が表題セルか文書の情報に無い")
    text = []
    for ws in wb.worksheets:
        for row in ws.iter_rows(values_only=True):
            text += [str(v) for v in row if v is not None]
    bad = 0
    for rec in recs:
        ws = wb[rec["sheet"]]
        cells = [norm(c.value) for r in (rec["r"], rec["r"] + 1) for c in ws[r] if c.value is not None]
        rowtxt = "|".join(cells)
        f = rec["fields"]
        key = next((k for k in PRIMARY if k in f and isinstance(f[k], str)), None)
        if key and norm(f[key]) not in rowtxt:
            bad += 1
            if bad <= 3:
                fail.append(f"{fn}: {rec['sheet']}!{rec['r']} の {key}={f[key]!r} が行に無い")
        for mk in ("marks", "plans"):
            for y, s in (f.get(mk) or {}).items():
                if norm(s) not in rowtxt:
                    bad += 1
                    if bad <= 3:
                        fail.append(f"{fn}: {rec['sheet']}!{rec['r']} の {mk}[{y}]={s} が行に無い")
        for y, v in (f.get("cost") or {}).items():
            if norm(v) not in rowtxt:
                bad += 1
                if bad <= 3:
                    fail.append(f"{fn}: {rec['sheet']}!{rec['r']} の cost[{y}]={v} が行に無い")
    return "\n".join(text), bad


def csv_check(fn, recs):
    raw = open(os.path.join(FILES, fn), "rb").read()
    txt = raw.decode("cp932")
    rows = list(csv.reader(io.StringIO(txt)))
    if "【架空】" not in rows[0][0]:
        fail.append(f"{fn}: 1 行目に【架空】が無い")
    head = rows[2]
    bad = 0
    for rec in recs:
        row = rows[rec["r"] - 1]
        if dict(zip(head, row)) != rec["fields"]:
            bad += 1
            fail.append(f"{fn}: {rec['r']} 行目が正解と違う")
    return txt, bad


def pdf_check(fn, recs):
    p = os.path.join(FILES, fn)
    info = subprocess.run(["pdfinfo", p], capture_output=True, text=True).stdout
    p1 = subprocess.run(["pdftotext", "-f", "1", "-l", "1", p, "-"], capture_output=True, text=True).stdout
    if "【架空】" not in p1 or "【架空】" not in info:
        fail.append(f"{fn}: 1 ページ目か文書の情報に【架空】が無い")
    full = subprocess.run(["pdftotext", p, "-"], capture_output=True, text=True).stdout
    flat = norm(full)
    pos, bad = 0, 0
    for rec in sorted(recs, key=lambda r: r.get("seq", r["r"] - 1)):
        f = rec["fields"]
        key = next((k for k in PRIMARY if k in f and isinstance(f[k], str)), None)
        if not key:
            continue
        probe = norm(re.sub(r"<[^>]+>", "", f[key]))[:18]
        k = flat.find(probe, max(0, pos - 400))
        if k < 0:
            bad += 1
            if bad <= 3:
                fail.append(f"{fn}: seq {rec['seq']} の {key} が本文に無い：{probe}")
        else:
            pos = k
    return full, bad


def main():
    n_rec = 0
    texts = {}
    ans = sorted(glob.glob(os.path.join(ANS, "files", "*.json")))
    files = sorted(os.listdir(FILES))
    stems = {os.path.splitext(f)[0]: f for f in files}
    for a in ans:
        d = json.load(open(a, encoding="utf-8"))
        fn = d["file"]
        if fn not in files:
            fail.append(f"{os.path.basename(a)}: 資料 {fn} が無い")
            continue
        recs = d["records"]
        n_rec += len(recs)
        if fn.endswith(".xlsx"):
            texts[fn], _ = xlsx_check(fn, recs)
        elif fn.endswith(".csv"):
            texts[fn], _ = csv_check(fn, recs)
        else:
            texts[fn], _ = pdf_check(fn, recs)
    for f in files:
        if os.path.splitext(f)[0] not in {os.path.splitext(json.load(open(a, encoding="utf-8"))["file"])[0] for a in ans}:
            fail.append(f"{f}: 正解ファイルが無い")
        if not f.startswith("【架空】"):
            fail.append(f"{f}: ファイル名が【架空】で始まらない")
    for fn, t in texts.items():
        for nm in REAL:
            if nm in t or nm in fn:
                fail.append(f"{fn}: 実在の名前「{nm}」に当たる")
    imgs = [f for f in files if f.lower().endswith((".jpg", ".jpeg", ".png", ".tif"))]
    if imgs:
        for f in imgs:
            fail.append(f"{f}: F08 には写真・スキャンを入れない設定")
    w = json.load(open(os.path.join(ANS, "world.json"), encoding="utf-8"))
    if w.get("schema") != "moyai.sample-world/0":
        fail.append("world.json: schema が moyai.sample-world/0 でない")
    if w.get("facility", {}).get("id") != "F08" or "【架空】" not in w.get("facility", {}).get("name", ""):
        fail.append("world.json: facility の id・name")
    if w.get("maker", {}).get("today") != "2026-11-13":
        fail.append("world.json: maker.today")
    aids = {a["id"] for a in w["assets"]}
    if len(aids) != len(w["assets"]) or len(aids) != 318:
        fail.append(f"world.json: assets の数・一意 {len(w['assets'])}")
    for a in w["assets"]:
        if a["part"] and a["part"] not in aids:
            fail.append(f"assets: {a['id']} の part {a['part']} が無い")
    wids = {x["id"] for x in w["works"]}
    for x in w["works"]:
        if x["asset"] not in aids:
            fail.append(f"works: {x['id']} の asset {x['asset']} が無い")
        if x["attr"] not in ("INSPECT", "OVERHAUL", "REPLACE", "UPGRADE"):
            fail.append(f"works: {x['id']} の attr")
        if x["manDays"] is not None and not isinstance(x["manDays"], (int, float)):
            fail.append(f"works: {x['id']} の manDays が数でない")
    recidx = set()
    for a in ans:
        dd = json.load(open(a, encoding="utf-8"))
        for rec in dd["records"]:
            recidx.add((dd["file"], rec.get("sheet"), rec["r"]))
            if dd["plan"] and rec["kind"] in ("item", "parent"):
                f = rec["fields"]
                if not all(k in f for k in ("path", "marks", "plans")):
                    fail.append(f"{dd['file']}: r{rec['r']} に path・marks・plans が無い")
                for mk in ("marks", "plans"):
                    for y, v in f.get(mk, {}).items():
                        if not re.match(r"^\d{4}$", y) or len(str(v)) > 8:
                            fail.append(f"{dd['file']}: r{rec['r']} の {mk} {y}={v}")
            elif not dd["plan"]:
                for k in ("manDays", "days"):
                    if k in rec["fields"] and not isinstance(rec["fields"][k], (int, float)):
                        fail.append(f"{dd['file']}: r{rec['r']} の {k} が数でない")
            if (dd["file"].endswith(".pdf") or dd["file"].endswith(".csv")) and rec.get("sheet"):
                fail.append(f"{dd['file']}: PDF・CSV の記録に sheet がある")
    if {x["file"] for x in w["files"]} != set(files):
        fail.append("world.json: files[] と files/ が合わない")
    for l in w["links"]:
        if l["file"] not in files:
            fail.append(f"links: {l['file']} が無い")
        if l["work"] not in wids:
            fail.append(f"links: 作業 {l['work']} が無い")
        if l["how"] not in ("direct", "annual", "bundle", "rollup", "upgrade"):
            fail.append(f"links: how {l['how']}")
        if (l["file"], l.get("sheet"), l["r"]) not in recidx:
            fail.append(f"links: {l['file']} {l.get('sheet')} r{l['r']} の記録が無い")
    plan_files = {json.load(open(a, encoding="utf-8"))["file"] for a in ans if json.load(open(a, encoding="utf-8"))["plan"]}
    for l in w["links"]:
        if l["file"] in plan_files and not isinstance(l.get("year"), int):
            fail.append(f"links: 星取表の結びに year が無い {l['file']} r{l['r']}")
    for e in w["errata"]:
        if e["type"] not in ("omitMark", "lastYearUnreflected") or (e["file"], e.get("sheet"), e["r"]) not in recidx:
            fail.append(f"errata: {e}")
        for x in e["works"]:
            if x not in wids:
                fail.append(f"errata: 作業 {x} が無い")
    for f in ("user.md", "README.md"):
        if not os.path.exists(os.path.join(BASE, f)):
            fail.append(f"{f} が無い")
    print(f"資料 {len(files)} 件、正解ファイル {len(ans)} 件、正解の行 {n_rec} 件、機器 {len(aids)} 台、作業 {len(wids)} 件、結び {len(w['links'])} 件、errata {len(w['errata'])} 件")
    print(f"画像（写真・スキャン）{len(imgs)} 件")
    if fail:
        print("NG", len(fail))
        for x in fail[:40]:
            print(" -", x)
        sys.exit(1)
    print("OK")


if __name__ == "__main__":
    main()
