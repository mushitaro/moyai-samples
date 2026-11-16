# -*- coding: utf-8 -*-
"""F06 本作りの検査
1. 正解の行がファイルの中身と合う（xlsx：シート・行に _check の文字がある／CSV：行に文字がある／PDF：seq 順に文字が出てくる）
2. どのファイルにも【架空】がある（xlsx：表題セルか文書の情報、CSV：1 行目、PDF：1 ページ目と文書の情報）
3. 実在の名前に当たらない（主なプラントメーカー・自治体・会社名の一覧と照合）
4. 写真の大きさ（この施設は手書きスキャン・写真なし。あれば 1,568〜2,000px・JPEG 30〜200KB を確認）
5. world.json の結び（links の work_id が works にある、実施済みの作業が資料の行かどうかの理由に必ず載る）
使い方：python3 check.py
"""
import csv
import glob
import io
import json
import os
import re
import subprocess
import sys

from openpyxl import load_workbook

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FILES = os.path.join(BASE, "files")
ADIR = os.path.join(BASE, "answer", "files")
REAL = ["三菱", "日立", "川崎重工", "荏原", "タクマ", "クボタ", "日鉄", "神鋼", "IHI", "JFE", "住友", "東芝", "富士電機", "横河", "栗田", "月島",
        "東京都", "横浜市", "大阪市", "名古屋市", "札幌市", "神戸市", "京都市", "福岡市", "川崎市", "さいたま市", "千葉市", "仙台市", "広島市",
        "港区", "日本環境衛生センター", "電気保安協会"]


def norm(s):
    return re.sub(r"\s+", "", str(s))


def check_xlsx(path, recs, errs):
    wb = load_workbook(path, read_only=False)
    cache = {}
    for rc in recs:
        sh = rc["sheet"]
        if sh not in wb.sheetnames:
            errs.append(f"{os.path.basename(path)}: シートなし {sh}")
            continue
        ws = wb[sh]
        key = (sh, rc["r"])
        if key not in cache:
            cache[key] = norm("".join(str(c.value) for c in ws[rc["r"]] if c.value is not None))
        row = cache[key]
        for s in rc.get("_check", []):
            if norm(s) not in row:
                errs.append(f"{os.path.basename(path)}: {sh}!{rc['r']} に「{s}」がない")
                break
    # 【架空】
    ok = "【架空】" in (wb.properties.title or "")
    for ws in wb.worksheets[:3]:
        for row in ws.iter_rows(min_row=1, max_row=3):
            for c in row:
                if c.value and "【架空】" in str(c.value):
                    ok = True
    if not ok:
        errs.append(f"{os.path.basename(path)}: 【架空】なし")
    text = []
    for ws in wb.worksheets:
        for row in ws.iter_rows(values_only=True):
            text.extend(str(v) for v in row if v is not None)
    return "".join(text)


def check_csv(path, recs, errs):
    raw = open(path, "rb").read()
    enc = "utf-8" if "新財務" in path else "cp932"
    txt = raw.decode(enc)
    lines = txt.splitlines()
    if "【架空】" not in lines[0]:
        errs.append(f"{os.path.basename(path)}: 1 行目に【架空】なし")
    for rc in recs:
        line = norm(lines[rc["r"] - 1])
        for s in rc.get("_check", []):
            if norm(s) not in line:
                errs.append(f"{os.path.basename(path)}: {rc['r']} 行目に「{s}」がない")
                break
    return txt


def check_pdf(path, recs, errs, warns):
    txt = subprocess.run(["pdftotext", "-layout", path, "-"], capture_output=True, text=True).stdout
    first = subprocess.run(["pdftotext", "-l", "1", path, "-"], capture_output=True, text=True).stdout
    info = subprocess.run(["pdfinfo", path], capture_output=True, text=True).stdout
    if "【架空】" not in first or "【架空】" not in info:
        errs.append(f"{os.path.basename(path)}: 1 ページ目か文書の情報に【架空】なし")
    flat = norm(txt)
    pos = 0
    disorder = 0
    for rc in sorted(recs, key=lambda x: x["seq"]):
        for s in rc.get("_check", [])[:1]:
            s = norm(s)[:8]
            k = flat.find(s, pos)
            if k < 0:
                if flat.find(s) < 0:
                    errs.append(f"{os.path.basename(path)}: seq {rc['seq']}「{s}」が PDF にない")
                else:
                    disorder += 1
            else:
                pos = k
    if disorder:
        warns.append(f"{os.path.basename(path)}: 順序が前後した行 {disorder}")
    return txt


def main():
    errs, warns = [], []
    files = sorted(os.listdir(FILES))
    nrec = 0
    alltext = {}
    for fn in files:
        path = os.path.join(FILES, fn)
        stem = os.path.splitext(fn)[0]
        ap = os.path.join(ADIR, stem + ".json")
        if not os.path.exists(ap):
            errs.append(f"{fn}: 正解なし")
            continue
        a = json.load(open(ap, encoding="utf-8"))
        recs = a["records"]
        nrec += len(recs)
        for rc in recs:
            if "_check" in rc.get("fields", {}):
                errs.append(f"{fn}: _check が fields の中にある")
            if fn.endswith(".pdf") and rc.get("r") != rc.get("seq", -9) + 1:
                errs.append(f"{fn}: PDF の記録の r が seq+1 でない")
                break
            if fn.endswith((".pdf", ".csv")) and rc.get("sheet") is not None:
                errs.append(f"{fn}: PDF・CSV の記録に sheet がある")
                break
            if a.get("plan") and rc["kind"] in ("parent", "item"):
                f = rc["fields"]
                if not isinstance(f.get("path"), list) or "marks" not in f or "plans" not in f:
                    errs.append(f"{fn}: 星取表の記録に path・marks・plans がない（r {rc.get('r')}）")
                    break
                if any(len(str(v)) > 8 for v in list(f["marks"].values()) + list(f["plans"].values())):
                    errs.append(f"{fn}: 8 字を越える印")
            if not a.get("plan") and rc["kind"] == "item" and "manDays" in rc["fields"] and rc["fields"]["manDays"] is not None \
                    and not isinstance(rc["fields"]["manDays"], (int, float)):
                errs.append(f"{fn}: manDays が数でない")
        if fn.endswith(".xlsx"):
            alltext[fn] = check_xlsx(path, recs, errs)
        elif fn.endswith(".csv"):
            alltext[fn] = check_csv(path, recs, errs)
        elif fn.endswith(".pdf"):
            alltext[fn] = check_pdf(path, recs, errs, warns)
        elif fn.lower().endswith((".jpg", ".jpeg", ".png")):
            from PIL import Image
            im = Image.open(path)
            if not (1568 <= max(im.size) <= 2000) or not (30_000 <= os.path.getsize(path) <= 200_000):
                errs.append(f"{fn}: 写真の大きさ {im.size} {os.path.getsize(path)}B")
        if "【架空】" not in fn:
            errs.append(f"{fn}: ファイル名に【架空】なし")
    for fn, t in alltext.items():
        for nm in REAL:
            if nm in t:
                errs.append(f"{fn}: 実在の名前らしき語「{nm}」")
    w = json.load(open(os.path.join(BASE, "answer", "world.json"), encoding="utf-8"))
    ex = json.load(open(os.path.join(BASE, "answer", "world.extra.json"), encoding="utf-8"))
    if w.get("schema") != "moyai.sample-world/0":
        errs.append("world.json: schema が違う")
    for k in ("facility", "profile", "maker", "assets", "works", "links", "errata", "files"):
        if k not in w:
            errs.append(f"world.json: {k} がない")
    if w["facility"].get("id") != "F06" or "【架空】" not in w["facility"].get("name", "") or w["maker"].get("today") != "2026-11-13":
        errs.append("world.json: facility / maker の値")
    ids = {x["id"] for x in w["works"]}
    aids = {x["id"] for x in w["assets"]}
    if any(x["asset"] not in aids for x in w["works"]):
        errs.append("world.json: 機器にない asset を持つ作業")
    bad = [l for l in w["links"] if l["work"] not in ids or l.get("how") not in ("direct", "annual", "bundle", "rollup", "upgrade")]
    if any(l["file"].endswith(".pdf") and not isinstance(l.get("r"), int) for l in w["links"]):
        errs.append("world.json: PDF を指す結びに r がない")
    if bad:
        errs.append(f"world.json: 存在しない work を指す結び {len(bad)}")
    if set(f["file"] for f in w["files"]) != set(files):
        errs.append("world.json: files[] が files/ と一致しない")
    for e in w["errata"]:
        if e["type"] not in ("omitMark", "lastYearUnreflected"):
            errs.append(f"world.json: errata の type {e['type']}")
    linked = {l["work"] for l in w["links"]}
    unrec = {u["id"] for u in ex["unrecorded_works"]}
    lost = [x["id"] for x in w["works"] if x["id"] not in linked and x["id"] not in unrec]
    if lost:
        errs.append(f"world.json: 資料にも理由にもない作業 {len(lost)}")
    photos = [f for f in files if f.lower().endswith((".jpg", ".jpeg", ".png"))]
    print(f"ファイル {len(files)}、正解の行 {nrec}、作業 {len(ids)}、結び {len(w['links'])}、errata {len(w['errata'])}、写真 {len(photos)}")
    for x in warns[:20]:
        print("注意:", x)
    if len(warns) > 20:
        print(f"注意 ほか {len(warns) - 20} 件")
    for x in errs[:50]:
        print("NG:", x)
    print("結果:", "OK" if not errs else f"NG {len(errs)} 件")
    return 0 if not errs else 1


if __name__ == "__main__":
    sys.exit(main())
