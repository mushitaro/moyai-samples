# -*- coding: utf-8 -*-
"""正解（answer/files）とファイルの中身の照合、【架空】の有無、実在の名前、写真の大きさ。
  python3 check.py   → 不一致を表示し、件数をまとめる。終了コード 0＝合格"""
import csv
import glob
import json
import os
import re
import subprocess
import sys
from openpyxl import load_workbook

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
FILES = os.path.join(ROOT, "files")
ANS = os.path.join(ROOT, "answer", "files")
# 実在する会社・自治体・施設の名前（当たったら不合格）。架空名と紛らわしい一般語は入れない。
REAL = ["日立", "三菱", "JFE", "荏原", "タクマ", "川崎重工", "IHI", "住友", "栗田", "クボタ", "神鋼", "日鉄", "東芝", "富士電機", "明電舎",
        "横河", "アズビル", "山武", "島津", "堀場", "新明和", "IHI", "前田建設", "鹿島", "大成", "清水建設", "大林", "東京都", "大阪市",
        "横浜市", "名古屋市", "札幌市", "福岡市", "日本ボイラ協会", "ボイラ・クレーン安全協会", "神鋼環境", "日立造船", "カナデビア"]
NORM = re.compile(r"\s+")


def norm(s):
    return NORM.sub("", str(s))


def row_text_xlsx(ws, r):
    out = []
    for c in ws[r]:
        if c.value is not None:
            out.append(str(c.value))
    return norm("".join(out))


def keys_to_check(f):
    vals = []
    for k in ("code", "name", "work", "processNo", "company", "text", "label", "memo", "equipment", "result"):
        v = f.get(k)
        if isinstance(v, (str, int)) and str(v).strip() and str(v) not in ("〃",):
            vals.append(str(v) if k not in ("text", "memo", "result") else str(v)[:12])
    for u in (f.get("units") or {}).values():
        vals.append(u["code"])
    for m in list((f.get("marks") or {}).values()) + list((f.get("plans") or {}).values()):
        vals.append(m)
    return vals


def check_xlsx(path, ans, errs):
    wb = load_workbook(path)
    for rec in ans["records"]:
        ws = wb[rec["sheet"]]
        t = row_text_xlsx(ws, rec["r"])
        for v in keys_to_check(rec["fields"]):
            if norm(v) not in t:
                errs.append(f"{os.path.basename(path)} {rec['sheet']}!{rec['r']}: 「{v}」が行に無い")
                break
        if ans.get("plan") and rec["kind"] in ("item", "parent"):
            fl = rec["fields"]
            if "path" not in fl or "marks" not in fl or "plans" not in fl:
                errs.append(f"{os.path.basename(path)} {rec['sheet']}!{rec['r']}: path・marks・plans が無い")
                continue
            if any(len(str(v)) > 8 for v in list(fl["marks"].values()) + list(fl["plans"].values())):
                errs.append(f"{os.path.basename(path)} {rec['sheet']}!{rec['r']}: 8 字を越える印")
            n = sum(1 for c in ws[rec["r"]] if c.value in ("○", "◎", "●", "△", "▲", "◇", "☆", "★"))
            sp2 = "sp2星取表" in os.path.basename(path)
            if not sp2 and fl["marks"] != fl["plans"]:
                errs.append(f"{os.path.basename(path)} {rec['sheet']}!{rec['r']}: 予定だけの様式で marks≠plans")
            exp = len(fl["marks"]) + (len(fl["plans"]) if sp2 else 0)
            if n != exp:
                errs.append(f"{os.path.basename(path)} {rec['sheet']}!{rec['r']}: 記号の数 {n} ≠ 正解 {exp}")
    # 【架空】
    t = (wb.properties.title or "") + "".join(str(ws["A1"].value or "") for ws in wb.worksheets)
    if "【架空】" not in t:
        errs.append(f"{os.path.basename(path)}: 【架空】が表題・プロパティに無い")
    txt = []
    for ws in wb.worksheets:
        for row in ws.iter_rows(values_only=True):
            txt += [str(v) for v in row if isinstance(v, str)]
    return "\n".join(txt)


def check_mandays(path, ans, errs):
    for rec in ans["records"]:
        f = rec["fields"]
        if "manDaysUnit" in f and f["manDaysUnit"] not in ("人日",):
            errs.append(f"{os.path.basename(path)}: manDaysUnit が人日でない")
        if f.get("manDays") is not None and "_hours" in rec:
            if abs(f["manDays"] * 8 - rec["_hours"]) > 0.005:
                errs.append(f"{os.path.basename(path)} {rec.get('sheet')}!{rec.get('r')}: 人日と時間が合わない")
            if not f.get("manDaysText"):
                errs.append(f"{os.path.basename(path)} {rec.get('sheet')}!{rec.get('r')}: manDaysText が無い")


def check_world(errs, files):
    w = json.load(open(os.path.join(ROOT, "answer", "world.json"), encoding="utf-8"))
    if w.get("schema") != "moyai.sample-world/0":
        errs.append("world.json: schema が違う")
    if w.get("facility", {}).get("id") != "F07" or "【架空】" not in w["facility"].get("name", ""):
        errs.append("world.json: facility が違う")
    if w.get("maker", {}).get("today") != "2026-11-13":
        errs.append("world.json: maker.today")
    if not w.get("profile", {}).get("furnace"):
        errs.append("world.json: profile.furnace が無い")
    aid = {a["id"] for a in w["assets"]}
    for a in w["assets"]:
        if a["part"] and a["part"] not in aid:
            errs.append(f"world.json: asset {a['id']} の part が無い")
        if not isinstance(a["installed"], int):
            errs.append(f"world.json: asset {a['id']} installed")
    wid = set()
    for x in w["works"]:
        wid.add(x["id"])
        if x["asset"] not in aid:
            errs.append(f"world.json: work {x['id']} の asset が無い")
        if x["attr"] not in ("INSPECT", "OVERHAUL", "REPLACE", "UPGRADE"):
            errs.append(f"world.json: work {x['id']} attr")
        if not isinstance(x["fy"], int):
            errs.append(f"world.json: work {x['id']} fy")
    fset = {os.path.basename(p) for p in files}
    for l in w["links"]:
        if l.get("how") not in ("direct", "annual", "bundle", "rollup", "upgrade"):
            errs.append(f"world.json: link how {l.get('how')}")
        if l["file"].endswith(".pdf") and not isinstance(l.get("r"), int):
            errs.append(f"world.json: PDF の link に r が無い")
        if l["work"] not in wid:
            errs.append(f"world.json: link の work {l['work']} が無い")
        if l["file"] not in fset:
            errs.append(f"world.json: link の file {l['file']} が無い")
        if "中長期整備計画" in l["file"] and not isinstance(l.get("year"), int):
            errs.append(f"world.json: 星取表の link に year が無い {l['file']} r{l['r']}")
    for e in w["errata"]:
        if e["type"] not in ("omitMark", "lastYearUnreflected") or e["file"] not in fset:
            errs.append(f"world.json: errata {e}")
    if {f["file"] for f in w["files"]} != fset:
        errs.append("world.json: files[] が files/ と合わない")
    return len(w["assets"]), len(w["works"]), len(w["links"]), len(w["errata"])


def check_csv(path, ans, errs):
    enc = "cp932" if "発注台帳" in path else "utf-8-sig"
    with open(path, encoding=enc, newline="") as f:
        lines = list(csv.reader(f))
    if "【架空】" not in "".join(lines[0]):
        errs.append(f"{os.path.basename(path)}: 1 行目に【架空】が無い")
    for rec in ans["records"]:
        row = lines[rec["r"] - 1]
        cells = set(row)
        for k, v in rec["fields"].items():
            if v is None or v == "":
                continue
            sv = str(v)
            if re.match(r"\d{4}-\d{2}-\d{2}$", sv):
                sv = sv.replace("-", "/")
            if sv not in cells:
                errs.append(f"{os.path.basename(path)} r{rec['r']}: {k}=「{sv}」が無い")
                break
    return "\n".join(",".join(l) for l in lines)


def check_pdf(path, ans, errs):
    txt = subprocess.run(["pdftotext", path, "-"], capture_output=True, text=True).stdout
    t = norm(txt)
    p1 = norm(subprocess.run(["pdftotext", "-f", "1", "-l", "1", path, "-"], capture_output=True, text=True).stdout)
    info = subprocess.run(["pdfinfo", path], capture_output=True, text=True).stdout
    if "【架空】" not in p1 or "【架空】" not in info:
        errs.append(f"{os.path.basename(path)}: 1 ページ目または文書情報に【架空】が無い")
    seqs = [r["seq"] for r in ans["records"]]
    if any(r.get("r") != r["seq"] + 1 for r in ans["records"]):
        errs.append(f"{os.path.basename(path)}: PDF の r ≠ seq+1")
    if seqs != list(range(len(seqs))):
        errs.append(f"{os.path.basename(path)}: seq が 0 から連番でない")
    for rec in ans["records"]:
        f = rec["fields"]
        key = None
        for k in ("code", "name", "label", "equipment", "品名", "stage", "text", "保全ｺｰﾄﾞ", "列"):
            if f.get(k):
                key = str(f[k])[:10]
                break
        if key and norm(key) not in t:
            errs.append(f"{os.path.basename(path)} seq{rec['seq']}: 「{key}」が本文に無い")
    return txt


def main():
    errs = []
    alltext = {}
    files = sorted(glob.glob(os.path.join(FILES, "*")))
    for p in files:
        stem = os.path.splitext(os.path.basename(p))[0]
        if not os.path.basename(p).startswith("【架空】"):
            errs.append(f"{os.path.basename(p)}: ファイル名が【架空】で始まらない")
        ap = os.path.join(ANS, stem + ".json")
        if not os.path.exists(ap):
            errs.append(f"{os.path.basename(p)}: 正解ファイルが無い")
            continue
        ans = json.load(open(ap, encoding="utf-8"))
        if ans["file"] != os.path.basename(p):
            errs.append(f"{stem}: 正解の file 名が違う")
        ext = os.path.splitext(p)[1].lower()
        check_mandays(p, ans, errs)
        for rec in ans["records"]:
            if ext in (".csv", ".pdf") and rec.get("sheet") is not None:
                errs.append(f"{os.path.basename(p)}: CSV・PDF の記録に sheet がある")
                break
        if ext == ".xlsx":
            alltext[p] = check_xlsx(p, ans, errs)
        elif ext == ".csv":
            alltext[p] = check_csv(p, ans, errs)
        elif ext == ".pdf":
            alltext[p] = check_pdf(p, ans, errs)
        elif ext in (".jpg", ".jpeg", ".png"):
            from PIL import Image
            im = Image.open(p)
            if not (1568 <= max(im.size) <= 2000):
                errs.append(f"{os.path.basename(p)}: 写真の大きさ {im.size}")
    for p, t in alltext.items():
        for name in REAL:
            if name in t:
                errs.append(f"{os.path.basename(p)}: 実在の名前「{name}」")
    wa = check_world(errs, files)
    nrec = sum(len(json.load(open(a, encoding="utf-8"))["records"]) for a in glob.glob(os.path.join(ANS, "*.json")))
    imgs = [p for p in files if p.lower().endswith((".jpg", ".jpeg", ".png"))]
    print(f"ファイル {len(files)}、正解の行 {nrec}、写真 {len(imgs)}（手書きスキャンはこの施設では作らない）")
    print(f"world.json：機器 {wa[0]}、作業 {wa[1]}、結び {wa[2]}、errata {wa[3]}")
    if errs:
        print(f"不合格 {len(errs)} 件")
        for e in errs[:int(os.environ.get('NERR', 60))]:
            print(" ", e)
        sys.exit(1)
    print("合格")


if __name__ == "__main__":
    main()
