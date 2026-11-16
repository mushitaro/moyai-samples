# -*- coding: utf-8 -*-
"""F02 本作りの検査。python3 check.py（full/gen で実行）
1. 正解の行（answer/files/*.json）がファイルの中身と合う
2. どのファイルにも【架空】がある（Excel は表題セルと文書プロパティ、CSV は 1 行目、PDF は 1 ページ目と文書情報）
3. 実在の名前に当たらない
4. 写真（画像）の大きさ（この施設は手書きスキャン・写真画像なし → 画像ファイルが無いことを確認）"""
import csv
import glob
import io
import json
import os
import re
import sys
from collections import Counter
import openpyxl
import pdfplumber

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FILES = os.path.join(BASE, "files")
ANS = os.path.join(BASE, "answer", "files")
REAL = ["日立造船", "カナデビア", "JFE", "タクマ", "荏原", "三菱重工", "川崎重工", "IHI", "クボタ", "栗田工業", "住友重機", "神鋼環境",
        "日鉄エンジ", "新明和", "極東開発", "メタウォーター", "東芝", "富士電機", "明電舎", "三機工業", "東京都", "大阪市", "横浜市",
        "名古屋市", "札幌市", "福岡市", "川崎市", "神戸市", "京都市", "日本環境衛生センター", "日立製作所"]
SKIP = {"manDays", "manDaysText", "path", "特記", "日付", "days", "週", "表", "comments", "hiddenMarks", "marks", "cost", "unit",
        "予定開始", "予定終了", "実績開始", "実績終了", "開始", "終了", "年月日", "職種別人員", "work", "cycle", "号炉", "計実", "台数",
        "1号炉", "2号炉", "共通", "区分", "No", "番号", "memo", "前回実施", "廃止まで", "重要度", "code", "設備区分名", "人員計", "人員", "工数",
        "時間", "進捗", "週工数", "特記", "daily", "plans"}
SHEETKEYS = {"天候", "気温", "作業場所", "使用機械", "品質記録", "安全・連絡", "特記", "作業内容"}
errs = []
stat = Counter()


def norm(s):
    return re.sub(r"\s+", "", str(s))


def err(f, m):
    errs.append(f"{f}: {m}")


def row_values(ws, r):
    vals = []
    for c in ws[r]:
        if c.value is not None:
            vals.append(str(c.value))
    for mr in ws.merged_cells.ranges:
        if mr.min_row <= r <= mr.max_row and mr.min_row != r:
            v = ws.cell(row=mr.min_row, column=mr.min_col).value
            if v is not None:
                vals.append(str(v))
    return vals


def check_xlsx(fn, recs):
    wb = openpyxl.load_workbook(os.path.join(FILES, fn))
    if "【架空】" not in (wb.properties.title or ""):
        err(fn, "文書プロパティの表題に【架空】が無い")
    first = wb.worksheets[0]
    if not any("【架空】" in str(c.value) for row in first.iter_rows(max_row=3) for c in row if c.value):
        err(fn, "表題セルに【架空】が無い")
    text = []
    for ws in wb.worksheets:
        for row in ws.iter_rows(values_only=True):
            text += [str(v) for v in row if v is not None]
    for rr in recs:
        ws = wb[rr["sheet"]] if rr["sheet"] else wb.worksheets[0]
        vals = row_values(ws, rr["r"])
        joined = norm("|".join(vals))
        sheet_txt = None
        for k, v in rr["fields"].items():
            if k in SKIP or v in (None, "") or isinstance(v, (dict, list)):
                continue
            hay = joined
            if k in SHEETKEYS and norm(v) not in joined:
                if sheet_txt is None:
                    sheet_txt = norm("|".join(str(c.value) for row in ws.iter_rows() for c in row if c.value is not None))
                hay = sheet_txt
            if norm(v) not in hay:
                err(fn, f"{rr['sheet']}!{rr['r']} {k}={v!r} が行に無い")
                break
        m = rr["fields"].get("marks")
        if m:
            cnt = Counter(vals)
            need = Counter(m.values())
            hid = rr.get("_hiddenMarks") or {}
            need.update(hid.values())
            for sym, n in need.items():
                if cnt[sym] < n:
                    err(fn, f"{rr['sheet']}!{rr['r']} 記号 {sym} が {n} 個必要（{cnt[sym]}）")
        for addr, txt in (rr["fields"].get("comments") or {}).items():
            cm = ws[addr].comment
            if not cm or norm(txt) not in norm(cm.text):
                err(fn, f"コメント {addr} が合わない")
        stat["rows"] += 1
    return "\n".join(text)


def check_csv(fn, recs):
    raw = open(os.path.join(FILES, fn), "rb").read()
    enc = "utf-8-sig" if raw.startswith(b"\xef\xbb\xbf") else "cp932"
    s = raw.decode(enc)
    lines = s.splitlines()
    if "【架空】" not in lines[0]:
        err(fn, "1 行目に【架空】が無い")
    head = next(csv.reader([lines[1]]))
    for rr in recs:
        row = next(csv.reader([lines[rr["r"] - 1]]))
        d = dict(zip(head, row))
        for k, v in rr["fields"].items():
            if norm(d.get(k, "")) != norm(v):
                err(fn, f"行 {rr['r']} {k}: {d.get(k)!r} != {v!r}")
                break
        stat["rows"] += 1
    return s


def near(x, full):
    """表のセル内で折り返した語（行の間に他の列の文字が入る）も拾う：近い範囲に順に現れれば可。"""
    if x in full:
        return True
    i = full.find(x[:1])
    while i >= 0:
        j, k = i, 0
        lim = i + len(x) * 12 + 40
        while j < min(len(full), lim) and k < len(x):
            if full[j] == x[k]:
                k += 1
            j += 1
        if k == len(x):
            return True
        i = full.find(x[:1], i + 1)
    return False


def check_pdf(fn, recs):
    with pdfplumber.open(os.path.join(FILES, fn)) as pdf:
        meta = pdf.metadata.get("Title", "")
        pages = [p.extract_text() or "" for p in pdf.pages]
    if "【架空】" not in meta:
        err(fn, "文書情報の表題に【架空】が無い")
    if "【架空】" not in pages[0]:
        err(fn, "1 ページ目に【架空】が無い")
    full = norm("".join(pages))
    for rr in recs:
        for k, v in rr["fields"].items():
            if k in SKIP or v in (None, "") or isinstance(v, (dict, list)):
                continue
            parts = [x for x in re.split(r"[\s/・、，,（）()「」:：]+", str(v)) if len(x) >= 2]
            if any(not near(norm(x)[:30], full) for x in parts):
                err(fn, f"seq {rr['seq']} {k}={str(v)[:30]!r} が本文に無い")
                break
        stat["rows"] += 1
    return "\n".join(pages)


def check_world(w, ans, fs):
    """物差しの形（moyai.sample-world/0）に合っているか。"""
    if w.get("schema") != "moyai.sample-world/0":
        errs.append("world.json: schema が moyai.sample-world/0 でない")
    if w.get("facility", {}).get("id") != "F02" or "【架空】" not in w.get("facility", {}).get("name", ""):
        errs.append("world.json: facility.id／name")
    if w.get("maker", {}).get("today") != "2026-11-13":
        errs.append("world.json: maker.today")
    if not w.get("profile", {}).get("furnace"):
        errs.append("world.json: profile.furnace が無い")
    if sorted(x["file"] for x in w["files"]) != sorted(fs):
        errs.append("world.json: files[] が files/ と一致しない")
    aid = {a["id"] for a in w["assets"]}
    for a in w["assets"]:
        for k in ("id", "name", "tag", "system", "furnace", "installed", "class", "part", "names"):
            if k not in a:
                errs.append(f"assets {a.get('id')}: {k} が無い")
                break
    ids = set()
    for o in w["works"]:
        ids.add(o["id"])
        if o["asset"] is not None and o["asset"] not in aid:
            errs.append(f"works {o['id']}: asset {o['asset']} が assets に無い")
        if o["attr"] not in ("INSPECT", "OVERHAUL", "REPLACE", "UPGRADE"):
            errs.append(f"works {o['id']}: attr")
        if o["status"] not in ("done", "planned", "postponed", "not_done"):
            errs.append(f"works {o['id']}: status")
        if o.get("manDays") is not None and not isinstance(o["manDays"], (int, float)):
            errs.append(f"works {o['id']}: manDays が数でない")
    recidx = {}
    for stem, a in ans.items():
        for rr in a["records"]:
            recidx[(a["file"], rr.get("r"), rr.get("seq"))] = rr
    for l in w["links"]:
        if l["work"] not in ids:
            errs.append(f"links: 存在しない作業 {l['work']}")
            break
        if (l["file"], l.get("r"), l.get("seq")) not in recidx:
            errs.append(f"links: 正解の記録に無い行 {l['file']} r={l.get('r')} seq={l.get('seq')}")
            break
        rr = recidx[(l["file"], l.get("r"), l.get("seq"))]
        if "_中長期整備計画_" in l["file"] and "year" not in l:
            errs.append(f"links: 星取表の結びに year が無い {l}")
            break
        if "_作業日報_" in l["file"] and "day" not in l and rr["kind"] == "item":
            errs.append(f"links: 日報の結びに day が無い {l}")
            break
        if l.get("how") not in ("direct", "annual", "bundle", "rollup", "upgrade"):
            errs.append(f"links: how {l.get('how')}")
            break
    for e in w["errata"]:
        if e["type"] not in ("omitMark", "lastYearUnreflected") or any(x not in ids for x in e["works"]):
            errs.append(f"errata: {e}")
        rr = recidx.get((e["file"], e["r"], None))
        if rr is None or str(e["year"]) in (rr["fields"].get("marks") or {}):
            errs.append(f"errata: 印が抜けていない／行が無い {e['file']} r={e['r']}")
    for stem, a in ans.items():
        for rr in a["records"]:
            f = rr["fields"]
            if a["plan"] and rr["kind"] in ("item", "parent"):
                if not isinstance(f.get("path"), list) or "marks" not in f or "plans" not in f:
                    errs.append(f"{a['file']}: 星取表の記録に path/marks/plans が無い r={rr.get('r')}")
                    break
            if "manDays" in f and not isinstance(f["manDays"], (int, float)):
                errs.append(f"{a['file']}: manDays が数でない r={rr.get('r')}")
                break
            if "days" in f and not isinstance(f["days"], (int, float)):
                errs.append(f"{a['file']}: days が数でない r={rr.get('r')}")
                break
            if (a["file"].endswith((".pdf", ".csv"))) and rr.get("sheet") is not None:
                errs.append(f"{a['file']}: PDF・CSV の記録に sheet がある")
                break
            if "_作業日報_" in a["file"] and rr["kind"] == "item" and "manDays" in f:
                # 人日の検算：1日1枚は 人員×時間÷8
                if "時間" in f and abs(f["manDays"] - f["人員"] * f["時間"] / 8) > 0.01:
                    errs.append(f"{a['file']}: manDays が人日でない r={rr.get('r')}")
                    break
    # files/ のハッシュ（変換の前に取った値と一致すること）
    hp = os.path.join(os.path.dirname(BASE), "work_files_sha256_before.txt")
    if os.path.exists(hp):
        import hashlib
        for line in open(hp, encoding="utf-8"):
            h, p_ = line.rstrip("\n").split("  ", 1)
            if hashlib.sha256(open(os.path.join(BASE, p_), "rb").read()).hexdigest() != h:
                errs.append(f"files/ が変わった：{p_}")
        stat["hash"] = 1


def main():
    ans = {os.path.basename(p)[:-5]: json.load(open(p, encoding="utf-8")) for p in glob.glob(os.path.join(ANS, "*.json"))}
    fs = sorted(os.listdir(FILES))
    for fn in fs:
        stem = os.path.splitext(fn)[0]
        if not fn.startswith("【架空】"):
            err(fn, "ファイル名が【架空】で始まらない")
        a = ans.get(stem)
        if not a:
            err(fn, "正解が無い")
            continue
        recs = a["records"]
        if fn.endswith(".xlsx"):
            text = check_xlsx(fn, recs)
        elif fn.endswith(".csv"):
            text = check_csv(fn, recs)
        elif fn.endswith(".pdf"):
            text = check_pdf(fn, recs)
        else:
            err(fn, "想定外の形式")
            continue
        for w in REAL:
            if w in text:
                err(fn, f"実在の名前らしい語: {w}")
        stat["files"] += 1
    imgs = [f for f in fs if f.lower().endswith((".jpg", ".jpeg", ".png"))]
    for f in imgs:
        from PIL import Image
        im = Image.open(os.path.join(FILES, f))
        if not (1568 <= max(im.size) <= 2000):
            err(f, f"画像の長辺 {max(im.size)}px")
    for stem in ans:
        if not any(os.path.splitext(f)[0] == stem for f in fs):
            err(stem, "正解はあるがファイルが無い")
    w = json.load(open(os.path.join(BASE, "answer", "world.json"), encoding="utf-8"))
    check_world(w, ans, fs)
    ids = {x["id"] for x in w["works"]}
    print(f"files={stat['files']} rows={stat['rows']} images={len(imgs)} works={len(ids)} links={len(w['links'])} hash_checked={stat['hash']}")
    if errs:
        print(f"NG {len(errs)} 件")
        for e in errs[:int(os.environ.get("NERR", "60"))]:
            print(" -", e)
        sys.exit(1)
    print("OK")


if __name__ == "__main__":
    main()
