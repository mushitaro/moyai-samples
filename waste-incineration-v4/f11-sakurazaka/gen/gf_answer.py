# -*- coding: utf-8 -*-
"""正解の行を書き出す（台帳・計画は出来上がった xlsx を読み直して行を拾う）"""
import re, os, json
import openpyxl
import world as W

YEARS = list(range(2020, 2028))

def norm(s):
    return re.sub(r"[\s　]", "", str(s)) if s is not None else ""

def comments_of(ws, r, maxc=40):
    out = {}
    for c in ws[r][:maxc]:
        if c.comment:
            out[c.coordinate] = c.comment.text
    return out

def ledger_records(fn):
    wb = openpyxl.load_workbook(fn)
    recs = []
    ws = wb["集計"]
    for row in ws.iter_rows(min_row=11):
        a = row[0].value
        if not a:
            continue
        vals = [c.value for c in row[1:9]]
        if a == "合計":
            recs.append(dict(sheet="集計", r=row[0].row, kind="total", fields={"設備": "合計", "計行数": vals[6], "計台数": vals[7]}))
        elif a in W.CAT.values():
            recs.append(dict(sheet="集計", r=row[0].row, kind="subtotal", fields={"設備": a, "1号炉欄数": vals[0], "1号炉台数": vals[1], "2号炉欄数": vals[2], "2号炉台数": vals[3],
                                                                               "共通欄数": vals[4], "共通台数": vals[5], "計欄数": vals[6], "計台数": vals[7]}))
    ws = wb["機器仕様台帳"]
    cur = None
    catname = None
    for row in ws.iter_rows(min_row=4):
        a = row[0].value
        r = row[0].row
        if isinstance(a, str) and re.match(r"^\d+\.　", a):
            catname = a.split("　", 1)[1]
            recs.append(dict(sheet=ws.title, r=r, kind="parent", fields={"path": [catname], "設備": catname}))
            cur = None
            continue
        m = re.match(r"^【(\w+-\d+)】　(.+)　（(.+)）$", a) if isinstance(a, str) else None
        if m:
            cur = dict(sheet=ws.title, r=r, kind="item", fields={"path": [catname, m.group(2)], "台帳No": m.group(1), "機器名称": m.group(2), "号炉": m.group(3),
                                                                 "図書番号": str(row[6].value).replace("図書 ", "")})
            recs.append(cur)
            continue
        if cur and r - cur["r"] <= 5:
            for li, vi in ((0, 1), (3, 4), (7, 8)):
                if len(row) > vi and row[li].value:
                    lab = norm(row[li].value)
                    v = row[vi].value
                    cur["fields"][lab] = v
                    if lab == "数量" and isinstance(v, str):
                        mm = re.match(r"(\d+) 台", v)
                        if mm:
                            cur["fields"]["台数"] = int(mm.group(1))
            cm = comments_of(ws, r)
            if cm:
                cur["fields"].setdefault("comments", {}).update(cm)
        if cur and isinstance(a, str) and a.startswith("【"):
            pass
    for rr in recs:
        if rr["kind"] == "item":
            cm = comments_of(ws, rr["r"])
            if cm:
                rr["fields"].setdefault("comments", {}).update(cm)
    if "変更履歴" in wb.sheetnames:
        ws = wb["変更履歴"]
        for row in ws.iter_rows(min_row=4):
            if row[0].value:
                recs.append(dict(sheet="変更履歴", r=row[0].row, kind="item", fields={"日付": row[0].value, "担当": row[1].value, "内容": row[2].value}))
    return recs


def plan_records(fn):
    wb = openpyxl.load_workbook(fn)
    ws = wb["整備計画"]
    recs = []
    # 見出しの行と列
    hdr = None
    for r in range(1, 10):
        if ws.cell(r, 1).value == "項番":
            hdr = r
            break
    heads = {norm(ws.cell(hdr, c).value): c for c in range(1, ws.max_column + 1) if ws.cell(hdr, c).value}
    ycol = {heads["R%d" % (y - 2018)]: y for y in YEARS}
    cB = heads.get("備考"); cQ = heads.get("部品供給"); cR = [c for k, c in heads.items() if k.startswith("改訂メモ")]
    cR = cR[0] if cR else None
    cat = None; name = None; sec = None
    for r in range(hdr + 1, ws.max_row + 1):
        a = ws.cell(r, 1).value; b = ws.cell(r, 2).value; e = ws.cell(r, 5).value; g = ws.cell(r, 7).value
        if isinstance(a, str) and re.match(r"^\d+　", a) and not e:
            cat = a.split("　", 1)[1]; sec = "eq" if not cat.startswith("法定") else "legal"
            recs.append(dict(sheet=ws.title, r=r, kind="parent", fields={"path": [cat]}))
            continue
        if isinstance(a, str) and a.startswith("年度別"):
            sec = "cost"; cat = a.split("（")[0]; cname = None
            continue
        marks = {str(ycol[c]): ws.cell(r, c).value for c in ycol if ws.cell(r, c).value not in (None, "")}
        memo = "／".join(str(x) for x in [ws.cell(r, cB).value if cB else None, ws.cell(r, cR).value if cR else None] if x)
        cm = comments_of(ws, r, ws.max_column)
        if sec == "eq" and (e or (isinstance(b, str) and b.strip().startswith("└"))):
            if b and not str(b).strip().startswith("└"):
                name = b
            unit = ws.cell(r, 3).value
            path = [cat, name, str(unit)]
            work = ws.cell(r, 6).value
            if isinstance(b, str) and b.strip().startswith("└"):
                path.append(b.strip().lstrip("└").strip())
                work = "部品手当て"
            cyc = {"毎年点検": 1, "隔年点検": 2}.get(work)
            f = {"path": path, "work": work, "code": e or "", "cycle": cyc, "cost": None, "unit": None, "marks": marks, "memo": memo, "comments": cm,
                 "台数": ws.cell(r, 4).value, "前回整備": ws.cell(r, 7).value}
            if cQ:
                f["部品供給"] = ws.cell(r, cQ).value
            recs.append(dict(sheet=ws.title, r=r, kind="item", fields=f))
        elif sec == "legal" and isinstance(a, str) and a.startswith("13-"):
            cyc_t = ws.cell(r, 6).value
            cyc = {"1年": 1, "2年": 2, "3年": 3, "6か月": 0.5, "1月": round(1 / 12, 3), "年2回以上": 0.5}.get(cyc_t)
            recs.append(dict(sheet=ws.title, r=r, kind="item", fields={"path": [cat, b], "work": "法定検査", "code": "", "cycle": cyc, "cycleText": cyc_t, "cost": None, "unit": None,
                                                                         "marks": marks, "memo": memo, "comments": cm}))
        elif sec == "cost" and (a or g in ("予算", "決算")):
            if a and a not in ("計", "予算 計", "決算 計"):
                cname = a
            kind = "total" if a in ("計", "予算 計", "決算 計") else "item"
            lab = g if g in ("予算", "決算") else None
            path = [cat, cname] + ([lab] if lab else []) if kind == "item" else [cat, a]
            vals = {str(ycol[c]): ws.cell(r, c).value for c in ycol if isinstance(ws.cell(r, c).value, (int, float))}
            recs.append(dict(sheet=ws.title, r=r, kind=kind, fields={"path": path, "work": "費用", "code": "", "cycle": None, "cost": vals, "unit": "千円", "marks": {}, "memo": memo, "comments": cm}))
    return recs


def dump(fn, recs, plan=False, extra=None):
    os.makedirs(os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "answer", "files"), exist_ok=True)
    stem = os.path.splitext(os.path.basename(fn))[0]
    out = {"file": os.path.basename(fn), "plan": plan, "records": recs}
    if extra:
        out.update(extra)
    p = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "answer", "files", stem + ".json")
    with open(p, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=1, default=str)
    return p
