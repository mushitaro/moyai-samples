# -*- coding: utf-8 -*-
"""機器台帳。
旧様式（市 i1、H17 作成、H27.4.1 現在）：
  シート「機器表（完成図書）」＝竣工時（H2.3）の完成図書の機器表を写したもの。1 行が 1 機器名、号炉の列が横に並ぶ
    （1号炉・2号炉・共通 の 数量／図書番号／市番号）。後で撤去した機器も行は残し、市番号の欄に「撤去」。
  シート「追加・更新」＝H5 以降に設置・更新した機器を、工事ごとに年度順に足したもの。1 行 1 台、旧番号（撤去機器の
    図書番号）と新番号を併記。
新様式（市 i2、H28.4 に区分を追加、R6.4.1 現在）：旧の 2 枚をコピーし、
  「機器表」の各号炉の市番号の右に「扱い」の列を差し込み、「追加・更新」の右端に「扱い」を足し、R2・R3 の行を i2 が追記。
  シート「予備品一覧」を新設（予備品で延命するための手持ち品）。"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from openpyxl import Workbook  # noqa: E402
from openpyxl.comments import Comment  # noqa: E402
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side  # noqa: E402
from openpyxl.utils import get_column_letter  # noqa: E402
from common import FAC_K, FONT, rng, set_props  # noqa: E402
from rec import Recs, path as out_path  # noqa: E402
import world  # noqa: E402
from world import SETSUBI, WORKNAME  # noqa: E402
from works import EQ, policy  # noqa: E402

thin, hair = Side(style="thin"), Side(style="hair")
B = Border(left=thin, right=thin, top=hair, bottom=hair)
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
F = Font(name=FONT, size=9)
FS = Font(name=FONT, size=8)
FH = Font(name=FONT, size=9, bold=True)
STRIKE = Font(name=FONT, size=9, strike=True, color="808080")
C = Alignment(horizontal="center", vertical="center", wrap_text=True)
L = Alignment(horizontal="left", vertical="center", wrap_text=True)
GRAY = PatternFill("solid", fgColor="D9D9D9")
BAND = PatternFill("solid", fgColor="DDEBF7")
HL = PatternFill("solid", fgColor="FFF2CC")
PFX = {"鷹羽エンジニアリング": "TFB", "霞台送風機": "KF", "碧水ポンプ工業": "HSP",
       "紫雲電機": "SE", "若葉計量器": "WTS", "黒潮クレーン": "KC", "千鳥破砕機": "CDS"}
LINE_NAME = {0: "共通", 1: "1号炉", 2: "2号炉"}
SHORT = {"停止まで使用": "使用", "予備品で延命": "延命", "部品取り（R3〜）": "部品取り",
         "停止まで使用（R2更新）": "使用(R2更新)"}


def model(e):
    r = rng("model" + e["base"])
    return f"{PFX[e['maker']]}-{r.randint(10, 99)}{r.choice(['', 'A', 'S', 'N'])}"


def ym(y):
    if y == 1989:
        return "H2.3"
    return f"H{y + 1 - 1988}.3"


def kw(spec):
    for tok in spec.split():
        if tok.endswith("kW"):
            return tok[:-2]
    return ""


def page(ws, last_col, last_row, title_rows, a3=True):
    ws.print_title_rows = title_rows
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A3 if a3 else ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_area = f"A1:{last_col}{last_row}"
    ws.oddFooter.center.text = "&A　&P / &N"


# ---------------------------------------------------------------- 機器表（完成図書）
AS = 2024          # 台帳の時点（年度。AS 年度の 4 月 1 日現在）


def main_rows():
    """完成図書に載っていた機器名ごとの行（竣工時の機器と、後に撤去した機器）。"""
    rows = {}
    order = []
    for e in world.tree_order(EQ):
        if e["year"] != 1989:
            continue
        key = (e["setsubi"], 0 if e["line"] == 0 else 1, e["keitou"], e["base"])
        if key not in rows:
            rows[key] = dict(setsubi=e["setsubi"], keitou=e["keitou"], name=e["base"], spec=e["spec"],
                             maker=e["maker"], model=model(e), cells={0: [], 1: [], 2: []}, removed=None)
            order.append(key)
        rows[key]["cells"][e["line"]].append(e)
    for rm in world.REMOVED:
        key = (rm["setsubi"], 0 if rm["line"] == 0 else 1, rm["keitou"], rm["base"])
        by = next(e for e in EQ if e["id"] == rm["by"])
        rm = dict(rm, id=rm["by"], duty=by["duty"])       # 撤去前は同じ市番号を使っていた
        if key not in rows:
            gone = rm["removed_fy"] < AS
            rows[key] = dict(setsubi=rm["setsubi"], keitou=rm["keitou"], name=rm["base"],
                             spec=by["spec"] if rm["base"] == by["base"] else "", maker=by["maker"] if rm["base"] == by["base"] else "鷹羽エンジニアリング",
                             model="" if gone else model(by), cells={0: [], 1: [], 2: []}, removed=rm["removed_fy"] if gone else None)
            # 撤去機器の行は、その系統の竣工時機器の並びの中に入れる
            pos = next((k for k, o in enumerate(order) if o[:3] == key[:3]), None)
            if pos is None:
                pos = next((k for k, o in enumerate(order) if o[:2] == key[:2]), len(order))
            order.insert(pos + sum(1 for o in order[pos:] if o[:3] == key[:3]), key)
        rows[key]["cells"][rm["line"]].append(rm)
    return [rows[k] for k in order]


def write_main(ws, with_pol, pol=None):
    ws["A1"] = f"{FAC_K} 機器台帳" + ("（停止まで）" if with_pol else "")
    ws["A1"].font = Font(name=FONT, size=14, bold=True)
    ws["A2"] = (f"{asof_text()}　（H17.4 i1 作成、H28.4 扱い区分追加 i2）" if with_pol else
                f"{asof_text()}　（H17.4 作成 環境施設課 i1。竣工図書 第3分冊「機器表」より転記）")
    ws["A2"].font = F
    if with_pol:
        ws["A1"].comment = Comment("旧表題：いずみ川清掃センター 機器台帳（H17.4 i1）", "i2")
    lines = (1, 2, 0)
    sub = ["数量", "図書番号", "市番号"] + (["扱い"] if with_pol else [])
    left = ["No", "設備", "系統", "機器名称", "型式", "仕様", "電動機\nkW", "メーカー"]
    lw = [4, 10, 9, 22, 8, 22, 6, 19]
    for j, (h, w) in enumerate(zip(left, lw), 1):
        ws.merge_cells(start_row=4, start_column=j, end_row=5, end_column=j)
        c = ws.cell(row=4, column=j, value=h)
        c.font, c.alignment, c.fill = FH, C, GRAY
        ws.column_dimensions[get_column_letter(j)].width = w
    col = len(left) + 1
    lcol = {}
    for ln in lines:
        lcol[ln] = col
        ws.merge_cells(start_row=4, start_column=col, end_row=4, end_column=col + len(sub) - 1)
        ws.cell(row=4, column=col, value=LINE_NAME[ln]).font = FH
        for k, h in enumerate(sub):
            c = ws.cell(row=5, column=col + k, value=h)
            c.font, c.alignment, c.fill = FH, C, GRAY
            ws.column_dimensions[get_column_letter(col + k)].width = [4.5, 8, 8, 8][k]
        col += len(sub)
    last = col + 1
    for h, w in (("設置", 6), ("備考", 34)):
        ws.merge_cells(start_row=4, start_column=col, end_row=5, end_column=col)
        c = ws.cell(row=4, column=col, value=h)
        c.font, c.alignment, c.fill = FH, C, GRAY
        ws.column_dimensions[get_column_letter(col)].width = w
        col += 1
    for rr in (4, 5):
        for j in range(1, last + 1):
            ws.cell(row=rr, column=j).border = BOX
            ws.cell(row=rr, column=j).alignment = C
            if not ws.cell(row=rr, column=j).fill.fgColor.rgb or rr == 4:
                ws.cell(row=rr, column=j).fill = GRAY
    r = 6
    no = 0
    prev_set, set_start = None, 6
    for row in main_rows():
        if row["setsubi"] != prev_set:
            if prev_set is not None and r - 1 > set_start:
                ws.merge_cells(start_row=set_start, start_column=2, end_row=r - 1, end_column=2)
            prev_set, set_start = row["setsubi"], r
        no += 1
        vals = [no, SETSUBI[row["setsubi"]], row["keitou"], row["name"], row["model"],
                row["spec"].replace(" 常用・予備", ""), kw(row["spec"]), row["maker"]]
        for j, v in enumerate(vals, 1):
            ws.cell(row=r, column=j, value=v)
        notes = []
        for ln in lines:
            items = row["cells"][ln]
            c0 = lcol[ln]
            if not items:
                continue
            ws.cell(row=r, column=c0, value=len(items))
            ws.cell(row=r, column=c0 + 1, value="\n".join(x["tag"] for x in items))
            if row["removed"]:
                ws.cell(row=r, column=c0 + 2, value=f"撤去\nH{row['removed'] + 1 - 1988}")
            else:
                ids = []
                for x in items:
                    star = "※" if with_pol and AS > 2020 and x["base"] == "計装用空気圧縮機" and x["name"].endswith("No.1") else ""
                    ids.append(x["id"] + star)
                ws.cell(row=r, column=c0 + 2, value="\n".join(ids))
                if with_pol:
                    ps = [SHORT[pol[x["id"]]] for x in items]
                    txt = ps[0] if len(set(ps)) == 1 else "\n".join(f"No.{k + 1} {p}" for k, p in enumerate(ps))
                    c = ws.cell(row=r, column=c0 + 3, value=txt)
                    if any(p != "使用" for p in ps):
                        c.fill = HL
                if any(x["duty"] for x in items) and "No.1常用・No.2予備" not in notes:
                    notes.append("No.1常用・No.2予備")
        if row["removed"]:
            by = [x["by"] for ln in lines for x in row["cells"][ln]]
            notes.append(f"H{row['removed'] + 1 - 1988} 撤去 →「追加・更新」{by[0]} ほか")
            for j in range(1, last + 1):
                ws.cell(row=r, column=j).font = STRIKE
        if with_pol and AS > 2020 and row["name"] == "計装用空気圧縮機":
            notes.append("※No.1 R2更新（追加・更新シート）")
        ws.cell(row=r, column=last - 1, value="H2.3" if no == 1 or row["removed"] else "〃")
        ws.cell(row=r, column=last, value="　".join(notes))
        nl = max(len(row["cells"][ln]) for ln in lines)
        ws.row_dimensions[r].height = max(24 if row["removed"] else 14, 12 * nl)
        for j in range(1, last + 1):
            c = ws.cell(row=r, column=j)
            if c.font != STRIKE:
                c.font = F
            c.border = B
            c.alignment = L if j in (3, 4, 6, 8, last) else C
        r += 1
    if r - 1 > set_start:
        ws.merge_cells(start_row=set_start, start_column=2, end_row=r - 1, end_column=2)
    for rr in range(6, r):
        ws.cell(row=rr, column=2).alignment = Alignment(vertical="top", wrap_text=True, text_rotation=255)
    ws.cell(row=r + 1, column=1, value="※ 竣工時の機器表（完成図書）による。H5以降の設置・更新は「追加・更新」シート。"
            "弁類・計装品は計装品台帳。盤は系統ごとに1台。").font = F
    if with_pol:
        ws.cell(row=r + 2, column=1, value="※ 扱い（H28追加）：使用＝停止まで使用、延命＝予備品で延命（予備品一覧）、部品取り＝予備機を部品取りに充当。").font = F
    ws.freeze_panes = "E6"
    page(ws, get_column_letter(last), r + 2, "4:5")


# ---------------------------------------------------------------- 追加・更新
def added_items():
    return sorted([e for e in EQ if e["year"] != 1989 and e["year"] < AS],
                  key=lambda e: (e["year"], list(SETSUBI).index(e["setsubi"]), e["line"], e["id"]))


def write_added(ws, with_pol, pol=None):
    ws["A1"] = f"{FAC_K} 機器台帳　追加・更新"
    ws["A1"].font = Font(name=FONT, size=13, bold=True)
    ws["A2"] = "H5以降に設置・更新した機器（工事ごとに追記）" + ("　R2・R3 は i2 追記" if with_pol else "")
    ws["A2"].font = F
    heads = ["No", "区分", "市番号", "図書番号", "機器名称", "号炉", "型式", "仕様", "メーカー", "製造番号",
             "設置年月", "旧番号（撤去）", "記入"] + (["扱い"] if with_pol else [])
    widths = [4, 6, 8, 8, 24, 6, 8, 24, 15, 10, 7, 12, 5] + ([10] if with_pol else [])
    for j, (h, w) in enumerate(zip(heads, widths), 1):
        c = ws.cell(row=4, column=j, value=h)
        c.font, c.alignment, c.border, c.fill = FH, C, BOX, GRAY
        ws.column_dimensions[get_column_letter(j)].width = w
    r, no, prev = 5, 0, None
    for e in added_items():
        if e["year"] != prev:
            prev = e["year"]
            ws.cell(row=r, column=2, value=WORKNAME.get(e["year"], f"H{e['year'] - 1988}"))
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=len(heads))
            for j in range(1, len(heads) + 1):
                ws.cell(row=r, column=j).fill = BAND
                ws.cell(row=r, column=j).border = BOX
                ws.cell(row=r, column=j).font = FH
            r += 1
        no += 1
        vals = [no, "更新" if e.get("old_tag") else "新設", e["id"], e["tag"] or "—", e["name"],
                LINE_NAME[e["line"]], model(e), e["spec"].replace(" 常用・予備", ""), e["maker"], e["serial"],
                ym(e["year"]), e.get("old_tag", ""), "i1"]
        if with_pol:
            vals.append(SHORT[pol[e["id"]]])
        for j, v in enumerate(vals, 1):
            c = ws.cell(row=r, column=j, value=v)
            c.font, c.border = F, B
            c.alignment = L if j in (5, 8, 9) else C
        if with_pol and vals[-1] != "使用":
            ws.cell(row=r, column=len(heads)).fill = HL
        r += 1
    if with_pol and AS > 2020:
        ws.cell(row=r, column=2, value="令和2〜3年度（i2 追記）")
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=len(heads))
        for j in range(1, len(heads) + 1):
            ws.cell(row=r, column=j).fill = BAND
            ws.cell(row=r, column=j).border = BOX
            ws.cell(row=r, column=j).font = FH
        r += 1
        add = [[no + 1, "更新", "0M-001", "—", "計装用空気圧縮機 No.1", "共通", "SE-R22", "22kW スクリュー式",
                "紫雲電機", "R2-3318", "R2.11", "I-001（旧機）", "i2", "使用(R2更新)"],
               [no + 2, "転用", "0A-019→0A-018", "—", "前処理破砕機 減速機（No.2→No.1）",
                "共通", "", "R3 前処理破砕機補修工事", "千鳥破砕機", "", "R3.12", "", "i2", "部品取り"]]
        if AS <= 2021:
            add = add[:1]
        for row in add:
            for j, v in enumerate(row, 1):
                c = ws.cell(row=r, column=j, value=v)
                c.font, c.border = F, B
                c.alignment = L if j in (5, 8, 9) else C
            ws.cell(row=r, column=len(heads)).fill = HL
            r += 1
    ws.freeze_panes = "C5"
    page(ws, get_column_letter(len(heads)), r, "4:4")
    return r


# ---------------------------------------------------------------- 予備品一覧（i2、H28〜）
SPARES = [
    ("散気ノズル（TFB-75用）", "1B-002・2B-002", 24, "R4購入 120本、R5 2号炉で96本使用", "R4"),
    ("ろ布（PTFE 480本組の予備）", "1E-007・2E-007", 16, "H28購入 48本、R4 1号炉で32本使用", "H28"),
    ("CPU基板", "0M-006", 1, "H28購入 2枚（生産中止前の最終手配）、R2 1枚使用", "H28"),
    ("I/O基板", "0M-006", 3, "H28購入", "H28"),
    ("給じんスクリュー羽根（肉盛済）", "1A-004・1A-005・2A-004・2A-005", 1, "1式。H29 メーカー定修で取外し品を再生", "H29"),
    ("インバータユニット（160kW）", "1F-006・2F-006", 1, "H28購入", "H28"),
    ("クレーン用インバータ", "0A-017", 1, "H25更新時の予備", "H25"),
    ("混練機パドル", "0K-005・0K-006", 1, "1式", "H28"),
    ("前処理破砕機 No.2（部品取り）", "0A-019", 1, "R3 減速機をNo.1へ転用。刃1式残", "R3"),
]


def spares():
    q = {"散気ノズル（TFB-75用）": (24 if AS > 2023 else None), "ろ布（PTFE 480本組の予備）": (16 if AS > 2022 else 48),
         "CPU基板": (1 if AS > 2020 else 2), "給じんスクリュー羽根（肉盛済）": (1 if AS > 2017 else None),
         "前処理破砕機 No.2（部品取り）": (1 if AS > 2021 else None)}
    why = {"散気ノズル（TFB-75用）": "R4購入 120本、R5 2号炉で96本使用",
           "ろ布（PTFE 480本組の予備）": "H28購入 48本" + ("、R4 1号炉で32本使用" if AS > 2022 else ""),
           "CPU基板": "H28購入 2枚（生産中止前の最終手配）" + ("、R2 1枚使用" if AS > 2020 else "")}
    out = []
    for (nm, tgt, n, w, fy) in SPARES:
        n2 = q.get(nm, n)
        if n2 is None:
            continue
        out.append((nm, tgt, n2, why.get(nm, w), fy))
    return out


def asof_text():
    if AS >= 2019:
        n = AS - 2018
        return f"令和{'元' if n == 1 else n}年4月1日現在"
    return f"平成{AS - 1988}年4月1日現在"


def write_spares(ws):
    ws["A1"] = f"{FAC_K} 予備品一覧（停止まで延命用）"
    ws["A1"].font = Font(name=FONT, size=13, bold=True)
    ws["A2"] = "H28.4 作成 i2　" + asof_text()
    ws["A2"].font = F
    heads = ["No", "品名", "対象機器（市番号）", "数量", "経緯", "手配年度", "保管"]
    for j, (h, w) in enumerate(zip(heads, [4, 28, 26, 6, 40, 8, 10]), 1):
        c = ws.cell(row=4, column=j, value=h)
        c.font, c.alignment, c.border, c.fill = FH, C, BOX, GRAY
        ws.column_dimensions[get_column_letter(j)].width = w
    SP = spares()
    for k, (nm, tgt, q, why, fy) in enumerate(SP, 1):
        for j, v in enumerate([k, nm, tgt, q, why, fy, "予備品倉庫" if "部品取り" not in nm else "破砕機室"], 1):
            c = ws.cell(row=4 + k, column=j, value=v)
            c.font, c.border = F, BOX
            c.alignment = L if j in (2, 3, 5) else C
    page(ws, "G", 4 + len(SP), "4:4", a3=False)


def harvest(R, wb):
    """書いたシートから正解の行を拾う。"""
    ws = wb["機器表（完成図書）"]
    heads = [ws.cell(row=5, column=j).value or ws.cell(row=4, column=j).value for j in range(1, ws.max_column + 1)]
    groups = [ws.cell(row=4, column=j).value for j in range(1, ws.max_column + 1)]
    cur = None
    for j in range(len(groups)):
        if groups[j] in ("1号炉", "2号炉", "共通"):
            cur = groups[j]
        elif ws.cell(row=5, column=j + 1).value is None:
            cur = None
        groups[j] = cur
    for r in range(6, ws.max_row + 1):
        no = ws.cell(row=r, column=1).value
        if not isinstance(no, int):
            continue
        f, lines = {}, {}
        for j in range(1, ws.max_column + 1):
            v = ws.cell(row=r, column=j).value
            if v is None or j == 1:
                continue
            if groups[j - 1]:
                lines.setdefault(groups[j - 1], {})[heads[j - 1]] = v
            else:
                f[heads[j - 1].replace("\n", "")] = v
        if "設備" not in f:
            f["設備"] = ws.cell(row=r, column=2).value
        f["lines"] = lines
        f["no"] = no
        f["struck"] = bool(ws.cell(row=r, column=4).font.strike)
        ids = [x.rstrip("※") for g in lines.values() for x in str(g.get("市番号", "")).split("\n") if x[:1] in "012"]
        R.add(ws.title, r, "item", f)          # 台帳の行は機器（作業ではない）。市番号は fields にある
    ws = wb["追加・更新"]
    heads = [ws.cell(row=4, column=j).value for j in range(1, ws.max_column + 1)]
    for r in range(5, ws.max_row + 1):
        if not isinstance(ws.cell(row=r, column=1).value, int):
            continue
        f = {h: ws.cell(row=r, column=j).value for j, h in enumerate(heads, 1) if ws.cell(row=r, column=j).value not in (None, "")}
        R.add(ws.title, r, "item", f)
    if "予備品一覧" in wb.sheetnames:
        ws = wb["予備品一覧"]
        heads = [ws.cell(row=4, column=j).value for j in range(1, ws.max_column + 1)]
        for r in range(5, ws.max_row + 1):
            if ws.cell(row=r, column=1).value is None:
                continue
            R.add(ws.title, r, "item", {h: ws.cell(row=r, column=j).value for j, h in enumerate(heads, 1)})


def ledger(asof):
    """asof＝年度（その年度の 4 月 1 日現在）。2016 より前は旧様式、以後は新様式。"""
    global AS
    AS = asof
    new = asof >= 2017
    pol = dict(policy())
    if asof <= 2021:
        pol["0A-019"] = "停止まで使用"
    if asof <= 2020:
        pol["0M-001"] = "停止まで使用"
    wb = Workbook()
    ws = wb.active
    ws.title = "機器表（完成図書）"
    write_main(ws, new, pol)
    write_added(wb.create_sheet("追加・更新"), new, pol)
    if new:
        write_spares(wb.create_sheet("予備品一覧"))
    set_props(wb, f"{FAC_K} 機器台帳 {asof_text()}")
    wa = asof_text().replace("4月1日現在", "")
    fn = (f"{FAC_K}_機器台帳_{wa}4月時点_新様式（市i2・扱い列と予備品一覧を追加）.xlsx" if new else
          f"{FAC_K}_機器台帳_{wa}4月時点_旧様式（市i1・完成図書機器表＋追加更新シート）.xlsx")
    wb.save(out_path(fn))
    R = Recs(fn)
    import openpyxl
    harvest(R, openpyxl.load_workbook(out_path(fn)))
    R.save()
    return fn


def count_check(path):
    """台帳から管理対象の台数を数え直す（撤去行は数えない）。"""
    import openpyxl
    wb = openpyxl.load_workbook(path)
    ids = set()
    ws = wb["機器表（完成図書）"]
    for row in ws.iter_rows(min_row=6, values_only=True):
        for v in row:
            if isinstance(v, str) and len(v) >= 6 and v[1:2].isalpha() and v[0] in "012" and v[2] == "-":
                for t in v.split("\n"):
                    ids.add(t.rstrip("※"))
    ws = wb["追加・更新"]
    for row in ws.iter_rows(min_row=5, values_only=True):
        if row[1] in ("新設", "更新") and row[2] and "→" not in str(row[2]):
            ids.add(row[2])
    return len(ids)


if __name__ == "__main__":
    for y in (2005, 2010, 2015, 2017, 2021, 2024):
        f = ledger(y)
        print(f, count_check(out_path(f)))
