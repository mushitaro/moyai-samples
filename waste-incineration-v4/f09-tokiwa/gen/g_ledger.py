# -*- coding: utf-8 -*-
"""機器台帳：SPC 保全システム台帳（2023年度末）→ 組合様式へ変換した台帳（令和6年度引渡し）"""
import datetime as dt
from openpyxl import Workbook
from openpyxl.comments import Comment
import world as W
from xlh import *
from openpyxl.styles import Alignment
from plan_gen import GROUP_ORDER, U_GROUPS, u_name, KIKAN
from rec import Rec
import contracts as C

_F = "アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワヲンァィゥェォャュョッー"
_H = "ｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖﾗﾘﾙﾚﾛﾜｦﾝｧｨｩｪｫｬｭｮｯｰ"
_D = {"ガ": "ｶﾞ", "ギ": "ｷﾞ", "グ": "ｸﾞ", "ゲ": "ｹﾞ", "ゴ": "ｺﾞ", "ザ": "ｻﾞ", "ジ": "ｼﾞ", "ズ": "ｽﾞ", "ゼ": "ｾﾞ", "ゾ": "ｿﾞ",
      "ダ": "ﾀﾞ", "ヂ": "ﾁﾞ", "ヅ": "ﾂﾞ", "デ": "ﾃﾞ", "ド": "ﾄﾞ", "バ": "ﾊﾞ", "ビ": "ﾋﾞ", "ブ": "ﾌﾞ", "ベ": "ﾍﾞ", "ボ": "ﾎﾞ",
      "パ": "ﾊﾟ", "ピ": "ﾋﾟ", "プ": "ﾌﾟ", "ペ": "ﾍﾟ", "ポ": "ﾎﾟ", "ヴ": "ｳﾞ"}

def hankaku(s):
    out = []
    for ch in s:
        if ch in _D: out.append(_D[ch])
        elif ch in _F: out.append(_H[_F.index(ch)])
        else: out.append(ch)
    return "".join(out)

# SPC 台帳で sv1 が半角カナで入れた系統（保全システム初期登録時の癖）
HALF_SYS = ("410", "220", "520")

def spc_name(e):
    n = e.name
    if e.sys in HALF_SYS:
        n = hankaku(n)
    return n

def maint_class(e):
    ins, ovh, life = W.POLICY.get(e.cls, W.DEFAULT_POLICY)
    if ins: return "TBM", f"{ins}Y"
    if ovh: return "TBM", f"{ovh}Y"
    if e.crit == "C": return "BDM", ""
    return "CBM", ""

import re
ACC = {"crane": ["巻上装置", "横行装置", "走行装置", "バケット（油圧開閉式）", "給電装置"],
       "crane_t": ["主巻装置", "補巻装置", "走行装置"], "crane_s": ["巻上装置"],
       "shredder": ["減速機", "カッタ軸（2 軸）", "スクリーン"],
       "idf": ["軸受箱（2）", "軸継手", "潤滑油ユニット"], "fan_l": ["軸受箱（2）", "軸継手", "吸込消音器"],
       "bf": ["パルス弁ヘッダ（6 室）", "ろ布（252 本・消耗品）"], "gasifier": ["散気ノズル（96 本・消耗品）", "炉内耐火物"],
       "melter": ["炉内耐火物（側壁・炉頂・出滓口）"], "chamber": ["耐火物"], "boiler": ["蒸発管群", "水冷壁"],
       "feeder": ["減速機"], "screw": ["減速機"], "nfscrew": ["減速機", "水冷ジャケット"], "chainconv": ["減速機", "チェーン・スクレーパ"],
       "elev": ["減速機", "バケット・チェーン"], "turbine": ["車室・ロータ", "調速装置"], "gen": ["固定子", "回転子"],
       "scr": ["触媒（2 層＋予備 1 層）"], "kneader": ["減速機"], "comp": ["アフタークーラ"], "egen": ["始動用空気槽", "ラジエータ"]}

def motor_of(e):
    ks = re.findall(r"([0-9.]+)kW", e.spec)
    if not ks or e.cls in ("inv", "heater", "mcc", "egen", "turbine", "gen"):
        return []
    out = []
    for k in ks:
        kw = float(k)
        v = "6.6kV" if kw >= 100 else ("400V" if kw >= 0.75 else "200V")
        p = "2P" if e.cls in ("bfp",) else ("6P" if e.cls in ("idf", "fan_l", "shredder") else "4P")
        out.append(f"{k}kW {p} {v} 全閉外扇" + (" INV" if "INV" in e.spec or e.cls in ("idf", "fan_l", "feeder") else ""))
    return out

def spc_ledger(V, path, fname):
    rec = Rec(fname)
    """SPC 保全システムの設備ツリー出力（系統 → 設備 → 機器 → 付属機器 → 電動機）。台帳対象はタグの付いた機器（階層 3）だけ"""
    EQV = [e for e in W.EQS if not (e.base == "排ガス水銀連続分析計" and V < 2018)]
    wb = Workbook(); ws = wb.active; ws.title = "設備ツリー"
    cols = ["階層", "設備コード", "名称", "号炉", "常用/予備", "形式", "仕様", "製造者", "製造番号", "設置年月", "更新年月",
            "重要度", "保全区分", "周期", "資産番号", "台帳対象", "備考"]
    merge(ws, 1, 1, 1, len(cols), f"{W.FAC}　設備台帳（保全システム 設備ツリー出力）", f(14, True), al("left"), None)
    merge(ws, 2, 1, 2, len(cols), f"{W.SPC}　保全グループ　{V + 1}/3/{29 if V % 2 else 28} 出力（{V}年度末）　出力者 {"sv1" if V <= 2016 else "sv2"}　台帳対象 {len(EQV)} 件（階層 3）", f(9), al("right"), None)
    for i, h in enumerate(cols):
        put(ws, 4, i + 1, h, f(9, True), al("center"), B_ALL, fill("D9E1F2"))
    r = 5
    order = sorted(EQV, key=lambda e: (e.furn if e.furn else 9, GROUP_ORDER.index(e.group), e.sys, e.tag))
    pg = ps = None
    n_motor = n_acc = 0
    LV = {1: fill("BDD7EE"), 2: fill("DDEBF7")}
    COLN = ["号炉", "常用/予備", "形式", "仕様", "製造者", "製造番号", "設置年月", "更新年月", "重要度", "保全区分", "周期", "資産番号", "台帳対象", "備考"]
    pth = {}
    def row(level, code, name, vals, bold=False):
        nonlocal r
        pth[level] = name
        for k in list(pth):
            if k > level: del pth[k]
        fl = {COLN[i]: v for i, v in enumerate(vals) if v not in (None, "")}
        rec.add("parent" if level <= 2 else "item", {"level": level, "code": code, "name": name, "path": [pth[k] for k in sorted(pth)],
                                                        "ledgerTarget": level == 3, **fl}, ws.title, r)
        put(ws, r, 1, level, f(8), al("center"), B_HAIR, LV.get(level))
        put(ws, r, 2, code, f(8), al("left"), B_HAIR, LV.get(level))
        c = put(ws, r, 3, name, f(8, bold), Alignment(horizontal="left", vertical="center", indent=(level - 1) * 2), B_HAIR, LV.get(level))
        for i, v in enumerate(vals):
            put(ws, r, 4 + i, v, f(8), al("left" if i in (3, 13) else "center", shrink=i == 3), B_HAIR, LV.get(level))
        ws.row_dimensions[r].outline_level = max(0, level - 1)
        r += 1
    for e in order:
        fl = {0: "共通", 1: "1号炉", 2: "2号炉"}[e.furn]
        if (e.furn, e.group) != pg:
            row(1, f"{e.furn}-G{GROUP_ORDER.index(e.group) + 1:02d}", f"{fl}　{e.group}", [None] * 14, True); pg = (e.furn, e.group); ps = None
        if e.sys + e.sub != ps:
            row(2, f"{e.furn}-{e.sys}", e.sub, [None] * 14, True); ps = e.sys + e.sub
        rep = W.replaced_before(e, V + 1)
        mc, cyc = maint_class(e)
        note = ""
        if e.tag == W.find(2, "過熱器").tag:
            note = "2015 減肉補修、2019 部分取替" if V >= 2019 else ("2015 減肉補修" if V >= 2015 else "")
        row(3, e.tag, spc_name(e), [{0: "共通", 1: "1", 2: "2"}[e.furn], e.duty, e.model, e.spec, e.maker.replace("【架空】", ""), e.serial,
                                    e.install.replace("-", "/"), f"{rep[-1] + 1}/03" if rep else None, e.crit, mc, cyc, e.spc_no, "○", note])
        for j, a in enumerate(ACC.get(e.cls, []), 1):
            n_acc += 1
            row(4, f"{e.tag}-a{j}", a, [None] * 11 + [None, None, ""])
        for j, m in enumerate(motor_of(e), 1):
            n_motor += 1
            row(5, f"{e.tag}-M{j}", "電動機" if j == 1 else f"電動機（{j}）", [None, None, f"TM-{e.idx % 7 + 1}", m, "【架空】東凪電機".replace("【架空】", ""),
                                                                       None, e.install.replace("-", "/"), None, None, None, None, None, None, "機器に付属"])
    widths(ws, {"A": 4.5, "B": 12, "C": 30, "D": 5, "E": 7, "F": 11, "G": 30, "H": 14, "I": 9, "J": 8, "K": 8, "L": 5, "M": 6, "N": 5, "O": 9, "P": 5, "Q": 22})
    ws.sheet_properties.outlinePr.summaryBelow = False
    ws.auto_filter.ref = f"A4:Q{r - 1}"
    ws.freeze_panes = "D5"
    page(ws, "landscape", 9, 1, 0, f"A1:{'Q' if V >= 2026 else ('P' if V == 2024 else 'O')}{r - 1}", "4:4")
    ws.oddFooter.center.text = "&P / &N"
    props(wb, f"{W.FAC} 設備台帳（{V}年度末・設備ツリー）", f"{'sv1' if V <= 2016 else 'sv2'}（【架空】ときわ北部エコサービス 保全グループ）", dt.datetime(V + 1, 3, 28, 17, 0))
    wb.save(path); rec.save()
    return n_acc, n_motor

STD_LIFE = {"pump": 15, "bfp": 15, "fan_l": 20, "idf": 20, "fan_m": 15, "fan_s": 15, "crane": 20, "shredder": 15,
            "mvalve": 15, "cvalve": 15, "analyzer": 10, "mcc": 20, "lp": 20, "inv": 15, "dcsc": 15, "dcsop": 7, "dcsio": 15,
            "conv": 15, "conv_d": 15, "turbine": 25, "gen": 25, "boiler": 25, "sh": 15, "eco": 15, "bf": 20, "scr": 20,
            "hv": 25, "lv": 25, "tr": 30, "tr_s": 30, "ups": 7, "dc": 10, "egen": 30}

SUBMAKER_U = {"idf": "峰岡送風機", "fan_l": "峰岡送風機", "pump": "八潮ポンプ"}

def union_ledger(V, path, fname):
    rec = Rec(fname)
    RW = f"令和{V - 2018}年4月1日"
    wb = Workbook()
    cols = ["管理番号", "大分類", "中分類", "機器名称", "号炉", "仕様", "製造者", "設置年度", "更新年度", "経過年数",
            "標準耐用年数", "重要度", "旧資産番号", "タグNo.", "備考", "（SPC）保全区分", "旧管理番号"]
    first = True
    total = 0
    R7NEW = {"溶融炉バーナ", "過熱器", "エコノマイザ", "誘引通風機", "誘引通風機インバータ盤", "流動空気押込送風機",
             "押込送風機インバータ盤", "溶融炉燃焼用送風機", "炉体冷却水ポンプ"}
    hist = []
    for gi, (sheet, groups) in enumerate(U_GROUPS):
        ws = wb.active if first else wb.create_sheet()
        first = False
        ws.title = sheet
        merge(ws, 1, 1, 1, 15, f"{W.FAC}　機器台帳　【{sheet}設備】", f(14, True), al("left"), None)
        merge(ws, 2, 1, 2, 15, f"{W.UNION}　施設課　{RW}現在（{'運営事業者台帳 2024/3/29 出力より変換' if V == 2024 else 'R6.4.1 版は運営事業者台帳 2024/3/29 出力より変換、以後毎年度更新'}）　u1", f(9), al("right"), None)
        for i, h in enumerate(cols):
            if i == 16 and V < 2026: continue
            put(ws, 4, i + 1, h, f(9, True), al("center", wrap=True), B_ALL, fill("E2EFDA" if i < 15 else ("D9D9D9" if i == 15 else "FFF2CC")))
        ws.row_dimensions[4].height = 26
        r = 5; n = 0
        eqs = [e for e in W.EQS if e.group in groups]
        eqs.sort(key=lambda e: (groups.index(e.group), e.sys, e.furn if e.furn else 9, e.idx))
        prev_sub = None
        nnew = len(eqs)
        for e in eqs:
            n += 1; total += 1
            rep = W.replaced_before(e, V)
            inst_fy = W.start_fy(e) - 1
            last = rep[-1] if rep else inst_fy
            nm = u_name(e)
            # 変換で半角カナが残った行（SPC の初期登録分のうち、変換表に無かった語）
            if e.sys in ("410",) and e.base in ("スートブロワ",):
                nm = spc_name(e)
            crit = {"A": "Ⅰ", "B": "Ⅱ", "C": "Ⅲ"}[e.crit]
            if e.sys == "540":
                crit = e.crit            # 変換漏れ（A/B/C のまま）
            mc, cyc = maint_class(e)
            sub = e.sub if e.sub != prev_sub else "〃"
            prev_sub = e.sub
            vals = [f"K{gi + 1:02d}-{n:04d}", groups[0][:-2] if len(groups) == 1 else e.group[:-2], sub, nm,
                    {0: "共通", 1: "1号", 2: "2号"}[e.furn], e.spec, e.maker.replace("【架空】", ""),
                    wareki_fy(inst_fy).replace("年度", ""), wareki_fy(rep[-1]).replace("年度", "") if rep else None,
                    V - last - 1 if (last == 2025 and V == 2026) else V - last, STD_LIFE.get(e.cls, 20), crit, e.spc_no, e.tag,
                    ("基幹改良 対象" if e.tag in KIKAN else "") + (f"（{e.duty}）" if e.duty in ("常用", "予備") else ""),
                    f"{mc} {cyc}".strip(), None]
            old_no = vals[0]
            if V >= 2026 and e.furn == 2 and e.tag in KIKAN and KIKAN[e.tag][1] == 2025:
                if e.base in R7NEW:
                    nnew += 1
                    vals[0] = f"K{gi + 1:02d}-{nnew:04d}"
                    vals[16] = old_no
                    vals[7] = "令和7"; vals[8] = None; vals[9] = 0
                    vals[6] = "鷹羽エンジニアリング" if e.cls not in SUBMAKER_U else SUBMAKER_U[e.cls]
                    vals[14] = "R7 基幹改良で更新（R8.1.16 引渡し）"
                    hist.append((old_no, vals[0], sheet, nm, {0: "共通", 1: "1号", 2: "2号"}[e.furn], KIKAN[e.tag][3]))
                else:
                    vals[14] = "R7 基幹改良で改良（" + KIKAN[e.tag][3][:14] + "）"
                    vals[8] = "令和7"
            COLU = ["管理番号", "大分類", "中分類", "機器名称", "号炉", "仕様", "製造者", "設置年度", "更新年度", "経過年数", "標準耐用年数", "重要度", "旧資産番号", "タグNo.", "備考", "（SPC）保全区分", "旧管理番号"]
            rec.add("item", {COLU[i]: v for i, v in enumerate(vals) if v not in (None, "") and not (i == 16 and V < 2026)}, ws.title, r)
            for i, v in enumerate(vals):
                if i == 16 and V < 2026: continue
                c = put(ws, r, i + 1, v, f(8), al("left" if i in (3, 5, 14) else "center", shrink=i in (5,)), B_ALL)
                if i == 15: c.font = f(8, color="808080")
                if i == 16 and v: c.fill = fill("FFF2CC")
            if vals[9] >= STD_LIFE.get(e.cls, 20):
                ws.cell(row=r, column=10).fill = fill("FCE4D6")
            r += 1
        widths(ws, {"A": 10, "B": 9, "C": 12, "D": 28, "E": 5, "F": 30, "G": 14, "H": 7, "I": 7, "J": 5, "K": 6, "L": 5, "M": 9, "N": 9, "O": 22, "P": 9, "Q": 10})
        if V >= 2025: ws.column_dimensions["P"].hidden = True   # SPC 由来の列。R7.4.1 版から隠した
        ws.cell(row=4, column=10).comment = Comment(f"{RW}時点。更新年度があればそこから数える" + ("（R7 更新分は 0）。" if V >= 2026 else "。"), "u1")
        ws.freeze_panes = "E5"
        page(ws, "landscape", 9, 1, 0, f"A1:{'Q' if V >= 2026 else ('P' if V == 2024 else 'O')}{r - 1}", "4:4")
        ws.oddFooter.center.text = f"{sheet}  &P / &N"
    # 後から足したシート：更新・除却の記録（R8.4.1 版から）
    ws = wb.create_sheet("更新・除却（R7〜）") if V >= 2026 else None
    if ws is None:
        hist = None
    if ws is not None:
        merge(ws, 1, 1, 1, 8, f"{W.FAC}　機器台帳　更新・除却の記録（令和7年度〜）", f(13, True), al("left"), None)
        put(ws, 2, 1, "基幹的設備改良工事で更新した機器は新しい管理番号を付け、旧番号を除却とした（各シートの「旧管理番号」列と対応）。", f(8))
        hh = ["除却日", "旧管理番号", "新管理番号", "シート", "機器名称", "号炉", "内容", "工事"]
        for i, h in enumerate(hh):
            put(ws, 4, 1 + i, h, f(9, True), al("center"), B_ALL, fill("FFF2CC"))
        rr = 5
        for old, newn, sh, nm, fu, txt in hist:
            rec.add("item", dict(zip(["除却日", "旧管理番号", "新管理番号", "シート", "機器名称", "号炉", "内容", "工事"], ["R8.1.16", old, newn, sh, nm, fu, txt, "基幹的設備改良工事（R7施工分）"])), ws.title, rr)
            for i, v in enumerate(["R8.1.16", old, newn, sh, nm, fu, txt, "基幹的設備改良工事（R7施工分）"]):
                put(ws, rr, 1 + i, v, f(8), al("left" if i in (4, 6, 7) else "center"), B_ALL)
            rr += 1
        for _ in range(4):
            for i in range(8): put(ws, rr, 1 + i, None, None, None, B_ALL)
            rr += 1
        widths(ws, {"A": 9, "B": 10, "C": 10, "D": 10, "E": 26, "F": 5, "G": 30, "H": 28})
        page(ws, "landscape", 9, 1, 1, f"A1:H{rr - 1}")
    # 集計シート
    ws = wb.create_sheet("集計", 0)
    merge(ws, 1, 1, 1, 6, f"{W.FAC}　機器台帳　集計（{RW}現在）", f(14, True), al("left"), None)
    for i, h in enumerate(["大分類（シート）", "1号炉", "2号炉", "共通", "計"]):
        put(ws, 3, 1 + i, h, f(9, True), al("center"), B_ALL, fill("E2EFDA"))
    r = 4; tt = [0, 0, 0]
    for sheet, groups in U_GROUPS:
        cs = [sum(1 for e in W.EQS if e.group in groups and e.furn == k) for k in (1, 2, 0)]
        tt = [a + b for a, b in zip(tt, cs)]
        rec.add("subtotal", {"sheet": sheet, "1号炉": cs[0], "2号炉": cs[1], "共通": cs[2], "計": sum(cs)}, ws.title, r)
        for i, v in enumerate([sheet] + cs + [sum(cs)]):
            put(ws, r, 1 + i, v, f(9), al("left" if i == 0 else "right"), B_ALL)
        r += 1
    rec.add("total", {"1号炉": tt[0], "2号炉": tt[1], "共通": tt[2], "計": sum(tt)}, ws.title, r)
    for i, v in enumerate(["合計"] + tt + [sum(tt)]):
        put(ws, r, 1 + i, v, f(9, True), al("left" if i == 0 else "right"), B_ALL, fill("EDEDED"))
    put(ws, r + 2, 1, "数え方：タグ番号単位。弁は電動弁・調節弁のみ、計装は分析計・主要流量計のみ（手動弁・一般計装は台帳外）。", f(8))
    put(ws, r + 3, 1, "常用・予備は別に数える。共通設備は炉へ重複計上しない。", f(8))
    widths(ws, {"A": 20, "B": 8, "C": 8, "D": 8, "E": 8})
    page(ws, "portrait", 9, 1, 1, f"A1:F{r + 3}")
    props(wb, f"{W.FAC} 機器台帳（{RW}現在）", "u1（【架空】ときわ北部環境組合 施設課）", dt.datetime(2024, 6, 14, 11, 0))
    wb.properties.modified = dt.datetime(V, 4, 3, 15, 0)
    wb.save(path); rec.save()
    return total
