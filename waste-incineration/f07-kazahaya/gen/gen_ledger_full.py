# -*- coding: utf-8 -*-
"""機器台帳：(旧) SPC 保全システム【架空】PM-NAVI の保全対象一覧 CSV（2015 年度から）
            (新) sp2 の Excel 管理表（2020 年度から併用。CSV を貼り付けて列を足した）"""
import csv
import io
from datetime import date
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter
from world import *
from common import *

ASOF = date(2026, 11, 5)
VER = 2026
KIND_SORT = {"機器": 0, "構造": 1, "電動機": 2, "計装": 3, "盤": 4}
UNITNAME = {"1": "1号炉", "2": "2号炉", "0": "共通"}
from works import WORKS
from contracts_full import get as get_contracts
_BYCODE_W = {}
_PENDING = {c["key"] for c in get_contracts() if c["status"] != "完了"}  # 報告書提出前の工事は PM-NAVI に未入力
for _w in WORKS:
    if _w["contract_key"] in _PENDING:
        continue
    _BYCODE_W.setdefault(_w["code"], []).append(_w)
CONSUM = ("ろ布", "ワイヤ", "熱電対", "蓄電池", "RO膜", "充填", "触媒", "パッキン", "チップ")


def fy_of(d):
    return d.year if d.month >= 4 else d.year - 1


def hozen_hoshiki(e):
    r = rule(e, VER >= 2020, VER)
    ts = {t for ws in r.values() for (t, w) in ws}
    if not r:
        return "BDM"
    if "L" in ts:
        return "法定+TBM" if ("O" in ts or "X" in ts) else "法定"
    names = {w for ws in r.values() for (t, w) in ws}
    if any(("測定" in n or "試験" in n or "診断" in n) for n in names):
        return "CBM"
    return "TBM"


def last_done(e, before=None):
    before = before or ASOF
    best = None
    for w in _BYCODE_W.get(e["code"], []):
        if w["end"] and date.fromisoformat(w["end"]) < before:
            d = date.fromisoformat(w["end"])
            if best is None or d >= best[0]:
                best = (d, w["work"])
    return best


def next_plan(e):
    r = rule(e, VER >= 2020, VER)
    fy = fy_of(ASOF)
    done = {(int(w["start"][:4]) if int(w["start"][5:7]) >= 4 else int(w["start"][:4]) - 1, w["work"])
            for w in _BYCODE_W.get(e["code"], []) if w["end"] and date.fromisoformat(w["end"]) < ASOF}
    for y in sorted(r):
        if y < fy:
            continue
        pend = [w for (t, w) in r[y] if (y, w) not in done]
        if pend:
            return (y, "・".join(dict.fromkeys(pend)))
    return None


def install_ym(e):
    ym = (2014, 10 + h(e["code"], 3))
    for w in sorted(_BYCODE_W.get(e["code"], []), key=lambda x: x["start"]):
        if w["t"] == "X" and w["end"] and date.fromisoformat(w["end"]) < ASOF and ("更新" in w["work"] or "交換" in w["work"]) \
                and not any(k in w["work"] for k in CONSUM):
            d = date.fromisoformat(w["end"])
            ym = (d.year, d.month)
    return ym


def status(e):
    for c in get_contracts():
        if c["unit"] == e["unit"] and c["period"][0] <= ASOF <= c["period"][1]:
            return "定修停止中"
    if e["role"] == "予備":
        return "予備"
    return "稼働中"


def rows():
    for e in ALL:
        yield e, last_done(e), next_plan(e), install_ym(e)


# ---------------------------------------------------------------- 旧：保全システム CSV（年度末出力、2026 は 11/05 出力）
def make_csv(asof):
    global ASOF, VER
    ASOF = asof
    VER = fy_of(asof) + (1 if asof.month == 3 else 0)
    name = f"{FN}_機器台帳_{asof:%Y-%m}出力_保全システムCSV.csv"
    head = ["保全対象コード", "保全対象名称", "設備区分", "系統", "号炉", "種別", "親保全対象コード", "常用予備区分", "仕様",
            "メーカー", "型式", "製造番号", "設置年月", "重要度", "保全方式", "前回保全日", "前回保全内容", "次回保全年度",
            "次回保全内容", "状態", "登録日", "更新日"]
    who = "sp1" if asof.year <= 2020 and asof < date(2020, 4, 1) else "sp2"
    ver = "Ver.3.1" if asof < date(2020, 4, 1) else ("Ver.4.0" if asof < date(2023, 4, 1) else "Ver.4.2")
    buf = io.StringIO()
    buf.write(f"#【架空】かざはやエコプラザ 保全対象一覧（全件） 出力日時:{asof:%Y/%m/%d} {8 + asof.day % 3:02d}:{asof.month * 3 % 60:02d} 出力者:{who} "
              f"件数:{len(ALL)} システム:【架空】PM-NAVI {ver}\r\n")
    w = csv.writer(buf, lineterminator="\r\n", quoting=csv.QUOTE_MINIMAL)
    w.writerow(head)
    r = 3
    for e, ld, nx, (iy, im) in rows():
        upd = ld[0] if ld else date(2015, 3, 31)
        vals = [e["code"], e["name"], e["cat"], e["sub"], UNITNAME[e["unit"]], e["kind"], e["motor_of"] or "",
                e["role"], e["spec"], maker_of(e), model_of(e), serial_of(e), f"{iy}/{im:02d}", e["imp"], hozen_hoshiki(e),
                ld[0].strftime("%Y/%m/%d") if ld else "", ld[1] if ld else "",
                nx[0] if nx else "", nx[1] if nx else "", status(e), "2015/03/31", upd.strftime("%Y/%m/%d")]
        w.writerow(vals)
        rec(name, "item", dict(zip(head, vals)), r=r)
        r += 1
    with open(outpath(name), "w", encoding="utf-8-sig", newline="") as f:
        f.write(buf.getvalue())
    return name


# ---------------------------------------------------------------- 新：sp2 Excel 管理表
# 構造：完成図書（水樹プラント）の「機器表」の並びを写し、同じ機器の 1号炉・2号炉を横に並べる。
#       炉別設備と共通設備でシートを分け、電動機は親機器の直下に「└ 電動機」行（1台ずつ別コード）。
SHEETS = [  # (シート名, 号炉区分, 設備区分)
    ("炉別_受入供給", "炉別", ["受入供給設備"]),
    ("炉別_燃焼", "炉別", ["燃焼設備"]),
    ("炉別_通風", "炉別", ["通風設備"]),
    ("炉別_ボイラ", "炉別", ["ボイラ設備"]),
    ("炉別_排ガス", "炉別", ["排ガス処理設備"]),
    ("炉別_灰出し", "炉別", ["灰出し設備"]),
    ("炉別_電気計装", "炉別", ["電気計装設備"]),
    ("共通_受入供給", "共通", ["受入供給設備"]),
    ("共通_薬剤", "共通", ["排ガス処理設備"]),
    ("共通_余熱利用", "共通", ["余熱利用設備"]),
    ("共通_給排水", "共通", ["給水設備", "排水処理設備"]),
    ("共通_灰出し", "共通", ["灰出し設備"]),
    ("共通_電気計装", "共通", ["電気設備", "計装設備"]),
    ("共通_雑設備", "共通", ["雑設備"]),
]

SPARE_V = [("燃焼火格子", 2024, "予備 8枚（2023定修残）"), ("誘引通風機", 2024, "軸受1組（2023購入）"),
           ("ボイラ給水ポンプ", 2018, "メカシール1組"), ("バグフィルタ本体", 2022, "ろ布 予備40本"),
           ("空冷復水器ファン減速機", 2024, "軸受一式1台分"), ("ごみクレーン", 2019, "ワイヤロープ1本")]
MEMO_V = [
    ("1KT-100", 2020, 2023, "IDF 2019軸受交換 振動 1.8mm/s"), ("1KT-100", 2022, 2023, "IDF 2021 振動上昇 4.6mm/s 臨時点検"),
    ("1KT-100", 2024, 2099, "IDF 2023軸受交換後 振動 2.1mm/s"), ("2KT-100", 2025, 2099, "誘引ファン 2024軸受交換"),
    ("1NS-030", 2020, 2023, "2023 更新予定（2020見直しで前倒し）"), ("1NS-030", 2024, 2099, "2023.10 全数更新（192枚）"),
    ("2NS-030", 2020, 2024, "2024 更新予定"), ("2NS-030", 2025, 2099, "2024.6 全数更新"),
    ("1HG-130", 2022, 2099, "バグ ろ布 2021交換"), ("2HG-130", 2023, 2099, "BFろ布 2022交換"),
    ("0YN-010", 2023, 2025, "2022 定期事業者検査（開放）異常なし"), ("0YN-010", 2026, 2099, "2026精密点検 施工中（～11/20）"),
    ("0YN-300", 2023, 2023, "2022 異音 臨時点検（油分析 鉄粉増）"), ("0YN-300", 2024, 2099, "2023 減速機OH（異音）"),
    ("1BO-070", 2022, 2023, "3SH 下段 最小6.3mm（2021測定）"), ("1BO-070", 2024, 2025, "3SH 下段 最小5.6mm（2023測定）"),
    ("1BO-070", 2026, 2099, "3SH 下段 最小5.1mm（2025測定）"), ("2BO-070", 2023, 2024, "3SH 下段 最小6.1mm（2022測定）"),
    ("2BO-070", 2025, 2025, "3SH 下段 最小5.7mm（2024測定）"), ("2BO-070", 2026, 2099, "3SH 下段 最小5.4mm（2026.6）"),
    ("0KS-010", 2025, 2099, "DCS部品 2029供給終了予告(2024.7通知)"),
]


def memo_of(code):
    t = [m for (c, a, b, m) in MEMO_V if c == code and a <= VER <= b]
    if code == "0UK-200" and VER >= 2021:
        y = VER - 1 if (VER - 1) % 2 else VER - 2
        t.append(f"性能検査 {y}.10 合格")
    return t[-1] if t else ""


def spare_of(base):
    t = [m for (b, a, m) in SPARE_V if b == base and VER >= a]
    if base == "空冷復水器ファン減速機" and VER >= 2025:
        return "軸受一式1台分／歯車1組(2024発注)"
    return t[-1] if t else ""


def _unit_cells(e):
    if e is None:
        return ["", "", "", "", ""]
    ld = last_done(e)
    nx = next_plan(e)
    iy, im = install_ym(e)
    return [e["code"], f"{iy}.{im:02d}", (f"{ld[0].strftime('%Y.%m')} {ld[1]}" if ld else ""),
            (f"{nx[0]} {nx[1]}" if nx else ""), memo_of(e["code"])]


def make_xlsx(ver, asof):
    global ASOF, VER
    ASOF, VER = asof, ver
    fname = f"{FN}_機器台帳_{ver}年度_SPC保全管理表.xlsx"
    wb = Workbook()
    setp(wb, "【架空】かざはやエコプラザ 機器管理表（保全対象一覧）", "【架空】かざはや環境パートナーズ sp2")
    ws0 = wb.active
    ws0.title = "集計"
    ws0["A1"] = "【架空】かざはやエコプラザ　機器管理表　集計"
    ws0["A1"].font = font(14, True)
    ws0["A2"] = f"PM-NAVI 保全対象一覧（{asof:%Y/%m/%d} 出力）を完成図書の機器表の並びに貼替え。機器と電動機は別に1台で数える（弁は対象外）。"
    ws0["A2"].font = font(9)
    hdr = ["シート", "設備区分", "行数", "1号炉", "2号炉", "共通", "台数計", "うち電動機"]
    for i, t in enumerate(hdr, 1):
        c = ws0.cell(4, i, t)
        c.font = font(10, True); c.fill = fill("D9E1F2"); c.border = ALL_THIN; c.alignment = C
    r0 = 5
    tot = [0] * 6
    body = []
    for sn, kubun, cats in SHEETS:
        if kubun == "炉別":
            es = [e for e in ALL if e["cat"] in cats and e["unit"] in ("1", "2")]
            keys = []
            for e in es:
                k = (e["base"], e["sfx"], bool(e["motor_of"]))
                if k not in keys:
                    keys.append(k)
            nrow = len(keys)
        else:
            es = [e for e in ALL if e["cat"] in cats and e["unit"] == "0"]
            nrow = len(es)
        n1 = sum(1 for e in es if e["unit"] == "1"); n2 = sum(1 for e in es if e["unit"] == "2")
        n0 = sum(1 for e in es if e["unit"] == "0"); nm = sum(1 for e in es if e["kind"] == "電動機")
        body.append((sn, kubun, cats, es))
        vals = [sn, "・".join(cats), nrow, n1, n2, n0, n1 + n2 + n0, nm]
        for i, v in enumerate(vals, 1):
            c = ws0.cell(r0, i, v); c.font = font(10); c.border = ALL_THIN; c.alignment = C if i > 2 else CL
        rec(fname, "subtotal", dict(zip(hdr, vals)), r=r0, sheet="集計")
        for k, v in enumerate([nrow, n1, n2, n0, n1 + n2 + n0, nm]):
            tot[k] += v
        r0 += 1
    for i, v in enumerate(["合計", ""] + tot, 1):
        c = ws0.cell(r0, i, v); c.font = font(10, True); c.border = box(t=med, b=med, l=thin, r=thin); c.alignment = C
    rec(fname, "total", dict(zip(hdr, ["合計", ""] + tot)), r=r0, sheet="集計")
    ws0.cell(r0 + 2, 1, "※2020.4 作成 sp2（毎年度コピーして更新）。炉別シートは1行に1号炉・2号炉の2台（完成図書 機器表と同じ並び）。台数は保全ｺｰﾄﾞの数で数える。").font = font(9)
    ws0.cell(r0 + 3, 1, "※毎年度 PM-NAVI から貼り直し。メモ・予備品は手入力。").font = font(9)
    for col, wdt in zip("ABCDEFGH", [16, 28, 7, 8, 8, 8, 8, 10]):
        ws0.column_dimensions[col].width = wdt
    ws0.page_setup.orientation = "portrait"; ws0.page_setup.paperSize = 9
    ws0.page_setup.fitToWidth = 1; ws0.sheet_properties.pageSetUpPr.fitToPage = True

    common_head = ["No.", "機器名", "種別", "常/予", "仕様", "メーカー", "重要度", "保全方式"]
    ublock = ["保全ｺｰﾄﾞ", "設置", "前回保全", "次回予定", "メモ"]
    uw = [9, 7, 22, 22, 18]
    cw = [5, 26, 6, 5, 30, 14, 5, 7]
    for sn, kubun, cats, es in body:
        ws = wb.create_sheet(sn)
        shift = 1 if (sn == "共通_余熱利用" and ver >= 2025) else 0  # 2025 年度版で A 列へ空列を挿入したまま（1 列ずれ）
        ws["A1"] = f"【架空】かざはやエコプラザ　機器管理表　{sn.replace('_', '（')}）"
        ws["A1"].font = font(13, True)
        ws["A2"] = f"出典：PM-NAVI 保全対象一覧 {asof:%Y/%m/%d}出力（並びは完成図書 機器表 MZ-KZ-M01 による） ＋ sp2 手入力"
        ws["A2"].font = font(8)
        groups = ["1号炉", "2号炉"] if kubun == "炉別" else ["共通"]
        ncol = len(common_head) + len(ublock) * len(groups) + 1
        upd = asof + timedelta(days=1 if ver == 2026 else 9)
        ws.cell(2, ncol + shift - 2, f"更新：{upd:%Y/%m/%d} sp2").font = font(8)
        # 2 段見出し
        for i, t in enumerate(common_head, 1 + shift):
            ws.merge_cells(start_row=4, start_column=i, end_row=5, end_column=i)
            ws.cell(4, i, t)
        col = len(common_head) + 1 + shift
        for g in groups:
            ws.merge_cells(start_row=4, start_column=col, end_row=4, end_column=col + len(ublock) - 1)
            ws.cell(4, col, g)
            for k, t in enumerate(ublock):
                ws.cell(5, col + k, t)
            col += len(ublock)
        ws.merge_cells(start_row=4, start_column=col, end_row=5, end_column=col)
        ws.cell(4, col, "予備品")
        last = col
        for rr in (4, 5):
            for cc in range(1 + shift, last + 1):
                c = ws.cell(rr, cc); c.font = font(9, True); c.border = ALL_THIN; c.alignment = C
                c.fill = fill("E2EFDA" if kubun == "炉別" else "FFF2CC")
        for i, wdt in enumerate(cw + uw * len(groups) + [16], 1 + shift):
            ws.column_dimensions[get_column_letter(i)].width = wdt
        if shift:
            ws.column_dimensions["A"].width = 2
        # 行の組立
        if kubun == "炉別":
            keys = []
            for e in es:
                k = (e["base"], e["sfx"], bool(e["motor_of"]))
                if k not in keys:
                    keys.append(k)
            lines = []
            for k in keys:
                pair = [next((e for e in es if (e["base"], e["sfx"], bool(e["motor_of"])) == k and e["unit"] == u), None) for u in ("1", "2")]
                lines.append(pair)
        else:
            lines = [[e] for e in es]
        r = 6
        n = 0
        prev_sub = None
        for pair in lines:
            e = pair[0]
            if e["sub"] != prev_sub and not e["motor_of"]:
                # 中区分の見出し行（完成図書の系統見出し）
                ws.merge_cells(start_row=r, start_column=1 + shift, end_row=r, end_column=last)
                ws.cell(r, 1 + shift, f"■ {e['sub']}").font = font(9, True)
                rec(fname, "parent", {"path": [sn, e["sub"]]}, r=r, sheet=sn)
                ws.cell(r, 1 + shift).fill = fill("F2F2F2")
                for cc in range(1 + shift, last + 1):
                    ws.cell(r, cc).border = box(t=thin, b=hair, l=thin if cc == 1 + shift else NONE, r=thin if cc == last else NONE)
                prev_sub = e["sub"]
                r += 1
            n += 1
            isM = bool(e["motor_of"])
            nm = "　└ 電動機" if isM else e["name"]
            vals = [n, nm, e["kind"], {"常用": "常", "予備": "予"}.get(e["role"], ""), e["spec"], maker_of(e).replace("【架空】", ""),
                    e["imp"], hozen_hoshiki(e)]
            for u in pair:
                vals += _unit_cells(u)
            vals.append("" if isM else spare_of(e["base"]))
            f = {"path": [sn, e["sub"], (BY_CODE[e["motor_of"]]["name"] + " 電動機") if isM else e["name"]], "name": nm.strip(),
                 "kind": e["kind"], "role": vals[3], "spec": e["spec"], "maker": vals[5], "importance": e["imp"], "method": vals[7],
                 "units": {}}
            for g, u in zip(groups, pair):
                if u is not None:
                    uc = _unit_cells(u)
                    f["units"][g] = dict(zip(["code", "installed", "last", "next", "memo"], uc))
            f["spare"] = vals[-1]
            rec(fname, "item", f, r=r, sheet=sn)
            for i, v in enumerate(vals, 1 + shift):
                c = ws.cell(r, i, v)
                c.font = font(8, color="0000C0" if (i - shift - len(common_head)) % len(ublock) == 0 and i - shift > len(common_head) and v else None)
                c.border = box(l=thin, r=thin, t=hair, b=hair)
                c.alignment = C if (i - shift) in (1, 3, 4, 7, 8) else CL
            if isM:
                for i in range(1 + shift, last + 1):
                    ws.cell(r, i).fill = fill("F7F7F7")
            if e["code"] == "1NS-030" and ver >= 2020:
                ws.cell(r, len(common_head) + 3 + shift).comment = Comment(
                    "sp2: 2020見直しで更新年を2027→2023に前倒し。2018測定で燃焼段可動火格子の摩耗が想定の約1.6倍。", "sp2")
            r += 1
            if sn == "炉別_排ガス" and e["code"] == "1HG-130":
                for k in range(2):  # 2020 年に入れた空行 2 行（予定メモ用）が残る
                    for i in range(1 + shift, last + 1):
                        ws.cell(r, i).border = box(l=thin, r=thin, t=hair, b=hair)
                    if k == 0 and ver >= 2026:
                        ws.cell(r, len(common_head) + len(ublock) * 2 + shift, "（2027/28 ろ布交換 見積依頼中）").font = font(8, color="C00000")
                        rec(fname, "item", {"path": [sn, "メモ"], "memo": "（2027/28 ろ布交換 見積依頼中）"}, r=r, sheet=sn)
                    r += 1
        ws.freeze_panes = f"{get_column_letter(3 + shift)}6"
        ws.print_title_rows = "4:5"
        ws.page_setup.orientation = "landscape"
        ws.page_setup.paperSize = 8
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.print_area = f"A1:{get_column_letter(last)}{r - 1}"
        ws.oddFooter.center.text = "&A  &P / &N"
    wb.save(outpath(fname))
    return fname


def run():
    out = []
    for fy in range(2015, 2026):
        out.append(("csv", fy, make_csv(date(fy + 1, 3, 31))))
    out.append(("csv", 2026, make_csv(date(2026, 11, 5))))
    for ver in range(2020, 2027):
        asof = date(ver, 3, 31) if ver < 2026 else date(2026, 11, 5)
        out.append(("xlsx", ver, make_xlsx(ver, asof)))
    return out


if __name__ == "__main__":
    for x in run():
        print(x)
