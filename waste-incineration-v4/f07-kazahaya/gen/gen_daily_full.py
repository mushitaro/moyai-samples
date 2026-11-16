# -*- coding: utf-8 -*-
"""作業日報（全日分）
 q1 工事日報（建設時様式 KZ-07 準用、2015–2018）
 q2 作業日報 MTS-日報-2019（2019–2023、SPC 提出用・保全ｺｰﾄﾞ順）
 q2 作業日報 MTS-日報-2024（2024–2026、2019 様式をコピーして：作業時間欄、工数 0.5h 単位の小数表示、会社別小計行）
 全停分は 2021 年度から副職長 q3 が記入（日付の書き方・工数の小数・安全欄が短い）。"""
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment
from common import *
from world import BY_CODE
from contracts_full import get, CREWS
from works import TASK_WORKS, BY_ID
import daytext as DT
import sample_daily as SD

WD = "月火水木金土日"
LINKS = []
JOB = {"機械A": "機械", "機械B": "機械", "機械C": "機械", "機械D": "機械", "電計": "電気", "電計2": "電気", "耐火": "築炉", "足場": "とび",
       "検査": "検査", "分析": "分析", "受変電": "電気", "洗浄": "洗浄", "SV": "指導員", "溶接": "溶接"}


def jd(d):
    return f"{d.year}年{d.month}月{d.day}日（{WD[d.weekday()]}）"


def fmt_of(c):
    y = c["fy"]
    return "q1" if y <= 2018 else ("q2_2019" if y <= 2023 else "q2_2024")


def writer_of(c):
    return c["zentei_writer"] if c["unit"] == "0" else c["foreman"]


def wids(c, no, codes):
    ws = TASK_WORKS.get((c["key"], no), [])
    if codes:
        ws = [w for w in ws if BY_ID[w]["code"] in codes] or ws
    return ws


# ---------------------------------------------------------------- q1 様式
TOOLS = {"grate_out": "2.8tチェーンブロック、電動ウインチ", "grate_in": "チェーンブロック、トルクレンチ", "wear": "デプスゲージ、ノギス",
         "refr": "ミキサ、バイブレータ、ブレーカ", "clean": "ブレーカ、送風機（仮設）", "idf": "油圧プーラ、ダイヤルゲージ、振動計",
         "thick": "超音波厚さ計", "slide": "溶接機（200A）", "pump": "プーラ、ダイヤルゲージ", "scaf_up": "足場材、ラチェット"}
KY = {"grate_out": "重量物の吊り荷の下に入らない", "refr": "粉じん（防じんマスク着用）", "clean": "炉内の酸素濃度測定、熱中症",
      "idf": "回転体の誤起動防止（施錠確認）", "scaf_up": "墜落防止（安全帯使用）", "scaf_down": "落下物、立入禁止措置",
      "slide": "溶接の火気管理", "thick": "照明確保、足元注意", "subst": "感電防止（検電・接地）"}


def q1_info(c, d):
    rows = DT.rows_for(c, d)
    comp = {}
    for r in rows:
        if r["crew"] == "SPC":
            continue
        cn = CREWS[r["task"]["crew"]][0]
        key = (cn, JOB.get(r["task"]["crew"], "作業"))
        comp[key] = comp.get(key, 0) + (r["n"] or 0)
    lines = []
    cats = []
    for r in rows:
        cat = DT.cat_of(r["task"])
        cats.append(cat)
        txt = r["text"]
        if c["unit"] in ("1", "2"):
            txt = txt.replace(f"{c['unit']}号炉", "１号炉" if c["unit"] == "1" else "２号炉")
        lines.append((txt, r))
        if r["note"] in ("追加", "手直し", "遅れ"):
            lines.append(("　→" + r["task"]["event"][2], r))
    tools = "、".join(dict.fromkeys(TOOLS[k] for k in cats if k in TOOLS)) or "一般工具"
    ky = "、".join(dict.fromkeys(KY[k] for k in cats if k in KY)) or "足元・整理整頓"
    ev = [r["task"]["event"][2] for r in rows if r["note"] in ("追加", "手直し")]
    order = ev[0] + "（SPC へ報告）" if ev else DT.pick(["特になし", "特になし", "作業終了時の火気確認を徹底", "通路の資材整理"], c["key"], d)
    return dict(weather=DT.weather(d), temp=f"{18 + DT.zlib.crc32(d.isoformat().encode()) % 12}℃",
                people=[(k[0], k[1], v) for k, v in comp.items()], work=lines, tools=tools, ky=ky, order=order)


def make_q1(c, fname):
    wb = Workbook()
    setp(wb, f"【架空】{c['name']} 工事日報", f"【架空】水樹テクノサービス {writer_of(c)}")
    wb.remove(wb.active)
    days = DT.report_days(c)
    cum = 0
    for no, d in enumerate(days, 1):
        sample = c.get("sample") == "C18" and d in SD.OLD_DAYS
        if sample:
            info0 = SD.OLD_DAYS[d]
            rows = DT.rows_for(c, d)
            info = dict(weather=info0["weather"], temp=info0["temp"], people=info0["people"],
                        work=[(t, None) for t in info0["work"]], tools=info0["tools"], ky=info0["ky"], order=info0["order"])
        else:
            info = q1_info(c, d)
        today_n = sum(p[2] for p in info["people"])
        cum += today_n
        sn = f"{d.month}.{d.day}"
        ws = wb.create_sheet(sn)
        for col, wdt in zip("ABCDEFGH", [16, 12, 12, 12, 10, 10, 10, 10]):
            ws.column_dimensions[col].width = wdt
        ws.merge_cells("A1:F1")
        ws["A1"] = "工　事　日　報"
        ws["A1"].font = font(18, True)
        ws["A1"].alignment = C
        ws["G1"] = f"No.{no}"
        ws["G1"].font = font(11)
        ws.merge_cells("A2:H2")
        ws["A2"] = "【架空】水樹テクノサービス株式会社　（水樹プラント 建設工事 様式 KZ-07 準用）"
        ws["A2"].font = font(8)
        ws["A2"].alignment = Alignment(horizontal="right")
        hdr = [("工事件名", "【架空】" + c["name"]), ("日　付", f"{jd(d)}　　天候 {info['weather']}　　気温 {info['temp']}"),
               ("作業時間", f"8:00 ～ 17:00　（昼休 12:00～13:00）　　職長 {writer_of(c)}")]
        r = 3
        for k, v in hdr:
            ws.cell(r, 1, k)
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
            ws.cell(r, 2, v)
            for col in range(1, 9):
                ws.cell(r, col).border = ALL_THIN
                ws.cell(r, col).font = font(10)
            ws.cell(r, 1).alignment = C
            ws.cell(r, 1).fill = fill("EDEDED")
            r += 1
        r += 1
        ws.cell(r, 1, "■ 出 面 （人）").font = font(10, True)
        r += 1
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
        for i, t in enumerate(["会　社　名", None, None, "職種", "本日", "累計"], 1):
            ws.cell(r, i, t)
        for col in range(1, 7):
            ws.cell(r, col).border = ALL_THIN
            ws.cell(r, col).font = font(9, True)
            ws.cell(r, col).alignment = C
            ws.cell(r, col).fill = fill("EDEDED")
        r += 1
        for comp, job, n in info["people"]:
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
            ws.cell(r, 1, comp)
            ws.cell(r, 4, job)
            ws.cell(r, 5, n)
            for col in range(1, 7):
                ws.cell(r, col).border = ALL_THIN
                ws.cell(r, col).font = font(9)
                ws.cell(r, col).alignment = C if col > 3 else CL
            rec(fname, "item", {"date": d.isoformat(), "weather": info["weather"], "company": comp, "job": job, "people": n,
                                "section": "出面"}, r=r, sheet=sn)
            r += 1
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
        ws.cell(r, 1, "合　計")
        ws.cell(r, 5, today_n)
        ws.cell(r, 6, cum)
        for col in range(1, 7):
            ws.cell(r, col).border = box(t=dbl)
            ws.cell(r, col).font = font(9, True)
            ws.cell(r, col).alignment = C
        rec(fname, "total", {"date": d.isoformat(), "people": today_n, "cumulative": cum, "section": "出面"}, r=r, sheet=sn)
        r += 2
        ws.cell(r, 1, "■ 作 業 内 容").font = font(10, True)
        r += 1
        top = r
        num = 0
        lines = info["work"] + [("", None)] * max(0, 9 - len(info["work"]))
        for (t, rw) in lines:
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
            if t and not t.startswith("　"):
                num += 1
            ws.cell(r, 1, (f"{num}. " if t and not t.startswith("　") else "") + t)
            ws.cell(r, 1).font = font(10)
            for col in range(1, 9):
                ws.cell(r, col).border = box(l=thin if col == 1 else NONE, r=thin if col == 8 else NONE,
                                             t=thin if r == top else dot, b=dot)
            if t:
                f = {"date": d.isoformat(), "weather": info["weather"], "work": t.strip(), "section": "作業内容"}
                if rw is not None:
                    LINKS.append(dict(file=fname, sheet=sn, r=r, works=wids(c, rw["no"], rw["codes"])))
                rec(fname, "item", f, r=r, sheet=sn)
            r += 1
        r += 1
        for k, v in [("使用機械・工具", info["tools"]), ("ＫＹ活動", info["ky"]), ("指示・連絡事項", info["order"]),
                     ("災害・事故", "無　　ヒヤリハット　無")]:
            ws.cell(r, 1, k)
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
            ws.cell(r, 2, v)
            for col in range(1, 9):
                ws.cell(r, col).border = ALL_THIN
                ws.cell(r, col).font = font(9)
                ws.cell(r, col).alignment = CLW
            ws.cell(r, 1).fill = fill("EDEDED")
            ws.row_dimensions[r].height = 30
            r += 1
        r += 1
        for i, t in enumerate(["SPC 立会確認", "職　長"]):
            col = 5 + 2 * i
            ws.merge_cells(start_row=r, start_column=col, end_row=r, end_column=col + 1)
            ws.cell(r, col, t)
            ws.merge_cells(start_row=r + 1, start_column=col, end_row=r + 3, end_column=col + 1)
            ws.cell(r + 1, col, c["spc"] if i == 0 else writer_of(c))
            ws.cell(r + 1, col).font = font(12, color="C00000")
            for rr in range(r, r + 4):
                for cc in (col, col + 1):
                    ws.cell(rr, cc).border = ALL_THIN
                    ws.cell(rr, cc).alignment = C
        ws.print_area = f"A1:H{r + 3}"
        ws.page_setup.orientation = "portrait"
        ws.page_setup.paperSize = 9
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 1
        ws.sheet_properties.pageSetUpPr.fitToPage = True
    wb.save(outpath(fname))
    return fname


# ---------------------------------------------------------------- q2 様式（2019 / 2024）
def gen_q2_info(c, d):
    rows = DT.rows_for(c, d)
    out = []
    for r in rows:
        code = " ".join(r["codes"]) if len(r["codes"]) <= 6 else f"{r['codes'][0]}ほか{len(r['codes']) - 1}件"
        out.append((r["crew"], r["no"], r["text"], code, r["n"], r["hrs"], r["prog"], r["note"] if r["note"] != "SPC運転" else "SPC運転"))
    ev = [r["task"]["event"] for r in rows if r["note"] in ("追加", "手直し", "遅れ")]
    special = "。".join(e[2] for e in ev) + ("。SPC " + c["spc"] + " へ報告。" if ev else "")
    cats = [DT.cat_of(r["task"]) for r in rows]
    wit = ""
    if any(k in cats for k in ("legal", "hydro", "subst", "deaer", "turb_end")):
        wit = f"SPC {c['spc']}、電気主任技術者" if "subst" in cats else f"SPC {c['spc']}"
    elif any(r["note"] in ("追加", "手直し") for r in rows):
        wit = f"SPC {c['spc']} 現地確認"
    return dict(weather=DT.weather(d), rows=out, special=special, witness=wit)


def _split(rows):
    coded, plain = [], []
    for (crew, no, txt, code, n, hrs, prog, note) in rows:
        if "～" in (code or "") or "ほか" in (code or ""):
            codes = [code]
        else:
            codes = code.split() if code else []
        if not codes:
            plain.append([("", "（ｺｰﾄﾞなし）", txt, crew, n, hrs, prog, no, note)])
            continue
        grp = []
        for k, cd in enumerate(codes):
            if cd in BY_CODE:
                nm = BY_CODE[cd]["name"]
            elif cd.startswith("0DE"):
                nm = "受変電設備 一式"
            else:
                nm = BY_CODE[cd.split("ほか")[0]]["name"] + "ほか"
            grp.append((cd, nm, txt if k == 0 else "〃", crew, n, hrs, prog, no, note if k == 0 else ""))
        coded.append(grp)
    coded.sort(key=lambda g: g[0][0])
    return coded, plain


def make_q2(c, fname, v2024):
    wb = Workbook()
    writer = writer_of(c)
    q3 = writer == "q3"
    setp(wb, f"【架空】{c['name']} 作業日報", f"【架空】水樹テクノサービス {writer}")
    wb.remove(wb.active)
    NC = 9
    for d in DT.report_days(c):
        if c.get("sample") == "C23" and d in SD.NEW_DAYS_C23:
            info = SD.NEW_DAYS_C23[d]
        elif c.get("sample") == "C23C" and d in SD.NEW_DAYS_C23C:
            info = SD.NEW_DAYS_C23C[d]
        else:
            info = gen_q2_info(c, d)
        sn = f"{d:%m%d}"
        ws = wb.create_sheet(sn)
        widths = [11, 22, 42, 7, 5, 6.5, 6, 5, 20]
        if v2024:
            widths = [11, 22, 40, 7, 5, 7, 6, 5, 18]
        for i, wdt in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = wdt
        ws.merge_cells("A1:E2")
        ws["A1"] = "作 業 日 報"
        ws["A1"].font = font(16, True)
        ws["A1"].alignment = Alignment(horizontal="left", vertical="center")
        for (rng, txt) in [("F1:G1", "SPC確認"), ("H1:I1", "PM-NAVI転記")]:
            ws.merge_cells(rng)
            ws[rng.split(":")[0]] = txt
        ws.merge_cells("F2:G3")
        ws["F2"] = c["spc"]
        ws.merge_cells("H2:I3")
        pend = c["status"] != "完了"
        ws["H2"] = "未" if (q3 or pend) else "済　（　/　）"
        for rr in range(1, 4):
            for cc in range(6, 10):
                ws.cell(rr, cc).border = ALL_THIN
                ws.cell(rr, cc).alignment = C
        ws["F1"].font = font(9)
        ws["H1"].font = font(9)
        ws["F2"].font = font(14, color="C00000")
        ws["H2"].font = font(10)
        ws["A3"] = f"記入 {writer}" + ("（全停）" if q3 else "")
        ws["A3"].font = font(9)
        ws["A4"] = "件名"
        ws.merge_cells("B4:C4")
        ws["B4"] = "【架空】" + c["name"]
        ws["D4"] = "日付"
        ws.merge_cells("E4:I4")
        dtxt = f"{d:%Y/%m/%d}（{WD[d.weekday()]}）" if q3 else jd(d)
        ws["E4"] = f"{dtxt}　{info['weather']}" + ("　8:00-17:00" if v2024 else "")
        for col in range(1, NC + 1):
            ws.cell(4, col).border = ALL_THIN
            ws.cell(4, col).font = font(10)
        ws["A4"].fill = fill("DEEAF6")
        ws["D4"].fill = fill("DEEAF6")
        hdr = ["保全ｺｰﾄﾞ", "機器名（PM-NAVI）", "作業内容", "班", "人数", "工数(h)", "進捗%", "工程", "備考"]
        r = 6
        for i, t in enumerate(hdr, 1):
            cc = ws.cell(r, i, t)
            cc.font = font(9 if i != 8 else 7, True)
            cc.fill = fill("DEEAF6" if not v2024 else "E2EFDA")
            cc.border = box(t=med, b=thin, l=thin, r=thin)
            cc.alignment = C
        r += 1
        coded, plain = _split(info["rows"])
        tn = th = 0
        used = 0
        bycomp = {}
        for gi, grp in enumerate(coded + plain):
            if gi == len(coded) and plain:
                ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=NC)
                ws.cell(r, 1, "共通作業（保全ｺｰﾄﾞなし）").font = font(8, True)
                ws.cell(r, 1).fill = fill("F2F2F2")
                for cc in range(1, NC + 1):
                    ws.cell(r, cc).border = box(l=thin, r=thin, t=thin, b=hair)
                r += 1
                used += 1
            top = r
            for k, vals in enumerate(grp):
                for i, v in enumerate(vals, 1):
                    if k > 0 and i in (4, 5, 6, 7, 8):
                        v = None
                    if i == 6 and isinstance(v, float) and not (q3 or v2024) and v == int(v):
                        v = int(v)
                    cc = ws.cell(r, i, v)
                    cc.font = font(9 if i != 8 else 8)
                    cc.border = box(l=thin, r=thin, t=thin if k == 0 else hair, b=hair)
                    cc.alignment = CLW if i in (2, 3, 9) else C
                    if i == 6 and (q3 or v2024) and isinstance(v, (int, float)):
                        cc.number_format = "0.0"
                cd, nm, txt, crew, n, hrs, prog, no, note = vals
                f = {"date": d.isoformat(), "weather": info["weather"], "code": cd, "name": nm, "work": txt, "note": note}
                if k == 0:
                    f.update({"crew": crew, "processNo": no, "people": n, "manDays": hrs, "manDaysText": ("" if hrs is None else (f"{hrs:.1f}" if (q3 or v2024) else f"{hrs:g}")),
                              "manDaysUnit": "h", "progress": prog})
                rec(fname, "item", f, r=r, sheet=sn)
                LINKS.append(dict(file=fname, sheet=sn, r=r, works=wids(c, no, [cd] if cd in BY_CODE else [])))
                ws.row_dimensions[r].height = 22 if len(grp) == 1 else 16
                r += 1
                used += 1
            if len(grp) > 1:
                for col in (4, 5, 6, 7, 8):
                    ws.merge_cells(start_row=top, start_column=col, end_row=r - 1, end_column=col)
            tn += grp[0][4] or 0
            th += grp[0][5] or 0
            crew = grp[0][3]
            if crew and crew != "SPC" and crew in CREWS:
                comp = CREWS[crew][0].replace("【架空】", "")
                a = bycomp.setdefault(comp, [0, 0])
                a[0] += grp[0][4] or 0
                a[1] += grp[0][5] or 0
        for _ in range(max(0, 12 - used)):
            for i in range(1, NC + 1):
                ws.cell(r, i).border = box(l=thin, r=thin, t=hair, b=hair)
            ws.row_dimensions[r].height = 18
            r += 1
        if v2024:
            for comp, (pn, ph) in bycomp.items():
                ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
                ws.cell(r, 1, f"小計　{comp}")
                ws.cell(r, 5, pn)
                ws.cell(r, 6, ph)
                ws.cell(r, 6).number_format = "0.0"
                for col in range(1, NC + 1):
                    cc = ws.cell(r, col)
                    cc.font = font(8)
                    cc.border = box(t=hair, b=hair, l=thin, r=thin)
                    cc.alignment = C
                rec(fname, "subtotal", {"date": d.isoformat(), "company": comp, "people": pn, "manDays": ph,
                                        "manDaysText": f"{ph:.1f}", "manDaysUnit": "h"}, r=r, sheet=sn)
                r += 1
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
        ws.cell(r, 1, "計（協力会社含む・SPC除く）" if not v2024 else "合計（協力会社含む・SPC除く）")
        ws.cell(r, 5, tn)
        ws.cell(r, 6, th if (q3 or v2024) or th != int(th) else int(th))
        if q3 or v2024:
            ws.cell(r, 6).number_format = "0.0"
        for col in range(1, NC + 1):
            cc = ws.cell(r, col)
            cc.font = font(9, True)
            cc.border = box(t=dbl, b=med, l=thin, r=thin)
            cc.alignment = C
        rec(fname, "total", {"date": d.isoformat(), "people": tn, "manDays": th, "manDaysUnit": "h"}, r=r, sheet=sn)
        r += 2
        safety = ("KY実施・入炉前酸素濃度測定 記録済　災害 無" if not q3 else "KY済 災害無")
        for k, v, hgt in [("特記事項（発見・手直し・変更）", info["special"], 64), ("立会・検査", info["witness"], 22), ("安全", safety, 22)]:
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
            ws.cell(r, 1, k)
            ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=NC)
            ws.cell(r, 3, v)
            for col in range(1, NC + 1):
                ws.cell(r, col).border = ALL_THIN
                ws.cell(r, col).font = font(9)
                ws.cell(r, col).alignment = CT if col == 3 else C
            ws.cell(r, 1).fill = fill("DEEAF6")
            ws.row_dimensions[r].height = hgt
            if v and k.startswith("特記"):
                rec(fname, "item", {"date": d.isoformat(), "section": "特記事項", "text": v}, r=r, sheet=sn)
            r += 1
        form = "MTS-日報-2024（q2）SPC提出用・保全ｺｰﾄﾞ順・会社別小計" if v2024 else "MTS-日報-2019（q2）SPC提出用・保全ｺｰﾄﾞ順"
        ws.cell(r, 1, "様式 " + form + ("　※全停分 q3 記入" if q3 else "")).font = font(7, color="808080")
        ws.print_area = f"A1:I{r}"
        ws.page_setup.orientation = "landscape"
        ws.page_setup.paperSize = 9
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 1
        ws.sheet_properties.pageSetUpPr.fitToPage = True
    wb.save(outpath(fname))
    return fname


def run():
    out = []
    for c in get():
        f = fmt_of(c)
        w = writer_of(c)
        y = c["fy"]
        if f == "q1":
            fn = f"{FN}_作業日報_{y}年度{c['short']}_{w}工事日報.xlsx"
            make_q1(c, fn)
        else:
            fn = f"{FN}_作業日報_{y}年度{c['short']}_{w}作業日報.xlsx"
            make_q2(c, fn, f == "q2_2024")
        out.append((c, fn))
    return out


if __name__ == "__main__":
    import time
    t0 = time.time()
    r = run()
    print(len(r), time.time() - t0)
