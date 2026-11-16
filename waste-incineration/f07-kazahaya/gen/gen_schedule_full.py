# -*- coding: utf-8 -*-
"""工程表（全契約）
 sp1 予定工程表（2015–2020、SPC 保全 sp1 が発注前に作る 1 工事 1 枚・日単位塗り）
   ※sp2 は 2020 年度も sp1 の様式のまま使った（工程表を q2 に任せたのは 2021 から）
 q2 予定実績工程表（2021–2026、メーカー子会社 職長 q2、半日単位・予定／実績 2 段・班・人員）"""
from datetime import timedelta
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill
from common import *
from contracts_full import get, CREWS, off, TODAY
from works import TASK_WORKS
from world import TEISHU

WD = "月火水木金土日"
LINKS = []
WHO = {"SPC運転": "SPC運転", "電計": "水樹TS電計", "電計2": "水樹TS電計", "足場": "北斗足場", "耐火": "東雲築炉", "機械A": "水樹TS",
       "機械B": "水樹TS", "機械C": "水樹TS", "機械D": "水樹TS", "分析": "分析ｾﾝﾀｰ", "検査": "非破壊検査", "受変電": "風早電設",
       "洗浄": "しぶき洗浄", "SV": "鳴瀬ﾀｰﾋﾞﾝ", "溶接": "かざはや溶接"}


def drange(a, b):
    d = a
    while d <= b:
        yield d
        d += timedelta(days=1)


def codes_text(cl):
    return " ".join(cl) if len(cl) <= 4 else f"{cl[0]}～{cl[-1]}（{len(cl)}件）"


def make_sp1(c, fname):
    y = c["fy"]
    wb = Workbook()
    setp(wb, f"【架空】{c['name']} 工程表", f"【架空】かざはや環境パートナーズ {c['spc']}")
    ws = wb.active
    ws.title = "工程表"
    S, E = c["period"]
    d0 = c["stop"] - timedelta(days=1)
    d1 = E + timedelta(days=2)
    days = list(drange(d0, d1))
    DC = 4
    last = DC + len(days)
    ws["A1"] = f"【架空】{c['name']}　工程表"
    ws["A1"].font = font(15, True)
    ws["A2"] = (f"工期：{S.year}年{S.month}月{S.day}日～{E.year}年{E.month}月{E.day}日　　発注：【架空】かざはや環境パートナーズ　"
                f"受注：【架空】水樹テクノサービス（職長 {c['foreman']}）")
    ws["A2"].font = font(9)
    made = S - timedelta(days=14)
    ws.cell(2, last, f"作成 {made.year}.{made.month}.{made.day} SPC保全 {c['spc']}").font = font(8)
    ws.cell(3, last, "（予定）").font = font(8, color="C00000")
    for col in (1, 2, 3):
        ws.merge_cells(start_row=4, start_column=col, end_row=6, end_column=col)
    ws.cell(4, 1, "No")
    ws.cell(4, 2, "作　業　項　目")
    ws.cell(4, 3, "担当")
    months = {}
    for i, d in enumerate(days):
        months.setdefault((d.year, d.month), []).append(i)
    for (yy, m), idx in months.items():
        ws.merge_cells(start_row=4, start_column=DC + idx[0], end_row=4, end_column=DC + idx[-1])
        ws.cell(4, DC + idx[0], f"{yy}年　{m}月")
    for i, d in enumerate(days):
        ws.cell(5, DC + i, d.day)
        ws.cell(6, DC + i, WD[d.weekday()])
    ws.merge_cells(start_row=4, start_column=last, end_row=6, end_column=last)
    ws.cell(4, last, "備考")
    for r in range(4, 7):
        for col in range(1, last + 1):
            cc = ws.cell(r, col)
            cc.font = font(9, True)
            cc.alignment = C
            cc.border = ALL_THIN
            cc.fill = fill("FFF2CC")
    for i, d in enumerate(days):
        if off(d, True):
            for r in (5, 6):
                ws.cell(r, DC + i).fill = fill("F8CBAD")
    BAR = fill("5B9BD5")
    r = 7
    u = c["unit"]
    head_rows = []
    if u in ("1", "2"):
        rise = [t for t in c["tasks"] if "炉立上げ" in t["name"]]
        rs = rise[0]["ps"] if rise else E
        head_rows.append(("", f"{u}号炉 停止期間", "", c["stop"], rs, f"▼{c['stop'].month}/{c['stop'].day} 6:00 ごみ焚切り　▲{rs.month}/{rs.day} ごみ投入"))
        if u == "1":
            z0, z1 = TEISHU[("0", y)]
            head_rows.append(("", "（参考）全炉停止時共通設備点検整備工事", "別契約", z0, z1, f"全停 {z0.month}/{z0.day}～{z1.month}/{z1.day}"))
    else:
        head_rows.append(("", "全炉停止期間（2号炉停止）", "", c["stop"], E, "1号炉は定修中"))
    for (no, nm, who, a, b, note) in head_rows:
        ws.cell(r, 2, nm)
        ws.cell(r, 3, who)
        ws.cell(r, last, note)
        for i, d in enumerate(days):
            if a <= d <= b:
                ws.cell(r, DC + i).fill = fill("A6A6A6") if "参考" not in nm else fill("D9D9D9")
        r += 1
    for t in c["tasks"]:
        ws.cell(r, 1, t["no"])
        ws.cell(r, 2, t["name"])
        ws.cell(r, 3, WHO.get(t["crew"], t["crew"]))
        for i, d in enumerate(days):
            if t["ps"] <= d <= t["pe"] and (not off(d, True) or t["crew"] in ("SPC運転", "受変電")):
                ws.cell(r, DC + i).fill = BAR
        if t["crew"] in ("耐火", "足場", "検査", "受変電"):
            ws.cell(r, last, f"{CREWS[t['crew']][0].replace('【架空】', '')} {CREWS[t['crew']][2]}名")
        rec(fname, "item", {"processNo": t["no"], "name": t["name"], "crew": WHO.get(t["crew"], t["crew"]), "codes": t["codes"],
                            "planStart": t["ps"].isoformat(), "planEnd": t["pe"].isoformat(), "note": ws.cell(r, last).value or ""},
            r=r, sheet="工程表")
        LINKS.append(dict(file=fname, sheet="工程表", r=r, works=TASK_WORKS.get((c["key"], t["no"]), [])))
        r += 1
    r += 3  # 空行 3 行（追加作業用に sp1 が毎回残す）
    endr = r - 1
    for rr in range(7, endr + 1):
        for col in range(1, last + 1):
            cc = ws.cell(rr, col)
            cc.border = box(l=thin if col <= DC or col == last else hair, r=thin if col < DC or col == last else hair, t=hair, b=hair)
            cc.font = font(9)
            cc.alignment = CL if col in (2, last) else C
        for i, d in enumerate(days):
            if off(d, True) and ws.cell(rr, DC + i).fill.fgColor.rgb in (None, "00000000"):
                ws.cell(rr, DC + i).fill = fill("FBE5D6")
    legend = "凡例　■ 作業予定（日曜・祝日は休工、運転班作業を除く）" if y <= 2019 else "凡例　■ 作業予定（日曜・祝日は休工、運転班作業を除く）※2020 sp2 作成（様式は前任のまま）"
    ws.cell(endr + 2, 1, f"{legend}　　確認：SPC所長 ＿＿＿　保全 {c['spc']}").font = font(9)
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 44
    ws.column_dimensions["C"].width = 10
    for i in range(len(days)):
        ws.column_dimensions[get_column_letter(DC + i)].width = 3.3
    ws.column_dimensions[get_column_letter(last)].width = 40
    ws.print_area = f"A1:{get_column_letter(last)}{endr + 2}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = 8
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    wb.save(outpath(fname))
    return fname


def make_q2(c, fname):
    wb = Workbook()
    setp(wb, f"【架空】{c['name']} 工程表（予定・実績）", "【架空】水樹テクノサービス q2")
    ws = wb.active
    ws.title = "工程表"
    S, E = c["period"]
    d0 = c["stop"] - timedelta(days=0)
    d1 = E + timedelta(days=1)
    days = list(drange(d0, d1))
    cut = c.get("cut")
    DC = 7
    last = DC + 2 * len(days)
    ws["A1"] = f"【架空】{c['name']}　工程表（予定・実績）"
    ws["A1"].font = font(15, True)
    ws["A2"] = (f"工期：{S:%Y年%m月%d日}～{E:%Y年%m月%d日}　発注者：【架空】かざはや環境パートナーズ株式会社　"
                f"受注者：【架空】水樹テクノサービス株式会社　職長：{c['foreman']}")
    ws["A2"].font = font(9)
    ver = f"第3版（実績記入）{E:%Y.%m.%d}" if not cut else f"第2版（実績記入 {cut:%m/%d} まで）"
    ws.cell(1, last, ver).font = font(9)
    ws.cell(2, last, f"作成：水樹TS q2　確認：SPC {c['spc']}").font = font(8)
    hdr = ["No", "作業項目", "対象（保全ｺｰﾄﾞ）", "班", "人員", "区分"]
    for i, t in enumerate(hdr, 1):
        ws.merge_cells(start_row=4, start_column=i, end_row=7, end_column=i)
        ws.cell(4, i, t)
    months = {}
    for i, d in enumerate(days):
        months.setdefault((d.year, d.month), []).append(i)
    for (yy, m), idx in months.items():
        ws.merge_cells(start_row=4, start_column=DC + 2 * idx[0], end_row=4, end_column=DC + 2 * idx[-1] + 1)
        ws.cell(4, DC + 2 * idx[0], f"{yy}年{m}月")
    for i, d in enumerate(days):
        ws.merge_cells(start_row=5, start_column=DC + 2 * i, end_row=5, end_column=DC + 2 * i + 1)
        ws.merge_cells(start_row=6, start_column=DC + 2 * i, end_row=6, end_column=DC + 2 * i + 1)
        ws.cell(5, DC + 2 * i, d.day)
        ws.cell(6, DC + 2 * i, WD[d.weekday()])
        ws.cell(7, DC + 2 * i, "前")
        ws.cell(7, DC + 2 * i + 1, "後")
    ws.merge_cells(start_row=4, start_column=last, end_row=7, end_column=last)
    ws.cell(4, last, "備考（実績）")
    for r in range(4, 8):
        for col in range(1, last + 1):
            cc = ws.cell(r, col)
            cc.font = font(8 if r == 7 else 9, True)
            cc.alignment = C
            cc.border = ALL_THIN
            cc.fill = fill("DEEAF6")
    for i, d in enumerate(days):
        if off(d, False):
            for r in (5, 6, 7):
                ws.cell(r, DC + 2 * i).fill = fill("F8CBAD")
                ws.cell(r, DC + 2 * i + 1).fill = fill("F8CBAD")
    PLAN = PatternFill("lightUp", fgColor="2F5597", bgColor="FFFFFF")
    ACT = fill("2F5597")
    ACT_X = fill("C00000")
    r = 8
    per_day = {d: 0 for d in days}
    for t in c["tasks"]:
        comp, leader, n = CREWS[t["crew"]]
        for k, lab in enumerate(("予定", "実績")):
            rr = r + k
            if k == 0:
                ws.cell(rr, 1, t["no"])
                ws.cell(rr, 2, t["name"])
                ws.cell(rr, 3, codes_text(t["codes"]))
                ws.cell(rr, 4, t["crew"] if t["crew"] != "SPC運転" else "SPC")
                ws.cell(rr, 5, n if n else "")
                for col in range(1, 6):
                    ws.merge_cells(start_row=rr, start_column=col, end_row=rr + 1, end_column=col)
            ws.cell(rr, 6, lab)
            a, b = (t["ps"], t["pe"]) if k == 0 else (t["as_"], t["ae"])
            for i, d in enumerate(days):
                if k == 1 and cut and d > cut:
                    continue
                if a <= d <= b and (not off(d, False) or t["crew"] in ("SPC運転", "受変電")):
                    for hh in (0, 1):
                        cc = ws.cell(rr, DC + 2 * i + hh)
                        cc.fill = PLAN if k == 0 else (ACT_X if d > t["pe"] else ACT)
                    if k == 0 and n:
                        per_day[d] += n
        ev = t.get("event")
        if ev:
            ws.cell(r + 1, last, ev[2])
        f = {"processNo": t["no"], "name": t["name"], "codes": t["codes"], "crew": t["crew"], "people": n or None,
             "planStart": t["ps"].isoformat(), "planEnd": t["pe"].isoformat()}
        if not cut or t["as_"] <= cut:
            f["actualStart"] = t["as_"].isoformat()
            if not cut or t["ae"] <= cut:
                f["actualEnd"] = t["ae"].isoformat()
        if ev:
            f["note"] = ev[2]
        rec(fname, "item", f, r=r, sheet="工程表")
        LINKS.append(dict(file=fname, sheet="工程表", r=r, works=TASK_WORKS.get((c["key"], t["no"]), [])))
        r += 2
    for (d, txt) in c.get("events", []):
        if d in days:
            i = days.index(d)
            ws.cell(r, DC + 2 * i, "▼")
    ws.cell(r, 2, "特記（実績）▼")
    ev_row = r
    r += 1
    ws.cell(r, 2, "予定人員（協力会社含む・SPC除く）")
    for i, d in enumerate(days):
        if per_day[d]:
            ws.merge_cells(start_row=r, start_column=DC + 2 * i, end_row=r, end_column=DC + 2 * i + 1)
            ws.cell(r, DC + 2 * i, per_day[d])
    man_row = r
    for rr in range(8, r + 1):
        for col in range(1, last + 1):
            cc = ws.cell(rr, col)
            cc.font = font(8 if col >= DC else 9)
            cc.alignment = CLW if col in (2, 3, last) else C
            top = thin if (rr - 8) % 2 == 0 or rr >= ev_row else hair
            cc.border = box(l=hair if col >= DC else thin, r=(thin if col % 2 == 0 and col >= DC else hair) if col >= DC else thin, t=top, b=hair)
    for i, d in enumerate(days):
        if ws.cell(ev_row, DC + 2 * i).value == "▼":
            ws.cell(ev_row, DC + 2 * i).font = font(9, True, "C00000")
        if off(d, False):
            for rr in range(8, man_row + 1):
                for hh in (0, 1):
                    cc = ws.cell(rr, DC + 2 * i + hh)
                    if cc.fill is None or cc.fill.fill_type is None:
                        cc.fill = fill("FBE5D6")
    ws.cell(man_row + 2, 2, "凡例　斜線＝予定　塗り＝実績　赤塗り＝予定終了日を超えた実績　▼＝特記事項　休工：日曜（祝日は作業）").font = font(9)
    for k, (d, t) in enumerate(c.get("events", [])):
        ws.cell(man_row + 3 + k, 2, f"▼{d.month}/{d.day}　{t}").font = font(9, color="C00000")
        rec(fname, "item", {"section": "特記", "date": d.isoformat(), "text": t}, r=man_row + 3 + k, sheet="工程表")
    endr = man_row + 2 + len(c.get("events", []))
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 36
    ws.column_dimensions["C"].width = 15
    ws.column_dimensions["D"].width = 6
    ws.column_dimensions["E"].width = 4.5
    ws.column_dimensions["F"].width = 4.5
    for i in range(2 * len(days)):
        ws.column_dimensions[get_column_letter(DC + i)].width = 1.9
    ws.column_dimensions[get_column_letter(last)].width = 30
    ws.freeze_panes = f"{get_column_letter(DC)}8"
    ws.print_area = f"A1:{get_column_letter(last)}{endr}"
    ws.print_title_rows = "4:7"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = 8
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    wb.save(outpath(fname))
    return fname


def run():
    out = []
    for c in get():
        y = c["fy"]
        if y <= 2020:
            fn = f"{FN}_工程表_{y}年度{c['short']}_{c['spc']}予定工程表.xlsx"
            make_sp1(c, fn)
        else:
            fn = f"{FN}_工程表_{y}年度{c['short']}_q2予定実績工程表.xlsx"
            make_q2(c, fn)
        out.append((c, fn))
    return out


if __name__ == "__main__":
    print(len(run()))
