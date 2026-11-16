# -*- coding: utf-8 -*-
"""工事報告書（定期整備工事、SPC→市）。Word → PDF（文字のある PDF、表の見出しはページごとにくり返す）。
2016〜2022：s1 様式（文章型）。2023〜：s2 様式（1 ページ目サマリー、表中心、写真帳）。金額は載せない。"""
import datetime as dt
import os
import world as W
import works as WK
import contracts as K
import gen_contract as G
import report_src as RS
from report_src import para, new_doc, footer, cover, page_break, photo_box, set_font, shade, repeat_header, fix_widths
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Mm
from rec import Rec, path
from xlhelp import WORK

VERB = {"OVERHAUL": "分解整備", "REPLACE": "更新", "UPGRADE": "改良", "INSPECT": "点検"}


def wareki_ym(d):
    if d >= dt.date(2019, 5, 1):
        return f"令和{'元' if d.year == 2019 else d.year - 2018}年{d.month}月"
    return f"平成{d.year - 1988}年{d.month}月"


def wareki_d(d):
    if d >= dt.date(2019, 5, 1):
        return f"令和{'元' if d.year == 2019 else d.year - 2018}年{d.month}月{d.day}日"
    return f"平成{d.year - 1988}年{d.month}月{d.day}日"


def fname(c):
    sty = "s1様式" if c["y"] <= 2022 else "s2様式"
    return f"【架空】ひばり野クリーンセンター_工事報告書_{c['y']}年度{c['f']}号炉定期整備工事_{sty}.pdf"


def idf_vib(f, y):
    if f == 1:
        return {2016: 2.5, 2017: 2.7, 2018: 2.9, 2019: 3.1, 2020: 3.2, 2021: 3.8, 2022: 4.1, 2023: 4.4, 2024: 4.6, 2025: 2.2, 2026: 2.4}[y]
    return round(2.0 + 0.1 * (y - 2016), 1)


def sh_min(f, y):
    if f == 1:
        return 3.6 if y == 2024 else (3.5 if y == 2025 else (4.4 if y >= 2026 else round(4.4 - 0.02 * (y - 2016), 1)))
    return round(4.4 - 0.03 * (y - 2016), 1)


def tbl(doc, R, head, rows, widths, size=8.5, works=None, section=None):
    t = doc.add_table(rows=1, cols=len(head))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(head):
        cc = t.rows[0].cells[i]
        cc.text = ""
        set_font(cc.paragraphs[0].add_run(h), size, True)
        cc.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        shade(cc, "D9E1F2")
    repeat_header(t.rows[0])
    for k, rw in enumerate(rows):
        cells = t.add_row().cells
        for i, val in enumerate(rw):
            cells[i].text = ""
            set_font(cells[i].paragraphs[0].add_run("" if val is None else str(val)), size)
        fl = {h: val for h, val in zip(head, rw) if val not in (None, "")}
        if section:
            fl["表"] = section
        R.add_seq("item", fl, (works[k] if works else None))
    fix_widths(t, widths)
    return t


def sentence(t, y):
    eqs = [W.eq(n, f) for n, f in t[7]]
    nm = "・".join(e.name for e in eqs)
    if t[8] == "INSPECT":
        return f"{nm}は点検を行い、運転に支障のある異常のないことを確認した。"
    e = eqs[0]
    a = W.actual_work(e, y) or set()
    if e.strat == "bf" and "X" in a:
        s = "バグフィルタはろ布 360 本を全数交換し、リテーナの変形品を取り替えた。"
    elif "R" in a and not ({"O", "X"} & a):
        s = f"{nm}は損耗部を確認のうえ、{G.REPAIR.get(e.strat, '部品交換').replace(' ', '')}により補修した。"
    else:
        s = f"{nm}は分解整備を行い、{G.PARTS_WORD.get(e.strat, '消耗部品')}を交換した。"
    if len(t) > 10 and t[10]:
        s += f"なお、{t[10].replace('1日延長', '作業を 1 日延長した').replace('1日後ろへ', '作業を 1 日繰り下げた')}。"
    return s


def next_plan(e, y):
    for yy in range(y + 1, 2031):
        w = W.plan_work(e, yy, y + 1)
        k = next((k for k in "XUOR" if k in w), None)
        if k:
            return yy, {"X": "更新", "U": "改良", "O": "分解整備", "R": "補修"}[k]
    return None


def build_s1(c):
    fn = fname(c)
    R = Rec(fn)
    y, f, tasks = c["y"], c["f"], c["tasks"]
    end = max(x for t in tasks for x in K.task_days(t, y))
    rep = end + dt.timedelta(days=35)
    doc = new_doc(f"【架空】ひばり野クリーンセンター {c['title']} 報告書", "SPC保全課 s1" if y <= 2021 else "SPC保全課 s2")
    cover(doc, c["title"], "工 事 報 告 書", wareki_ym(rep), "ひばり野市長　様", "ひばり野環境サービス株式会社",
          "（施設：【架空】ひばり野クリーンセンター）")
    stop = min(x for t in tasks if t[2].startswith("炉立下げ") for x in K.task_days(t, y))
    up = max(x for t in tasks if t[2].startswith("炉立上げ") for x in K.task_days(t, y))
    para(doc, "1. 工事概要", 13, True)
    for k_, v_ in (("(1) 工事件名", c["title"]), ("(2) 施工場所", f"【架空】ひばり野クリーンセンター　{f}号炉系統"),
                   ("(3) 契約工期", f"{wareki_d(c['period'][0])} 〜 {wareki_d(c['period'][1])}"),
                   ("(4) 炉停止期間", f"{wareki_d(stop)} 〜 {wareki_d(up)}"),
                   ("(5) 発注者", "ひばり野環境サービス株式会社（長期包括運営委託 年度補修計画による）"),
                   ("(6) 施工者", "鷹羽エンジニアリング株式会社（協力会社 野分工業株式会社）")):
        para(doc, f"{k_}　{v_}", 10.5, False, None, 2, 5)
        R.add_seq("item", {"項目": k_[4:], "内容": v_})
    para(doc)
    para(doc, "2. 工事内容", 13, True)
    areas = []
    for t in tasks:
        if t[8] in ("—",) or t[2].startswith(("炉内清掃", "ボイラ水洗", "ボイラ水管")):
            continue
        a = t[1].replace("共通（", "").replace("）", "")
        if a not in areas:
            areas.append(a)
    for i, a in enumerate(areas, 1):
        ts = [t for t in tasks if t[1].replace("共通（", "").replace("）", "") == a and t[8] != "—"]
        txt = "".join(sentence(t, y) for t in ts if not t[2].startswith(("炉内清掃", "ボイラ水洗")) and "UT" not in t[2]
                      and "定期事業者検査" not in t[2] and "水圧" not in t[2])
        if a == "燃焼ガス冷却設備":
            th, pth = G.boiler_min(f, y), G.boiler_min(f, y - 1) if y > 2016 else None
            legal = any("定期事業者検査" in t[2] for t in tasks)
            pre = (f"{f}号ボイラは電気事業法に基づく定期事業者検査の年にあたるため、水洗の後に肉厚測定（UT）を行い、開放検査および水圧試験を受検した。"
                   if legal else f"{f}号ボイラは水洗の後、水管および過熱器の肉厚測定（UT）を行った。")
            txt = pre + f"水管の最小肉厚は {th:.1f}mm であった。" + txt
        if not txt:
            continue
        para(doc, f"2.{i} {a}", 11, True, None, 2)
        para(doc, txt, 10, False, None, 6, 5)
        wk = []
        for t in ts:
            for nm, ff in t[7]:
                w = WK.wid(W.eq(nm, ff), y)
                if w and w not in wk:
                    wk.append(w)
        R.add_seq("item", {"設備": a, "本文": txt}, wk)
    para(doc, "3. 点検結果および所見", 13, True)
    para(doc, "点検の結果、上記の補修を行った箇所を除き、運転に支障のある異常は認められなかった。主な測定結果を以下に示す。", 10, False, None, 4, 5)
    py = y - 2 if y >= 2018 else None
    rows = [(f"{f}号ボイラ 水管", "最小肉厚", f"{G.boiler_min(f, py):.1f}mm" if py else "－", f"{G.boiler_min(f, y):.1f}mm",
             "減肉傾向 要監視" if f == 2 and y >= 2019 else "良"),
            (f"{f}号過熱器", "最小肉厚", f"{sh_min(f, py):.1f}mm" if py else "－", f"{sh_min(f, y):.1f}mm", "良"),
            (f"{f}号誘引通風機", "軸受振動", f"{idf_vib(f, py):.1f}mm/s" if py else "－", f"{idf_vib(f, y):.1f}mm/s",
             "上昇傾向 要監視" if idf_vib(f, y) >= 3.8 else "良")]
    tbl(doc, R, ["機器", "測定項目", f"前回（{py}）" if py else "前回", f"今回（{y}）", "所見"], rows, [32, 24, 26, 26, 42], section="点検結果")
    para(doc)
    para(doc, "4. 使用材料", 13, True)
    parts = G.parts_for(c)
    rows = [(p["name"], p["spec"], f"{p['used']}{p['unit']}", "SPC支給" + (f"（ほか予備 {p['spare']}{p['unit']}）" if p["spare"] else "")) for p in parts]
    rows += [("キャスタブル耐火物", "高アルミナ質", "1式", "施工者手配"), ("当て板・溶接材料", "SS400 ほか", "1式", "施工者手配")]
    tbl(doc, R, ["品名", "仕様", "数量", "備考"], rows, [52, 48, 20, 50], section="使用材料")
    para(doc)
    para(doc, "5. 今後の対応", 13, True)
    items = []
    if G.FINDINGS.get((y, f)):
        items.append(G.FINDINGS[(y, f)] + "。")
    for nm in ("燃焼ストーカ", "誘引通風機", "灰押出装置", "ボイラ本体", "バグフィルタ", "触媒反応塔"):
        e = W.eq(nm, f)
        np_ = next_plan(e, y)
        if np_ and np_[0] <= y + 2:
            items.append(f"{nm}は {np_[0]} 年度に{np_[1]}を計画する。")
    for i, s in enumerate(items[:5], 1):
        para(doc, f"({i}) {s}", 10, False, None, 2, 5)
        R.add_seq("item", {"No": i, "本文": s})
    para(doc, "以上", 10.5, False, "r")
    para(doc, "添付：作業日報（写し）、工程表、写真帳（別冊）", 9)
    footer(doc, f"{c['title']} 報告書")
    d = os.path.join(WORK, f"rep_{c['cid']}.docx")
    doc.save(d)
    RS.to_pdf(d, path(fn))
    return fn


def build_s2(c):
    fn = fname(c)
    R = Rec(fn)
    y, f, tasks = c["y"], c["f"], c["tasks"]
    end = max(x for t in tasks for x in K.task_days(t, y))
    rep = end + dt.timedelta(days=35)
    if (y, f) == (2024, 1):
        rep = dt.date(2024, 12, 10)
    doc = new_doc(f"【架空】ひばり野クリーンセンター {c['title']} 報告書", "SPC保全課 s2")
    cover(doc, c["title"], "工 事 報 告 書", f"{rep.year}年{rep.month}月{rep.day}日", "ひばり野市長　様", "ひばり野環境サービス株式会社",
          "（施設：【架空】ひばり野クリーンセンター）")
    stop = min(x for t in tasks if t[2].startswith("炉立下げ") for x in K.task_days(t, y))
    up = max(x for t in tasks if t[2].startswith("炉立上げ") for x in K.task_days(t, y))
    rows_ = K.daily_rows(tasks, y, seed_add={(2024, 1): 24, (2025, 2): 25}.get((y, f), y * 10 + f))
    mh = sum(m * h for v in rows_.values() for (t, d_, m, h) in v if t[3] != "SPC運転")
    para(doc, "工事報告 サマリー", 14, True, "c", 6)
    changed = [t for t in tasks if t[6] and t[6] != t[5]]
    if (y, f) == (2024, 1):
        add_ = "IDF ランナ肉盛・バランス修正、灰押出装置 側板当て板 1.5m²（変更契約 10/18）"
        dl = "燃焼ストーカ火格子（支給品）2回目納入が 2日遅れ → 10/19（土）作業で吸収、炉立上げは予定どおり"
    else:
        add_ = "なし"
        dl = "、".join(f"{t[2]}（{t[10]}）" for t in changed if len(t) > 10 and t[10]) or "なし"
    ov = [("工事件名", c["title"]), ("施設・系統", f"【架空】ひばり野クリーンセンター　{f}号炉系統"),
          ("契約工期", f"{c['period'][0]:%Y/%m/%d} 〜 {c['period'][1]:%Y/%m/%d}"), ("炉停止", f"{stop:%Y/%m/%d} 〜 {up:%Y/%m/%d}"),
          ("施工", "鷹羽エンジニアリング㈱（協力会社 野分工業㈱、築炉 白砂耐火工業、足場 白鷺鳶工業、UT 東栄非破壊検査）"),
          ("延べ工数", f"協力会社日報集計 {mh:,.0f} 人・時（SPC 運転・立会を除く）"), ("追加・変更", add_), ("遅延", dl)]
    tbl(doc, R, ["項目", "内容"], ov, [35, 135], 9, section="サマリー")
    para(doc, "")
    if (y, f) == (2024, 1):
        srows = RS.SUMMARY_2024
        sw = []
        for rw in srows:
            ids = []
            for t in tasks:
                for nm, ff in t[7]:
                    e = W.eq(nm, ff)
                    if e.no in rw[0] or e.full in rw[1]:
                        w = WK.wid(e, y)
                        if w and w not in ids:
                            ids.append(w)
            sw.append(ids)
    else:
        srows, sw = [], []
        for t in tasks:
            if t[8] == "—" or t[2].startswith("ボイラ水洗"):
                continue
            eqs = [W.eq(n, ff) for n, ff in t[7]]
            if not eqs:
                continue
            nos = "/".join(e.no for e in eqs[:2]) + (" 他" if len(eqs) > 2 else "")
            nm = "・".join(e.full for e in eqs[:3]) + (" 他" if len(eqs) > 3 else "")
            kb = {"INSPECT": "点検", "OVERHAUL": "補修" if "補修" in t[2] else "分解整備", "REPLACE": "更新", "UPGRADE": "改良"}[t[8]]
            res = "完了" if t[8] != "INSPECT" else "良"
            if "UT" in t[2]:
                res = f"水管 最小{G.boiler_min(f, y):.1f}mm"
            np_ = next_plan(eqs[0], y)
            e0 = eqs[0]
            a0 = W.actual_work(e0, y) or set()
            if t[8] == "INSPECT":
                what = t[2]
            elif e0.strat == "bf" and "X" in a0:
                what = "ろ布 360本 全数交換"
            elif "補修" in t[2]:
                what = G.REPAIR.get(e0.strat, "部品交換").replace(" ", "")
            else:
                what = "分解整備、" + G.PARTS_WORD.get(e0.strat, "消耗部品") + " 交換"
            srows.append((nos, nm, kb, what, res, str(np_[0]) if np_ else "－"))
            sw.append([w for w in (WK.wid(e, y) for e in eqs) if w])
    tbl(doc, R, ["資産番号", "機器", "区分", "実施内容", "結果", "次回"], srows, [22, 36, 15, 50, 27, 18], 7.5, works=sw, section="対象機器")
    page_break(doc)
    para(doc, "1. 工事概要", 13, True)
    for k_, v_ in (("(1) 工事件名", c["title"]), ("(2) 施工場所", f"【架空】ひばり野クリーンセンター　{f}号炉系統"),
                   ("(3) 契約工期", f"{c['period'][0].year}年{c['period'][0].month}月{c['period'][0].day}日 〜 {c['period'][1].year}年{c['period'][1].month}月{c['period'][1].day}日"),
                   ("(4) 炉停止期間", f"{stop.year}年{stop.month}月{stop.day}日 〜 {up.year}年{up.month}月{up.day}日"),
                   ("(5) 発注者", "ひばり野環境サービス株式会社（長期包括運営委託 年度補修計画による）"),
                   ("(6) 施工者", "鷹羽エンジニアリング株式会社（協力会社 野分工業株式会社）")):
        para(doc, f"{k_}　{v_}", 10.5, False, None, 2, 5)
        R.add_seq("item", {"項目": k_[4:], "内容": v_})
    para(doc, "2. 実施経過（予定と実績）", 13, True)
    rows, ws_ = [], []
    for t in tasks:
        if t[3] == "SPC運転" and not t[2].startswith(("炉立下げ", "炉立上げ")):
            continue
        ps, pe = t[5]
        a = t[6] or t[5]
        rows.append((t[2], t[3], f"{ps}〜{pe}", f"{a[0]}〜{a[1]}", "変更" if t[6] and t[6] != t[5] else ""))
        ws_.append([w for w in (WK.wid(W.eq(n, ff), y) for n, ff in t[7]) if w])
    tbl(doc, R, ["作業項目", "班", "予定", "実績", "差"], rows, [70, 22, 25, 25, 12], 8, works=ws_, section="実施経過")
    para(doc)
    para(doc, "3. 機器別の記録", 13, True)
    para(doc, "3.1 ボイラ・過熱器 肉厚測定（UT）", 11, True)
    py = y - 2
    if (y, f) == (2024, 1):
        ut = [("水管 前壁", "84", "4.5", "4.0", "4.1", "良"), ("水管 側壁", "132", "4.5", "4.2", "4.2", "良"),
              ("過熱器 1次", "48", "4.5", "4.2", "4.3", "良"), ("過熱器 2次 入口側", "48", "4.5", "3.6", "4.3", "要補修（2026）")]
    else:
        ut = [("水管", "216", "4.5", f"{G.boiler_min(f, y):.1f}", f"{G.boiler_min(f, py):.1f}", "良" if G.boiler_min(f, y) >= 3.6 else "要監視"),
              ("過熱器", "96", "4.5", f"{sh_min(f, y):.1f}", f"{sh_min(f, py):.1f}", "良" if sh_min(f, y) > 3.8 else "要補修")]
    tbl(doc, R, ["部位", "測定点数", "公称", "最小", f"前回({py})", "判定"], ut, [40, 22, 18, 18, 25, 37], 8.5, section="UT")
    para(doc, "3.2 誘引通風機", 11, True)
    if (y, f) == (2024, 1):
        iv = [("軸受振動（反負荷側）", "4.6mm/s", "2.1mm/s", "管理値 4.5mm/s"), ("軸受温度", "68℃", "52℃", "負荷運転時"),
              ("ランナ羽根 入口側", "摩耗あり（最大 2.5mm）", "肉盛・仕上げ", "追加作業")]
    else:
        iv = [("軸受振動（反負荷側）", f"{idf_vib(f, y):.1f}mm/s", "－", "管理値 4.5mm/s（運転中測定）")]
    tbl(doc, R, ["項目", "整備前", "整備後", "備考"], iv, [45, 35, 35, 55], 8.5, section="誘引通風機")
    para(doc)
    para(doc, "4. 支給品・材料の使用", 13, True)
    if (y, f) == (2024, 1):
        mrows = RS.MAT_2024
    else:
        mrows = [(p["name"], p["spec"], str(p["qty"]), str(p["used"]), str(p["spare"]), "予備品として保管" if p["spare"] else "")
                 for p in G.parts_for(c)]
    tbl(doc, R, ["品名", "型式", "支給数", "使用", "残（予備）", "備考"], mrows, [42, 34, 15, 15, 18, 46], 8, section="支給品")
    para(doc, "施工者手配材料：キャスタブル耐火物、当て板 SS400、溶接材料 一式", 9, False, None, 6)
    para(doc, "5. 所見と次回計画への反映", 13, True)
    if (y, f) == (2024, 1):
        items = ["1号過熱器 2次入口側：2026年度 補修を中長期整備計画に追加する（2025年度改訂）。",
                 "1号バグフィルタ ろ布：試験結果により全数交換を 2025年度から 2026年度へ延伸する。",
                 "1号誘引通風機：整備後の振動は良好。廃止までの分解整備は行わず、傾向監視とする。",
                 "燃焼ストーカ火格子の支給品は、次回から 1回納入（定修 2週間前）とするよう購買に依頼した。"]
    else:
        items = []
        if G.FINDINGS.get((y, f)):
            items.append(G.FINDINGS[(y, f)] + "。")
        for nm in ("燃焼ストーカ", "誘引通風機", "ボイラ本体", "過熱器", "バグフィルタ"):
            np_ = next_plan(W.eq(nm, f), y)
            if np_ and np_[0] <= y + 2:
                items.append(f"{f}号{nm}：{np_[0]}年度 {np_[1]}。")
        if not items:
            items.append("今回の結果による計画の変更はない。")
    for i, s in enumerate(items, 1):
        para(doc, f"({i}) {s}", 10, False, None, 2, 3)
        R.add_seq("item", {"No": i, "本文": s})
    page_break(doc)
    para(doc, "写 真 帳", 14, True, "c")
    caps = [f"写真{i + 1} {t[2]}" for i, t in enumerate([t for t in tasks if t[8] in ("OVERHAUL", "REPLACE")][:4])]
    while len(caps) < 4:
        caps.append(f"写真{len(caps) + 1} 炉内 全景")
    photo_box(doc, caps)
    footer(doc, f"{c['title']} 報告書")
    d = os.path.join(WORK, f"rep_{c['cid']}.docx")
    doc.save(d)
    RS.to_pdf(d, path(fn))
    return fn


def build(c):
    if (c["y"], c["f"]) == (2026, 1):
        return None            # 2026-11-13 時点で未提出（12 月提出予定）
    return build_s1(c) if c["y"] <= 2022 else build_s2(c)
