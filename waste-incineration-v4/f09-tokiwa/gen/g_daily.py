# -*- coding: utf-8 -*-
"""作業日報（全年度）：x1 2009 様式（〜2013）／2014 様式（2014–2023）／2024 様式（2024–）／メーカー週報（基幹改良）"""
import datetime as dt, random
from openpyxl import Workbook
import world as W, works as K
from xlh import *
from rec import Rec
from daily_gen import CREW_K, CREW_N, WEATHER15, PHRASES

def wid(c, t):
    return [f"{c.cid}:{t.tid}:{x}" for x in t.tags] if t.tags else [f"{c.cid}:{t.tid}"]

def days_of(c):
    days = {}
    for t in c.tasks:
        for d, txt, crew in t.steps:
            days.setdefault(d, []).append((t, txt, crew))
    return dict(sorted(days.items()))

def hours(d, txt, c):
    if "【追加】" in txt:
        return "8:00～19:00", 10
    if c.cid == "T2025-1" and d == dt.date(2025, 5, 23):
        return "8:00～18:30", 9.5
    return "8:00～17:00", 8

def era(d):
    if d >= dt.date(2019, 5, 1):
        n = d.year - 2018
        return f"R{'元' if n == 1 else n}年{d.month}月{d.day}日"
    return f"H{d.year - 1988}年{d.month}月{d.day}日"

def weather(d):
    if d.year == 2015 and d.month == 10 and d.day in WEATHER15:
        return WEATHER15[d.day]
    return ["晴", "晴", "曇", "雨", "晴", "曇時々晴", "晴"][(d.toordinal() * 7) % 7]

def tokki(c, d, items):
    if c.cid == "T2015-2":
        return {13: "過熱器下部ベンド No.14　肉厚 2.6mm（基準 3.2mm）　sv1 さんへ報告。補修方法は明日指示。",
                14: "定期事業者検査 内部 終了。過熱器 No.14 は肉盛補修の指示あり（追加）。", 15: "肉盛溶接のため溶接工 2 名 19:00 まで残業。",
                10: "不燃物抜出スクリュー 工場へ搬出（15日戻り予定）。", 19: "ボイラ閉止 肉盛の影響で 1 日遅れ。水圧試験 21日に変更（sv1 了承）。"}.get(d.day, "")
    if c.cid == "T2025-1":
        return {22: "BF 第4室 リテーナ（ケージ）曲がり 3 本。sv2 さんへ報告、予備品で交換の了承を得た。",
                23: "ケージ交換のため BF 班 18:30 まで。BF 完了は 27日の見込み（1 日遅れ）。",
                20: "過熱器 No.9 ベンド 2.5mm。計算必要厚 2.1mm 以上のため継続使用（基幹改良 R8 で更新予定）。",
                16: "溶融炉 側壁下段 残存厚 最小 118mm（2023 年 2号炉 95mm）。",
                26: "破砕機 1号 停止、2号機で受入継続（SPC 運転班 了承）。"}.get(d.day if d.month == 5 else 0, "")
    out = []
    for t, txt, crew in items:
        if txt.startswith("【追加】"):
            out.append(txt.replace("【追加】", "") + "。SPC へ報告し追加で施工、残業。")
        elif "定期事業者検査" in txt:
            out.append("定期事業者検査 内部 終了。")
        elif "最小" in txt and "mm" in txt:
            out.append(txt.split("、")[-1] + " を SPC へ報告。")
    return " ".join(out[:2])

# ============================================================== x1 2009／2014 様式（縦）
def x1_v(c, path, fname, form):
    """form: '2009'（人数のみ、KY 欄なし）または '2014'（職種別人数、KY・使用機材）"""
    rec = Rec(fname)
    wb = Workbook(); wb.remove(wb.active)
    days = days_of(c)
    shuku = c.short if c.kind != "dcs" else "DCS更新工事（現地工事）"
    for n, (d, items) in enumerate(days.items(), 1):
        ws = wb.create_sheet(f"{d.month}.{d.day}")
        ncol = 13 if form == "2014" else 8
        merge(ws, 1, 1, 1, ncol - 1, "【架空】作　業　日　報", f(16, True), al("center"), None)
        put(ws, 1, ncol, f"No.{n}", f(9), al("right"))
        merge(ws, 2, 1, 2, 2, "工事件名", f(9), al("center"), B_ALL, fill("F2F2F2"))
        kname = c.title.replace("ときわ北部クリーンセンター", "ときわ北部ＣＣ").replace("【架空】", "")
        merge(ws, 2, 3, 2, ncol, kname, f(10), al("left"), B_ALL)
        merge(ws, 3, 1, 3, 2, "日　付", f(9), al("center"), B_ALL, fill("F2F2F2"))
        merge(ws, 3, 3, 3, ncol - 6 if form == "2014" else 4, f"{era(d)}（{YOUBI[d.weekday()]}）", f(10), al("left"), B_ALL)
        hc = ncol - 5 if form == "2014" else 5
        merge(ws, 3, hc, 3, hc + 1, "天　候", f(9), al("center"), B_ALL, fill("F2F2F2"))
        merge(ws, 3, hc + 2, 3, ncol, weather(d), f(10), al("left"), B_ALL)
        merge(ws, 4, 1, 4, 2, "施工会社", f(9), al("center"), B_ALL, fill("F2F2F2"))
        merge(ws, 4, 3, 4, ncol - 6 if form == "2014" else 4, "北嶺プラント工業㈱", f(10), al("left"), B_ALL)
        merge(ws, 4, hc, 4, hc + 1, "職　長", f(9), al("center"), B_ALL, fill("F2F2F2"))
        merge(ws, 4, hc + 2, 4, ncol, "x1", f(10), al("left"), B_ALL)
        r = 6
        merge(ws, r, 1, r + 1, 2, "作業箇所", f(9, True), al("center"), B_ALL, fill("F2F2F2"))
        merge(ws, r, 3, r + 1, 4, "作　業　内　容", f(9, True), al("center"), B_ALL, fill("F2F2F2"))
        if form == "2014":
            merge(ws, r, 5, r, 11, "人　員（名）", f(9, True), al("center"), B_ALL, fill("F2F2F2"))
            for i, nm in enumerate(CREW_N): put(ws, r + 1, 5 + i, nm, f(7), al("center", shrink=True), B_ALL, fill("F2F2F2"))
            merge(ws, r, 12, r + 1, 12, "計", f(9, True), al("center"), B_ALL, fill("F2F2F2"))
            merge(ws, r, 13, r + 1, 13, "時間", f(9, True), al("center"), B_ALL, fill("F2F2F2"))
        else:
            merge(ws, r, 5, r + 1, 6, "人　員", f(9, True), al("center"), B_ALL, fill("F2F2F2"))
            merge(ws, r, 7, r + 1, 8, "時　間", f(9, True), al("center"), B_ALL, fill("F2F2F2"))
        r += 2
        tot = {k: 0 for k in CREW_K}; tall = 0
        own = [x for x in items if x[0].company != W.REFR]
        other = [x for x in items if x[0].company == W.REFR]
        prev_place = None
        lines = 0
        for i, (t, txt, crew) in enumerate(own + other):
            place = (f"{c.furn}号炉" if c.furn else "中央制御室") if t.company != W.REFR else "（丸越築炉）"
            if t.tags and W.BY_TAG[t.tags[0]].furn == 0 and c.kind != "dcs" and t.company != W.REFR:
                place = "共通"
            content = f"{t.name.split(' ')[0]}：{txt}" if t.tags and not txt.startswith(t.name.split(' ')[0]) else txt
            tm, hr = hours(d, txt, c)
            pv = "〃" if place == prev_place else place
            prev_place = place
            merge(ws, r, 1, r, 2, pv, f(8), al("center"), B_ALL)
            merge(ws, r, 3, r, 4, content, f(8), al("left", wrap=True), B_ALL)
            cc = dict(crew)
            if i == 0 and place != "（丸越築炉）": cc["職"] = 1
            ssum = sum(v for v in cc.values() if v)
            if form == "2014":
                for j, k in enumerate(CREW_K):
                    v = cc.get(k) or None
                    put(ws, r, 5 + j, v, f(9), al("center"), B_ALL)
                    if v: tot[k] += v
                put(ws, r, 12, ssum or None, f(9), al("center"), B_ALL)
                tmv = tm if tm != "8:00～17:00" else ("〃" if i else tm)
                put(ws, r, 13, tmv, f(7), al("center", shrink=True), B_ALL)
            else:
                merge(ws, r, 5, r, 6, f"{ssum}名" if ssum else None, f(9), al("center"), B_ALL)
                tmv = tm if (i == 0 or tm != "8:00～17:00") else "〃"
                merge(ws, r, 7, r, 8, tmv, f(8), al("center"), B_ALL)
            tall += ssum
            ws.row_dimensions[r].height = 30 if len(content) > 26 else 18
            fields = {"date": d.isoformat(), "weather": weather(d), "place": place if pv == "〃" else pv, "content": content,
                      "people": ssum, "time": tm, "manDays": round(ssum * hr / 8, 2), "manDaysText": f"{ssum}名" if form == "2009" else str(ssum)}
            if form == "2014":
                fields["crew"] = {CREW_N[j]: cc[k] for j, k in enumerate(CREW_K) if cc.get(k)}
            rec.add("item", fields, ws.title, r, wid(c, t))
            r += 1; lines += 1
        for _ in range(max(0, 10 - lines)):
            merge(ws, r, 1, r, 2, None, None, None, B_ALL); merge(ws, r, 3, r, 4, None, None, None, B_ALL)
            for j in range(5, ncol + 1): put(ws, r, j, None, None, None, B_ALL)
            ws.row_dimensions[r].height = 18; r += 1
        merge(ws, r, 1, r, 4, "合　　計", f(9, True), al("center"), B_ALL, fill("F2F2F2"))
        if form == "2014":
            for j, k in enumerate(CREW_K): put(ws, r, 5 + j, tot[k] or None, f(9, True), al("center"), B_ALL, fill("F2F2F2"))
            put(ws, r, 12, tall, f(9, True), al("center"), B_ALL, fill("F2F2F2")); put(ws, r, 13, None, None, None, B_ALL, fill("F2F2F2"))
        else:
            merge(ws, r, 5, r, 6, f"{tall}名", f(9, True), al("center"), B_ALL, fill("F2F2F2")); merge(ws, r, 7, r, 8, None, None, None, B_ALL, fill("F2F2F2"))
        rec.add("total", {"date": d.isoformat(), "people": tall}, ws.title, r)
        r += 2
        if form == "2014":
            ky = ["狭い炉内での挟まれ・転倒　→　声かけ・照明確保", "マンホール開放時の高温・粉じん　→　温度確認、防じんマスク",
                  "足場上の工具落下　→　工具ひも付け", "溶接作業の火気　→　消火器配置・監視人", "重量物の吊込み　→　合図者を決める"][(d.toordinal()) % 5]
            if c.cid == "T2015-2":
                ky = {8: "マンホール開放時の高温・粉じん　→　温度確認、防じんマスク", 9: "足場上の工具落下　→　工具ひも付け",
                      13: "雨天 屋外通路のすべり　→　足元確認", 15: "溶接作業の火気　→　消火器配置・監視人", 22: "珪砂投入時の粉じん　→　防じんマスク"}.get(d.day, "狭い炉内での挟まれ・転倒　→　声かけ・照明確保")
            merge(ws, r, 1, r, 2, "ＫＹ", f(9), al("center"), B_ALL, fill("F2F2F2")); merge(ws, r, 3, r, ncol, ky, f(8), al("left"), B_ALL)
            rec.add("item", {"date": d.isoformat(), "KY": ky}, ws.title, r); r += 1
            kz = "チェーンブロック、電動工具"
            if c.cid == "T2015-2":
                kz = {8: "エアホッパ、ジェットヒーター、送風機", 9: "高所作業車、エアブロー", 12: "超音波厚さ計、溶接機 2 台", 13: "超音波厚さ計、溶接機",
                      15: "溶接機 2 台、PT 薬剤", 16: "UT 探傷器、チェーンブロック", 22: "珪砂 フレコン 6 袋、クレーン"}.get(d.day, kz)
            elif any("肉厚" in x[1] for x in items): kz = "超音波厚さ計、チェーンブロック"
            elif any("溶接" in x[1] or "【追加】" in x[1] for x in items): kz = "溶接機、チェーンブロック"
            merge(ws, r, 1, r, 2, "使用機材", f(9), al("center"), B_ALL, fill("F2F2F2")); merge(ws, r, 3, r, ncol, kz, f(8), al("left"), B_ALL)
            rec.add("item", {"date": d.isoformat(), "equipmentUsed": kz}, ws.title, r); r += 1
        tk = tokki(c, d, items)
        merge(ws, r, 1, r + 1, 2, "特記事項", f(9), al("center"), B_ALL, fill("F2F2F2"))
        merge(ws, r, 3, r + 1, ncol, tk, f(8), al("left", "top", wrap=True), B_ALL)
        if tk: rec.add("item", {"date": d.isoformat(), "note": tk}, ws.title, r)
        ws.row_dimensions[r].height = 18; ws.row_dimensions[r + 1].height = 18
        r += 3
        sc = ncol - 4
        put(ws, r, sc, "SPC確認", f(8), al("center"), B_ALL); ws.merge_cells(start_row=r, start_column=sc, end_row=r, end_column=sc + 1)
        put(ws, r, sc + 2, "職長", f(8), al("center"), B_ALL); ws.merge_cells(start_row=r, start_column=sc + 2, end_row=r, end_column=sc + 3)
        merge(ws, r + 1, sc, r + 2, sc + 1, "sv1" if c.fy <= 2016 else "sv2", f(9), al("center"), B_ALL)
        merge(ws, r + 1, sc + 2, r + 2, sc + 3, "x1", f(9), al("center"), B_ALL)
        if form == "2014":
            widths(ws, {"A": 5, "B": 5, "C": 22, "D": 22, "E": 4.5, "F": 4.5, "G": 4.5, "H": 4.5, "I": 4.5, "J": 4.5, "K": 4.5, "L": 4.5, "M": 9})
        else:
            widths(ws, {"A": 5, "B": 5, "C": 26, "D": 26, "E": 5, "F": 5, "G": 6, "H": 9})
        page(ws, "portrait", 9, 1, 1, f"A1:{L(ncol)}{r + 2}")
    props(wb, f"【架空】ときわ北部CC {c.short} 作業日報（{c.fy}）", "x1（【架空】北嶺プラント工業）", dt.datetime(c.end.year, c.end.month, c.end.day, 9))
    wb.save(path); rec.save()
    return len(days)

# ============================================================== x1 2024 様式（横）
def x1_2024(c, path, fname):
    rec = Rec(fname)
    from contracts import B24
    wb = Workbook(); wb.remove(wb.active)
    days = days_of(c)
    total_steps = {t.tid: len(t.steps) for t in c.tasks}
    done = {}
    for n, (d, items) in enumerate(days.items(), 1):
        ws = wb.create_sheet(f"{d.month}.{d.day}")
        merge(ws, 1, 1, 1, 14, "【架空】作　業　日　報", f(16, True), al("center"), None)
        put(ws, 1, 15, f"No.{n}", f(9), al("right"))
        put(ws, 2, 1, "工事件名", f(9), al("center"), B_ALL, fill("F2F2F2"))
        merge(ws, 2, 2, 2, 5, c.title.replace("ときわ北部クリーンセンター", "ときわ北部ＣＣ"), f(10), al("left"), B_ALL)
        put(ws, 2, 6, "日　付", f(9), al("center"), B_ALL, fill("F2F2F2"))
        merge(ws, 2, 7, 2, 15, f"令和{d.year - 2018}年{d.month}月{d.day}日（{YOUBI[d.weekday()]}）", f(10), al("left"), B_ALL)
        put(ws, 3, 1, "施工会社", f(9), al("center"), B_ALL, fill("F2F2F2"))
        merge(ws, 3, 2, 3, 5, "北嶺プラント工業㈱　職長 x1", f(10), al("left"), B_ALL)
        put(ws, 3, 6, "天候", f(9), al("center"), B_ALL, fill("F2F2F2"))
        wt = f"{['晴', '晴', '曇', '雨', '晴'][d.day % 5]}　{18 + d.day % 7}℃"
        merge(ws, 3, 7, 3, 15, wt, f(10), al("left"), B_ALL)
        r = 5
        for i, h in enumerate(["班", "作業箇所（機器）", "作業内容", "人数", "工数\n(人･h)", "進捗\n(%)", "写真No.", "備考"]):
            put(ws, r, 1 + i, h, f(9, True), al("center", wrap=True), B_ALL, fill("DDEBF7"))
        for j, nm in enumerate(CREW_N):
            put(ws, r, 9 + j, nm, f(7), al("center", wrap=True), B_ALL, fill("F2F2F2"))
        put(ws, r - 1, 9, "職種別人員（名）", f(8, True), al("left"))
        ws.row_dimensions[r].height = 28
        r += 1
        tp = th = photo = 0; ctot = {}
        prev_team = None
        for t, txt, crew in sorted(items, key=lambda x: (x[0].team, x[0].tid)):
            done[t.tid] = done.get(t.tid, 0) + 1
            tm, hr = hours(d, txt, c)
            ppl = sum(crew.values())
            team = t.team if c.cid == "T2025-1" else B24.get(t.team, t.team)
            if t.company == W.REFR:
                place = f"{c.furn}号 旋回溶融炉（丸越築炉）"
            elif t.tags:
                place = "、".join(("共通 " if W.BY_TAG[x].furn == 0 else f"{W.BY_TAG[x].furn}号 ") + W.BY_TAG[x].name for x in t.tags[:2])
                if len(t.tags) > 2: place += f" ほか{len(t.tags) - 2}台"
            else:
                place = f"{c.furn}号炉 全般"
            prog = round(done[t.tid] / total_steps[t.tid] * 100) if total_steps[t.tid] else None
            if c.cid == "T2025-1" and t.tid == "T05" and d <= dt.date(2025, 5, 27):
                prog = {19: 15, 20: 30, 21: 45, 22: 55, 23: 65, 24: 80, 26: 95, 27: 100}[d.day]
            photo += 1 if t.tags else 0
            tv = team if team != prev_team else "〃"
            prev_team = team
            put(ws, r, 1, tv, f(8), al("center"), B_ALL)
            put(ws, r, 2, place, f(8), al("left", wrap=True), B_ALL)
            put(ws, r, 3, txt, f(8), al("left", wrap=True), B_ALL)
            put(ws, r, 4, ppl, f(9), al("center"), B_ALL)
            put(ws, r, 5, ppl * hr, f(9), al("center"), B_ALL, fmt="0.0" if hr % 1 else "0")
            put(ws, r, 6, prog, f(9), al("center"), B_ALL)
            ph = f"{d.month}{d.day:02d}-{photo:02d}" if t.tags else None
            put(ws, r, 7, ph, f(7), al("center"), B_ALL)
            bk = "残業 " + str(hr - 8).rstrip("0").rstrip(".") + "h" if hr > 8 else ("ケージ予備品 3本使用" if "ケージ 3 本交換" in txt else None)
            put(ws, r, 8, bk, f(7), al("left", shrink=True), B_ALL)
            for j, k in enumerate(CREW_K):
                put(ws, r, 9 + j, crew.get(k) or None, f(8), al("center"), B_ALL)
                ctot[k] = ctot.get(k, 0) + (crew.get(k) or 0)
            ws.row_dimensions[r].height = 30
            rec.add("item", {"date": d.isoformat(), "weather": wt, "team": team if tv == "〃" else tv, "place": place, "content": txt, "people": ppl,
                             "manDays": round(ppl * hr / 8, 2), "manDaysText": f"{ppl * hr:g}", "progress": prog, "photo": ph, "memo": bk,
                             "crew": {CREW_N[j]: crew[k] for j, k in enumerate(CREW_K) if crew.get(k)}}, ws.title, r, wid(c, t))
            tp += ppl; th += ppl * hr; r += 1
        put(ws, r, 1, "－", f(8), al("center"), B_ALL); put(ws, r, 2, "職長", f(8), al("left"), B_ALL)
        put(ws, r, 3, "作業指示・安全管理・SPC 打合せ", f(8), al("left"), B_ALL)
        put(ws, r, 4, 1, f(9), al("center"), B_ALL); put(ws, r, 5, 8, f(9), al("center"), B_ALL)
        for cc in range(6, 9): put(ws, r, cc, None, None, None, B_ALL)
        for j, k in enumerate(CREW_K): put(ws, r, 9 + j, 1 if k == "職" else None, f(8), al("center"), B_ALL)
        ctot["職"] = ctot.get("職", 0) + 1
        rec.add("item", {"date": d.isoformat(), "place": "職長", "content": "作業指示・安全管理・SPC 打合せ", "people": 1, "manDays": 1.0, "manDaysText": "8", "crew": {"職長": 1}}, ws.title, r)
        tp += 1; th += 8; r += 1
        for _ in range(max(0, 9 - len(items))):
            for cc in range(1, 16): put(ws, r, cc, None, None, None, B_ALL)
            ws.row_dimensions[r].height = 22; r += 1
        merge(ws, r, 1, r, 3, "合　　計", f(9, True), al("center"), B_ALL, fill("DDEBF7"))
        put(ws, r, 4, tp, f(9, True), al("center"), B_ALL, fill("DDEBF7"))
        put(ws, r, 5, th, f(9, True), al("center"), B_ALL, fill("DDEBF7"), fmt="0.0" if th % 1 else "0")
        for cc in range(6, 9): put(ws, r, cc, None, None, None, B_ALL, fill("DDEBF7"))
        for j, k in enumerate(CREW_K): put(ws, r, 9 + j, ctot.get(k) or None, f(9, True), al("center"), B_ALL, fill("DDEBF7"))
        rec.add("total", {"date": d.isoformat(), "people": tp, "manDays": round(th / 8, 2), "manDaysText": f"{th:g}"}, ws.title, r)
        r += 2
        tk = tokki(c, d, items)
        merge(ws, r, 1, r + 1, 1, "特記", f(9), al("center"), B_ALL, fill("F2F2F2"))
        merge(ws, r, 2, r + 1, 15, tk, f(8), al("left", "top", wrap=True), B_ALL)
        if tk: rec.add("item", {"date": d.isoformat(), "note": tk}, ws.title, r)
        r += 2
        put(ws, r, 1, "KY", f(8), al("center"), B_ALL, fill("F2F2F2"))
        merge(ws, r, 2, r, 6, "炉内作業の酸欠・粉じん → 測定・マスク／高所 → 安全帯", f(7), al("left"), B_ALL)
        put(ws, r, 7, "確認 sv2", f(8), al("center"), B_ALL); put(ws, r, 8, "記入 x1", f(8), al("center"), B_ALL)
        widths(ws, {"A": 8, "B": 22, "C": 40, "D": 5, "E": 7, "F": 6, "G": 8, "H": 16, "I": 4.5, "J": 4.5, "K": 4.5, "L": 4.5, "M": 4.5, "N": 4.5, "O": 4.5})
        page(ws, "landscape", 9, 1, 1, f"A1:O{r}")
    props(wb, f"【架空】ときわ北部CC {c.fy}年度 {c.short} 作業日報", "x1（【架空】北嶺プラント工業）", dt.datetime(c.end.year, c.end.month, c.end.day, 9), "x1")
    wb.save(path); rec.save()
    return len(days)

# ============================================================== メーカー週報
def maker(c, path, fname):
    rec = Rec(fname)
    wb = Workbook(); wb.remove(wb.active)
    rnd = random.Random(W.SEED + 3 + c.fy)
    w0 = c.start - dt.timedelta(days=c.start.weekday())
    last = min(c.end, W.TODAY)
    nweeks = ((last - w0).days // 7) + 1
    weeks = [w0 + dt.timedelta(weeks=i) for i in range(nweeks)]
    if c.fy == 2025:
        holidays = set(K.K3_HOLIDAYS) | {dt.date(2025, 9, 15), dt.date(2025, 9, 23), dt.date(2025, 11, 3), dt.date(2026, 1, 12)}
        INSP = {dt.date(2025, 12, 22): "水圧試験 u1立会", dt.date(2025, 12, 23): "使用前自主検査", dt.date(2025, 10, 9): "鉄皮補修 組合確認",
                dt.date(2025, 11, 19): "築炉出来形 u1", dt.date(2026, 1, 16): "引渡し確認", dt.date(2025, 11, 12): "エコ RT 結果確認"}
        EVENT = {dt.date(2025, 10, 21): "過熱器管材 納期遅れ連絡（約3週）→ u1 報告", dt.date(2025, 10, 1): "鉄皮 減肉範囲 想定より広い → 当て板追加 協議",
                 dt.date(2025, 11, 7): "過熱器管材 搬入", dt.date(2025, 12, 5): "変更契約（工期延長）締結"}
    else:
        holidays = {dt.date(2026, 9, 21), dt.date(2026, 9, 22), dt.date(2026, 9, 23), dt.date(2026, 10, 12), dt.date(2026, 11, 3)}
        INSP = {dt.date(2026, 10, 1): "鉄皮補修 組合確認", dt.date(2026, 11, 6): "エコ RT 結果確認"}
        EVENT = {dt.date(2026, 9, 30): "鉄皮 減肉 事前調査どおり。当て板 2 日延長", dt.date(2026, 9, 14): "過熱器管材 搬入（前倒し手配）"}
    cum = 0
    for wi, w in enumerate(weeks):
        days = [w + dt.timedelta(days=k) for k in range(7)]
        ws = wb.create_sheet(f"{wi + 1}w({w.month}.{w.day})")
        nc = 11
        merge(ws, 1, 1, 1, nc, "【架空】工　事　週　報", f(15, True), al("center"), None)
        merge(ws, 2, 1, 2, nc, f"工事名：ときわ北部クリーンセンター基幹的設備改良工事（令和7〜9年度）　{'令和7' if c.fy == 2025 else '令和8'}年度施工分（{c.furn}号炉）", f(9), al("left"), None)
        merge(ws, 3, 1, 3, nc, f"第{wi + 1}週　{wareki_date(w, 'H')}〜{wareki_date(days[-1], 'H')}　受注者：鷹羽エンジニアリング㈱　現場代理人 m1　記入 m2", f(9), al("left"), None)
        for cc, h in ((1, "工種"), (2, "作業"), (3, "対象")):
            merge(ws, 5, cc, 6, cc, h, f(9, True), al("center"), B_ALL, fill("DEEAF6"))
        for k, d in enumerate(days):
            off = d.weekday() == 6 or d in holidays
            put(ws, 5, 4 + k, f"{d.month}/{d.day}", f(9, True), al("center"), B_ALL, fill("F2DCDB") if off else fill("DEEAF6"))
            put(ws, 6, 4 + k, YOUBI[d.weekday()] + ("　休" if off else ""), f(8), al("center"), B_ALL, fill("F2DCDB") if off else fill("DEEAF6"))
        merge(ws, 5, nc, 6, nc, "週の出来高・記事", f(9, True), al("center", wrap=True), B_ALL, fill("DEEAF6"))
        r = 7
        active = [(n_, t) for n_, t in enumerate(c.tasks) if t[6] <= days[-1] and t[7] >= w and t[6] <= last]
        mk = [0] * 7; sc = [0] * 7
        prev = None
        for n_, (kou, name, tags, attr, ps, pe, as_, ae, crew) in active:
            pool = PHRASES.get(name, [name]); span = (ae - as_).days + 1
            put(ws, r, 1, kou if kou != prev else "〃", f(8, kou != prev), al("center"), B_ALL); prev = kou
            put(ws, r, 2, name, f(8), al("left", wrap=True), B_ALL)
            put(ws, r, 3, "\n".join(tags) if tags else "—", f(7), al("center", wrap=True), B_ALL)
            last_ph = None; daily = {}
            for k, d in enumerate(days):
                off = d.weekday() == 6 or d in holidays or d > last
                txt = None
                if as_ <= d <= ae and not off:
                    ph = pool[min(len(pool) - 1, (d - as_).days * len(pool) // span)]
                    txt = ph if ph != last_ph else "〃"; last_ph = ph; daily[d.isoformat()] = ph
                    if crew:
                        n2 = crew if d.weekday() != 5 else max(1, crew // 2)
                        n2 = max(1, n2 + rnd.randint(-2, 2)) if n2 > 3 else n2
                        m2 = max(1, round(n2 * 0.2)); mk[k] += m2; sc[k] += n2 - m2
                    else:
                        mk[k] = max(mk[k], 2)
                put(ws, r, 4 + k, txt, f(7), al("center" if txt == "〃" else "left", "center", True), B_ALL, fill("F2DCDB") if (d.weekday() == 6 or d in holidays) else None)
            done = min(100, round(((min(days[-1], ae, last) - as_).days + 1) / span * 100))
            note = f"出来高 {done}%" + ("（完了）" if ae <= min(days[-1], last) else "")
            if (as_, ae) != (ps, pe) and as_ <= days[-1]: note += " 予定から変更"
            put(ws, r, nc, note, f(7), al("left", wrap=True), B_ALL)
            ws.row_dimensions[r].height = 30
            wk = [f"{c.cid}:K{n_:02d}:{t}" for t in tags] or [f"{c.cid}:K{n_:02d}"]
            rec.add("item", {"week": wi + 1, "kou": kou, "name": name, "equipment": " ".join(tags), "daily": daily, "progress": done, "memo": note}, ws.title, r, wk)
            r += 1
        for _ in range(max(0, 8 - len(active))):
            for cc in range(1, nc + 1): put(ws, r, cc, None, None, None, B_ALL)
            ws.row_dimensions[r].height = 22; r += 1
        for lab, sub, vals in (("人員", "鷹羽", mk), ("", "協力会社", sc), ("", "計", [a + b for a, b in zip(mk, sc)])):
            put(ws, r, 1, lab, f(9, True), al("center"), B_ALL, fill("F2F2F2"))
            merge(ws, r, 2, r, 3, sub, f(8), al("center"), B_ALL, fill("F2F2F2"))
            for k in range(7): put(ws, r, 4 + k, vals[k] or None, f(9, sub == "計"), al("center"), B_ALL)
            put(ws, r, nc, f"週計 {sum(vals)}" + (f"（累計 {cum + sum(vals)}）" if sub == "計" else ""), f(8, sub == "計"), al("left"), B_ALL)
            rec.add("total" if sub == "計" else "subtotal", {"week": wi + 1, "row": sub, "people": {days[k].isoformat(): vals[k] for k in range(7) if vals[k]},
                                                             "manDays": sum(vals), "manDaysText": f"週計 {sum(vals)}"}, ws.title, r)
            r += 1
        cum += sum(mk) + sum(sc)
        for lab, src in (("天候", None), ("安全指示", None), ("立会・検査", INSP), ("記事", EVENT)):
            merge(ws, r, 1, r, 3, lab, f(8, True), al("center"), B_ALL, fill("F2F2F2"))
            vals = {}
            for k, d in enumerate(days):
                off = d.weekday() == 6 or d in holidays or d > last
                if lab == "天候": v = None if off else ["晴", "曇", "晴", "雨", "晴", "曇"][d.toordinal() % 6]
                elif lab == "安全指示": v = None if off else ["吊込み 合図者固定", "炉内 酸欠測定", "火気 監視人", "高所 二丁掛け", "粉じん マスク"][d.toordinal() % 5]
                else: v = src.get(d)
                if v: vals[d.isoformat()] = v
                put(ws, r, 4 + k, v, f(7), al("left", "center", True), B_ALL, fill("F2DCDB") if (d.weekday() == 6 or d in holidays) else None)
            put(ws, r, nc, None, None, None, B_ALL)
            if vals: rec.add("item", {"week": wi + 1, "row": lab, "daily": vals}, ws.title, r)
            ws.row_dimensions[r].height = 26 if lab in ("立会・検査", "記事") else 16
            r += 1
        widths(ws, {"A": 7, "B": 20, "C": 8, "K": 16})
        for k in range(7): ws.column_dimensions[L(4 + k)].width = 13
        page(ws, "landscape", 9, 1, 1, f"A1:{L(nc)}{r - 1}")
    props(wb, f"【架空】ときわ北部CC 基幹的設備改良工事 工事週報 {'R7' if c.fy == 2025 else 'R8'}", "m2（【架空】鷹羽エンジニアリング 工事部）",
          dt.datetime(last.year, last.month, last.day, 18))
    wb.save(path); rec.save()
    return len(weeks)
