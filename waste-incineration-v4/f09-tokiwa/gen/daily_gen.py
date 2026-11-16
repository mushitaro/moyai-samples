# -*- coding: utf-8 -*-
"""作業日報：x1 2014 様式（2015 2号炉）／x1 2024 様式（2025 1号炉）／メーカー週報型（基幹改良）"""
import datetime as dt, random
from openpyxl import Workbook
import world as W, works as K
from xlh import *

CREW_K = ["職", "機", "溶", "築", "と", "電", "作"]
CREW_N = ["職長", "機械工", "溶接工", "築炉工", "とび工", "電工", "作業員"]
WEATHER15 = {8: "晴", 9: "晴", 10: "曇", 12: "晴", 13: "雨", 14: "曇", 15: "晴", 16: "晴", 17: "晴", 19: "曇", 20: "晴",
             21: "晴", 22: "曇時々雨", 23: "晴", 24: "晴", 27: "晴", 28: "曇", 30: "晴"}

def hours_for(d, txt, contract):
    if "【追加】" in txt:
        return "8:00～19:00", 10
    if contract == "K2" and d == dt.date(2025, 5, 23):
        return "8:00～18:30", 9.5
    return "8:00～17:00", 8

# ============================================================== x1 2014 様式（縦・1日1枚）
def x1_2015(path):
    wb = Workbook(); wb.remove(wb.active)
    days = K.k1k2_days(K.K1)
    for n, (d, items) in enumerate(days.items(), 1):
        ws = wb.create_sheet(f"{d.month}.{d.day}")
        merge(ws, 1, 1, 1, 12, "【架空】作　業　日　報", f(16, True), al("center"), None)
        put(ws, 1, 13, f"No.{n}", f(9), al("right"))
        merge(ws, 2, 1, 2, 2, "工事件名", f(9), al("center"), B_ALL, fill("F2F2F2"))
        merge(ws, 2, 3, 2, 13, "平成27年度　ときわ北部ＣＣ　2号炉定期整備工事", f(10), al("left"), B_ALL)
        merge(ws, 3, 1, 3, 2, "日　付", f(9), al("center"), B_ALL, fill("F2F2F2"))
        merge(ws, 3, 3, 3, 7, f"H27年{d.month}月{d.day}日（{YOUBI[d.weekday()]}）", f(10), al("left"), B_ALL)
        merge(ws, 3, 8, 3, 9, "天　候", f(9), al("center"), B_ALL, fill("F2F2F2"))
        merge(ws, 3, 10, 3, 13, WEATHER15.get(d.day, "晴"), f(10), al("left"), B_ALL)
        merge(ws, 4, 1, 4, 2, "施工会社", f(9), al("center"), B_ALL, fill("F2F2F2"))
        merge(ws, 4, 3, 4, 7, "北嶺プラント工業㈱", f(10), al("left"), B_ALL)
        merge(ws, 4, 8, 4, 9, "職　長", f(9), al("center"), B_ALL, fill("F2F2F2"))
        merge(ws, 4, 10, 4, 13, "x1", f(10), al("left"), B_ALL)
        # 表
        r = 6
        merge(ws, r, 1, r + 1, 2, "作業箇所", f(9, True), al("center"), B_ALL, fill("F2F2F2"))
        merge(ws, r, 3, r + 1, 4, "作　業　内　容", f(9, True), al("center"), B_ALL, fill("F2F2F2"))
        merge(ws, r, 5, r, 11, "人　員（名）", f(9, True), al("center"), B_ALL, fill("F2F2F2"))
        for i, nm in enumerate(CREW_N):
            put(ws, r + 1, 5 + i, nm, f(7), al("center", shrink=True), B_ALL, fill("F2F2F2"))
        merge(ws, r, 12, r + 1, 12, "計", f(9, True), al("center"), B_ALL, fill("F2F2F2"))
        merge(ws, r, 13, r + 1, 13, "時間", f(9, True), al("center"), B_ALL, fill("F2F2F2"))
        r += 2
        tot = {k: 0 for k in CREW_K}
        prev_place = None
        rows_used = 0
        # 職長 1 名を先頭行に
        lines = []
        own = [(t, txt, crew) for t, txt, crew in items if t.company == W.SUB1 or t.company == "SPC運転"]
        other = [(t, txt, crew) for t, txt, crew in items if t.company == W.REFR]
        for t, txt, crew in own + other:
            place = "2号炉" if t.company != W.REFR else "（丸越築炉）"
            lines.append((place, f"{t.name.split(' ')[0]}：{txt}" if t.tags else txt, crew, hours_for(d, txt, "K1")))
        for i, (place, txt, crew, (tm, hr)) in enumerate(lines):
            pv = "〃" if place == prev_place else place
            prev_place = place
            merge(ws, r, 1, r, 2, pv, f(8), al("center"), B_ALL)
            merge(ws, r, 3, r, 4, txt, f(8), al("left", wrap=True), B_ALL)
            c = dict(crew)
            if i == 0 and place == "2号炉":
                c["職"] = 1
            s = 0
            for j, k in enumerate(CREW_K):
                v = c.get(k) or None
                put(ws, r, 5 + j, v, f(9), al("center"), B_ALL)
                if v: tot[k] += v; s += v
            put(ws, r, 12, s or None, f(9), al("center"), B_ALL)
            put(ws, r, 13, tm if tm != "8:00～17:00" else "〃" if i else tm, f(7), al("center", shrink=True), B_ALL)
            ws.row_dimensions[r].height = 30 if len(txt) > 26 else 18
            r += 1; rows_used += 1
        for _ in range(max(0, 10 - rows_used)):
            merge(ws, r, 1, r, 2, None, None, None, B_ALL); merge(ws, r, 3, r, 4, None, None, None, B_ALL)
            for j in range(5, 14): put(ws, r, j, None, None, None, B_ALL)
            ws.row_dimensions[r].height = 18
            r += 1
        merge(ws, r, 1, r, 4, "合　　計", f(9, True), al("center"), B_ALL, fill("F2F2F2"))
        for j, k in enumerate(CREW_K):
            put(ws, r, 5 + j, tot[k] or None, f(9, True), al("center"), B_ALL, fill("F2F2F2"))
        put(ws, r, 12, sum(tot.values()), f(9, True), al("center"), B_ALL, fill("F2F2F2"))
        put(ws, r, 13, None, None, None, B_ALL, fill("F2F2F2"))
        r += 2
        ky = {8: "マンホール開放時の高温・粉じん　→　温度確認、防じんマスク",
              9: "足場上の工具落下　→　工具ひも付け", 13: "雨天 屋外通路のすべり　→　足元確認", 15: "溶接作業の火気　→　消火器配置・監視人",
              22: "珪砂投入時の粉じん　→　防じんマスク"}.get(d.day, "狭い炉内での挟まれ・転倒　→　声かけ・照明確保")
        merge(ws, r, 1, r, 2, "ＫＹ", f(9), al("center"), B_ALL, fill("F2F2F2"))
        merge(ws, r, 3, r, 13, ky, f(8), al("left"), B_ALL); r += 1
        kizai = {8: "エアホッパ、ジェットヒーター、送風機", 9: "高所作業車、エアブロー", 12: "超音波厚さ計、溶接機 2 台", 13: "超音波厚さ計、溶接機",
                 15: "溶接機 2 台、PT 薬剤", 16: "UT 探傷器、チェーンブロック", 22: "珪砂 フレコン 6 袋、クレーン"}.get(d.day, "チェーンブロック、電動工具")
        merge(ws, r, 1, r, 2, "使用機材", f(9), al("center"), B_ALL, fill("F2F2F2"))
        merge(ws, r, 3, r, 13, kizai, f(8), al("left"), B_ALL); r += 1
        tok = {13: "過熱器下部ベンド No.14　肉厚 2.6mm（基準 3.2mm）　sv1 さんへ報告。補修方法は明日指示。",
               14: "定期事業者検査 内部 終了。過熱器 No.14 は肉盛補修の指示あり（追加）。",
               15: "肉盛溶接のため溶接工 2 名 19:00 まで残業。",
               10: "不燃物抜出スクリュー 工場へ搬出（15日戻り予定）。",
               19: "ボイラ閉止 肉盛の影響で 1 日遅れ。水圧試験 21日に変更（sv1 了承）。"}.get(d.day, "")
        merge(ws, r, 1, r + 1, 2, "特記事項", f(9), al("center"), B_ALL, fill("F2F2F2"))
        merge(ws, r, 3, r + 1, 13, tok, f(8), al("left", "top", wrap=True), B_ALL)
        ws.row_dimensions[r].height = 18; ws.row_dimensions[r + 1].height = 18
        r += 3
        put(ws, r, 9, "SPC確認", f(8), al("center"), B_ALL); ws.merge_cells(start_row=r, start_column=9, end_row=r, end_column=10)
        put(ws, r, 11, "職長", f(8), al("center"), B_ALL); ws.merge_cells(start_row=r, start_column=11, end_row=r, end_column=12)
        merge(ws, r + 1, 9, r + 2, 10, "sv1", f(9), al("center"), B_ALL)
        merge(ws, r + 1, 11, r + 2, 12, "x1", f(9), al("center"), B_ALL)
        widths(ws, {"A": 5, "B": 5, "C": 22, "D": 22, "E": 4.5, "F": 4.5, "G": 4.5, "H": 4.5, "I": 4.5, "J": 4.5, "K": 4.5, "L": 4.5, "M": 9})
        page(ws, "portrait", 9, 1, 1, f"A1:M{r + 2}")
    props(wb, "【架空】ときわ北部CC 2号炉定期整備工事 作業日報（H27）", "x1（【架空】北嶺プラント工業）", dt.datetime(2015, 10, 31, 9, 0))
    wb.save(path)
    return len(days)

# ============================================================== x1 2024 様式（横・1日1枚・班別・工数）
def x1_2025(path):
    wb = Workbook(); wb.remove(wb.active)
    days = K.k1k2_days(K.K2)
    total_steps = {t.tid: len(t.steps) for t in K.K2}
    done = {}
    rnd = random.Random(W.SEED + 25)
    for n, (d, items) in enumerate(days.items(), 1):
        ws = wb.create_sheet(f"{d.month}.{d.day}")
        merge(ws, 1, 1, 1, 14, "【架空】作　業　日　報", f(16, True), al("center"), None)
        put(ws, 1, 15, f"No.{n}", f(9), al("right"))
        put(ws, 2, 1, "工事件名", f(9), al("center"), B_ALL, fill("F2F2F2"))
        merge(ws, 2, 2, 2, 5, "2025年度　ときわ北部ＣＣ　1号炉定期整備工事", f(10), al("left"), B_ALL)
        put(ws, 2, 6, "日　付", f(9), al("center"), B_ALL, fill("F2F2F2"))
        merge(ws, 2, 7, 2, 15, f"令和7年{d.month}月{d.day}日（{YOUBI[d.weekday()]}）", f(10), al("left"), B_ALL)
        put(ws, 3, 1, "施工会社", f(9), al("center"), B_ALL, fill("F2F2F2"))
        merge(ws, 3, 2, 3, 5, "北嶺プラント工業㈱　職長 x1", f(10), al("left"), B_ALL)
        put(ws, 3, 6, "天候", f(9), al("center"), B_ALL, fill("F2F2F2"))
        merge(ws, 3, 7, 3, 15, f"{['晴', '晴', '曇', '雨', '晴'][d.day % 5]}　{18 + d.day % 7}℃", f(10), al("left"), B_ALL)
        r = 5
        heads = ["班", "作業箇所（機器）", "作業内容", "人数", "工数\n(人･h)", "進捗\n(%)", "写真No.", "備考", ""]
        for i, h in enumerate(heads[:8]):
            put(ws, r, 1 + i, h, f(9, True), al("center", wrap=True), B_ALL, fill("DDEBF7"))
        # 2014 様式の職種別人員を右端に残し、そこから人数・工数を出す（x1 の書き癖）
        for j, nm in enumerate(CREW_N):
            put(ws, r, 9 + j, nm, f(7), al("center", wrap=True), B_ALL, fill("F2F2F2"))
        put(ws, r - 1, 9, "職種別人員（名）", f(8, True), al("left"))
        ws.row_dimensions[r].height = 28
        r += 1
        tp = 0; th = 0; photo = 0; ctot = {}
        items_sorted = sorted(items, key=lambda x: (x[0].team, x[0].tid))
        prev_team = None
        for t, txt, crew in items_sorted:
            done[t.tid] = done.get(t.tid, 0) + 1
            tm, hr = hours_for(d, txt, "K2")
            ppl = sum(crew.values())
            if t.tid == "T03":
                place = "1号 旋回溶融炉（丸越築炉）"
            elif t.tags:
                place = "、".join(("共通 " if W.BY_TAG[x].furn == 0 else "1号 ") + W.BY_TAG[x].name for x in t.tags[:2])
            else:
                place = "1号炉 全般"
            prog = round(done[t.tid] / total_steps[t.tid] * 100) if total_steps[t.tid] else None
            if t.tid == "T05" and d <= dt.date(2025, 5, 27):
                prog = {19: 15, 20: 30, 21: 45, 22: 55, 23: 65, 24: 80, 26: 95, 27: 100}[d.day]
            photo += 1 if t.tags else 0
            put(ws, r, 1, t.team if t.team != prev_team else "〃", f(8), al("center"), B_ALL)
            prev_team = t.team
            put(ws, r, 2, place, f(8), al("left", wrap=True), B_ALL)
            put(ws, r, 3, txt, f(8), al("left", wrap=True), B_ALL)
            put(ws, r, 4, ppl, f(9), al("center"), B_ALL)
            put(ws, r, 5, ppl * hr, f(9), al("center"), B_ALL, fmt="0.0" if hr % 1 else "0")
            put(ws, r, 6, prog, f(9), al("center"), B_ALL)
            put(ws, r, 7, f"{d.month}{d.day:02d}-{photo:02d}" if t.tags else None, f(7), al("center"), B_ALL)
            put(ws, r, 8, "残業 1.5h" if hr > 8 else ("ケージ予備品 3本使用" if "ケージ 3 本交換" in txt else None), f(7), al("left", shrink=True), B_ALL)
            for j, k in enumerate(CREW_K):
                put(ws, r, 9 + j, crew.get(k) or None, f(8), al("center"), B_ALL)
                ctot[k] = ctot.get(k, 0) + (crew.get(k) or 0)
            ws.row_dimensions[r].height = 30
            tp += ppl; th += ppl * hr
            r += 1
        # 職長は別行（2024 様式から工数に入れる）
        put(ws, r, 1, "－", f(8), al("center"), B_ALL); put(ws, r, 2, "職長", f(8), al("left"), B_ALL)
        put(ws, r, 3, "作業指示・安全管理・SPC 打合せ", f(8), al("left"), B_ALL)
        put(ws, r, 4, 1, f(9), al("center"), B_ALL); put(ws, r, 5, 8, f(9), al("center"), B_ALL)
        for c in range(6, 9): put(ws, r, c, None, None, None, B_ALL)
        for j, k in enumerate(CREW_K): put(ws, r, 9 + j, 1 if k == "職" else None, f(8), al("center"), B_ALL)
        ctot["職"] = ctot.get("職", 0) + 1
        tp += 1; th += 8; r += 1
        for _ in range(max(0, 9 - len(items))):
            for c in range(1, 16): put(ws, r, c, None, None, None, B_ALL)
            ws.row_dimensions[r].height = 22
            r += 1
        merge(ws, r, 1, r, 3, "合　　計", f(9, True), al("center"), B_ALL, fill("DDEBF7"))
        put(ws, r, 4, tp, f(9, True), al("center"), B_ALL, fill("DDEBF7"))
        put(ws, r, 5, th, f(9, True), al("center"), B_ALL, fill("DDEBF7"), fmt="0.0" if th % 1 else "0")
        for c in range(6, 9): put(ws, r, c, None, None, None, B_ALL, fill("DDEBF7"))
        for j, k in enumerate(CREW_K): put(ws, r, 9 + j, ctot.get(k) or None, f(9, True), al("center"), B_ALL, fill("DDEBF7"))
        r += 2
        tok = {22: "BF 第4室 リテーナ（ケージ）曲がり 3 本。sv2 さんへ報告、予備品で交換の了承を得た。",
               23: "ケージ交換のため BF 班 18:30 まで。BF 完了は 27日の見込み（1 日遅れ）。",
               20: "過熱器 No.9 ベンド 2.5mm。計算必要厚 2.1mm 以上のため継続使用（基幹改良 R8 で更新予定）。",
               16: "溶融炉 側壁下段 残存厚 最小 118mm（2023 年 2号炉 95mm）。",
               26: "破砕機 1号 停止、2号機で受入継続（SPC 運転班 了承）。"}.get(d.day if d.month == 5 else 0, "")
        merge(ws, r, 1, r + 1, 1, "特記", f(9), al("center"), B_ALL, fill("F2F2F2"))
        merge(ws, r, 2, r + 1, 15, tok, f(8), al("left", "top", wrap=True), B_ALL)
        r += 2
        put(ws, r, 1, "KY", f(8), al("center"), B_ALL, fill("F2F2F2"))
        merge(ws, r, 2, r, 6, "炉内作業の酸欠・粉じん → 測定・マスク／高所 → 安全帯", f(7), al("left"), B_ALL)
        put(ws, r, 7, "確認 sv2", f(8), al("center"), B_ALL); put(ws, r, 8, "記入 x1", f(8), al("center"), B_ALL)
        widths(ws, {"A": 8, "B": 22, "C": 40, "D": 5, "E": 7, "F": 6, "G": 8, "H": 16, "I": 4.5, "J": 4.5, "K": 4.5, "L": 4.5, "M": 4.5, "N": 4.5, "O": 4.5})
        page(ws, "landscape", 9, 1, 1, f"A1:O{r}")
    props(wb, "【架空】ときわ北部CC 2025年度 1号炉定期整備工事 作業日報", "x1（【架空】北嶺プラント工業）", dt.datetime(2025, 6, 7, 9, 0), "x1")
    wb.save(path)
    return len(days)

# ============================================================== メーカー 週報型
PHRASES = {
    "仮設・搬入路・資材置場": ["仮囲い設置、資材置場整地", "搬入路 敷鉄板", "現場事務所・休憩所設置", "仮設電源・仮設照明 配線", "揚重機 据付・点検"],
    "旋回溶融炉 耐火物解体": ["炉頂部 れんが解体", "上段側壁 解体・搬出", "中段側壁 解体（ブレーカ）", "下段側壁・出滓口 解体", "炉底 解体、ガラ搬出", "炉内清掃・鉄皮露出"],
    "鉄皮補修（スタッド・当て板）": ["鉄皮 肉厚測定（減肉範囲が想定より広い）", "当て板 溶接（追加分を含む）", "スタッド 植替え", "当て板・スタッド 検査"],
    "旋回溶融炉 耐火物築炉": ["炉底 キャスタブル打設", "下段側壁 れんが積み", "出滓口ブロック 据付", "中段側壁 れんが積み", "上段側壁 れんが積み", "炉頂 キャスタブル打設", "目地・仕上げ"],
    "二次燃焼室 耐火物更新": ["既設キャスタブル はつり", "アンカー溶接", "キャスタブル 吹付け", "養生・仕上がり確認"],
    "溶融炉バーナ 更新": ["既設バーナ 撤去", "新バーナ 据付", "燃料・空気配管 接続", "単体着火試験"],
    "炉体冷却水ポンプA・B 更新": ["既設ポンプ A 撤去", "A 据付・配管接続", "既設ポンプ B 撤去", "B 据付・配管接続", "芯出し・単体運転"],
    "過熱器 撤去": ["保温撤去、管寄せ切断", "過熱器管群 吊下ろし", "撤去品 搬出", "据付部 清掃・墨出し"],
    "過熱器 据付・溶接": ["管材受入れ・検品", "管寄せ 据付", "管群 吊込み", "現地溶接（RT 検査並行）", "支持金物 取付", "保温復旧"],
    "エコノマイザ 撤去": ["ケーシング開放", "管群 切断・吊下ろし", "撤去品 搬出"],
    "エコノマイザ 据付": ["管群 吊込み・据付", "現地溶接", "RT 検査", "ケーシング 復旧"],
    "ボイラ水圧試験・使用前検査": ["水張り、水圧試験（6.0MPa）", "使用前自主検査（組合 立会）"],
    "誘引通風機 撤去": ["電動機・軸受台 撤去", "ケーシング・羽根車 撤去搬出"],
    "誘引通風機 据付・芯出し": ["基礎補修・アンカー", "ケーシング 据付", "羽根車・軸受 据付", "電動機 据付・芯出し", "ダクト接続"],
    "押込送風機・燃焼用送風機 更新": ["押込送風機 撤去", "押込送風機 据付", "燃焼用送風機 撤去", "燃焼用送風機 据付", "ダクト接続・芯出し"],
    "インバータ盤 更新（誘引・押込）": ["既設盤 撤去", "新盤 据付", "ケーブル 接続", "単体試験"],
    "炉コントローラ盤 制御改造": ["I/O 増設", "ソフト改修 入替", "ループチェック", "インターロック試験"],
    "炉乾燥（溶融炉）": ["乾燥バーナ 設置", "昇温（150℃ 保持）", "昇温（350℃ 保持）", "昇温（600℃ 保持）・降温"],
    "総合試運転・性能確認": ["ごみ投入開始、負荷 50%", "負荷 80%", "定格負荷 確認", "排ガス測定", "スラグ出滓 確認", "性能確認 終了"],
    "炉立下げ（運営事業者）": ["2号炉 立下げ（運営事業者）"],
}

def maker_k3(path):
    """メーカーの週報：行＝工種・作業、列＝日（その日の作業・出来高を記入）。下に日ごとの人数・安全・立会"""
    wb = Workbook(); wb.remove(wb.active)
    rnd = random.Random(W.SEED + 3)
    w0 = dt.date(2025, 8, 18)
    weeks = [w0 + dt.timedelta(weeks=i) for i in range(22)]
    holidays = set(K.K3_HOLIDAYS) | {dt.date(2025, 9, 15), dt.date(2025, 9, 23), dt.date(2025, 10, 13), dt.date(2025, 11, 3), dt.date(2025, 11, 24), dt.date(2026, 1, 12)}
    holidays -= {dt.date(2025, 10, 13), dt.date(2025, 11, 24)}   # 山場の祝日は出勤
    INSP = {dt.date(2025, 12, 22): "水圧試験 u1立会", dt.date(2025, 12, 23): "使用前自主検査",
            dt.date(2025, 10, 9): "鉄皮補修 組合確認", dt.date(2025, 11, 19): "築炉出来形 u1",
            dt.date(2026, 1, 16): "引渡し確認", dt.date(2025, 11, 12): "エコ RT 結果確認"}
    EVENT = {dt.date(2025, 10, 21): "過熱器管材 納期遅れ連絡（約3週）→ u1 報告", dt.date(2025, 10, 1): "鉄皮 減肉範囲 想定より広い → 当て板追加 協議",
             dt.date(2025, 11, 7): "過熱器管材 搬入", dt.date(2025, 12, 5): "変更契約（工期延長）締結"}
    cum = 0
    for wi, w in enumerate(weeks):
        days = [w + dt.timedelta(days=k) for k in range(7)]
        ws = wb.create_sheet(f"{wi + 1}w({w.month}.{w.day})")
        nc = 3 + 7 + 1
        merge(ws, 1, 1, 1, nc, "【架空】工　事　週　報", f(15, True), al("center"), None)
        merge(ws, 2, 1, 2, nc, f"工事名：{W.FAC.replace('【架空】', '')}基幹的設備改良工事（令和7〜9年度）　令和7年度施工分（2号炉）", f(9), al("left"), None)
        merge(ws, 3, 1, 3, nc, f"第{wi + 1}週　{wareki_date(w, 'H')}〜{wareki_date(days[-1], 'H')}　受注者：鷹羽エンジニアリング㈱　現場代理人 m1　記入 m2", f(9), al("left"), None)
        merge(ws, 5, 1, 6, 1, "工種", f(9, True), al("center"), B_ALL, fill("DEEAF6"))
        merge(ws, 5, 2, 6, 2, "作業", f(9, True), al("center"), B_ALL, fill("DEEAF6"))
        merge(ws, 5, 3, 6, 3, "対象", f(9, True), al("center"), B_ALL, fill("DEEAF6"))
        for k, d in enumerate(days):
            off = d.weekday() == 6 or d in holidays
            put(ws, 5, 4 + k, f"{d.month}/{d.day}", f(9, True), al("center"), B_ALL, fill("F2DCDB") if off else fill("DEEAF6"))
            put(ws, 6, 4 + k, YOUBI[d.weekday()] + ("　休" if off else ""), f(8), al("center"), B_ALL, fill("F2DCDB") if off else fill("DEEAF6"))
        merge(ws, 5, nc, 6, nc, "週の出来高・記事", f(9, True), al("center", wrap=True), B_ALL, fill("DEEAF6"))
        r = 7
        active = [t for t in K.K3 if t[6] <= days[-1] and t[7] >= w]
        mk = [0] * 7; sc = [0] * 7
        prev = None
        for kou, name, tags, attr, ps, pe, as_, ae, crew in active:
            pool = PHRASES.get(name, [name])
            span = (ae - as_).days + 1
            put(ws, r, 1, kou if kou != prev else "〃", f(8, kou != prev), al("center"), B_ALL)
            prev = kou
            put(ws, r, 2, name, f(8), al("left", wrap=True), B_ALL)
            put(ws, r, 3, "\n".join(tags) if tags else "—", f(7), al("center", wrap=True), B_ALL)
            last_ph = None
            for k, d in enumerate(days):
                off = d.weekday() == 6 or d in holidays
                txt = None
                if as_ <= d <= ae and not off:
                    pos = (d - as_).days
                    ph = pool[min(len(pool) - 1, pos * len(pool) // span)]
                    txt = ph if ph != last_ph else "〃"
                    last_ph = ph
                    if crew:
                        n_ = crew if d.weekday() != 5 else max(1, crew // 2)
                        n_ = max(1, n_ + rnd.randint(-2, 2)) if n_ > 3 else n_
                        m_ = max(1, round(n_ * 0.2))
                        mk[k] += m_; sc[k] += n_ - m_
                    else:
                        mk[k] = max(mk[k], 2)
                put(ws, r, 4 + k, txt, f(7), al("center" if txt == "〃" else "left", "center", True), B_ALL, fill("F2DCDB") if off else None)
            done = min(100, round(((min(days[-1], ae) - as_).days + 1) / span * 100))
            note = f"出来高 {done}%" + ("（完了）" if ae <= days[-1] else "")
            if (as_, ae) != (ps, pe) and as_ <= days[-1]:
                note += " 予定から変更"
            put(ws, r, nc, note, f(7), al("left", wrap=True), B_ALL)
            ws.row_dimensions[r].height = 30
            r += 1
        for _ in range(max(0, 8 - len(active))):
            for c in range(1, nc + 1): put(ws, r, c, None, None, None, B_ALL)
            ws.row_dimensions[r].height = 22; r += 1
        # 日ごとの人数・記事
        rows = [("人員", "鷹羽", mk), ("", "協力会社", sc), ("", "計", [a + b for a, b in zip(mk, sc)])]
        for lab, sub, vals in rows:
            put(ws, r, 1, lab, f(9, True), al("center"), B_ALL, fill("F2F2F2"))
            merge(ws, r, 2, r, 3, sub, f(8), al("center"), B_ALL, fill("F2F2F2"))
            for k in range(7):
                put(ws, r, 4 + k, vals[k] or None, f(9, sub == "計"), al("center"), B_ALL)
            put(ws, r, nc, f"週計 {sum(vals)}" + (f"（累計 {cum + sum(vals)}）" if sub == "計" else ""), f(8, sub == "計"), al("left"), B_ALL)
            r += 1
        cum += sum(mk) + sum(sc)
        for lab, src in (("天候", None), ("安全指示", None), ("立会・検査", INSP), ("記事", EVENT)):
            merge(ws, r, 1, r, 3, lab, f(8, True), al("center"), B_ALL, fill("F2F2F2"))
            for k, d in enumerate(days):
                off = d.weekday() == 6 or d in holidays
                if lab == "天候":
                    v = None if off else ["晴", "曇", "晴", "雨", "晴", "曇"][d.toordinal() % 6]
                elif lab == "安全指示":
                    v = None if off else ["吊込み 合図者固定", "炉内 酸欠測定", "火気 監視人", "高所 二丁掛け", "粉じん マスク"][d.toordinal() % 5]
                else:
                    v = src.get(d)
                put(ws, r, 4 + k, v, f(7), al("left", "center", True), B_ALL, fill("F2DCDB") if off else None)
            put(ws, r, nc, None, None, None, B_ALL)
            ws.row_dimensions[r].height = 26 if lab in ("立会・検査", "記事") else 16
            r += 1
        widths(ws, {"A": 7, "B": 20, "C": 8, "K": 16})
        for k in range(7): ws.column_dimensions[L(4 + k)].width = 13
        page(ws, "landscape", 9, 1, 1, f"A1:{L(nc)}{r - 1}")
    props(wb, "【架空】ときわ北部CC 基幹的設備改良工事 工事週報 R7", "m2（【架空】鷹羽エンジニアリング 工事部）", dt.datetime(2026, 1, 19, 18, 0))
    wb.save(path)
    return len(weeks)
