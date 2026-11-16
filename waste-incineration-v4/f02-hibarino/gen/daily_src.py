# -*- coding: utf-8 -*-
"""作業日報。
旧 = 協力会社 w1 が 2017 年に作った「1日1枚（縦）」様式。2024 年度は職長交代直後の w2 がそのまま使用。
新 = w2 が 2025 年度に作った「1週1枚（横・日別欄）」様式。
市追加工事 = 鷹羽エンジ（メーカー）の工事日報（1日1枚・職種別人員・品質記録）。"""
import datetime as dt
import os
import random
from openpyxl import Workbook
from openpyxl.utils import get_column_letter as L
import world as W
import contracts as K
from xlhelp import *

WEATHER = ["晴", "晴", "曇", "晴のち曇", "曇", "雨", "曇時々雨", "晴"]

NOTE_2024 = {
    dt.date(2024, 10, 3): "SPC 運転課：1号炉 冷却中（炉内 210℃）。縁切り板 挿入 煙道・飛灰の2箇所 SPC 立会",
    dt.date(2024, 10, 4): "13:30 炉内 60℃以下 確認（SPC 朝比奈様）。入炉許可 受領",
    dt.date(2024, 10, 9): "作動油 採取 1L → 分析依頼（結果 後日）",
    dt.date(2024, 10, 11): "BF ろ布 3本 抜取 → 強度試験へ発送（筑紫フィルタ）。新品3本 取付",
    dt.date(2024, 10, 12): "UT：過熱器 2次入口側 外面減肉 最小3.6mm（公称4.5mm）。SPC 小鳥遊様へ報告、写真撮影済。"
                           "灰押出装置 側板下部にも腐食あり → 当て板 1.5m2 追加（SPC 了承）",
    dt.date(2024, 10, 15): "火格子 支給品 2回目（40枚）本日入荷予定のところ 未着。SPC 購買に確認依頼",
    dt.date(2024, 10, 16): "ボイラ定期事業者検査（開放）実施、指摘なし。IDF ランナ 羽根入口側 摩耗 → 肉盛 追加 SPC と協議・了承。"
                           "火格子 2回目 10/17 入荷の回答",
    dt.date(2024, 10, 17): "火格子 2回目 40枚 入荷 15:00、外観・寸法 受入確認（SPC 立会）",
    dt.date(2024, 10, 19): "土曜 作業（火格子 遅れ分の取戻し）。燃焼ストーカ 全列 完了",
    dt.date(2024, 10, 21): "IDF 試運転 振動 2.1mm/s（交換前 4.6mm/s）。二次送風機 着手を 10/22 へ 1日繰下げ",
    dt.date(2024, 10, 25): "ボイラ水圧試験 4.5MPa 30分 漏れなし。検査員・SPC 立会",
    dt.date(2024, 10, 30): "インタロック確認 全項目 OK（SPC 立会）",
    dt.date(2024, 11, 2): "8:00 ごみ投入開始。IDF 振動 監視 異常なし",
    dt.date(2024, 11, 5): "定格負荷 確認、SPC へ引渡し。日報 終了",
}
KY = ["高所作業 安全帯 使用の徹底", "炉内 酸欠・粉じん 測定後 入炉", "重量物 吊り作業 合図者 1名", "グラインダ 火花 養生",
      "足場 昇降時 3点支持", "溶接 ヒューム 送風機 使用", "回転機 起動前 縁切り確認", "熱中症 予防 水分補給"]


def build_old(path):
    con = K.CONTRACTS["OH2024-1"]
    rows = K.daily_rows(K.T2024, 2024, seed_add=24)
    r_ = random.Random(W.SEED + 241)
    wb = Workbook()
    wb.remove(wb.active)
    props(wb, f"【架空】ひばり野クリーンセンター {con['title']} 作業日報", "野分工業 職長 w2", dt.datetime(2024, 11, 6, 17, 0))
    sheets = []
    for x, lines in rows.items():
        work = [l for l in lines if l[0][3] != "SPC運転"]
        if not work:
            continue
        ws = wb.create_sheet(f"{x.month}.{x.day:02d}")
        sheets.append(ws.title)
        for c, w in {1: 11, 2: 14, 3: 38, 4: 5, 5: 6, 6: 7, 7: 7}.items():
            ws.column_dimensions[L(c)].width = w
        ws.merge_cells("A1:G1")
        put(ws, 1, 1, "【架空】　作　業　日　報", font(16, True), CENTER)
        put(ws, 3, 1, "工事件名", font(9, True), CENTER, B_ALL, fill("F2F2F2"))
        ws.merge_cells("B3:C3")
        put(ws, 3, 2, con["title"], font(9), LEFT, B_ALL)
        put(ws, 3, 4, "日付", font(9, True), CENTER, B_ALL, fill("F2F2F2"))
        ws.merge_cells("E3:G3")
        put(ws, 3, 5, f"{x.year}年{x.month}月{x.day}日（{K.WD[x.weekday()]}）", font(9), CENTER, B_ALL)
        put(ws, 4, 1, "施設", font(9, True), CENTER, B_ALL, fill("F2F2F2"))
        ws.merge_cells("B4:C4")
        put(ws, 4, 2, "【架空】ひばり野クリーンセンター　1号炉", font(9), LEFT, B_ALL)
        put(ws, 4, 4, "天候", font(9, True), CENTER, B_ALL, fill("F2F2F2"))
        ws.merge_cells("E4:G4")
        put(ws, 4, 5, r_.choice(WEATHER), font(9), CENTER, B_ALL)
        put(ws, 5, 1, "作業時間", font(9, True), CENTER, B_ALL, fill("F2F2F2"))
        ws.merge_cells("B5:C5")
        long_ = any(l[3] > 8 for l in work)
        put(ws, 5, 2, "8:00 〜 17:00" + ("（一部 18:30 まで）" if long_ else ""), font(9), LEFT, B_ALL)
        put(ws, 5, 4, "職長", font(9, True), CENTER, B_ALL, fill("F2F2F2"))
        ws.merge_cells("E5:G5")
        put(ws, 5, 5, "森野", font(9), CENTER, B_ALL)
        hdr = ["班", "作業場所", "作業内容", "人員", "時間", "工数", "進捗"]
        for i, h in enumerate(hdr, 1):
            put(ws, 7, i, h, font(9, True), CENTER, B_ALL, fill("DDEBF7"))
        r = 8
        prevplace = None
        tm = th = 0
        for (t, desc, men, hrs) in work:
            crew = t[3]
            place = "1炉 " + t[1].replace("設備", "") if t[1] != "共通" else "1炉 全般"
            if t[0] in ("T15",):
                place = "1炉 給水ポンプ室"
            if t[0] in ("T20", "T21"):
                place = "1炉 通風機室"
            if t[0] in ("T26", "T03"):
                place = "1炉 電気室"
            ds = K.task_days(t, 2024)
            wds = [d for d in ds if d.weekday() != 6]
            k = wds.index(x) + 1 if x in wds else 0
            prog = round(100 * k / len(wds)) if wds else 0
            put(ws, r, 1, crew, font(9), CENTER, B_ALL)
            put(ws, r, 2, "〃" if place == prevplace else place, font(9), CENTER if place == prevplace else LEFT, B_ALL)
            prevplace = place
            put(ws, r, 3, desc, font(9), LEFT, B_ALL)
            put(ws, r, 4, men, font(9), CENTER, B_ALL)
            put(ws, r, 5, hrs, font(9), CENTER, B_ALL, None, "0.0")
            put(ws, r, 6, round(men * hrs, 1), font(9), RIGHT, B_ALL, None, "0.0")
            put(ws, r, 7, f"{prog}%" if prog < 100 else "完", font(9), CENTER, B_ALL)
            ws.row_dimensions[r].height = 26 if len(desc) > 22 else 18
            tm += men
            th += men * hrs
            r += 1
        while r < 20:
            box(ws, r, 1, r, 7)
            ws.row_dimensions[r].height = 18
            r += 1
        put(ws, r, 1, "合　計", font(9, True), CENTER, B_ALL, fill("F2F2F2"))
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        put(ws, r, 2, None, None, None, B_ALL)
        put(ws, r, 4, tm, font(9, True), CENTER, B_ALL)
        put(ws, r, 5, None, None, None, B_ALL)
        put(ws, r, 6, round(th, 1), font(9, True), RIGHT, B_ALL, None, "0.0")
        put(ws, r, 7, None, None, None, B_ALL)
        r += 2
        put(ws, r, 1, "特記事項", font(9, True), CENTER, B_ALL, fill("F2F2F2"))
        ws.merge_cells(start_row=r, start_column=2, end_row=r + 2, end_column=7)
        put(ws, r, 2, NOTE_2024.get(x, ""), font(9), LEFT_TOP, B_ALL)
        ws.merge_cells(start_row=r, start_column=1, end_row=r + 2, end_column=1)
        box(ws, r, 1, r + 2, 7)
        r += 3
        put(ws, r, 1, "安全(KY)", font(9, True), CENTER, B_ALL, fill("F2F2F2"))
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
        put(ws, r, 2, KY[(x.toordinal()) % len(KY)], font(9), LEFT, B_ALL)
        box(ws, r, 1, r, 7)
        r += 2
        for j, h in enumerate(("職長", "SPC確認")):
            c = 5 + j
            put(ws, r, c, h, font(8), CENTER, B_ALL)
            put(ws, r + 1, c, "森野" if h == "職長" else "小鳥遊", font(9, color="C00000"), CENTER, B_ALL)
            ws.row_dimensions[r + 1].height = 30
        put(ws, r, 1, "野分工業株式会社", font(9))
        page(ws, landscape=False, paper=9, fit_w=1, fit_h=1, area=f"A1:G{r + 1}", footer="野分工業(株)　梶本")
    wb.save(path)
    return sheets


# ---------------------------------------------------------------- 新様式（週報式）
NOTE_2025 = {
    dt.date(2025, 10, 1): "タービン 縁切り。1号炉 運転継続（バイパス）",
    dt.date(2025, 10, 7): "ボイラ清掃 付着灰多く1日延長",
    dt.date(2025, 10, 8): "UT：前回取替部 良好。隣接 F-11・F-27 最小3.4mm → 肉盛",
    dt.date(2025, 10, 10): "肉盛部 PT 合格",
    dt.date(2025, 10, 14): "ボイラ定期事業者検査 指摘なし。水圧試験 4.5MPa 合格",
    dt.date(2025, 10, 9): "押込送風機 ランナ 良好、軸受 交換",
    dt.date(2025, 10, 17): "タービン 動翼 PT 指示なし",
    dt.date(2025, 10, 18): "潤滑油 フラッシング 延長（土曜）",
    dt.date(2025, 10, 24): "インタロック確認 OK",
}


def build_new(path):
    con = K.CONTRACTS["OH2025-2"]
    rows = K.daily_rows(K.T2025, 2025, seed_add=25)
    r_ = random.Random(W.SEED + 251)
    wb = Workbook()
    wb.remove(wb.active)
    props(wb, f"【架空】ひばり野クリーンセンター {con['title']} 作業日報（週報式）", "野分工業 職長 w2", dt.datetime(2025, 11, 4, 17, 30))
    start = dt.date(2025, 9, 29)
    weeks = []
    w0 = start
    while w0 <= dt.date(2025, 10, 29):
        weeks.append(w0)
        w0 += dt.timedelta(days=7)
    names = []
    for wk, w0 in enumerate(weeks, 1):
        days = [w0 + dt.timedelta(days=i) for i in range(6)]
        ws = wb.create_sheet(f"第{wk}週")
        names.append(ws.title)
        cw = {1: 4, 2: 10, 3: 24, 4: 8}
        for c, w in cw.items():
            ws.column_dimensions[L(c)].width = w
        C0 = 5
        for i in range(6):
            ws.column_dimensions[L(C0 + 2 * i)].width = 3.6
            ws.column_dimensions[L(C0 + 2 * i + 1)].width = 15
        CS = C0 + 12
        ws.column_dimensions[L(CS)].width = 7
        ws.column_dimensions[L(CS + 1)].width = 6
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=CS + 1)
        put(ws, 1, 1, "【架空】　作 業 日 報（週）", font(16, True), CENTER)
        ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=2)
        put(ws, 3, 1, "工事件名", font(9, True), CENTER, B_ALL, fill("F2F2F2"))
        ws.merge_cells(start_row=3, start_column=3, end_row=3, end_column=8)
        put(ws, 3, 3, con["title"], font(9), LEFT, B_ALL)
        put(ws, 3, 9, "期間", font(9, True), CENTER, B_ALL, fill("F2F2F2"))
        ws.merge_cells(start_row=3, start_column=10, end_row=3, end_column=12)
        put(ws, 3, 10, f"{days[0]:%Y/%m/%d}（月）〜{days[-1]:%m/%d}（土）", font(9), CENTER, B_ALL)
        put(ws, 3, 13, "職長", font(9, True), CENTER, B_ALL, fill("F2F2F2"))
        put(ws, 3, 14, "森野", font(9), CENTER, B_ALL)
        ws.merge_cells(start_row=4, start_column=1, end_row=4, end_column=2)
        put(ws, 4, 1, "施設", font(9, True), CENTER, B_ALL, fill("F2F2F2"))
        ws.merge_cells(start_row=4, start_column=3, end_row=4, end_column=8)
        put(ws, 4, 3, "【架空】ひばり野クリーンセンター　2号炉（タービン・空冷復水器を含む）", font(9), LEFT, B_ALL)
        box(ws, 3, 1, 4, 8)
        R = 6
        for c, h in ((1, "No"), (2, "設備"), (3, "作業項目"), (4, "班")):
            ws.merge_cells(start_row=R, start_column=c, end_row=R + 1, end_column=c)
            put(ws, R, c, h, font(9, True), CENTER, B_ALL, fill("E2EFDA"))
        for i, x in enumerate(days):
            c = C0 + 2 * i
            ws.merge_cells(start_row=R, start_column=c, end_row=R, end_column=c + 1)
            put(ws, R, c, f"{x.month}/{x.day}（{K.WD[x.weekday()]}）" + ("祝" if x == dt.date(2025, 10, 13) else ""), font(9, True),
                CENTER, B_ALL, fill("E2EFDA"))
            put(ws, R + 1, c, "人", font(8, True), CENTER, B_ALL, fill("E2EFDA"))
            put(ws, R + 1, c + 1, "作業内容", font(8, True), CENTER, B_ALL, fill("E2EFDA"))
        ws.merge_cells(start_row=R, start_column=CS, end_row=R + 1, end_column=CS)
        put(ws, R, CS, "週工数\n(人h)", font(8, True), CENTER, B_ALL, fill("E2EFDA"))
        ws.merge_cells(start_row=R, start_column=CS + 1, end_row=R + 1, end_column=CS + 1)
        put(ws, R, CS + 1, "進捗", font(8, True), CENTER, B_ALL, fill("E2EFDA"))
        r = R + 2
        active = []
        for t in K.T2025:
            if t[3] == "SPC運転":
                continue
            td = K.task_days(t, 2025)
            if any(d in days for d in td):
                active.append(t)
        daymen = {x: 0 for x in days}
        dayh = {x: 0.0 for x in days}
        for n, t in enumerate(active, 1):
            put(ws, r, 1, n, font(8), CENTER, B_ALL)
            put(ws, r, 2, t[1].replace("設備", ""), font(8), LEFT, B_ALL)
            put(ws, r, 3, t[2], font(8), LEFT, B_ALL)
            put(ws, r, 4, t[3], font(8), CENTER, B_ALL)
            wh = 0.0
            for i, x in enumerate(days):
                c = C0 + 2 * i
                hit = [l for l in rows.get(x, []) if l[0][0] == t[0]]
                if hit:
                    _, desc, men, hrs = hit[0]
                    put(ws, r, c, men, font(8), CENTER, B_ALL)
                    put(ws, r, c + 1, desc + ("（残業）" if hrs > 8 else ""), font(8), LEFT, B_ALL, fill("FFFFFF"))
                    wh += men * hrs
                    daymen[x] += men
                    dayh[x] += men * hrs
                else:
                    put(ws, r, c, None, None, None, B_ALL)
                    put(ws, r, c + 1, None, None, None, B_ALL, fill("F7F7F7"))
            td = [d for d in K.task_days(t, 2025) if d.weekday() != 6]
            done = sum(1 for d in td if d <= days[-1])
            prog = round(100 * done / len(td)) if td else 0
            put(ws, r, CS, round(wh, 1), font(8), RIGHT, B_ALL, None, "0.0")
            put(ws, r, CS + 1, "完" if prog >= 100 else f"{prog}%", font(8), CENTER, B_ALL)
            ws.row_dimensions[r].height = 22
            r += 1
        for lab, vals, nf in (("日計 人員", daymen, "0"), ("日計 工数(人h)", dayh, "0.0")):
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
            put(ws, r, 1, lab, font(8, True), CENTER, B_ALL, fill("F2F2F2"))
            for i, x in enumerate(days):
                c = C0 + 2 * i
                ws.merge_cells(start_row=r, start_column=c, end_row=r, end_column=c + 1)
                put(ws, r, c, round(vals[x], 1) if vals[x] else None, font(8, True), CENTER, B_ALL, fill("F2F2F2"), nf)
            put(ws, r, CS, round(sum(dayh.values()), 1) if lab.startswith("日計 工数") else None, font(8, True), RIGHT, B_ALL,
                fill("F2F2F2"), "0.0")
            put(ws, r, CS + 1, None, None, None, B_ALL, fill("F2F2F2"))
            r += 1
        for lab in ("天候", "特記", "安全(KY)", "SPC確認"):
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
            put(ws, r, 1, lab, font(8, True), CENTER, B_ALL, fill("F2F2F2"))
            for i, x in enumerate(days):
                c = C0 + 2 * i
                ws.merge_cells(start_row=r, start_column=c, end_row=r, end_column=c + 1)
                v = None
                if lab == "天候":
                    v = r_.choice(WEATHER) if daymen[x] else None
                elif lab == "特記":
                    v = NOTE_2025.get(x)
                elif lab == "安全(KY)":
                    v = KY[x.toordinal() % len(KY)] if daymen[x] else None
                elif lab == "SPC確認":
                    v = "小鳥遊" if daymen[x] else None
                put(ws, r, c, v, font(8, color="C00000" if lab == "SPC確認" else None), LEFT_TOP if lab == "特記" else CENTER, B_ALL)
            ws.merge_cells(start_row=r, start_column=CS, end_row=r, end_column=CS + 1)
            put(ws, r, CS, None, None, None, B_ALL)
            ws.row_dimensions[r].height = 36 if lab == "特記" else 16
            r += 1
        put(ws, r + 1, 1, "野分工業株式会社　※毎日 当日の欄を記入し、土曜に SPC へ提出（2025年度から週1枚に変更）", font(8))
        page(ws, landscape=True, paper=8, fit_w=1, fit_h=1, area=f"A1:{L(CS + 1)}{r + 1}")
    wb.save(path)
    return names


# ---------------------------------------------------------------- 市追加工事（メーカー日報）
CITY_DAY = {
    "10/10": ("資機材搬入（溶接機4台・はつり機・RT 器材は後日）、2号ボイラ 第1放射室 前壁 養生。午後 耐火物はつり 着手（F-12〜F-18 範囲）",
              {"世話役": 1, "築炉工": 3, "普通作業員": 2}, "", "新規入場者教育 6名"),
    "10/11": ("前壁 耐火物はつり F-12〜F-26・S-03〜S-05 範囲 完了。はつりガラ 搬出",
              {"世話役": 1, "築炉工": 4, "普通作業員": 2}, "", ""),
    "10/12": ("午前 水管 肉厚再測定（UT）F-12〜F-26・S-03〜S-05、切断位置けがき。午後 F-14・F-15・F-16 切断開始",
              {"世話役": 1, "溶接工": 4, "検査員": 2}, "UT 再測定 63点：最小 F-14 3.1mm", ""),
    "10/13": ("既設水管 F-14〜F-16・F-21〜F-23 切断・撤去 完了（各 1.8m、計6本）", {"世話役": 1, "溶接工": 4, "普通作業員": 1}, "",
              "撤去管 2本 サンプル保管（市 確認用）"),
    "10/14": ("開先加工 F-14〜F-16（上下 6継手）", {"世話役": 1, "溶接工": 4}, "開先角度・ルート面 確認", ""),
    "10/16": ("開先加工 F-21〜F-23、新管 6本 仮付け", {"世話役": 1, "溶接工": 4}, "新管 材料確認（ミルシート照合）", "市 監督員 日下部様 材料確認"),
    "10/17": ("新管 本溶接 F-14〜F-16 上側継手（TIG＋被覆アーク）", {"世話役": 1, "溶接工": 4}, "溶接条件 記録", ""),
    "10/18": ("本溶接 F-14〜F-16 下側、F-21 上下。肉盛溶接 F-12・F-13・F-17 着手", {"世話役": 1, "溶接工": 6}, "", ""),
    "10/19": ("本溶接 F-22・F-23 上下 完了（計12継手）。肉盛 F-18〜F-20・F-24・F-25", {"世話役": 1, "溶接工": 6}, "", ""),
    "10/20": ("午前 肉盛 F-26・S-03〜S-05 完了（計12箇所）。午後 RT 撮影 12継手（17:00〜立入制限）",
              {"世話役": 1, "溶接工": 2, "検査員": 3}, "RT 12継手 撮影", "RT 放射線管理区域 設定"),
    "10/21": ("RT 判定：F-22 下側継手 ブローホール（3類）不合格 → はつり・補修溶接、再撮影",
              {"世話役": 1, "溶接工": 2, "検査員": 2}, "RT 11継手 合格、F-22下 再RT 合格", "不合格1継手 市・SPC へ報告"),
    "10/23": ("肉盛部 PT（12箇所）、肉盛部 肉厚測定", {"世話役": 1, "検査員": 2, "溶接工": 1}, "PT 指示模様なし、肉盛後 4.6〜5.2mm", ""),
    "10/24": ("午前 水圧試験 4.5MPa 30分保持（市 日下部様・SPC 小鳥遊様 立会）漏れなし。午後 耐火物復旧 アンカー溶接",
              {"世話役": 1, "溶接工": 2, "築炉工": 2, "普通作業員": 2}, "水圧試験 合格", ""),
    "10/25": ("耐火物復旧 キャスタブル 型枠・打設 F-12〜F-20 範囲", {"世話役": 1, "築炉工": 4}, "", ""),
    "10/26": ("キャスタブル 打設 F-21〜S-05 範囲 完了 計 4.2m2、養生", {"世話役": 1, "築炉工": 4}, "", ""),
    "10/27": ("型枠 脱型、片付け・清掃、SPC へ作業場所 返却", {"世話役": 1, "築炉工": 2, "普通作業員": 1}, "", "現地作業 完了"),
}
TRADES = ["世話役", "溶接工", "築炉工", "検査員", "普通作業員"]
MACH = {"溶接工": "溶接機（TIG・アーク）", "築炉工": "はつり機・ミキサ", "検査員": "UT 探傷器 / RT 装置 / PT 器材"}


def wareki(x):
    return f"令和{x.year - 2018}年{x.month}月{x.day}日（{K.WD[x.weekday()]}）"


def build_city(path):
    con = K.CONTRACTS["CITY-R5"]
    r_ = random.Random(W.SEED + 5)
    wb = Workbook()
    wb.remove(wb.active)
    props(wb, "【架空】令和5年度 ひばり野クリーンセンター 2号ボイラ水管補修工事 工事日報", "鷹羽エンジニアリング 現場代理人 m1",
          dt.datetime(2023, 10, 28, 9, 0))
    names = []
    for k, (txt, men, q, note) in CITY_DAY.items():
        x = dt.date(2023, *K.d(k))
        ws = wb.create_sheet(f"{x.month}{x.day:02d}")
        names.append(ws.title)
        for c, w in {1: 12, 2: 12, 3: 12, 4: 12, 5: 12, 6: 12, 7: 12}.items():
            ws.column_dimensions[L(c)].width = w
        ws.merge_cells("A1:G1")
        put(ws, 1, 1, "【架空】工 事 日 報", font(15, True), CENTER)
        ws.merge_cells("F2:G2")
        put(ws, 2, 6, "鷹羽エンジニアリング株式会社", font(8), RIGHT)
        put(ws, 3, 1, "工事名", font(9, True), CENTER, B_ALL, fill("EDEDED"))
        ws.merge_cells("B3:G3")
        put(ws, 3, 2, con["title"], font(9), LEFT, B_ALL)
        box(ws, 3, 1, 3, 7)
        put(ws, 4, 1, "年月日", font(9, True), CENTER, B_ALL, fill("EDEDED"))
        ws.merge_cells("B4:D4")
        put(ws, 4, 2, wareki(x), font(9), CENTER, B_ALL)
        put(ws, 4, 5, "天候", font(9, True), CENTER, B_ALL, fill("EDEDED"))
        put(ws, 4, 6, r_.choice(WEATHER), font(9), CENTER, B_ALL)
        put(ws, 4, 7, f"{r_.randint(14, 22)}℃", font(9), CENTER, B_ALL)
        box(ws, 4, 1, 4, 7)
        put(ws, 5, 1, "作業場所", font(9, True), CENTER, B_ALL, fill("EDEDED"))
        ws.merge_cells("B5:G5")
        put(ws, 5, 2, "2号ボイラ 第1放射室 前壁（2号炉 定期整備 炉停止中）　作業時間 8:00〜17:00" + ("（RT 17:00〜20:30）" if k == "10/20" else ""), font(9), LEFT, B_ALL)
        box(ws, 5, 1, 5, 7)
        put(ws, 7, 1, "作業内容", font(9, True), CENTER, B_ALL, fill("EDEDED"))
        ws.merge_cells("A7:G7")
        ws.merge_cells("A8:G12")
        put(ws, 8, 1, txt, font(10), LEFT_TOP)
        box(ws, 8, 1, 12, 7)
        put(ws, 14, 1, "職種別人員", font(9, True), LEFT)
        for i, tname in enumerate(TRADES + ["計"], 1):
            put(ws, 15, i, tname, font(9, True), CENTER, B_ALL, fill("EDEDED"))
            v = sum(men.values()) if tname == "計" else men.get(tname)
            put(ws, 16, i, v, font(10), CENTER, B_ALL)
        put(ws, 18, 1, "使用機械", font(9, True), CENTER, B_ALL, fill("EDEDED"))
        ws.merge_cells("B18:G18")
        put(ws, 18, 2, "、".join(MACH[t] for t in TRADES if t in men and t in MACH) or "－", font(9), LEFT, B_ALL)
        box(ws, 18, 1, 18, 7)
        put(ws, 19, 1, "品質記録", font(9, True), CENTER, B_ALL, fill("EDEDED"))
        ws.merge_cells("B19:G19")
        put(ws, 19, 2, q or "－", font(9), LEFT, B_ALL)
        box(ws, 19, 1, 19, 7)
        put(ws, 20, 1, "安全・連絡", font(9, True), CENTER, B_ALL, fill("EDEDED"))
        ws.merge_cells("B20:G21")
        put(ws, 20, 2, (note + "　" if note else "") + f"KY：{KY[x.toordinal() % len(KY)]}", font(9), LEFT_TOP, B_ALL)
        ws.merge_cells("A20:A21")
        box(ws, 20, 1, 21, 7)
        for j, h in enumerate(("監督員", "現場代理人")):
            c = 6 + j
            put(ws, 23, c, h, font(8), CENTER, B_ALL)
            put(ws, 24, c, None if h == "監督員" else "宇佐見", font(9, color="C00000"), CENTER, B_ALL)
        ws.row_dimensions[24].height = 30
        page(ws, landscape=False, paper=9, fit_w=1, fit_h=1, area="A1:G24")
    wb.save(path)
    return names


FN = {
    "old": "【架空】ひばり野クリーンセンター_作業日報_2024年度1号炉定期整備工事_旧様式1日1枚.xlsx",
    "new": "【架空】ひばり野クリーンセンター_作業日報_2025年度2号炉定期整備工事_新様式週1枚.xlsx",
    "city": "【架空】ひばり野クリーンセンター_作業日報_令和5年度2号ボイラ水管補修工事_メーカー様式.xlsx",
}

if __name__ == "__main__":
    p = {k: os.path.join(OUT, v) for k, v in FN.items()}
    s1 = build_old(p["old"])
    s2 = build_new(p["new"])
    s3 = build_city(p["city"])
    print(len(s1), len(s2), len(s3))
    print(render(p["old"], [s1.index("10.12") + 1, s1.index("10.16") + 1]))
    print(render(p["new"], [3]))
    print(render(p["city"], [s3.index("1021") + 1]))
