# -*- coding: utf-8 -*-
"""作業日報
 旧：q1（メーカー子会社 職長）が建設時の様式を続けて使った「工事日報」（2015–2018、1 日 1 枚・縦型・人員累計・KY 欄）
 新：q2 が 2019 年度から使う「作業日報」（1 日 1 枚・横型・班ごとの行に工程 No・保全ｺｰﾄﾞ・工数・進捗）
     全停契約分は副職長 q3 が同じ様式で記入（書き癖が違う）
"""
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from contracts import *
from common import *

WD = "月火水木金土日"


def jd(d):
    return f"{d.year}年{d.month}月{d.day}日（{WD[d.weekday()]}）"


# ---------------------------------------------------------------- 旧様式（2018 年度 1 号炉）
OLD_DAYS = {
    date(2018, 9, 6): dict(
        weather="晴", temp="28℃",
        people=[("【架空】水樹テクノサービス", "機械", 10), ("【架空】東雲築炉工業", "築炉", 5), ("【架空】かざはや分析センター", "分析", 2)],
        work=["１号炉 炉内清掃・クリンカ除去　完了（東雲築炉）",
              "火格子摩耗測定　乾燥段　実施（可動・固定 各列3点、ノギス・デプスゲージ）",
              "　〃　　　　　燃焼段　可動側 1～4列 実施",
              "長抜差スートブロワ No.4　ランス引抜き・パッキン押え取外し",
              "BF ろ布抜取　6室×2本（分析センター 立会 SPC sp1）",
              "炉内仮設照明 増設（燃焼段・後燃焼段）"],
        tools="2.8tチェーンブロック、電動ウインチ、デプスゲージ、送風機（仮設）",
        ky="炉内作業時の酸素濃度測定（18%以上確認後入炉）、落下物注意",
        order="入炉前の酸欠測定は毎回記録すること（SPC指示）",
    ),
    date(2018, 9, 7): dict(
        weather="曇", temp="26℃",
        people=[("【架空】水樹テクノサービス", "機械", 10), ("【架空】東雲築炉工業", "築炉", 5)],
        work=["火格子摩耗測定　燃焼段（可動側 5～8列・固定側 全列）、後燃焼段　実施",
              "　→燃焼段 可動火格子の摩耗大　最大 8.6mm（管理値 15mm）　１列目・３列目に多い",
              "　→SPC sp1 へ報告済（16:30）、明日 測定点を追加（各列5点）",
              "耐火物点検　開始（打診・目視、乾燥段側壁から）",
              "長抜差スートブロワ No.4　ランス・パッキン交換　完了、手回し確認"],
        tools="デプスゲージ、超音波厚さ計（耐火物裏の鉄皮）、打診ハンマー",
        ky="測定時の手指の挟まれ（火格子間）、照明の確保",
        order="摩耗測定結果は表にまとめて報告書に添付（sp1）",
    ),
    date(2018, 9, 13): dict(
        weather="雨", temp="22℃",
        people=[("【架空】水樹テクノサービス", "機械", 10), ("【架空】東雲築炉工業", "築炉", 5)],
        work=["給じん装置　ライナ（摺動板）取付　完了、手動ストローク確認",
              "主灰押出装置　ライナ溶接（プッシャ底板 2枚）　明日 仕上げ",
              "耐火物部分補修　燃焼段右側壁 0.9m² キャスタブル施工（東雲築炉）",
              "　〃　　　　　　炉出口天井 れんが 2枚差替え"],
        tools="溶接機（200A）2台、ミキサ、バイブレータ",
        ky="溶接作業の火気管理（消火器配置）、雨天時の通路滑り",
        order="特になし",
    ),
}


def crew_count_old(d):
    """工程（実績）から、その日に現場にいた人数（累計の計算用）。"""
    n = 0
    for (no, nm, code, crew, ps, pe, as_, ae) in C18["tasks"]:
        if as_ <= d <= ae and not is_off(d) and crew != "SPC運転":
            n += CREWS[crew][2]
    return n


def make_old():
    fname = f"{FN}_作業日報_2018年度1号炉定期整備工事_q1工事日報.xlsx"
    wb = Workbook()
    setp(wb, "【架空】2018年度 かざはやエコプラザ 1号炉定期整備工事 工事日報", "【架空】水樹テクノサービス q1")
    wb.remove(wb.active)
    s, e = C18["period"]
    workdays = [d for d in drange(s, e) if not is_off(d) or d == date(2018, 9, 8)]
    for d, info in OLD_DAYS.items():
        ws = wb.create_sheet(f"{d.month}.{d.day}")
        no = workdays.index(d) + 1
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
        rows = [
            ("工事件名", "【架空】" + C18["name"]),
            ("日　付", f"{jd(d)}　　天候 {info['weather']}　　気温 {info['temp']}"),
            ("作業時間", "8:00 ～ 17:00　（昼休 12:00～13:00）　　職長 q1"),
        ]
        r = 3
        for k, v in rows:
            ws.cell(r, 1, k)
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
            ws.cell(r, 2, v)
            for col in range(1, 9):
                ws.cell(r, col).border = ALL_THIN
                ws.cell(r, col).font = font(10)
            ws.cell(r, 1).alignment = C
            ws.cell(r, 1).fill = fill("EDEDED")
            ws.row_dimensions[r].height = 20
            r += 1
        r += 1
        # 人員
        ws.cell(r, 1, "■ 出 面 （人）").font = font(10, True)
        r += 1
        hdr = ["会　社　名", "", "", "職種", "本日", "累計", "", ""]
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
        for i, t in enumerate(hdr, 1):
            ws.cell(r, i, t if t else None)
        for col in range(1, 7):
            ws.cell(r, col).border = ALL_THIN
            ws.cell(r, col).font = font(9, True)
            ws.cell(r, col).alignment = C
            ws.cell(r, col).fill = fill("EDEDED")
        r += 1
        # 累計：工程実績から（9/3 以降の各日の人数の和）
        cum_total = sum(crew_count_old(x) for x in drange(s, d))
        tot = 0
        for comp, job, n in info["people"]:
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
            ws.cell(r, 1, comp)
            ws.cell(r, 4, job)
            ws.cell(r, 5, n)
            tot += n
            for col in range(1, 7):
                ws.cell(r, col).border = ALL_THIN
                ws.cell(r, col).font = font(9)
                ws.cell(r, col).alignment = C if col > 3 else CL
            r += 1
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
        ws.cell(r, 1, "合　計")
        ws.cell(r, 5, tot)
        ws.cell(r, 6, cum_total)
        for col in range(1, 7):
            ws.cell(r, col).border = box(t=dbl)
            ws.cell(r, col).font = font(9, True)
            ws.cell(r, col).alignment = C
        r += 2
        ws.cell(r, 1, "■ 作 業 内 容").font = font(10, True)
        r += 1
        top = r
        lines = info["work"] + [""] * (9 - len(info["work"]))
        num = 0
        for i, t in enumerate(lines):
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
            if t and not t.startswith("　"):
                num += 1
            ws.cell(r, 1, (f"{num}. " if t and not t.startswith("　") else "") + t)
            ws.cell(r, 1).font = font(10)
            for col in range(1, 9):
                ws.cell(r, col).border = box(l=thin if col == 1 else NONE, r=thin if col == 8 else NONE,
                                             t=thin if r == top else dot, b=dot)
            ws.row_dimensions[r].height = 19
            r += 1
        for col in range(1, 9):
            ws.cell(r - 1, col).border = box(l=thin if col == 1 else NONE, r=thin if col == 8 else NONE, t=dot, b=thin)
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
            ws.cell(r + 1, col, "sp1" if i == 0 else "q1")
            ws.cell(r + 1, col).font = font(12, color="C00000")
            for rr in range(r, r + 4):
                for cc in (col, col + 1):
                    ws.cell(rr, cc).border = ALL_THIN
                    ws.cell(rr, cc).alignment = C
            ws.cell(r, col).font = font(9)
        ws.print_area = f"A1:H{r + 3}"
        ws.page_setup.orientation = "portrait"
        ws.page_setup.paperSize = 9
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 1
        ws.sheet_properties.pageSetUpPr.fitToPage = True
    wb.save(outpath(fname))
    return fname


# ---------------------------------------------------------------- 新様式（2023 年度）
# (班, 工程No, 作業内容, 対象, 人数, 工数h, 進捗%, 備考)
NEW_DAYS_C23 = {
    date(2023, 10, 5): dict(weather="晴", rows=[
        ("足場", 3, "炉内足場 組立（燃焼室・ボイラ下部）", "", 4, 32, 70, ""),
        ("耐火", 4, "炉内クリンカ除去（乾燥段～燃焼段上部）", "1NS-010", 5, 40, 40, "ブレーカ使用"),
        ("機械B", 10, "誘引通風機 カップリング切離し・軸受ハウジング開放", "1KT-100", 4, 32, 20, "分解前振動 5.8mm/s"),
    ], special="", witness="SPC sp2 巡回（10:30）"),
    date(2023, 10, 9): dict(weather="曇", rows=[
        ("機械A", 6, "燃焼火格子 撤去 可動側1～4列（64枚）", "1NS-030", 6, 54, 30, "祝日作業"),
        ("耐火", 8, "耐火物 打診点検（乾燥段・燃焼段 側壁）", "1NS-010", 5, 40, 15, ""),
        ("機械B", 10, "誘引通風機 軸受交換（DE・NDE）、インペラ摩耗部 肉盛溶接 開始", "1KT-100", 4, 32, 55, "肉盛 前縁 6枚"),
    ], special="", witness=""),
    date(2023, 10, 11): dict(weather="晴", rows=[
        ("機械A", 6, "燃焼火格子 撤去 完了（可動96・固定96 計192枚）、火格子受台 清掃・点検", "1NS-030", 6, 48, 100, ""),
        ("耐火", 8, "側壁下部 損傷部 調査・マーキング（燃焼段 左右）", "1NS-010", 5, 40, 30, "下記特記"),
        ("機械B", 10, "誘引通風機 インペラ肉盛 完了・仕上げ研削", "1KT-100", 4, 32, 75, ""),
        ("電計", 11, "誘引通風機電動機 軸受交換（6.6kV 280kW）", "1KT-100M", 3, 24, 70, ""),
        ("検査", 9, "肉厚測定 1次・2次・3次過熱器", "1BO-050 1BO-060 1BO-070", 2, 16, 60, "3SH 下段 最小5.6mm"),
        ("機械C", 12, "BF逆洗 電磁弁ダイヤフラム交換 室1～3（60個）", "1HG-140", 3, 24, 50, ""),
        ("分析", 13, "ろ布抜取 6本（各室1本）、触媒 抜取 2本（1・2層）", "1HG-130 1HG-380", 2, 8, 100, "結果 11月中旬"),
    ], special="燃焼段 左右側壁下部（火格子上 0～400mm）に耐火物の剥離・欠損。範囲 計2.4m²、最大深さ85mm。"
                 "当初補修範囲（乾燥段側壁 1.2m²）の外。15:40 SPC sp2 へ報告、写真送付。明日10:00 変更協議。",
        witness="SPC sp2 現地確認（16:00）"),
    date(2023, 10, 12): dict(weather="晴", rows=[
        ("機械A", 7, "燃焼火格子 新品取付 1～2列（48枚）、ボルト本締め", "1NS-030", 6, 54, 25, ""),
        ("耐火", 8, "側壁下部 追加補修 はつり（承認後 11:00～）", "1NS-010", 7, 63, 45, "2名増員"),
        ("機械B", 10, "誘引通風機 動バランス調整・組立", "1KT-100", 4, 32, 90, "G2.5"),
        ("電計", 11, "誘引通風機電動機 絶縁診断 完了（PI 3.8）", "1KT-100M", 3, 24, 100, ""),
        ("検査", 9, "肉厚測定 エコノマイザ・蒸気ドラム", "1BO-090 1BO-020", 2, 16, 85, ""),
        ("機械C", 12, "BFダイヤフラム 室4～6、室別入口ダンパ3 分解", "1HG-140 1HG-170", 3, 24, 80, ""),
    ], special="10:00 変更協議（SPC sp2・水樹TS q2）：側壁下部 2.4m² 追加補修を承認。耐火 2名増員し工期内で吸収。"
                 "材料は東雲築炉 在庫のキャスタブル（高アルミナ質）を使用、不足分 1.2t は東雲築炉で手配。",
        witness="SPC sp2"),
    date(2023, 10, 16): dict(weather="晴", rows=[
        ("機械A", 7, "燃焼火格子 取付 5～8列", "1NS-030", 6, 48, 70, ""),
        ("耐火", 8, "側壁下部 キャスタブル施工（追加部 左側）", "1NS-010", 7, 56, 75, ""),
        ("機械B", 14, "ボイラ給水ポンプ1A 分解（カップリング・メカシール取外し）", "1BO-100", 4, 32, 20, "メカシール漏れ跡あり"),
        ("機械C", 16, "押込送風機ダンパ駆動装置 分解", "1KT-020", 3, 24, 25, ""),
        ("電計", 17, "熱電対交換 炉出口1・2、水銀計 旧機撤去", "1NS-330 1NS-340 1HG-470", 3, 24, 30, ""),
        ("検査", 15, "減温塔 内面点検・肉厚測定", "1HG-010", 2, 16, 50, ""),
    ], special="10/15（日）全停 受変電停電。7:30 復電確認、1号炉関係の仮設電源・照明 異常なし。", witness=""),
    date(2023, 10, 19): dict(weather="雨", rows=[
        ("機械A", 7, "燃焼火格子 全数取付完了・動作確認、クリアランス測定 → 3列 手直し開始", "1NS-030", 6, 57, 90, "残業1.5h"),
        ("耐火", 8, "追加補修部 乾燥養生、仕上げ", "1NS-010", 2, 16, 95, ""),
        ("機械B", 14, "ボイラ給水ポンプ1A 組立（メカシール・軸受・ライナリング新品）", "1BO-100", 4, 32, 80, ""),
        ("機械C", 16, "NH3注入ポンプA ダイヤフラム交換、噴霧水ポンプB電動機 軸受交換", "1HG-400 1HG-030M", 3, 24, 100, ""),
        ("電計", 17, "水銀連続分析計 新機据付・サンプリング配管接続", "1HG-470", 3, 24, 60, ""),
    ], special="燃焼火格子 可動/固定間クリアランス 基準2～5mm に対し 6・7・8列目が最大7.5mm。"
                 "固定火格子の受け金具の位置ずれ（据付時の締付順）による。3列を緩めて再調整、明日完了見込み。",
        witness="SPC sp2 クリアランス測定立会"),
    date(2023, 10, 26): dict(weather="晴", rows=[
        ("機械A", 22, "ボイラ定期事業者検査 受検（10:00～15:00）合格", "1BO-010", 6, 48, 100, ""),
        ("足場", 19, "足場解体 完了・搬出", "", 4, 32, 100, ""),
        ("電計", 23, "復旧前 配線・端子 確認（熱電対・水銀計・IDF電動機）", "", 3, 24, 30, ""),
    ], special="", witness="SPC sp2、電気主任技術者（ボイラ・タービン主任技術者 SPC sp5）"),
    date(2023, 10, 31): dict(weather="晴", rows=[
        ("SPC", 25, "炉立上げ 助燃バーナ点火 9:00、昇温", "", None, None, 50, "SPC運転"),
        ("機械A", 25, "立上げ監視（火格子動作・ボルト増締め、IDF振動 2.1mm/s）", "1NS-030 1KT-100", 3, 24, None, ""),
        ("電計", 25, "計装 指示確認（炉出口温度・水銀計）", "1NS-330 1HG-470", 2, 16, None, ""),
    ], special="", witness=""),
}

NEW_DAYS_C23C = {
    date(2023, 10, 15): dict(weather="晴", rows=[
        ("受変電", 3, "特高GIS・主変・高圧盤・所内変 年次点検、保護継電器試験", "0DE-010～160", 8, 64.0, 100, "停電 08:05-16:40"),
        ("受変電", 4, "非常用発電機 負荷試験（30分・実負荷）", "0DE-270", 0, 0.0, 100, "上記人員で実施"),
        ("電計2", 2, "停電操作 立会・仮設照明", "", 2, 18.0, 100, ""),
        ("機械D", 10, "蒸気だめ 開放準備（保温材撤去）", "0YN-190 0YN-200", 3, 24.0, 20, ""),
    ], special="復電 16:40。主変 油中ガス 採油済（結果待ち）。", witness="SPC sp2、電気主任技術者（外部委託）"),
    date(2023, 10, 19): dict(weather="雨", rows=[
        ("機械D", 8, "ACCファン減速機3 分解・点検、軸受交換", "0YN-300", 3, 27.0, 50, "下記"),
        ("機械D", 9, "脱気器 性能検査 受検（10:00）", "0YN-400", 0, 0.0, 100, "合格"),
        ("SV", 5, "タービン主油ポンプ 組立後の確認（立会）", "0YN-030", 1, 4.0, 100, "鳴瀬タービン"),
    ], special="減速機3 低速段歯面に軽微なピッチング（面積率 約3%）。継続使用可と判断、2025年度に歯車交換を提言。"
                 "異音の原因は低速軸 出力側軸受の損傷（外輪剥離）。",
        witness="SPC sp2"),
}


def _split_rows(rows):
    """班ごとの記入を保全ｺｰﾄﾞ 1 件 1 行に分ける。人数・工数はその作業の最初の行だけに書き、続く行は縦に結合する。
    ｺｰﾄﾞの無い作業（足場・運転など）は最後に「共通作業」としてまとめる。"""
    from world import BY_CODE
    coded, plain = [], []
    for (crew, no, txt, code, n, hrs, prog, note) in rows:
        codes = code.split() if code else []
        if not codes:
            plain.append([("", "（ｺｰﾄﾞなし）", txt, crew, n, hrs, prog, no, note)])
            continue
        grp = []
        for k, cd in enumerate(codes):
            nm = BY_CODE[cd]["name"] if cd in BY_CODE else "受変電設備 一式"
            grp.append((cd, nm, txt if k == 0 else "〃", crew, n, hrs, prog, no, note if k == 0 else ""))
        coded.append(grp)
    coded.sort(key=lambda g: g[0][0])
    return coded, plain


def make_new(c, days, fname, writer, author):
    """q2 様式（2019〜）：SPC 提出用。行は保全ｺｰﾄﾞ順（PM-NAVI に転記しやすい並び）、工程No は右の小さな列。"""
    wb = Workbook()
    setp(wb, f"【架空】{c['name']} 作業日報", f"【架空】水樹テクノサービス {writer}")
    wb.remove(wb.active)
    q3 = writer == "q3"
    NC = 9
    for d, info in days.items():
        ws = wb.create_sheet(f"{d:%m%d}")
        widths = [11, 22, 42, 7, 5, 6.5, 6, 5, 20]
        for i, wdt in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = wdt
        ws.merge_cells("A1:E2")
        ws["A1"] = "作 業 日 報"
        ws["A1"].font = font(16, True)
        ws["A1"].alignment = Alignment(horizontal="left", vertical="center")
        # 右上：SPC 確認欄（大きめ。確認者・PM-NAVI 転記の済／未）
        for (rng, txt) in [("F1:G1", "SPC確認"), ("H1:I1", "PM-NAVI転記")]:
            ws.merge_cells(rng)
            ws[rng.split(":")[0]] = txt
        ws.merge_cells("F2:G3")
        ws["F2"] = "sp2"
        ws.merge_cells("H2:I3")
        ws["H2"] = "済　（　/　）" if not q3 else "未"
        for rr in range(1, 4):
            for cc in range(6, 10):
                ws.cell(rr, cc).border = ALL_THIN
                ws.cell(rr, cc).alignment = C
        ws["F1"].font = font(9); ws["H1"].font = font(9)
        ws["F2"].font = font(14, color="C00000"); ws["H2"].font = font(10)
        ws["A3"] = f"記入 {writer}" + ("（全停）" if q3 else "")
        ws["A3"].font = font(9)
        ws["A4"] = "件名"
        ws.merge_cells("B4:C4")
        ws["B4"] = "【架空】" + c["name"]
        ws["D4"] = "日付"
        ws.merge_cells("E4:I4")
        ws["E4"] = (f"{d:%Y/%m/%d}（{WD[d.weekday()]}）　{info['weather']}" if q3 else f"{jd(d)}　{info['weather']}")
        for col in range(1, NC + 1):
            ws.cell(4, col).border = ALL_THIN
            ws.cell(4, col).font = font(10)
        ws["A4"].fill = fill("DEEAF6"); ws["D4"].fill = fill("DEEAF6")
        hdr = ["保全ｺｰﾄﾞ", "機器名（PM-NAVI）", "作業内容", "班", "人数", "工数(h)", "進捗%", "工程", "備考"]
        r = 6
        for i, t in enumerate(hdr, 1):
            cc = ws.cell(r, i, t)
            cc.font = font(9 if i != 8 else 7, True)
            cc.fill = fill("DEEAF6")
            cc.border = box(t=med, b=thin, l=thin, r=thin)
            cc.alignment = C
        r += 1
        coded, plain = _split_rows(info["rows"])
        tn = th = 0
        used = 0
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
                    cc = ws.cell(r, i, v if v is not None else None)
                    cc.font = font(9 if i != 8 else 8)
                    cc.border = box(l=thin, r=thin, t=thin if k == 0 else hair, b=hair)
                    cc.alignment = CLW if i in (2, 3, 9) else C
                    if i == 6 and q3 and isinstance(v, float):
                        cc.number_format = "0.0"
                ws.row_dimensions[r].height = 22 if len(grp) == 1 else 16
                r += 1
                used += 1
            if len(grp) > 1:
                for col in (4, 5, 6, 7, 8):
                    ws.merge_cells(start_row=top, start_column=col, end_row=r - 1, end_column=col)
            tn += grp[0][4] or 0
            th += grp[0][5] or 0
        for _ in range(max(0, 12 - used)):
            for i in range(1, NC + 1):
                ws.cell(r, i).border = box(l=thin, r=thin, t=hair, b=hair)
            ws.row_dimensions[r].height = 18
            r += 1
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
        ws.cell(r, 1, "計（協力会社含む・SPC除く）")
        ws.cell(r, 5, tn)
        ws.cell(r, 6, th)
        if q3:
            ws.cell(r, 6).number_format = "0.0"
        for col in range(1, NC + 1):
            cc = ws.cell(r, col)
            cc.font = font(9, True)
            cc.border = box(t=dbl, b=med, l=thin, r=thin)
            cc.alignment = C
        r += 2
        for k, v, hgt in [("特記事項（発見・手直し・変更）", info["special"], 64), ("立会・検査", info["witness"], 22),
                          ("安全", "KY実施・入炉前酸素濃度測定 記録済　災害 無" if not q3 else "KY済 停電操作は二重確認 災害無", 22)]:
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
            r += 1
        ws.cell(r, 1, "様式 MTS-日報-2019（q2）SPC提出用・保全ｺｰﾄﾞ順" + ("　※全停分 q3 記入" if q3 else "")).font = font(7, color="808080")
        ws.print_area = f"A1:I{r}"
        ws.page_setup.orientation = "landscape"
        ws.page_setup.paperSize = 9
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 1
        ws.sheet_properties.pageSetUpPr.fitToPage = True
    wb.save(outpath(fname))
    return fname


def run():
    return [make_old(),
            make_new(C23, NEW_DAYS_C23, f"{FN}_作業日報_2023年度1号炉定期整備工事_q2作業日報.xlsx", "q2", "q2"),
            make_new(C23C, NEW_DAYS_C23C, f"{FN}_作業日報_2023年度全炉停止時共通設備点検整備工事_q2作業日報.xlsx", "q3", "q3")]


if __name__ == "__main__":
    for f in run():
        print(f)
